#!/usr/bin/env python3
"""Supabase 영속화 — 업로드 파일은 Storage, 작업 기록은 Postgres에 남긴다.

지금까지 업로드 파일과 취합 결과는 임시 폴더 + 프로세스 메모리에만 있어
서버를 재시작하면 사라졌다(PRD §10 잔여 리스크). 이 모듈이 그 기록을
design.md §3의 테이블·버킷에 남겨, 재시작 후에도 파일과 결과가 남고
F4-2 작업 이력의 전제가 채워진다.

경계: 진행 중이던 작업을 재시작 후 이어서 하는 것은 여전히 안 된다.
파싱된 작업 집합은 메모리에 있고 되살리려면 파일을 다시 내려받아 다시
읽어야 한다. 이 모듈이 보장하는 것은 "기록과 산출물이 남는다"까지다.

인증이 꺼진 개발 모드(AUTH_DISABLED=1)에서는 소유자가 없어 전부 건너뛴다.
"""

from __future__ import annotations

import json
import mimetypes
import os
import uuid
from pathlib import Path

import httpx

import auth

UPLOAD_BUCKET = "uploads"
RESULT_BUCKET = "results"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HWPX_MIME = "application/hwp+zip"
TIMEOUT = 60          # 대용량 xlsx 업로드를 감당할 여유


def enabled() -> bool:
    return auth.configured()


def _base() -> str:
    return os.environ["SUPABASE_URL"].rstrip("/")


def _headers(extra: dict | None = None) -> dict:
    key = auth._secret_key()
    return {"apikey": key, "Authorization": f"Bearer {key}", **(extra or {})}


def _rest(method: str, path: str, prefer: str | None = None, **kw) -> httpx.Response:
    extra = {"Content-Type": "application/json"}
    if prefer:
        extra["Prefer"] = prefer
    res = httpx.request(method, f"{_base()}/rest/v1{path}",
                        headers=_headers(extra), timeout=TIMEOUT, **kw)
    if res.status_code >= 400:
        raise RuntimeError(f"Supabase REST {method} {path} → {res.status_code} {res.text[:300]}")
    return res


def _insert(table: str, row: dict) -> dict:
    res = _rest("POST", f"/{table}", prefer="return=representation", json=row)
    return res.json()[0]


# ── 프로젝트 ──────────────────────────────────────────────────────────────────
def create_project(owner_id: str, name: str) -> str:
    """취합 작업 1건 = 프로젝트 1개. 사용자가 고르는 UI는 아직 없어 자동 생성한다."""
    return _insert("projects", {"owner_id": owner_id, "name": name})["id"]


# ── 업로드 파일 ───────────────────────────────────────────────────────────────
def put_upload(project_id: str, uploader_id: str, local_path: Path, original_name: str,
               kind: str = "excel") -> dict:
    """파일을 uploads 버킷에 올리고 uploaded_files 행을 남긴다.

    Storage 키에는 원본 파일명을 쓰지 않는다 — 한글·공백이 섞인 이름이 키로
    들어가면 인코딩 문제가 생기고, 원본명은 DB에 그대로 보관하면 충분하다.
    확장자·MIME은 kind에서 정한다(excel=xlsx / hwpx=hwpx).
    """
    file_id = str(uuid.uuid4())
    ext, mime = _kind_format(kind)
    storage_path = f"{project_id}/{file_id}.{ext}"
    data = local_path.read_bytes()
    _put_object(UPLOAD_BUCKET, storage_path, data, mime)
    row = _insert("uploaded_files", {
        "id": file_id, "project_id": project_id, "uploader_id": uploader_id,
        "storage_path": storage_path, "original_name": original_name,
        "kind": kind, "status": "uploaded", "size_bytes": len(data),
    })
    return row


def _kind_format(kind: str) -> tuple[str, str]:
    """프로젝트 유형 → (확장자, MIME)."""
    return ("hwpx", HWPX_MIME) if kind == "hwpx" else ("xlsx", XLSX_MIME)


def _put_object(bucket: str, path: str, data: bytes, content_type: str | None = None) -> None:
    ctype = content_type or mimetypes.guess_type(path)[0] or "application/octet-stream"
    res = httpx.post(f"{_base()}/storage/v1/object/{bucket}/{path}",
                     headers=_headers({"Content-Type": ctype, "x-upsert": "true"}),
                     content=data, timeout=TIMEOUT)
    if res.status_code >= 400:
        raise RuntimeError(f"Storage 업로드 실패 {bucket}/{path} → {res.status_code} {res.text[:300]}")


def get_object(bucket: str, path: str) -> bytes:
    res = httpx.get(f"{_base()}/storage/v1/object/{bucket}/{path}", headers=_headers(), timeout=TIMEOUT)
    if res.status_code >= 400:
        raise RuntimeError(f"Storage 다운로드 실패 {bucket}/{path} → {res.status_code}")
    return res.content


# ── 검토 결과 ─────────────────────────────────────────────────────────────────
def save_review(file_id: str, grade: str, issues: list[dict]) -> None:
    """review_results.badge는 화면의 정상/이상 이진 표기가 아니라 내부 등급이다.

    §6.1의 이진 표기는 화면 전용이고, 취합 포함 여부(§7)는 등급으로 정해지므로
    기록에는 오류/경고/정상 등급을 남긴다.
    """
    # 재검토하면 같은 파일에 결과가 쌓이므로 이전 것을 지우고 다시 넣는다
    _rest("DELETE", "/review_results", params={"file_id": f"eq.{file_id}"})
    _insert("review_results", {
        "file_id": file_id, "badge": grade, "issue_count": len(issues),
        "issues_json": issues,
    })
    _rest("PATCH", "/uploaded_files", params={"id": f"eq.{file_id}"}, json={"status": "reviewed"})


# ── 취합 잡 ───────────────────────────────────────────────────────────────────
def create_job(project_id: str, mode: str, included_file_ids: list[str],
               summary_fields: list[str], kind: str = "excel") -> str:
    return _insert("aggregation_jobs", {
        "project_id": project_id, "kind": kind, "mode": mode, "status": "running",
        "included_file_ids": included_file_ids, "summary_fields_json": summary_fields,
        "progress": 0, "stats_json": {},
    })["id"]


def update_job(job_id: str, **fields) -> None:
    if fields:
        _rest("PATCH", "/aggregation_jobs", params={"id": f"eq.{job_id}"}, json=fields)


def put_result(project_id: str, job_id: str, local_path: Path, label: str) -> str:
    """취합 결과·오류 리포트를 results 버킷에 올리고 경로를 돌려준다.

    확장자는 만들어진 파일 그대로 따라간다 — 엑셀 취합은 xlsx, 한글 병합은 hwpx다.
    """
    ext = local_path.suffix.lstrip(".").lower() or "bin"
    mime = HWPX_MIME if ext == "hwpx" else XLSX_MIME
    path = f"{project_id}/{job_id}_{label}.{ext}"
    _put_object(RESULT_BUCKET, path, local_path.read_bytes(), mime)
    return path


# ── 이력 (F4-2의 재료) ────────────────────────────────────────────────────────
def list_projects(owner_id: str, q: str = "", kind: str = "", limit: int = 200) -> list[dict]:
    """이력 목록. 검색·유형 필터는 파일명까지 봐야 하므로 여기서 걸러낸다.

    PostgREST로 중첩 테이블(파일명)을 걸러내려면 쿼리가 복잡해지는데,
    §6.1 규모(사용자당 90일 보관)에서는 목록이 크지 않아 이득이 없다.
    """
    res = _rest("GET", "/projects", params={
        "owner_id": f"eq.{owner_id}",
        "select": "id,name,created_at,"
                  "uploaded_files(original_name,kind,status),"
                  "aggregation_jobs(id,kind,mode,status,progress,result_url,error_message,created_at)",
        "order": "created_at.desc", "limit": str(limit)})
    rows = res.json()

    needle = (q or "").strip().lower()
    for row in rows:
        row["file_names"] = [f["original_name"] for f in row.get("uploaded_files") or []]
        row["jobs"] = sorted(row.pop("aggregation_jobs", None) or [],
                             key=lambda j: j["created_at"], reverse=True)
        # 유형은 취합 잡이 있으면 그 kind, 없으면 업로드 파일의 kind를 따른다
        kinds = {j["kind"] for j in row["jobs"]} | {f["kind"] for f in row.get("uploaded_files") or []}
        row["kind"] = ("excel" if "excel" in kinds else next(iter(kinds), "excel"))
        row["file_count"] = len(row["file_names"])
        row.pop("uploaded_files", None)

    if kind:
        rows = [r for r in rows if r["kind"] == kind]
    if needle:
        rows = [r for r in rows
                if needle in r["name"].lower()
                or any(needle in n.lower() for n in r["file_names"])]
    return rows


def project_detail(owner_id: str, project_id: str) -> dict | None:
    """소유자 확인을 쿼리에 포함해, 남의 프로젝트는 조회 자체가 비게 한다."""
    res = _rest("GET", "/projects", params={
        "id": f"eq.{project_id}", "owner_id": f"eq.{owner_id}",
        "select": "id,name,created_at,"
                  "uploaded_files(id,original_name,status,created_at,"
                  "review_results(badge,issue_count)),"
                  "aggregation_jobs(id,mode,status,progress,result_url,error_message,created_at)"})
    rows = res.json()
    return rows[0] if rows else None


# ── Admin (F4-3) ──────────────────────────────────────────────────────────────
def _count(table: str, **filters) -> int:
    """행 수만 센다. Prefer: count=exact면 본문 대신 Content-Range로 받는다."""
    params = {"select": "id", "limit": "1", **{k: v for k, v in filters.items()}}
    res = _rest("GET", f"/{table}", prefer="count=exact", params=params)
    rng = res.headers.get("content-range", "")
    return int(rng.split("/")[-1]) if "/" in rng and rng.split("/")[-1].isdigit() else 0


def dashboard() -> dict:
    """§3.2 시스템 현황. API 비용은 집계 경로가 없어 None으로 돌려준다."""
    jobs = _rest("GET", "/aggregation_jobs", params={"select": "status"}).json()
    by_status: dict[str, int] = {}
    for job in jobs:
        by_status[job["status"]] = by_status.get(job["status"], 0) + 1

    sizes = _rest("GET", "/uploaded_files", params={"select": "size_bytes"}).json()
    known = [row["size_bytes"] for row in sizes if row.get("size_bytes") is not None]
    return {
        "jobs_total": len(jobs),
        "jobs_by_status": by_status,
        "users_total": _count("profiles"),
        "users_active": _count("profiles", status="eq.active"),
        "files_total": len(sizes),
        # 이 컬럼 추가 전에 올라온 행은 크기를 모른다. 0으로 단정하지 않고 따로 알린다
        "storage_bytes": sum(known),
        "storage_unknown_files": len(sizes) - len(known),
        "api_cost": None,          # OpenAI 사용량 집계 경로가 아직 없다
        "retention_days": get_policy().get("retention_days"),
    }


def list_accounts() -> list[dict]:
    res = _rest("GET", "/profiles", params={
        "select": "id,employee_no,reset_email,role,status,created_at",
        "order": "created_at.asc"})
    return res.json()


def patch_account(user_id: str, fields: dict) -> dict | None:
    allowed = {k: v for k, v in fields.items() if k in ("role", "status")}
    if not allowed:
        return None
    res = _rest("PATCH", "/profiles", prefer="return=representation",
                params={"id": f"eq.{user_id}"}, json=allowed)
    rows = res.json()
    return rows[0] if rows else None


def delete_account(user_id: str) -> None:
    """계정과 그 자료를 지운다.

    DB는 auth.users → profiles → projects → uploaded_files까지 CASCADE로
    사라지지만 Storage 객체는 대상이 아니다. 프로젝트 목록을 알 수 있는 동안
    Storage를 먼저 비운다.
    """
    projects = _rest("GET", "/projects",
                     params={"owner_id": f"eq.{user_id}", "select": "id"}).json()
    for project in projects:
        for bucket in (UPLOAD_BUCKET, RESULT_BUCKET):
            _empty_folder(bucket, project["id"])
    res = httpx.delete(f"{_base()}/auth/v1/admin/users/{user_id}",
                       headers=_headers(), timeout=TIMEOUT)
    if res.status_code >= 400:
        raise RuntimeError(f"계정 삭제 실패: {res.status_code} {res.text[:200]}")


def _empty_folder(bucket: str, prefix: str) -> None:
    listed = httpx.post(f"{_base()}/storage/v1/object/list/{bucket}",
                        headers=_headers({"Content-Type": "application/json"}),
                        json={"prefix": f"{prefix}/", "limit": 1000}, timeout=TIMEOUT)
    if listed.status_code >= 400:
        return
    names = [f"{prefix}/{obj['name']}" for obj in listed.json()]
    if names:
        httpx.request("DELETE", f"{_base()}/storage/v1/object/{bucket}",
                      headers=_headers({"Content-Type": "application/json"}),
                      json={"prefixes": names}, timeout=TIMEOUT)


def get_policy() -> dict:
    rows = _rest("GET", "/retention_policy", params={"select": "key,value"}).json()
    out = {}
    for row in rows:
        value = row["value"]
        out[row["key"]] = int(value) if str(value).isdigit() else value
    return out


def set_policy(key: str, value, actor_id: str) -> dict:
    _rest("POST", "/retention_policy", prefer="resolution=merge-duplicates",
          json={"key": key, "value": str(value), "updated_by": actor_id})
    return get_policy()


def log_action(actor_id: str | None, action: str, target: str = "") -> None:
    """감사 로그. 기록 실패가 본 작업을 막지는 않는다."""
    try:
        _insert("audit_logs", {"actor_id": actor_id, "action": action, "target": target})
    except Exception:
        pass


def list_logs(q: str = "", limit: int = 200) -> list[dict]:
    res = _rest("GET", "/audit_logs", params={
        "select": "id,action,target,created_at,profiles(employee_no)",
        "order": "created_at.desc", "limit": str(limit)})
    rows = res.json()
    for row in rows:
        profile = row.pop("profiles", None) or {}
        row["actor"] = profile.get("employee_no") or "(삭제된 계정)"
    needle = (q or "").strip().lower()
    if needle:
        rows = [r for r in rows
                if needle in r["action"].lower() or needle in (r["target"] or "").lower()
                or needle in r["actor"].lower()]
    return rows


# ── F1 양식 생성 (문답 세션 · 템플릿 · 작성기준) ───────────────────────────────
FORM_BUCKET = RESULT_BUCKET       # 생성된 양식도 results 버킷에 둔다(버킷을 늘리지 않는다)


def create_intake_session(project_id: str, first_message: dict | None = None) -> dict:
    """문답 세션을 시작한다. 프로젝트당 active 세션은 하나뿐이다(부분 유니크 인덱스).

    첨부가 먼저 오는 경로(F6)에서는 첫 메시지가 아직 없다 — 빈 대화로 시작한다.
    """
    return _insert("intake_sessions", {
        "project_id": project_id, "status": "active",
        "messages_json": [first_message] if first_message else [],
        "spec_json": {}, "turn_count": 0,
    })


def get_intake_session(session_id: str) -> dict | None:
    rows = _rest("GET", "/intake_sessions", params={"id": f"eq.{session_id}", "select": "*"}).json()
    return rows[0] if rows else None


def update_intake_session(session_id: str, **fields) -> None:
    if fields:
        fields.setdefault("updated_at", "now()")
        _rest("PATCH", "/intake_sessions", params={"id": f"eq.{session_id}"}, json=fields)


def owner_of_project(project_id: str) -> str | None:
    """프로젝트 소유자. 세션·템플릿 접근을 소유자로 스코핑하기 위해 쓴다 (§6.2)."""
    rows = _rest("GET", "/projects",
                 params={"id": f"eq.{project_id}", "select": "owner_id"}).json()
    return rows[0]["owner_id"] if rows else None


def create_form_template(project_id: str, session_id: str, spec: dict) -> str:
    return _insert("form_templates", {
        "project_id": project_id, "intake_session_id": session_id,
        "spec_json": spec, "status": "generating", "version": 1, "output_format": "xlsx",
    })["id"]


def finish_form_template(template_id: str, workbook: dict, file_url: str) -> None:
    _rest("PATCH", "/form_templates", params={"id": f"eq.{template_id}"},
          json={"status": "done", "workbook_json": workbook, "file_url": file_url,
                "updated_at": "now()"})


def fail_form_template(template_id: str, reason: str) -> None:
    _rest("PATCH", "/form_templates", params={"id": f"eq.{template_id}"},
          json={"status": "failed", "error_message": reason[:500], "updated_at": "now()"})


def template_of_session(session_id: str) -> dict | None:
    """이 문답으로 만든 최신 양식. F1-7 대화형 수정이 무엇을 고칠지 찾는 데 쓴다."""
    rows = _rest("GET", "/form_templates", params={
        "intake_session_id": f"eq.{session_id}", "status": "eq.done",
        "select": "*", "order": "version.desc", "limit": "1"}).json()
    return rows[0] if rows else None


def revise_form_template(template_id: str, version: int, spec: dict, workbook: dict,
                         file_url: str) -> None:
    """F1-7 수정 결과를 같은 행에 새 버전으로 올린다.

    행을 새로 만들지 않는 이유는 목록·취합·이력이 늘 최신 버전을 가리켜야 하기 때문이다.
    이전 버전 파일은 버전별 경로로 Storage에 남는다(`put_form`).
    """
    _rest("PATCH", "/form_templates", params={"id": f"eq.{template_id}"},
          json={"version": version, "spec_json": spec, "workbook_json": workbook,
                "file_url": file_url, "updated_at": "now()"})


def get_form_template(template_id: str) -> dict | None:
    rows = _rest("GET", "/form_templates", params={"id": f"eq.{template_id}", "select": "*"}).json()
    return rows[0] if rows else None


def put_form(project_id: str, template_id: str, local_path: Path, version: int = 1) -> str:
    """생성된 양식을 Storage에 올린다. 버전별 경로로 남겨 과거 버전이 덮이지 않게 한다.

    경로는 프로젝트 폴더 **바로 아래 평면**이어야 한다. 정리 코드(`delete_account` →
    `_empty_folder`)는 `{project_id}/`를 훑어 나온 객체만 지우는데, 목록 API는 한 단계
    아래 폴더를 객체가 아닌 항목으로 돌려준다. 종전 `forms/{project_id}/...`는 두 가지가
    다 어긋나 있어서 계정을 지워도 생성한 양식이 Storage에 그대로 남았다.
    """
    path = f"{project_id}/form_{template_id}_v{version}.xlsx"
    _put_object(RESULT_BUCKET, path, local_path.read_bytes(), XLSX_MIME)
    return path


def put_attachment(project_id: str, local_path: Path, original_name: str) -> str:
    """문답 첨부를 Storage에 올린다 (F1-3·F6-2).

    경로는 다른 용도와 같은 규칙 — **프로젝트 폴더 바로 아래 평면**이다(put_form 주석 참조).
    F6 등록은 이 객체를 그대로 양식 파일로 삼는다. 다시 쓰지 않으므로 등록 전후 바이트가
    같다는 것이 구조적으로 보장된다(인지 §7 E1 무변경 등록).
    """
    suffix = Path(original_name).suffix.lower() or ".bin"
    path = f"{project_id}/attach_{uuid.uuid4().hex}{suffix}"
    mime = XLSX_MIME if suffix == ".xlsx" else HWPX_MIME if suffix == ".hwpx" else None
    _put_object(FORM_BUCKET, path, local_path.read_bytes(), mime)
    return path


def register_form_template(project_id: str, session_id: str, spec: dict, file_url: str) -> str:
    """등록된 외부 양식을 기록한다 (F6-5).

    F1과 달리 workbook_json이 없다 — 원본 파일 자체가 산출물이라 셀 단위 재현 명세가
    없는 것이 정상이다(인지 E1). 그래서 `source`로 구분한다(20260804060000 참조).
    구조 추출 스냅샷은 남기지 않는다 — 원본이 무변형으로 있어 언제든 다시 뽑는다.
    """
    return _insert("form_templates", {
        "project_id": project_id, "intake_session_id": session_id, "spec_json": spec,
        "status": "done", "version": 1, "output_format": "xlsx",
        "source": "recognized_external", "file_url": file_url})["id"]


def save_field_rules(template_id: str, rows: list[dict]) -> None:
    """spec에서 파생한 작성기준을 갈아끼운다(F1-5). 재생성 시 이전 것을 남기지 않는다."""
    _rest("DELETE", "/field_rules", params={"form_template_id": f"eq.{template_id}"})
    for row in rows:
        _insert("field_rules", {"form_template_id": template_id, **row})


def list_form_templates(owner_id: str, limit: int = 100) -> list[dict]:
    """본인이 만든 양식 목록. 취합 화면에서 작성기준을 고르는 재료다 (F1-6)."""
    res = _rest("GET", "/form_templates", params={
        "select": "id,project_id,status,version,spec_json,file_url,source,created_at,"
                  "projects!inner(owner_id,name)",
        "projects.owner_id": f"eq.{owner_id}", "status": "eq.done",
        "order": "created_at.desc", "limit": str(limit)})
    out = []
    for row in res.json():
        spec = row.get("spec_json") or {}
        out.append({
            "id": row["id"], "project_id": row["project_id"], "version": row["version"],
            "created_at": row["created_at"], "file_url": row.get("file_url"),
            "source": row.get("source") or "ai_generated",
            "title": spec.get("form_title") or (row.get("projects") or {}).get("name") or "양식",
            "field_count": len(spec.get("fields") or []),
        })
    return out


def rules_of_template(owner_id: str, template_id: str) -> dict | None:
    """양식의 작성기준을 취합 엔진이 먹는 형태로 돌려준다. 남의 것은 None."""
    rows = _rest("GET", "/form_templates", params={
        "id": f"eq.{template_id}", "select": "spec_json,projects!inner(owner_id)",
        "projects.owner_id": f"eq.{owner_id}"}).json()
    if not rows:
        return None
    import formgen as fg
    return fg.rules_for_aggregate(rows[0].get("spec_json") or {})


def set_ai_consent(user_id: str) -> None:
    _rest("PATCH", "/profiles", params={"id": f"eq.{user_id}"}, json={"ai_consent_at": "now()"})


def ai_consented(user_id: str) -> bool:
    rows = _rest("GET", "/profiles",
                 params={"id": f"eq.{user_id}", "select": "ai_consent_at"}).json()
    return bool(rows and rows[0].get("ai_consent_at"))
