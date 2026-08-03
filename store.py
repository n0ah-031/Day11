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
def put_upload(project_id: str, uploader_id: str, local_path: Path, original_name: str) -> dict:
    """파일을 uploads 버킷에 올리고 uploaded_files 행을 남긴다.

    Storage 키에는 원본 파일명을 쓰지 않는다 — 한글·공백이 섞인 이름이 키로
    들어가면 인코딩 문제가 생기고, 원본명은 DB에 그대로 보관하면 충분하다.
    """
    file_id = str(uuid.uuid4())
    storage_path = f"{project_id}/{file_id}.xlsx"
    _put_object(UPLOAD_BUCKET, storage_path, local_path.read_bytes(), XLSX_MIME)
    row = _insert("uploaded_files", {
        "id": file_id, "project_id": project_id, "uploader_id": uploader_id,
        "storage_path": storage_path, "original_name": original_name,
        "kind": "excel", "status": "uploaded",
    })
    return row


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
               summary_fields: list[str]) -> str:
    return _insert("aggregation_jobs", {
        "project_id": project_id, "kind": "excel", "mode": mode, "status": "running",
        "included_file_ids": included_file_ids, "summary_fields_json": summary_fields,
        "progress": 0, "stats_json": {},
    })["id"]


def update_job(job_id: str, **fields) -> None:
    if fields:
        _rest("PATCH", "/aggregation_jobs", params={"id": f"eq.{job_id}"}, json=fields)


def put_result(project_id: str, job_id: str, local_path: Path, label: str) -> str:
    """취합 결과·오류 리포트를 results 버킷에 올리고 경로를 돌려준다."""
    path = f"{project_id}/{job_id}_{label}.xlsx"
    _put_object(RESULT_BUCKET, path, local_path.read_bytes(), XLSX_MIME)
    return path


# ── 이력 (F4-2의 재료) ────────────────────────────────────────────────────────
def list_projects(owner_id: str, limit: int = 50) -> list[dict]:
    res = _rest("GET", "/projects", params={
        "owner_id": f"eq.{owner_id}", "select": "id,name,created_at",
        "order": "created_at.desc", "limit": str(limit)})
    return res.json()


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
