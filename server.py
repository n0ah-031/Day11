#!/usr/bin/env python3
"""F2 엑셀 취합 웹 서버.

aggregate.py(CLI 엔진)를 수정 없이 import해 브라우저 흐름에 얹는다:
  업로드 → 검토(1단계 + 파일 단위 게이팅 후 2단계) → 취합 → 다운로드

세션은 모듈 전역 dict. 서버를 재시작하면 사라지는 것이 정상이다(임시 작업 도구).
실행: uvicorn server:app
"""

from __future__ import annotations

import os
import tempfile
import threading
import uuid
from pathlib import Path

from fastapi import Body, Cookie, Depends, FastAPI, File, HTTPException, Response, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import aggregate as ag
import auth
import store

BASE = Path(__file__).parent
MAX_FILE_MB = 50            # 파일 1개 상한
MAX_SESSION_FILES = 30      # 세션 누적 파일 수 상한
MAX_SESSION_MB = 500        # 세션 누적 용량 상한
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

ag.load_env(BASE / ".env")
app = FastAPI(title="엑셀 취합")


def _auth_disabled() -> bool:
    """명시적으로 끈 경우에만 인증을 건너뛴다.

    키가 없으면 조용히 열리는 것이 아니라 503으로 닫힌다 — 설정 누락이
    접근 제어 구멍이 되지 않도록 fail-closed로 둔다.
    """
    return os.environ.get("AUTH_DISABLED", "").lower() in ("1", "true", "yes")


if _auth_disabled():
    print("경고: AUTH_DISABLED=1 — 인증 없이 실행합니다. 로컬 개발·테스트 전용입니다.")


def require_user(access_token: str | None = Cookie(default=None, alias=auth.ACCESS_COOKIE)) -> dict:
    if _auth_disabled():
        return {"id": None, "employee_no": "(인증 비활성)", "role": "user", "status": "active"}
    return auth.current_user(access_token)


User = Depends(require_user)

# 세션: sid → {"dir": 임시폴더, "files": [UploadedFile], "bytes": 누적 바이트}
SESSIONS: dict[str, dict] = {}
# 잡: jid → 진행 상황. 세션과 마찬가지로 프로세스 메모리이며 재시작하면 사라진다
JOBS: dict[str, dict] = {}


def _apply_dept_names(session: dict) -> None:
    """사용자가 정한 부서명을 파일에 다시 씌운다.

    재검토·취합 때 원본을 다시 읽으면 uf.dept가 파일명으로 돌아가므로,
    파일을 만질 때마다 이걸 통과시킨다.
    """
    names = session.get("dept_names") or {}
    for fid, uf in enumerate(session["files"]):
        chosen = (names.get(str(fid)) or "").strip()
        if chosen:
            uf.dept = chosen


def _session(sid: str, user: dict) -> dict:
    session = SESSIONS.get(sid)
    # 남의 세션은 존재 자체를 알리지 않는다 (§6.2 사용자별 데이터 격리)
    if session is None or session["owner"] != user["id"]:
        raise HTTPException(404, "세션을 찾을 수 없습니다. 페이지를 새로고침 후 다시 시도해주세요.")
    return session


def _start_job(fn, owner) -> dict:
    """긴 작업을 백그라운드 스레드로 돌리고 즉시 job_id를 돌려준다.

    명세 규모(30개/10만 행)에서 읽기·합성이 수십 초라 동기 응답으로는
    진행 상황을 알릴 방법이 없다. fn은 report(phase, done, total)를 받아
    단계를 남기고, 클라이언트는 /api/job/{jid}를 폴링한다.
    """
    jid = uuid.uuid4().hex
    job = JOBS[jid] = {"status": "running", "phase": "준비 중", "done": 0, "total": 0,
                       "result": None, "error": None, "owner": owner}

    def report(phase: str, done: int = 0, total: int = 0) -> None:
        job.update(phase=phase, done=done, total=total)

    def run() -> None:
        try:
            job["result"] = fn(report)
            job.update(status="done", phase="완료")
        except HTTPException as exc:
            job.update(status="error", error=str(exc.detail))
        except Exception as exc:                       # 잡 안에서 죽어도 폴링으로 사유가 전달돼야 한다
            job.update(status="error", error=f"{type(exc).__name__}: {exc}")

    threading.Thread(target=run, daemon=True).start()
    return {"job_id": jid}


@app.post("/api/auth/signup")
def auth_signup(body: dict = Body(...)) -> dict:
    profile = auth.signup(body.get("employee_no", ""), body.get("password", ""),
                          body.get("reset_email", ""))
    return {"profile": profile}


@app.post("/api/auth/login")
def auth_login(response: Response, body: dict = Body(...)) -> dict:
    result = auth.login(body.get("employee_no", ""), body.get("password", ""))
    auth.set_session_cookies(response, result["token"])
    # 토큰은 httpOnly 쿠키로만 나간다. 응답 본문에는 담지 않는다
    return {"profile": result["profile"]}


@app.post("/api/auth/logout")
def auth_logout(response: Response) -> dict:
    auth.clear_session_cookies(response)
    return {"ok": True}


@app.post("/api/auth/reset-password-request")
def auth_reset(body: dict = Body(...)) -> dict:
    auth.request_password_reset(body.get("employee_no", ""))
    # 사번 존재 여부와 무관하게 같은 응답 (계정 열거 방지)
    return {"ok": True}


@app.get("/api/auth/me")
def auth_me(user: dict = User) -> dict:
    return {"profile": user, "auth_enabled": not _auth_disabled()}


def require_admin(user: dict = User) -> dict:
    """§3.2 Admin 전용. 권한이 없으면 403."""
    if not _auth_disabled() and user.get("role") != "admin":
        raise HTTPException(403, "관리자 권한이 필요합니다.")
    if not store.enabled():
        raise HTTPException(503, "Supabase가 설정되지 않아 관리자 기능을 쓸 수 없습니다.")
    return user


Admin = Depends(require_admin)


@app.get("/api/admin/dashboard")
def admin_dashboard(admin: dict = Admin) -> dict:
    return store.dashboard()


@app.get("/api/admin/accounts")
def admin_accounts(admin: dict = Admin) -> dict:
    return {"accounts": store.list_accounts()}


@app.patch("/api/admin/accounts/{user_id}")
def admin_patch_account(user_id: str, body: dict = Body(...), admin: dict = Admin) -> dict:
    if user_id == admin.get("id") and (body.get("role") == "user"
                                       or body.get("status") == "suspended"):
        # 마지막 관리자가 스스로를 잠가버리면 CLI 말고는 되돌릴 방법이 없다
        raise HTTPException(400, "본인 계정의 권한·상태는 이 화면에서 바꿀 수 없습니다.")
    row = store.patch_account(user_id, body)
    if row is None:
        raise HTTPException(404, "계정을 찾을 수 없거나 바꿀 수 있는 항목이 없습니다.")
    changed = ", ".join(f"{k}={v}" for k, v in body.items() if k in ("role", "status"))
    store.log_action(admin.get("id"), "계정 변경", f"{row['employee_no']} ({changed})")
    return {"account": row}


@app.delete("/api/admin/accounts/{user_id}")
def admin_delete_account(user_id: str, admin: dict = Admin) -> dict:
    if user_id == admin.get("id"):
        raise HTTPException(400, "본인 계정은 삭제할 수 없습니다.")
    target = next((a for a in store.list_accounts() if a["id"] == user_id), None)
    if target is None:
        raise HTTPException(404, "계정을 찾을 수 없습니다.")
    store.delete_account(user_id)
    # 계정이 사라지면 누가 지웠는지만 남는다(대상 사번은 문자열로 보존)
    store.log_action(admin.get("id"), "계정 삭제", target["employee_no"])
    return {"deleted": target["employee_no"]}


@app.put("/api/admin/policy/retention")
def admin_set_retention(body: dict = Body(...), admin: dict = Admin) -> dict:
    days = body.get("retention_days")
    if not isinstance(days, int) or not 1 <= days <= 3650:
        raise HTTPException(400, "보관 기간은 1~3650일 사이의 정수로 입력해주세요.")
    policy = store.set_policy("retention_days", days, admin.get("id"))
    store.log_action(admin.get("id"), "정책 변경", f"retention_days={days}")
    return {"policy": policy}


@app.get("/api/admin/logs")
def admin_logs(q: str = "", admin: dict = Admin) -> dict:
    return {"logs": store.list_logs(q=q)}


@app.get("/api/history")
def history(q: str = "", type: str = "", user: dict = User) -> dict:
    """F4-2 이력. 본인 프로젝트만 돌려준다 (§6.2).

    type은 design.md 화면 9의 유형 필터(전체/엑셀/hwpx)에 대응한다.
    hwpx(F3)는 미착수라 지금은 항상 빈 목록이 된다.
    """
    if not store.enabled() or not user["id"]:
        return {"projects": []}
    kind = {"excel": "excel", "hwpx": "hwpx"}.get(type, "")
    return {"projects": store.list_projects(user["id"], q=q, kind=kind)}


@app.get("/api/history/{project_id}")
def history_detail(project_id: str, user: dict = User) -> dict:
    if not store.enabled() or not user["id"]:
        raise HTTPException(404, "이력을 찾을 수 없습니다.")
    detail = store.project_detail(user["id"], project_id)
    if detail is None:
        # 남의 프로젝트는 존재 자체를 알리지 않는다
        raise HTTPException(404, "이력을 찾을 수 없습니다.")
    return detail


@app.get("/api/history/{project_id}/download/{job_id}")
def history_download(project_id: str, job_id: str, user: dict = User):
    """이력에서 과거 취합 결과를 내려받는다. 세션이 사라진 뒤에도 동작한다."""
    detail = history_detail(project_id, user)
    job = next((j for j in detail.get("aggregation_jobs", []) if j["id"] == job_id), None)
    if job is None or not job.get("result_url"):
        raise HTTPException(404, "결과 파일이 없습니다.")
    data = store.get_object(store.RESULT_BUCKET, job["result_url"])
    return Response(content=data, media_type=XLSX_MIME,
                    headers={"Content-Disposition": 'attachment; filename="merged.xlsx"'})


@app.get("/api/job/{jid}")
def job_status(jid: str, user: dict = User) -> dict:
    job = JOBS.get(jid)
    if job is None or job["owner"] != user["id"]:
        raise HTTPException(404, "작업을 찾을 수 없습니다. 페이지를 새로고침 후 다시 시도해주세요.")
    return {k: v for k, v in job.items() if k != "owner"}


def _file_view(fid: int, uf, siblings: list[str] | None = None) -> dict:
    return {
        "fid": fid,
        "name": uf.name,
        "dept": uf.dept,
        # 파일명에서 뽑은 부서명 후보. 사용자가 고르거나 직접 입력한다(§9)
        "dept_candidates": ag.suggest_dept_names(uf.path.stem, siblings),
        "readable": uf.readable,
        # §2.1 읽기 실패 사유는 read_file이 남긴 첫 이슈의 사유 그대로
        "reason": (uf.issues[0].reason if not uf.readable and uf.issues else None),
        # headers는 UI의 키 컬럼 선택지 재료
        "sheets": [{"name": s.name, "header_row": s.header_row, "rows": len(s.rows),
                    "cols": len(s.headers), "images": len(s.images), "headers": s.headers}
                   for s in uf.sheets],
    }


@app.post("/api/session")
def create_session(user: dict = User) -> dict:
    """취합 작업 1건 = 프로젝트 1개. 사용자가 고르는 UI는 아직 없어 자동 생성한다."""
    sid = uuid.uuid4().hex
    session = SESSIONS[sid] = {"dir": Path(tempfile.mkdtemp(prefix="agg-")), "files": [],
                               "bytes": 0, "owner": user["id"], "project_id": None,
                               "file_ids": {}, "dept_names": {}}
    if store.enabled() and user["id"]:
        # 이름에 시각을 넣지 않는다 — 서버 시간대로 굳어버려 DB의 created_at(UTC)을
        # 보는 사람 시간대로 변환한 값과 어긋난다. 시각은 created_at만 쓴다.
        session["project_id"] = store.create_project(user["id"], "엑셀 취합")
    return {"sid": sid, "project_id": session["project_id"]}


@app.post("/api/session/{sid}/files")
async def upload_files(sid: str, files: list[UploadFile] = File(...), user: dict = User) -> dict:
    """저장까지만 동기로 하고, 느린 구조 인식은 잡으로 넘긴다.

    거부 사유(확장자·용량·개수)는 즉시 400으로 돌려줘야 하므로 여기서 검사한다.
    """
    session = _session(sid, user)
    pending: list[Path] = []
    for up in files:
        name = Path(up.filename or "").name
        if not name.lower().endswith(".xlsx"):
            raise HTTPException(400, f"'{name}'은(는) .xlsx 파일이 아닙니다. "
                                     "엑셀에서 'Excel 통합 문서(*.xlsx)'로 저장 후 재시도해주세요.")
        data = await up.read()
        if len(data) > MAX_FILE_MB * 1024 * 1024:
            raise HTTPException(400, f"'{name}'의 용량이 {MAX_FILE_MB}MB를 초과합니다.")
        if len(session["files"]) + len(pending) + 1 > MAX_SESSION_FILES:
            raise HTTPException(400, f"업로드 가능한 파일은 최대 {MAX_SESSION_FILES}개입니다.")
        if session["bytes"] + len(data) > MAX_SESSION_MB * 1024 * 1024:
            raise HTTPException(400, f"업로드 총 용량이 {MAX_SESSION_MB}MB를 초과합니다.")
        path = session["dir"] / name
        path.write_bytes(data)
        session["bytes"] += len(data)
        pending.append(path)

    def work(report):
        for i, path in enumerate(pending):
            report(f"{path.name} 구조를 읽는 중", i, len(pending))
            uf = ag.read_file(path)
            session["files"].append(uf)
            if session["project_id"]:
                # 파일 자체를 Storage에 남겨야 재시작 후에도 원본이 살아 있다
                row = store.put_upload(session["project_id"], session["owner"], path, path.name)
                session["file_ids"][len(session["files"]) - 1] = row["id"]
        report("구조 인식 완료", len(pending), len(pending))
        # 부서명 후보는 같이 올라온 파일명을 함께 봐야 정확해진다(공통어 = 제목)
        stems = [f.path.stem for f in session["files"]]
        _apply_dept_names(session)
        # fid는 세션 내 인덱스이므로 누적된 전체 목록을 돌려준다
        return {"files": [_file_view(i, uf, stems) for i, uf in enumerate(session["files"])]}

    return _start_job(work, user["id"])


@app.post("/api/session/{sid}/review")
def review(sid: str, body: dict = Body(default={}), user: dict = User) -> dict:
    session = _session(sid, user)
    no_ai = bool((body or {}).get("no_ai"))
    model = os.environ.get("OPENAI_MODEL", "gpt-5-mini")
    # 키 컬럼·선택 입력 컬럼을 사용자가 고르면 엔진의 rules 경로로 넘긴다
    key_col = (body or {}).get("key_col")
    rules: dict = {key_col: {"key": True}} if key_col else {}
    # 비고·특이사항처럼 비워둘 수 있는 칸. 지정이 없으면 엔진 기본값(전 컬럼 필수)
    for col in (body or {}).get("optional_cols") or []:
        rules.setdefault(col, {})["required"] = False
    # 부서명은 화면에서 고르거나 직접 입력한 값을 쓴다. 파일명은 기본값일 뿐이다
    if isinstance((body or {}).get("dept_names"), dict):
        session["dept_names"] = {str(k): v for k, v in body["dept_names"].items()}

    def work(report):
        total = len(session["files"])
        _apply_dept_names(session)
        # 재검토 시 판정이 누적되지 않도록 초기화한다. 취합이 한 번이라도 돌았으면
        # preprocess가 셀 값을 고쳐놓았으므로 그때만 원본에서 다시 읽는다
        # (파일당 재파싱이 검토 시간의 절반을 차지한다).
        if session.get("preprocessed"):
            fresh = []
            for i, uf in enumerate(session["files"]):
                report(f"{uf.name} 원본을 다시 읽는 중", i, total)
                fresh.append(ag.read_file(uf.path))
            session["files"] = fresh
            session["preprocessed"] = False
        else:
            for uf in session["files"]:
                uf.issues.clear()
                uf.fixes.clear()
                uf.ai_unverified = False
                for sheet in uf.sheets:
                    sheet.col_types = []

        # 1단계는 전부 로컬 계산이라 빠르다(명세 규모에서 0.2초)
        for fid, uf in enumerate(session["files"]):
            report(f"{uf.name} 규칙 검증 중", fid, total)
            if uf.readable:
                ag.review_stage1(uf, rules)

        # 2단계는 파일당 십수 초 걸리는 API 호출이라 한꺼번에 병렬로 돌린다.
        # 게이팅(§4.2)은 review_stage2_many가 파일 단위로 적용한다.
        if not no_ai:
            gated = [uf for uf in session["files"] if uf.readable and not uf.issues]
            if gated:
                done = 0

                def ai_progress(uf):
                    nonlocal done
                    done += 1
                    report(f"AI 재검증 중 ({done}/{len(gated)})", done, len(gated))

                report("AI 재검증 중", 0, len(gated))
                ag.review_stage2_many(session["files"], model, on_done=ai_progress)

        out, columns = [], []
        for fid, uf in enumerate(session["files"]):
            for sheet in uf.sheets:
                columns += [h for h in sheet.headers if h]
            file_id = session["file_ids"].get(fid)
            if file_id:
                store.save_review(file_id, uf.grade,
                                  [{"sheet": i.sheet, "cell": i.cell, "column": i.attr,
                                    "kind": i.kind, "stage": i.stage, "grade": i.grade,
                                    "reason": i.reason} for i in uf.issues])
            out.append({
                "fid": fid,
                "name": uf.name,
                "status": uf.status,                          # §6.1 정상/이상 이진 표기
                "ai_unverified": uf.ai_unverified,            # §6.4 정상(AI 미검증)
                "default_checked": uf.grade != ag.ERROR,      # §7 내부 등급 오류만 기본 해제
                "selectable": bool(uf.readable and uf.sheets),
                "issues": [{"sheet": i.sheet, "cell": i.cell, "column": i.attr,
                            "message": f"{i.tag} {i.reason}"} for i in uf.issues],
            })
        report("검토 완료", total, total)
        return {"files": out, "columns": list(dict.fromkeys(columns))}

    return _start_job(work, user["id"])


@app.post("/api/session/{sid}/aggregate")
def aggregate(sid: str, body: dict = Body(default={}), user: dict = User) -> dict:
    session = _session(sid, user)
    if isinstance(body.get("dept_names"), dict):
        session["dept_names"] = {str(k): v for k, v in body["dept_names"].items()}
    _apply_dept_names(session)
    files = session["files"]
    included = set(body.get("included") or [])
    # included는 사용자의 최종 선택이므로 등급으로 재필터링하지 않는다 (§7 강제 포함)
    picked = [i for i, uf in enumerate(files) if i in included and uf.readable and uf.sheets]
    selected = [files[i] for i in picked]

    if not selected:
        raise HTTPException(400, "취합 가능한 파일이 없습니다.")

    mode = body.get("mode", "B")
    job_id = None
    if session["project_id"]:
        job_id = store.create_job(session["project_id"], mode,
                                  [session["file_ids"][i] for i in picked if i in session["file_ids"]],
                                  body.get("summary_cols") or [])
        session["job_id"] = job_id

    def work(report_fn):
        steps = 4

        def step(label: str, done: int) -> None:
            report_fn(label, done, steps)
            if job_id:
                store.update_job(job_id, progress=int(done / steps * 100))

        try:
            step("오류 리포트를 쓰는 중", 0)
            report_path = session["dir"] / "error_report.xlsx"
            ag.write_report(files, report_path)
            session["report"] = report_path

            step("자동교정하는 중", 1)
            for uf in selected:
                ag.preprocess(uf)
            session["preprocessed"] = True   # 셀 값이 바뀌었으니 재검토 시 원본을 다시 읽어야 한다

            step("시트를 합치는 중", 2)
            wb, notes = ag.synthesize(selected, mode,
                                      body.get("group_map") or {}, body.get("summary_cols") or [])

            step("결과 파일을 저장하는 중", 3)
            result = session["dir"] / "merged.xlsx"
            wb.save(result)
            session["result"] = result
        except Exception as exc:
            # 실패도 기록에 남겨야 이력에서 무슨 일이 있었는지 알 수 있다
            if job_id:
                store.update_job(job_id, status="failed", error_message=f"{type(exc).__name__}: {exc}")
            raise
        step("취합 완료", steps)

        # Fix에는 컬럼명이 없으므로 같은 셀의 이슈에서 작성기준을 끌어온다
        attr_of = {(i.file, i.sheet, i.cell): i.attr for uf in selected for i in uf.issues}
        payload = {
            "metrics": {
                "aggregated": len(selected),
                "sheets": len(wb.sheetnames),
                "autofixed": sum(len(uf.fixes) for uf in selected),
                "excluded": len(files) - len(selected),
            },
            "result_sheets": wb.sheetnames,
            "excluded_files": [{"name": uf.name, "issue_count": len(uf.issues)}
                               for i, uf in enumerate(files) if i not in set(picked)],
            "autofixes": [{"sheet": f.sheet, "cell": f.cell,
                           "column": attr_of.get((f.file, f.sheet, f.cell), ""),
                           "before": f.original, "after": f.corrected}
                          for uf in selected for f in uf.fixes],
            # §10.1 강제 포함한 오류 등급 파일의 오류 이슈 — 다운로드는 막지 않는다
            "forced_notice": [i.line() for uf in selected if uf.grade == ag.ERROR
                              for i in uf.issues if i.grade == ag.ERROR],
            "notes": list(dict.fromkeys(notes)),
        }
        if job_id:
            result_url = store.put_result(session["project_id"], job_id, result, "merged")
            store.put_result(session["project_id"], job_id, report_path, "report")
            store.update_job(job_id, status="done", progress=100, result_url=result_url,
                             stats_json=payload["metrics"])
            store.log_action(session["owner"], "취합 완료",
                            f"모드 {mode} · {len(selected)}건")
        return payload

    return _start_job(work, user["id"])


@app.get("/api/session/{sid}/download/{kind}")
def download(sid: str, kind: str, user: dict = User):
    session = _session(sid, user)
    path = session.get("result" if kind == "result" else "report")
    if not path or not Path(path).exists():
        raise HTTPException(404, "아직 생성되지 않은 파일입니다.")
    return FileResponse(path, filename=Path(path).name,
                        media_type=XLSX_MIME)


# 정적 UI는 반드시 마지막에 마운트한다(먼저 걸면 /api 라우트가 가려진다)
app.mount("/", StaticFiles(directory=BASE / "ui", html=True))
