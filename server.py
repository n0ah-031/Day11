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

from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import aggregate as ag

BASE = Path(__file__).parent
MAX_FILE_MB = 50            # 파일 1개 상한
MAX_SESSION_FILES = 30      # 세션 누적 파일 수 상한
MAX_SESSION_MB = 500        # 세션 누적 용량 상한

ag.load_env(BASE / ".env")
app = FastAPI(title="엑셀 취합")

# 세션: sid → {"dir": 임시폴더, "files": [UploadedFile], "bytes": 누적 바이트}
SESSIONS: dict[str, dict] = {}
# 잡: jid → 진행 상황. 세션과 마찬가지로 프로세스 메모리이며 재시작하면 사라진다
JOBS: dict[str, dict] = {}


def _session(sid: str) -> dict:
    if sid not in SESSIONS:
        raise HTTPException(404, "세션을 찾을 수 없습니다. 페이지를 새로고침 후 다시 시도해주세요.")
    return SESSIONS[sid]


def _start_job(fn) -> dict:
    """긴 작업을 백그라운드 스레드로 돌리고 즉시 job_id를 돌려준다.

    명세 규모(30개/10만 행)에서 읽기·합성이 수십 초라 동기 응답으로는
    진행 상황을 알릴 방법이 없다. fn은 report(phase, done, total)를 받아
    단계를 남기고, 클라이언트는 /api/job/{jid}를 폴링한다.
    """
    jid = uuid.uuid4().hex
    job = JOBS[jid] = {"status": "running", "phase": "준비 중", "done": 0, "total": 0,
                       "result": None, "error": None}

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


@app.get("/api/job/{jid}")
def job_status(jid: str) -> dict:
    job = JOBS.get(jid)
    if job is None:
        raise HTTPException(404, "작업을 찾을 수 없습니다. 페이지를 새로고침 후 다시 시도해주세요.")
    return job


def _file_view(fid: int, uf) -> dict:
    return {
        "fid": fid,
        "name": uf.name,
        "readable": uf.readable,
        # §2.1 읽기 실패 사유는 read_file이 남긴 첫 이슈의 사유 그대로
        "reason": (uf.issues[0].reason if not uf.readable and uf.issues else None),
        # headers는 UI의 키 컬럼 선택지 재료
        "sheets": [{"name": s.name, "header_row": s.header_row, "rows": len(s.rows),
                    "cols": len(s.headers), "images": len(s.images), "headers": s.headers}
                   for s in uf.sheets],
    }


@app.post("/api/session")
def create_session() -> dict:
    sid = uuid.uuid4().hex
    SESSIONS[sid] = {"dir": Path(tempfile.mkdtemp(prefix="agg-")), "files": [], "bytes": 0}
    return {"sid": sid}


@app.post("/api/session/{sid}/files")
async def upload_files(sid: str, files: list[UploadFile] = File(...)) -> dict:
    """저장까지만 동기로 하고, 느린 구조 인식은 잡으로 넘긴다.

    거부 사유(확장자·용량·개수)는 즉시 400으로 돌려줘야 하므로 여기서 검사한다.
    """
    session = _session(sid)
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
            session["files"].append(ag.read_file(path))
        report("구조 인식 완료", len(pending), len(pending))
        # fid는 세션 내 인덱스이므로 누적된 전체 목록을 돌려준다
        return {"files": [_file_view(i, uf) for i, uf in enumerate(session["files"])]}

    return _start_job(work)


@app.post("/api/session/{sid}/review")
def review(sid: str, body: dict = Body(default={})) -> dict:
    session = _session(sid)
    no_ai = bool((body or {}).get("no_ai"))
    model = os.environ.get("OPENAI_MODEL", "gpt-5-mini")
    # 키 컬럼·선택 입력 컬럼을 사용자가 고르면 엔진의 rules 경로로 넘긴다
    key_col = (body or {}).get("key_col")
    rules: dict = {key_col: {"key": True}} if key_col else {}
    # 비고·특이사항처럼 비워둘 수 있는 칸. 지정이 없으면 엔진 기본값(전 컬럼 필수)
    for col in (body or {}).get("optional_cols") or []:
        rules.setdefault(col, {})["required"] = False

    def work(report):
        total = len(session["files"])
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

        out, columns = [], []
        for fid, uf in enumerate(session["files"]):
            report(f"{uf.name} 검토 중", fid, total)
            if uf.readable:
                ag.review_stage1(uf, rules)
                # 파일 단위 게이팅: 1단계 위반이 하나라도 있으면 2단계로 진입하지 않는다 (§4.2)
                if not uf.issues and not no_ai:
                    ag.review_stage2(uf, model)
            for sheet in uf.sheets:
                columns += [h for h in sheet.headers if h]
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

    return _start_job(work)


@app.post("/api/session/{sid}/aggregate")
def aggregate(sid: str, body: dict = Body(default={})) -> dict:
    session = _session(sid)
    files = session["files"]
    included = set(body.get("included") or [])
    # included는 사용자의 최종 선택이므로 등급으로 재필터링하지 않는다 (§7 강제 포함)
    picked = [i for i, uf in enumerate(files) if i in included and uf.readable and uf.sheets]
    selected = [files[i] for i in picked]

    if not selected:
        raise HTTPException(400, "취합 가능한 파일이 없습니다.")

    def work(report_fn):
        steps = 4
        report_fn("오류 리포트를 쓰는 중", 0, steps)
        report_path = session["dir"] / "error_report.xlsx"
        ag.write_report(files, report_path)
        session["report"] = report_path

        report_fn("자동교정하는 중", 1, steps)
        for uf in selected:
            ag.preprocess(uf)
        session["preprocessed"] = True      # 셀 값이 바뀌었으니 재검토 시 원본을 다시 읽어야 한다

        report_fn("시트를 합치는 중", 2, steps)
        wb, notes = ag.synthesize(selected, body.get("mode", "B"),
                                  body.get("group_map") or {}, body.get("summary_cols") or [])

        report_fn("결과 파일을 저장하는 중", 3, steps)
        result = session["dir"] / "merged.xlsx"
        wb.save(result)
        session["result"] = result
        report_fn("취합 완료", steps, steps)

        # Fix에는 컬럼명이 없으므로 같은 셀의 이슈에서 작성기준을 끌어온다
        attr_of = {(i.file, i.sheet, i.cell): i.attr for uf in selected for i in uf.issues}
        return {
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

    return _start_job(work)


@app.get("/api/session/{sid}/download/{kind}")
def download(sid: str, kind: str):
    session = _session(sid)
    path = session.get("result" if kind == "result" else "report")
    if not path or not Path(path).exists():
        raise HTTPException(404, "아직 생성되지 않은 파일입니다.")
    return FileResponse(path, filename=Path(path).name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# 정적 UI는 반드시 마지막에 마운트한다(먼저 걸면 /api 라우트가 가려진다)
app.mount("/", StaticFiles(directory=BASE / "ui", html=True))
