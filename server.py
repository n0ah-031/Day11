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


def _session(sid: str) -> dict:
    if sid not in SESSIONS:
        raise HTTPException(404, "세션을 찾을 수 없습니다. 페이지를 새로고침 후 다시 시도해주세요.")
    return SESSIONS[sid]


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
    session = _session(sid)
    for up in files:
        name = Path(up.filename or "").name
        if not name.lower().endswith(".xlsx"):
            raise HTTPException(400, f"'{name}'은(는) .xlsx 파일이 아닙니다. "
                                     "엑셀에서 'Excel 통합 문서(*.xlsx)'로 저장 후 재시도해주세요.")
        data = await up.read()
        if len(data) > MAX_FILE_MB * 1024 * 1024:
            raise HTTPException(400, f"'{name}'의 용량이 {MAX_FILE_MB}MB를 초과합니다.")
        if len(session["files"]) + 1 > MAX_SESSION_FILES:
            raise HTTPException(400, f"업로드 가능한 파일은 최대 {MAX_SESSION_FILES}개입니다.")
        if session["bytes"] + len(data) > MAX_SESSION_MB * 1024 * 1024:
            raise HTTPException(400, f"업로드 총 용량이 {MAX_SESSION_MB}MB를 초과합니다.")
        path = session["dir"] / name
        path.write_bytes(data)
        session["bytes"] += len(data)
        session["files"].append(ag.read_file(path))
    # fid는 세션 내 인덱스이므로 누적된 전체 목록을 돌려준다
    return {"files": [_file_view(i, uf) for i, uf in enumerate(session["files"])]}


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
    # 재검토 시 이슈가 누적되지 않도록 원본에서 다시 읽는다
    session["files"] = [ag.read_file(uf.path) for uf in session["files"]]

    out, columns = [], []
    for fid, uf in enumerate(session["files"]):
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
    return {"files": out, "columns": list(dict.fromkeys(columns))}


@app.post("/api/session/{sid}/aggregate")
def aggregate(sid: str, body: dict = Body(default={})) -> dict:
    session = _session(sid)
    files = session["files"]
    included = set(body.get("included") or [])
    # included는 사용자의 최종 선택이므로 등급으로 재필터링하지 않는다 (§7 강제 포함)
    picked = [i for i, uf in enumerate(files) if i in included and uf.readable and uf.sheets]
    selected = [files[i] for i in picked]

    report = session["dir"] / "error_report.xlsx"
    ag.write_report(files, report)
    session["report"] = report
    if not selected:
        raise HTTPException(400, "취합 가능한 파일이 없습니다.")

    for uf in selected:
        ag.preprocess(uf)
    wb, notes = ag.synthesize(selected, body.get("mode", "B"),
                              body.get("group_map") or {}, body.get("summary_cols") or [])
    result = session["dir"] / "merged.xlsx"
    wb.save(result)
    session["result"] = result

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
