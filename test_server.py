#!/usr/bin/env python3
"""server.py 자체 점검. 실행: python3 test_server.py (프레임워크 없음)"""

import io
import os
import sys
import time

import openpyxl
from fastapi import HTTPException
from fastapi.testclient import TestClient

# 취합 API 시나리오는 Supabase 없이 돌아야 하므로 인증을 명시적으로 끈다.
# 인증 자체는 test_auth.py에서 실제 Supabase를 상대로 검증한다.
os.environ["AUTH_DISABLED"] = "1"

import server  # noqa: E402

client = TestClient(server.app)


def book_bytes(rows: list[list]) -> bytes:
    """1행 제목 → 2행 공백 → 3행 헤더 구조의 xlsx를 메모리로 만든다 (§2.2)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "예산"
    ws["A1"] = "실적 보고"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.append([None, None, None])
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


CLEAN = [["사번", "부서", "예산액"], ["A1", "기획부", 100], ["A2", "기획부", 200]]
BROKEN = [["사번", "부서", "예산액"], ["B1", "총무부", None]]      # 필수값 누락 → 오류
# 첫 컬럼이 행마다 반복 + 키워드 헤더 없음 → 자동 추정이 키를 부서로 잡아 오탐
REPEATED = [["부서", "성명", "예산액"],
            ["기획부", "김하나", 100],
            ["기획부", "이두리", 200]]
# 비고는 실무에서 대개 비워둔다 → 지정이 없으면 전 행이 필수값 누락으로 잡힌다
WITH_BLANK_NOTE = [["사번", "부서", "예산액", "비고"],
                   ["A1", "기획부", 100, None],
                   ["A2", "기획부", 200, None]]


def upload(sid: str, items: list[tuple[str, bytes]]):
    return client.post(f"/api/session/{sid}/files",
                       files=[("files", (n, d, "application/octet-stream")) for n, d in items])


def wait(res, timeout: float = 60.0):
    """잡 기반 엔드포인트: job_id를 받아 완료까지 폴링하고 결과만 돌려준다."""
    assert res.status_code == 200, res.text
    jid = res.json()["job_id"]
    deadline = time.time() + timeout
    seen = []
    while time.time() < deadline:
        job = client.get(f"/api/job/{jid}").json()
        seen.append(job["phase"])
        if job["status"] == "done":
            return job["result"]
        if job["status"] == "error":
            raise AssertionError(f"잡 실패: {job['error']}")
        time.sleep(0.005)
    raise AssertionError(f"잡이 {timeout}초 안에 끝나지 않았다. 단계: {seen[-3:]}")


def new_session() -> str:
    res = client.post("/api/session")
    assert res.status_code == 200, res.text
    return res.json()["sid"]


def hwpx_bytes(tag: str) -> bytes:
    """hwpx fixture는 test_hwpx.py 것을 그대로 쓴다(구조를 두 곳에서 관리하지 않는다)."""
    import tempfile
    from pathlib import Path

    import test_hwpx as th
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / f"{tag}.hwpx"
        th._fixture(p, tag, b"\x89PNG-" + tag.encode())
        return p.read_bytes()


def hwpx_session() -> str:
    res = client.post("/api/session", json={"kind": "hwpx"})
    assert res.status_code == 200 and res.json()["kind"] == "hwpx", res.text
    return res.json()["sid"]


def test_form_requires_store():
    """양식 생성은 기록이 남아야 의미가 있다 — Supabase 없이는 503으로 닫는다.

    LLM 호출은 하지 않는다(설정 검사에서 먼저 막힌다). AUTH_DISABLED로 도는 이 파일에서는
    user['id']가 없으므로 여기까지만 확인하고, 실제 흐름은 test_auth.py가 검증한다.
    """
    res = client.post("/api/form/session", json={"message": "예산 양식 만들어줘"})
    assert res.status_code == 503, res.text
    assert "Supabase" in res.json()["detail"], res.text

    # 빈 요청은 LLM을 부르기 전에 400으로 거른다
    assert client.post("/api/form/session", json={"message": "  "}).status_code == 400
    # 없는 문답은 존재를 알리지 않는다
    assert client.post("/api/form/없는세션/messages", json={"message": "x"}).status_code == 404
    assert client.get("/api/form/없는세션").status_code == 404
    assert client.get("/api/form/template/없는것/download").status_code == 404
    # F6 등록도 같은 전제를 쓴다 — 기록이 없으면 등록 자체가 성립하지 않는다
    att = client.post("/api/form/attachments",
                      files=[("files", ("양식.xlsx", b"PK\x03\x04", server.XLSX_MIME))])
    assert att.status_code == 503 and "Supabase" in att.json()["detail"], att.text
    assert client.post("/api/form/없는세션/register").status_code == 404
    assert client.post("/api/form/없는세션/revise", json={"message": "빼줘"}).status_code == 404
    # 인증이 꺼져 있으면 동의는 통과 상태로 본다(로컬 개발)
    assert client.get("/api/form/consent").json() == {"consented": True}
    assert client.get("/api/form/templates").json() == {"templates": []}
    print("  ✓ 양식 생성 사전 조건(Supabase 미설정 503 / 빈 요청 400 / 없는 문답 404)")


def test_hwpx_end_to_end():
    """한글 병합: 업로드 → 드래그 순서(역순)로 병합 → hwpx 다운로드."""
    import zipfile

    sid = hwpx_session()
    files = wait(client.post(f"/api/hwpx/session/{sid}/files",
                             files=[("files", ("가.hwpx", hwpx_bytes("A"), "application/octet-stream")),
                                    ("files", ("나.hwpx", hwpx_bytes("B"), "application/octet-stream"))
                                    ]))["files"]
    assert [f["fid"] for f in files] == [0, 1], files
    assert all(f["readable"] and f["sections"] == 1 for f in files), files

    # 화면에서 드래그로 뒤집은 순서를 그대로 보낸다
    out = wait(client.post(f"/api/hwpx/session/{sid}/merge", json={"order": [1, 0]}))
    assert out["metrics"] == {"merged": 2, "sections": 2, "size_kb": out["metrics"]["size_kb"]}
    assert out["order"] == ["나.hwpx", "가.hwpx"], out["order"]
    assert any("settings.xml" in n for n in out["notes"]), out["notes"]
    assert out["untouched_refs"] == {"linkListIDRef": 2, "linkListNextIDRef": 2}, out["untouched_refs"]

    res = client.get(f"/api/hwpx/session/{sid}/download")
    assert res.status_code == 200 and res.headers["content-type"] == server.HWPX_MIME
    z = zipfile.ZipFile(io.BytesIO(res.content))
    assert z.read("mimetype") == b"application/hwp+zip"
    assert 'secCnt="2"' in z.read("Contents/header.xml").decode()
    # 지정한 순서가 구역 순서다 — 첫 구역이 두 번째로 올린 파일이어야 한다
    assert "B" in z.read("Contents/section0.xml").decode()
    assert "A" in z.read("Contents/section1.xml").decode()
    print("  ✓ 한글 병합 end-to-end(업로드→지정 순서 병합→hwpx 다운로드)")


def test_hwpx_rejected():
    """확장자·개수 거부와, 열리지 않는 파일의 사유 표시."""
    sid = hwpx_session()
    res = client.post(f"/api/hwpx/session/{sid}/files",
                      files=[("files", ("표.xlsx", book_bytes(CLEAN), "application/octet-stream"))])
    assert res.status_code == 400 and ".hwp" in res.json()["detail"], res.text

    # zip이 아닌 파일은 업로드는 되지만 '열 수 없음'으로 표시되고 병합 대상에서 빠진다
    files = wait(client.post(f"/api/hwpx/session/{sid}/files",
                             files=[("files", ("깨진.hwpx", b"not a zip", "application/octet-stream")),
                                    ("files", ("가.hwpx", hwpx_bytes("A"), "application/octet-stream"))
                                    ]))["files"]
    assert files[0]["readable"] is False and "hwp" in files[0]["reason"], files[0]
    assert files[1]["readable"] is True, files[1]

    # 병합 가능한 파일이 1개뿐 → 400
    res = client.post(f"/api/hwpx/session/{sid}/merge", json={"order": [0, 1]})
    assert res.status_code == 400 and "2개 이상" in res.json()["detail"], res.text

    res = client.get(f"/api/hwpx/session/{sid}/download")
    assert res.status_code == 404, res.text
    print("  ✓ 한글 병합 거부(.xlsx / 열 수 없는 파일 / 대상 2개 미만 / 결과 없음)")


def test_session_sweep():
    """F3-5 임시 파일 자동 정리 — TTL이 지난 세션 폴더는 새 세션이 생길 때 지워진다."""
    sid = hwpx_session()
    session = server.SESSIONS[sid]
    tmpdir = session["dir"]
    assert tmpdir.exists()

    session["created"] -= (server.SESSION_TTL_H * 3600) + 60      # 만료시킨다
    hwpx_session()                                                # 새 세션 → 청소 실행
    assert sid not in server.SESSIONS, "만료 세션이 남아 있다"
    assert not tmpdir.exists(), "임시 폴더가 지워지지 않았다"
    print("  ✓ 만료 세션 임시 폴더 자동 정리")


def test_end_to_end():
    """업로드 → 검토 → 정상 파일만 취합 → 결과 xlsx 다운로드."""
    sid = new_session()
    files = wait(upload(sid, [("기획부.xlsx", book_bytes(CLEAN)),
                              ("총무부.xlsx", book_bytes(BROKEN))]))["files"]
    assert [f["fid"] for f in files] == [0, 1], files
    assert files[0]["readable"] and files[0]["sheets"][0]["header_row"] == 3, files[0]

    data = wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    good, bad = data["files"][0], data["files"][1]
    assert good["status"] == "정상" and good["default_checked"], good
    assert bad["status"] == "이상" and not bad["default_checked"], bad
    assert bad["issues"] and bad["issues"][0]["cell"], bad["issues"]
    assert "예산액" in data["columns"], data["columns"]

    body = wait(client.post(f"/api/session/{sid}/aggregate",
                            json={"mode": "B", "included": [0], "group_map": {}, "summary_cols": []}))
    assert body["metrics"] == {"aggregated": 1, "sheets": 1, "autofixed": 0, "excluded": 1}, body["metrics"]
    assert body["result_sheets"] == ["예산"], body["result_sheets"]
    assert body["excluded_files"][0]["name"] == "총무부.xlsx", body["excluded_files"]
    assert body["forced_notice"] == [], body["forced_notice"]

    res = client.get(f"/api/session/{sid}/download/result")
    assert res.status_code == 200 and res.content[:2] == b"PK", res.status_code
    assert client.get(f"/api/session/{sid}/download/report").content[:2] == b"PK"
    print("  ✓ end-to-end(업로드→검토→취합→다운로드)")


def test_forced_include():
    """오류 파일을 강제로 포함하면 취합은 되고 특이사항 안내가 따라붙는다 (§7, §10.1)."""
    sid = new_session()
    wait(upload(sid, [("기획부.xlsx", book_bytes(CLEAN)), ("총무부.xlsx", book_bytes(BROKEN))]))
    wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    body = wait(client.post(f"/api/session/{sid}/aggregate",
                            json={"mode": "B", "included": [0, 1], "group_map": {}, "summary_cols": []}))
    assert body["metrics"]["aggregated"] == 2, body["metrics"]
    assert body["forced_notice"], "강제 포함한 오류 파일의 안내가 있어야 한다"
    print("  ✓ 강제 포함 시 forced_notice 안내")


def test_upload_rejected():
    """xlsx가 아닌 파일은 업로드 단계에서 거부된다."""
    sid = new_session()
    res = upload(sid, [("메모.txt", b"hello")])
    assert res.status_code == 400, res.status_code
    assert ".xlsx" in res.json()["detail"], res.json()
    assert client.post("/api/session/없는세션/review", json={}).status_code == 404
    print("  ✓ 업로드 거부(.xlsx 아님) + 없는 세션 404")


def test_key_column():
    """자동 판정은 고유 컬럼을 키로 잡고, key_col 지정이 그것을 덮어쓴다 (F2-13)."""
    sid = new_session()
    files = wait(upload(sid, [("영업부.xlsx", book_bytes(REPEATED))]))["files"]
    assert files[0]["sheets"][0]["headers"] == ["부서", "성명", "예산액"], files

    # 첫 컬럼(부서)이 행마다 반복되지만, 자동 판정이 고유한 '성명'을 키로 잡아
    # 충돌 오탐이 나지 않는다. 종전에는 첫 컬럼을 그냥 키로 삼아 '이상'이었다
    data = wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    assert data["files"][0]["status"] == "정상", data["files"][0]

    # 사용자가 반복되는 컬럼을 키로 지정하면 그 판단이 우선한다 → 충돌로 잡힌다
    data = wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True, "key_col": "부서"}))
    bad = data["files"][0]
    assert bad["status"] == "이상", bad
    assert any("충돌" in i["message"] or "키" in i["message"] for i in bad["issues"]), bad["issues"]
    print("  ✓ 키 컬럼 자동 판정(고유 컬럼) + key_col 지정이 우선")


def test_optional_columns():
    """비고처럼 비워두는 칸 때문에 파일 전체가 '이상'이 되지 않아야 한다."""
    sid = new_session()
    wait(upload(sid, [("기획부.xlsx", book_bytes(WITH_BLANK_NOTE))]))

    # 지정 없이는 엔진 기본값(전 컬럼 필수) → 빈 비고가 필수값 누락으로 잡힌다
    data = wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))["files"][0]
    assert data["status"] == "이상", data
    assert any("누락" in i["message"] and i["column"] == "비고" for i in data["issues"]), data["issues"]
    assert data["default_checked"] is False, "오류 등급이라 기본 해제돼야 한다"

    # 비고를 선택 입력으로 지정하면 정상으로 돌아오고 취합 대상에 포함된다
    data = wait(client.post(f"/api/session/{sid}/review",
                            json={"no_ai": True, "optional_cols": ["비고"]}))["files"][0]
    assert data["status"] == "정상", data
    assert data["default_checked"] is True, data
    # 다른 컬럼의 진짜 누락까지 덮어버리면 안 된다
    sid2 = new_session()
    wait(upload(sid2, [("총무부.xlsx", book_bytes(BROKEN))]))
    data = wait(client.post(f"/api/session/{sid2}/review",
                            json={"no_ai": True, "optional_cols": ["비고"]}))["files"][0]
    assert data["status"] == "이상", data
    print("  ✓ optional_cols 지정으로 선택 기재 칸 오탐 해소")


def test_unreadable_stays_bad():
    """읽지 못한 파일은 검토에서 '정상'이 될 수 없다.

    재검토 초기화가 read_file이 남긴 실패 사유까지 지우고, 1단계는 읽을 수 없는 파일을
    건너뛰기 때문에 이슈 0건 → '정상'으로 표시됐다. 브라우저에서 손상 파일을 올려보고서야
    드러났다 — 업로드 단계는 사유를 정확히 보여주는데 검토 화면만 '정상'이었다.
    """
    sid = new_session()
    broken = b"PK\x03\x04" + b"\x00" * 200          # zip 머리만 있고 내용이 깨진 파일
    files = wait(upload(sid, [("깨진.xlsx", broken), ("기획부.xlsx", book_bytes(CLEAN))]))["files"]
    assert files[0]["readable"] is False and files[0]["reason"], files[0]

    # 첫 검토와 재검토 모두 사유를 들고 '이상'이어야 하고, 이슈가 누적되지도 않아야 한다
    for turn in (1, 2):
        data = wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))["files"]
        bad = data[0]
        assert bad["status"] == "이상", (turn, bad)
        assert len(bad["issues"]) == 1, (turn, bad["issues"])
        assert "열 수 없습니다" in bad["issues"][0]["message"], bad["issues"]
        assert bad["selectable"] is False and bad["default_checked"] is False, bad
        assert data[1]["status"] == "정상", (turn, data[1])
    print("  ✓ 읽지 못한 파일은 검토에서도 '이상'(재검토에도 사유 유지)")


def test_dept_names():
    """부서명은 파일명에서 후보를 제시하고, 사용자가 정한 값이 결과에 쓰인다 (§9)."""
    sid = new_session()
    names = ["2025년 점검 리스트(대구).xlsx",
             "2025년 점검 리스트(양식)_판교지사.xlsx"]
    files = wait(upload(sid, [(n, book_bytes(CLEAN)) for n in names]))["files"]

    # 같이 올라온 파일명의 공통어('점검', '리스트')는 제목이므로 뒤로 밀린다
    assert files[0]["dept_candidates"][0] == "대구", files[0]["dept_candidates"]
    assert files[1]["dept_candidates"][0] == "판교지사", files[1]["dept_candidates"]
    # 기본값은 종전대로 파일명 전체 — 사용자가 고르기 전까지는 바뀌지 않는다
    assert files[0]["dept"] == "2025년 점검 리스트(대구)", files[0]["dept"]

    # 사용자가 정한 이름이 결과의 '부서' 컬럼에 들어간다
    wait(client.post(f"/api/session/{sid}/review",
                     json={"no_ai": True, "dept_names": {"0": "대구", "1": "판교지사"}}))
    body = wait(client.post(f"/api/session/{sid}/aggregate",
                            json={"mode": "B", "included": [0, 1]}))
    assert body["metrics"]["aggregated"] == 2, body["metrics"]

    res = client.get(f"/api/session/{sid}/download/result")
    wb = openpyxl.load_workbook(io.BytesIO(res.content))
    ws = wb["예산"]
    depts = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    assert depts == {"대구", "판교지사"}, depts

    # 모드 A는 시트명에도 쓰인다
    wait(client.post(f"/api/session/{sid}/review",
                     json={"no_ai": True, "dept_names": {"0": "대구", "1": "판교지사"}}))
    wait(client.post(f"/api/session/{sid}/aggregate", json={"mode": "A", "included": [0, 1]}))
    wb = openpyxl.load_workbook(io.BytesIO(
        client.get(f"/api/session/{sid}/download/result").content))
    assert sorted(wb.sheetnames) == ["대구_예산", "판교지사_예산"], wb.sheetnames
    print("  ✓ 부서명 후보 추천 + 사용자 지정이 결과에 반영")


def test_job_progress():
    """긴 작업은 즉시 job_id를 돌려주고, 폴링으로 단계·진행률·오류가 전달돼야 한다."""
    sid = new_session()
    res = upload(sid, [(f"부서{i}.xlsx", book_bytes(CLEAN)) for i in range(4)])
    assert res.status_code == 200, res.text
    jid = res.json()["job_id"]
    assert "files" not in res.json(), "업로드는 결과를 바로 주지 않고 잡으로 넘겨야 한다"

    job, phases = None, []
    for _ in range(2000):
        job = client.get(f"/api/job/{jid}").json()
        phases.append((job["phase"], job["done"], job["total"]))
        if job["status"] != "running":
            break
        time.sleep(0.002)
    assert job["status"] == "done", job
    assert job["total"] == 4, job
    assert any("구조를 읽는 중" in p for p, _, _ in phases), phases[:5]
    assert len(job["result"]["files"]) == 4, job["result"]

    # 취합 단계도 같은 계약을 따른다
    wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    res = client.post(f"/api/session/{sid}/aggregate",
                      json={"mode": "B", "included": [0, 1, 2, 3]})
    assert "job_id" in res.json(), res.json()
    assert wait(res)["metrics"]["aggregated"] == 4

    # 잡 안에서 터진 예외는 500이 아니라 error 상태로 전달된다
    res = client.post(f"/api/session/{sid}/aggregate", json={"mode": "B", "included": [99]})
    assert res.status_code == 400, res.status_code      # 선택 0건은 즉시 거부
    assert client.get("/api/job/없는잡").status_code == 404
    print("  ✓ 잡 진행률 폴링(단계·진행률·완료·404)")


def pivot_bytes(sheet_name: str, base: int) -> bytes:
    """월×지표 총괄표. 실제 양식처럼 **B열부터** 시작하고 시트명이 파일마다 다르다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["B1"] = f"{sheet_name} 연간 집계"
    ws["B3"], ws["C3"], ws["D3"] = "구분", "신고", "조치"
    for i, month in enumerate(("1월", "2월")):
        r = 4 + i
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = month, base + i, base + i + 10
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_pivot_sheets():
    """총괄표류: 시트명이 파일마다 달라도 피벗으로 지정하면 한 시트 지사별 1행이 된다."""
    sid = new_session()
    files = wait(upload(sid, [("강남지사.xlsx", pivot_bytes("강남지사 총괄표", 1)),
                              ("판교지사.xlsx", pivot_bytes("판교지사 총괄표", 5))]))["files"]
    names = sorted(s["name"] for f in files for s in f["sheets"])
    assert names == ["강남지사 총괄표", "판교지사 총괄표"], names

    # 피벗 지정 없이 취합하면 시트가 파일 수만큼 갈린다(종전 동작이 그대로임을 확인)
    wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    plain = wait(client.post(f"/api/session/{sid}/aggregate",
                             json={"mode": "B", "included": [0, 1]}))
    assert len(plain["result_sheets"]) == 2, plain["result_sheets"]

    # 피벗으로 지정하면 한 시트로 모인다. 표 선택이 바뀌면 원본에서 다시 읽는다
    out = wait(client.post(f"/api/session/{sid}/review", json={
        "no_ai": True, "sheets": ["강남지사 총괄표", "판교지사 총괄표"],
        "pivot_sheets": ["강남지사 총괄표", "판교지사 총괄표"]}))
    # 컬럼 이름이 `행 라벨 + 원래 컬럼명`으로 바뀐다(원래 컬럼명만으로는 남지 않는다)
    assert "1월 신고" in out["columns"] and "2월 조치" in out["columns"], out["columns"]
    assert "신고" not in out["columns"] and "구분" not in out["columns"], out["columns"]
    assert all(f["status"] == "정상" for f in out["files"]), out["files"]

    merged = wait(client.post(f"/api/session/{sid}/aggregate",
                              json={"mode": "B", "included": [0, 1]}))
    assert merged["result_sheets"] == ["총괄표(전개)"], merged["result_sheets"]

    dl = client.get(f"/api/session/{sid}/download/result")
    assert dl.status_code == 200, dl.status_code
    ws = openpyxl.load_workbook(io.BytesIO(dl.content))["총괄표(전개)"]
    assert ws.max_row == 3 and ws.max_column == 5, (ws.max_row, ws.max_column)
    assert [c.value for c in ws[1]][:2] == ["부서", "1월 신고"], [c.value for c in ws[1]]
    assert [ws.cell(row=r, column=2).value for r in (2, 3)] == [1, 5]

    # 되돌릴 수 있어야 한다 — 지정을 비우면 다시 시트가 갈린다
    back = wait(client.post(f"/api/session/{sid}/review",
                            json={"no_ai": True, "pivot_sheets": []}))
    assert "신고" in back["columns"] and "1월 신고" not in back["columns"], back["columns"]
    print("  ✓ 총괄표류 피벗 전개(시트명 통합·1행·되돌리기)")


def test_clock_skew_tolerance():
    """발급자와 검증자의 시계가 어긋나도 방금 발급된 토큰을 받아야 한다.

    실측에서 이 머신이 Supabase보다 3~4초 느려 로그인 직후 모든 요청이 401이 됐다
    (`iat`가 미래라 PyJWT가 `ImmatureSignatureError`). 사용자에게는 "로그인은 됐는데
    아무것도 안 된다"로 보이고, 시계가 다시 맞으면 사라져 원인을 찾기 어렵다.

    시계가 잘 맞는 머신에서는 이 결함이 드러나지 않으므로, 미래 `iat` 토큰을 직접 만들어
    검증 경로에 태운다(네트워크 없이 돈다 — 우리 키로 서명하고 JWKS만 바꿔 끼운다).
    """
    import time
    import jwt
    from cryptography.hazmat.primitives.asymmetric import ec

    import auth

    key = ec.generate_private_key(ec.SECP256R1())
    now = int(time.time())

    def make(iat_offset: int) -> str:
        return jwt.encode({"sub": "11111111-1111-1111-1111-111111111111",
                           "aud": "authenticated", "iat": now + iat_offset,
                           "exp": now + 3600}, key, algorithm="ES256")

    real = auth._jwk_client
    auth._jwk_client = lambda: type("K", (), {
        "get_signing_key_from_jwt": staticmethod(
            lambda tok: type("S", (), {"key": key.public_key()})())})()
    try:
        # 발급 시각이 우리 시계보다 앞서 있어도(오차 범위 안) 통과해야 한다
        claims = auth.verify_token(make(+5))
        assert claims["sub"].startswith("1111"), claims
        assert auth.CLOCK_SKEW_SEC >= 30, auth.CLOCK_SKEW_SEC
        # 오차 범위를 크게 벗어난 토큰은 여전히 거부한다(여유가 검증을 무력화하면 안 된다)
        try:
            auth.verify_token(make(auth.CLOCK_SKEW_SEC + 600))
            raise AssertionError("한참 미래의 토큰이 통과했다")
        except HTTPException as exc:
            assert exc.status_code == 401, exc
    finally:
        auth._jwk_client = real
    print("  ✓ 시계 오차 허용(방금 발급된 토큰 통과 / 한참 미래 토큰은 거부)")


def main() -> int:
    # 업로드 파일은 서버가 세션별 임시 폴더에 두므로 여기서 따로 만들 것이 없다
    for fn in (test_end_to_end, test_forced_include, test_upload_rejected, test_key_column,
               test_optional_columns, test_unreadable_stays_bad, test_dept_names,
               test_job_progress,
               test_hwpx_end_to_end, test_hwpx_rejected, test_session_sweep,
               test_form_requires_store, test_pivot_sheets,
               test_clock_skew_tolerance):
        fn()
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
