#!/usr/bin/env python3
"""server.py 자체 점검. 실행: python3 test_server.py (프레임워크 없음)"""

import io
import sys

import openpyxl
from fastapi.testclient import TestClient

import server

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


def new_session() -> str:
    res = client.post("/api/session")
    assert res.status_code == 200, res.text
    return res.json()["sid"]


def test_end_to_end():
    """업로드 → 검토 → 정상 파일만 취합 → 결과 xlsx 다운로드."""
    sid = new_session()
    res = upload(sid, [("기획부.xlsx", book_bytes(CLEAN)), ("총무부.xlsx", book_bytes(BROKEN))])
    assert res.status_code == 200, res.text
    files = res.json()["files"]
    assert [f["fid"] for f in files] == [0, 1], files
    assert files[0]["readable"] and files[0]["sheets"][0]["header_row"] == 3, files[0]

    res = client.post(f"/api/session/{sid}/review", json={"no_ai": True})
    assert res.status_code == 200, res.text
    data = res.json()
    good, bad = data["files"][0], data["files"][1]
    assert good["status"] == "정상" and good["default_checked"], good
    assert bad["status"] == "이상" and not bad["default_checked"], bad
    assert bad["issues"] and bad["issues"][0]["cell"], bad["issues"]
    assert "예산액" in data["columns"], data["columns"]

    res = client.post(f"/api/session/{sid}/aggregate",
                      json={"mode": "B", "included": [0], "group_map": {}, "summary_cols": []})
    assert res.status_code == 200, res.text
    body = res.json()
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
    assert upload(sid, [("기획부.xlsx", book_bytes(CLEAN)),
                        ("총무부.xlsx", book_bytes(BROKEN))]).status_code == 200
    assert client.post(f"/api/session/{sid}/review", json={"no_ai": True}).status_code == 200
    res = client.post(f"/api/session/{sid}/aggregate",
                      json={"mode": "B", "included": [0, 1], "group_map": {}, "summary_cols": []})
    assert res.status_code == 200, res.text
    body = res.json()
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
    """첫 컬럼이 반복되면 키 충돌로 오탐 → key_col로 고유 컬럼을 지정하면 정상."""
    sid = new_session()
    res = upload(sid, [("영업부.xlsx", book_bytes(REPEATED))])
    assert res.status_code == 200, res.text
    assert res.json()["files"][0]["sheets"][0]["headers"] == ["부서", "성명", "예산액"], res.json()

    res = client.post(f"/api/session/{sid}/review", json={"no_ai": True})
    assert res.json()["files"][0]["status"] == "이상", res.json()["files"][0]

    res = client.post(f"/api/session/{sid}/review", json={"no_ai": True, "key_col": "성명"})
    assert res.json()["files"][0]["status"] == "정상", res.json()["files"][0]
    print("  ✓ key_col 지정으로 키 충돌 오탐 해소")


def test_optional_columns():
    """비고처럼 비워두는 칸 때문에 파일 전체가 '이상'이 되지 않아야 한다."""
    sid = new_session()
    res = upload(sid, [("기획부.xlsx", book_bytes(WITH_BLANK_NOTE))])
    assert res.status_code == 200, res.text

    # 지정 없이는 엔진 기본값(전 컬럼 필수) → 빈 비고가 필수값 누락으로 잡힌다
    data = client.post(f"/api/session/{sid}/review", json={"no_ai": True}).json()["files"][0]
    assert data["status"] == "이상", data
    assert any("누락" in i["message"] and i["column"] == "비고" for i in data["issues"]), data["issues"]
    assert data["default_checked"] is False, "오류 등급이라 기본 해제돼야 한다"

    # 비고를 선택 입력으로 지정하면 정상으로 돌아오고 취합 대상에 포함된다
    data = client.post(f"/api/session/{sid}/review",
                       json={"no_ai": True, "optional_cols": ["비고"]}).json()["files"][0]
    assert data["status"] == "정상", data
    assert data["default_checked"] is True, data
    # 다른 컬럼의 진짜 누락까지 덮어버리면 안 된다
    sid2 = new_session()
    upload(sid2, [("총무부.xlsx", book_bytes(BROKEN))])
    data = client.post(f"/api/session/{sid2}/review",
                       json={"no_ai": True, "optional_cols": ["비고"]}).json()["files"][0]
    assert data["status"] == "이상", data
    print("  ✓ optional_cols 지정으로 선택 기재 칸 오탐 해소")


def main() -> int:
    # 업로드 파일은 서버가 세션별 임시 폴더에 두므로 여기서 따로 만들 것이 없다
    for fn in (test_end_to_end, test_forced_include, test_upload_rejected, test_key_column,
               test_optional_columns):
        fn()
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
