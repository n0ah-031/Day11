#!/usr/bin/env python3
"""HANDOFF §5-3 확인 — 서식 상속 규칙 1을 실제 Supabase 상대로 검증한다.

test_server.py는 Supabase 없이 돌아 `_form_store_ready`가 항상 거짓이므로 규칙 1
분기 자체를 실행하지 못한다. 이 스크립트는 인증을 켠 상태로 실제 프로젝트를 상대로
F6 등록 → 취합을 태워 다음 셋을 본다.

  ① 규칙 1  — 등록 양식을 고르면 그 양식의 서식이 쓰인다(회신본 다수결이 아니라)
  ② 규칙 2  — 양식을 고르지 않으면 회신본 공통 서식이 쓰인다(대조군)
  ③ 소유권  — 남의 양식 id를 넣으면 거부되고, 500이 아니라 규칙 2로 내려간다

LLM은 스텁이다(F6 인지 자체는 test_auth가 이미 덮는다) — 유료 호출 0건.
테스트 계정은 `zz-test-` 접두사로 만들고 끝에 지운다.
"""

import json
import os
import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path

REPO = Path("/Users/jonny/Desktop/해커톤프로젝트/.claude/worktrees/supabase-gamja-20aa04")
sys.path.insert(0, str(REPO))
os.chdir(REPO)

import openpyxl                                        # noqa: E402
import aggregate as ag                                 # noqa: E402

ag.load_env(REPO / ".env")
os.environ.pop("AUTH_DISABLED", None)                  # 인증을 켠 상태로 돌린다

from fastapi.testclient import TestClient              # noqa: E402
import auth, server, store                             # noqa: E402
import formgen as fg                                   # noqa: E402
from test_formgen import Stub                          # noqa: E402
from test_auth import _delete_user, _delete_test_logs, _delete_storage_for_owner  # noqa: E402

SRC = REPO / "sampledata/2025년도 집중안전점검 지열지점 사업장 자체점검 리스트(양식)_판교지사.xlsx"
REPLY_A = REPO / "sampledata/2025년도 집중안전점검 지열지점 사업장 자체점검 리스트(대구).xlsx"
REPLY_B = REPO / "sampledata/2025년도 집중안전점검 지열지점 사업장 자체점검 리스트(용인).xlsx"
SHEET = "2025년 대상 리스트"
FORM_WIDTH = 77.7          # 회신본에 없는 값 — 이 값이 결과에 나오면 규칙 1이 돈 것이다
RESET_EMAIL = "aggregation-test@example.com"

made_users: list[str] = []
_wait_client: TestClient | None = None


def wait(res, timeout: float = 180.0):
    assert res.status_code == 200, res.text
    jid = res.json()["job_id"]
    deadline = time.time() + timeout
    while time.time() < deadline:
        job = _wait_client.get(f"/api/job/{jid}").json()
        if job["status"] == "done":
            return job["result"]
        if job["status"] == "error":
            raise AssertionError(f"잡 실패: {job['error']}")
        time.sleep(0.05)
    raise AssertionError("잡이 시간 안에 끝나지 않았다")


def signup() -> tuple[TestClient, str, str]:
    emp, pw = f"zz-test-{uuid.uuid4().hex[:10]}", "test-" + uuid.uuid4().hex[:12]
    c = TestClient(server.app)
    res = c.post("/api/auth/signup",
                 json={"employee_no": emp, "password": pw, "reset_email": RESET_EMAIL})
    assert res.status_code == 200, res.text
    uid = res.json()["profile"]["id"]
    made_users.append(uid)
    assert c.post("/api/auth/login", json={"employee_no": emp, "password": pw}).status_code == 200
    return c, uid, emp


def form_file(tmp: Path) -> Path:
    """실제 양식을 복사해 열너비만 특이값으로 바꾼다 — 회신본과 구별되게."""
    path = tmp / "zz-test 등록양식.xlsx"
    shutil.copy(SRC, path)
    wb = openpyxl.load_workbook(path)
    wb[SHEET].column_dimensions["A"].width = FORM_WIDTH
    wb.save(path)
    return path


def register_form(c: TestClient, tmp: Path) -> tuple[str, list[str]]:
    """F6 등록을 실제 엔드포인트로 태운다. LLM만 스텁이라 유료 호출은 0건이다."""
    path = form_file(tmp)
    # 입력이 외부 API로 나가므로 동의 전에는 F1·F6 전체가 403이다(명세 §9.1)
    assert c.get("/api/form/consent").json()["consented"] is False
    assert c.post("/api/form/consent").status_code == 200
    up = c.post("/api/form/attachments",
                files=[("files", (path.name, path.read_bytes(), server.XLSX_MIME))])
    assert up.status_code == 200, up.text
    made = up.json()
    sid = made["session_id"]
    assert made["mode"] == "recognize", made
    xlsx = next(a for a in made["attachments"] if a["kind"] == "xlsx")
    table = next(t for t in xlsx["tables"] if t["name"] == SHEET)
    headers = table["headers"]
    print(f"  인지한 표: {SHEET} · 항목 {len(headers)}개 → {headers}")

    # 항목명은 헤더 문자열 그대로여야 한다(인지 루브릭이 1:1을 강제한다).
    # 필수/선택을 갈라 놓는다 — 전부 같으면 '근거 없는 기본값'으로 막힌다
    types = {"개소": "number", "준공연도": "number", "점검일시": "date", "온도차(℃)": "number"}
    fields = [{"name": h, "type": types.get(h, "text"),
               "required": h not in ("보수여부", "온도차(℃)"),
               "format": "YYYY-MM-DD" if types.get(h) == "date" else None,
               "notes": None, "confidence": "high"} for h in headers]
    spec = {"form_title": "zz-test 지열지점 자체점검", "locale": "ko",
            "selected_sheets": [SHEET], "fields": fields}

    real = fg._client
    try:
        fg._client = lambda: Stub([{"intent": "recognize", "reply": "이렇게 읽었습니다",
                                    "spec_json": spec, "recognize_complete": True}])
        turn = c.post(f"/api/form/{sid}/messages", json={"message": "이 양식 그대로 받을 거야"})
        assert turn.status_code == 200, turn.text
        assert turn.json()["spec_complete"] is True, turn.json().get("gaps")
    finally:
        fg._client = real

    reg = c.post(f"/api/form/{sid}/register")
    assert reg.status_code == 200, reg.text
    tid = reg.json()["template_id"]
    # 등록은 원본을 다시 쓰지 않는다 — 바이트 무변경이 구조적으로 보장된다(인지 E1)
    row = store.get_form_template(tid)
    stored = store.get_object(store.RESULT_BUCKET, row["file_url"])
    assert stored == path.read_bytes(), "등록 전후 바이트가 다르다"
    print(f"  등록 완료: template_id={tid[:8]}… · 바이트 무변경 확인 ({len(stored):,})")
    return tid, headers


def aggregate(c: TestClient, template_id: str | None) -> tuple[dict, bytes]:
    global _wait_client
    _wait_client = c
    sid = c.post("/api/session").json()["sid"]
    wait(c.post(f"/api/session/{sid}/files", files=[
        ("files", (REPLY_A.name, REPLY_A.read_bytes(), server.XLSX_MIME)),
        ("files", (REPLY_B.name, REPLY_B.read_bytes(), server.XLSX_MIME)),
    ]))
    review = wait(c.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    picked = [f["fid"] for f in review["files"]]
    body = {"mode": "B", "included": picked, "group_map": {}, "summary_cols": [],
            "sheets": [SHEET]}
    if template_id:
        body["template_id"] = template_id
    out = wait(c.post(f"/api/session/{sid}/aggregate", json=body))
    dl = c.get(f"/api/session/{sid}/download/result")
    assert dl.status_code == 200, dl.status_code
    return out, dl.content


def width_of(data: bytes) -> float | None:
    import io
    ws = openpyxl.load_workbook(io.BytesIO(data))[SHEET]
    dim = ws.column_dimensions.get("B")            # 부서가 앞에 끼므로 양식 A열 → 결과 B열
    return round(dim.width, 1) if dim and dim.width else None


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="rule1-"))
    ok = True
    try:
        print("== 계정 준비 ==")
        ca, uid_a, emp_a = signup()
        cb, uid_b, emp_b = signup()
        print(f"  A={emp_a}  B={emp_b}")

        print("\n== F6 등록 (LLM 스텁, 유료 호출 0건) ==")
        tid, headers = register_form(ca, tmp)

        print("\n== ① 규칙 1: 양식을 골라 취합 ==")
        out1, data1 = aggregate(ca, tid)
        w1 = width_of(data1)
        print(f"  format_source : {out1['format_source']}")
        print(f"  결과 B열 너비 : {w1}   (양식={FORM_WIDTH})")
        for label, cond in [("등록 양식이 기준으로 표기됨", "등록 양식" in out1["format_source"]),
                            (f"양식의 열너비 {FORM_WIDTH}가 결과에 적용됨", w1 == FORM_WIDTH)]:
            print(f"  {'✓' if cond else '✗'} {label}")
            ok &= cond

        print("\n== ② 규칙 2: 양식 없이 취합 (대조군) ==")
        out2, data2 = aggregate(ca, None)
        w2 = width_of(data2)
        print(f"  format_source : {out2['format_source']}")
        print(f"  결과 B열 너비 : {w2}")
        for label, cond in [("공통 서식이 기준으로 표기됨", "공통 서식" in out2["format_source"]),
                            ("양식 너비가 쓰이지 않음", w2 != FORM_WIDTH)]:
            print(f"  {'✓' if cond else '✗'} {label}")
            ok &= cond

        print("\n== ③ 소유권: B가 A의 양식 id로 취합 ==")
        out3, data3 = aggregate(cb, tid)
        w3 = width_of(data3)
        print(f"  format_source : {out3['format_source']}")
        print(f"  결과 B열 너비 : {w3}")
        for label, cond in [("500이 아니라 취합이 정상 완료됨", bool(out3["result_sheets"])),
                            ("남의 양식 서식이 쓰이지 않음", w3 != FORM_WIDTH),
                            ("규칙 2로 내려감", "공통 서식" in out3["format_source"])]:
            print(f"  {'✓' if cond else '✗'} {label}")
            ok &= cond

        print("\n== ④ 없는 양식 id (규칙 1 조회 실패 폴백) ==")
        out4, data4 = aggregate(ca, "not-a-uuid-at-all")
        print(f"  format_source : {out4['format_source']}")
        cond = bool(out4["result_sheets"]) and "공통 서식" in out4["format_source"]
        print(f"  {'✓' if cond else '✗'} 잘못된 id에도 취합이 죽지 않고 규칙 2로 내려감")
        ok &= cond
    finally:
        print("\n== 정리 ==")
        for uid in made_users:
            _delete_storage_for_owner(uid)
            _delete_user(uid)
        _delete_test_logs()
        shutil.rmtree(tmp, ignore_errors=True)
        print(f"  테스트 계정 {len(made_users)}개 · Storage · 감사 로그 정리 완료")
    print("\n전체 통과" if ok else "\n실패 항목 있음")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
