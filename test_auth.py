#!/usr/bin/env python3
"""인증(F4-1) 점검. 실행: python3 test_auth.py

test_server.py와 달리 **실제 Supabase 프로젝트를 상대로** 돈다.
.env의 SUPABASE_* 값이 필요하며 네트워크를 쓴다.

테스트 계정은 매 실행마다 새로 만들고 끝나면 지운다. 사번에 `zz-test-`
접두사를 붙여 실계정과 구분한다.
"""

from __future__ import annotations

import io
import os
import sys
import time
import uuid

import httpx
import openpyxl

import aggregate as ag
from pathlib import Path

ag.load_env(Path(__file__).parent / ".env")
os.environ.pop("AUTH_DISABLED", None)          # 인증을 켠 상태로 돌려야 한다

from fastapi.testclient import TestClient       # noqa: E402

import auth                                     # noqa: E402
import server                                   # noqa: E402

EMP = f"zz-test-{uuid.uuid4().hex[:10]}"
PW = "test-" + uuid.uuid4().hex[:12]
RESET_EMAIL = "aggregation-test@example.com"
created_user_id: str | None = None
client_for_wait: TestClient | None = None
storage_paths: list[str] = []


def book_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "예산"
    ws["A1"] = "실적 보고"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.append([None, None, None])
    for row in [["사번", "부서", "예산액"], ["A1", "기획부", 100]]:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def wait(res, timeout: float = 120.0):
    """잡 기반 엔드포인트: job_id를 받아 완료까지 폴링하고 결과만 돌려준다."""
    assert res.status_code == 200, res.text
    jid = res.json()["job_id"]
    deadline = time.time() + timeout
    while time.time() < deadline:
        job = client_for_wait.get(f"/api/job/{jid}").json()
        if job["status"] == "done":
            return job["result"]
        if job["status"] == "error":
            raise AssertionError(f"잡 실패: {job['error']}")
        time.sleep(0.02)
    raise AssertionError("잡이 시간 안에 끝나지 않았다")


def test_protected_without_login():
    """로그인 없이는 취합 API에 접근할 수 없어야 한다."""
    anon = TestClient(server.app)
    for method, path in [("post", "/api/session"), ("get", "/api/auth/me"),
                         ("get", "/api/job/아무거나")]:
        res = getattr(anon, method)(path)
        assert res.status_code == 401, f"{path} → {res.status_code} (401이어야 함)"
    print("  ✓ 미로그인 차단(401)")


def test_signup_and_login():
    global created_user_id
    client = TestClient(server.app)

    # 입력 검증
    bad = client.post("/api/auth/signup",
                      json={"employee_no": "a", "password": PW, "reset_email": RESET_EMAIL})
    assert bad.status_code == 400 and "사번" in bad.json()["detail"], bad.text
    bad = client.post("/api/auth/signup",
                      json={"employee_no": EMP, "password": "short", "reset_email": RESET_EMAIL})
    assert bad.status_code == 400 and "비밀번호" in bad.json()["detail"], bad.text

    res = client.post("/api/auth/signup",
                      json={"employee_no": EMP, "password": PW, "reset_email": RESET_EMAIL})
    assert res.status_code == 200, res.text
    profile = res.json()["profile"]
    created_user_id = profile["id"]
    assert profile["employee_no"] == EMP and profile["role"] == "user", profile
    assert profile["status"] == "active", profile

    # 같은 사번 재가입은 막힌다
    dup = client.post("/api/auth/signup",
                      json={"employee_no": EMP, "password": PW, "reset_email": RESET_EMAIL})
    assert dup.status_code == 409, dup.text

    # 대소문자만 다른 사번도 같은 사번이다. 가상 이메일이 소문자라 여기서
    # 걸러내지 않으면 Auth 쪽에서 이메일 중복으로 늦게 터진다
    dup2 = client.post("/api/auth/signup",
                       json={"employee_no": EMP.upper(), "password": PW, "reset_email": RESET_EMAIL})
    assert dup2.status_code == 409, f"대소문자 변형이 통과했다: {dup2.status_code} {dup2.text}"
    print("  ✓ 가입(입력 검증 + 사번 중복 409, 대소문자 무관)")

    # 틀린 비밀번호 — 사번 존재 여부를 노출하지 않는 같은 메시지
    wrong = client.post("/api/auth/login", json={"employee_no": EMP, "password": "wrong-pw-123"})
    missing = client.post("/api/auth/login", json={"employee_no": "zz-none", "password": PW})
    assert wrong.status_code == missing.status_code == 401, (wrong.text, missing.text)
    assert wrong.json()["detail"] == missing.json()["detail"], "계정 열거가 가능하면 안 된다"

    # 대문자로 입력해도 같은 계정으로 로그인된다
    upper = client.post("/api/auth/login", json={"employee_no": EMP.upper(), "password": PW})
    assert upper.status_code == 200, upper.text
    assert upper.json()["profile"]["employee_no"] == EMP, upper.json()

    ok = client.post("/api/auth/login", json={"employee_no": EMP, "password": PW})
    assert ok.status_code == 200, ok.text
    assert ok.json()["profile"]["employee_no"] == EMP
    # 토큰은 본문에 실리지 않고 httpOnly 쿠키로만 나간다
    assert "access_token" not in ok.text, "토큰이 응답 본문에 노출됐다"
    cookie = ok.cookies.get(auth.ACCESS_COOKIE)
    assert cookie, ok.cookies
    set_cookie = " ".join(ok.headers.get_list("set-cookie")).lower()
    assert "httponly" in set_cookie and "samesite=lax" in set_cookie, set_cookie
    print("  ✓ 로그인(계정 열거 방지 + httpOnly·SameSite 쿠키)")
    global client_for_wait
    client_for_wait = client
    return client


def test_authenticated_flow(client: TestClient):
    """로그인한 클라이언트는 취합 전 구간을 쓸 수 있다."""
    me = client.get("/api/auth/me")
    assert me.status_code == 200 and me.json()["profile"]["employee_no"] == EMP, me.text
    assert me.json()["auth_enabled"] is True

    sid = client.post("/api/session").json()["sid"]
    res = client.post(f"/api/session/{sid}/files",
                      files=[("files", ("기획부.xlsx", book_bytes(), "application/octet-stream"))])
    assert res.status_code == 200, res.text
    jid = res.json()["job_id"]
    for _ in range(2000):
        job = client.get(f"/api/job/{jid}").json()
        if job["status"] != "running":
            break
    assert job["status"] == "done", job
    assert "owner" not in job, "소유자 정보가 응답에 새어나가면 안 된다"
    print("  ✓ 로그인 후 취합 API 사용 가능")
    return sid


def test_isolation(sid: str):
    """다른 사용자의 세션·잡은 보이지 않아야 한다 (§6.2)."""
    other_emp = f"zz-test-{uuid.uuid4().hex[:10]}"
    other = TestClient(server.app)
    made = other.post("/api/auth/signup",
                      json={"employee_no": other_emp, "password": PW, "reset_email": RESET_EMAIL})
    assert made.status_code == 200, made.text
    other_id = made.json()["profile"]["id"]
    try:
        assert other.post("/api/auth/login",
                          json={"employee_no": other_emp, "password": PW}).status_code == 200
        # 남의 sid는 존재 자체를 알리지 않는다
        assert other.post(f"/api/session/{sid}/review", json={"no_ai": True}).status_code == 404
        assert other.get(f"/api/session/{sid}/download/result").status_code == 404
        print("  ✓ 사용자 간 세션 격리(404)")
    finally:
        _delete_storage_for_owner(other_id)
        _delete_user(other_id)


def test_role_and_suspend(client: TestClient):
    """manage_users.py로 바꾼 권한·상태가 즉시 반영돼야 한다."""
    import manage_users

    assert manage_users.main(["promote", EMP.upper()]) == 0, "대소문자 무관하게 찾아야 한다"
    assert client.get("/api/auth/me").json()["profile"]["role"] == "admin"

    assert manage_users.main(["demote", EMP]) == 0
    assert client.get("/api/auth/me").json()["profile"]["role"] == "user"

    # 정지는 이미 발급된 토큰에도 즉시 걸려야 한다 (매 요청 status 확인)
    assert manage_users.main(["suspend", EMP]) == 0
    blocked = client.get("/api/auth/me")
    assert blocked.status_code == 403, f"{blocked.status_code} {blocked.text}"
    assert "정지" in blocked.json()["detail"]
    # 정지 상태에서는 새로 로그인도 막힌다
    assert client.post("/api/auth/login",
                       json={"employee_no": EMP, "password": PW}).status_code == 403

    assert manage_users.main(["activate", EMP]) == 0
    assert client.get("/api/auth/me").status_code == 200
    print("  ✓ 권한 변경 + 계정 정지(발급된 토큰에도 즉시 적용)")


def test_persistence(client: TestClient):
    """업로드 파일은 Storage, 작업 기록은 DB에 남아야 한다 (세션 휘발성 해소)."""
    import store

    made = client.post("/api/session").json()
    sid, project_id = made["sid"], made["project_id"]
    assert project_id, "취합 시작 시 프로젝트가 자동 생성돼야 한다"

    res = client.post(f"/api/session/{sid}/files",
                      files=[("files", ("기획부.xlsx", book_bytes(), "application/octet-stream"))])
    jid = res.json()["job_id"]
    for _ in range(2000):
        if client.get(f"/api/job/{jid}").json()["status"] != "running":
            break
    wait(client.post(f"/api/session/{sid}/review", json={"no_ai": True}))
    agg = wait(client.post(f"/api/session/{sid}/aggregate", json={"mode": "B", "included": [0]}))
    assert agg["metrics"]["aggregated"] == 1, agg

    # 이력에 남았는지
    detail = client.get(f"/api/history/{project_id}")
    assert detail.status_code == 200, detail.text
    d = detail.json()
    assert len(d["uploaded_files"]) == 1, d["uploaded_files"]
    upl = d["uploaded_files"][0]
    assert upl["original_name"] == "기획부.xlsx" and upl["status"] == "reviewed", upl
    assert upl["review_results"] and upl["review_results"][0]["badge"] == "정상", upl
    job = d["aggregation_jobs"][0]
    assert job["status"] == "done" and job["progress"] == 100, job
    assert job["mode"] == "B" and job["result_url"], job

    # 원본 파일이 Storage에 실제로 있는지 (재시작해도 남는 근거)
    files = store._rest("GET", "/uploaded_files",
                        params={"id": f"eq.{upl['id']}", "select": "storage_path"}).json()
    blob = store.get_object(store.UPLOAD_BUCKET, files[0]["storage_path"])
    assert blob[:2] == b"PK" and len(blob) > 1000, len(blob)

    # 세션이 사라진 뒤에도 이력에서 결과를 내려받을 수 있는지
    import server
    server.SESSIONS.clear()
    dl = client.get(f"/api/history/{project_id}/download/{job['id']}")
    assert dl.status_code == 200 and dl.content[:2] == b"PK", dl.status_code

    # 목록에도 보이고, 남의 이력은 404
    projects = client.get("/api/history").json()["projects"]
    assert any(p["id"] == project_id for p in projects), projects
    print("  ✓ 영속화(Storage 원본 + DB 기록 + 세션 없이 이력 다운로드)")
    return project_id


def test_logout(client: TestClient):
    assert client.post("/api/auth/logout").status_code == 200
    assert client.get("/api/auth/me").status_code == 401, "로그아웃 후에도 접근되면 안 된다"
    print("  ✓ 로그아웃")


def _delete_user(user_id: str | None) -> None:
    if not user_id:
        return
    key = auth._secret_key()
    headers = {"apikey": key, "Authorization": f"Bearer {key}"}
    # profiles는 auth.users FK가 ON DELETE CASCADE라 함께 지워진다
    httpx.delete(f"{os.environ['SUPABASE_URL']}/auth/v1/admin/users/{user_id}",
                 headers=headers, timeout=20)


def _delete_storage_for_owner(owner_id: str | None) -> None:
    """이 사용자의 모든 프로젝트 폴더를 정리한다.

    DB는 profiles → projects → uploaded_files가 CASCADE로 지워지지만
    Storage 객체는 대상이 아니라 남는다. 사용자를 지우기 전에 훑어야
    프로젝트 목록을 알 수 있다.
    """
    if not owner_id:
        return
    key = auth._secret_key()
    headers = {"apikey": key, "Authorization": f"Bearer {key}",
               "Content-Type": "application/json"}
    base = os.environ["SUPABASE_URL"].rstrip("/")
    rows = httpx.get(f"{base}/rest/v1/projects", headers=headers, timeout=20,
                     params={"owner_id": f"eq.{owner_id}", "select": "id"})
    for row in (rows.json() if rows.status_code == 200 else []):
        _delete_storage(row["id"])


def _delete_storage(project_id: str | None) -> None:
    if not project_id:
        return
    key = auth._secret_key()
    headers = {"apikey": key, "Authorization": f"Bearer {key}",
               "Content-Type": "application/json"}
    base = os.environ["SUPABASE_URL"].rstrip("/")
    for bucket in ("uploads", "results"):
        listed = httpx.post(f"{base}/storage/v1/object/list/{bucket}", headers=headers, timeout=20,
                            json={"prefix": f"{project_id}/", "limit": 200})
        if listed.status_code >= 400:
            continue
        names = [f"{project_id}/{o['name']}" for o in listed.json()]
        if names:
            httpx.request("DELETE", f"{base}/storage/v1/object/{bucket}", headers=headers,
                          json={"prefixes": names}, timeout=30)


def main() -> int:
    if not auth.configured():
        print("건너뜀: .env에 SUPABASE_URL·SUPABASE_PUBLISHABLE_KEY·SUPABASE_SECRET_KEY가 필요합니다.")
        return 0
    try:
        test_protected_without_login()
        client = test_signup_and_login()
        sid = test_authenticated_flow(client)
        test_isolation(sid)
        test_role_and_suspend(client)
        test_persistence(client)
        test_logout(client)
    finally:
        # 사용자를 지우면 프로젝트도 CASCADE로 사라지므로 Storage를 먼저 훑는다
        _delete_storage_for_owner(created_user_id)
        _delete_user(created_user_id)
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
