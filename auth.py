#!/usr/bin/env python3
"""F4-1 인증 — 사번 로그인을 Supabase Auth에 위임한다.

Supabase Auth는 이메일/비밀번호 기반이고 PRD는 사번을 요구하므로,
design.md §2대로 사번을 `{사번}@internal.local` 가상 이메일로 치환한다.
프론트는 사번만 알고, 가상 이메일은 이 모듈 밖으로 나가지 않는다.

토큰은 Supabase가 발급한 JWT를 그대로 쓰고 요청마다 검증한다(design.md §2).
서명은 프로젝트 JWKS(ES256 공개키)로 로컬 검증하므로 요청마다 Supabase를
호출하지 않는다.

전달 방식만 명세와 다르다 — Authorization 헤더 대신 httpOnly 쿠키에 담는다.
브라우저 JS가 토큰을 만질 수 없어 XSS로 새지 않고, 프론트가 토큰을 보관할
코드도 없어진다. 쿠키 자동 전송에 따르는 CSRF는 SameSite=Lax로 막는다.
"""

from __future__ import annotations

import os
import re

import httpx
import jwt
from fastapi import Cookie, HTTPException, Response
from jwt import PyJWKClient

VIRTUAL_EMAIL_DOMAIN = "internal.local"     # design.md §2
ACCESS_COOKIE = "chwihap_at"
REFRESH_COOKIE = "chwihap_rt"
EMPLOYEE_NO_RE = re.compile(r"^[A-Za-z0-9._-]{3,32}$")
MIN_PASSWORD_LEN = 8

_jwks: PyJWKClient | None = None


def _cfg(name: str) -> str:
    value = os.environ.get(name, "")
    if not value:
        raise HTTPException(503, f"{name}이(가) 설정되지 않았습니다. .env를 확인해주세요.")
    return value


def _secret_key() -> str:
    """서버 전용 키. 신형 `sb_secret_...`을 우선하고 legacy service_role로 물러선다.

    `supabase projects api-keys`는 신형 secret 키를 마스킹해서 내려주므로
    CLI만으로는 채울 수 없다. 대시보드에서 복사해 .env의 SUPABASE_SECRET_KEY에
    넣으면 그쪽이 쓰이고, 없으면 legacy service_role을 쓴다.
    """
    for name in ("SUPABASE_SECRET_KEY", "SUPABASE_SERVICE_ROLE_KEY"):
        value = os.environ.get(name, "")
        if value:
            if not value.isascii():
                raise HTTPException(503, f"{name} 값이 손상되었습니다(마스킹된 값일 수 있습니다).")
            return value
    raise HTTPException(503, "SUPABASE_SECRET_KEY 또는 SUPABASE_SERVICE_ROLE_KEY가 필요합니다.")


def configured() -> bool:
    """인증을 켤 수 있는 상태인지."""
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_PUBLISHABLE_KEY")
                and (os.environ.get("SUPABASE_SECRET_KEY")
                     or os.environ.get("SUPABASE_SERVICE_ROLE_KEY")))


def _normalize(employee_no: str) -> str:
    """사번은 대소문자를 구분하지 않는다.

    가상 이메일을 소문자로 만들면서 저장값은 입력 그대로 두면, Postgres의
    대소문자 구분 UNIQUE 때문에 `Admin01`과 `admin01`이 중복 검사를 통과한
    뒤 Auth 쪽에서 같은 이메일로 충돌한다. 한 곳에서 소문자로 맞춰 둔다.
    """
    return (employee_no or "").strip().lower()


def _virtual_email(employee_no: str) -> str:
    return f"{employee_no}@{VIRTUAL_EMAIL_DOMAIN}"


def _auth_url(path: str) -> str:
    return f"{_cfg('SUPABASE_URL')}/auth/v1{path}"


def _rest_url(path: str) -> str:
    return f"{_cfg('SUPABASE_URL')}/rest/v1{path}"


def _secret_headers() -> dict:
    """RLS를 우회하는 서버 전용 키. 서버 안에서만 쓰고 응답에 싣지 않는다."""
    key = _secret_key()
    return {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}


def _detail(res: httpx.Response, fallback: str) -> str:
    try:
        body = res.json()
    except Exception:
        return fallback
    for field in ("msg", "message", "error_description", "error", "hint"):
        if isinstance(body, dict) and body.get(field):
            return str(body[field])
    return fallback


# ── 가입 ──────────────────────────────────────────────────────────────────────
def signup(employee_no: str, password: str, reset_email: str) -> dict:
    """사번으로 가입. 가입 제한은 없고 Admin이 status로 사후 통제한다 (F4-1)."""
    employee_no = _normalize(employee_no)
    reset_email = (reset_email or "").strip()
    if not EMPLOYEE_NO_RE.match(employee_no):
        raise HTTPException(400, "사번은 영문·숫자 3~32자로 입력해주세요.")
    if len(password or "") < MIN_PASSWORD_LEN:
        raise HTTPException(400, f"비밀번호는 {MIN_PASSWORD_LEN}자 이상이어야 합니다.")
    if "@" not in reset_email or "." not in reset_email.split("@")[-1]:
        raise HTTPException(400, "비밀번호 재설정용 이메일 형식이 올바르지 않습니다.")

    with httpx.Client(timeout=20) as client:
        # 사번 선점 여부를 먼저 본다. auth 사용자만 만들어지고 profiles가 실패하면
        # 로그인은 되는데 프로필이 없는 상태가 되므로 그 전에 걸러낸다.
        dup = client.get(_rest_url("/profiles"), headers=_secret_headers(),
                         params={"employee_no": f"eq.{employee_no}", "select": "id"})
        if dup.status_code == 200 and dup.json():
            raise HTTPException(409, f"이미 사용 중인 사번입니다: {employee_no}")

        # 가상 이메일은 실제로 메일을 받을 수 없으므로 확인 절차 없이 생성한다.
        created = client.post(
            _auth_url("/admin/users"), headers=_secret_headers(),
            json={"email": _virtual_email(employee_no), "password": password,
                  "email_confirm": True})
        if created.status_code >= 400:
            raise HTTPException(400, _detail(created, "가입에 실패했습니다."))
        user_id = created.json()["id"]

        # reset_email·role·status는 profiles가 단일 출처다. user_metadata는
        # 사용자가 고칠 수 있어 권한 판단에 쓰면 안 된다.
        profile = client.post(
            _rest_url("/profiles"), headers={**_secret_headers(), "Prefer": "return=representation"},
            json={"id": user_id, "employee_no": employee_no, "reset_email": reset_email,
                  "role": "user", "status": "active"})
        if profile.status_code >= 400:
            # 프로필 없는 유령 계정이 남지 않도록 되돌린다
            client.delete(_auth_url(f"/admin/users/{user_id}"), headers=_secret_headers())
            raise HTTPException(400, _detail(profile, "프로필 생성에 실패했습니다."))
        return profile.json()[0]


# ── 로그인 ────────────────────────────────────────────────────────────────────
def login(employee_no: str, password: str) -> dict:
    employee_no = _normalize(employee_no)
    if not employee_no or not password:
        raise HTTPException(400, "사번과 비밀번호를 입력해주세요.")
    with httpx.Client(timeout=20) as client:
        res = client.post(
            _auth_url("/token"), params={"grant_type": "password"},
            headers={"apikey": _cfg("SUPABASE_PUBLISHABLE_KEY"), "Content-Type": "application/json"},
            json={"email": _virtual_email(employee_no), "password": password})
        if res.status_code >= 400:
            # 사번 존재 여부를 알려주지 않는다(계정 열거 방지)
            raise HTTPException(401, "사번 또는 비밀번호가 올바르지 않습니다.")
        token = res.json()
        profile = _fetch_profile(client, token["user"]["id"])
        if profile is None:
            raise HTTPException(403, "프로필이 없는 계정입니다. 관리자에게 문의해주세요.")
        if profile.get("status") != "active":
            raise HTTPException(403, "정지된 계정입니다. 관리자에게 문의해주세요.")
        return {"token": token, "profile": profile}


def _fetch_profile(client: httpx.Client, user_id: str) -> dict | None:
    res = client.get(_rest_url("/profiles"), headers=_secret_headers(),
                     params={"id": f"eq.{user_id}",
                             "select": "id,employee_no,reset_email,role,status,created_at"})
    if res.status_code >= 400:
        raise HTTPException(502, _detail(res, "프로필 조회에 실패했습니다."))
    rows = res.json()
    return rows[0] if rows else None


def request_password_reset(employee_no: str) -> None:
    """가상 이메일이 아니라 profiles.reset_email로 보낸다 (design.md §2).

    사번이 없어도 같은 응답을 주어 계정 존재 여부를 노출하지 않는다.
    """
    with httpx.Client(timeout=20) as client:
        res = client.get(_rest_url("/profiles"), headers=_secret_headers(),
                         params={"employee_no": f"eq.{_normalize(employee_no)}",
                                 "select": "reset_email"})
        rows = res.json() if res.status_code == 200 else []
        if not rows:
            return
        client.post(_auth_url("/recover"),
                    headers={"apikey": _cfg("SUPABASE_PUBLISHABLE_KEY"),
                             "Content-Type": "application/json"},
                    json={"email": rows[0]["reset_email"]})


# ── 세션 쿠키 ─────────────────────────────────────────────────────────────────
def set_session_cookies(response: Response, token: dict) -> None:
    secure = os.environ.get("COOKIE_SECURE", "").lower() in ("1", "true", "yes")
    common = {"httponly": True, "samesite": "lax", "secure": secure, "path": "/"}
    response.set_cookie(ACCESS_COOKIE, token["access_token"],
                        max_age=token.get("expires_in", 3600), **common)
    if token.get("refresh_token"):
        response.set_cookie(REFRESH_COOKIE, token["refresh_token"], max_age=30 * 24 * 3600, **common)


def clear_session_cookies(response: Response) -> None:
    for name in (ACCESS_COOKIE, REFRESH_COOKIE):
        response.delete_cookie(name, path="/")


def refresh(refresh_token: str) -> dict:
    with httpx.Client(timeout=20) as client:
        res = client.post(
            _auth_url("/token"), params={"grant_type": "refresh_token"},
            headers={"apikey": _cfg("SUPABASE_PUBLISHABLE_KEY"), "Content-Type": "application/json"},
            json={"refresh_token": refresh_token})
        if res.status_code >= 400:
            raise HTTPException(401, "세션이 만료되었습니다. 다시 로그인해주세요.")
        return res.json()


# ── 요청 검증 ─────────────────────────────────────────────────────────────────
def _jwk_client() -> PyJWKClient:
    global _jwks
    if _jwks is None:
        # PyJWKClient가 키를 캐시하므로 요청마다 JWKS를 받아오지 않는다
        _jwks = PyJWKClient(f"{_cfg('SUPABASE_URL')}/auth/v1/.well-known/jwks.json")
    return _jwks


# 토큰을 발급하는 쪽(Supabase)과 검증하는 쪽(이 서버)의 시계는 언제나 조금 어긋난다.
# 여유가 없으면 **방금 발급된 토큰이 `iat`가 미래라는 이유로 거부된다** — 실측에서 이 머신이
# 서버보다 3~4초 느려 로그인 직후 모든 요청이 401이 됐다(`ImmatureSignatureError`).
# 사용자에게는 "로그인은 됐는데 아무것도 안 된다"로 보이고, 시계가 다시 맞으면 사라져
# 원인을 찾기 어렵다. exp에도 같은 여유가 적용되지만 1분은 세션 수명(1시간)에 비해 무해하다.
CLOCK_SKEW_SEC = 60


def verify_token(access_token: str) -> dict:
    """JWT 서명·만료를 로컬에서 검증하고 클레임을 돌려준다."""
    try:
        key = _jwk_client().get_signing_key_from_jwt(access_token).key
        return jwt.decode(access_token, key, algorithms=["ES256", "RS256"],
                          audience="authenticated", leeway=CLOCK_SKEW_SEC,
                          options={"require": ["exp", "sub"]})
    except jwt.ExpiredSignatureError:
        raise HTTPException(401, "세션이 만료되었습니다. 다시 로그인해주세요.")
    except Exception:
        raise HTTPException(401, "로그인이 필요합니다.")


def current_user(access_token: str | None = Cookie(default=None, alias=ACCESS_COOKIE)) -> dict:
    """FastAPI 의존성. 로그인하지 않았으면 401."""
    if not configured():
        raise HTTPException(503, "인증이 설정되지 않았습니다. .env의 SUPABASE_* 값을 확인해주세요.")
    if not access_token:
        raise HTTPException(401, "로그인이 필요합니다.")
    claims = verify_token(access_token)
    with httpx.Client(timeout=20) as client:
        profile = _fetch_profile(client, claims["sub"])
    if profile is None:
        raise HTTPException(401, "로그인이 필요합니다.")
    if profile.get("status") != "active":
        # 정지 처리가 이미 발급된 토큰에도 즉시 반영되도록 매 요청 확인한다
        raise HTTPException(403, "정지된 계정입니다. 관리자에게 문의해주세요.")
    return profile


def require_admin(user: dict) -> dict:
    if user.get("role") != "admin":
        raise HTTPException(403, "관리자 권한이 필요합니다.")
    return user
