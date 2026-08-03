#!/usr/bin/env python3
"""계정 관리 CLI (F4-3 Admin 콘솔이 나오기 전까지 쓰는 임시 도구).

Admin 화면이 아직 없어 role·status를 바꿀 방법이 없으므로 서버 키로 직접 고친다.
비밀번호는 다루지 않는다 — 가입은 화면에서 본인이 하고, 이 도구는 이미 있는
계정의 권한·상태만 바꾼다.

    python3 manage_users.py list
    python3 manage_users.py promote admin01      # role → admin
    python3 manage_users.py demote  admin01      # role → user
    python3 manage_users.py suspend 12345        # status → suspended
    python3 manage_users.py activate 12345       # status → active
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import httpx

import aggregate as ag
import auth

ag.load_env(Path(__file__).parent / ".env")


def _headers() -> dict:
    key = auth._secret_key()
    return {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}


def _base() -> str:
    return os.environ["SUPABASE_URL"].rstrip("/")


def cmd_list() -> int:
    res = httpx.get(f"{_base()}/rest/v1/profiles", headers=_headers(), timeout=20,
                    params={"select": "employee_no,role,status,reset_email,created_at",
                            "order": "created_at"})
    res.raise_for_status()
    rows = res.json()
    if not rows:
        print("등록된 계정이 없습니다.")
        return 0
    print(f"{'사번':<24} {'권한':<7} {'상태':<10} {'재설정 이메일':<32} 가입일")
    for r in rows:
        print(f"{r['employee_no']:<24} {r['role']:<7} {r['status']:<10} "
              f"{r['reset_email']:<32} {r['created_at'][:10]}")
    return 0


def _patch(employee_no: str, payload: dict, label: str) -> int:
    employee_no = auth._normalize(employee_no)
    res = httpx.patch(f"{_base()}/rest/v1/profiles",
                      headers={**_headers(), "Prefer": "return=representation"},
                      params={"employee_no": f"eq.{employee_no}"}, json=payload, timeout=20)
    if res.status_code >= 400:
        print(f"실패: {res.status_code} {res.text}", file=sys.stderr)
        return 1
    rows = res.json()
    if not rows:
        print(f"'{employee_no}' 사번을 찾을 수 없습니다. `list`로 확인해주세요.", file=sys.stderr)
        return 1
    r = rows[0]
    print(f"{label}: {r['employee_no']} → role={r['role']}, status={r['status']}")
    return 0


COMMANDS = {
    "promote": lambda emp: _patch(emp, {"role": "admin"}, "권한 상향"),
    "demote": lambda emp: _patch(emp, {"role": "user"}, "권한 하향"),
    "suspend": lambda emp: _patch(emp, {"status": "suspended"}, "계정 정지"),
    "activate": lambda emp: _patch(emp, {"status": "active"}, "계정 활성"),
}


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not auth.configured():
        print(".env에 SUPABASE_URL·SUPABASE_PUBLISHABLE_KEY와 "
              "SUPABASE_SECRET_KEY(또는 SUPABASE_SERVICE_ROLE_KEY)가 필요합니다.", file=sys.stderr)
        return 2
    if not argv or argv[0] == "list":
        return cmd_list()
    action = argv[0]
    if action not in COMMANDS or len(argv) < 2:
        print(__doc__)
        return 2
    return COMMANDS[action](argv[1])


if __name__ == "__main__":
    sys.exit(main())
