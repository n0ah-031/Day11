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

보관 기간 정리도 여기서 돌린다 — 주기 실행은 cron에 맡긴다(서버에 타이머를 두면
재시작이 잦아 언제 무엇이 지워졌는지 알 수 없다).

    python3 manage_users.py purge-expired --dry-run   # 지울 대상만 센다
    python3 manage_users.py purge-expired             # 실제로 지운다
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


def cmd_purge_expired(dry_run: bool) -> int:
    """보관 기간이 지난 작업과 그 파일을 지운다 (F4-3).

    양식을 담은 프로젝트는 대상에서 빠진다(두고두고 다시 쓰는 자산이라 보관 기간의
    대상이 아니다). 몇 건이 그렇게 빠졌는지 함께 찍는다.
    """
    import store
    days = store.get_policy().get("retention_days")
    days = int(days) if str(days).isdigit() else 90
    plan = store.expired_projects(days)
    print(f"기준 {days}일 (이전: {plan['cutoff'][:10]}) · 대상 작업 {len(plan['projects'])}건 · "
          f"파일 {plan['files']}건 · {plan['bytes'] / 1024 / 1024:.1f}MB")
    if plan["kept_with_templates"]:
        print(f"  · 양식이 있어 남겨둔 프로젝트 {plan['kept_with_templates']}건")
    for row in plan["projects"]:
        print(f"  - {row['created_at'][:10]} {row['name']} (파일 {row['files']}건)")
    if dry_run:
        print("--dry-run: 지우지 않았습니다.")
        return 0
    if not plan["projects"]:
        return 0
    store.purge_projects([row["id"] for row in plan["projects"]])
    store.log_action(None, "보관 기간 정리",
                     f"기준 {days}일 · 작업 {len(plan['projects'])}건 · 파일 {plan['files']}건 (CLI)")
    print(f"{len(plan['projects'])}건을 지웠습니다.")
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not auth.configured():
        print(".env에 SUPABASE_URL·SUPABASE_PUBLISHABLE_KEY와 "
              "SUPABASE_SECRET_KEY(또는 SUPABASE_SERVICE_ROLE_KEY)가 필요합니다.", file=sys.stderr)
        return 2
    if not argv or argv[0] == "list":
        return cmd_list()
    if argv[0] == "purge-expired":
        return cmd_purge_expired("--dry-run" in argv)
    action = argv[0]
    if action not in COMMANDS or len(argv) < 2:
        print(__doc__)
        return 2
    return COMMANDS[action](argv[1])


if __name__ == "__main__":
    sys.exit(main())
