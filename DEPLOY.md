# 배포 (Fly.io)

컨테이너 이미지와 설정은 저장소에 있습니다(`Dockerfile`·`fly.toml`·`.dockerignore`).
**실제 배포는 계정이 필요해 사람이 실행합니다.** 아래 순서대로 하면 됩니다.

## 왜 Fly.io인가

- **Supabase가 서울(ap-northeast-2)**입니다. 요청 한 번에 Supabase REST를 여러 번 오가므로
  (인증 확인 → 기록 → Storage) 가까운 리전이 그대로 체감됩니다. Fly의 도쿄(`nrt`)가 가장 가깝습니다.
- **세션·잡이 프로세스 메모리**라 인스턴스가 하나여야 하고 잠들면 안 됩니다. Fly는 이걸 설정으로
  못 박을 수 있습니다(`min_machines_running = 1`, `max_machines_running = 1`, `auto_stop_machines = false`).
- Vercel 같은 서버리스는 맞지 않습니다 — 요청마다 인스턴스가 갈리면 업로드 세션을 잃고,
  2단계 AI 재검증(파일당 20초대)이 함수 실행 시간에 걸립니다.

Railway·Render도 같은 Dockerfile로 됩니다. 다만 리전 선택이 좁고(미국 위주) 무료 등급이
잠들 수 있어, 그 두 성질이 이 앱과 맞지 않습니다.

## 배포 절차

```bash
brew install flyctl        # 이 머신에는 아직 없습니다
fly auth login             # 브라우저로 로그인
fly launch --no-deploy --copy-config    # fly.toml을 그대로 씁니다(앱 이름만 정하면 됩니다)
```

비밀값은 `fly.toml`에 넣지 않고 secret으로 넣습니다(`.env`는 커밋되지 않습니다).

```bash
fly secrets set \
  OPENAI_API_KEY="$(grep '^OPENAI_API_KEY=' .env | cut -d= -f2-)" \
  OPENAI_MODEL="$(grep '^OPENAI_MODEL=' .env | cut -d= -f2-)" \
  SUPABASE_URL="$(grep '^SUPABASE_URL=' .env | cut -d= -f2-)" \
  SUPABASE_PUBLISHABLE_KEY="$(grep '^SUPABASE_PUBLISHABLE_KEY=' .env | cut -d= -f2-)" \
  SUPABASE_SERVICE_ROLE_KEY="$(grep '^SUPABASE_SERVICE_ROLE_KEY=' .env | cut -d= -f2-)"

fly deploy
fly open /login.html
```

## 배포 전에 반드시 확인할 것

| 항목 | 왜 |
|---|---|
| `AUTH_DISABLED`를 **넣지 않는다** | 이 값이 켜지면 인증이 통째로 없어집니다. 키가 없으면 열리는 게 아니라 503으로 닫히도록 되어 있으니(fail-closed) 그냥 두면 됩니다 |
| `COOKIE_SECURE=1` | `fly.toml`에 이미 있습니다. 없으면 HTTPS에서도 쿠키에 Secure가 붙지 않습니다 |
| 머신 1대 유지 | 늘리면 세션을 잃습니다. `fly scale count 1`을 넘기지 마세요 |
| 첫 관리자 지정 | 배포 후 `/login.html`에서 가입하고, 로컬에서 `python3 manage_users.py promote <사번>`으로 올립니다(Admin 화면은 관리자만 들어갑니다) |

## 확인한 것 (로컬 도커)

배포 전에 같은 이미지를 로컬에서 띄워 확인했습니다.

- 빌드 성공, 이미지 741MB(가장 큰 것은 cryptography 12MB·pillow 18MB·openai 5.7MB)
- 화면 7개 전부 200, 로그인 없이 `POST /api/session`은 401(인증 게이트 동작)
- 컨테이너 안에서 **실제 Supabase를 상대로 전 구간**: 가입 → 로그인 → 업로드 2개 →
  검토(정상 2건) → 모드 B 취합(2파일·1시트) → 결과 다운로드(4행) → 이력 1건

## 알아둘 한계

- **재시작하면 진행 중이던 작업이 사라집니다.** 기록과 산출물은 Supabase에 남지만, 업로드
  세션과 잡 진행률은 프로세스 메모리입니다. 배포(=재시작) 중에 취합을 돌리던 사용자는 다시
  올려야 합니다.
- **인스턴스를 늘리려면 세션·잡을 밖으로 빼야 합니다**(DB나 Redis). 그 전까지는 머신 1대가
  이 앱의 구조적 상한입니다.
- 업로드 상한은 파일당 50MB·합계 500MB이고 Admin 화면에서 바꿉니다. 큰 업로드를 자주 쓰면
  머신 메모리(현재 1GB)를 함께 올리세요 — 30개 파일 취합 실측에서 최대 RSS 378MB였습니다.
