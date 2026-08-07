# 배포 — Oracle Cloud Always Free

VM 한 대에 도커로 올립니다. 구성은 `docker-compose.yml`(앱 + Caddy) 하나입니다.

## 왜 이 방식인가

이 앱이 호스팅에 거는 조건은 셋이고, 전부 구조에서 나옵니다 — 취향이 아닙니다.

| 조건 | 왜 | 안 지키면 |
|---|---|---|
| **인스턴스 1개** | 업로드 세션·잡 진행률이 프로세스 메모리(dict + 임시 폴더)에 있다 | 요청이 다른 인스턴스로 가면서 "세션을 찾을 수 없습니다" |
| **응답 후에도 CPU가 돌 것** | 긴 작업을 백그라운드 스레드로 돌리고 요청에는 즉시 `job_id`만 준다(`server._start_job`) | 취합이 진행되지 않는다 |
| **요청이 오래 걸려도 안 끊길 것** | 2단계 AI 재검증이 파일당 20초대(30개 파일 실측 112초, HANDOFF §4) | 큰 취합이 잘린다 |

VM은 이 셋을 그냥 만족합니다. 서버리스에서 설정으로 우회해야 했던 것들이 문제가 되지 않고,
잠들지 않고, 무료입니다.

## 1. VM 만들기

> **Always Free는 홈 리전에서만 무료입니다.** 홈 리전은 가입할 때 정해집니다. 인스턴스를
> 만들 때 콘솔 좌측 상단 리전이 가입 때 고른 그 리전인지 반드시 확인하세요 — 다른 리전에
> 만들면 그냥 과금됩니다.
>
> 가입 목록에 한국(서울)이 없는 경우가 있습니다(오라클도 "원하는 리전이 없으면 인접 리전을
> 고르라"고 안내합니다). 그때는 **도쿄 → 오사카 → 싱가포르** 순으로 가까운 곳을 고르세요.
> Supabase가 서울이라 왕복이 30~40ms 늘지만, 이 앱의 무거운 작업은 CPU와 AI 호출이라
> (30개 파일 기준 엑셀 6.7초·AI 112초) 비율로는 미미합니다.

Oracle Cloud 콘솔 → Compute → Instances → Create instance

| 항목 | 값 |
|---|---|
| Region | 홈 리전(권장: **ap-seoul-1** 서울 — Supabase와 같은 도시) |
| Shape | **VM.Standard.A1.Flex** (ARM Ampere) · **1 OCPU · 6GB** |
| Image | Canonical Ubuntu 24.04 (aarch64) |
| Boot volume | 기본 50GB 그대로 |
| SSH 키 | "Generate a key pair for me"로 만들고 **private key를 반드시 받아두세요**(다시 못 받습니다) |

Always Free 한도는 A1이 **전체 2 OCPU · 12GB**, 블록 스토리지 합계 200GB입니다. 위 구성이면
한도의 절반만 씁니다.

> **A1(ARM)이 "Out of capacity"로 안 만들어지는 일이 잦습니다.** 순서대로 시도하세요.
>
> 1. **AD 바꾸기** — 오류가 권하는 방법이지만 서울·도쿄는 AD가 하나뿐인 경우가 많습니다.
> 2. **잠시 뒤 재시도** — 용량은 수시로 풀립니다(새벽 시간대가 잘 납니다). 크기를 바꿔
>    (1 OCPU/6GB ↔ 2 OCPU/12GB) 시도하면 다른 호스트 풀에 걸려 되기도 합니다.
> 3. **`VM.Standard.E2.1.Micro`로 진행** — 2대까지 무료고 대개 바로 만들어집니다.
>    **1/8 OCPU · 1GB**라 느립니다(엑셀 파싱이 CPU를 쓰는 작업입니다). 1GB에서는 도커
>    빌드가 OOM으로 죽을 수 있어 `setup-vm.sh`가 **메모리 2GB 미만이면 스왑 2GB를 자동으로
>    만듭니다.**
>
> **E2로 시작해도 잃는 것이 없습니다.** 데이터는 전부 Supabase에 있고 VM은 컨테이너만
> 돌리므로, 나중에 A1이 나면 새 VM에서 `setup-vm.sh`를 한 번 돌리고 DNS(또는 접속 IP)만
> 바꾸면 됩니다.

## 2. 콘솔에서 포트 열기

**VCN → Security List → Ingress Rules**에 `0.0.0.0/0` TCP **80·443**을 추가합니다.
VM 안쪽 방화벽은 아래 스크립트가 엽니다 — **두 군데를 다 열어야** 접속이 됩니다(오라클에서
가장 많이 걸리는 지점입니다).

## 3. 한 번에 배포

VM에 접속해 스크립트를 받아 실행합니다. 방화벽(안쪽)·도커 설치·코드 내려받기·기동을 다 합니다.

```bash
ssh -i <키> ubuntu@<VM_IP>

curl -fsSL https://raw.githubusercontent.com/n0ah-031/Day11/claude/handoff-work-progress-f416b0/setup-vm.sh -o setup-vm.sh
bash setup-vm.sh                       # HTTP (도메인 없을 때)
# bash setup-vm.sh chwihap.example.com # 도메인이 있으면 자동 HTTPS
```

처음 실행하면 `.env`가 없다고 멈춥니다(키는 저장소에 없습니다 — 그게 맞습니다).
안내대로 로컬에서 복사한 뒤 다시 실행하세요:

```bash
scp -i <키> .env ubuntu@<VM_IP>:~/Day11/.env    # 로컬에서
```

`restart: unless-stopped`라 VM을 재부팅해도 자동으로 다시 뜹니다.

확인·운영:

```bash
sudo docker compose ps                 # app·caddy 둘 다 Up
sudo docker compose logs -f app        # 로그
git pull && sudo docker compose up -d --build   # 갱신(진행 중 세션은 사라집니다)
```

## 5. 첫 관리자

`/login.html`에서 가입한 뒤, 로컬(또는 VM)에서:

```bash
python3 manage_users.py promote <사번>
```

## 도메인과 HTTPS

- **도메인이 있으면** `SITE_ADDRESS`에 도메인을 넣고 A 레코드를 VM IP로 걸어두면 Caddy가
  Let's Encrypt 인증서를 자동으로 받아 갱신합니다. 이때 `COOKIE_SECURE=1`(기본값)입니다.
- **없으면** HTTP로만 뜹니다. 그 경우 **반드시 `COOKIE_SECURE=0`**으로 내려야 로그인이
  됩니다 — Secure 쿠키는 HTTPS에서만 전송되기 때문입니다. 대신 **인증 쿠키가 평문으로
  오갑니다.** 발표용 임시 운영이면 감수할 만하지만, 실제로 쓰실 거면 도메인을 붙이세요.

## 로컬 도커로 확인한 것

배포 설정만 만들고 "됐다"고 하지 않기 위해, 같은 구성을 로컬에서 띄워 확인했습니다.

- `docker compose up` → app·caddy 기동, 화면 3개 200
- `setup-vm.sh`: 문법 검사, `.env` 없을 때 안내하고 멈추는 동작, 그리고 설치할 패키지
  (`docker.io`·`docker-compose-v2`·`git`·`iptables-persistent`)가 **arm64 Ubuntu 24.04에
  실제로 있는지** 확인했습니다. 다만 방화벽·systemd 단계는 컨테이너에서 재현할 수 없어
  **실제 VM에서 처음 돌려보는 것은 남아 있습니다**
- **Caddy를 통해** 로그인 없이 `POST /api/session` → 401(인증 게이트 동작)
- **Caddy를 통해 실제 Supabase 상대로 전 구간**: 가입 → 로그인 → 업로드 → 검토(정상) →
  모드 B 취합 → 결과 다운로드(4,961바이트). 테스트 계정은 정리했습니다

실제 오라클 VM 배포는 계정이 필요해 실행하지 않았습니다.

## 알아둘 한계

- **재시작하면 진행 중이던 작업이 사라집니다**(`docker compose up -d --build`로 갱신할 때도
  마찬가지). 기록과 산출물은 Supabase에 남고, 업로드 세션·잡 진행률만 사라집니다.
- **복제하지 마세요**(`docker compose up --scale app=2` 금지). 세션이 메모리에 있는 한
  인스턴스 1개가 구조적 상한입니다.
- 업로드 상한은 파일당 50MB·합계 500MB이고 Admin 화면에서 바꿉니다. Caddy의 요청 본문
  상한(`Caddyfile`의 `max_size 500MB`)도 함께 맞춰야 합니다.
- Supabase 무료 플랜은 오래 안 쓰면 프로젝트가 일시정지됩니다. 발표 전에 한 번 접속해
  깨워두세요.
