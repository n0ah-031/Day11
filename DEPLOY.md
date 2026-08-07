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
**서울 리전(ap-seoul-1)**이라 Supabase(서울)와 같은 도시이며, 잠들지 않고, 무료입니다.

## 1. VM 만들기

> **가입할 때 홈 리전을 서울로 고르세요.** Always Free 자원은 **홈 리전에서만** 무료이고,
> 홈 리전은 가입 시 정해집니다(나중에 바꾸기 어렵습니다). 다른 리전에 만들면 그냥 과금됩니다.
> 이미 다른 홈 리전으로 가입하셨다면 그 리전에 만드세요 — Supabase(서울)와 멀어져 느려지지만
> 동작은 합니다.

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

> **A1(ARM)이 "Out of capacity"로 안 만들어지는 일이 잦습니다.** 가용성 도메인(AD)을 바꿔가며
> 재시도하는 게 첫 번째 방법입니다. 그래도 안 되면 `VM.Standard.E2.1.Micro`(2대까지 무료)가
> 있는데 **1/8 OCPU · 1GB**로 성능이 많이 낮습니다 — 메모리는 데모 규모에 충분하지만(30개
> 파일 취합 실측 최대 RSS 378MB) 엑셀 파싱이 CPU를 쓰는 작업이라 체감이 느립니다. A1이 날
> 때까지 기다렸다 쓰는 편을 권합니다.

## 2. 포트 열기 — **두 군데를 다 열어야 합니다**

Oracle에서 가장 많이 걸리는 지점입니다. VCN 보안 목록만 열고 끝내면 접속이 안 됩니다.

```bash
# (A) 콘솔: VCN → Security List → Ingress Rules 에 0.0.0.0/0 TCP 80, 443 추가

# (B) VM 안의 방화벽. Ubuntu 이미지는 기본 iptables 규칙이 막고 있습니다
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 443 -j ACCEPT
sudo netfilter-persistent save
```

## 3. 도커 설치

```bash
sudo apt-get update && sudo apt-get install -y docker.io docker-compose-v2 git
sudo usermod -aG docker $USER && newgrp docker
```

## 4. 앱 올리기

```bash
git clone https://github.com/n0ah-031/Day11.git && cd Day11
git checkout claude/handoff-work-progress-f416b0

# 키는 저장소에 없습니다. 로컬 .env를 복사해 오세요(내용은 HANDOFF §0.3)
scp .env ubuntu@<VM_IP>:~/Day11/.env      # 로컬에서 실행

# 도메인이 있으면 (자동 HTTPS)
SITE_ADDRESS=chwihap.example.com docker compose up -d --build

# 도메인이 없으면 (HTTP만) — 쿠키 Secure를 꺼야 로그인이 됩니다
SITE_ADDRESS=:80 COOKIE_SECURE=0 docker compose up -d --build
```

`restart: unless-stopped`라 VM을 재부팅해도 자동으로 다시 뜹니다.

확인:

```bash
docker compose ps          # app·caddy 둘 다 Up
curl -I http://<VM_IP>/login.html
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
