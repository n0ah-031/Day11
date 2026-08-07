#!/usr/bin/env bash
# Oracle Cloud Ubuntu VM에서 한 번 실행하면 배포까지 끝난다.
#
#   ssh -i <키> ubuntu@<VM_IP>
#   curl -fsSL https://raw.githubusercontent.com/n0ah-031/Day11/claude/handoff-work-progress-f416b0/setup-vm.sh -o setup-vm.sh
#   bash setup-vm.sh                      # HTTP로 띄운다(도메인 없을 때)
#   bash setup-vm.sh chwihap.example.com  # 도메인이 있으면 자동 HTTPS
#
# 하는 일: VM 안쪽 방화벽 열기 → 도커 설치 → 코드 받기 → 컨테이너 기동.
# **VCN 보안 목록(콘솔)은 이 스크립트가 못 엽니다** — 콘솔에서 따로 80·443을 열어야 한다.

set -euo pipefail

REPO="${REPO:-https://github.com/n0ah-031/Day11.git}"
BRANCH="${BRANCH:-claude/handoff-work-progress-f416b0}"
DIR="${DIR:-$HOME/Day11}"
DOMAIN="${1:-}"

echo "==> 1/6 방화벽 (VM 안쪽)"
# Oracle의 Ubuntu 이미지는 기본 iptables 규칙이 80·443을 막는다. ufw를 쓰는 이미지도 있어
# 둘 다 처리한다. 이미 열려 있으면 그냥 지나간다.
if command -v ufw >/dev/null && sudo ufw status 2>/dev/null | grep -q "Status: active"; then
  sudo ufw allow 80/tcp && sudo ufw allow 443/tcp
else
  for port in 80 443; do
    sudo iptables -C INPUT -p tcp --dport "$port" -j ACCEPT 2>/dev/null \
      || sudo iptables -I INPUT 1 -p tcp --dport "$port" -j ACCEPT
  done
  # 재부팅 후에도 남기기. iptables-persistent가 없으면 설치한다(질문 없이)
  if ! command -v netfilter-persistent >/dev/null; then
    sudo DEBIAN_FRONTEND=noninteractive apt-get update -y
    sudo DEBIAN_FRONTEND=noninteractive apt-get install -y iptables-persistent
  fi
  sudo netfilter-persistent save
fi

echo "==> 2/6 스왑 (메모리가 작은 VM 대비)"
# E2.1.Micro는 1GB뿐이라 도커 빌드 중 OOM으로 죽는다. 2GB 미만이면 스왑을 만들어 둔다.
# A1(6GB 이상)이면 이 단계는 건너뛴다.
MEM_MB=$(free -m | awk '/^Mem:/{print $2}')
if [ "${MEM_MB:-9999}" -lt 2048 ] && [ ! -f /swapfile ]; then
  echo "  메모리 ${MEM_MB}MB → 스왑 2GB 생성"
  sudo fallocate -l 2G /swapfile || sudo dd if=/dev/zero of=/swapfile bs=1M count=2048
  sudo chmod 600 /swapfile
  sudo mkswap /swapfile >/dev/null
  sudo swapon /swapfile
  grep -q '^/swapfile' /etc/fstab || echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab >/dev/null
else
  echo "  메모리 ${MEM_MB}MB → 스왑 불필요"
fi

echo "==> 3/6 도커"
sudo DEBIAN_FRONTEND=noninteractive apt-get update -y
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y docker.io docker-compose-v2 git
sudo systemctl enable --now docker
sudo usermod -aG docker "$USER" || true      # 다음 로그인부터 sudo 없이 docker를 쓴다

echo "==> 4/6 코드"
if [ -d "$DIR/.git" ]; then
  git -C "$DIR" fetch --all --quiet && git -C "$DIR" checkout "$BRANCH" --quiet && git -C "$DIR" pull --quiet
else
  git clone --branch "$BRANCH" "$REPO" "$DIR"
fi

echo "==> 5/6 키 확인"
if [ ! -f "$DIR/.env" ]; then
  cat >&2 <<EOF

  $DIR/.env 가 없습니다. 키는 저장소에 들어 있지 않습니다(그래야 맞습니다).
  로컬에서 아래를 실행해 복사한 뒤, 이 스크립트를 다시 돌리세요:

      scp -i <키> .env ${USER:-ubuntu}@<VM_IP>:$DIR/.env

  필요한 값은 OPENAI_API_KEY · OPENAI_MODEL · SUPABASE_URL ·
  SUPABASE_PUBLISHABLE_KEY · SUPABASE_SERVICE_ROLE_KEY 입니다(HANDOFF §0.3).
EOF
  exit 1
fi

echo "==> 6/6 기동"
cd "$DIR"
if [ -n "$DOMAIN" ]; then
  # 도메인이 있으면 Caddy가 Let's Encrypt 인증서를 자동으로 받는다 → 쿠키에 Secure를 붙인다
  sudo SITE_ADDRESS="$DOMAIN" COOKIE_SECURE=1 docker compose up -d --build
  URL="https://$DOMAIN"
else
  # HTTP로만 뜬다. Secure 쿠키는 HTTPS에서만 전송되므로 꺼야 로그인이 된다.
  # 이 경우 인증 쿠키가 평문으로 오간다 — 임시 운영에서만 쓰라
  sudo SITE_ADDRESS=":80" COOKIE_SECURE=0 docker compose up -d --build
  URL="http://$(curl -fsS --max-time 5 ifconfig.me 2>/dev/null || echo '<VM_IP>')"
fi

echo
sudo docker compose ps
cat <<EOF

배포 완료: ${URL}/login.html

  · 화면이 안 열리면 콘솔에서 VCN → Security List → Ingress에 TCP 80·443이 열려 있는지 보세요
    (이 스크립트는 VM 안쪽만 엽니다)
  · 첫 관리자: /login.html 에서 가입한 뒤  python3 manage_users.py promote <사번>
  · 로그:  sudo docker compose logs -f app
  · 갱신:  git pull && sudo docker compose up -d --build   (진행 중이던 세션은 사라집니다)
EOF
