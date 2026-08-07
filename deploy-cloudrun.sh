#!/usr/bin/env bash
# Cloud Run 배포. `.env`의 키를 Secret Manager에 넣고 서비스를 올린다.
#
# 이 스크립트가 정하는 값들에는 각각 이유가 있다(바꾸기 전에 읽어보라):
#
#   --no-cpu-throttling   이 앱은 긴 작업을 백그라운드 스레드로 돌리고 요청에는 즉시
#                         job_id만 돌려준다(server._start_job). Cloud Run 기본값(요청 기반
#                         과금)은 응답을 보낸 뒤 CPU를 조이므로 그 스레드가 멈춘다.
#                         **이 플래그를 빼면 취합이 진행되지 않는다.**
#   --max-instances=1     세션·잡이 프로세스 메모리라 인스턴스가 둘이면 요청이 다른
#                         인스턴스로 가면서 "세션을 찾을 수 없습니다"가 뜬다.
#   --min-instances=0     유휴 15분 뒤 0으로 내려간다(=잠든다). 그때 진행 중이던 세션은
#                         사라지고 끝난 결과는 Supabase에 남는다. 항상 켜두려면 1로 바꾼다
#                         (그만큼 계속 과금된다).
#   --timeout=900         2단계 AI 재검증이 파일당 20초대다. 30개 파일이면 기본 300초로 모자란다.
#   --region              도쿄. Supabase가 서울이라 가장 가깝고, Cloud Run 요금 Tier 1이라
#                         무료 범위 대상이다(서울 asia-northeast3은 Tier 2).
#   --allow-unauthenticated
#                         Cloud Run 레벨 인증을 열고 앱 자체 로그인(F4-1)을 쓴다는 뜻이다.
#                         앱은 키가 없으면 열리는 게 아니라 503으로 닫힌다(fail-closed).
#
# 실행 전 준비: gcloud CLI 설치 → `gcloud auth login` → `gcloud config set project <ID>`
# 사용법: ./deploy-cloudrun.sh [서비스명]

set -euo pipefail

SERVICE="${1:-chwihap}"
REGION="${REGION:-asia-northeast1}"
ENV_FILE="${ENV_FILE:-.env}"
SECRETS=(OPENAI_API_KEY SUPABASE_URL SUPABASE_PUBLISHABLE_KEY SUPABASE_SERVICE_ROLE_KEY)

command -v gcloud >/dev/null || { echo "gcloud가 없습니다. https://cloud.google.com/sdk/docs/install" >&2; exit 1; }
[ -f "$ENV_FILE" ] || { echo "$ENV_FILE 이 없습니다. 키는 이 파일에서만 읽습니다." >&2; exit 1; }

PROJECT="$(gcloud config get-value project 2>/dev/null)"
[ -n "$PROJECT" ] && [ "$PROJECT" != "(unset)" ] || { echo "gcloud config set project <ID> 를 먼저 하세요." >&2; exit 1; }
echo "프로젝트 $PROJECT · 리전 $REGION · 서비스 $SERVICE"

gcloud services enable run.googleapis.com cloudbuild.googleapis.com secretmanager.googleapis.com

# 키는 Secret Manager에 둔다. 환경변수로 넣으면 콘솔에서 그대로 보인다.
for name in "${SECRETS[@]}"; do
  value="$(grep "^${name}=" "$ENV_FILE" | cut -d= -f2- || true)"
  [ -n "$value" ] || { echo "$ENV_FILE 에 $name 이 없습니다." >&2; exit 1; }
  if gcloud secrets describe "$name" >/dev/null 2>&1; then
    printf '%s' "$value" | gcloud secrets versions add "$name" --data-file=- >/dev/null
  else
    printf '%s' "$value" | gcloud secrets create "$name" --data-file=- --replication-policy=automatic >/dev/null
  fi
  echo "  secret $name 갱신"
done

# 런타임 서비스 계정이 secret을 읽을 수 있어야 한다
NUMBER="$(gcloud projects describe "$PROJECT" --format='value(projectNumber)')"
RUNTIME="${NUMBER}-compute@developer.gserviceaccount.com"
for name in "${SECRETS[@]}"; do
  gcloud secrets add-iam-policy-binding "$name" \
    --member="serviceAccount:${RUNTIME}" --role=roles/secretmanager.secretAccessor >/dev/null
done

MODEL="$(grep '^OPENAI_MODEL=' "$ENV_FILE" | cut -d= -f2- || echo gpt-5-mini)"

# ENV_NAME=시크릿이름:latest 목록
PAIRS=""
for name in "${SECRETS[@]}"; do PAIRS="${PAIRS}${name}=${name}:latest,"; done
PAIRS="${PAIRS%,}"

gcloud run deploy "$SERVICE" \
  --source . \
  --region "$REGION" \
  --platform managed \
  --allow-unauthenticated \
  --port 8080 \
  --memory 1Gi \
  --cpu 1 \
  --no-cpu-throttling \
  --min-instances 0 \
  --max-instances 1 \
  --concurrency 40 \
  --timeout 900 \
  --set-env-vars "COOKIE_SECURE=1,AI_CONCURRENCY=6,OPENAI_MODEL=${MODEL}" \
  --set-secrets "$PAIRS"

URL="$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')"
echo
echo "배포 완료: ${URL}/login.html"
echo "첫 관리자는 가입 후 로컬에서: python3 manage_users.py promote <사번>"
