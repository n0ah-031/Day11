# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 저장소 구조 — 먼저 읽을 것

**`main`에 코드와 문서가 전부 있습니다.** 2026-09-08에 구현 브랜치를 `main`으로
합치고 `claude/*` 워크트리·브랜치를 걷어냈습니다 — 세션마다 워크트리가 옛 브랜치로
되돌아가 코드가 사라져 보이는 문제가 있었고, 그 원인이 워크트리였습니다.
**작업은 저장소 루트에서 `main`으로 하세요.**

저장소 루트에는 `HANDOFF.md`가 있습니다 —
**실질적인 진입점이며 CLAUDE.md보다 항상 최신입니다.** 환경 확인 절차(§0), 실행 방법(§1),
구현/미구현 범위(§2), 코드 구조(§3), 실제 업무 파일 검증 근거(§4), 다음 할 일(§5),
알고 있어야 할 상태(§6)가 들어 있습니다. 코드에 손대기 전에 §0과 §6을 읽으세요.

기준 문서는 `docs/PRD_v2.0.md`입니다(가장 최신). `docs/`에 v1.3~v2.0이 함께 있으니
번호가 가장 큰 것을 보세요.

## 명령어

의존성(패키지 매니페스트 없음 — 직접 설치):

```bash
pip install openpyxl pillow openai fastapi uvicorn python-multipart httpx pyjwt cryptography
```

웹 서버 (저장소 루트에서 실행 — `load_env`가 cwd의 `.env`를 읽습니다):

```bash
python3 -m uvicorn server:app --port 8000
```

`/login.html`로 시작합니다. Supabase 없이 취합만 보려면 `AUTH_DISABLED=1`을 붙이되,
**그 모드에서 F1 양식 생성은 503입니다**(`user["id"]`가 없음). `.claude/launch.json`의
`chwihap`(인증 끔) / `chwihap-auth`(인증 켬)가 각 설정입니다.

테스트 — pytest가 아니라 각 파일이 자체 러너입니다. 개별 시나리오만 돌리는 옵션은 없습니다:

```bash
python3 test_aggregate.py && python3 test_server.py && python3 test_hwpx.py && python3 test_formgen.py && python3 test_auth.py
```

앞의 넷(9·12·5·9 시나리오)은 네트워크·API 키 없이 돕니다(LLM 스텁, `AUTH_DISABLED=1`).
**`test_auth.py`(14 시나리오)만 실제 Supabase를 상대로 돌며** 네트워크와 `.env`의
`SUPABASE_*`가 필요합니다. 끝난 뒤 DB는 **계정 1건(`admin01`)·감사 로그 0건·버킷 양쪽 비어 있는
상태**여야 정상입니다 — 남아 있으면 teardown이 실패한 것이니 원인을 찾으세요.

CLI (웹과 같은 코드 경로를 탑니다):

```bash
python3 aggregate.py <입력폴더> --mode B --rules rules.json
python3 formgen.py "부서별 예산 집행 현황 양식 만들어줘" -o form.xlsx --rules-out rules.json
python3 hwpx_merge.py a.hwpx b.hwpx -o merged.hwpx
python3 manage_users.py promote admin01
```

`formgen.py`는 **실제 OpenAI 호출이 일어납니다**(문답 턴마다 1회 + 저작 1회).
`aggregate.py`는 `--no-ai`로 2단계를 끌 수 있습니다.

## 아키텍처

명세상 스택은 React + FastAPI + Supabase지만 **실제 구현은 CLI + FastAPI + 정적 HTML**입니다.
빌드 체인이 없습니다 — `ui/*.html`은 Tailwind CDN + 바닐라 JS이고 `server.py`가 `/api` 라우트
뒤에 `ui/`를 정적 마운트합니다. (진행 문서에 적힌 React/Vite/Zustand/SheetJS 데모는 이 브랜치의
구현이 아닙니다.)

핵심 계층:

| 파일 | 역할 |
|---|---|
| `aggregate.py` | F2 취합 엔진 + CLI (1,300줄 단일 파일 — 흐름이 선형이라 분리 이득이 없음) |
| `formgen.py` | F1 AI 양식 생성 + F6 기존 양식 인지·등록 엔진 + CLI |
| `hwpx_merge.py` | F3 hwpx(OWPML) 직접 병합. 표준 라이브러리만 씀, `aggregate.py`와 독립 |
| `server.py` | FastAPI. 세션·잡은 프로세스 메모리 `dict` + 임시 디렉터리 |
| `auth.py` | 사번 → `{사번}@internal.local` 가상 이메일로 Supabase Auth 위임. JWKS(ES256) 로컬 검증 |
| `store.py` | Supabase 영속화 (Storage `uploads`/`results`, 기록 insert/patch, 이력 조회) |

**`server.py`는 `aggregate.py`를 import만 하고 수정하지 않습니다.** `main()`과 동일한 순서
(`read_file` → `review_stage1` → 파일 단위 게이팅 → `review_stage2_many` → `preprocess` →
`synthesize` → `write_report`)를 호출하므로 CLI와 웹이 같은 경로를 탑니다. 이 대응을 깨지 마세요.

취합 검토는 2단계입니다 — 1차 규칙 검증(전량 스캔) 후, **파일 내 위반 셀이 1개라도 있으면 그
파일은 AI 호출 자체를 하지 않습니다**(부분 통과 없음). 이미지는 1차만 수행합니다.

지켜야 하는 구조적 결정:

- **`aggregate.mask`가 외부 전송 직전 유일한 통과 지점입니다.** 마스커를 새로 만들지 말고,
  전송 payload 전체(인접 컬럼·분포 샘플 포함)가 이 함수를 지나게 하세요.
- **`formgen`의 구조 판정은 `aggregate`의 헤더·표 인식을 재사용합니다.** 두 번째 판정기를 만들면
  취합이 보는 헤더와 인지가 보는 헤더가 갈라집니다.
- **LLM 산출물은 신뢰하지 않습니다.** `validate_workbook`(수식 화이트리스트·외부 참조 금지·셀 수
  상한·시트명·스타일 속성)을 통과한 것만 파일이 됩니다. 모델 키 이름이 회차마다 흔들리면
  `normalize_workbook`에 별칭을 더하되 **값을 만들어내지는 마세요** — 비었으면 비운 채로 걸리게 둡니다.
- **F6 등록은 원본 바이트를 다시 쓰지 않습니다.** 첨부로 올라온 Storage 객체를 그대로 양식 파일로
  삼습니다. openpyxl로 열고 다시 저장하는 코드를 끼우면 조건부서식·매크로가 조용히 사라집니다.
- 생성 버튼은 **서버가 내린 `spec_complete`로만** 켭니다(클라이언트 자체 판정 금지).
- **F1 항목명은 헤더 문자열 그대로**여야 합니다(`부서명*`의 `*`까지). 취합이 이 이름으로 열을 찾습니다.
- Storage 경로는 `{project_id}/` 바로 아래 평면으로 두세요 — 정리 코드가 그 규칙에 의존합니다.
- `rule_type` 어휘는 `formgen.FIELD_TYPES`와 **DB check 제약 두 곳**에 있습니다. 유형을 더하면
  마이그레이션도 함께 내야 합니다(안 하면 23514로 죽습니다). `test_auth.py`가 두 곳을 대조합니다.

## 환경·비밀값

`.env`(gitignore됨)가 **유일한 키 파일입니다.** 과거에 `.env.example`에 실제 키를 적어 커밋
직전까지 간 사고가 있어 템플릿 파일 자체를 삭제했습니다 — 다시 만들지 마세요.

`OPENAI_API_KEY` · `OPENAI_MODEL`(`gpt-5-mini`) · `SUPABASE_URL` ·
`SUPABASE_PUBLISHABLE_KEY` · `SUPABASE_SERVICE_ROLE_KEY`. 그 외 `AUTH_DISABLED`,
`AI_CONCURRENCY`(기본 6), `COOKIE_SECURE`.

- OpenAI 키가 없거나 틀려도 죽지 않습니다 — 2단계 AI 검증만 건너뛰고 `정상(AI 미검증)`으로 표시합니다.
- Supabase 키가 없으면 **인증이 열리는 게 아니라 503으로 닫힙니다**(fail-closed).
- **유료 OpenAI 호출은 사용자가 개발 중 승인했습니다(2026-08-04). 돌릴 때 토큰 사용량을 함께
  보고하세요.** 규모 감각: 양식 생성 한 바퀴 6호출·16k 토큰·약 $0.022. 단
  **2단계 AI 병렬화 벤치마크(30개 파일 실제 키)는 전에 거부되었으니 별도로 물어보세요.**

## Supabase 스키마 변경

**`supabase db push`는 쓸 수 없습니다** — 원격 전용 베이스라인(`0001`·`0002`)의 SQL이 저장소에
없어 이력 검증에서 거부됩니다. `migration repair --status reverted` 우회는 실제 적용된
마이그레이션을 미적용으로 기록하게 되므로 쓰지 않습니다. 절차:

```bash
pbcopy < supabase/migrations/<파일>.sql
supabase migration repair --status applied <타임스탬프>
```

마이그레이션 파일을 쓴 뒤 **원격 DB 적용은 사용자에게 부탁하고**(대시보드 SQL 에디터), 반영을
확인한 다음 기록만 맞춥니다.

프로젝트는 `chwihap-app`(ref `cfchnprkizcmfymlezyo`) 하나만 씁니다 — 같은 계정의 다른 프로젝트에는
무관한 시스템이 돌고 있어 건드리지 않습니다. **RLS는 켜져 있고 정책이 0개**라 `anon`·`authenticated`가
전면 거부이며 모든 접근이 백엔드(service key)를 경유합니다. 소유자 기반 정책을 추가하면 오히려
`authenticated`에게 읽기를 열어주는 셈이니, 프론트가 Supabase를 직접 호출하는 구조로 바꿀 때
함께 하세요.

## 실제 파일로 테스트할 때

`sampledata/`(gitignore됨)에 실제 업무 파일이 있습니다 — 지사별 자체점검 xlsx 8개,
`quarterly/` 분기별 실적 12개, `hwpx/` 3개. 업로드 화면에서 세 가지를 지정하지 않으면 오탐이 납니다:

| 항목 | 미지정 시 |
|---|---|
| 키 컬럼 | 자동 판정(키워드 → 값이 고유한 첫 컬럼 → 첫 컬럼) |
| 선택 입력 컬럼 | **전 컬럼 필수** — 비고처럼 비워두는 칸이 있으면 그 파일이 '이상'이 되어 취합에서 기본 제외 |
| 부서명 | 파일명 전체가 결과의 `부서` 컬럼에 들어감 |

새 계열 양식을 만나면 **원본 구조를 먼저 덤프해** 엔진 판정과 대조하세요 — 두 계열 모두 이 방식으로
원인을 갈라냈습니다.

`admin01` 소유의 남은 DB 행은 테스트 흔적이 아닐 수 있습니다(사용자가 직접 앱을 씁니다).
**소유자를 확인하기 전에 지우지 마세요.** 브라우저 확인용으로 API로 만든 계정은 teardown이 돌지
않으므로 `test_auth._delete_storage_for_owner` → `_delete_user` → `_delete_test_logs`를 직접
부르세요. 감사 로그 액션을 새로 추가하면 `_delete_test_logs` 정리 목록에도 넣어야 합니다.

## 문서 규약

PRD는 팀 공유용 개요로 유지하고, 기능별 세부 스펙은 `docs/` 아래 별도 문서로 분리합니다
(`취합기능_기술명세_v0.1.md`, `생성기능_기술명세_v0.1.md`, `인지기능_기술명세_v0.1.md`).
PRD 변경은 §10 결정 사항 로그에 항목을 남기고, 각 요구사항 셀에 `(vX.Y)` 표기로 어느 버전에서
바뀌었는지 밝히는 방식입니다. 커밋 메시지는 한국어로, **무엇을 했는지와 왜 그렇게 했는지**를
함께 적습니다.

알려진 문서 오류: `docs/design.md` §3의 `field_rules.rule_type`은 4종으로 적혀 있으나 실제는
기술명세 §3.4대로 6종입니다. 구현된 요구사항 중 PRD에 아직 올라가지 않은 것들이 쌓여 있습니다
(HANDOFF §5-7 "PRD v2.0 정리").
