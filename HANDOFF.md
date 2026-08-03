# 핸드오프: 취합 프로그램 (F2 엑셀 취합)

| 항목 | 내용 |
|---|---|
| 작성일 | 2026-08-03 (갱신) |
| 브랜치 | `claude/handoff-work-progress-f416b0` (이전: `claude/day11-handoff-continuation-2021a1`, `claude/docs-aggregation-program-31c2ae`) |
| 기준 문서 | [docs/PRD_v1.9.md](docs/PRD_v1.9.md) (최신), [docs/취합기능_기술명세_v0.1.md](docs/취합기능_기술명세_v0.1.md), [docs/design.md](docs/design.md) |
| 상태 | **로그인 → 업로드 → 검토 → 취합 → 다운로드 전 구간 동작.** 브라우저 실측 확인 완료. 인증(F4-1) 연결, 종전 미검증 2건(이미지 anchor·대량 성능) 해소 |

---

## 0. 먼저 할 일 — 본인 API 키 설정

**이 저장소에는 API 키가 없습니다.** `.env`는 `.gitignore`에 등록되어 커밋되지 않으므로, 이어받는 쪽에서 직접 만들어야 합니다.

저장소 루트에 `.env` 파일을 만들고 아래 한 줄을 넣으세요(값은 본인 OpenAI 키):

```
OPENAI_API_KEY=sk-...
```

- 키를 `.env` **외의 파일에 적지 마세요.** 이전에 템플릿 파일(`.env.example`)에 실제 키를 적어 커밋 직전까지 간 사고가 있었고, 그래서 템플릿 파일 자체를 삭제했습니다. 키 파일은 `.env` 하나뿐입니다.
- 키가 없거나 틀려도 프로그램은 죽지 않습니다. 명세 §13.3에 따라 2단계 AI 검증만 건너뛰고 `정상(AI 미검증)`으로 표시하며 취합은 정상 진행됩니다.

---

## 1. 실행 방법

```bash
pip install openpyxl openai fastapi uvicorn python-multipart
```

### 웹 (권장)

```bash
python3 -m uvicorn server:app --port 8000
```

브라우저에서 `http://localhost:8000/aggregate.html` — 업로드 → 검토·범위선택 → 합성모드 → 결과 4단계입니다. `http://localhost:8000/` 은 생성 홈이고 사이드바 '취합'으로 이동합니다.

### CLI

```bash
python3 aggregate.py <입력폴더> --mode B
```

주요 옵션:

| 옵션 | 기본값 | 설명 |
|---|---|---|
| `--mode` | `B` | 합성 모드 `A`/`B`/`C`/`D` (§9) |
| `--model` | `gpt-5-mini` | LLM 모델. `.env`의 `OPENAI_MODEL`로도 지정 가능 |
| `--no-ai` | off | 2단계 AI 재검증 생략(규칙검증만) |
| `--out` | `merged.xlsx` | 결과 파일 경로 |
| `--report` | `error_report.xlsx` | 오류 리포트 경로 |
| `--include-anomalous` | off | 오류 등급 파일까지 강제로 취합 포함(§7) |
| `--rules` | 없음 | 작성기준 JSON 경로(§5.4 확장 스키마). 키 컬럼 `"key": true`, 선택 입력 컬럼 `"required": false` |
| `--group-map` | 없음 | 모드 C 그룹 매핑 JSON `{시트명: 그룹명}` |
| `--summary-cols` | 없음 | 모드 D 요약 지표 컬럼(콤마 구분) |
| `--include-hidden` | off | 숨김 시트도 포함 |

자체 점검:

```bash
python3 test_aggregate.py
python3 test_server.py
python3 test_auth.py
```

앞의 두 개는 7개·6개 시나리오가 전부 통과해야 정상이며 **네트워크·API 키 없이** 돌아갑니다(AI 계층은 스텁, 인증은 `AUTH_DISABLED=1`로 끔).

`test_auth.py`는 다릅니다 — **실제 Supabase 프로젝트를 상대로** 10개 시나리오를 돌아 네트워크와 `.env`의 `SUPABASE_*`가 필요합니다. 테스트 계정은 매 실행 만들고 지웁니다(사번 접두사 `zz-test-`).

---

## 2. 구현 범위 — 무엇이 되고 무엇이 안 되는지

### 구현 완료

| 명세 | 구현 내용 |
|---|---|
| §2 업로드·구조 인식 | 시트 열거, 헤더 추정(병합 제목행 skip + 다음 행 타입 대비 검증), 데이터 영역 인식(공백 2행 허용), 숨김 시트 제외 |
| §2.1 오류 안내 | 포맷 오류·헤더 미발견·빈 시트·데이터 없음·파일 손상 5종을 사유별 메시지로 구분 |
| §3 유형 분류 | 작성기준 우선 → 키워드(국/영문) 1차 후보 → 값 타입 샘플링 최종 판정. 불확실 시 자유서술형 |
| §4.1 1단계 규칙검증 | **전량 스캔**(첫 위반에서 중단하지 않음), 누락 → 정합성 → 중복/충돌 순서 |
| §4.2 파일 단위 게이팅 | 파일 내 위반 셀이 1개라도 있으면 그 파일은 AI 호출 자체를 하지 않음 |
| §4.2 컨텍스트 규칙 | 코드표·자유서술형은 인접 컬럼 포함, 패턴 기반은 미포함, 정량적은 같은 컬럼 분포 통계 포함 |
| §4.2.1 개별 파싱 실패 | `이상(AI 재검증 실패)` + `[확인필요]` 태그, 내부 등급은 오류 |
| §5 유형별 규칙 | 정성(코드표/패턴/자유서술) · 정량(타입·범위·자릿수) · 이미지(확장자·용량·해상도·해시중복·anchor 범위) |
| §6 이상판별 출력 | 정상/이상 이진 표시 + 사유에 `[경미]`/`[심각]`/`[확인필요]` 태그. 검증 단계는 화면 비표시 |
| §7 대표 등급 | 파일 내 최악 등급 채택(오류 > 경고 > 정상), 오류는 기본 제외 |
| §8 전처리 | 경고 항목 자동교정 + **원본값 보존**(감사 추적). AI 발견 이상은 자동교정 제외 |
| §9 합성 A/B/C/D | 모드 D 요약은 SUMIF 수식으로 생성(하드코딩 아님), 컬럼 부재 시 `N/A` |
| §10 리포트 | 2시트 구성 — `오류 목록`(검증 단계 구분 포함) + `자동교정 이력` |
| §13.2 마스킹 | 주민번호·전화·이메일 자동 치환. **전송 payload 전체**에 적용(인접 컬럼·분포 샘플 포함) |
| §13.3 장애 폴백 | AI 서비스 장애 시 2단계 생략 → `정상(AI 미검증)` |
| F2-13 키 컬럼 지정 | 업로드 단계에서 작업 단위 전역 키 컬럼 1개 선택(기본 `자동`). 작성기준 `{헤더명: {"key": true}}`로 전달 (PRD v1.8 신규) |
| 선택 입력 컬럼 지정 | 업로드 단계에서 비워둬도 되는 컬럼을 다중 선택. 작성기준 `{헤더명: {"required": false}}`로 전달. 미지정 시 종전대로 전 컬럼 필수 (PRD v1.9 F2-14) |
| 진행률 폴링 | 업로드 파싱·검토·취합이 즉시 `{job_id}`를 돌려주고 `GET /api/job/{jid}`로 단계·진행률을 폴링. 잡 상태는 프로세스 메모리 (PRD v1.9 F2-15) |
| F4-1 사번 로그인 | 사번 → `{사번}@internal.local` 가상 이메일로 Supabase Auth 위임. 사번은 대소문자 무관. 토큰은 httpOnly·SameSite=Lax 쿠키, 검증은 JWKS(ES256) 로컬. 취합 API 전체에 인증 + 세션·잡 소유자 격리(§6.2) |
| 영속화 | 업로드 xlsx는 Storage `uploads`, 결과·리포트는 `results`. 기록은 `projects`·`uploaded_files`·`review_results`·`aggregation_jobs`. 취합 시작 시 프로젝트 자동 생성 |
| F4-2 작업 이력 | [ui/history.html](ui/history.html) — 검색(작업명·파일명)·유형 필터(전체/엑셀/한글), 과거 결과 재다운로드. 세션이 사라진 뒤에도 동작 |
| F4-3 Admin 콘솔 | [ui/admin.html](ui/admin.html) — 지표 카드, 계정 권한·정지·삭제, 보관 기간, 사용 로그(검색). 관리자만 접근하고 본인 계정은 스스로 잠그거나 지울 수 없다 |
| 감사 로그 | 계정 변경·삭제, 정책 변경, 취합 완료를 `audit_logs`에 남긴다. 행위자 계정이 지워지면 `actor_id`만 NULL이 되고 기록은 남는다 |

### 미구현 (의도적 제외)

명세는 React + FastAPI + Supabase 풀스택이지만, 지금은 **CLI + FastAPI + 정적 HTML**입니다. React 빌드 체인과 Supabase를 빼고 취합 코어를 먼저 실제로 동작시키는 쪽을 택했습니다.

| 미구현 | 사유 / 다음 단계 |
|---|---|
| 작업 재개 | 재시작 후 **기록과 산출물은 남지만** 중단된 취합을 그 자리에서 이어서 하지는 못한다. 파싱된 작업 집합이 메모리에 있어 되살리려면 파일을 다시 내려받아 다시 읽어야 한다 |
| 진행 중 잡 상태 | `aggregation_jobs`에 남기지만 폴링용 상태는 프로세스 메모리라, 재시작하면 진행률 추적이 끊긴다(기록은 남음) |
| 업로드 제한 화면 조정 | 파일당 50MB·30개·500MB는 `server.py` 상수다. Admin 화면에 값만 보여주고 바꾸는 기능은 없다 |
| 보관 기간 자동 삭제 | 값은 `retention_policy`에 저장되지만 기간이 지난 자료를 실제로 지우는 배치가 없다 |
| hwpx 병합(F3) | 미착수. 명세상 별도 기능 |
| 이미지 AI Vision | 명세 §14에서 v2 백로그로 지정된 항목 |
| 파일 간 교차 중복 | 명세 §14 v2 백로그. 중복 판정은 파일 내부로 한정 |

---

## 3. 코드 구조

| 파일 | 역할 |
|---|---|
| [aggregate.py](aggregate.py) | 취합 엔진 + CLI. 아래 표 참조 |
| [server.py](server.py) | FastAPI. 세션·잡 모두 프로세스 메모리 `dict` + 임시 디렉터리라 **서버 재시작 시 소실**. 긴 작업은 `_start_job`으로 스레드에 넘기고 `GET /api/job/{jid}` 폴링. `/api` 라우트 뒤에 `ui/`를 정적 마운트 |
| [ui/aggregate.html](ui/aggregate.html) | 취합 4단계 단일 페이지. Tailwind CDN + 바닐라 JS, 빌드 없음. 토큰·셸은 `index.html`에서 그대로 이식 |
| [auth.py](auth.py) | F4-1 인증. 가상 이메일 치환·가입·로그인·JWT 검증. 가상 이메일은 이 모듈 밖으로 안 나간다 |
| [manage_users.py](manage_users.py) | 계정 CLI(`list`/`promote`/`demote`/`suspend`/`activate`). Admin 화면이 생겼으니 첫 관리자 지정·복구용으로만 남긴다 |
| [store.py](store.py) | Supabase 영속화. Storage 업로드·다운로드, 기록 insert/patch, 이력 조회 |
| [ui/login.html](ui/login.html) | 로그인·회원가입 (design.md 화면 1). aggregate.html과 토큰·테마 셸 공유 |
| [ui/history.html](ui/history.html) | 작업 이력 (design.md 화면 9). 검색·유형 필터·재다운로드 |
| [ui/admin.html](ui/admin.html) | Admin 콘솔 (design.md 화면 10). 지표·계정·정책·사용 로그 |
| [test_auth.py](test_auth.py) | 인증·영속화·이력·Admin 10개 시나리오. **실제 Supabase를 상대로 돌고 네트워크가 필요하다** |
| [test_server.py](test_server.py) | API 6개 시나리오(end-to-end / 강제 포함 / 업로드 거부 / 키 컬럼 / 선택 입력 컬럼 / 잡 진행률) |

`server.py`는 `aggregate.py`를 import만 하고 수정하지 않습니다. `main()`과 동일한 순서(`read_file` → `review_stage1` → 게이팅 → `review_stage2_many` → `preprocess` → `synthesize` → `write_report`)를 호출하므로 CLI와 웹이 같은 코드 경로를 탑니다.

### aggregate.py 내부

| 영역 | 역할 |
|---|---|
| `load_env` | `.env` → 환경변수 (표준 라이브러리만, python-dotenv 미사용) |
| `Issue` / `Fix` / `SheetData` / `UploadedFile` | 판정 결과 데이터 구조. 명세 §11의 `ReviewResult`/`ErrorItem`에 대응 |
| `read_file` · `_find_header` · `_read_images` | §2 구조 인식 (헤더·데이터 영역·이미지 anchor) |
| `classify_columns` | §3 유형 분류 |
| `review_stage1` · `_validate_text` · `_validate_number` · `_review_duplicates` · `_review_images` | §4.1·§5 유형별 1단계 규칙 |
| `mask` | §13.2 마스킹. **외부 전송 직전 공통 통과 지점** |
| `_ai_items` · `_ai_call` · `review_stage2_many` · `review_stage2` | §4.2 2단계 배치 재검증 + §4.2.1 부분 실패 격리. `_many`가 (파일, 배치)를 평평한 목록으로 만들어 단일 동시성 상한으로 병렬 호출한다 |
| `preprocess` | §8 경고 항목 자동교정 + 원본값 보존 |
| `synthesize` · `_add_summary_sheet` · `_shift_anchor` | §9 합성 모드 A/B/C/D + 이미지 anchor 재배치 |
| `write_report` | §10 리포트 2시트 |
| `main` | CLI 조립 |

단일 파일인 이유: 취합 흐름이 선형이라 모듈 분리 이득이 없었습니다.

---

## 4. 검증 근거

- `python3 test_aggregate.py` — 7개 시나리오 통과 (구조인식·전량스캔 / AI 폴백 / 마스킹 / 합성 4모드 / 이미지 anchor 재배치 / 2단계 병렬·실패격리 / CLI end-to-end)
- 실제 OpenAI 호출 검증(`gpt-5-mini`): 1단계를 전부 통과하지만 내용이 문맥상 틀린 샘플로 확인
  - 무관한 서술(`"오늘 점심은 김치찌개가…"` in 사업내용) → 부적합 판정
  - 이상치 금액(980,000,000 vs 다른 행 수백만원) → 부적합 판정
  - 인접 컬럼 교차 검증: 사업내용↔예산액 불일치를 양방향으로 지적 → §4.2 인접 컨텍스트 동작 확인
  - 세 건 모두 경고 등급이라 취합은 차단되지 않음(2/2건 포함) → §6.2 오탐 방어 동작 확인

- `python3 test_server.py` — 6개 시나리오 통과
- **브라우저 실측**(uvicorn + 실제 xlsx 3개, API 키 없이): 업로드 시 시트·헤더행·행수 인식 → 검토에서 정상 2 / 이상 1, 오류 파일만 기본 해제(선택 2/3건) → 모드 D 취합 → `종합요약` 시트가 `=SUMIF('예산'!A2:A4,$A2,'예산'!E2:E4)` 수식으로 생성됨(§9 하드코딩 금지 충족) → 결과·리포트 다운로드 200
- 오류 파일까지 강제 체크(3/3건) → §10.1 **특이사항** 안내가 셀 위치까지 표시됨
- 라이트/다크 양쪽 렌더 확인, 콘솔 에러 0건
- **키 컬럼(F2-13)**: 첫 컬럼이 부서명처럼 반복되는 xlsx로 `자동` 검토 → '이상'(키 충돌 오탐), 같은 파일에 고유 키 컬럼 지정 → '정상'. 화면에서 전환 확인
- **2단계 AI 실측(v1.9 회차, `gpt-5-mini`)**: 1단계를 통과하는 4행 파일 1개에 **16.0초**. 무관한 서술·인접 컬럼 불일치·이상치 금액을 모두 경미 등급으로 지적해 취합은 차단되지 않음
- **모드 C**: 4개 시트(1~4월)×2파일을 `1·2월→상반기 / 3월→하반기 / 4월 미매핑`으로 취합 → 상반기 9행·하반기 5행·미분류 5행, `구분` 컬럼에 원본 시트명 보존, 미매핑 안내 표시. API·UI 양쪽 확인

### 이번 회차에 해소한 미검증 항목

**이미지 anchor 재배치(§9)** — 실측 결과 3건 모두 깨져 있어 수정했습니다([aggregate.py](aggregate.py) `_shift_anchor`).

| 증상 | 원인 |
|---|---|
| 모드 A에서 사진이 2행 아래 엉뚱한 데이터 행에 붙음 | 출력은 헤더를 1행으로 당겨 쓰는데 anchor를 안 옮겼음(주석은 "이동 없음"이라 단정) |
| 모드 B/C/D에서 증빙사진이 예산액 칸 위에 얹힘 | 앞에 끼워 넣는 `구분`·`부서`만큼 열이 밀리는데 열 오프셋 미적용 |
| 셀에 맞춰 줄여 놓은 사진이 원본 픽셀 크기로 되살아남 | 표시 크기는 anchor의 `ext`에 있는데 openpyxl의 `im.width/height`는 원본 픽셀을 돌려줌 |

두 셀 앵커는 `OneCellAnchor`로 뭉개지 않고 그대로 재구성합니다(점유 범위·셀내 EMU 오프셋·`editAs` 보존). 실측: 원본 `D4:F6` → 모드 A `D2:F4`, 모드 B `E2:G4`. `test_aggregate.py`에 시나리오 추가.

**대량 파일 성능** — 30개 파일 / 총 99,990행 / 3.4MB, AI 계층 제외, M-series macOS 기준:

| 단계 | 소요 |
|---|---|
| `read_file`(구조 인식) | 3.48s (52%) |
| `save`(결과 쓰기) | 2.02s (30%) |
| `synthesize`(모드 B) | 0.93s (14%) |
| `review_stage1` | 0.16s |
| `preprocess` | 0.11s |
| **합계** | **6.71s** (최대 RSS 378MB, 오탐 0건) |

82%가 openpyxl의 xlsx 파싱·직렬화라 미세 최적화 여지가 없습니다.

웹 경로 실측(업로드→검토→취합)은 **6.9초**입니다. 검토 단계가 0.0초인 것은 종전에 업로드에서 한 번, 검토에서 또 한 번 하던 재파싱을 없앴기 때문입니다(취합이 돌아 `preprocess`가 셀을 고친 뒤에만 원본을 다시 읽습니다).

> **측정 주의** — 이 표의 초기값(합계 36.3초)은 `tracemalloc`을 켠 채 잰 값이라 5배 넘게 부풀려져 있었습니다. 파이썬에서 openpyxl 같은 할당 집약적 코드를 잴 때 `tracemalloc`은 측정 자체를 크게 왜곡합니다. 메모리는 `resource.getrusage`의 RSS로 따로 재세요.

**2단계 AI가 실제 지배적 비용** — 실제 키(`gpt-5-mini`)로 4행짜리 파일 1개를 재검증하는 데 **16.0초**가 걸렸습니다. 1단계 규칙검증 전체(0.16초)의 100배입니다. 즉 **진행률 표시가 필요한 이유는 엑셀 처리(7초)가 아니라 AI 재검증**입니다.

그래서 병렬화했습니다([aggregate.py](aggregate.py) `review_stage2_many`). (파일, 배치)를 하나의 평평한 작업 목록으로 만들어 **단일 동시성 상한**으로 돌립니다 — 파일별로 스레드 풀을 중첩하면 동시 호출 수를 통제할 수 없습니다. 동시 호출 수는 `AI_CONCURRENCY`(기본 6)로 조절합니다.

실패 격리는 파일 단위로 유지합니다(§13.3) — 어느 배치가 터지면 그 파일만 `정상(AI 미검증)`이 되고 다른 파일 판정은 살아 있습니다.

**미측정**: 실제 키로 돌린 순차 대비 병렬 배수. 회귀 테스트는 `_ai_call`을 지연·실패 스텁으로 바꿔 호출 스케줄만 결정적으로 검증합니다(네트워크·비용 없음). 실제 배수는 업무 파일로 한 번 재보는 편이 정확합니다 — 모델 응답 시간과 계정 rate limit에 달려 있습니다.

**미검증(남음)**: 실제 업무 서식에서의 헤더 인식 정확도(실제 파일 필요), 30개 규모의 2단계 전체 소요(1개 16초 기준 추정치만 있음).

---

## 5. 이어받는 쪽에 권하는 순서

1. `.env`에 키를 넣고 `python3 test_aggregate.py && python3 test_server.py && python3 test_auth.py`로 환경 확인
2. uvicorn 띄우고 **실제 업무 엑셀**로 한 번 돌려서 헤더 인식이 실제 서식에서 맞는지 확인 — 구조 인식은 실제 파일에서 가장 깨지기 쉬운 부분입니다
3. **업로드 화면에서 키 컬럼과 선택 입력 컬럼을 반드시 지정하세요.** 실제 서식에서 오탐이 나는 지점은 지금까지 이 둘뿐이었습니다
   - 키 컬럼: 지정이 없으면 **첫 컬럼**이 키입니다([aggregate.py](aggregate.py) `_pick_key_column`). 첫 컬럼이 부서명처럼 행마다 반복되면 전 행이 키 충돌로 '이상' 처리됩니다
   - 선택 입력 컬럼: 지정이 없으면 **전 컬럼이 필수**입니다. 비고·특이사항처럼 비워두는 칸이 있으면 그 파일 전체가 '이상'이 되어 취합에서 기본 제외됩니다
   - CLI에서는 `--rules` JSON의 `"key": true` / `"required": false`로 지정합니다
4. **2단계 AI 병렬화 실측** — 병렬 경로는 들어갔지만(§4 참조) 실제 키로 순차 대비 배수를 재보지 않았습니다. 업무 파일 여러 개로 한 번 재고, 느리면 `AI_CONCURRENCY`를 올려보세요(기본 6). 계정 rate limit에 걸리면 낮추면 됩니다
5. **첫 관리자 지정** — 계정 관리는 이제 `/admin.html`에서 하지만, 첫 관리자는 화면에 들어갈 수 없으니 CLI로 올립니다.

   ```bash
   python3 manage_users.py promote admin01
   ```

   비밀번호는 이 도구가 다루지 않습니다 — 가입은 `/login.html`에서 본인이 합니다. 사번은 대소문자를 구분하지 않습니다

6. **RLS 정책은 지금 쓸 일이 아닙니다.** RLS는 켜져 있고 정책이 0개라 `anon`·`authenticated`는 전면 거부이며, 모든 접근이 백엔드(service key)를 경유합니다. 공개 키로 조회하면 전 테이블이 빈 배열임을 테스트가 지킵니다. 소유자 기반 정책을 쓰면 오히려 `authenticated`에게 읽기를 **열어주는** 셈이라, 프론트가 Supabase를 직접 호출하도록 구조를 바꿀 때 함께 쓰는 것이 맞습니다
7. **Supabase** — 프로젝트는 정해졌습니다. 새 계정의 `chwihap-app`(ref `cfchnprkizcmfymlezyo`, ap-northeast-2 서울)이며 `supabase link` 완료 상태입니다

   기존 `ksw1727@gmail.com's Project`는 쓰지 않습니다. 그 `public` 스키마에는 사고 보고 시스템이 19개 테이블로 돌아가고 있어 섞이면 안 됩니다. Preview branch는 병합하면 결국 같은 production DB로 들어가 격리 수단이 못 되고, 전용 스키마를 써도 `supabase_migrations` 이력을 공유해 두 저장소의 `db push`가 간섭합니다 — 별도 프로젝트만이 완전히 분리됩니다

   **스키마는 이미 있습니다.** 저장소 밖(대시보드)에서 만들어진 것으로 보이며 명세의 테이블 10개(`aggregation_jobs`·`review_results`·`uploaded_files`·`form_templates`·`field_rules`·`group_mappings`·`projects`·`profiles`·`audit_logs`·`retention_policy`)와 버킷 `uploads`·`results`가 있습니다. 데이터는 전부 0건입니다

   **보안 결함을 하나 고쳤습니다** — 그 10개 테이블 전부 RLS가 꺼진 채 `anon`에 SELECT·INSERT·UPDATE·DELETE 권한이 있었습니다. anon 키는 프론트엔드에 실려 공개되는 값이라 키만 있으면 누구나 전 테이블을 읽고 지울 수 있는 상태였고, `supabase db advisors`도 10건 전부를 ERROR/EXTERNAL로 지적했습니다. [supabase/migrations/20260803065832_enable_rls.sql](supabase/migrations/20260803065832_enable_rls.sql)로 RLS를 켰고 재진단 ERROR 0건입니다. 정책은 인증(F4-1)이 들어와 접근 모델이 정해진 뒤에 씁니다

8. **미해결 — 마이그레이션 baseline**. 원격 이력의 `0001`·`0002`는 SQL이 저장소에 없어 remote-only로 남아 있습니다. 저장소만으로 DB를 재구축할 수 없다는 뜻입니다. `supabase db pull`이 Docker를 요구하는데 이 머신에 Docker Desktop이 없습니다 — 설치 후 `supabase db pull baseline --linked`로 한 번 떠두면 해소됩니다

9. **CLI 로그인 주의** — `supabase login`이 기본 프로필의 토큰을 새 계정으로 덮었습니다. 기존 ksw1727 계정을 다시 쓰려면 재로그인이 필요합니다. 또 `~/.supabase/profile`이 설정 파일 없는 프로필명(`chwihap`)을 가리켜 `db query`가 `failed to read profile`로 죽길래 `profile.disabled`로 옮겨뒀습니다(되돌리려면 파일명만 복구)
