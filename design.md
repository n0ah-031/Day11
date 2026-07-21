# Design: 취합앱 (자료 취합 전용 프로그램)

| 항목 | 내용 |
|---|---|
| 문서 버전 | v1.1 |
| 작성일 | 2026-07-21 |
| 기준 PRD | 취합앱_PRD_v1.3.md |
| 상태 | 초안 (User 리뷰 대기) |
| 목적 | PRD v1.3에서 design.md 단계로 위임한 상세 아키텍처·데이터모델·API·화면별 UI/인터랙션 명세, 그리고 잔여 리스크 노트 2건(모드 C 그룹 매핑 관리 방식, 모드 D 요약 지표 정의 방식)의 확정 결과를 기록 |

> **v1.1 변경 사항**: Stitch로 생성한 홈 화면 비주얼 시안 5종 중 라이트 모드는 **컨셉 02 파스텔 프렌들리**, 다크 모드는 **컨셉 03 다크 모드 프로페셔널**로 확정. §6에 공통 디자인 시스템 및 라이트/다크 모드 토글 UI 스펙 추가(§6.0).

---

## 1. 아키텍처 개요

### 1.1 전체 구성

```
[React(Vite/TS) SPA] ──HTTPS/JWT──▶ [FastAPI 백엔드]
                                        │
                        ┌───────────────┼──────────────────┐
                        ▼               ▼                  ▼
                 [Supabase Postgres] [Supabase Storage] [비동기 작업 처리]
                  (사용자/이력/메타)   (업로드·결과 파일)   (FastAPI BackgroundTasks
                                        │                   + jobs 테이블 폴링)
                                        ▼
                              [OpenAI Adapter 계층] ──▶ OpenAI GPT API
```

- **프론트엔드**: React + Vite + TypeScript. 서버 상태는 TanStack Query로 관리해 FastAPI를 호출.
- **백엔드**: Python + FastAPI. 엑셀 구조 인식(openpyxl), hwpx(OWPML, zip+xml) 직접 파싱, 규칙 기반 검토 로직을 모두 이 계층에서 처리.
- **데이터/스토리지**: Supabase 단일 사용.
  - Postgres: 계정, 프로젝트, 양식/작성기준, 검토결과, 취합 작업, 로그 등 모든 메타데이터.
  - Storage: 업로드 원본 파일과 결과물(양식 파일, 취합 결과, hwpx 병합 결과, 오류 리포트) 바이너리.
  - Auth: 로그인/세션 발급을 위임 (§2 참조).
- **AI 연동**: OpenAI GPT API를 어댑터 계층(`ai/adapter.py` 형태) 뒤에 감춰 벤더 교체 가능성 확보(PRD §6.5).
- **비동기 처리**: 엑셀 대량 처리(F2)·hwpx 병합(F3)은 별도 인프라(Redis/Celery) 없이, `aggregation_jobs` 테이블을 상태 저장소로 쓰고 FastAPI `BackgroundTasks`로 실행. 해커톤 일정상 인프라를 최소화하는 트레이드오프이며, 서버 재시작 시 재시도/동시성 제어는 MVP 범위에서 직접 구현하지 않음(§8 리스크 참조).

### 1.2 기술 스택 요약

| 영역 | 선택 | 비고 |
|---|---|---|
| 프론트엔드 | React 18 + Vite + TypeScript, TanStack Query | SPA, 반응형 웹으로 모바일 대응(PRD §7 범위 제외: 모바일 전용 앱 아님) |
| 백엔드 | Python 3.x + FastAPI | 엑셀/hwpx 처리 생태계(openpyxl 등) 활용 |
| DB/Storage/Auth | Supabase (Postgres + Storage + Auth) | 인프라 구성 최소화 |
| AI | OpenAI GPT API (어댑터 계층 경유) | PRD §6.5, zero data retention 조건 |
| 비동기 처리 | Postgres `aggregation_jobs` 테이블 + FastAPI BackgroundTasks | 추가 인프라(Redis 등) 없이 진행률/상태 폴링 구현 |

---

## 2. 인증 및 계정 설계 (F4-1)

PRD는 로그인 ID로 **사번**을 요구하지만 Supabase Auth는 기본적으로 이메일/비밀번호 기반이다. 다음과 같이 매핑한다.

- 회원가입 시 사용자가 입력하는 값: **사번**, **비밀번호**, **비밀번호 재설정용 이메일**(`reset_email`).
- 백엔드는 `{사번}@internal.local` 형태의 **가상 이메일**을 생성해 Supabase Auth의 `email/password` 가입 API를 그대로 호출한다.
- `profiles` 테이블에 `employee_no`(사번, unique) ↔ `auth.users.id` ↔ `reset_email` 매핑을 저장한다.
- 로그인 시 클라이언트는 사번/비밀번호만 입력하고, 백엔드가 사번을 가상 이메일로 치환해 Supabase Auth에 위임한다. 프론트는 사번을 직접 알 필요가 없으며 항상 백엔드를 경유한다.
- 비밀번호 재설정은 가상 이메일이 아닌 `profiles.reset_email`로 별도 발송(Supabase Auth의 기본 재설정 메일 발송 대상을 재설정용 이메일로 오버라이드).
- 가입 제한 없음(자유 가입), Admin이 `profiles.status`(active/suspended)로 사후 통제(F4-1, 잔여 리스크: 외부인 가입 모니터링은 운영 절차로 별도 수립).
- 세션: Supabase Auth가 발급하는 JWT를 그대로 사용, FastAPI는 요청마다 JWT를 검증하고 `profiles.role`로 권한(User/Admin) 분기.

---

## 3. 데이터 모델

| 테이블 | 주요 컬럼 | 설명 |
|---|---|---|
| `profiles` | id(=auth.users.id), employee_no(unique), reset_email, role(user/admin), status(active/suspended), created_at | Supabase Auth 사용자와 1:1, 사번↔가상이메일 매핑 |
| `projects` | id, owner_id, name, created_at | 양식 생성 1건 = 프로젝트 1개. 양식·작성기준·취합 작업이 귀속되는 단위 |
| `form_templates` | id, project_id, spec_json, file_url, created_at | F1 결과물(문답형으로 확정된 스펙 + 생성된 hwpx/xlsx) |
| `field_rules` | id, form_template_id, field_name, rule_type(date/name/amount/dept_code), rule_config_json | 필드별 작성 기준(F1-5), F2-5 전처리에서 참조 |
| `uploaded_files` | id, project_id, uploader_id, storage_path, original_name, kind(excel/hwpx), status, created_at | 업로드 파일 메타(바이너리는 Storage) |
| `review_results` | id, file_id, badge(정상/경고/오류), issue_count, issues_json | F2-2~F2-3 검토 결과. `issues_json`은 유형/건수/위치/사유 배열 |
| `group_mappings` | id, project_id, sheet_name, group_name | **모드 C 확정**: 화면 UI로 관리, 프로젝트 단위로 저장·재사용 |
| `aggregation_jobs` | id, project_id, kind(excel/hwpx), mode(A/B/C/D, hwpx는 null), status(queued/running/done/failed), included_file_ids, summary_fields_json, result_url, progress, created_at | F2 취합 + F3 병합 공용 작업 테이블. **모드 D 확정**: `summary_fields_json`은 실행마다 사용자가 지정한 값을 그대로 저장(고정값 아님) |
| `audit_logs` | id, actor_id, action, target, created_at | Admin 접근 로그 및 전체 사용 로그(§6.2) |
| `retention_policy` | key, value(기본 90일), updated_by | Admin이 변경 가능한 보관 정책(F4-3) |

**보관/삭제**: `uploaded_files`와 `aggregation_jobs.result_url`이 가리키는 Storage 객체는 `retention_policy` 기준(기본 90일) 경과 시 정리 배치(Supabase `pg_cron` 스케줄 함수)로 삭제하고 메타 행은 소프트 삭제(`deleted_at`) 처리. hwpx 병합용 임시 파일은 작업 완료 즉시(동기적으로) 삭제.

---

## 4. 핵심 파이프라인 설계

### 4.1 F1. AI 양식 생성 — 문답형 정보 수집

상태 머신: `INTAKE`(순차 질문) → `SPEC_COMPLETE`(생성 버튼 활성) → `GENERATING` → `DONE`

- 대상 부서, 포함 항목뿐 아니라 **항목별 입력 형식(날짜 표기, 숫자 단위 등)과 필수/선택 여부**까지 확인해야 `SPEC_COMPLETE`로 전환된다(F1-2).
- 매 턴: 프론트가 사용자 응답을 백엔드로 전달 → 백엔드가 누적 `spec_json` + 최신 응답을 OpenAI 어댑터에 전달 → 다음 질문 또는 완료 판정을 받아 프론트에 반환.
- 프론트는 우측 패널에 `spec_json`을 필드별 카드로 실시간 렌더링. `SPEC_COMPLETE` 전에는 생성 버튼 비활성.
- 완료 후 OpenAI에 최종 spec으로 xlsx/hwpx 렌더링을 요청, 결과를 `form_templates` + `field_rules`로 저장.
- 대화형 수정(F1-7)은 생성 완료 후 별도 턴으로, 기존 `form_template`을 갱신.
- 문서 파일 첨부(F1-3)는 `INTAKE` 단계에서 컨텍스트로만 사용, 별도 저장은 하지 않음(양식 생성 참고 자료일 뿐 이력 관리 대상 아님).

### 4.2 F2. 엑셀 취합

파이프라인: 업로드 → 구조 인식 → 검토 → 배지 산정 → 사용자 범위 선택 → 전처리 → 합성(A/B/C/D) → 다운로드

- **검토/배지 산정**: PRD F2-3 확정 기준표(오류=적색/필수값누락·키충돌·구조불일치, 경고=황색/자동변환가능·완전중복·비필수누락, 정보=회색/자동수정내역)를 규칙 엔진으로 그대로 구현. `field_rules`가 있으면 그 기준으로, 없으면 사용자가 그 자리에서 기준을 지정(F2-5).
- **범위 선택**: 기본 체크 상태(오류=해제, 정상/경고=체크)는 서버가 `review_results.badge`로 계산해 내려주고, 체크박스 토글은 프론트 상태로 처리 후 취합 실행 요청 시점에만 `included_file_ids`로 서버에 전달(매 클릭마다 서버 호출하지 않음).
- **합성 모드**:
  - A(원본 보존형)·B(시트명 통합형): 추가 입력 없이 바로 실행.
  - C(그룹 통합형): 화면에서 `group_mappings`를 편집(시트명→그룹명). 매핑 없는 시트는 "미분류" 그룹으로 자동 편입 후 경고 표시.
  - D(원본보존+요약형): 실행 시점마다 사용자가 요약 대상 지표(컬럼)를 선택 → `aggregation_jobs.summary_fields_json`에 저장, 종합요약 시트는 SUMIF/SUMIFS 수식으로 생성(하드코딩 금지, 지정 컬럼 부재 시 "N/A").
- **오류 리포트(F2-12)**: `review_results.issues_json`을 파일명·시트·셀 위치·오류유형·사유·작성기준 열을 포함한 xlsx로 직렬화.

### 4.3 F3. 한글(hwpx) 병합

- 업로드 → 프론트에서 드래그로 순서 지정(실제 배열 재정렬, 정적 그립 아님) → `aggregation_jobs`(kind=hwpx) 생성 → FastAPI BackgroundTasks가 OWPML(zip+xml)을 직접 파싱·병합.
- `progress` 필드를 처리 중 갱신, 프론트는 폴링으로 진행률 표시 및 완료 알림.
- M1 PoC에서 복잡 서식(개체·스타일 충돌) 병합 품질이 기준 미달일 경우 상용 SDK 전환 게이트 판단(PRD §6.5, §8) — 본 design.md는 1차로 직접 파싱 구현을 전제로 함.

### 4.4 비동기 작업 공통 설계

- `aggregation_jobs`가 F2(대량 처리 시 동기 3분 초과분)·F3 공통 작업 테이블.
- FastAPI가 `BackgroundTasks`로 작업을 실행하며 진행 중 `status`/`progress`를 갱신. 클라이언트는 `GET /aggregation-jobs/{id}`로 폴링.
- 서버 프로세스 재시작 시 진행 중이던 작업은 재시도되지 않고 `failed`로 남는다(MVP 범위, §8 리스크 참조).

---

## 5. API 설계 (엔드포인트 개요)

| 그룹 | 엔드포인트 | 설명 |
|---|---|---|
| 인증 | `POST /auth/signup` | 사번/비밀번호/재설정이메일로 가입 → 내부적으로 가상이메일 생성 후 Supabase Auth 위임 |
| | `POST /auth/login` | 사번/비밀번호 → 가상이메일 치환 후 로그인, JWT 반환 |
| | `POST /auth/reset-password-request` | `reset_email`로 재설정 메일 발송 |
| | `GET /auth/me` | 현재 사용자 프로필/권한 조회 |
| 프로젝트 | `POST /projects` / `GET /projects` / `GET /projects/{id}` | 프로젝트 생성/이력 조회 |
| 양식 생성(F1) | `POST /projects/{id}/form-intake/messages` | 문답 1턴 처리, 다음 질문 또는 `spec_complete` 반환 |
| | `POST /projects/{id}/form-templates/generate` | 최종 spec으로 양식 파일 생성 |
| | `POST /form-templates/{id}/revise` | 생성 후 대화형 수정(F1-7) |
| 엑셀 취합(F2) | `POST /projects/{id}/files` | 엑셀 파일 업로드(multipart) |
| | `POST /projects/{id}/files/review` | 업로드 파일 구조 인식 + 검토 실행, 파일별 배지/이슈 반환 |
| | `GET /files/{id}/report` | 오류 리포트 다운로드(xlsx) |
| | `GET /projects/{id}/group-mappings` / `PUT ...` | 모드 C 그룹 매핑 조회/수정 |
| | `POST /projects/{id}/aggregate` | `{mode, included_file_ids, group_mappings?, summary_fields?}` → `aggregation_job` 생성 |
| | `GET /aggregation-jobs/{id}` | 상태/진행률 폴링(F2, F3 공용) |
| | `GET /aggregation-jobs/{id}/download` | 결과 파일 다운로드 |
| 한글 병합(F3) | `POST /projects/{id}/files` (kind=hwpx) | hwpx 업로드 |
| | `POST /projects/{id}/hwpx-merge` | `{file_order}` → `aggregation_job`(kind=hwpx) 생성 |
| 이력(F4-2) | `GET /history?type=excel\|hwpx\|all&q=` | 본인 작업 이력 검색/필터 |
| Admin(F4-3) | `GET /admin/dashboard` | 전체 작업 건수·활성 사용자·API 비용·보관 용량 |
| | `GET /admin/accounts` / `PATCH /admin/accounts/{id}` | 계정 목록/승인·정지·권한 변경 |
| | `PUT /admin/policy/retention` | 보관 기간 등 정책 변경 |
| | `GET /admin/logs` | 전체 사용 로그 조회 |

모든 엔드포인트는 FastAPI가 OpenAPI 스키마를 자동 생성하며, 상세 요청/응답 필드 타입은 구현 단계(Pydantic 모델)에서 정의한다.

---

## 6. 화면별 UI/인터랙션 명세 (화면 1~10)

### 6.0 화면 공통 · 디자인 시스템 및 라이트/다크 모드 (Stitch 시안 확정)

Stitch(AI UI 디자인 도구)로 홈 화면(화면 2 기준) 비주얼 컨셉 5종(미니멀 화이트/그레이·파스텔 프렌들리·다크 모드 프로페셔널·웜톤 크리에이티브·엔터프라이즈 네이비/블루)을 생성해 User 리뷰를 거쳤다. 기능·문구·정보 구조는 5종 모두 동일하게 유지했고, 톤앤매너만 달랐다. 확정 결과는 다음과 같다.

- **라이트 모드 = 컨셉 02 "파스텔 프렌들리"**: 라벤더·민트·피치 파스텔 톤, 카드/버튼 모서리를 크게 둥글림(대형 radius), 그림자는 옅고 부드럽게(soft drop shadow). 친근하고 접근성 높은 인상.
- **다크 모드 = 컨셉 03 "다크 모드 프로페셔널"**: 딥 네이비~차콜 배경에 비비드 블루-바이올렛 포인트 컬러, 무거운 드롭섀도우 대신 은은한 글로우/보더로 깊이감 표현. 하이테크하고 신뢰감 있는 인상.

**참고 팔레트** (Stitch 생성 화면 스크린샷에서 실측 추출한 값. 정확한 디자인 토큰은 구현 단계에서 Stitch 결과물과 대조해 최종 확정)

| 토큰 | 라이트(컨셉 02) | 다크(컨셉 03) |
|---|---|---|
| background | `#F0F0F8` (라벤더 화이트) | `#0C0C24` (딥 네이비-블랙) |
| surface(카드) | `#FFFFFF` | `#181830` ~ `#0C1824` |
| primary(강조/버튼) | `#8880F8` (라벤더) | `#483CE4` (비비드 블루-바이올렛) |
| 보조 파스텔 | mint `#D0F0E8`, peach `#F8E8D0` | — |
| radius | 대형(예: 16px+, pill 컴포넌트는 full) | 표준~중형(기존 8px 계열 유지, 글로우 보더로 강조) |
| 그림자/깊이 | soft drop shadow | 그림자 최소화, 은은한 글로우/보더 강조 |

**라이트/다크 모드 토글**

- 모든 화면(비로그인 상태의 화면 1 포함) 상단바 우측에 라이트/다크 모드 토글(해/달 아이콘 스위치)을 배치한다. 화면 2의 상단 네비(이력/로그아웃/Admin 콘솔 링크) 바로 옆에 위치.
- 토글 클릭 시 새로고침 없이 즉시 전체 화면 테마를 전환한다.
- 기본값: 최초 진입 시 OS `prefers-color-scheme`(라이트/다크) 감지값을 기본으로 적용한다.
- 사용자가 토글을 직접 조작하면 그 선택을 우선하여 브라우저 `localStorage`에 저장하고, 이후 재방문 시 저장된 값을 사용한다. 계정 간 동기화(서버 저장)는 MVP 범위에서 다루지 않는다.

### 화면 1 · 로그인 및 회원가입 (F4-1)
- 중앙 카드 레이아웃, 로그인/회원가입 탭 전환.
- 로그인 탭: 사번, 비밀번호, "로그인" 버튼, 비밀번호 찾기 링크.
- 회원가입 탭 전환 시 **비밀번호 확인**, **비밀번호 재설정용 이메일** 필드가 자동으로 나타남(탭 전환에 필드가 동적으로 추가/제거).
- 계정이 Admin에 의해 정지된 경우 로그인 시 "관리자에 의해 계정이 정지되었습니다" 인라인 오류.
- 우측 상단에 라이트/다크 모드 토글 노출(§6.0 참조, 비로그인 상태에서도 사용 가능).
- API: `POST /auth/login`, `POST /auth/signup`

### 화면 2 · 일반 User 홈 (F4-4)
- 상단 탭바 '생성'/'취합', 기본 선택은 '생성'.
- '생성' 탭: 인사말 + 큰 프롬프트 입력창 + 추천 프롬프트 칩(예: "부서별 예산 집행 현황 양식 만들어줘") — GPT/Gemini 홈 화면 스타일.
- 프롬프트 입력 시 화면 3(문답형 생성)으로 전환.
- '취합' 탭 선택 시 화면 4(엑셀 업로드)로 전환. 탭 상태는 세션 동안 유지.
- 상단 네비: 이력(화면 9) 진입, 로그아웃, Admin 계정인 경우 Admin 콘솔(화면 10) 링크 추가 노출, 라이트/다크 모드 토글(§6.0 참조).

### 화면 3 · AI 양식 생성 (F1-1~F1-2, F1-7)
- 좌측: 챗봇형 문답 UI(순차 질문·응답), 하단 고정 입력창, 문서 첨부 버튼(hwpx/xlsx/png).
- 우측: "지금까지 확인된 사양" 카드 — 대상 부서, 포함 항목, 항목별 입력 형식·필수여부가 실시간 누적 표시.
- "생성" 버튼은 `SPEC_COMPLETE` 전까지 비활성(비활성 상태에 이유 툴팁 표시).
- 생성 완료 후: 결과 미리보기 + 다운로드 + 대화형 수정 입력창(F1-7, 새로운 턴으로 처리).
- API: `POST .../form-intake/messages`, `POST .../form-templates/generate`, `POST .../revise`

### 화면 4 · 엑셀 업로드 및 구조 인식 (F2-1)
- 드래그앤드롭/파일 선택 업로드 영역, 업로드 제한 안내(파일당 50MB, 작업당 최대 30개/총 500MB).
- 업로드 즉시 파일별 카드에 인식된 시트 목록 + 상태 배지(인식 완료/확인 필요) 노출.
- "검토 시작" 버튼 → 화면 5.
- API: `POST .../files`, `POST .../files/review`

### 화면 5 · 검토 결과 및 취합 범위 선택 (F2-3, F2-4)
- 파일 리스트(행 = 체크박스 + 파일명 + 상태 배지 3종: 정상/경고 N/오류 N).
- 기본 선택: 정상·경고 체크, 오류 해제(사용자 변경 가능).
- 배지/행 클릭 시 인라인 확장(오류 유형/건수/위치/사유).
- 체크박스 토글마다 상단 "선택된 파일 n/m건"과 하단 진행 버튼 "취합을 진행하시겠습니까 (n/m)" 문구가 실시간 갱신(클라이언트 상태 기반, 서버 재호출 없음).
- 상단 "오류 리포트 다운로드" 버튼.
- API: 목록 조회 시 `review_results` 포함, `GET /files/{id}/report`

### 화면 6 · 합성 모드 선택 (F2-6~F2-9)
- 모드 A/B/C/D 카드형 선택 UI, 완전한 문장으로 모드 설명, 결과 시트 개수 공식 표기(n×m / m / k / m+1).
- **모드 C 선택 시**: 그룹 매핑 편집 영역 노출 — 시트명 목록에 그룹명을 드롭다운/직접입력으로 지정. 매핑 없는 시트는 "미분류"로 자동 편입되고 경고 배지 표시.
- **모드 D 선택 시**: 요약 대상 지표(컬럼) 다중 선택 UI 노출(해당 실행에만 적용, 고정값 아님).
- 카드 선택 시 하단 "결과 미리보기" 영역에 생성될 시트 탭 목록과 예시 데이터 표시.
- "취합 실행" → `aggregation_job` 생성 후 진행 상태 표시(폴링).
- API: `GET/PUT .../group-mappings`, `POST .../aggregate`, `GET /aggregation-jobs/{id}`

### 화면 7 · 취합 결과 및 다운로드 (F2-11, F2-12)
- 지표 카드 4개(취합 파일/생성 시트/자동 수정/제외 파일) + 아이콘. 그리드는 CSS 미디어쿼리 기준폭으로 좁은 폭 2×2 / 넓은 폭 4×1로만 전환(auto-fit 금지, 3+1 같은 중간 배치 방지).
- 결과 파일·제외 파일: 반응형 2열 그리드 카드.
- 자동 수정 항목: 아이콘 기반 카드 그리드.
- 다운로드 버튼(결과 파일), 오류 리포트 재다운로드 링크.
- API: `GET /aggregation-jobs/{id}`, `GET /aggregation-jobs/{id}/download`

### 화면 8 · 한글(hwpx) 병합 (F3-1, F3-2)
- 업로드 영역 + 업로드된 파일 개수 실시간 표시.
- 목록의 각 행을 실제 드래그하면 병합 순서가 즉시 변경(정적 그립 아이콘이 아닌 동작하는 재정렬).
- "병합 시작" → 비동기 작업 생성, 진행률 표시, 완료 시 알림 + 다운로드.
- API: `POST .../files`(kind=hwpx), `POST .../hwpx-merge`, `GET /aggregation-jobs/{id}`

### 화면 9 · 작업 이력 (F4-2)
- 검색창(파일명/양식명) + 유형 필터(전체/엑셀/hwpx) 클릭 시 목록이 실제로 필터링.
- 목록 항목: 유형 아이콘, 이름, 생성일, 상태, "재사용" 액션(양식은 화면 3으로 spec 프리필, 취합 결과는 재다운로드).
- User는 본인 작업 이력만 조회.
- API: `GET /history?type=&q=`

### 화면 10 · Admin 콘솔 (F4-3)
- 지표 카드: 전체 작업 건수, 활성 사용자, API 비용, 보관 용량.
- 계정 목록 테이블: 사번, 이메일(길 경우 말줄임 처리), 상태, 가입일, 액션(정지/삭제/권한 변경).
- 정책 입력 영역: 보관 기간(일) 입력칸(너비·패딩 보정으로 잘림 방지), 업로드 제한 조정.
- 전체 사용 로그 조회(검색/필터).
- User에게는 노출되지 않는 Admin 전용 화면(화면 2가 일반 User 홈으로 확정되어, 시스템 현황 대시보드는 이 화면 하나로 통합).
- API: `GET /admin/dashboard`, `GET /admin/accounts`, `PATCH /admin/accounts/{id}`, `PUT /admin/policy/retention`, `GET /admin/logs`

---

## 7. 보안 설계 (§6.2 반영)

- 전 구간 HTTPS(TLS 1.2+), Supabase 저장 데이터 암호화(at-rest).
- 비밀번호: Supabase Auth 기본 해시 저장 방식 사용(bcrypt 계열), 로그인 시도 제한은 백엔드에서 rate limit 적용.
- **외부 AI API 전송 정책**: OpenAI GPT API로 문서 데이터 전송 전, 사용자에게 고지 및 동의 절차(양식 생성 최초 진입 시 1회 동의 UI). zero data retention 조건 적용.
- **개인정보 자동 마스킹**: OpenAI로 전송하는 컨텍스트/스펙 텍스트에 정규식 기반 마스킹을 어댑터 계층 진입 전 공통 필터로 적용 — 주민등록번호·전화번호·이메일은 자동 치환, 성명 열은 사용자가 마스킹 여부 선택(F1-3 첨부 파일 및 F1-2 문답 텍스트 모두 대상).
- 사용자별 데이터 격리: 모든 조회 쿼리는 `owner_id`/`uploader_id` 기준으로 스코핑(Supabase RLS 정책으로 이중 보장 권장).
- Admin 접근 및 전체 사용 로그는 `audit_logs`에 기록.

---

## 8. 잔여 리스크 확정 결과 및 남은 리스크

PRD §10 잔여 리스크 노트 중 design.md 단계 확정 대상 2건:

| 항목 | 확정 결과 | 반영 위치 |
|---|---|---|
| 모드 C 그룹 매핑 관리 방식 | 화면 UI로 관리(별도 설정 파일 없음), 프로젝트 단위로 `group_mappings` 테이블에 저장·재사용 | §3, §4.2, 화면 6 |
| 모드 D 요약 지표 정의 | 고정하지 않고 취합 모드 선택·실행 시마다 사용자가 지정, `aggregation_jobs.summary_fields_json`에 실행별로 저장 | §3, §4.2, 화면 6 |

design.md 작성 과정에서 새로 식별된 리스크:

- **비동기 처리 내구성**: Redis/Celery 없이 DB 폴링 + BackgroundTasks로 구현하기로 결정(해커톤 일정 우선). 서버 재시작 시 진행 중 작업은 복구되지 않고 `failed` 처리됨 — 사용량이 늘거나 MVP 이후 안정화 단계에서 큐 인프라 도입 재검토 필요.
- **가상 이메일 인증 방식의 Supabase 의존**: 사번을 가상 이메일로 매핑하는 방식은 Supabase Auth의 email/password 플로우를 그대로 재사용하기 위한 절충안. Supabase Auth 정책(이메일 형식 검증, 중복 체크 등)이 변경되면 매핑 로직 영향 가능.
- **hwpx 전환 게이트**(PRD 기존 리스크 유지): M1 PoC 결과에 따라 상용 SDK 전환 여부 판단 필요, 본 design.md는 직접 파싱 구현을 전제로 작성됨.
