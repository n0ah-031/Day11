-- AI 양식 생성(F1)에 필요한 테이블·컬럼을 만든다.
--
-- form_templates·field_rules는 이미 있었다(design.md §3의 최소 형태). 다만 생성기능
-- 기술명세 §3.3이 요구하는 상태·버전·저작물 컬럼이 없어 문답에서 생성까지의 진행을
-- 기록할 수 없었다. 없는 것만 더한다.
--
-- intake_sessions는 신설이다. 문답은 여러 턴에 걸쳐 사양이 누적되는 과정이라,
-- 프로세스 메모리에만 두면 서버가 재시작될 때 사용자가 답한 내용이 전부 사라진다.
-- F2·F3의 세션은 업로드한 파일이 Storage에 남아 있어 다시 시작할 수 있지만, 문답은
-- 되살릴 근거가 없다.
--
-- profiles.ai_consent_at은 AI 전송 동의(명세 §9.1)를 1회 기록한다. 사용자가 입력한
-- 내용이 외부 API로 나가므로 동의 없이는 F1 API 전체를 403으로 막는다.

-- ── intake_sessions (신설) ───────────────────────────────────────────────────
create table if not exists public.intake_sessions (
  id                uuid primary key default gen_random_uuid(),
  project_id        uuid not null references public.projects(id) on delete cascade,
  status            text not null default 'active',
  messages_json     jsonb not null default '[]'::jsonb,
  spec_json         jsonb not null default '{}'::jsonb,
  attachments_json  jsonb not null default '[]'::jsonb,
  turn_count        int  not null default 0,
  created_at        timestamptz not null default now(),
  updated_at        timestamptz not null default now(),
  constraint intake_sessions_status_check
    check (status in ('active', 'spec_complete', 'closed'))
);

comment on table public.intake_sessions is
  'F1 문답 세션. messages_json·spec_json은 사용자 입력 원문을 담는다 — 마스킹은 외부 API 전송 직전에만 적용한다(명세 §9.2).';

-- 프로젝트당 진행 중인 문답은 하나뿐이다. 닫힌 세션은 여러 개 남을 수 있다.
create unique index if not exists intake_sessions_one_active
  on public.intake_sessions (project_id)
  where status = 'active';

create index if not exists intake_sessions_project_idx
  on public.intake_sessions (project_id, created_at desc);

alter table public.intake_sessions enable row level security;
-- 정책은 두지 않는다. 다른 테이블과 같은 원칙 — anon·authenticated 전면 거부이고
-- 모든 접근은 service key를 쓰는 백엔드를 경유한다(20260803065832 참조).

-- ── form_templates 확장 (명세 §3.3) ─────────────────────────────────────────
alter table public.form_templates
  -- 세션이 정리돼도 생성된 양식은 남아야 하므로 CASCADE가 아니라 SET NULL이다.
  add column if not exists intake_session_id uuid
    references public.intake_sessions(id) on delete set null,
  add column if not exists status text not null default 'generating',
  add column if not exists version int not null default 1,
  add column if not exists workbook_json jsonb,
  add column if not exists output_format text not null default 'xlsx',
  add column if not exists error_message text,
  add column if not exists updated_at timestamptz not null default now();

comment on column public.form_templates.workbook_json is
  'LLM이 저작한 셀 단위 파일 명세(명세 §6.3). 대화형 수정(F1-7)의 입력이자 미리보기 원본.';
comment on column public.form_templates.status is
  'generating / done / failed. 실패 사유는 error_message에 남긴다.';

-- 기존 행(있다면)은 파일이 이미 있으므로 done으로 본다.
update public.form_templates set status = 'done' where file_url is not null and status = 'generating';

alter table public.form_templates
  drop constraint if exists form_templates_status_check;
alter table public.form_templates
  add constraint form_templates_status_check
  check (status in ('generating', 'done', 'failed'));

create index if not exists form_templates_project_idx
  on public.form_templates (project_id, created_at desc);

-- ── profiles.ai_consent_at (명세 §9.1) ──────────────────────────────────────
alter table public.profiles add column if not exists ai_consent_at timestamptz;

comment on column public.profiles.ai_consent_at is
  'AI 전송 동의 시각. NULL이면 F1 API를 403으로 막는다. 취합(F2)의 2단계 AI 재검증은 이 동의와 별개다(기존 동작 유지).';
