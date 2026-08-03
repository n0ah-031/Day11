-- public 스키마 전 테이블에 RLS를 켠다.
--
-- 배경: 이 스키마는 저장소 밖(대시보드 SQL)에서 만들어졌고 RLS가 꺼진 채였다.
-- Supabase는 public 스키마를 Data API(PostgREST)로 노출하고 anon·authenticated
-- 역할에 기본 권한을 부여하므로, RLS가 없으면 프론트엔드에 실려 공개되는 anon 키만
-- 있으면 누구나 전 테이블을 읽고 쓸 수 있다. `supabase db advisors`가 10개 테이블
-- 전부를 rls_disabled_in_public(ERROR, EXTERNAL)로 지적했다.
--
-- 정책은 아직 만들지 않는다. 정책 없이 RLS만 켜면 anon·authenticated에는 전면 거부가
-- 되고, 서버가 쓰는 service_role은 RLS를 우회하므로 현재 동작에는 영향이 없다.
-- 인증(F4-1)이 구현되어 실제 접근 모델이 정해지면 그때 소유권 기반 정책을 추가한다.
-- 지금 정책을 미리 쓰면 존재하지 않는 접근 모델을 추측하는 셈이 된다.

alter table public.profiles          enable row level security;
alter table public.projects          enable row level security;
alter table public.form_templates    enable row level security;
alter table public.field_rules       enable row level security;
alter table public.group_mappings    enable row level security;
alter table public.uploaded_files    enable row level security;
alter table public.review_results    enable row level security;
alter table public.aggregation_jobs  enable row level security;
alter table public.audit_logs        enable row level security;
alter table public.retention_policy  enable row level security;
