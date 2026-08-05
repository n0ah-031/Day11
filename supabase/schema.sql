-- 원격 DB의 현재 스키마 전체 스냅샷 (supabase db dump --linked)
--
-- **왜 이 파일이 있는가.** 원격 마이그레이션 이력의 `0001`·`0002`는 SQL이 저장소에 없다
-- (대시보드 SQL 에디터로 만든 기본 스키마다). 그래서 `supabase/migrations/`만 순서대로
-- 돌리면 **테이블이 하나도 서지 않는다** — 기반 테이블이 없어 전부 실패한다.
-- 실제로 확인했다: 빈 Postgres에 마이그레이션 7개를 순서대로 돌려 public 테이블 0개.
--
-- **새 환경 만들기**: 이 파일 하나를 적용하면 현재 상태(테이블 11개·제약·RLS)가 그대로 선다.
-- 확인 방법도 같다 — 빈 Postgres에 이 파일을 돌리면 테이블 11개·RLS 11개가 나온다.
-- 이후에 추가되는 변경만 `supabase/migrations/`에 쌓는다.
--
-- **Supabase 밖에서 돌릴 때 나오는 오류 2건은 정상이다** — `supabase_vault` 확장과
-- `supabase_realtime` publication은 Supabase 플랫폼이 제공하는 것이라 맨 Postgres에는 없다.
-- 앱 스키마와는 무관하다.
--
-- **스키마를 바꾸면 이 파일도 다시 뽑는다**: `supabase db dump --linked -f supabase/schema.sql`
-- (Docker 데몬이 필요하다. 이 머신은 colima를 쓴다 — `colima start`)
--
-- 뽑은 시각: 2026-08-05 (마이그레이션 20260804060000까지 반영됨)

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;


COMMENT ON SCHEMA "public" IS 'standard public schema';



CREATE EXTENSION IF NOT EXISTS "pg_stat_statements" WITH SCHEMA "extensions";






CREATE EXTENSION IF NOT EXISTS "pgcrypto" WITH SCHEMA "extensions";






CREATE EXTENSION IF NOT EXISTS "supabase_vault" WITH SCHEMA "vault";






CREATE EXTENSION IF NOT EXISTS "uuid-ossp" WITH SCHEMA "extensions";





SET default_tablespace = '';

SET default_table_access_method = "heap";


CREATE TABLE IF NOT EXISTS "public"."aggregation_jobs" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "project_id" "uuid" NOT NULL,
    "kind" "text" NOT NULL,
    "mode" "text",
    "status" "text" DEFAULT 'queued'::"text" NOT NULL,
    "included_file_ids" "jsonb" DEFAULT '[]'::"jsonb" NOT NULL,
    "summary_fields_json" "jsonb" DEFAULT '[]'::"jsonb" NOT NULL,
    "result_url" "text",
    "progress" integer DEFAULT 0 NOT NULL,
    "error_message" "text",
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "stats_json" "jsonb" DEFAULT '{}'::"jsonb" NOT NULL,
    CONSTRAINT "aggregation_jobs_kind_check" CHECK (("kind" = ANY (ARRAY['excel'::"text", 'hwpx'::"text"]))),
    CONSTRAINT "aggregation_jobs_mode_check" CHECK (("mode" = ANY (ARRAY['A'::"text", 'B'::"text", 'C'::"text", 'D'::"text"]))),
    CONSTRAINT "aggregation_jobs_status_check" CHECK (("status" = ANY (ARRAY['queued'::"text", 'running'::"text", 'done'::"text", 'failed'::"text"])))
);


ALTER TABLE "public"."aggregation_jobs" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."audit_logs" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "actor_id" "uuid",
    "action" "text" NOT NULL,
    "target" "text",
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL
);


ALTER TABLE "public"."audit_logs" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."field_rules" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "form_template_id" "uuid",
    "field_name" "text" NOT NULL,
    "rule_type" "text" NOT NULL,
    "rule_config_json" "jsonb" DEFAULT '{}'::"jsonb" NOT NULL,
    CONSTRAINT "field_rules_rule_type_check" CHECK (("rule_type" = ANY (ARRAY['date'::"text", 'name'::"text", 'amount'::"text", 'dept_code'::"text", 'text'::"text", 'number'::"text"])))
);


ALTER TABLE "public"."field_rules" OWNER TO "postgres";


COMMENT ON COLUMN "public"."field_rules"."rule_type" IS '생성기능 기술명세 §3.4의 6종. formgen.FIELD_TYPES와 같은 어휘를 쓴다.';



CREATE TABLE IF NOT EXISTS "public"."form_templates" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "project_id" "uuid" NOT NULL,
    "spec_json" "jsonb" DEFAULT '{}'::"jsonb" NOT NULL,
    "file_url" "text",
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "intake_session_id" "uuid",
    "status" "text" DEFAULT 'generating'::"text" NOT NULL,
    "version" integer DEFAULT 1 NOT NULL,
    "workbook_json" "jsonb",
    "output_format" "text" DEFAULT 'xlsx'::"text" NOT NULL,
    "error_message" "text",
    "updated_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "source" "text" DEFAULT 'ai_generated'::"text" NOT NULL,
    CONSTRAINT "form_templates_source_check" CHECK (("source" = ANY (ARRAY['ai_generated'::"text", 'recognized_external'::"text"]))),
    CONSTRAINT "form_templates_status_check" CHECK (("status" = ANY (ARRAY['generating'::"text", 'done'::"text", 'failed'::"text"])))
);


ALTER TABLE "public"."form_templates" OWNER TO "postgres";


COMMENT ON COLUMN "public"."form_templates"."status" IS 'generating / done / failed. 실패 사유는 error_message에 남긴다.';



COMMENT ON COLUMN "public"."form_templates"."workbook_json" IS 'LLM이 저작한 셀 단위 파일 명세(명세 §6.3). 대화형 수정(F1-7)의 입력이자 미리보기 원본.';



CREATE TABLE IF NOT EXISTS "public"."group_mappings" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "project_id" "uuid" NOT NULL,
    "sheet_name" "text" NOT NULL,
    "group_name" "text" NOT NULL
);


ALTER TABLE "public"."group_mappings" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."intake_sessions" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "project_id" "uuid" NOT NULL,
    "status" "text" DEFAULT 'active'::"text" NOT NULL,
    "messages_json" "jsonb" DEFAULT '[]'::"jsonb" NOT NULL,
    "spec_json" "jsonb" DEFAULT '{}'::"jsonb" NOT NULL,
    "attachments_json" "jsonb" DEFAULT '[]'::"jsonb" NOT NULL,
    "turn_count" integer DEFAULT 0 NOT NULL,
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "updated_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    CONSTRAINT "intake_sessions_status_check" CHECK (("status" = ANY (ARRAY['active'::"text", 'spec_complete'::"text", 'closed'::"text"])))
);


ALTER TABLE "public"."intake_sessions" OWNER TO "postgres";


COMMENT ON TABLE "public"."intake_sessions" IS 'F1 문답 세션. messages_json·spec_json은 사용자 입력 원문을 담는다 — 마스킹은 외부 API 전송 직전에만 적용한다(명세 §9.2).';



CREATE TABLE IF NOT EXISTS "public"."profiles" (
    "id" "uuid" NOT NULL,
    "employee_no" "text" NOT NULL,
    "reset_email" "text" NOT NULL,
    "role" "text" DEFAULT 'user'::"text" NOT NULL,
    "status" "text" DEFAULT 'active'::"text" NOT NULL,
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "ai_consent_at" timestamp with time zone,
    CONSTRAINT "profiles_role_check" CHECK (("role" = ANY (ARRAY['user'::"text", 'admin'::"text"]))),
    CONSTRAINT "profiles_status_check" CHECK (("status" = ANY (ARRAY['active'::"text", 'suspended'::"text"])))
);


ALTER TABLE "public"."profiles" OWNER TO "postgres";


COMMENT ON COLUMN "public"."profiles"."ai_consent_at" IS 'AI 전송 동의 시각. NULL이면 F1 API를 403으로 막는다. 취합(F2)의 2단계 AI 재검증은 이 동의와 별개다(기존 동작 유지).';



CREATE TABLE IF NOT EXISTS "public"."projects" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "owner_id" "uuid" NOT NULL,
    "name" "text" NOT NULL,
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL
);


ALTER TABLE "public"."projects" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."retention_policy" (
    "key" "text" NOT NULL,
    "value" "text" NOT NULL,
    "updated_by" "uuid"
);


ALTER TABLE "public"."retention_policy" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."review_results" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "file_id" "uuid" NOT NULL,
    "badge" "text" NOT NULL,
    "issue_count" integer DEFAULT 0 NOT NULL,
    "issues_json" "jsonb" DEFAULT '[]'::"jsonb" NOT NULL,
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    CONSTRAINT "review_results_badge_check" CHECK (("badge" = ANY (ARRAY['정상'::"text", '경고'::"text", '오류'::"text"])))
);


ALTER TABLE "public"."review_results" OWNER TO "postgres";


CREATE TABLE IF NOT EXISTS "public"."uploaded_files" (
    "id" "uuid" DEFAULT "gen_random_uuid"() NOT NULL,
    "project_id" "uuid" NOT NULL,
    "uploader_id" "uuid" NOT NULL,
    "storage_path" "text" NOT NULL,
    "original_name" "text" NOT NULL,
    "kind" "text" NOT NULL,
    "status" "text" DEFAULT 'uploaded'::"text" NOT NULL,
    "created_at" timestamp with time zone DEFAULT "now"() NOT NULL,
    "deleted_at" timestamp with time zone,
    "size_bytes" bigint,
    CONSTRAINT "uploaded_files_kind_check" CHECK (("kind" = ANY (ARRAY['excel'::"text", 'hwpx'::"text"])))
);


ALTER TABLE "public"."uploaded_files" OWNER TO "postgres";


COMMENT ON COLUMN "public"."uploaded_files"."size_bytes" IS '업로드 당시 파일 크기(바이트). 이 컬럼 추가 전에 올라온 행은 NULL.';



ALTER TABLE ONLY "public"."aggregation_jobs"
    ADD CONSTRAINT "aggregation_jobs_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."audit_logs"
    ADD CONSTRAINT "audit_logs_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."field_rules"
    ADD CONSTRAINT "field_rules_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."form_templates"
    ADD CONSTRAINT "form_templates_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."group_mappings"
    ADD CONSTRAINT "group_mappings_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."intake_sessions"
    ADD CONSTRAINT "intake_sessions_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."profiles"
    ADD CONSTRAINT "profiles_employee_no_key" UNIQUE ("employee_no");



ALTER TABLE ONLY "public"."profiles"
    ADD CONSTRAINT "profiles_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."projects"
    ADD CONSTRAINT "projects_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."retention_policy"
    ADD CONSTRAINT "retention_policy_pkey" PRIMARY KEY ("key");



ALTER TABLE ONLY "public"."review_results"
    ADD CONSTRAINT "review_results_pkey" PRIMARY KEY ("id");



ALTER TABLE ONLY "public"."uploaded_files"
    ADD CONSTRAINT "uploaded_files_pkey" PRIMARY KEY ("id");



CREATE INDEX "form_templates_project_idx" ON "public"."form_templates" USING "btree" ("project_id", "created_at" DESC);



CREATE UNIQUE INDEX "intake_sessions_one_active" ON "public"."intake_sessions" USING "btree" ("project_id") WHERE ("status" = 'active'::"text");



CREATE INDEX "intake_sessions_project_idx" ON "public"."intake_sessions" USING "btree" ("project_id", "created_at" DESC);



ALTER TABLE ONLY "public"."aggregation_jobs"
    ADD CONSTRAINT "aggregation_jobs_project_id_fkey" FOREIGN KEY ("project_id") REFERENCES "public"."projects"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."audit_logs"
    ADD CONSTRAINT "audit_logs_actor_id_fkey" FOREIGN KEY ("actor_id") REFERENCES "public"."profiles"("id") ON DELETE SET NULL;



ALTER TABLE ONLY "public"."field_rules"
    ADD CONSTRAINT "field_rules_form_template_id_fkey" FOREIGN KEY ("form_template_id") REFERENCES "public"."form_templates"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."form_templates"
    ADD CONSTRAINT "form_templates_intake_session_id_fkey" FOREIGN KEY ("intake_session_id") REFERENCES "public"."intake_sessions"("id") ON DELETE SET NULL;



ALTER TABLE ONLY "public"."form_templates"
    ADD CONSTRAINT "form_templates_project_id_fkey" FOREIGN KEY ("project_id") REFERENCES "public"."projects"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."group_mappings"
    ADD CONSTRAINT "group_mappings_project_id_fkey" FOREIGN KEY ("project_id") REFERENCES "public"."projects"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."intake_sessions"
    ADD CONSTRAINT "intake_sessions_project_id_fkey" FOREIGN KEY ("project_id") REFERENCES "public"."projects"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."profiles"
    ADD CONSTRAINT "profiles_id_fkey" FOREIGN KEY ("id") REFERENCES "auth"."users"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."projects"
    ADD CONSTRAINT "projects_owner_id_fkey" FOREIGN KEY ("owner_id") REFERENCES "public"."profiles"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."retention_policy"
    ADD CONSTRAINT "retention_policy_updated_by_fkey" FOREIGN KEY ("updated_by") REFERENCES "public"."profiles"("id") ON DELETE SET NULL;



ALTER TABLE ONLY "public"."review_results"
    ADD CONSTRAINT "review_results_file_id_fkey" FOREIGN KEY ("file_id") REFERENCES "public"."uploaded_files"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."uploaded_files"
    ADD CONSTRAINT "uploaded_files_project_id_fkey" FOREIGN KEY ("project_id") REFERENCES "public"."projects"("id") ON DELETE CASCADE;



ALTER TABLE ONLY "public"."uploaded_files"
    ADD CONSTRAINT "uploaded_files_uploader_id_fkey" FOREIGN KEY ("uploader_id") REFERENCES "public"."profiles"("id") ON DELETE CASCADE;



ALTER TABLE "public"."aggregation_jobs" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."audit_logs" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."field_rules" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."form_templates" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."group_mappings" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."intake_sessions" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."profiles" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."projects" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."retention_policy" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."review_results" ENABLE ROW LEVEL SECURITY;


ALTER TABLE "public"."uploaded_files" ENABLE ROW LEVEL SECURITY;




ALTER PUBLICATION "supabase_realtime" OWNER TO "postgres";


GRANT USAGE ON SCHEMA "public" TO "postgres";
GRANT USAGE ON SCHEMA "public" TO "anon";
GRANT USAGE ON SCHEMA "public" TO "authenticated";
GRANT USAGE ON SCHEMA "public" TO "service_role";





































































































































































GRANT ALL ON TABLE "public"."aggregation_jobs" TO "anon";
GRANT ALL ON TABLE "public"."aggregation_jobs" TO "authenticated";
GRANT ALL ON TABLE "public"."aggregation_jobs" TO "service_role";



GRANT ALL ON TABLE "public"."audit_logs" TO "anon";
GRANT ALL ON TABLE "public"."audit_logs" TO "authenticated";
GRANT ALL ON TABLE "public"."audit_logs" TO "service_role";



GRANT ALL ON TABLE "public"."field_rules" TO "anon";
GRANT ALL ON TABLE "public"."field_rules" TO "authenticated";
GRANT ALL ON TABLE "public"."field_rules" TO "service_role";



GRANT ALL ON TABLE "public"."form_templates" TO "anon";
GRANT ALL ON TABLE "public"."form_templates" TO "authenticated";
GRANT ALL ON TABLE "public"."form_templates" TO "service_role";



GRANT ALL ON TABLE "public"."group_mappings" TO "anon";
GRANT ALL ON TABLE "public"."group_mappings" TO "authenticated";
GRANT ALL ON TABLE "public"."group_mappings" TO "service_role";



GRANT ALL ON TABLE "public"."intake_sessions" TO "anon";
GRANT ALL ON TABLE "public"."intake_sessions" TO "authenticated";
GRANT ALL ON TABLE "public"."intake_sessions" TO "service_role";



GRANT ALL ON TABLE "public"."profiles" TO "anon";
GRANT ALL ON TABLE "public"."profiles" TO "authenticated";
GRANT ALL ON TABLE "public"."profiles" TO "service_role";



GRANT ALL ON TABLE "public"."projects" TO "anon";
GRANT ALL ON TABLE "public"."projects" TO "authenticated";
GRANT ALL ON TABLE "public"."projects" TO "service_role";



GRANT ALL ON TABLE "public"."retention_policy" TO "anon";
GRANT ALL ON TABLE "public"."retention_policy" TO "authenticated";
GRANT ALL ON TABLE "public"."retention_policy" TO "service_role";



GRANT ALL ON TABLE "public"."review_results" TO "anon";
GRANT ALL ON TABLE "public"."review_results" TO "authenticated";
GRANT ALL ON TABLE "public"."review_results" TO "service_role";



GRANT ALL ON TABLE "public"."uploaded_files" TO "anon";
GRANT ALL ON TABLE "public"."uploaded_files" TO "authenticated";
GRANT ALL ON TABLE "public"."uploaded_files" TO "service_role";









ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON SEQUENCES TO "postgres";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON SEQUENCES TO "anon";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON SEQUENCES TO "authenticated";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON SEQUENCES TO "service_role";






ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON FUNCTIONS TO "postgres";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON FUNCTIONS TO "anon";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON FUNCTIONS TO "authenticated";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON FUNCTIONS TO "service_role";






ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON TABLES TO "postgres";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON TABLES TO "anon";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON TABLES TO "authenticated";
ALTER DEFAULT PRIVILEGES FOR ROLE "postgres" IN SCHEMA "public" GRANT ALL ON TABLES TO "service_role";































