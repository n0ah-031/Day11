-- 정책 변경 기록이 행위자 계정 삭제를 막지 않게 한다.
--
-- 직전 마이그레이션(audit_logs)과 같은 부류다. retention_policy.updated_by도
-- profiles(id)를 ON DELETE 절 없이 참조해, 정책을 한 번 바꾼 관리자는 삭제할
-- 수 없었다. 정책 행 자체는 남아야 하므로 SET NULL로 둔다 — 마지막 변경자만
-- 잃고 값은 유지된다.
--
-- profiles를 참조하는 FK를 전수 점검한 결과 남은 것은 이것뿐이다:
--   projects.owner_id        CASCADE   (사용자 자료는 함께 사라지는 것이 맞다)
--   uploaded_files.uploader_id CASCADE (같음)
--   audit_logs.actor_id      SET NULL  (직전 마이그레이션)
--   retention_policy.updated_by NO ACTION → SET NULL  (이 마이그레이션)

alter table public.retention_policy
  drop constraint if exists retention_policy_updated_by_fkey;

alter table public.retention_policy
  add constraint retention_policy_updated_by_fkey
  foreign key (updated_by) references public.profiles(id) on delete set null;
