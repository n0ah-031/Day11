-- 감사 로그가 행위자 계정 삭제를 막지 않게 한다.
--
-- audit_logs.actor_id가 profiles(id)를 ON DELETE 절 없이(= NO ACTION) 참조하고
-- 있었다. 로그를 쓰기 시작한 뒤로는 뭔가를 한 사용자를 지울 수 없었다 —
-- auth.users 삭제 → profiles CASCADE 삭제까지 가서 이 제약에 걸려 전체가
-- 실패한다. 정작 관리자가 지우려는 대상이 그런 사용자다.
--
-- SET NULL로 바꾼다. 감사 로그는 행위자가 사라진 뒤에도 남아야 의미가 있고
-- (그게 감사 기록의 목적이다), 조회 쪽은 actor가 NULL이면 '(삭제된 계정)'으로
-- 표시한다. CASCADE로 로그까지 지우면 삭제한 사실 자체가 증발한다.

alter table public.audit_logs
  drop constraint if exists audit_logs_actor_id_fkey;

alter table public.audit_logs
  add constraint audit_logs_actor_id_fkey
  foreign key (actor_id) references public.profiles(id) on delete set null;
