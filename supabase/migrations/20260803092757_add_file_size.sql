-- 업로드 파일 크기를 기록한다.
--
-- Admin 콘솔(F4-3)의 '보관 용량' 지표에 필요하다. 실제 바이트는 Storage에
-- 있지만 PostgREST는 storage 스키마에 닿지 않고, 노출시키려면 뷰나
-- SECURITY DEFINER 함수를 public에 두어야 해서 대가가 크다. 업로드 시점에
-- 이미 알고 있는 값이라 컬럼 하나로 끝낸다.
--
-- 기존 행은 크기를 모르므로 NULL로 남는다(0으로 채우면 없는 값을 0으로
-- 단정하는 셈이다). 지표는 NULL을 제외해 합산한다.

alter table public.uploaded_files add column if not exists size_bytes bigint;

comment on column public.uploaded_files.size_bytes is
  '업로드 당시 파일 크기(바이트). 이 컬럼 추가 전에 올라온 행은 NULL.';
