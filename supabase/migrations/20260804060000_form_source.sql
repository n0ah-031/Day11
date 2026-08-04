-- F6 기존 양식 인지·등록 — 양식이 어떻게 생겼는지를 기록한다.
--
-- 인지 기능 기술명세 §3.2는 form_templates에 source·structural_json을, intake_sessions에
-- mode·primary_attachment_id·structural_json을 더하라고 한다. 그중 **source 하나만** 만든다.
--
-- 나머지 넷을 만들지 않는 이유:
--   · intake_sessions.attachments_json이 이미 있고(F1 첨부용으로 만들어 두고 안 쓰던 컬럼)
--     구조 추출 결과·등록 대상 지정은 그 첨부 항목에 붙는 정보다. 첨부 배열 안에
--     structural·primary를 두면 파일과 그 파일의 구조가 한자리에 있고, 세션 mode는
--     `등록 대상으로 지정된 xlsx가 있는가`로 그대로 파생된다(별도 상태를 두면 둘이 어긋난다).
--   · form_templates.structural_json은 등록 원본이 무변형으로 Storage에 남으므로
--     언제든 aggregate.read_file로 다시 뽑을 수 있다. 스냅샷을 따로 두면 원본과
--     스냅샷이라는 두 개의 진실이 생긴다.
--
-- source는 파생시킬 수 없어서 컬럼으로 둔다. `workbook_json is null`로 구분하려 하면
-- 저작에 실패한 F1 양식(status=failed)과 등록 양식이 같은 모양이 되고, 등록 양식은
-- 원본이 산출물이라 workbook_json이 영구히 비어 있는 것이 정상이다(명세 E1·§3.2).
-- 이 구분은 F1-7 대화형 수정이 손대면 안 되는 양식을 가리는 근거이기도 하다(E3).

alter table public.form_templates
  add column if not exists source text not null default 'ai_generated';

alter table public.form_templates
  drop constraint if exists form_templates_source_check;
alter table public.form_templates
  add constraint form_templates_source_check
  check (source in ('ai_generated', 'recognized_external'));

comment on column public.form_templates.source is
  'ai_generated(F1 저작) / recognized_external(F6 등록). 등록 양식은 원본 파일이 산출물이므로 workbook_json이 NULL이고 재저작 대상이 아니다(인지기능 명세 E1·E3).';
