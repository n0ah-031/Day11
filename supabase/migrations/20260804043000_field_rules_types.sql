-- field_rules.rule_type 제약을 기술명세에 맞춘다.
--
-- 베이스라인(remote-only 0001·0002)이 만든 제약은 date/name/amount/dept_code 4종만
-- 허용한다. design.md §3의 초기 스케치를 그대로 옮긴 값이다. 그런데 생성기능
-- 기술명세 §3.4는 여기에 text·number를 더한 6종을 규정하고, formgen.FIELD_TYPES도
-- 6종이다.
--
-- 실무 양식에서 가장 흔한 유형이 바로 그 text·number라서, F1 생성은 예외 없이
-- 작성기준 저장(store.save_field_rules)에서 23514로 죽었다. 브라우저에서 문답 6턴을
-- 마치고 '양식 생성'을 눌렀을 때 처음 드러났다 — test_formgen.py는 store를 타지 않아
-- 잡히지 않는 지점이다.
--
-- 값을 좁히는 쪽(text·number를 4종 중 하나로 뭉개기)은 택하지 않았다. 부서명·사번·
-- 교육과정명·비고를 전부 name으로 적으면 근거 없는 유형을 기록하게 된다.

alter table public.field_rules
  drop constraint if exists field_rules_rule_type_check;
alter table public.field_rules
  add constraint field_rules_rule_type_check
  check (rule_type in ('date', 'name', 'amount', 'dept_code', 'text', 'number'));

comment on column public.field_rules.rule_type is
  '생성기능 기술명세 §3.4의 6종. formgen.FIELD_TYPES와 같은 어휘를 쓴다.';
