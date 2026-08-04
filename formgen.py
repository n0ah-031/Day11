"""F1 AI 양식 생성 — 문답으로 사양을 모아 xlsx 양식과 작성기준을 만든다.

두 번의 LLM 역할이 있다(생성기능 기술명세 D1·D3).

1. **문답**(`intake_turn`) — 질문 순서·완료 판정을 LLM이 주도한다. 대신 백엔드가 두 가지
   가드를 건다: 완결성 루브릭(§4.3)과 최대 턴 수. LLM이 "다 됐다"고 해도 사양에 빈 곳이
   있으면 생성으로 넘기지 않는다.
2. **저작**(`author_workbook`) — LLM이 셀 단위로 파일 내용을 짓고(`workbook_json`),
   `materialize`는 그걸 해석 없이 xlsx로 옮기기만 한다. 레이아웃 규칙을 코드에 넣지 않는
   이유는, 양식마다 제목행·안내문·시트 분리 요구가 달라서 규칙을 넣는 만큼 갇히기 때문이다.

**LLM 산출물은 신뢰하지 않는다.** `validate_workbook`이 통과시킨 것만 파일로 만든다
(수식 함수 화이트리스트·외부 참조 금지·셀 수 상한·스타일 속성 화이트리스트).

마스킹은 `aggregate.mask`를 그대로 쓴다 — 외부 전송 직전 공통 통과 지점이 이미 거기 있고
(명세 §13.2), 두 번째 마스커를 만들면 규칙이 갈라진다. 저장은 원문이고 마스킹은 전송할
때만 한다(§9.2).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import aggregate as ag

TURN_LIMIT = 30           # §4.3 최대 턴 가드
RECOGNIZE_TURN_LIMIT = 10  # 인지 §4.3 — 구조가 이미 주어져 확인 질문이 적은 것이 정상
AUTHOR_RETRY = 2          # §6.3 검증 실패 시 재저작 횟수
MAX_SHEETS = 10
MAX_CELLS = 10_000
ATTACH_CHARS = 8_000      # §6.5 첨부 추출 텍스트 절단
SAMPLE_ROWS = 3           # 인지 §6.2 sample_rows — 형식 추론에 필요한 만큼만

FIELD_TYPES = ("text", "number", "amount", "date", "name", "dept_code")
STYLE_KEYS = ("bold", "italic", "size", "bg", "color", "border", "align", "number_format")
ALLOWED_FUNCS = ("SUM", "SUMIF", "SUMIFS", "COUNT", "IF", "TODAY")
VALIDATION_TYPES = ("date", "list", "whole", "decimal")

REF_RE = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}$")
RANGE_RE = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}(:\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})?$")
FUNC_RE = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(")
BAD_SHEET_CHARS = set(r"\/?*[]:")

INTAKE_SYSTEM = """당신은 공공기관 업무 양식을 설계하는 인터뷰어입니다.
사용자가 배포할 엑셀 양식의 사양을 문답으로 확정하는 것이 목표입니다.

반드시 확인할 것:
1) 양식 이름
2) 배포 대상(전체 부서인지, 특정 부서 목록인지)
3) 포함할 항목(열)
4) 항목별 입력 형식 — 날짜 표기(예: YYYY-MM-DD), 금액 단위·구분기호 유무, 코드표 사용 여부
5) 항목별 필수/선택 구분

규칙:
- 한 번에 한 주제만 묻습니다. 여러 질문을 한꺼번에 쏟지 마세요.
- 사용자가 "알아서 해달라"고 하면 실무 관례에 맞는 기본값을 제안하고 동의를 구합니다.
- 이미 답이 나온 것을 다시 묻지 마세요.
- **사용자가 정정하면 그 지시가 이깁니다.** 예를 들어 "비고만 선택이고 나머지는 필수"라고
  하면 비고를 제외한 모든 항목의 required를 true로 바꿉니다. 앞서 당신이 제안했던 값을
  그대로 남겨두지 마세요. 필수/선택이 틀리면 회신본 검토에서 그대로 오판으로 이어집니다.
- 확인된 내용은 매 턴 spec_json 전체 갱신본으로 누적합니다.
- 남은 확인 사항이 없을 때만 spec_complete를 true로 둡니다.

반드시 아래 JSON만 출력합니다.
{
  "reply": "사용자에게 보여줄 한국어 메시지(질문 또는 확인 요청)",
  "spec_json": {
    "form_title": "양식 이름",
    "target_depts": ["전체 부서"],
    "fields": [
      {"name": "항목명", "type": "text|number|amount|date|name|dept_code",
       "required": true, "format": "형식 설명 또는 null", "notes": "비고 또는 null"}
    ],
    "layout_hints": "레이아웃 요구(자유 텍스트) 또는 null",
    "locale": "ko"
  },
  "spec_complete": false,
  "coverage": {"confirmed_topics": [], "remaining_topics": []}
}
type이 date이거나 amount인 항목은 format을 반드시 채웁니다."""

AUTHOR_SYSTEM = f"""당신은 확정된 사양(spec_json)을 받아 배포용 엑셀 양식 파일의 내용을
셀 단위로 저작합니다. 회신자가 값을 채워 되돌려줄 빈 양식입니다.

실무 관례를 따르세요:
- 최상단에 양식 제목 행(가로 병합, 굵게, 가운데)
- 필요하면 제목 아래 작성 안내 문구 한두 줄
- 헤더 행은 굵게 + 배경색 + 테두리, 필수 항목은 이름 끝에 * 표시
- 헤더 아래로 회신자가 채울 빈 행을 넉넉히(100행 안팎) 두고, 그 범위에 입력 형식에 맞는
  유효성 검사를 겁니다
- 유효성 검사 type은 {", ".join(VALIDATION_TYPES)} 만 씁니다(정수는 whole). 드롭다운은
  type을 list로 두고 선택지를 source 배열에 넣습니다
- **선택지가 사양에 실제 목록으로 주어진 경우에만** 드롭다운을 만듭니다. "코드표 사용",
  "전체 부서"처럼 목록이 아닌 설명만 있으면 드롭다운을 만들지 말고 자유 입력으로 두세요 —
  없는 선택지를 지어내면 회신자가 값을 넣을 수 없습니다
- 열 너비는 항목 성격에 맞게 지정
- 헤더 행에 freeze_panes를 걸어 스크롤해도 보이게 합니다

제약(위반 시 파일로 만들지 않습니다):
- 시트 {MAX_SHEETS}개 이하, 셀 합계 {MAX_CELLS}개 이하
- 시트명 31자 이하, {"".join(sorted(BAD_SHEET_CHARS))} 사용 금지, 중복 금지
- 수식은 {", ".join(ALLOWED_FUNCS)} 만 사용, 다른 파일 참조 금지
- style 속성은 {", ".join(STYLE_KEYS)} 만 사용

반드시 아래 JSON만 출력합니다.
{{
  "sheets": [
    {{
      "name": "시트명",
      "columns": [{{"index": 1, "width": 14}}],
      "merges": ["A1:F1"],
      "freeze_panes": "A4",
      "styles": {{"title": {{"bold": true, "size": 14, "align": "center"}}}},
      "cells": [{{"ref": "A1", "value": "제목", "style": "title"}}],
      "validations": [{{"range": "B4:B103", "type": "date",
                        "prompt": "YYYY-MM-DD 형식으로 입력", "error": "날짜 형식이 올바르지 않습니다"}}]
    }}
  ]
}}"""


RECOGNIZE_SYSTEM = f"""당신은 이미 쓰이고 있는 엑셀 양식을 읽고 그 양식의 작성기준을
알아내는 분석가입니다. 사용자는 이 양식을 **새로 만들 생각이 없고 그대로 쓰려고** 합니다.
당신이 하는 일은 양식을 고치는 것이 아니라, 회신본을 검토할 때 쓸 항목 정보를 확정하는
것입니다.

먼저 사용자의 의도를 판정합니다(intent).
- recognize: 첨부한 양식을 그대로 쓰겠다는 뜻(예: "이 양식으로 받을 거야", "기존 양식이야")
- generate: 첨부는 참고일 뿐 새 양식을 만들어 달라는 뜻(예: "이거 비슷하게 새로 만들어줘")
- ambiguous: 둘 중 어느 쪽인지 판단할 수 없음. reply에 어느 쪽인지 묻는 질문을 담습니다

structure에는 파일에서 찾은 **표 목록**이 들어 있습니다(표마다 헤더·기존 표시형식·기존
유효성검사·샘플 행). 이것을 근거로 항목별 자료 유형·입력 형식·필수 여부를 추론합니다.

**표가 둘 이상이면 어느 표로 회신을 받을지 먼저 정해야 합니다.** 실제 양식에는 작성 가이드,
비워 둔 대장, 월×지표 총괄표처럼 회신 대상이 아닌 표가 섞여 있습니다. 이름과 헤더만으로
확실하지 않으면 **추측하지 말고 사용자에게 물어보세요.** 정한 표 이름을 selected_sheets에
그대로(structure의 name과 글자 단위로 같게) 담고, fields는 **고른 표의 헤더만** 다룹니다.
표가 하나면 selected_sheets에 그 하나를 담습니다.

지켜야 할 것:
- **항목 이름(name)은 헤더 문자열을 글자 그대로 옮깁니다.** 다듬거나 번역하거나 * 표시를
  떼지 마세요. 이 이름으로 회신본의 열을 찾기 때문에, 한 글자만 달라도 그 항목의 검토가
  통째로 빗나갑니다.
- 헤더를 빠뜨리거나 없는 항목을 더하지 마세요. structure의 헤더와 fields는 1:1입니다.
- **필수/선택은 근거가 있을 때만 confidence를 high로 둡니다.** 근거가 되는 것은 헤더 끝의
  `*` 표시, 양식의 작성 안내 문구, 사용자가 말해준 것입니다. 표시가 없다는 것만으로 선택이라고
  단정하지 마세요. 샘플 행이 다 채워져 있다는 것도 근거가 아닙니다(회신자가 성실했을 뿐입니다).
  근거가 없으면 **low로 두고 어느 항목이 필수인지 물어보세요** — 여기가 틀리면 회신본 검토에서
  누락을 못 잡거나 없는 누락을 잡습니다.
- existing_number_format이 `yyyy-mm-dd`·`#,##0` 처럼 형식을 알려주면 그대로 format에
  옮깁니다(추측이 아니라 파일에 적힌 값입니다). 값이 null이면 형식을 지정하지 않은 열이라
  근거가 없다는 뜻입니다 — 지어내지 마세요.
- existing_validation의 list source는 그 항목의 코드표입니다. notes에 남기세요.
- 근거가 있는 항목은 confidence를 "high", 파일만 봐서는 알 수 없어 사용자에게 물어야 하는
  항목은 "low"로 둡니다. **모르는 것을 그럴듯하게 채우지 말고 low로 두고 물어보세요.**
- 질문은 한 번에 한 주제만, low인 항목에 대해서만 합니다. 다 확인됐으면 그때
  recognize_complete를 true로 둡니다.
- type은 {", ".join(FIELD_TYPES)} 중 하나입니다. date·amount는 format을 반드시 채웁니다.

반드시 아래 JSON만 출력합니다.
{{
  "intent": "recognize|generate|ambiguous",
  "reply": "사용자에게 보여줄 한국어 메시지",
  "spec_json": {{
    "form_title": "양식 이름(파일명·제목 행에서)",
    "target_depts": ["전체 부서"],
    "selected_sheets": ["회신받을 표 이름(structure의 name 그대로)"],
    "fields": [
      {{"name": "헤더 그대로", "type": "text|number|amount|date|name|dept_code",
       "required": true, "format": "형식 또는 null", "notes": "비고 또는 null",
       "confidence": "high|low"}}
    ],
    "layout_hints": null,
    "locale": "ko"
  }},
  "recognize_complete": false
}}"""


class FormGenError(Exception):
    """생성을 진행할 수 없는 상태. 사용자에게 그대로 보여줄 문장을 담는다."""


# ── LLM 호출 ─────────────────────────────────────────────────────────────────

def _client():
    """OpenAI 클라이언트.

    F2의 2단계 재검증은 AI가 없으면 건너뛰고 취합을 계속하지만(§13.3), 양식 생성은
    AI가 기능 자체다. 조용히 빈 양식을 만드는 대신 사유를 들고 실패한다(§9.4).
    """
    try:
        from openai import OpenAI
        return OpenAI()
    except Exception as exc:
        raise FormGenError(
            "AI 서비스에 연결할 수 없어 양식을 생성할 수 없습니다. "
            f"잠시 후 다시 시도해주세요. (사유: {exc})") from exc


def _ask(system: str, payload: str, model: str, client=None) -> dict:
    """JSON 응답 1회. payload는 마스킹을 이미 통과한 문자열이어야 한다."""
    client = client or _client()
    try:
        res = client.chat.completions.create(
            model=model,
            response_format={"type": "json_object"},
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": payload}],
        )
        text = res.choices[0].message.content
        # 유료 호출이라 사용량을 남긴다. 스텁 클라이언트에는 usage가 없다.
        usage = getattr(res, "usage", None)
        if usage is not None:
            print(f"[formgen] {model} 토큰 in={usage.prompt_tokens} "
                  f"out={usage.completion_tokens} total={usage.total_tokens}", file=sys.stderr)
    except FormGenError:
        raise
    except Exception as exc:
        raise FormGenError(f"AI 응답을 받지 못했습니다. 다시 시도해주세요. (사유: {exc})") from exc
    try:
        return json.loads(text)
    except (TypeError, json.JSONDecodeError) as exc:
        raise FormGenError("AI 응답을 해석할 수 없습니다. 다시 시도해주세요.") from exc


def _masked(obj) -> str:
    """외부로 나가는 payload는 전부 마스킹을 통과시킨다 (§9.2·§13.2)."""
    return ag.mask(json.dumps(obj, ensure_ascii=False))


# ── 문답 ─────────────────────────────────────────────────────────────────────

def rubric_gaps(spec: dict) -> list[str]:
    """사양이 생성 가능한 상태인지 확인하고, 부족한 것을 돌려준다 (§4.3 가드 1).

    LLM이 spec_complete=true를 줘도 이 목록이 비지 않으면 완료로 인정하지 않는다.
    """
    gaps = []
    if not (spec.get("form_title") or "").strip():
        gaps.append("양식 이름")
    fields = spec.get("fields") or []
    if not fields:
        gaps.append("포함할 항목")
    for f in fields:
        name = (f.get("name") or "").strip() or "(이름 없는 항목)"
        if f.get("type") not in FIELD_TYPES:
            gaps.append(f"'{name}'의 자료 유형")
        if not isinstance(f.get("required"), bool):
            gaps.append(f"'{name}'의 필수/선택 구분")
        if f.get("type") in ("date", "amount") and not (f.get("format") or "").strip():
            gaps.append(f"'{name}'의 입력 형식")
    return list(dict.fromkeys(gaps))


def intake_turn(messages: list[dict], spec: dict, model: str,
                attachments: list[dict] | None = None, client=None) -> dict:
    """문답 한 턴. `{reply, spec_json, spec_complete, coverage}`를 돌려준다.

    messages는 원문 그대로 받고(저장본), 전송 직전에만 마스킹한다.
    """
    turn = len([m for m in messages if m.get("role") == "user"])
    payload = {
        "conversation": [{"role": m.get("role"), "content": m.get("content")} for m in messages],
        "current_spec": spec or {},
        "attachments": [{"name": a.get("original_name"), "kind": a.get("kind"),
                         "content": (a.get("extracted") or "")[:ATTACH_CHARS]}
                        for a in (attachments or [])],
    }
    if turn >= TURN_LIMIT:
        # §4.3 가드 2 — 무한 문답 방지. 지금까지 정보로 마감하게 한다
        payload["instruction"] = ("턴 수 상한에 도달했습니다. 더 묻지 말고 지금까지 확인된 "
                                 "내용으로 사양을 마감하세요.")

    out = _ask(INTAKE_SYSTEM, _masked(payload), model, client)
    new_spec = out.get("spec_json") if isinstance(out.get("spec_json"), dict) else (spec or {})
    coverage = out.get("coverage") if isinstance(out.get("coverage"), dict) else {}
    coverage.setdefault("confirmed_topics", [])
    coverage.setdefault("remaining_topics", [])

    gaps = rubric_gaps(new_spec)
    complete = bool(out.get("spec_complete")) and not gaps
    if bool(out.get("spec_complete")) and gaps:
        # 부족한 항목을 남은 주제로 되돌려 다음 턴에 LLM이 다시 묻게 한다
        coverage["remaining_topics"] = list(dict.fromkeys(list(coverage["remaining_topics"]) + gaps))

    return {
        "reply": str(out.get("reply") or "").strip() or "조금 더 알려주세요.",
        "spec_json": new_spec,
        "spec_complete": complete,
        "coverage": coverage,
        "gaps": gaps,
        # 지금 처리한 턴까지의 개수. `turn+1`은 아직 오지 않은 턴을 세는 셈이라
        # intake_sessions.turn_count가 실제보다 1 많게 남았다(6턴 문답이 7로 기록).
        "turn": turn,
    }


# ── F6 기존 양식 인지 ─────────────────────────────────────────────────────────
# 구조 판정을 새로 쓰지 않는다. 실제 업무 양식 두 계열로 검증된 헤더·표 인식이
# aggregate에 이미 있고(계층 헤더 합성, 한 시트 여러 표, 합계·서식 행 제외), 두 번째
# 판정기를 만들면 **취합이 보는 헤더와 인지가 보는 헤더가 갈라진다.** 그러면 인지로 만든
# 작성기준의 항목명이 취합 때 열을 못 찾는다.

# 성명으로 보이는 열 판정. `서명`을 넣으면 **부서명**이 걸려 부서 값이 통째로 가려진다
# (실제로 걸렸다). 서명 열은 대개 `담당자 서명`이라 담당자로 잡힌다.
NAME_HEADERS = ("성명", "이름", "담당자", "작성자")


def _blank_form_headers(grid: list[list], merged: list) -> tuple | None:
    """값이 하나도 없는 빈 양식의 헤더를 찾는다 (인지 E5 — 빈 양식도 등록 대상).

    `aggregate._find_header`는 헤더 **다음 행에 값이 있을 때만** 헤더로 확정한다. 취합에서는
    맞다 — 데이터가 없으면 취합할 것이 없다. 하지만 배포 전 빈 양식은 헤더 아래가 전부
    비어 있는 것이 정상이라 그 규칙으로는 하나도 못 찾는다. 그래서 이 경로만 따로 둔다.
    """
    best, count = -1, 0
    for i, row in enumerate(grid):
        filled = [v for v in row if not ag._is_blank(v)]
        if len(filled) >= 2 and all(isinstance(v, str) for v in filled) and len(filled) > count:
            best, count = i, len(filled)
    if best < 0:
        return None
    top, bottom = ag._header_block(grid, best, merged)
    return (top, ag._compose_headers(grid, top, bottom, merged), [], [], [], [])


def _validation_map(ws) -> dict:
    """열 문자 → 그 열에 걸린 기존 유효성 검사 (인지 §6.2 existing_validation).

    드롭다운 목록은 그 항목의 코드표라는 근거다 — 추론이 아니라 파일에 적혀 있는 값이다.
    """
    out: dict[str, dict] = {}
    for dv in getattr(ws.data_validations, "dataValidation", []) or []:
        source = None
        formula = str(dv.formula1 or "")
        if dv.type == "list" and formula.startswith('"') and formula.endswith('"'):
            source = [s.strip() for s in formula[1:-1].split(",") if s.strip()]
        info = {"type": dv.type, "source": source}
        for rng in str(dv.sqref or "").split():
            letters = re.findall(r"([A-Z]{1,3})", rng.replace("$", ""))
            for letter in letters[:1] or []:
                out.setdefault(letter, info)
    return out


def read_structure(path: Path) -> dict:
    """xlsx 양식의 구조를 뽑는다 (인지 §6.2). 파일은 열어서 읽기만 한다(E1 무변경).

    `sample_rows`는 예시 데이터가 든 파일에서만 채워지고 빈 양식이면 빈 배열이다(E5).
    성명으로 보이는 열의 값은 아예 담지 않는다(§9.2 J4) — 예시 데이터는 실제 회신본일
    가능성이 높고, 유형 판정에는 헤더만으로 충분해서 값을 들고 있을 이득이 없다.
    """
    import openpyxl
    from openpyxl.utils import get_column_letter

    if path.suffix.lower() != ".xlsx":
        raise FormGenError(f"등록할 수 있는 양식은 xlsx뿐입니다({path.suffix}는 등록 대상이 "
                           "아닙니다).")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:
        raise FormGenError(f"양식 파일을 열 수 없습니다(손상 또는 암호 보호 가능): {exc}")

    sheets, omitted = [], []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        merged = [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]
        tables = ag._split_tables(grid, merged) or []
        if not tables:
            blank = _blank_form_headers(grid, merged)
            if blank is None:
                continue                     # 표가 아닌 시트(작성 안내 등) — 등록 대상 아님
            tables = [blank]
        dvs = _validation_map(ws)
        for part, (hi, headers, rows, row_numbers, *_rest) in enumerate(tables, 1):
            # 세로 병합 값 채우기는 `read_file`이 `_split_tables` 뒤에 따로 한다. 여기서도
            # 통과시키지 않으면 병합된 `지사`·`개소`가 샘플에서 빈 칸으로 보여 모델이
            # '선택 항목'으로 오추론한다(실측 양식이 A5:A17처럼 병합돼 있다).
            ag._fill_merged(rows, row_numbers, merged)
            if not any(h for h in headers):
                continue
            if len(sheets) >= MAX_SHEETS:
                omitted.append(ws.title if len(tables) == 1 else f"{ws.title} ({part})")
                continue
            cols, samples = [], []
            for idx, header in enumerate(headers):
                if not header:
                    continue
                letter = get_column_letter(idx + 1)
                cell = ws.cell(row=hi + 2, column=idx + 1)
                # `General`은 표시형식을 지정하지 않았다는 뜻이다. 그대로 올려보내면
                # 모델이 그걸 입력 형식으로 옮겨 적어 'General'이 화면에 뜬다(실측).
                fmt = cell.number_format
                cols.append({"letter": letter, "header": header,
                             "existing_number_format": None if fmt == "General" else fmt,
                             "existing_validation": dvs.get(letter)})
            for row in rows[:SAMPLE_ROWS]:
                sample = {}
                for idx, header in enumerate(headers):
                    if not header or idx >= len(row):
                        continue
                    value = row[idx]
                    if any(k in header for k in NAME_HEADERS) and not ag._is_blank(value):
                        value = "(성명 마스킹)"
                    sample[header] = value if value is None or isinstance(
                        value, (str, int, float)) else str(value)
                samples.append(sample)
            sheets.append({
                "name": ws.title if len(tables) == 1 else f"{ws.title} ({part})",
                "header_row": hi + 1,
                "headers": [h for h in headers if h],
                "columns": cols,
                "sample_rows": samples,
            })

    if not sheets:
        raise FormGenError("이 파일에서 표(열 제목 행)를 찾지 못해 양식으로 등록할 수 "
                           "없습니다. 제목 행과 입력 칸 사이에 열 제목 행이 있는지 "
                           "확인해주세요.")
    out = {"sheets": sheets}
    if omitted:
        # 조용히 자르지 않는다 — 무엇이 빠졌는지 화면에서 밝힌다
        out["omitted_sheets"] = omitted
    return out


def selected_tables(spec: dict, structural: dict) -> list[dict]:
    """회신받을 표. 파일에 표가 하나면 고를 것이 없어 그것으로 본다.

    둘 이상이면 사용자가 대화에서 정한 것만 쓴다 — 실제 양식은 한 파일에 작성 가이드·
    비워 둔 대장·총괄표가 섞여 있어(실측: 한 파일에 표 6개·헤더 142개) 전 표를 항목으로
    다루면 인지 자체가 성립하지 않는다. 이름 규칙으로 맞히지 않는 것은 F2의 '취합할 표
    고르기'와 같은 이유다.
    """
    sheets = (structural or {}).get("sheets") or []
    if len(sheets) == 1:
        return list(sheets)
    picked = [str(n) for n in (spec.get("selected_sheets") or [])]
    return [s for s in sheets if s.get("name") in picked]


def structure_headers(structural: dict, spec: dict | None = None) -> list[str]:
    """회신받을 표의 헤더 전체(중복 제거, 순서 유지)."""
    seen: dict[str, None] = {}
    for sheet in selected_tables(spec or {}, structural):
        for header in sheet.get("headers") or []:
            name = str(header).strip()
            if name:
                seen.setdefault(name, None)
    return list(seen)


REQUIRED_MARKS = ("*", "필수")


def _required_evidence(spec: dict, structural: dict, messages: list[dict] | None) -> bool:
    """필수/선택을 판단할 근거가 있는가 — 양식의 표시이거나 사용자가 말해준 것.

    빈 양식·회신본 어느 쪽도 '이 칸을 꼭 채워야 하는지'를 파일에 담고 있지 않다. 값이 다
    채워져 있다는 것도 근거가 아니다(회신자가 성실했을 뿐이다).
    """
    for sheet in selected_tables(spec, structural):
        for header in sheet.get("headers") or []:
            if any(mark in str(header) for mark in REQUIRED_MARKS):
                return True
    return any(("필수" in (m.get("content") or "") or "선택" in (m.get("content") or ""))
               for m in messages or [] if m.get("role") == "user")


def recognize_gaps(spec: dict, structural: dict, messages: list[dict] | None = None) -> list[str]:
    """인지 §4.3 완결성 루브릭. `rubric_gaps`에 **헤더 1:1 대응**을 더한 것이다.

    1:1을 강하게 보는 이유는 취합이 항목명으로 열을 찾기 때문이다. 항목이 하나 빠지면
    그 열은 작성기준 없이 검토되고, 이름이 다듬어지면 그 항목의 검토가 통째로 빗나간다.
    """
    gaps = rubric_gaps(spec)
    sheets = (structural or {}).get("sheets") or []
    if len(sheets) > 1 and not selected_tables(spec, structural):
        names = ", ".join(str(s.get("name")) for s in sheets)
        return [f"회신받을 표 (이 파일에서 찾은 표: {names})"] + gaps
    headers = structure_headers(structural, spec)
    names = [(f.get("name") or "").strip() for f in spec.get("fields") or []]
    for header in headers:
        if header not in names:
            gaps.append(f"양식의 '{header}' 항목 정보")
    for name in names:
        if name and name not in headers:
            gaps.append(f"'{name}'은 양식에 없는 항목입니다")
    # 확신도가 낮은 항목은 확정으로 보지 않는다 (인지 §6.3·E7).
    fields = spec.get("fields") or []
    for f in fields:
        if str(f.get("confidence") or "").lower() == "low":
            gaps.append(f"'{(f.get('name') or '').strip()}' 확인 필요")

    # 근거가 없는데 전 항목의 필수/선택이 똑같이 나오면 그것은 판정이 아니라 기본값이다.
    # 실측에서 모델이 8개 항목을 전부 '선택 · 확인됨'으로 내려보냈다 — 프롬프트로
    # "근거 없으면 물어봐"라고 못 박아도 회차마다 흔들려서 여기서 막는다. 그대로 등록하면
    # 작성기준이 아무 누락도 잡지 못하는데 화면에는 '확인됨'으로 뜬다.
    if fields and len({bool(f.get("required")) for f in fields}) == 1 \
            and not _required_evidence(spec, structural, messages):
        gaps.append("필수 입력 항목 지정(양식에 * 표시가 없어 파일만으로는 알 수 없습니다)")
    return list(dict.fromkeys(gaps))


def recognize_turn(messages: list[dict], spec: dict, structural: dict, model: str,
                   attachments: list[dict] | None = None, client=None) -> dict:
    """인지 모드 한 턴. 의도 판정(F6-1)과 구조 인지(F6-3)를 한 호출로 한다.

    명세 §6.1은 `classify_intent`와 `recognize_turn`을 따로 두되 한 호출로 합쳐도 된다고
    했다. 합친다 — 첨부한 양식이 있을 때만 이 경로를 타므로, 의도를 물을 상황과 구조를
    읽을 상황이 언제나 같이 온다. 호출을 나누면 매 턴 값을 두 번 내는 셈이다.
    """
    turn = len([m for m in messages if m.get("role") == "user"])
    payload = {
        "conversation": [{"role": m.get("role"), "content": m.get("content")} for m in messages],
        "current_spec": spec or {},
        "structure": structural,
        "attachments": [{"name": a.get("original_name"), "kind": a.get("kind"),
                         "content": (a.get("extracted") or "")[:ATTACH_CHARS]}
                        for a in (attachments or [])],
    }
    # 백엔드가 아직 요구하는 것을 알려준다. 이것 없이는 모델이 "확인할 질문 없습니다"라고
    # 하고 사용자만 gaps 패널을 보게 되어 대화가 수렴하지 않는다.
    still = recognize_gaps(spec or {}, structural, messages) if spec else []
    if still:
        payload["still_missing"] = still
    if turn >= RECOGNIZE_TURN_LIMIT:
        payload["instruction"] = ("턴 수 상한에 도달했습니다. 더 묻지 말고 파일에서 확인된 "
                                 "내용으로 항목 정보를 마감하세요.")

    out = _ask(RECOGNIZE_SYSTEM, _masked(payload), model, client)
    intent = out.get("intent")
    if intent not in ("recognize", "generate", "ambiguous"):
        intent = "ambiguous"                 # 판정을 못 받으면 사용자에게 되묻는다(R6)
    new_spec = out.get("spec_json") if isinstance(out.get("spec_json"), dict) else (spec or {})

    gaps = recognize_gaps(new_spec, structural, messages)
    complete = intent == "recognize" and bool(out.get("recognize_complete")) and not gaps
    return {
        "intent": intent,
        "reply": str(out.get("reply") or "").strip() or "조금 더 알려주세요.",
        "spec_json": new_spec,
        "spec_complete": complete,
        "coverage": {"confirmed_topics": [], "remaining_topics": gaps},
        "gaps": gaps,
        "turn": turn,
    }


# ── 저작 · 검증 ───────────────────────────────────────────────────────────────

def _check_value(value, where: str) -> list[str]:
    if isinstance(value, bool) or value is None:
        return [] if value is None else [f"{where}: 참/거짓 값은 쓸 수 없습니다"]
    if isinstance(value, (int, float)):
        return []
    if not isinstance(value, str):
        return [f"{where}: 문자열·숫자만 쓸 수 있습니다"]
    if not value.startswith("="):
        return []
    bad = []
    used = {m.group(1).upper() for m in FUNC_RE.finditer(value)}
    for func in used - {f.upper() for f in ALLOWED_FUNCS}:
        bad.append(f"{where}: 허용되지 않은 함수 {func}")
    if "[" in value or "]" in value:
        bad.append(f"{where}: 다른 파일을 참조하는 수식은 쓸 수 없습니다")
    return bad


def normalize_workbook(wb: dict) -> dict:
    """같은 뜻의 다른 이름을 표준 키로 옮긴다. **값을 만들어내지는 않는다.**

    실측에서 모델이 드롭다운 목록을 `source` 대신 `values`로, 정수 검사를 `whole` 대신
    `integer`로 내보냈다. 회차마다 흔들리는 부분이라 프롬프트만으로는 막히지 않고, 이름이
    다를 뿐 같은 제약이라 그대로 받아 옮긴다. 여기서 걸러지지 않는 것은 검증이 막는다.
    """
    if not isinstance(wb, dict):
        return wb
    for sheet in wb.get("sheets") or []:
        if not isinstance(sheet, dict):
            continue
        for dv in sheet.get("validations") or []:
            if not isinstance(dv, dict):
                continue
            if dv.get("type") in ("integer", "int"):
                dv["type"] = "whole"
            if not dv.get("source"):
                for alias in ("values", "options", "list", "items"):
                    if isinstance(dv.get(alias), list) and dv[alias]:
                        dv["source"] = dv[alias]
                        break
    return wb


def validate_workbook(wb: dict) -> list[str]:
    """명세 §6.3 검증. 위반 목록을 돌려준다(빈 목록이면 통과).

    LLM 산출물을 그대로 파일로 만들지 않기 위한 관문이다.
    """
    problems: list[str] = []
    sheets = wb.get("sheets") if isinstance(wb, dict) else None
    if not isinstance(sheets, list) or not sheets:
        return ["시트가 없습니다"]
    if len(sheets) > MAX_SHEETS:
        problems.append(f"시트가 {len(sheets)}개입니다(최대 {MAX_SHEETS}개)")

    seen_names, total_cells = set(), 0
    for i, sheet in enumerate(sheets):
        where = f"시트 {i + 1}"
        if not isinstance(sheet, dict):
            problems.append(f"{where}: 형식이 올바르지 않습니다")
            continue
        name = str(sheet.get("name") or "").strip()
        if not name:
            problems.append(f"{where}: 시트명이 비었습니다")
        if len(name) > 31:
            problems.append(f"{where}: 시트명이 31자를 넘습니다")
        if set(name) & BAD_SHEET_CHARS:
            problems.append(f"{where}: 시트명에 쓸 수 없는 문자가 있습니다")
        if name in seen_names:
            problems.append(f"{where}: 시트명 '{name}'이 중복됩니다")
        seen_names.add(name)

        cells = sheet.get("cells") or []
        if not isinstance(cells, list):
            problems.append(f"{where}: cells 형식이 올바르지 않습니다")
            cells = []
        total_cells += len(cells)
        styles = sheet.get("styles") if isinstance(sheet.get("styles"), dict) else {}
        for key, style in styles.items():
            if not isinstance(style, dict):
                problems.append(f"{where}: 스타일 '{key}' 형식이 올바르지 않습니다")
                continue
            for attr in set(style) - set(STYLE_KEYS):
                problems.append(f"{where}: 스타일 '{key}'의 허용되지 않은 속성 {attr}")

        for cell in cells:
            if not isinstance(cell, dict) or not REF_RE.match(str(cell.get("ref") or "")):
                problems.append(f"{where}: 셀 위치 표기가 올바르지 않습니다 ({cell!r})"[:120])
                continue
            problems += _check_value(cell.get("value"), f"{where} {cell['ref']}")

        for rng in sheet.get("merges") or []:
            if not RANGE_RE.match(str(rng)) or ":" not in str(rng):
                problems.append(f"{where}: 병합 범위 표기가 올바르지 않습니다 ({rng!r})")
        fp = sheet.get("freeze_panes")
        if fp and not REF_RE.match(str(fp)):
            problems.append(f"{where}: freeze_panes 표기가 올바르지 않습니다 ({fp!r})")

        for dv in sheet.get("validations") or []:
            if not isinstance(dv, dict) or not RANGE_RE.match(str(dv.get("range") or "")):
                problems.append(f"{where}: 유효성 검사 범위 표기가 올바르지 않습니다")
                continue
            if dv.get("type") not in VALIDATION_TYPES:
                problems.append(f"{where}: 지원하지 않는 유효성 검사 종류 {dv.get('type')!r}")
            if dv.get("type") == "list" and not (dv.get("source") or []):
                problems.append(f"{where}: 드롭다운 목록이 비었습니다")

    if total_cells > MAX_CELLS:
        problems.append(f"셀이 {total_cells:,}개입니다(최대 {MAX_CELLS:,}개)")
    return problems


def author_workbook(spec: dict, model: str, attachments: list[dict] | None = None,
                    client=None, on_retry=None) -> dict:
    """사양대로 workbook_json을 저작한다. 검증을 통과한 것만 돌려준다.

    검증에 걸리면 위반 목록을 그대로 알려 다시 저작하게 한다(최대 AUTHOR_RETRY회).
    """
    client = client or _client()
    payload = {
        "spec": spec,
        "attachments": [{"name": a.get("original_name"), "kind": a.get("kind"),
                         "content": (a.get("extracted") or "")[:ATTACH_CHARS]}
                        for a in (attachments or [])],
    }
    problems: list[str] = []
    for attempt in range(AUTHOR_RETRY + 1):
        if problems:
            payload["fix_these"] = problems
            if on_retry:
                on_retry(attempt, problems)
        wb = normalize_workbook(_ask(AUTHOR_SYSTEM, _masked(payload), model, client))
        problems = validate_workbook(wb)
        if not problems:
            return wb
    raise FormGenError("양식 파일을 만들 수 없었습니다. 검증에서 걸린 항목: "
                       + "; ".join(problems[:5]))


# ── xlsx 변환 ────────────────────────────────────────────────────────────────

def materialize(wb: dict, out: Path) -> list[str]:
    """workbook_json을 xlsx로 옮긴다. 해석·보정을 하지 않는 1:1 변환이다.

    안전 기본값만 둔다 — 정의 안 된 스타일은 무서식, 너비 미지정 열은 12.
    돌려주는 것은 **낮춰서 처리한 것들의 목록**이다(빈 목록이면 그대로 옮겨졌다는 뜻).
    """
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    notes: list[str] = []
    book = openpyxl.Workbook()
    book.remove(book.active)
    for sheet in wb["sheets"]:
        ws = book.create_sheet(str(sheet.get("name") or "양식")[:31])
        styles = sheet.get("styles") if isinstance(sheet.get("styles"), dict) else {}

        for cell in sheet.get("cells") or []:
            target = ws[cell["ref"]]
            target.value = cell.get("value")
            style = styles.get(cell.get("style")) or {}
            if not style:
                continue                                  # 정의 없는 스타일 참조 → 무서식
            if style.get("bold") or style.get("italic") or style.get("size") or style.get("color"):
                target.font = Font(bold=bool(style.get("bold")), italic=bool(style.get("italic")),
                                   size=style.get("size"),
                                   color=str(style["color"]).lstrip("#") if style.get("color") else None)
            if style.get("bg"):
                fill = str(style["bg"]).lstrip("#")
                target.fill = PatternFill("solid", start_color=fill, end_color=fill)
            if style.get("border"):
                side = Side(style=str(style["border"]))
                target.border = Border(left=side, right=side, top=side, bottom=side)
            if style.get("align"):
                target.alignment = Alignment(horizontal=str(style["align"]), vertical="center",
                                             wrap_text=True)
            if style.get("number_format"):
                target.number_format = str(style["number_format"])

        for col in sheet.get("columns") or []:
            try:
                idx = int(col.get("index"))
            except (TypeError, ValueError):
                continue
            ws.column_dimensions[get_column_letter(idx)].width = col.get("width") or 12

        for rng in sheet.get("merges") or []:
            ws.merge_cells(str(rng))
        if sheet.get("freeze_panes"):
            ws.freeze_panes = str(sheet["freeze_panes"])

        for dv in sheet.get("validations") or []:
            kind = dv.get("type")
            if kind == "list":
                items = ",".join(str(s).replace(",", " ") for s in dv.get("source") or [])
                obj = DataValidation(type="list", formula1=f'"{items}"', allow_blank=True)
            elif kind == "date":
                # 날짜 검사는 비교식을 요구한다. 실무 데이터를 걸러내지 않는 하한을 쓴다
                obj = DataValidation(type="date", operator="greaterThan",
                                     formula1="DATE(1900,1,1)", allow_blank=True)
            else:
                lo, hi = dv.get("min"), dv.get("max")
                if lo is not None and hi is not None:
                    obj = DataValidation(type=kind, operator="between", formula1=str(lo),
                                         formula2=str(hi), allow_blank=True)
                elif lo is not None:
                    obj = DataValidation(type=kind, operator="greaterThanOrEqual",
                                         formula1=str(lo), allow_blank=True)
                else:
                    # 숫자 검사인데 범위가 없다. operator·수식 없는 숫자 검사는 엑셀이
                    # 손상으로 보므로 그대로 쓸 수 없고, 상·하한을 지어내면 사용자가
                    # 요구하지 않은 제약이 양식에 박힌다. 안내 문구만 남기고 밝힌다.
                    obj = DataValidation(allow_blank=True)
                    notes.append(f"{ws.title} {dv['range']}: 숫자 범위가 정해지지 않아 "
                                 "입력 안내만 넣고 형식 검사는 걸지 않았습니다")
            obj.prompt, obj.promptTitle = dv.get("prompt") or "", "입력 형식"
            obj.error, obj.errorTitle = dv.get("error") or "", "입력 오류"
            obj.showInputMessage = bool(dv.get("prompt"))
            obj.showErrorMessage = bool(dv.get("error"))
            ws.add_data_validation(obj)
            obj.add(str(dv["range"]))

    book.save(out)
    return notes


# ── 작성기준 파생 (F1-5·F1-6) ────────────────────────────────────────────────

def derive_field_rules(spec: dict) -> list[dict]:
    """spec.fields → field_rules 행. LLM을 다시 부르지 않는 기계적 파생이다."""
    rows = []
    for f in spec.get("fields") or []:
        name = (f.get("name") or "").strip()
        if not name:
            continue
        config: dict = {"required": bool(f.get("required"))}
        if (f.get("format") or "").strip():
            config["format"] = f["format"].strip()
        if (f.get("notes") or "").strip():
            config["notes"] = f["notes"].strip()
        rows.append({
            "field_name": name,
            "rule_type": f["type"] if f.get("type") in FIELD_TYPES else "text",
            "rule_config_json": config,
        })
    return rows


def rules_for_aggregate(spec: dict) -> dict:
    """취합 엔진(F2)이 그대로 먹는 작성기준으로 바꾼다.

    `aggregate.review_stage1`의 rules는 `{헤더명: {"required": bool, "key": bool}}` 형태다.
    **키 컬럼은 넣지 않는다** — 사양에 대응 개념이 없어서 첫 항목이나 dept_code를 키로
    단정하면 없는 근거로 값을 만드는 셈이다. F2의 자동 판정에 맡기고 화면에서 바꾼다.
    """
    out = {}
    for f in spec.get("fields") or []:
        name = (f.get("name") or "").strip()
        if name:
            out[name] = {"required": bool(f.get("required"))}
    return out


# ── CLI ──────────────────────────────────────────────────────────────────────

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="AI 양식 생성 (문답 → xlsx)")
    ap.add_argument("prompt", nargs="?", help="첫 요청. 생략하면 입력받는다")
    ap.add_argument("-o", "--out", type=Path, default=Path("form.xlsx"))
    ap.add_argument("--rules-out", type=Path, help="작성기준 JSON 경로(취합의 --rules에 그대로 쓴다)")
    ap.add_argument("--model", default=None, help="기본값은 .env의 OPENAI_MODEL")
    a = ap.parse_args(argv)

    ag.load_env(Path(__file__).parent / ".env")
    import os
    model = a.model or os.environ.get("OPENAI_MODEL", "gpt-5-mini")

    first = a.prompt or input("어떤 양식이 필요하신가요? > ").strip()
    if not first:
        print("요청 내용이 없습니다.")
        return 1

    messages: list[dict] = [{"role": "user", "content": first}]
    spec: dict = {}
    try:
        while True:
            turn = intake_turn(messages, spec, model)
            spec = turn["spec_json"]
            messages.append({"role": "assistant", "content": turn["reply"]})
            print(f"\nAI: {turn['reply']}")
            if turn["spec_complete"]:
                break
            if turn["gaps"]:
                print(f"  (아직 필요한 것: {', '.join(turn['gaps'])})")
            answer = input("\n나: ").strip()
            if not answer:
                print("답변이 없어 문답을 멈춥니다.")
                return 1
            messages.append({"role": "user", "content": answer})

        print(f"\n사양 확정: {spec.get('form_title')} · 항목 {len(spec.get('fields') or [])}개")
        print("양식 파일을 만드는 중…")
        wb = author_workbook(spec, model,
                             on_retry=lambda n, p: print(f"  · 검증 미통과 {len(p)}건 → 재저작 {n}회차"))
        notes = materialize(wb, a.out)
        print(f"결과: {a.out}")
        for n in notes:
            print(f"  · 낮춰서 처리: {n}")

        rules = rules_for_aggregate(spec)
        if a.rules_out:
            a.rules_out.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"작성기준: {a.rules_out}  (취합 시 --rules {a.rules_out})")
        else:
            print("작성기준:", json.dumps(rules, ensure_ascii=False))
    except FormGenError as exc:
        print(f"\n실패: {exc}", file=sys.stderr)
        return 1
    except (KeyboardInterrupt, EOFError):
        print("\n중단했습니다.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
