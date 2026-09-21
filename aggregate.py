#!/usr/bin/env python3
"""F2 엑셀 취합 CLI.

docs/취합기능_기술명세_v0.1.md 구현:
  업로드(폴더 스캔) → 구조 인식 → 1단계 규칙 검증 → (파일 단위 게이팅) 2단계 AI 재검증
  → 이상판별 출력 → 범위 선택 → 전처리(자동교정) → 합성 A/B/C/D → merged.xlsx + error_report.xlsx

부서명은 파일명(확장자 제외)에서 가져온다. 부서별 회신 파일 1개 = 부서 1개 전제(§9 모드 A/B).
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlsx_format as xf
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import EMU_to_pixels, pixels_to_EMU

# ── 시스템 기본값 (§5.4: 작성기준 스키마에 값이 없으면 이 값을 사용) ──────────────
BLANK_ROW_TOLERANCE = 2          # §2.2 데이터 영역 중간 연속 공백 행 허용치
TYPE_SAMPLE_ROWS = 20            # §3 값 타입 분포 샘플링 행 수
FREEFORM_MAX_LENGTH = 500        # §5.4
IMAGE_ALLOWED_EXT = ("jpg", "jpeg", "png")
IMAGE_MAX_SIZE_MB = 10
IMAGE_MIN_RESOLUTION = (200, 200)
IMAGE_MAX_COUNT_PER_KEY = 1
SHEET_NAME_LIMIT = 31            # 엑셀 시트명 상한 (§9 모드 A 자르기 규칙)
# 실제 업무 양식의 표 끝에 붙는 행들 — 레코드가 아니라 표 장식이다
TOTAL_LABELS = ("합계", "소계", "총계", "누계", "총합계", "계")
NOTE_PREFIXES = ("*", "※", "주)", "비고)", "◦", "ㅇ", "•")
# 회신 파일의 소속을 담는 컬럼 이름 — 전 행이 비어 있으면 부서명으로 채우고(review_stage1),
# 값이 부서명과 같으면 합성에서 '부서' 컬럼을 따로 붙이지 않고 이 컬럼을 쓴다(synthesize).
# '주관부서'·'담당부서'는 다른 부서 몫이 비어 있는 것이 정상이라(§9) 넣지 않는다 —
# 정확히 일치하는 이름만 본다.
DEPT_HEADERS = ("지사", "지사명", "부서", "부서명", "사업소", "소속", "팀", "팀명", "기관", "기관명")
NOTES_MAX_CHARS = 3000           # 양식 안내문 컨텍스트 상한 — 안내는 길어도 기준은 앞쪽에 있다

ERROR, WARN, OK = "오류", "경고", "정상"
GRADE_ORDER = {ERROR: 2, WARN: 1, OK: 0}   # §7 대표 등급 = 최악 등급

# 기본 가정(작성기준이 없는 컬럼은 필수) 폭주 감지 — 하나의 가정이 같은 유형의
# 이슈를 시트에서 이만큼 만들어내면 데이터가 아니라 가정이 틀렸다는 신호로 본다.
# 실측 근거: 월별 점검표(빈 칸이 정상인 매트릭스형)가 필수값 누락 1,051건으로
# 보고돼 정작 봐야 할 칸이 묻혔다(9차). 두 조건을 모두 요구한다 — 절대치 하한은
# 작은 표의 진짜 누락 몇십 건을 접지 않기 위해, 비율 하한은 큰 표에서 진짜
# 누락이 수백 건인 경우(드물지만 전부 심각)를 접지 않기 위해서다.
DEFAULT_MISS_FLOOR = 30          # 미선언 컬럼 누락이 이 건수 이상이고
DEFAULT_MISS_RATIO = 0.3         # 미선언 컬럼 검사 칸의 이 비율 이상일 때만 접는다

NUM_KEYWORDS = ("날짜", "금액", "수량", "비율", "예산", "집행", "건수", "인원",
                "date", "amount", "count", "rate", "qty", "budget", "total")
CODE_KEYWORDS = ("부서", "성명", "이름", "코드", "직급", "구분",
                 "dept", "department", "name", "code", "id")
KEY_KEYWORDS = ("사번", "아이디", "코드", "id", "key", "no")

# §13.2 개인정보 자동 마스킹 — 외부 API로 나가는 모든 값에 예외 없이 적용
MASK_PATTERNS = (
    (re.compile(r"\d{6}\s*-\s*\d{7}"), "[주민번호]"),
    (re.compile(r"01\d[-\s]?\d{3,4}[-\s]?\d{4}"), "[전화번호]"),
    (re.compile(r"0\d{1,2}[-\s]\d{3,4}[-\s]\d{4}"), "[전화번호]"),
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "[이메일]"),
)


def mask(value) -> str:
    text = "" if value is None else str(value)
    for pattern, replacement in MASK_PATTERNS:
        text = pattern.sub(replacement, text)
    return text


# ── 데이터 모델 (§11) ──────────────────────────────────────────────────────────
@dataclass
class Issue:
    file: str
    sheet: str
    cell: str            # 파일/시트 단위 오류는 "(해당없음)" (§2.1)
    kind: str            # 누락 / 정합성 / 중복 / 충돌 / 소속범위 / AI 재검증 실패
    stage: str           # 1단계 / 2단계 / 2단계(실패)
    grade: str           # 오류 / 경고  (§6.2 내부 등급, 화면에는 노출 안 함)
    reason: str
    attr: str = ""

    @property
    def tag(self) -> str:
        # §6.3 사유 텍스트에 붙는 참고용 심각도 표기
        if self.kind == "AI 재검증 실패":
            return "[확인필요]"
        if self.grade == OK:
            return "[안내]"
        return "[심각]" if self.grade == ERROR else "[경미]"

    def line(self) -> str:
        return f"{self.file} - {self.sheet} - {self.cell}: {self.tag} {self.reason}"


@dataclass
class Fix:
    """§8 자동교정 이력 (감사 추적용: 원본값 → 교정값)."""
    file: str
    sheet: str
    cell: str
    original: str
    corrected: str
    rule: str


@dataclass
class SheetData:
    name: str
    header_row: int                  # 엑셀 실제 행 번호 (1-base), 0이면 인식 실패
    headers: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)
    row_numbers: list[int] = field(default_factory=list)   # rows[i]의 엑셀 실제 행 번호
    images: list[dict] = field(default_factory=list)
    col_types: list[str] = field(default_factory=list)
    merged: list[tuple[int, int, int, int]] = field(default_factory=list)  # (min_row, min_col, max_row, max_col)
    # _fill_merged가 값을 복사해 채운 세로 병합 구간 — (rows 첫 인덱스, 마지막 인덱스, 컬럼 0-base).
    # 결과에 다시 병합으로 되돌리는 데 쓴다(_restore_vmerges).
    vmerges: list[tuple[int, int, int]] = field(default_factory=list)
    # 원본 표 끝의 합계 행 — 취합에서는 뺐지만(이중 합산 방지) 결과 끝에 SUM 수식으로 다시 만든다.
    # total_cols는 원본 합계 행에 숫자가 있던 컬럼(0-base). 준공연도처럼 더하지 않는 컬럼을
    # 짐작으로 더하지 않기 위해 원본이 더한 컬럼만 더한다.
    total_label: str = ""
    total_label_col: int = 0
    total_cols: list[int] = field(default_factory=list)


@dataclass
class UploadedFile:
    path: Path
    dept: str
    sheets: list[SheetData] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    fixes: list[Fix] = field(default_factory=list)
    ai_unverified: bool = False      # §6.4 정상(AI 미검증)
    readable: bool = True
    pair_period: tuple[int, int] | None = None   # 계획-실적 검사에 실제로 쓴 기간(월)
    note_sheets: list[str] = field(default_factory=list)   # 표가 아닌 시트(작성 안내 등)
    notes_text: str = ""             # 그 시트들의 문구 — AI 검토·취합 계획의 '양식안내' 컨텍스트

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def grade(self) -> str:
        return max((i.grade for i in self.issues), key=lambda g: GRADE_ORDER[g], default=OK)

    @property
    def violations(self) -> list:
        """등급이 있는 이슈만. 안내(정상 등급) 이슈는 판정·2단계 게이팅에 영향을 주지 않는다."""
        return [i for i in self.issues if i.grade != OK]

    @property
    def status(self) -> str:
        # §6.1 화면 표기는 정상/이상 이진화 — 안내만 있는 파일은 정상이다
        return "이상" if self.violations else "정상"


# 조직명에 흔히 붙는 꼬리 — 파일명에서 부서명 후보를 고를 때의 강한 신호
ORG_SUFFIXES = ("지사", "사업소", "본부", "지역본부", "센터", "지점", "영업소",
                "사업단", "부", "과", "팀", "실", "국", "청", "시", "군", "구")
# 파일명에 섞이지만 부서명은 아닌 말들
NOT_DEPT_WORDS = ("양식", "서식", "최종", "수정", "사본", "복사본", "제출", "제출용",
                  "취합", "결과", "회신", "완료", "확인", "검토", "원본", "샘플",
                  "final", "copy", "draft", "temp")


def suggest_dept_names(stem: str, others: list[str] | None = None) -> list[str]:
    """파일명에서 부서명 후보를 뽑는다. 앞쪽이 더 그럴듯한 순서.

    부서명은 파일명에서 가져오는데(§9) 실제 회신 파일은 부서명 위치가 제각각이다.
    `(수원사업소)2025년도…`, `…리스트(대구)`, `…리스트(양식)_판교지사`처럼 앞·뒤·
    구분자 뒤에 흩어져 있어 한 가지 규칙으로는 못 잡는다. 그래서 규칙을 여럿 두고
    후보를 제시하고, 최종 선택은 사용자에게 맡긴다.

    others에 같이 올라온 다른 파일명을 주면 정확도가 크게 올라간다. 취합은 같은
    양식을 배포해 회신받은 파일들이므로 **모든 파일에 공통으로 들어간 말은 제목이고
    부서명은 파일마다 다른 부분**이다 — 이게 조직명 꼬리(지사·사업소)보다 강한
    신호다. 예: '지열지점'은 지점으로 끝나지만 8개 파일 전부에 있으니 제목이다.
    """
    stem = (stem or "").strip()
    candidates: list[str] = []

    def add(value: str) -> None:
        value = (value or "").strip(" _-·.")
        if not value or len(value) > 30:
            return
        if value.replace(" ", "").lower() in [w.lower() for w in NOT_DEPT_WORDS]:
            return
        if value not in candidates:
            candidates.append(value)

    # ① 괄호 안 — 실제 파일에서 가장 흔한 자리
    for group in re.findall(r"[（(\[]([^）)\]]+)[）)\]]", stem):
        add(group)
    # ② 구분자로 끊은 조각 (뒤에서부터 — 파일명 꼬리에 붙는 경우가 많다)
    without_groups = re.sub(r"[（(\[][^）)\]]*[）)\]]", " ", stem)
    parts = [p for p in re.split(r"[_\-–—\s]+", without_groups) if p]
    for part in reversed(parts):
        add(part)

    siblings = [s for s in (others or []) if s and s != stem]

    def common(value: str) -> int:
        """다른 파일명에도 들어 있으면 제목의 일부다(부서명은 파일마다 다르다)."""
        if not siblings:
            return 0
        hits = sum(1 for s in siblings if value in s)
        return 1 if hits >= max(1, len(siblings) // 2) else 0

    # 공통어가 아닌 것 → 조직명 꼬리를 가진 것 → 짧은 것 순
    def rank(value: str) -> tuple:
        return (common(value), 0 if value.endswith(ORG_SUFFIXES) else 1, len(value))

    candidates.sort(key=rank)
    # 마지막 보루로 파일명 전체(종전 동작)를 남겨 둔다
    if stem and stem not in candidates:
        candidates.append(stem)
    return candidates[:4]


# ── 2. 업로드 및 구조 인식 (F2-1) ─────────────────────────────────────────────
def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_dept_header(header) -> bool:
    return str(header or "").replace(" ", "").strip() in DEPT_HEADERS


def _total_row_of(grid: list[list], dropped: list[int]) -> tuple[str, int, list[int]]:
    """제외한 행 중 합계 행의 (표기, 표기 컬럼, 숫자가 있던 컬럼들). 없으면 ("", 0, []).

    data_only로 읽으므로 수식 칸은 캐시된 값이다 — 재계산 없이 저장된 파일이면 None이라
    합계 컬럼을 못 찾고, 그때는 결과에 합계 행을 만들지 않는다(지어내지 않는다).
    """
    for number in dropped:
        row = grid[number - 1] if 0 < number <= len(grid) else []
        filled = [(c, v) for c, v in enumerate(row) if not _is_blank(v)]
        if filled and str(filled[0][1]).replace(" ", "").strip() in TOTAL_LABELS:
            cols = [c for c, v in filled[1:] if _is_number(v)]
            if cols:
                return str(filled[0][1]).strip(), filled[0][0], cols
    return "", 0, []


def _is_summary_or_note(row: list) -> bool:
    """합계 행·각주 행인지. 데이터 레코드가 아니므로 취합에서 뺀다.

    합계 행을 레코드로 취급하면 오탐(빈 칸이 필수값 누락으로 잡힘)에 그치지 않고,
    부서별로 세로 누적하는 모드 B/D에서 합계가 한 번 더 더해져 결과가 틀린다.
    """
    filled = [v for v in row if not _is_blank(v)]
    if not filled:
        return False
    # 표가 A열부터 시작한다고 볼 수 없다 — 실제 양식은 B열부터 그리는 것이 흔하고,
    # 그러면 row[0]만 보는 판정은 '합계' 표기를 놓쳐 합계 행이 레코드로 들어온다.
    head = str(filled[0]).replace(" ", "").strip()
    if head in TOTAL_LABELS:
        return True
    # 칸 하나에 안내 문구만 있는 행 (표 아래 각주)
    if len(filled) == 1 and isinstance(filled[0], str):
        return filled[0].strip().startswith(NOTE_PREFIXES)
    return False


def _is_template_row(row: list, headers: list[str]) -> bool:
    """순번만 채워진 빈 서식 행인지.

    실무 양식은 순번을 1..N까지 미리 적어 두고 회신자가 그중 일부만 채운다(실측: 의견수렴
    시트에 25행이 미리 매겨져 있고 10건만 작성). 이 행을 레코드로 보면 남은 15행 × 컬럼 수
    만큼 '필수값 누락'이 쌓여 그 파일이 오류가 되고, 취합 결과에도 빈 행이 들어간다.

    판정은 좁게 둔다 — **표의 첫 컬럼에 숫자 하나만** 있고 나머지가 전부 빈 행.
    """
    if len(headers) < 3:
        return False
    filled = [(i, v) for i, v in enumerate(row) if not _is_blank(v)]
    if len(filled) != 1:
        return False
    first = next((i for i, h in enumerate(headers) if h), None)
    return filled[0][0] == first and _is_number(filled[0][1])


def _looks_like_data(row: list) -> bool:
    """헤더 후보 다음 행이 데이터인지: 숫자/날짜가 하나라도 섞여 있으면 데이터로 본다 (§2.2)."""
    return any(not isinstance(v, str) and not _is_blank(v) for v in row)


HEADER_BLOCK_MAX_TIERS = 3   # 실무 계층 헤더는 2~3단이다. 이보다 두꺼우면 헤더가 아니다


def _plausible_block(grid: list[list], hi: int,
                     merged: list[tuple[int, int, int, int]]) -> bool:
    """후보 행의 헤더 블록이 그럴듯한 두께인지.

    ●만 찍는 월별 점검표에서 이 검사가 없으면 첫 숫자(예산 칸)가 우연히 나오는 자리의
    윗행이 헤더로 확정되고, 데이터 구역의 그룹 병합(A7:A16 같은)을 따라 블록이 10행짜리로
    부풀어 법령 본문이 컬럼명이 된다(실측, 9차). 진짜 계층 헤더는 2~3단을 넘지 않는다.
    """
    top, bottom = _header_block(grid, hi, merged)
    return bottom - top + 1 <= HEADER_BLOCK_MAX_TIERS


def _find_header(grid: list[list],
                 merged: list[tuple[int, int, int, int]] = ()) -> int:
    """헤더 행 인덱스(0-base)를 찾는다. 못 찾으면 -1. (§2.2 제목행 오인식 방지)"""
    for i, row in enumerate(grid):
        filled = [v for v in row if not _is_blank(v)]
        if len(filled) < 2:
            continue                      # 단일 병합 제목행 → 헤더 후보 제외
        if not all(isinstance(v, str) for v in filled):
            continue                      # 헤더는 순수 텍스트
        nxt = grid[i + 1] if i + 1 < len(grid) else []
        if not [v for v in nxt if not _is_blank(v)]:
            continue                      # 다음 행 공백 → 후보 제외
        if _looks_like_data(nxt) and _plausible_block(grid, i, list(merged)):
            return i                      # 다음 행이 숫자/날짜를 보임 → 헤더 확정
    # 전 컬럼이 텍스트인 시트(숫자 없음)는 위 규칙으로 확정되지 않으므로,
    # 제목행이 아닌 첫 텍스트 행을 헤더로 채택한다.
    for i, row in enumerate(grid):
        filled = [v for v in row if not _is_blank(v)]
        if len(filled) >= 2 and all(isinstance(v, str) for v in filled):
            if [v for v in (grid[i + 1] if i + 1 < len(grid) else []) if not _is_blank(v)] \
                    and _plausible_block(grid, i, list(merged)):
                return i
    return -1


def _header_block(grid: list[list], hi: int,
                  merged: list[tuple[int, int, int, int]]) -> tuple[int, int]:
    """헤더가 몇 행에 걸쳐 있는지 `(첫 행, 마지막 행)`을 0-base로 돌려준다.

    계층 헤더의 실제 표기법은 정해져 있다 — 단일 컬럼은 상·하위 행을 **세로로 병합**하고
    (`B35:C36`), 2단 컬럼은 상위를 **가로로 병합**한 뒤 하위 행에 자식을 둔다(`AE35:AI35`
    아래 가능성·중대성·위험성). 그래서 병합만 보면 범위가 정확히 나온다.

    `_find_header`는 '다음 행이 데이터로 보이는' 행을 고르므로 계층 헤더에서는 **맨 아래 단**을
    집는다. 위로 걸쳐 있는 병합을 따라 올라가 첫 단을 찾는다.
    """
    top = hi
    while True:
        # 이 행이 위에서 시작한 병합 안에 있으면 그 시작 행이 헤더의 윗단이다
        above = [r0 for r0, _, r1, _ in merged if r0 <= top and top + 1 <= r1 and r0 < top + 1]
        if not above:
            break
        top = max(above) - 1
    bottom = max([hi] + [r1 - 1 for r0, _, r1, _ in merged if r0 == top + 1])
    return top, min(bottom, len(grid) - 1)


def _compose_headers(grid: list[list], top: int, bottom: int,
                     merged: list[tuple[int, int, int, int]]) -> list[str]:
    """계층 헤더를 컬럼당 한 줄로 합친다. `조치 전 위험성평가 중대성` 처럼.

    합치지 않으면 하위 단만 남아 `중대성`이 조치 전·후 두 컬럼에 같은 이름으로 붙는다.
    상위 이름은 **자기 값이 있는 칸에만** 붙인다 — 가로 병합 범위 전체에 뿌리면 한 이름이
    여러 컬럼에 중복된다.
    """
    ncols = max((len(grid[r]) for r in range(top, bottom + 1)), default=0)
    tiers = list(range(top, bottom + 1))
    # 상위 단은 가로 병합 범위로 펼쳐 둔다(하위 칸의 부모를 찾기 위해서만 쓴다)
    spread = []
    for r in tiers:
        row = [grid[r][c] if c < len(grid[r]) else None for c in range(ncols)]
        for r0, c0, r1, c1 in merged:
            if r0 == r + 1 and c1 > c0:
                value = grid[r0 - 1][c0 - 1] if c0 - 1 < len(grid[r0 - 1]) else None
                for c in range(c0 - 1, min(c1, ncols)):
                    if _is_blank(row[c]):
                        row[c] = value
        spread.append(row)

    names = []
    for c in range(ncols):
        leaf = next((i for i in range(len(tiers) - 1, -1, -1)
                     if c < len(grid[tiers[i]]) and not _is_blank(grid[tiers[i]][c])), None)
        if leaf is None:
            names.append("")
            continue
        parts = [str(spread[i][c]).strip() for i in range(leaf) if not _is_blank(spread[i][c])]
        parts.append(str(grid[tiers[leaf]][c]).strip())
        out: list[str] = []
        for part in parts:                 # 세로 병합이면 같은 말이 위아래로 반복된다
            if not out or out[-1] != part:
                out.append(part)
        names.append(" ".join(out))
    while names and names[-1] == "":
        names.pop()
    return names


def _split_tables(grid: list[list], merged: list[tuple[int, int, int, int]]
                  ) -> list[tuple[int, list[str], list[list], list[int], list[int]]]:
    """시트를 표 단위로 나눈다. 표당 `(헤더 인덱스, 헤더, 행, 원본 행번호, 제외 행번호)`.

    대개 표는 하나지만, 실제 양식은 한 시트에 표를 쌓아 두기도 한다(§4의 분기별 실적
    양식은 재해유형×월 피벗 아래에 아차사고 상세 목록이 또 있다).
    """
    out = []
    start = 0
    while start < len(grid):
        sub = grid[start:]
        # 계층 헤더는 여러 행이다. 병합으로 범위를 잡고 한 줄로 합친 뒤, 데이터는 그 아래부터.
        sub_merged = [(r0 - start, c0, r1 - start, c1) for r0, c0, r1, c1 in merged
                      if r1 - start >= 1]
        hi = _find_header(sub, sub_merged)
        if hi < 0:
            break
        top, bottom = _header_block(sub, hi, sub_merged)
        headers = _compose_headers(sub, top, bottom, sub_merged)
        hi = top
        rows: list[list] = []
        row_numbers: list[int] = []
        dropped: list[int] = []
        blanks: list[int] = []
        blank_run = 0
        j = bottom + 1
        while j < len(sub):
            row = list(sub[j])[:len(headers)] + [None] * max(0, len(headers) - len(sub[j]))
            if all(_is_blank(v) for v in row):
                blank_run += 1
                if blank_run > BLANK_ROW_TOLERANCE:
                    break
                j += 1
                continue
            blank_run = 0
            if _is_summary_or_note(row):
                dropped.append(start + j + 1)
                j += 1
                continue
            if _is_template_row(row, headers):
                blanks.append(start + j + 1)
                j += 1
                continue
            if rows and _starts_new_table(row, rows):
                break                       # 여기부터는 다음 표다. j를 넘기지 않는다
            rows.append(row)
            row_numbers.append(start + j + 1)
            j += 1
        # 표 뒤에 미리 병합해 둔 빈 행이 길게 이어지고 그 끝에 합계·각주가 붙는 양식이 있다
        # (실측: 용인 회신본은 5행 데이터, 62행 합계). 빈 구간을 건너 처음 만나는 행이
        # 합계·각주면 이 표의 것으로 본다 — 데이터나 다음 표의 헤더면 손대지 않는다.
        k = j
        while k < len(sub) and all(_is_blank(v) for v in sub[k]):
            k += 1
        while k < len(sub):
            row = list(sub[k])[:len(headers)] + [None] * max(0, len(headers) - len(sub[k]))
            if not _is_summary_or_note(row):
                break
            dropped.append(start + k + 1)
            j = k + 1                       # 소비한 행은 다음 표 탐색에서 다시 보지 않는다
            k += 1
        out.append((start + hi, headers, rows, row_numbers, dropped, blanks))
        start += max(j, hi + 1)             # 최소 한 행은 전진해 무한 루프를 막는다
    return out


def _sheet_selected(title: str, selected: set[str]) -> bool:
    """이 시트를 읽어야 하는지. 한 시트가 표 여러 개로 나뉘면 이름이 '시트명 (2)'가 되므로
    그중 하나라도 골라져 있으면 시트를 읽어야 한다."""
    return title in selected or any(s.startswith(f"{title} (") for s in selected)


def _starts_new_table(row: list, rows: list[list]) -> bool:
    """이 행이 **다음 표의 헤더**인지. 실제 양식은 한 시트에 표를 여러 개 쌓아 둔다.

    빈 행 한 줄만 두고 다음 표가 시작되면 `BLANK_ROW_TOLERANCE`로는 끊기지 않아, 둘째
    표의 헤더와 데이터가 첫 표의 행으로 섞여 들어간다(실측: 그 시트 이슈 132건 중 132건이
    이것 때문이었다).

    판정은 지금까지 읽은 행에서 만든 타입 프로파일을 쓴다 — 숫자로 채워져 온 컬럼에
    글자가 들어오면서 그 행이 온통 글자면 새 표의 헤더다. '전 컬럼이 텍스트인 표'의
    데이터 행은 프로파일이 깨지지 않으므로 걸리지 않는다.
    """
    filled = [v for v in row if not _is_blank(v)]
    if len(filled) < 2 or not all(isinstance(v, str) for v in filled):
        return False
    for c in range(len(row)):
        if _is_blank(row[c]) or not isinstance(row[c], str):
            continue
        seen = [r[c] for r in rows if c < len(r) and not _is_blank(r[c])]
        if len(seen) < 2:
            continue
        numeric = sum(1 for v in seen if not isinstance(v, str))
        if numeric >= len(seen) * 0.8:
            return True
    return False


def _read_images(ws) -> list[dict]:
    """이미지 anchor 전체 점유 범위(from~to)와 메타데이터를 파싱한다 (§2, §5.3)."""
    out = []
    for im in getattr(ws, "_images", []):
        anchor = im.anchor
        frm = getattr(anchor, "_from", None)
        if frm is None:
            continue
        to = getattr(anchor, "to", None)
        try:
            data = im._data()
        except Exception:
            data = b""
        width = height = 0
        try:
            from PIL import Image as PILImage
            with PILImage.open(io.BytesIO(data)) as pim:
                width, height = pim.size
        except Exception:
            width, height = int(im.width or 0), int(im.height or 0)
        # 표시 크기는 anchor의 ext에만 있다(openpyxl의 im.width/height는 원본 픽셀).
        ext = getattr(anchor, "ext", None)
        display = (EMU_to_pixels(ext.cx), EMU_to_pixels(ext.cy)) if ext else None
        out.append({
            "from": (frm.row + 1, frm.col + 1),                       # (row, col) 1-base
            "to": (to.row + 1, to.col + 1) if to else (frm.row + 1, frm.col + 1),
            "single_cell_anchor": to is None,
            "marker_from": (frm.col, frm.row, frm.colOff, frm.rowOff),  # 0-base + EMU 미세 오프셋
            "marker_to": (to.col, to.row, to.colOff, to.rowOff) if to else None,
            "edit_as": getattr(anchor, "editAs", None),
            "display": display,
            "ext": (getattr(im, "format", "") or "").lower(),
            "bytes": len(data),
            "data": data,
            "sha256": hashlib.sha256(data).hexdigest() if data else "",
            "resolution": (width, height),
        })
    return out


PIVOT_SHEET_NAME = "총괄표(전개)"


def _flatten_pivot(sheet: SheetData) -> SheetData:
    """월×지표 피벗표를 **한 줄로 눕힌다** — `{행라벨} {컬럼명}` 컬럼에 값 하나씩.

    총괄표는 두 가지 이유로 종전 방식이 성립하지 않았다. ① 시트명이 파일마다 다르므로
    (`강남지사 총괄표`) 모드 B가 파일 수만큼 시트를 갈라 놓는다. ② 안쪽이 월×지표
    피벗이라 세로로 쌓을 축이 없다 — 같은 `1월` 행이 파일마다 있고, 쌓으면 지사 구분이
    행 라벨에 묻힌다.

    한 줄로 눕히면 지사별 1행이 되어 **기존 모드가 그대로 먹는다**(모드 B로 쌓으면
    지사 12행, `부서` 컬럼으로 구분). 시트명을 한 이름으로 모으는 것도 그래서다.

    행 라벨은 **이름이 있는 첫 컬럼**에서 읽는다(`구분` 아래 `1월`·`2월`…). 0번 컬럼으로
    단정하면 안 된다 — 실제 총괄표는 B열부터 시작해서 A열 헤더가 비어 있다(실측에서 걸렸다.
    `_is_summary_or_note`가 `row[0]`만 보다 합계 행을 놓쳤던 것과 같은 종류).
    이미지는 버린다: 눕힌 뒤에는 어느 칸에 붙어 있었는지 말할 수 없다.
    """
    label_idx = next((i for i, h in enumerate(sheet.headers) if h), 0)
    headers: list[str] = []
    values: list = []
    seen: dict[str, int] = {}
    for row in sheet.rows:
        cell = row[label_idx] if label_idx < len(row) else None
        label = "" if _is_blank(cell) else str(cell).strip()
        for idx, header in enumerate(sheet.headers):
            if idx <= label_idx or not header:
                continue
            name = f"{label} {header}".strip()
            # 라벨이 같은 행이 두 번 나오면 이름이 겹친다. 조용히 덮지 않고 번호를 붙인다
            seen[name] = seen.get(name, 0) + 1
            if seen[name] > 1:
                name = f"{name} ({seen[name]})"
            headers.append(name)
            values.append(row[idx] if idx < len(row) else None)
    row_number = sheet.row_numbers[0] if sheet.row_numbers else sheet.header_row + 1
    return SheetData(PIVOT_SHEET_NAME, sheet.header_row, headers, [values] if headers else [],
                     [row_number] if headers else [])


def read_file(path: Path, include_hidden: bool = False,
              include_sheets: set[str] | None = None,
              pivot_sheets: set[str] | None = None) -> UploadedFile:
    """xlsx 한 개를 읽는다.

    `include_sheets`를 주면 그 시트만 읽는다. 실제 양식에는 취합 대상이 아닌 시트가
    섞여 있다 — 작성 가이드, 비워 둔 대장, 지사별로 이름이 다른 총괄표. 어느 시트가
    데이터인지는 담당자가 알고 있으므로 이름 규칙으로 맞히지 않고 골라 받는다.

    `pivot_sheets`에 든 표는 월×지표 피벗으로 보고 한 줄로 눕힌다(`_flatten_pivot`).
    피벗인지도 이름으로 맞히지 않는다 — 담당자가 고른다.
    """
    uf = UploadedFile(path=path, dept=path.stem)
    if path.suffix.lower() != ".xlsx":
        uf.readable = False
        uf.issues.append(Issue(path.name, "(해당없음)", "(해당없음)", "정합성", "1단계", ERROR,
                               f"지원하지 않는 파일 형식({path.suffix}). .xlsx만 업로드할 수 있습니다. "
                               "해결방법: 엑셀에서 '다른 이름으로 저장' → 'Excel 통합 문서(*.xlsx)'로 변경 후 재시도."))
        return uf
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:                       # 손상/암호 보호 (§2.1)
        uf.readable = False
        uf.issues.append(Issue(path.name, "(해당없음)", "(해당없음)", "정합성", "1단계", ERROR,
                               f"파일을 열 수 없습니다(손상 또는 암호 보호 가능): {exc}. "
                               "해결방법: 엑셀에서 정상적으로 열리는지, 암호가 걸려있지 않은지 확인 후 재시도."))
        return uf

    skipped: list[Issue] = []           # 표가 아닌 시트 — 유효한 시트가 하나도 없을 때만 오류로 올린다
    for ws in wb.worksheets:
        if ws.sheet_state != "visible" and not include_hidden:
            continue
        # 고르지 않은 시트는 사유도 남기지 않는다 — 사용자가 뺀 것이라 조치할 것이 없다.
        # 시트를 표 여러 개로 나눠 읽는 경우(_split) 그 이름들도 함께 본다.
        if include_sheets is not None and not _sheet_selected(ws.title, include_sheets):
            continue
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        images = _read_images(ws)
        if not any(not _is_blank(v) for row in grid for v in row) and not images:
            skipped.append(Issue(path.name, ws.title, "(해당없음)", "정합성", "1단계", ERROR,
                                 "시트에 데이터가 없습니다. 해결방법: 빈 시트를 삭제하거나 데이터를 입력 후 재시도."))
            continue
        merged = [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]
        tables = _split_tables(grid, merged)
        if not tables:
            skipped.append(Issue(path.name, ws.title, "(해당없음)", "정합성", "1단계", ERROR,
                                 "헤더(열 제목) 행을 찾을 수 없습니다. 해결방법: 제목·로고 행과 데이터 사이에 "
                                 "명확한 열 제목 행이 있는지 확인해주세요."))
            uf.note_sheets.append(ws.title)
            lines = [str(v).strip() for row in grid for v in row
                     if isinstance(v, str) and v.strip()]
            uf.notes_text = (uf.notes_text + "\n" + "\n".join(lines))[:NOTES_MAX_CHARS].strip()
            continue

        added = 0
        for part, (hi, headers, rows, row_numbers, dropped, blanks) in enumerate(tables, 1):
            # 표가 둘 이상이면 뒤 표에 이름을 붙인다. 시작 행이 파일마다 같으므로
            # 모드 B에서 같은 표끼리 묶인다(실측: 12개 회신본 모두 35행에서 둘째 표 시작).
            name = ws.title if len(tables) == 1 else f"{ws.title} ({part})"
            # 어떤 행이 데이터인지는 원본 기준으로 먼저 정하고, 병합 값은 그 뒤에 채운다.
            # 순서가 바뀌면 미리 병합해 둔 빈 템플릿 행까지 값이 생겨 데이터로 되살아난다.
            vmerges = _fill_merged(rows, row_numbers, merged)
            total = _total_row_of(grid, dropped)
            if dropped:
                # 합계·각주 행은 실제 양식의 정상 구조라 결함이 아니다 — 경고로 두면 합계 행을
                # 가진 모든 회신본이 '이상'으로 뜬다(실측: 용인). 등급 없는 안내로 남긴다
                uf.issues.append(Issue(path.name, name, f"{dropped[0]}행" if len(dropped) == 1
                                       else f"{dropped[0]}~{dropped[-1]}행",
                                       "정합성", "1단계", OK,
                                       f"합계·각주로 보이는 {len(dropped)}개 행을 취합에서 제외했습니다"
                                       f"(행 {', '.join(map(str, dropped))}). 데이터 행이라면 "
                                       "첫 칸의 '합계' 같은 표기를 지워주세요."))
            if blanks:
                uf.issues.append(Issue(path.name, name, f"{blanks[0]}~{blanks[-1]}행",
                                       "정합성", "1단계", WARN,
                                       f"순번만 있고 내용이 비어 있는 {len(blanks)}개 행을 취합에서 "
                                       f"제외했습니다(행 {blanks[0]}~{blanks[-1]}). 작성하려던 "
                                       "행이라면 내용을 채워주세요."))
            if not rows:
                if len(tables) == 1:
                    skipped.append(Issue(path.name, name, "(해당없음)", "누락", "1단계", ERROR,
                                         "헤더는 있으나 입력된 데이터가 없습니다. "
                                         "해결방법: 데이터를 입력 후 재시도."))
                continue
            if include_sheets is not None and name not in include_sheets:
                continue
            # 이미지는 자기 표의 행 범위에 있는 것만 딸려간다. 시트 하나의 이미지를
            # 전 표에 다 붙이면 남의 표 사진까지 그 표의 이슈로 잡힌다.
            lo, hla = row_numbers[0], row_numbers[-1]
            mine = [im for im in images if lo <= im["from"][0] <= hla] if len(tables) > 1 else images
            data = SheetData(name, hi + 1, headers, rows, row_numbers, mine, merged=merged,
                             vmerges=vmerges,
                             total_label=total[0], total_label_col=total[1], total_cols=total[2])
            if pivot_sheets and name in pivot_sheets:
                # 피벗으로 고른 표는 한 줄로 눕히고 한 이름으로 모은다(§9 모드가 그대로 먹는다)
                flat = _flatten_pivot(data)
                if not flat.headers:
                    uf.issues.append(Issue(path.name, name, "(해당없음)", "정합성", "1단계", WARN,
                                           "피벗형으로 지정했지만 눕힐 값이 없어 건너뛰었습니다. "
                                           "표에 데이터 행이 있는지 확인해주세요."))
                    continue
                if len(data.images) > 0:
                    uf.issues.append(Issue(path.name, name, "(해당없음)", "정합성", "1단계", WARN,
                                           f"피벗형으로 눕히면서 이 표의 사진 {len(data.images)}장을 "
                                           "결과에서 제외했습니다(눕힌 뒤에는 붙어 있던 칸을 "
                                           "말할 수 없습니다)."))
                data = flat
            uf.sheets.append(data)
            added += 1
        if added and len(tables) > 1:
            uf.issues.append(Issue(path.name, ws.title, "(해당없음)", "정합성", "1단계", WARN,
                                   f"이 시트에서 표 {len(tables)}개를 찾아 따로 읽었습니다"
                                   f"(시작 행 {', '.join(str(t[0] + 1) for t in tables)}). "
                                   "한 표여야 한다면 표 사이의 빈 행을 지워주세요."))

    # 취합할 시트가 하나라도 있으면, 표가 아닌 시트는 파일을 막을 사유가 아니다.
    # 실제 업무 양식에는 '작성 주의사항'처럼 안내문만 있는 시트가 흔하다 —
    # 이걸 오류로 보면 파일 전체가 '이상'이 되어 취합에서 기본 제외된다(§7). 경고로 두면
    # 멀쩡한 파일이 화면에 '이상'으로 뜬다(실측: 용인 회신본) — 등급 없는 안내로 남긴다.
    if uf.sheets:
        for issue in skipped:
            uf.issues.append(Issue(issue.file, issue.sheet, issue.cell, issue.kind, issue.stage,
                                   OK, "표 구조가 아니어서 취합에서 제외했습니다"
                                         "(작성 안내·빈 시트 등). 취합 대상 시트라면 "
                                         "제목 행과 데이터 사이에 열 제목 행이 있는지 확인해주세요."))
    else:
        uf.issues.extend(skipped)
    return uf


def _fill_merged(rows: list[list], row_numbers: list[int],
                 merged: list[tuple[int, int, int, int]]) -> list[tuple[int, int, int]]:
    """병합 셀의 값을 그 범위에 속한 데이터 행에 채우고, 채운 구간을 돌려준다.

    엑셀에서 세로로 병합된 칸은 사람 눈에는 구간 전체의 값이지만, 파일에는
    좌상단 한 칸만 값이 있고 나머지는 비어 있다. 채우지 않으면 실제 업무 양식의
    '지사'·'개소'처럼 여러 행에 걸친 칸이 전부 필수값 누락으로 잡힌다.

    다만 채운 값은 **검증용**이다. 그대로 결과에 쓰면 '개소'처럼 구간 하나를 세는
    컬럼이 행마다 반복되어 합계가 부풀어 오른다(실측: 판교 2개소가 2+2, 수원 3개소가
    3+3+3 → 결과 합 44 vs 원본 합 34). 그래서 채운 구간을 돌려주고 결과에서는
    `_restore_vmerges`가 다시 병합으로 되돌린다.

    이미 데이터로 확정된 행만 대상이다 — 양식이 미리 병합해 둔 빈 행까지 채우면
    값이 생겨 데이터 행으로 되살아난다(실제 양식은 5~60행을 미리 병합해 둔다).
    """
    if not rows:
        return []
    runs: list[tuple[int, int, int]] = []
    index_of = {number: i for i, number in enumerate(row_numbers)}
    for min_row, min_col, max_row, max_col in merged:
        if min_row == max_row and min_col == max_col:
            continue
        source = index_of.get(min_row)
        if source is None:
            continue                       # 병합 시작 행이 데이터가 아니면(제목 등) 건너뛴다
        targets = [index_of[n] for n in range(min_row, max_row + 1) if n in index_of]
        width = len(rows[source])
        for c in range(min_col - 1, min(max_col, width)):
            value = rows[source][c]
            if _is_blank(value):
                continue
            for target in targets:
                if _is_blank(rows[target][c]):
                    rows[target][c] = value
            # 중간에 뺀 행(합계·빈 행)이 끼어 끊긴 구간은 되돌리지 않는다 — 병합하면
            # 그 병합에 없던 행까지 묶인다. 값은 채운 채 두고 검증만 받는다
            if max_row > min_row and len(targets) > 1 and                     targets[-1] - targets[0] + 1 == len(targets):
                runs.append((targets[0], targets[-1], c))
    return runs


def _restore_vmerges(ws, sheet: "SheetData", first: int, col_offset: int,
                     skip_col: int = 0) -> None:
    """원본에서 세로 병합이던 칸을 결과에도 병합으로 되돌린다.

    openpyxl은 병합할 때 좌상단 외의 값을 지우므로, 반복해 채워 둔 값이 한 번만
    남고 합계 SUM도 한 번만 센다. `first`는 이 파일 블록의 첫 출력 행,
    `col_offset`은 앞에 끼운 '구분'·'부서' 컬럼 수다.
    """
    for lo, hi, col in sheet.vmerges:
        column = col + col_offset + 1
        if column == skip_col:
            continue           # 부서 구분 컬럼은 파일 블록 단위로 따로 병합한다(이중 병합 금지)
        ws.merge_cells(start_row=first + lo, start_column=column,
                       end_row=first + hi, end_column=column)


# ── 3. 데이터 유형 분류 ───────────────────────────────────────────────────────
def _is_number(value) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return _to_number(value) is not None


def _to_number(value):
    """콤마·통화기호·공백을 제거하고 숫자로 변환. 실패 시 None (§5.2 경고 조건)."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        return None
    cleaned = re.sub(r"[,\s₩$€%]", "", value.strip())
    try:
        return float(cleaned) if cleaned else None
    except ValueError:
        return None


def classify_columns(sheet: SheetData, rules: dict) -> list[str]:
    """컬럼별 유형: numeric / code_list / pattern / freeform / image (§3)."""
    image_cols = {c for im in sheet.images for c in range(im["from"][1], im["to"][1] + 1)}
    types = []
    for idx, header in enumerate(sheet.headers):
        rule = rules.get(header, {})
        if rule.get("type"):
            types.append(rule["type"])
            continue
        if (idx + 1) in image_cols:
            types.append("image")
            continue
        low = header.lower()
        hint = None
        if any(k in low for k in NUM_KEYWORDS):
            hint = "numeric"
        elif any(k in low for k in CODE_KEYWORDS):
            hint = "code_list"
        sample = [r[idx] for r in sheet.rows[:TYPE_SAMPLE_ROWS] if not _is_blank(r[idx])]
        if sample:
            # 값 타입 분포가 최종 기준 (키워드 힌트보다 우선, §3).
            # 다수결(>50%)로 판정한다 — 오류값이 일부 섞인 숫자 컬럼도 numeric으로 봐야
            # 그 오류값을 정합성 위반으로 잡아낼 수 있다.
            numeric_ratio = sum(_is_number(v) for v in sample) / len(sample)
            if numeric_ratio > 0.5:
                types.append("numeric")
                continue
            if hint == "numeric":
                hint = None
        if hint == "code_list" and rule.get("master_list"):
            types.append("code_list")
        elif rule.get("pattern_regex") or "이메일" in low or "email" in low:
            types.append("pattern")
        elif hint == "code_list":
            types.append("code_list")
        else:
            types.append("freeform")          # §3 분류 불확실 시 기본값(가장 관대)
    return types


# ── 4.1 / 5. 1단계 규칙기반 검증 ──────────────────────────────────────────────
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+\.[\w.]+$")


def _validate_text(value: str, ctype: str, rule: dict):
    """(grade, reason, corrected) 반환. grade=None이면 정상 (§5.1)."""
    if ctype == "code_list":
        master = rule.get("master_list") or []
        if not master:
            return None, "", None
        if value in master:
            return None, "", None
        norm = {m.strip().casefold(): m for m in master}
        hit = norm.get(value.strip().casefold())
        if hit:
            return WARN, f"코드표와 공백·대소문자 차이만 존재(자동 교정 가능): '{value}' → '{hit}'", hit
        return ERROR, f"코드표 마스터 목록에 없는 값: '{value}'", None
    if ctype == "pattern":
        pattern = rule.get("pattern_regex")
        regex = re.compile(pattern) if pattern else EMAIL_RE
        if regex.match(value):
            return None, "", None
        squeezed = re.sub(r"\s+", "", value)
        if regex.match(squeezed):
            return WARN, f"구분자·공백 차이(자동 교정 가능): '{value}' → '{squeezed}'", squeezed
        return ERROR, f"허용 형식에 맞지 않는 값: '{value}'", None
    # freeform
    for banned in rule.get("banned_words") or []:
        if banned in value:
            return ERROR, f"금칙어 포함: '{banned}'", None
    limit = int(rule.get("max_length") or FREEFORM_MAX_LENGTH)
    if len(value) > limit:
        return WARN, f"길이 초과({len(value)}자 > {limit}자, 자동 절삭 가능)", value[:limit]
    return None, "", None


def _validate_number(raw, rule: dict):
    number = _to_number(raw)
    if number is None:
        return ERROR, f"숫자로 해석할 수 없는 값: '{raw}'", None
    lo, hi = rule.get("min"), rule.get("max")
    if lo is not None and number < lo:
        return ERROR, f"허용 범위 미달({number} < {lo})", None
    if hi is not None and number > hi:
        return ERROR, f"허용 범위 초과({number} > {hi})", None
    places = rule.get("decimal_places")
    if places is not None:
        rounded = round(number, int(places))
        if abs(rounded - number) > 1e-12:
            return WARN, f"소수점 자릿수 규칙 위반(자동 반올림 가능): {number} → {rounded}", rounded
        number = rounded
    if isinstance(raw, str):
        return WARN, f"콤마·통화기호 제거 후 정상(자동 교정 가능): '{raw}' → {number}", number
    return None, "", None


def _pick_key_column(headers: list[str], rows: list[list] | None = None) -> int:
    """중복·충돌 판정의 기준 컬럼을 고른다 (F2-13의 `자동`).

    ① 헤더에 사번·코드 같은 키워드가 있으면 그 컬럼
    ② 없으면 값이 행마다 고유한 첫 컬럼 — 실제 업무 양식은 첫 컬럼이 지사·부서명처럼
       반복되는 경우가 흔해서, 첫 컬럼을 그냥 키로 삼으면 전 행이 충돌로 오탐된다
    ③ 그래도 못 찾으면 첫 컬럼 (사용자가 화면에서 지정하는 것이 최선)
    """
    for idx, header in enumerate(headers):
        if any(k in header.lower() for k in KEY_KEYWORDS):
            return idx
    if rows:
        for idx in range(len(headers)):
            values = [str(r[idx]).strip() for r in rows
                      if idx < len(r) and not _is_blank(r[idx])]
            if len(values) == len(rows) and len(set(values)) == len(values):
                return idx
    return 0


def _short_name(name: str, limit: int = 24) -> str:
    """안내문에 넣는 컬럼명. 계층 헤더가 뭉친 컬럼명은 수백 자라 안내문을 덮는다."""
    flat = " ".join(str(name).split())
    return flat if len(flat) <= limit else flat[:limit] + "…"


# ── 계획-실적 페어 검사 (9차) ─────────────────────────────────────────────────
# 월별 점검표(계획/실적 2행 페어)의 진짜 검토 논리: 조사 기간 안에서 계획(●)이 있는
# 달은 실적도 있어야 한다. 기간은 파일명·시트명에서 추정하되(사용자 결정: 자동 추정),
# 화면·CLI에서 직접 지정하면 그 값을 쓴다. 추정도 지정도 없으면 검사하지 않는다 —
# 기간을 모르는 채 전체 12개월을 검사하면 미래 달이 전부 미이행으로 오탐된다.
PERIOD_QUARTERS = {"1분기": (1, 3), "2분기": (4, 6), "3분기": (7, 9), "4분기": (10, 12),
                   "상반기": (1, 6), "하반기": (7, 12)}
PAIR_PLAN, PAIR_ACTUAL = "계획", "실적"
_MONTH_RE = re.compile(r"(\d{1,2})\s*월")
_RANGE_RE = re.compile(r"(\d{1,2})\s*월?\s*[~\-–]\s*(\d{1,2})\s*월")


def _infer_period(text: str) -> tuple[int, int] | None:
    """'1분기'·'상반기'·'1~3월'·'3월' 표기에서 (시작월, 끝월)을 뽑는다. 없으면 None."""
    for token, span in PERIOD_QUARTERS.items():
        if token in text:
            return span
    m = _RANGE_RE.search(text)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        if 1 <= lo <= hi <= 12:
            return lo, hi
    # 연도('2026년')의 숫자를 월로 오인하지 않도록 단독 월은 마지막에, 좁게 본다
    months = [int(g) for g in _MONTH_RE.findall(text) if 1 <= int(g) <= 12]
    if len(months) == 1:
        return months[0], months[0]
    return None


def _month_columns(headers: list[str]) -> dict[int, int]:
    """{컬럼 인덱스: 월}. '2025년 1월'처럼 합성된 계층 헤더에서도 월을 찾는다."""
    out = {}
    for idx, header in enumerate(headers):
        m = _MONTH_RE.search(header or "")
        if m and 1 <= int(m.group(1)) <= 12:
            out[idx] = int(m.group(1))
    return out


def _review_plan_actual(uf: UploadedFile, sheet: SheetData,
                        period: tuple[int, int]) -> None:
    """계획 행에 값이 있는 조사 기간 내 달은 짝 실적 행에도 값이 있어야 한다.

    구분 컬럼은 이름으로 맞히지 않는다 — 값이 계획/실적 두 가지뿐인 컬럼을 찾는다.
    짝은 계획 행 바로 다음에 오는 실적 행이다(이 계열 양식의 고정 표기법).
    미이행은 파일 작성 오류가 아니라 업무 사실일 수 있으므로 **경고**로 보고만 하고
    파일을 반려(오류) 대상으로 만들지 않는다(사용자 결정, 9차).
    """
    months = _month_columns(sheet.headers)
    if not months:
        return
    marker = None
    for idx in range(len(sheet.headers)):
        values = {str(r[idx]).strip() for r in sheet.rows
                  if idx < len(r) and not _is_blank(r[idx])}
        if values == {PAIR_PLAN, PAIR_ACTUAL}:
            marker = idx
            break
    if marker is None:
        return
    uf.pair_period = period
    lo, hi = period
    # 부서별 회신 파일에는 다른 부서 몫의 행이 비어 있는 것이 정상이다(부서 1개 =
    # 파일 1개 전제, §9). 주관부서류 컬럼이 있고 회신 부서와 맞는 행이 있으면 그
    # 행만 검사한다. 하나도 안 맞으면 전체를 검사한다 — 조용히 다 건너뛰는 것보다
    # 경고가 몇 건 더 뜨는 쪽이 안전하다(등급이 경고라 반려로 이어지지 않는다).
    dept_idx = next((i for i, h in enumerate(sheet.headers)
                     if h and ("주관부서" in h or "담당부서" in h or h.strip() == "부서")),
                    None)
    mine: set[int] | None = None
    if dept_idx is not None and uf.dept:
        matched = {i for i, row in enumerate(sheet.rows)
                   if not _is_blank(row[dept_idx])
                   and str(row[dept_idx]).strip() in uf.dept}
        if matched:
            mine = matched
    for i, row in enumerate(sheet.rows):
        if str(row[marker]).strip() != PAIR_PLAN:
            continue
        if mine is not None and i not in mine:
            continue
        nxt = sheet.rows[i + 1] if i + 1 < len(sheet.rows) else None
        if nxt is None or str(nxt[marker]).strip() != PAIR_ACTUAL:
            continue                     # 짝이 깨진 행은 판정하지 않는다(추측 금지)
        for idx, month in months.items():
            if not (lo <= month <= hi):
                continue
            if _is_blank(row[idx]) or not _is_blank(nxt[idx]):
                continue
            cell = f"{get_column_letter(idx + 1)}{sheet.row_numbers[i + 1]}"
            uf.issues.append(Issue(
                uf.name, sheet.name, cell, "계획실적", "1단계", WARN,
                f"계획 미이행 확인 필요 — {month}월 계획이 있으나 실적이 비어 있습니다"
                f"(검사 기간 {lo}~{hi}월)", sheet.headers[idx]))


def review_stage1(uf: UploadedFile, rules: dict,
                  period: tuple[int, int] | None = None) -> None:
    """파일 내 모든 셀을 끝까지 스캔한다(첫 위반에서 중단하지 않음, §4.1).

    `period`는 계획-실적 페어 검사의 조사 기간(월). 지정이 없으면 파일명에서
    추정하고, 그래도 없으면 페어 검사만 건너뛴다(다른 검사는 그대로 돈다).
    """
    if period is None:
        period = _infer_period(uf.name)
    uf.pair_period = None                # 재검토 시 지난 판이 남지 않게 초기화
    for sheet in uf.sheets:
        sheet.col_types = classify_columns(sheet, rules)
        default_miss: list[Issue] = []   # 미선언 컬럼(기본 가정)에서 나온 누락만 모은다
        col_folds: list[Issue] = []      # 컬럼 전체 미기입으로 접은 안내 — 폭주 감지가 다시 접는다
        unruled_checked = 0              # 그 컬럼들에서 실제로 검사한 칸 수
        for idx, header in enumerate(sheet.headers):
            ctype = sheet.col_types[idx]
            if ctype == "image" or not header:
                continue
            rule = rules.get(header, {})
            required = rule.get("required", True)
            # 사용자가/양식이 선언한 기준에서 나온 누락은 절대 접지 않는다 —
            # 접기는 '기본 가정이 틀렸다'는 신호에만 반응해야 한다
            declared = header in rules
            if not declared:
                unruled_checked += len(sheet.rows)
            letter = get_column_letter(idx + 1)
            col_miss: list[Issue] = []       # 이 컬럼의 누락 — 전 행이 비면 1건으로 접는다
            for i, row in enumerate(sheet.rows):
                cell = f"{letter}{sheet.row_numbers[i]}"
                raw = row[idx]
                # 누락 → 정합성 → 중복/충돌 순서. 누락이면 이후 검사 생략 (§4.1)
                if _is_blank(raw):
                    # 작성기준이 선택 입력으로 정한 칸은 비어 있는 것이 정상이므로
                    # 보고하지 않는다. 담당자가 조치할 것이 없고, 비고처럼 대개
                    # 비워두는 컬럼에서는 행 수만큼 알림이 쌓여 리포트를 덮는다.
                    if required:
                        issue = Issue(uf.name, sheet.name, cell, "누락", "1단계", ERROR,
                                      "필수값 누락", header)
                        uf.issues.append(issue)
                        col_miss.append(issue)
                        if not declared:
                            default_miss.append(issue)
                    continue
                if ctype == "numeric":
                    grade, reason, corrected = _validate_number(raw, rule)
                else:
                    grade, reason, corrected = _validate_text(str(raw), ctype, rule)
                if grade:
                    uf.issues.append(Issue(uf.name, sheet.name, cell, "정합성", "1단계", grade,
                                           reason, header))
                    if grade == WARN and corrected is not None:
                        uf.fixes.append(Fix(uf.name, sheet.name, cell, str(raw), str(corrected), reason))

            # 컬럼 전체가 비었으면 '깜빡한 칸'이 아니라 '안 쓴 컬럼'이다 — 칸마다 오류를 내면
            # 파일이 통째로 취합에서 빠진다(실측: 중앙지사 회신본이 지사·개소 8칸으로 제외).
            # 칸 오류를 걷어내고 경고 1건으로 접는다. 소속 컬럼(지사·부서)이면 회신 부서명으로
            # 채운다 — 파일 하나가 부서 하나라는 전제(§9)에서 그 값은 이미 알고 있다.
            # 행이 하나뿐이면 '안 쓴 컬럼'과 '빠뜨린 칸'을 가를 수 없어 접지 않는다 — 소속
            # 컬럼은 예외다. 적힌 값이 전부 그 파일의 부서면(조직 꼬리 무시) 빈 칸의 값은
            # 정해져 있으므로, 전부 비었든 일부만 비었든(실측: 분당 회신본은 병합 범위가 한 행
            # 짧아 마지막 행만 비었다) 부서명으로 채운다.
            fill_dept = _is_dept_header(header) and bool(uf.dept) and bool(col_miss) and all(
                _org_key(row[idx]) == _org_key(uf.dept)
                for row in sheet.rows if not _is_blank(row[idx]))
            if not declared and col_miss and (
                    fill_dept or (len(col_miss) == len(sheet.rows) and len(sheet.rows) >= 2)):
                folded_ids = {id(i) for i in col_miss}
                # default_miss에는 남긴다 — 폭주 감지는 '빈 칸이 몇 개였나'를 봐야 한다
                uf.issues[:] = [i for i in uf.issues if id(i) not in folded_ids]
                blank_rows = [n for n, row in zip(sheet.row_numbers, sheet.rows) if _is_blank(row[idx])]
                span = f"{letter}{blank_rows[0]}:{letter}{blank_rows[-1]}" \
                    if len(blank_rows) > 1 else f"{letter}{blank_rows[0]}"
                if fill_dept:
                    for number in blank_rows:
                        uf.fixes.append(Fix(uf.name, sheet.name, f"{letter}{number}", "", uf.dept,
                                            "소속 컬럼 빈 칸 → 부서명 자동 채움"))
                    reason = (f"'{header}' 컬럼이 전부 비어 있어 부서명 '{uf.dept}'으로 채웠습니다"
                              f"({len(blank_rows)}칸)" if len(blank_rows) == len(sheet.rows) else
                              f"'{header}' 컬럼의 빈 칸 {len(blank_rows)}개를 부서명 '{uf.dept}'으로 "
                              f"채웠습니다(적힌 값이 전부 이 부서)")
                else:
                    reason = (f"'{header}' 컬럼이 전부 비어 있습니다({len(sheet.rows)}칸). 칸 누락이 "
                              "아니라 미기입 컬럼으로 보고 1건으로 접었습니다. 비워 두는 컬럼이면 "
                              "선택 입력 컬럼으로 지정해주세요.")
                fold = Issue(uf.name, sheet.name, span, "누락", "1단계", WARN, reason, header)
                uf.issues.append(fold)
                col_folds.append(fold)

        # 기본 가정 폭주 감지 — 미선언 컬럼의 누락이 임계(건수·비율 모두)를 넘으면
        # 개별 보고를 접고 '가정 재검토' 안내 1건으로 바꾼다. 빈 칸이 정상인
        # 점검표(월별 매트릭스 등)에서 수천 건이 쏟아져 진짜 신호를 덮는 것을 막는다.
        # 파일을 정상으로 만들지는 않는다(경고 1건 유지) — 빈 칸의 의미는 구조로
        # 추측할 수 없으므로, 판정은 사용자가 컬럼을 지정한 뒤의 재검토가 한다.
        if len(default_miss) >= DEFAULT_MISS_FLOOR and unruled_checked \
                and len(default_miss) / unruled_checked >= DEFAULT_MISS_RATIO:
            folded = {id(i) for i in default_miss} | {id(i) for i in col_folds}
            uf.issues[:] = [i for i in uf.issues if id(i) not in folded]
            cols: list[str] = []
            for i in default_miss:
                if i.attr and i.attr not in cols:
                    cols.append(i.attr)
            shown = ", ".join(_short_name(c) for c in cols[:5]) \
                + (f" 외 {len(cols) - 5}개" if len(cols) > 5 else "")
            uf.issues.append(Issue(
                uf.name, sheet.name, "(해당없음)", "가정", "1단계", WARN,
                f"필수 가정 재검토 필요 — 작성기준이 지정되지 않은 컬럼 {len(cols)}개에서 "
                f"필수값 누락 {len(default_miss)}건이 나왔습니다(검사한 칸의 "
                f"{len(default_miss) / unruled_checked:.0%}). 빈 칸이 정상인 점검표로 "
                f"보여 개별 보고 대신 이 안내로 접었습니다. 키 컬럼과 선택 입력 컬럼을 "
                f"지정한 뒤 다시 검토해주세요. 대상 컬럼: {shown}"))

        _review_duplicates(uf, sheet, rules)
        _review_images(uf, sheet, rules)
        if period:
            _review_plan_actual(uf, sheet, period)


def _review_duplicates(uf: UploadedFile, sheet: SheetData, rules: dict) -> None:
    """동일 키 그룹핑 → 그룹 내 값 비교. 완전 중복=경고, 값 상이=충돌(오류) (§5.1)."""
    if not sheet.headers:
        return
    declared_idx = next((i for i, h in enumerate(sheet.headers)
                         if rules.get(h, {}).get("key")), None)
    key_idx = declared_idx if declared_idx is not None \
        else _pick_key_column(sheet.headers, sheet.rows)
    letter = get_column_letter(key_idx + 1)
    found: list[Issue] = []
    fixes: list[Fix] = []
    groups: dict[str, list[int]] = {}
    for i, row in enumerate(sheet.rows):
        key = row[key_idx]
        if _is_blank(key):
            continue
        groups.setdefault(str(key).strip(), []).append(i)
    for key, indexes in groups.items():
        if len(indexes) < 2:
            continue
        first = [str(v) for v in sheet.rows[indexes[0]]]
        for i in indexes[1:]:
            cell = f"{letter}{sheet.row_numbers[i]}"
            if [str(v) for v in sheet.rows[i]] == first:
                found.append(Issue(uf.name, sheet.name, cell, "중복", "1단계", WARN,
                                   f"키 '{key}' 완전 중복 행(자동 제거 가능)", sheet.headers[key_idx]))
                fixes.append(Fix(uf.name, sheet.name, cell, f"중복 행(키 {key})", "행 제거", "완전 중복 행 자동 제거"))
            else:
                found.append(Issue(uf.name, sheet.name, cell, "충돌", "1단계", ERROR,
                                   f"키 '{key}'가 같으나 값이 상이한 행 존재", sheet.headers[key_idx]))

    # 키 가정 폭주 감지 — 자동 판정한 키에서 충돌·중복이 임계를 넘으면 개별 보고 대신
    # '키 지정' 안내 1건으로 접는다(필수 가정 접기와 같은 원칙). 병합된 분류 컬럼이
    # 키로 잘못 잡히면 행 수만큼 충돌이 쏟아진다(실측: 월별 점검표에서 71건).
    # 사용자가 지정한 키에서 나온 충돌은 몇 건이든 접지 않는다. 중복 행의 '자동 제거'
    # 제안(fixes)도 함께 걷는다 — 틀린 키 가정으로 행을 지우면 안 된다.
    if declared_idx is None and sheet.rows \
            and len(found) >= DEFAULT_MISS_FLOOR \
            and len(found) / len(sheet.rows) >= DEFAULT_MISS_RATIO:
        uf.issues.append(Issue(
            uf.name, sheet.name, "(해당없음)", "가정", "1단계", WARN,
            f"키 가정 재검토 필요 — 키 컬럼을 지정하지 않아 "
            f"'{_short_name(sheet.headers[key_idx])}' 컬럼을 키로 가정했는데, 충돌·중복이 "
            f"{len(found)}건(데이터 행의 {len(found) / len(sheet.rows):.0%}) 나왔습니다. "
            f"분류처럼 값이 반복되는 컬럼을 키로 잘못 가정했을 가능성이 큽니다. "
            f"개별 보고 대신 이 안내로 접었습니다 — 키 컬럼을 지정한 뒤 다시 검토해주세요.",
            sheet.headers[key_idx]))
        return
    uf.issues.extend(found)
    uf.fixes.extend(fixes)


def _review_images(uf: UploadedFile, sheet: SheetData, rules: dict) -> None:
    """이미지는 1단계만 수행하고 종료 (§5.3). 내용 판단은 하지 않는다."""
    if not sheet.images:
        return
    rule = next((r for h, r in rules.items() if r.get("type") == "image"), {})
    allowed = tuple(rule.get("allowed_extensions") or IMAGE_ALLOWED_EXT)
    max_bytes = float(rule.get("max_file_size_mb") or IMAGE_MAX_SIZE_MB) * 1024 * 1024
    min_res = tuple(rule.get("min_resolution") or IMAGE_MIN_RESOLUTION)
    max_count = int(rule.get("max_count_per_key") or IMAGE_MAX_COUNT_PER_KEY)
    seen_hash: dict[str, str] = {}
    # 명세 §5.3은 "동일 키(동일 행/개체)에 대해 **필드당** 허용 개수"다. 행 단위로 세면
    # 사진 열이 둘인 양식(문제 예시·개선 예시)에서 정상 입력이 전부 위반으로 잡힌다.
    per_key: dict[tuple[int, int], int] = {}

    for im in sheet.images:
        cell = f"{get_column_letter(im['from'][1])}{im['from'][0]}"
        if im["ext"] and im["ext"] not in allowed:
            uf.issues.append(Issue(uf.name, sheet.name, cell, "정합성", "1단계", ERROR,
                                   f"허용되지 않는 이미지 형식: {im['ext']} (허용: {', '.join(allowed)})"))
        if im["bytes"] > max_bytes:
            uf.issues.append(Issue(uf.name, sheet.name, cell, "정합성", "1단계", ERROR,
                                   f"이미지 용량 상한 초과({im['bytes'] / 1024 / 1024:.1f}MB > {max_bytes / 1024 / 1024:.0f}MB)"))
        width, height = im["resolution"]
        if width and (width < min_res[0] or height < min_res[1]):
            uf.issues.append(Issue(uf.name, sheet.name, cell, "정합성", "1단계", ERROR,
                                   f"최소 해상도 미달({width}×{height} < {min_res[0]}×{min_res[1]})"))
        # 소속 범위 정합성: anchor 전체 범위가 병합 영역과 일치하면 정상,
        # 병합 셀이 아니면서 2개 이상 데이터 행에 걸치면 소속 모호 → 경고 (§5.3)
        if im["to"][0] > im["from"][0]:
            spans_merged = any(mr[0] <= im["from"][0] and im["to"][0] <= mr[2]
                               for mr in sheet.merged)
            if not spans_merged:
                uf.issues.append(Issue(uf.name, sheet.name, cell, "소속범위", "1단계", WARN,
                                       f"이미지가 {im['from'][0]}~{im['to'][0]}행에 걸쳐 있고 병합 셀 영역과 "
                                       "일치하지 않아 소속 레코드가 모호합니다(자동 처리 불가, 담당자 확인 권고)"))
        if im["sha256"]:
            if im["sha256"] in seen_hash:
                uf.issues.append(Issue(uf.name, sheet.name, cell, "중복", "1단계", WARN,
                                       f"동일 이미지 파일 재사용({seen_hash[im['sha256']]}와 해시 일치, 담당자 확인 권고)"))
            else:
                seen_hash[im["sha256"]] = cell
        key = (im["from"][0], im["from"][1])
        per_key[key] = per_key.get(key, 0) + 1

    for (row, col), count in per_key.items():
        if count > max_count:
            uf.issues.append(Issue(uf.name, sheet.name, f"{get_column_letter(col)}{row}",
                                   "충돌", "1단계", ERROR,
                                   f"한 칸에 이미지가 {count}장 있습니다(허용 {max_count}장)"))


# ── 4.2 2단계 AI(LLM) 재검증 ─────────────────────────────────────────────────
AI_SYSTEM_PROMPT = (
    "당신은 공공기관 자료 취합 담당자를 돕는 검토자입니다. 회신 파일의 표를 행 단위로 받습니다. "
    "형식 오류·누락은 이미 규칙으로 검증했으므로, 문맥상 부적합(안내가 정한 기간·범위를 벗어난 값, "
    "질문과 무관한 서술, 문맥상 맞지 않는 분류)만 찾습니다. 판정 절차를 지키세요. "
    "① 항목마다 적용되는 기준을 '양식안내'와 '컬럼.작성기준'에서 찾는다. 기준이 없는 항목은 적합이다. "
    "② 값과 기준을 실제로 비교한다 — 날짜는 기간의 시작·끝을 포함해 안에 있는지, 숫자는 부등호 방향까지 "
    "확인한다(예: '3도 이상'이면 3.2·24.7은 적합). "
    "③ 기준을 실제로 위반한 값만 findings에 넣고, 위반한 기준 문장을 '위반기준'에, 값과 기준을 대조한 "
    "한 문장을 reason에 적는다. "
    "'분포'(건수·평균·최소·최대)는 참고 정보이지 한계가 아니다 — 분포의 최대값과 같거나 그 범위 안의 값을 "
    "부적합으로 판정하지 마라. 어느 컬럼의 기준도 없는 값(예: 보수여부의 표기)은 판정하지 마라. "
    "적합한 값을 findings에 넣으면 담당자가 헛일을 한다 — 확실하지 않으면 적합이다. "
    "행마다 verdict를 내고, findings가 비어 있으면 verdict는 적합이다. 항목 이름은 컬럼 헤더 글자 그대로다."
)
# 응답 스키마를 강제한다(structured outputs). json_object만 지정하면 id 누락·verdict 오타가
# 'AI 재검증 실패' 오류로 떨어졸다 — 스키마가 있으면 그 실패 경로 자체가 거의 사라진다.
AI_RESULT_SCHEMA = {
    "name": "review_results", "strict": True,
    "schema": {
        "type": "object", "additionalProperties": False, "required": ["results"],
        "properties": {"results": {"type": "array", "items": {
            "type": "object", "additionalProperties": False,
            "required": ["id", "verdict", "findings"],
            "properties": {
                "id": {"type": "string"},
                "verdict": {"type": "string", "enum": ["적합", "부적합"]},
                "findings": {"type": "array", "items": {
                    "type": "object", "additionalProperties": False,
                    "required": ["항목", "위반기준", "reason"],
                    "properties": {"항목": {"type": "string"}, "위반기준": {"type": "string"},
                                   "reason": {"type": "string"}}}}}}}},
    },
}

# 유료 호출 계측 — 호출 수·토큰을 프로세스 단위로 누적한다. 실행 단위 사용량은
# 스냅샷 차로 구한다(usage_snapshot). stderr 로그만으로는 비용을 화면·DB에 못 싣는다.
AI_USAGE = {"calls": 0, "prompt_tokens": 0, "completion_tokens": 0}
_USAGE_LOCK = threading.Lock()


def usage_snapshot() -> dict:
    with _USAGE_LOCK:
        return dict(AI_USAGE)


def usage_delta(before: dict) -> dict:
    now = usage_snapshot()
    return {k: now[k] - before.get(k, 0) for k in now}


def _record_usage(model: str, response, n_items: int, label: str) -> None:
    usage = getattr(response, "usage", None)
    with _USAGE_LOCK:
        AI_USAGE["calls"] += 1
        if usage is not None:
            AI_USAGE["prompt_tokens"] += usage.prompt_tokens or 0
            AI_USAGE["completion_tokens"] += usage.completion_tokens or 0
    if usage is not None:
        print(f"[aggregate:{label}] {model} 토큰 in={usage.prompt_tokens} "
              f"out={usage.completion_tokens} total={usage.total_tokens} (항목 {n_items}건)",
              file=sys.stderr)


def _stage1_blocks(uf: UploadedFile) -> tuple[set, set]:
    """1단계 위반이 난 셀과 컬럼. 이 칸은 2단계로 보내지 않는다 — 판정이 이미 났고, 보내면
    돈만 든다. 게이팅을 파일 단위에서 칸 단위로 낮춘 것이다(실측: 파일 단위면 실제 회신본
    8개 중 2단계에 닿는 파일이 거의 없었다)."""
    cells, cols = set(), set()
    for issue in uf.issues:
        if issue.grade == OK:
            continue
        if re.fullmatch(r"[A-Z]{1,3}\d+", issue.cell or ""):
            cells.add((issue.sheet, issue.cell))
        elif issue.attr:
            cols.add((issue.sheet, issue.attr))
    return cells, cols


def _ai_items(uf: UploadedFile) -> list[dict]:
    """AI 재검증 대상 — 행 하나가 항목 하나다. 전송되는 모든 값에 마스킹 적용 (§13.2).

    셀마다 보내면 같은 행의 참고값이 셀 수만큼 반복 전송된다(실측: 파일당 16~20초).
    행으로 묶으면 참고값이 한 번만 가고, 모델도 행 안의 관계를 한눈에 본다.
    """
    cells, cols = _stage1_blocks(uf)
    items = []
    for sheet in uf.sheets:
        for i, row in enumerate(sheet.rows):
            number = sheet.row_numbers[i]
            values = {}
            for idx, header in enumerate(sheet.headers):
                ctype = sheet.col_types[idx] if idx < len(sheet.col_types) else "freeform"
                if ctype == "image" or not header or idx >= len(row) or _is_blank(row[idx]):
                    continue
                if (sheet.name, f"{get_column_letter(idx + 1)}{number}") in cells \
                        or (sheet.name, header) in cols:
                    continue
                values[header] = mask(row[idx])
            if values:
                items.append({"id": f"{sheet.name}!{number}", "sheet": sheet.name,
                              "row": number, "항목": values})
    return items


def _ai_context(uf: UploadedFile, rules: dict | None = None) -> dict:
    """한 파일의 호출에 한 번만 붙는 컨텍스트 — 양식 안내문, 컬럼별 작성기준·유형·분포.

    '작성기준'이 헤더 문자열 하나뿐이면 모델은 기간·범위 같은 진짜 기준을 알 수 없다.
    양식의 작성 주의사항 시트(예: '24.11.01~25.02.28 기간, 3도 이상')가 그 기준이고,
    F1·F6 작성기준(rules)이 있으면 그것도 싣는다.
    """
    columns: dict[str, dict] = {}
    for sheet in uf.sheets:
        for idx, header in enumerate(sheet.headers):
            ctype = sheet.col_types[idx] if idx < len(sheet.col_types) else "freeform"
            if ctype == "image" or not header:
                continue
            info: dict = {"유형": ctype}
            rule = (rules or {}).get(header)
            if rule:
                info["작성기준"] = {k: v for k, v in rule.items() if k != "key"}
            if ctype == "numeric":
                numbers = [n for n in (_to_number(r[idx]) for r in sheet.rows if idx < len(r))
                           if n is not None]
                if numbers:
                    info["분포"] = {"건수": len(numbers), "평균": round(sum(numbers) / len(numbers), 1),
                                  "최소": min(numbers), "최대": max(numbers)}
            columns[f"{sheet.name}!{header}"] = info
    ctx: dict = {"부서": uf.dept, "컬럼": columns}
    if uf.notes_text:
        ctx["양식안내"] = mask(uf.notes_text)
    return ctx


def _ai_call(client, model: str, items: list[dict], context: dict | None = None) -> dict[str, dict]:
    payload = {"컨텍스트": context or {},
               "행": [{k: v for k, v in it.items() if k != "sheet"} for it in items]}
    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_schema", "json_schema": AI_RESULT_SCHEMA},
        messages=[
            {"role": "system", "content": AI_SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
    )
    _record_usage(model, response, len(items), "재검증")
    parsed = json.loads(response.choices[0].message.content)
    return {str(r.get("id")): r for r in parsed.get("results", []) if r.get("id")}


def _ai_client():
    """OpenAI 클라이언트. 만들 수 없으면 None (§13.3 서비스 전체 이용 불가).

    타임아웃·재시도를 명시한다 — 기본값이면 한 배치가 걸릴 때 취합 전체가 기다린다.
    """
    try:
        from openai import OpenAI
        return OpenAI(timeout=float(os.environ.get("AI_TIMEOUT_SEC", "90")),
                      max_retries=int(os.environ.get("AI_MAX_RETRIES", "2")))
    except Exception as exc:
        print(f"  · AI 재검증 생략(서비스 이용 불가: {exc}) → 정상(AI 미검증)", file=sys.stderr)
        return None


_OK_WORDS = ("적합합니다", "부합합니다", "적합함", "부합함", "문제 없", "문제없", "벗어나지 않")
_BAD_WORDS = ("부적합", "벗어남", "벗어납", "초과", "미달", "위반", "않습니다", "없습니다", "아닙니다")


def _self_contradicting(reason: str) -> bool:
    """부적합 판정의 사유가 오히려 '적합하다'고 말하는가.

    '…기간을 벗어나지 않아야 하지만, 작성된 일정은 적합합니다'처럼 결론이 적합인 문장.
    부정 표현('벗어납니다', '초과')이 함께 있으면 진짜 위반 설명일 수 있으니 살린다.
    """
    text = reason.replace(" ", "")
    return any(w.replace(" ", "") in text for w in _OK_WORDS) \
        and not any(w in text for w in _BAD_WORDS)


def _apply_verdicts(uf: UploadedFile, items: list[dict], verdicts: dict[str, dict]) -> None:
    for it in items:
        result = verdicts.get(it["id"])
        sheet = next((s for s in uf.sheets if s.name == it["sheet"]), None)
        if not result or result.get("verdict") not in ("적합", "부적합"):
            # §4.2.1 서비스는 정상이나 개별 항목 판정 불가 → 이상(AI 재검증 실패), 오류 등급
            uf.issues.append(Issue(uf.name, it["sheet"], f"{it['row']}행", "AI 재검증 실패", "2단계(실패)",
                                   ERROR,
                                   "AI 재검증 실패 — 사유: AI 응답을 해석할 수 없어 최종 판정을 내리지 못했습니다."
                                   "(1단계 규칙기반 검증은 통과했으나, 2단계에서 예상된 '적합/부적합' 형식이 아닌 "
                                   "응답을 받아 처리가 중단되었습니다.) 담당자가 직접 값을 확인해주세요."))
            continue
        if result["verdict"] != "부적합":
            continue
        # §6.2 LLM 판단은 확정적이지 않으므로 경고 등급. 자동교정 대상에서 제외(§8)
        findings = result.get("findings") or [{"항목": "", "reason": result.get("reason", "(사유 없음)")}]
        for finding in findings:
            reason = str(finding.get("reason") or "(사유 없음)")
            rule = str(finding.get("위반기준") or "").strip()
            if _self_contradicting(reason):
                # 실측(gpt-4o-mini): '…작성된 일정은 적합합니다'라는 사유로 부적합을 낸 건이 4건.
                # 사유가 적합이라고 말하면 판정이 아니라 잡음이다 — 담당자에게 넘기지 않는다
                continue
            header = str(finding.get("항목") or "")
            idx = sheet.headers.index(header) if sheet and header in sheet.headers else None
            cell = f"{get_column_letter(idx + 1)}{it['row']}" if idx is not None else f"{it['row']}행"
            text = f"AI 재검증 부적합: {reason}" + (f" [기준: {rule}]" if rule else "")
            uf.issues.append(Issue(uf.name, it["sheet"], cell, "정합성", "2단계", WARN, text, header))


# ── 표기 통일 후보 ─────────────────────────────────────────────────────────
# 같은 뜻을 다르게 적은 값('완료'·'보수완료')은 셀 단위 판정으로는 보이지 않는다. 컬럼의
# 고유값 집합을 모아 한 번 묶어 달라고 하고, 결과는 제안(경고)으로만 남긴다 — 자동교정하지
# 않는다(§8: AI 발견 이상은 자동교정 제외). 기관 취합본도 이 통일을 손으로 했다.
AI_NORMALIZE_PROMPT = (
    "같은 항목(컬럼)에 적힌 표기들을 항목별로 받습니다(값: 건수). 뜻이 같은데 표기만 다른 값들을 "
    "묶고 대표 표기를 정하세요. 대표는 가장 많이 쓰였거나 '양식안내'가 요구하는 표기입니다. 뜻이 "
    "다른 값(예: '보수완료'와 '보수중')은 절대 묶지 마세요. 확실한 것만 묶고, 묶을 것이 없는 "
    "항목은 groups를 빈 배열로 둡니다. 표기는 받은 값 글자 그대로 돌려줍니다."
)
AI_NORMALIZE_SCHEMA = {
    "name": "normalization", "strict": True,
    "schema": {
        "type": "object", "additionalProperties": False, "required": ["columns"],
        "properties": {"columns": {"type": "array", "items": {
            "type": "object", "additionalProperties": False, "required": ["항목", "groups"],
            "properties": {
                "항목": {"type": "string"},
                "groups": {"type": "array", "items": {
                    "type": "object", "additionalProperties": False, "required": ["대표", "표기"],
                    "properties": {"대표": {"type": "string"},
                                   "표기": {"type": "array", "items": {"type": "string"}}}}}}}}},
    },
}
NORMALIZE_MAX_LEN = 20      # 코드값은 짧다 — 긴 서술은 표기 통일 대상이 아니다
NORMALIZE_MAX_DISTINCT = 15  # 고유값이 이보다 많으면 코드표가 아니라 자유 입력이다


def _ai_normalize_call(client, model: str, columns: dict, notes: str) -> dict[str, dict[str, str]]:
    """{헤더: {값: 건수}} → {헤더: {표기: 대표}}. 대표와 같은 표기는 넣지 않는다."""
    payload = {"양식안내": notes, "항목": columns}
    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_schema", "json_schema": AI_NORMALIZE_SCHEMA},
        messages=[{"role": "system", "content": AI_NORMALIZE_PROMPT},
                  {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}])
    _record_usage(model, response, len(columns), "표기통일")
    parsed = json.loads(response.choices[0].message.content)
    out: dict[str, dict[str, str]] = {}
    for col in parsed.get("columns", []):
        header = str(col.get("항목") or "")
        for group in col.get("groups", []):
            canon = str(group.get("대표") or "").strip()
            for variant in group.get("표기", []):
                variant = str(variant).strip()
                if canon and variant and variant != canon:
                    out.setdefault(header, {})[variant] = canon
    return out


def _suggest_normalization(files: list[UploadedFile], client, model: str) -> None:
    counts: dict[str, dict[str, int]] = {}
    for uf in files:
        for sheet in uf.sheets:
            for idx, header in enumerate(sheet.headers):
                ctype = sheet.col_types[idx] if idx < len(sheet.col_types) else "freeform"
                if not header or ctype not in ("code_list", "freeform"):
                    continue
                for row in sheet.rows:
                    if idx >= len(row) or _is_blank(row[idx]):
                        continue
                    value = mask(row[idx]).strip()
                    if 0 < len(value) <= NORMALIZE_MAX_LEN:
                        counts.setdefault(header, {})[value] = counts.setdefault(header, {}).get(value, 0) + 1
    candidates = {h: v for h, v in counts.items() if 2 <= len(v) <= NORMALIZE_MAX_DISTINCT}
    if not candidates:
        return
    notes = next((mask(uf.notes_text) for uf in files if uf.notes_text), "")
    try:
        mapping = _ai_normalize_call(client, model, candidates, notes)
    except Exception as exc:                       # 제안은 부가물이다 — 실패해도 검토를 막지 않는다
        print(f"  · 표기 통일 후보 생략({exc})", file=sys.stderr)
        return
    for uf in files:
        for sheet in uf.sheets:
            for idx, header in enumerate(sheet.headers):
                variants = mapping.get(header)
                if not variants:
                    continue
                hits: dict[str, list[int]] = {}
                for i, row in enumerate(sheet.rows):
                    if idx < len(row) and not _is_blank(row[idx]):
                        value = mask(row[idx]).strip()
                        if value in variants:
                            hits.setdefault(value, []).append(sheet.row_numbers[i])
                if not hits:
                    continue
                letter = get_column_letter(idx + 1)
                rows_hit = sorted(n for ns in hits.values() for n in ns)
                span = f"{letter}{rows_hit[0]}" if len(rows_hit) == 1 else f"{letter}{rows_hit[0]}:{letter}{rows_hit[-1]}"
                detail = ", ".join(f"'{v}' {len(ns)}건 → '{variants[v]}'" for v, ns in hits.items())
                uf.issues.append(Issue(uf.name, sheet.name, span, "표기", "2단계", WARN,
                                       f"표기 통일 후보: {detail} (다른 회신본과 표기가 다릅니다. "
                                       "자동으로 바꾸지 않았습니다)", header))


def review_stage2_many(files: list[UploadedFile], model: str, batch_size: int = 100,
                       workers: int | None = None, on_done=None,
                       rules: dict | None = None) -> dict:
    """여러 파일의 2단계를 한 번에 병렬로 돌린다. 반환값은 이 실행의 API 사용량.

    게이팅은 **칸 단위**다 — 1단계 위반 셀·컬럼만 빼고 나머지 칸은 보낸다(`_stage1_blocks`).
    다만 오류 등급 파일은 기본 제외 대상(§7)이라 보내지 않는다(강제 포함 전까지는 비용만
    든다). 파일 간 호출은 서로 독립이므로 (파일, 배치)를 평평한 목록으로 만들어 단일
    동시성 상한으로 돌린다.

    실패 격리는 파일 단위로 유지한다(§13.3). 어느 배치가 터지면 그 파일만
    `정상(AI 미검증)`이 되고 다른 파일의 판정은 그대로 살아 있다.

    on_done(uf)는 파일 하나의 판정이 끝날 때마다 호출된다(진행률 표시용).
    """
    before = usage_snapshot()
    targets = [uf for uf in files if uf.readable and uf.sheets and uf.grade != ERROR]
    if not targets:
        return usage_delta(before)
    client = _ai_client()
    if client is None:
        for uf in targets:
            uf.ai_unverified = True
        return usage_delta(before)

    items_of = {id(uf): _ai_items(uf) for uf in targets}
    context_of = {id(uf): _ai_context(uf, rules) for uf in targets}
    jobs = [(uf, items_of[id(uf)][s:s + batch_size])
            for uf in targets
            for s in range(0, len(items_of[id(uf)]), batch_size)]
    if not jobs:
        return usage_delta(before)

    if workers is None:
        workers = int(os.environ.get("AI_CONCURRENCY", "6"))
    workers = max(1, min(workers, len(jobs)))

    verdicts: dict[int, dict[str, dict]] = {id(uf): {} for uf in targets}
    failed: set[int] = set()
    lock = threading.Lock()

    def run(job) -> None:
        uf, batch = job
        try:
            got = _ai_call(client, model, batch, context_of[id(uf)])
        except Exception as exc:                   # §13.3 서비스 장애·타임아웃
            with lock:
                if id(uf) not in failed:
                    failed.add(id(uf))
                    print(f"  · {uf.name}: AI 재검증 생략(서비스 장애: {exc}) → 정상(AI 미검증)",
                          file=sys.stderr)
            return
        with lock:
            verdicts[id(uf)].update(got)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        list(pool.map(run, jobs))

    for uf in targets:
        items = items_of[id(uf)]
        if id(uf) in failed:
            uf.ai_unverified = True
        elif items:
            missing = [it for it in items if it["id"] not in verdicts[id(uf)]]
            if missing:
                # ponytail: 실패 항목만 1회 재시도(개별 호출 대신 한 번에 묶어서).
                try:
                    verdicts[id(uf)].update(_ai_call(client, model, missing, context_of[id(uf)]))
                except Exception:
                    pass
            _apply_verdicts(uf, items, verdicts[id(uf)])
        if on_done:
            on_done(uf)

    _suggest_normalization([uf for uf in targets if id(uf) not in failed], client, model)
    return usage_delta(before)


def review_stage2(uf: UploadedFile, model: str, batch_size: int = 100,
                  rules: dict | None = None) -> dict:
    """파일 하나의 2단계."""
    return review_stage2_many([uf], model, batch_size=batch_size, rules=rules)


# ── 취합 계획 제안 ──────────────────────────────────────────────────────────
# 업로드 직후 양식 구조를 한 번 읽고 화면의 설정(키 컬럼·선택 입력 컬럼·부서명·취합할 표·
# 피벗)을 미리 채운다. 담당자가 손으로 취합할 때 첫 단계에서 하는 일이 이것이다. 결과는
# 사용자가 이미 통제하는 설정값이라 틀려도 화면에서 고칠 수 있다. 판단이 갈리면 질문으로 남긴다.
AI_PLAN_PROMPT = (
    "당신은 여러 부서가 같은 양식으로 회신한 엑셀 파일을 하나로 취합하는 담당자를 돕습니다. "
    "파일별 표 구조(헤더·샘플 행·행수), 파일명에서 뽑은 부서명 후보, 양식의 작성 안내문을 받습니다. "
    "취합 설정을 제안하세요. key_col: 행마다 고유한 컬럼 하나(부서·지사처럼 반복되는 컬럼은 키가 "
    "아닙니다. 고유 컬럼이 없으면 null). optional_cols: 비고처럼 비워 두는 것이 정상인 컬럼. "
    "dept_names: 파일별 부서명(fid → 이름). 후보 중 고르되, 모든 파일에 공통인 말은 제목이라 "
    "부서명이 아닙니다. sheets: 취합 대상 표 이름(작성 안내·빈 대장·총괄표는 대상이 아닙니다). "
    "pivot_sheets: 월×지표 총괄표처럼 한 줄로 눕혀야 하는 표. notes: 담당자가 알아야 할 관찰 "
    "(표기 불일치, 기간 밖 값 등) 최대 5개. questions: 확실하지 않아 확인이 필요한 것 최대 3개. "
    "이름은 받은 헤더·표 이름을 글자 그대로 쓰고, 없는 이름을 만들지 마세요."
)
AI_PLAN_SCHEMA = {
    "name": "aggregation_plan", "strict": True,
    "schema": {
        "type": "object", "additionalProperties": False,
        "required": ["key_col", "optional_cols", "dept_names", "sheets", "pivot_sheets", "notes", "questions"],
        "properties": {
            "key_col": {"type": ["string", "null"]},
            "optional_cols": {"type": "array", "items": {"type": "string"}},
            "dept_names": {"type": "array", "items": {
                "type": "object", "additionalProperties": False, "required": ["fid", "name"],
                "properties": {"fid": {"type": "integer"}, "name": {"type": "string"}}}},
            "sheets": {"type": "array", "items": {"type": "string"}},
            "pivot_sheets": {"type": "array", "items": {"type": "string"}},
            "notes": {"type": "array", "items": {"type": "string"}},
            "questions": {"type": "array", "items": {"type": "string"}},
        },
    },
}
PLAN_SAMPLE_ROWS = 2
PLAN_SAMPLE_FILES = 8      # 샘플 행은 앞 몇 파일만 — 헤더·행수는 전 파일


def _plan_call(client, model: str, payload: dict) -> dict:
    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_schema", "json_schema": AI_PLAN_SCHEMA},
        messages=[{"role": "system", "content": AI_PLAN_PROMPT},
                  {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}])
    _record_usage(model, response, len(payload.get("파일", [])), "취합계획")
    return json.loads(response.choices[0].message.content)


def plan_aggregation(files: list[UploadedFile], model: str, client=None) -> dict | None:
    """업로드된 파일들의 구조를 보고 취합 설정을 제안한다. AI를 못 쓰면 None.

    반환값의 이름은 전부 실제 헤더·표 이름에 대조해 걸러 낸다 — 모델이 지어낸 이름이
    화면 설정에 들어가면 안 된다.
    """
    readable = [(fid, uf) for fid, uf in enumerate(files) if uf.readable and uf.sheets]
    if not readable:
        return None
    client = client or _ai_client()
    if client is None:
        return None
    stems = [uf.path.stem for _, uf in readable]
    payload_files = []
    for n, (fid, uf) in enumerate(readable):
        tables = []
        for sheet in uf.sheets:
            table = {"name": sheet.name, "headers": [h for h in sheet.headers if h], "rows": len(sheet.rows)}
            if n < PLAN_SAMPLE_FILES:
                table["sample"] = [{h: mask(r[i]) for i, h in enumerate(sheet.headers)
                                    if h and i < len(r) and not _is_blank(r[i])}
                                   for r in sheet.rows[:PLAN_SAMPLE_ROWS]]
            tables.append(table)
        entry = {"fid": fid, "파일명": uf.name, "부서명후보": suggest_dept_names(uf.path.stem, stems),
                 "표": tables}
        if uf.note_sheets:
            entry["표아닌시트"] = uf.note_sheets
        payload_files.append(entry)
    notes = next((mask(uf.notes_text)[:1500] for _, uf in readable if uf.notes_text), "")
    try:
        raw = _plan_call(client, model, {"파일": payload_files, "양식안내": notes})
    except Exception as exc:
        print(f"  · 취합 계획 제안 생략({exc})", file=sys.stderr)
        return None

    headers = {h for _, uf in readable for s in uf.sheets for h in s.headers if h}
    names = {s.name for _, uf in readable for s in uf.sheets}
    fids = {fid for fid, _ in readable}
    key_col = raw.get("key_col")
    plan = {
        "key_col": key_col if isinstance(key_col, str) and key_col in headers else None,
        "optional_cols": [c for c in dict.fromkeys(raw.get("optional_cols") or []) if c in headers],
        "dept_names": {str(d["fid"]): str(d["name"]).strip()
                       for d in (raw.get("dept_names") or [])
                       if isinstance(d, dict) and d.get("fid") in fids
                       and str(d.get("name") or "").strip() and len(str(d["name"])) <= 30},
        "sheets": [s for s in dict.fromkeys(raw.get("sheets") or []) if s in names],
        "pivot_sheets": [s for s in dict.fromkeys(raw.get("pivot_sheets") or []) if s in names],
        "notes": [str(x) for x in (raw.get("notes") or [])][:5],
        "questions": [str(x) for x in (raw.get("questions") or [])][:3],
    }
    return plan


# ── 8. 전처리 ────────────────────────────────────────────────────────────────
def preprocess(uf: UploadedFile) -> None:
    """1단계 경고 항목에 자동 교정을 적용한다. 원본값은 Fix에 보존 (§8 감사 추적).

    2단계(AI) 판정은 자동 교정하지 않는다. 이미지도 변환하지 않는다.
    """
    by_cell = {(f.sheet, f.cell): f for f in uf.fixes if f.corrected != "행 제거"}
    drop = {(f.sheet, int(re.sub(r"\D", "", f.cell) or 0))
            for f in uf.fixes if f.corrected == "행 제거"}
    for sheet in uf.sheets:
        keep_rows, keep_numbers = [], []
        for i, row in enumerate(sheet.rows):
            row_number = sheet.row_numbers[i]
            if (sheet.name, row_number) in drop:
                continue
            new_row = list(row)
            for idx in range(len(new_row)):
                fix = by_cell.get((sheet.name, f"{get_column_letter(idx + 1)}{row_number}"))
                if fix:
                    number = _to_number(fix.corrected)
                    new_row[idx] = number if number is not None and _is_number(row[idx]) else fix.corrected
            keep_rows.append(new_row)
            keep_numbers.append(row_number)
        sheet.rows, sheet.row_numbers = keep_rows, keep_numbers


# ── 9. 합성 (F2-6~F2-10) ─────────────────────────────────────────────────────
def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:SHEET_NAME_LIMIT]
    candidate, n = base, 1
    while candidate in used:
        suffix = f"~{n}"
        candidate = base[:SHEET_NAME_LIMIT - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def _shift_anchor(im: dict, row_offset: int, col_offset: int):
    """anchor의 종류·점유 범위·표시 크기는 그대로 두고 위치만 옮긴다 (§9 재배치 규칙)."""
    fc, fr, fco, fro = im["marker_from"]
    frm = AnchorMarker(col=max(0, fc + col_offset), row=max(0, fr + row_offset),
                       colOff=fco, rowOff=fro)
    if im["marker_to"] is None:
        w, h = im["display"] or im["resolution"]
        return OneCellAnchor(_from=frm, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
    tc, tr, tco, tro = im["marker_to"]
    to = AnchorMarker(col=max(0, tc + col_offset), row=max(0, tr + row_offset),
                      colOff=tco, rowOff=tro)
    return TwoCellAnchor(editAs=im["edit_as"] or "twoCell", _from=frm, to=to)


def _row_offset(sheet: SheetData, target_first_row: int) -> int:
    """원본 데이터 첫 행이 결과에서 target_first_row로 가도록 하는 이동량."""
    origin = sheet.row_numbers[0] if sheet.row_numbers else sheet.header_row + 1
    return target_first_row - origin


def _add_images(ws, sheet: SheetData, row_offset: int, col_offset: int = 0) -> None:
    """이미지를 원본과 같은 데이터 행·컬럼에 다시 붙인다 (§9)."""
    for im in sheet.images:
        if not im["data"]:
            continue
        try:
            picture = XLImage(io.BytesIO(im["data"]))
        except Exception:
            continue
        ws.add_image(picture, _shift_anchor(im, row_offset, col_offset))


def _write_block(ws, headers: list[str], rows: list[list], start_row: int) -> int:
    if start_row == 1:
        for c, header in enumerate(headers, start=1):
            ws.cell(row=1, column=c, value=header)
        start_row = 2
    for row in rows:
        for c, value in enumerate(row, start=1):
            ws.cell(row=start_row, column=c, value=value)
        start_row += 1
    return start_row


def _format_for(entries, template_path) -> xf.FormatTemplate:
    """출력 시트 하나의 기준 서식. 규칙 1(등록 양식) → 규칙 2(전 파일 다수결).

    다수결에 넣는 순서는 파일명 오름차순이다 — 동수일 때 첫 파일이 이기게 하려면
    순서가 결정적이어야 한다.
    """
    if template_path:
        for _, sheet in entries:
            tpl = xf.capture(template_path, sheet.name, sheet.header_row)
            if tpl:
                tpl.source_label = f"등록 양식 '{sheet.name}'"
                return tpl
    caps = [xf.capture(uf.path, sheet.name, sheet.header_row)
            for uf, sheet in sorted(entries, key=lambda e: e[0].name)]
    return xf.consensus(caps) or xf.basic()


def _org_key(name) -> str:
    """조직명 비교 키 — 공백과 조직 꼬리(지사·사업소·부…)를 뗀다.

    파일명은 '수원사업소'·'화성지사'로 오고 회신본의 지사 칸은 '수원'·'화성'으로 적혀 있다
    (실측 8개 중 3개). 같은 조직을 표기만 다르게 부른 것이라 꼬리를 떼고 비교한다.
    꼬리는 긴 것부터, 남는 글자가 있을 때까지만 뗀다('구' 하나만 남는 일은 없다).
    """
    key = str(name or "").replace(" ", "").strip()
    suffixes = sorted(ORG_SUFFIXES, key=len, reverse=True)
    while True:
        hit = next((s for s in suffixes if key.endswith(s) and len(key) > len(s)), None)
        if hit is None:
            return key
        key = key[:-len(hit)]


def _reusable_dept_col(headers: list[str], entries) -> int | None:
    """양식의 소속 컬럼(지사·부서)을 부서 구분으로 그대로 쓸 수 있으면 그 인덱스(0-base).

    모든 파일의 모든 행에서 값이 그 파일의 부서명과 같은 조직이어야 한다(`_org_key`).
    다르면(회신본이 소속을 다르게 적었거나 사용자가 다른 부서명을 골랐으면) '부서' 컬럼을
    앞에 붙이는 종전 동작이 안전하다 — 어느 쪽 값이 맞는지 여기서 판정할 근거가 없다.
    """
    for idx, header in enumerate(headers):
        if not _is_dept_header(header):
            continue
        if all(idx < len(row) and not _is_blank(row[idx])
               and _org_key(row[idx]) == _org_key(uf.dept)
               for uf, sheet in entries for row in sheet.rows):
            return idx
    return None


def synthesize(files: list[UploadedFile], mode: str, group_map: dict, summary_cols: list[str],
               all_files: list[UploadedFile] | None = None, template_path=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    notes: list[str] = []
    labels: list[str] = []           # 시트마다 쓴 기준 서식 — 개요·안내에 밝힌다
    total_rows = 0

    if mode == "A":
        for uf in files:
            for sheet in uf.sheets:
                ws = wb.create_sheet(_safe_sheet_name(f"{uf.dept}_{sheet.name}", used))
                # 모드 A는 원본 보존형이라 규칙 1·2를 타지 않는다 — 각 시트가 자기 파일의
                # 서식을 쓴다. 다수결로 남의 열너비를 씌우면 원본 보존이 아니게 된다
                tpl = xf.capture(uf.path, sheet.name, sheet.header_row) or xf.basic()
                labels.append(tpl.source_label)
                first = xf.open_sheet(ws, tpl, sheet.headers)
                last = _write_block(ws, [], sheet.rows, first) - 1
                _restore_vmerges(ws, sheet, first, 0)
                _add_images(ws, sheet, _row_offset(sheet, first))
                xf.close_sheet(ws, tpl, sheet.headers, first, last)
                total_rows += len(sheet.rows)
        label = labels[0] if len(set(labels)) == 1 else f"파일별 원본 서식 {len(set(labels))}종"
        notes.append(f"기준 서식: {label}")
        run = build_summary(all_files if all_files is not None else files, files,
                            mode, label, total_rows)
        xf.overview_sheet(wb, 0, run)
        return wb, notes, run

    # B/C/D 공통: 세로 누적. C는 그룹 단위, B/D는 시트명 단위.
    buckets: dict[str, list[tuple[UploadedFile, SheetData]]] = {}
    for uf in files:
        for sheet in uf.sheets:
            if mode == "C":
                key = group_map.get(sheet.name)
                if key is None:
                    key = "미분류"
                    notes.append(f"'{sheet.name}' 시트는 그룹 매핑에 없어 '미분류' 그룹으로 편입했습니다.")
            else:
                key = sheet.name
            buckets.setdefault(key, []).append((uf, sheet))

    dept_rows: dict[str, dict[str, tuple[int, int]]] = {}   # 시트 → 부서 → (첫행, 마지막행)
    header_rows: dict[str, int] = {}                        # 시트 → 헤더 행(요약 수식이 참조)
    dept_cols: dict[str, str] = {}                          # 시트 → 부서 구분 컬럼의 헤더명

    for key, entries in buckets.items():
        base_headers = entries[0][1].headers
        title = _safe_sheet_name(key, used)
        ws = wb.create_sheet(title)
        tpl = _format_for(entries, template_path)
        labels.append(tpl.source_label)
        # 양식에 이미 소속 컬럼(지사·부서)이 있고 값이 부서명과 같으면 그 컬럼이 부서 구분이다 —
        # 같은 값을 든 '부서' 컬럼을 하나 더 붙이면 결과가 원본 양식과 달라지고 정보도 늘지 않는다
        same = [(uf, s) for uf, s in entries if s.headers == base_headers]
        dept_idx = _reusable_dept_col(base_headers, same)
        extra = (["구분"] if mode == "C" else []) + ([] if dept_idx is not None else ["부서"])
        out_headers = extra + base_headers
        dept_col = len(extra) + dept_idx + 1 if dept_idx is not None else len(extra)   # 1-base
        dept_cols[title] = out_headers[dept_col - 1]
        cursor = xf.open_sheet(ws, tpl, out_headers)
        first_data_row = cursor
        header_rows[title] = cursor - 1
        err_ws = None
        err_tpl = None
        err_headers: list[str] = []
        err_first = 0
        for uf, sheet in entries:
            if sheet.headers != base_headers:
                # §9 모드 B 예외: 부서 간 헤더 불일치 → 오류 시트로 분리
                if err_ws is None:
                    err_ws = wb.create_sheet(_safe_sheet_name(f"{key}_헤더불일치", used))
                    err_tpl = xf.capture(uf.path, sheet.name, sheet.header_row) or xf.basic()
                    err_headers = ["부서", "시트"] + sheet.headers
                    err_first = xf.open_sheet(err_ws, err_tpl, err_headers)
                notes.append(f"'{uf.dept}'의 '{sheet.name}' 시트 헤더가 기준과 달라 '{err_ws.title}' 시트로 분리했습니다.")
                rows = [[uf.dept, sheet.name] + list(r) for r in sheet.rows]
                _write_block(err_ws, [], rows, max(err_ws.max_row + 1, err_first))
                total_rows += len(rows)   # 헤더불일치 시트로 빠진 행도 결과에 실제로 존재한다
                continue
            prefix = ([sheet.name] if mode == "C" else []) + ([] if dept_idx is not None else [uf.dept])
            rows = [prefix + list(r) for r in sheet.rows]
            first = cursor
            cursor = _write_block(ws, [], rows, cursor)
            if rows:
                dept_rows.setdefault(title, {})[uf.dept] = (first, cursor - 1)
                if mode != "D":
                    _restore_vmerges(ws, sheet, first, len(extra), skip_col=dept_col)
            # extra(구분·부서)가 앞에 끼므로 이미지도 그만큼 오른쪽으로 밀어야 한다
            _add_images(ws, sheet, _row_offset(sheet, first), col_offset=len(extra))
        total_rows += cursor - first_data_row
        # 원본이 합계 행을 갖고 있었으면 결과 끝에도 둔다 — 값이 아니라 SUM 수식으로, 원본이
        # 더한 컬럼만. 취합 중에는 뺐던 행이라(이중 합산 방지) 여기서 한 번만 다시 만든다
        total_src = next((s for _, s in same if s.total_cols), None)
        if total_src is not None and cursor > first_data_row:
            off = len(extra)
            ws.cell(row=cursor, column=total_src.total_label_col + 1 + off,
                    value=total_src.total_label)
            for c in total_src.total_cols:
                col_letter = get_column_letter(c + 1 + off)
                ws.cell(row=cursor, column=c + 1 + off,
                        value=f"=SUM({col_letter}{first_data_row}:{col_letter}{cursor - 1})")
            cursor += 1
        xf.close_sheet(ws, tpl, out_headers, first_data_row, cursor - 1)
        # 부서 구간의 소속 칸을 세로 병합한다 — 기관이 손으로 만드는 취합본의 모양이다.
        # 모드 D는 제외한다: 병합하면 첫 칸 외에는 값이 비어 종합요약의 SUMIF가 첫 행만 더한다
        if mode != "D":
            for lo, hi in dept_rows.get(title, {}).values():
                if hi > lo:
                    ws.merge_cells(start_row=lo, start_column=dept_col, end_row=hi, end_column=dept_col)
        if err_ws is not None:
            xf.close_sheet(err_ws, err_tpl, err_headers, err_first, err_ws.max_row)

    # 표가 아닌 시트(작성 주의사항 등)가 모든 파일에 같은 이름으로 있으면 결과에 한 번 옮긴다 —
    # 회신 양식의 일부이고 취합본을 받는 사람도 같은 안내를 본다. 등록 양식이 있으면 거기서,
    # 없으면 첫 파일에서 가져온다(같은 양식을 배포해 회신받았다는 전제, §9)
    shared_notes = [n for n in (files[0].note_sheets if files else [])
                    if all(n in uf.note_sheets for uf in files)]
    for pos, name in enumerate(shared_notes):
        for src in ([template_path] if template_path else []) + [files[0].path]:
            if xf.copy_sheet(src, name, wb, _safe_sheet_name(name, used), index=pos):
                notes.append(f"'{name}' 시트는 표가 아니라 {Path(src).name}에서 그대로 옮겼습니다.")
                break

    if mode == "D":
        _add_summary_sheet(wb, files, summary_cols, dept_rows, dept_cols, notes, header_rows)

    label = labels[0] if len(set(labels)) == 1 else f"시트별 기준 서식 {len(set(labels))}종"
    notes.append(f"기준 서식: {label}")
    run = build_summary(all_files if all_files is not None else files, files,
                        mode, label, total_rows)
    xf.overview_sheet(wb, 0, run)     # 결과 파일만 받아도 무엇이 빠졌는지 보이게
    return wb, notes, run


def _add_summary_sheet(wb, files, summary_cols, dept_rows, dept_cols, notes, header_rows) -> None:
    """종합요약 시트를 최상단에 추가. 수치는 SUMIF 수식 기반(하드코딩 금지, §9)."""
    ws = wb.create_sheet("종합요약", 0)
    ws.cell(row=1, column=1, value="부서")
    for c, col in enumerate(summary_cols, start=2):
        ws.cell(row=1, column=c, value=col)
    depts = sorted({uf.dept for uf in files})
    for r, dept in enumerate(depts, start=2):
        ws.cell(row=r, column=1, value=dept)
        for c, col in enumerate(summary_cols, start=2):
            terms = []
            for title, ranges in dept_rows.items():
                target = wb[title]
                # 제목 블록을 복원하면서 헤더가 1행이 아니게 됐다 — 시트마다 실제 헤더 행을 본다.
                # dept_rows에 시트가 있으면 header_rows에도 반드시 있어야 한다(같은 루프에서
                # 함께 채워진다) — 없으면 조용히 1행을 보는 대신 KeyError로 바로 드러낸다
                hr = header_rows[title]
                headers = [target.cell(row=hr, column=i).value
                           for i in range(1, target.max_column + 1)]
                if col not in headers or dept not in ranges:
                    continue
                value_letter = get_column_letter(headers.index(col) + 1)
                dept_letter = get_column_letter(headers.index(dept_cols[title]) + 1)
                lo, hi = ranges[dept]
                terms.append(f"SUMIF('{title}'!{dept_letter}{lo}:{dept_letter}{hi},"
                             f"$A{r},'{title}'!{value_letter}{lo}:{value_letter}{hi})")
            if terms:
                ws.cell(row=r, column=c, value="=" + "+".join(terms))
            else:
                ws.cell(row=r, column=c, value="N/A")     # §9 지정 컬럼 부재 시
                notes.append(f"요약 지표 '{col}'을(를) 찾을 수 없어 '{dept}' 행에 N/A로 표기했습니다.")
    xf.close_sheet(ws, xf.basic(), ["부서"] + list(summary_cols), 2, ws.max_row)


def _reason_of(uf: UploadedFile) -> str:
    """제외·경고 사유를 사람이 읽을 한 줄로 줄인다. 같은 사유는 건수로 묶는다."""
    if not uf.readable:
        return "파일을 읽지 못했습니다"
    pool = [i for i in uf.issues if i.grade == ERROR] or \
           [i for i in uf.issues if i.grade == WARN] or \
           [i for i in uf.issues if i.grade == OK]
    if not pool:
        return ""
    counts: dict[str, int] = {}
    for issue in pool:
        # 사유 뒤의 괄호는 셀 위치 같은 개별 정보라 묶을 때는 뗀다. AI 사유는 건마다 문장이
        # 달라 그대로 두면 개요 한 칸이 문단이 된다 — 종류('AI 재검증 부적합')로 묶는다
        label = issue.reason.split("(")[0].strip()
        if issue.stage.startswith("2단계"):
            label = issue.reason.split(":")[0].strip()
        counts[label] = counts.get(label, 0) + 1
    shown = list(counts.items())[:3]
    text = " · ".join(f"{label} {n}건" for label, n in shown)
    if len(counts) > len(shown):
        # 잘랐다는 사실을 밝힌다 — 조용히 줄이면 그게 전부인 것처럼 읽힌다
        text += f" 외 {len(counts) - len(shown)}종"
    return text


def build_summary(all_files, selected, mode: str, source_label: str,
                  result_rows: int = 0) -> xf.RunSummary:
    """취합 개요에 실을 현황. CLI·웹이 같은 값을 쓰도록 엔진에서 만든다."""
    picked = {id(uf) for uf in selected}
    rows = [xf.FileRow(dept=uf.dept, name=uf.name, status=uf.status,
                       included=id(uf) in picked,
                       rows=sum(len(s.rows) for s in uf.sheets),
                       reason=_reason_of(uf))
            for uf in all_files]
    return xf.RunSummary(mode=mode, source_label=source_label,
                         result_rows=result_rows, files=rows)


# ── 10.2 오류 리포트 ─────────────────────────────────────────────────────────
def write_report(files: list[UploadedFile], path: Path, overview=None) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "오류 목록"
    issue_headers = ["파일명", "시트", "셀 위치", "오류 유형", "검증 단계", "사유", "작성 기준"]
    ws.append(issue_headers)
    for uf in files:
        for issue in uf.issues:
            ws.append([issue.file, issue.sheet, issue.cell, issue.kind, issue.stage,
                       f"{issue.tag} {issue.reason}", issue.attr])
    ws2 = wb.create_sheet("자동교정 이력")
    fix_headers = ["파일명", "시트", "셀 위치", "원본값", "교정값", "적용 규칙"]
    ws2.append(fix_headers)
    for uf in files:
        for fix in uf.fixes:
            ws2.append([fix.file, fix.sheet, fix.cell, fix.original, fix.corrected, fix.rule])
    if ws2.max_row == 1:
        # 헤더만 있는 빈 시트는 만들다 만 것처럼 보인다 — 없다는 사실을 적는다
        ws2.append(["해당 없음"])

    # 리포트는 원본 양식과 무관한 문서라 상속 대상이 아니다. 기본서식만 입힌다
    xf.close_sheet(ws, xf.basic(), issue_headers, 2, ws.max_row)
    xf.close_sheet(ws2, xf.basic(), fix_headers, 2, ws2.max_row)
    if overview is not None:
        xf.overview_sheet(wb, 0, overview)
    wb.save(path)


# ── 6/7. 이상판별 출력 및 범위 선택 ───────────────────────────────────────────
def print_review(files: list[UploadedFile]) -> None:
    print("\n=== 이상판별 결과 ===")
    for uf in files:
        status = uf.status
        if status == "정상" and uf.ai_unverified:
            status = "정상(AI 미검증)"
        print(f"[{status}] {uf.name}")
        for issue in uf.issues:
            print(f"    {issue.line()}")


def load_env(path: Path = Path(".env")) -> None:
    """.env의 KEY=VALUE를 환경변수로 올린다 (§13.1: 키는 코드에 하드코딩 금지).

    이미 설정된 환경변수는 덮어쓰지 않는다(셸에서 export한 값이 우선).
    """
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip("'\""))


def main(argv=None) -> int:
    load_env()
    parser = argparse.ArgumentParser(description="엑셀 취합 (F2)")
    parser.add_argument("input_dir", help="취합할 .xlsx 파일이 있는 폴더")
    parser.add_argument("--mode", choices=list("ABCD"), default="B", help="합성 모드 (기본 B)")
    parser.add_argument("--out", default="merged.xlsx", help="취합 결과 파일 경로")
    parser.add_argument("--report", default="error_report.xlsx", help="오류 리포트 파일 경로")
    parser.add_argument("--rules", help="작성기준 JSON (§5.4). 없으면 시스템 기본값")
    parser.add_argument("--group-map", help="모드 C 그룹 매핑 JSON {시트명: 그룹명}")
    parser.add_argument("--summary-cols", default="", help="모드 D 요약 지표 컬럼(콤마 구분)")
    parser.add_argument("--include-anomalous", action="store_true",
                        help="오류 등급 파일도 강제로 취합에 포함 (§7)")
    parser.add_argument("--no-ai", action="store_true", help="2단계 AI 재검증 생략")
    parser.add_argument("--model", default=os.environ.get("OPENAI_MODEL", "gpt-5-mini"))
    parser.add_argument("--include-hidden", action="store_true", help="숨김 시트 포함")
    parser.add_argument("--pivot-sheets", default="",
                        help="월×지표 피벗표(총괄표 등) 시트명, 콤마 구분. 한 줄로 눕혀 "
                             f"'{PIVOT_SHEET_NAME}' 한 시트로 모은다")
    parser.add_argument("--period", default=None,
                        help="계획-실적 검사 기간(월), 예: 1-3 또는 3. 생략하면 파일명에서 "
                             "추정('1분기'·'상반기'·'1~3월' 표기)")
    args = parser.parse_args(argv)

    period = None
    if args.period:
        m = re.fullmatch(r"(\d{1,2})(?:\s*[~\-]\s*(\d{1,2}))?", args.period.strip())
        if not m or not (1 <= int(m.group(1)) <= int(m.group(2) or m.group(1)) <= 12):
            print(f"--period 형식이 잘못됐습니다: '{args.period}' (예: 1-3)", file=sys.stderr)
            return 2
        period = (int(m.group(1)), int(m.group(2) or m.group(1)))

    rules = json.loads(Path(args.rules).read_text(encoding="utf-8")) if args.rules else {}
    group_map = json.loads(Path(args.group_map).read_text(encoding="utf-8")) if args.group_map else {}
    summary_cols = [c.strip() for c in args.summary_cols.split(",") if c.strip()]

    paths = sorted(p for p in Path(args.input_dir).iterdir()
                   if p.is_file() and not p.name.startswith("~$"))
    if not paths:
        print(f"'{args.input_dir}'에 파일이 없습니다.", file=sys.stderr)
        return 1

    files = []
    # 부서명 후보의 '공통어'는 같은 양식으로 회신한 xlsx끼리 봐야 한다. 폴더에 pdf·zip이
    # 섞이면(실측 27개) 공통어가 사라져 제목의 일부('지열지점')가 부서명으로 뽑힌다
    stems = [p.stem for p in paths if p.suffix.lower() == ".xlsx"]
    for path in paths:
        print(f"· {path.name} 검토 중")
        uf = read_file(path, args.include_hidden,
                       pivot_sheets={n.strip() for n in args.pivot_sheets.split(',') if n.strip()}
                       or None)
        # 부서명은 파일명 전체가 아니라 후보 1순위다 — 같이 올라온 파일명의 공통어(제목)를
        # 걷어낸 나머지(웹은 같은 후보를 칩으로 보여 준다). 검토 전에 정해야
        # 소속 컬럼 자동 채움(review_stage1)이 이 이름을 쓴다.
        uf.dept = suggest_dept_names(path.stem, stems)[0]
        if uf.readable:
            review_stage1(uf, rules, period=period)
            if uf.pair_period:
                print(f"  · 계획-실적 검사 기간: {uf.pair_period[0]}~{uf.pair_period[1]}월"
                      + ("" if period else " (파일명에서 추정)"))
            if uf.grade == ERROR:
                # 오류 등급 파일은 기본 제외 대상이라 2단계로 보내지 않는다(§4.2·§7).
                # 경고만 있는 파일은 위반 칸만 빼고 나머지 칸을 보낸다(칸 단위 게이팅)
                print("  · 1단계 오류 → 2단계 AI 재검증 미진입(기본 제외 파일, API 호출 절약)")
        files.append(uf)

    if not args.no_ai:
        # 2단계는 파일마다 API 호출이라 순차로 돌리면 파일 수에 비례해 늘어난다.
        # 파일 간 호출은 서로 독립적이므로 병렬로 묶는다.
        pending = [uf for uf in files if uf.readable and uf.sheets and uf.grade != ERROR]
        if pending:
            print(f"· 2단계 AI 재검증 {len(pending)}건 (동시 호출 "
                  f"{os.environ.get('AI_CONCURRENCY', '6')}개)")
            used = review_stage2_many(files, args.model, rules=rules)
            if used["calls"]:
                print(f"  · API 호출 {used['calls']}회 · 토큰 in={used['prompt_tokens']} "
                      f"out={used['completion_tokens']}")

    print_review(files)

    selected = [uf for uf in files if uf.readable and uf.sheets
                and (uf.grade != ERROR or args.include_anomalous)]
    forced = [uf for uf in selected if uf.grade == ERROR]
    print(f"\n취합 대상: {len(selected)}/{len(files)}건")
    if not selected:
        write_report(files, Path(args.report),
                     overview=build_summary(files, [], args.mode, ""))
        print(f"취합 가능한 파일이 없습니다. 오류 리포트: {args.report}")
        return 2

    for uf in selected:
        preprocess(uf)

    wb, notes, run = synthesize(selected, args.mode, group_map, summary_cols, all_files=files)
    wb.save(args.out)
    write_report(files, Path(args.report), overview=run)

    # 개요는 데이터 시트가 아니라 안내다. 세면 '결과 시트 N개'가 틀린 값이 된다
    # (server.py의 같은 계산과 이유가 같다)
    data_sheets = [s for s in wb.sheetnames if s != xf.OVERVIEW_SHEET]
    print(f"\n결과: {args.out} (시트 {len(data_sheets)}개)")
    print(f"오류 리포트: {args.report}")
    if forced:
        # §10.1 특이사항 안내 — 다운로드를 막지 않는다
        print("\n[특이사항] 강제로 포함한 항목 중 형식 오류가 있어 결과 파일의 일부 값·수식이 "
              "정상 계산되지 않았을 수 있습니다(예: #VALUE!). 아래를 확인 후 원본을 수정해주세요.")
        for uf in forced:
            for issue in (i for i in uf.issues if i.grade == ERROR):
                print(f"    {issue.line()}")
    for note in dict.fromkeys(notes):
        print(f"[안내] {note}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
