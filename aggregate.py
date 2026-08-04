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

ERROR, WARN, OK = "오류", "경고", "정상"
GRADE_ORDER = {ERROR: 2, WARN: 1, OK: 0}   # §7 대표 등급 = 최악 등급

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


@dataclass
class UploadedFile:
    path: Path
    dept: str
    sheets: list[SheetData] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    fixes: list[Fix] = field(default_factory=list)
    ai_unverified: bool = False      # §6.4 정상(AI 미검증)
    readable: bool = True

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def grade(self) -> str:
        return max((i.grade for i in self.issues), key=lambda g: GRADE_ORDER[g], default=OK)

    @property
    def status(self) -> str:
        # §6.1 화면 표기는 정상/이상 이진화
        return "이상" if self.issues else "정상"


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


def _find_header(grid: list[list]) -> int:
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
        if _looks_like_data(nxt):
            return i                      # 다음 행이 숫자/날짜를 보임 → 헤더 확정
    # 전 컬럼이 텍스트인 시트(숫자 없음)는 위 규칙으로 확정되지 않으므로,
    # 제목행이 아닌 첫 텍스트 행을 헤더로 채택한다.
    for i, row in enumerate(grid):
        filled = [v for v in row if not _is_blank(v)]
        if len(filled) >= 2 and all(isinstance(v, str) for v in filled):
            if [v for v in (grid[i + 1] if i + 1 < len(grid) else []) if not _is_blank(v)]:
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
        hi = _find_header(sub)
        if hi < 0:
            break
        # 계층 헤더는 여러 행이다. 병합으로 범위를 잡고 한 줄로 합친 뒤, 데이터는 그 아래부터.
        sub_merged = [(r0 - start, c0, r1 - start, c1) for r0, c0, r1, c1 in merged
                      if r1 - start >= 1]
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


def read_file(path: Path, include_hidden: bool = False,
              include_sheets: set[str] | None = None) -> UploadedFile:
    """xlsx 한 개를 읽는다.

    `include_sheets`를 주면 그 시트만 읽는다. 실제 양식에는 취합 대상이 아닌 시트가
    섞여 있다 — 작성 가이드, 비워 둔 대장, 지사별로 이름이 다른 총괄표. 어느 시트가
    데이터인지는 담당자가 알고 있으므로 이름 규칙으로 맞히지 않고 골라 받는다.
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
            continue

        added = 0
        for part, (hi, headers, rows, row_numbers, dropped, blanks) in enumerate(tables, 1):
            # 표가 둘 이상이면 뒤 표에 이름을 붙인다. 시작 행이 파일마다 같으므로
            # 모드 B에서 같은 표끼리 묶인다(실측: 12개 회신본 모두 35행에서 둘째 표 시작).
            name = ws.title if len(tables) == 1 else f"{ws.title} ({part})"
            # 어떤 행이 데이터인지는 원본 기준으로 먼저 정하고, 병합 값은 그 뒤에 채운다.
            # 순서가 바뀌면 미리 병합해 둔 빈 템플릿 행까지 값이 생겨 데이터로 되살아난다.
            _fill_merged(rows, row_numbers, merged)
            if dropped:
                uf.issues.append(Issue(path.name, name, f"{dropped[0]}행" if len(dropped) == 1
                                       else f"{dropped[0]}~{dropped[-1]}행",
                                       "정합성", "1단계", WARN,
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
            uf.sheets.append(SheetData(name, hi + 1, headers, rows, row_numbers, mine,
                                       merged=merged))
            added += 1
        if added and len(tables) > 1:
            uf.issues.append(Issue(path.name, ws.title, "(해당없음)", "정합성", "1단계", WARN,
                                   f"이 시트에서 표 {len(tables)}개를 찾아 따로 읽었습니다"
                                   f"(시작 행 {', '.join(str(t[0] + 1) for t in tables)}). "
                                   "한 표여야 한다면 표 사이의 빈 행을 지워주세요."))

    # 취합할 시트가 하나라도 있으면, 표가 아닌 시트는 파일을 막을 사유가 아니다.
    # 실제 업무 양식에는 '작성 주의사항'처럼 안내문만 있는 시트가 흔하다 —
    # 이걸 오류로 보면 파일 전체가 '이상'이 되어 취합에서 기본 제외된다(§7).
    if uf.sheets:
        for issue in skipped:
            uf.issues.append(Issue(issue.file, issue.sheet, issue.cell, issue.kind, issue.stage,
                                   WARN, "표 구조가 아니어서 취합에서 제외했습니다"
                                         "(작성 안내·빈 시트 등). 취합 대상 시트라면 "
                                         "제목 행과 데이터 사이에 열 제목 행이 있는지 확인해주세요."))
    else:
        uf.issues.extend(skipped)
    return uf


def _fill_merged(rows: list[list], row_numbers: list[int],
                 merged: list[tuple[int, int, int, int]]) -> None:
    """병합 셀의 값을 그 범위에 속한 데이터 행에 채운다.

    엑셀에서 세로로 병합된 칸은 사람 눈에는 구간 전체의 값이지만, 파일에는
    좌상단 한 칸만 값이 있고 나머지는 비어 있다. 채우지 않으면 실제 업무 양식의
    '지사'·'개소'처럼 여러 행에 걸친 칸이 전부 필수값 누락으로 잡힌다.

    이미 데이터로 확정된 행만 대상이다 — 양식이 미리 병합해 둔 빈 행까지 채우면
    값이 생겨 데이터 행으로 되살아난다(실제 양식은 5~60행을 미리 병합해 둔다).
    """
    if not rows:
        return
    index_of = {number: i for i, number in enumerate(row_numbers)}
    for min_row, min_col, max_row, max_col in merged:
        if min_row == max_row and min_col == max_col:
            continue
        source = index_of.get(min_row)
        if source is None:
            continue                       # 병합 시작 행이 데이터가 아니면(제목 등) 건너뛴다
        width = len(rows[source])
        for c in range(min_col - 1, min(max_col, width)):
            value = rows[source][c]
            if _is_blank(value):
                continue
            for number in range(min_row, max_row + 1):
                target = index_of.get(number)
                if target is not None and _is_blank(rows[target][c]):
                    rows[target][c] = value


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


def review_stage1(uf: UploadedFile, rules: dict) -> None:
    """파일 내 모든 셀을 끝까지 스캔한다(첫 위반에서 중단하지 않음, §4.1)."""
    for sheet in uf.sheets:
        sheet.col_types = classify_columns(sheet, rules)
        for idx, header in enumerate(sheet.headers):
            ctype = sheet.col_types[idx]
            if ctype == "image" or not header:
                continue
            rule = rules.get(header, {})
            required = rule.get("required", True)
            letter = get_column_letter(idx + 1)
            for i, row in enumerate(sheet.rows):
                cell = f"{letter}{sheet.row_numbers[i]}"
                raw = row[idx]
                # 누락 → 정합성 → 중복/충돌 순서. 누락이면 이후 검사 생략 (§4.1)
                if _is_blank(raw):
                    # 작성기준이 선택 입력으로 정한 칸은 비어 있는 것이 정상이므로
                    # 보고하지 않는다. 담당자가 조치할 것이 없고, 비고처럼 대개
                    # 비워두는 컬럼에서는 행 수만큼 알림이 쌓여 리포트를 덮는다.
                    if required:
                        uf.issues.append(Issue(uf.name, sheet.name, cell, "누락", "1단계", ERROR,
                                               "필수값 누락", header))
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

        _review_duplicates(uf, sheet, rules)
        _review_images(uf, sheet, rules)


def _review_duplicates(uf: UploadedFile, sheet: SheetData, rules: dict) -> None:
    """동일 키 그룹핑 → 그룹 내 값 비교. 완전 중복=경고, 값 상이=충돌(오류) (§5.1)."""
    if not sheet.headers:
        return
    key_idx = next((i for i, h in enumerate(sheet.headers) if rules.get(h, {}).get("key")),
                   _pick_key_column(sheet.headers, sheet.rows))
    letter = get_column_letter(key_idx + 1)
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
                uf.issues.append(Issue(uf.name, sheet.name, cell, "중복", "1단계", WARN,
                                       f"키 '{key}' 완전 중복 행(자동 제거 가능)", sheet.headers[key_idx]))
                uf.fixes.append(Fix(uf.name, sheet.name, cell, f"중복 행(키 {key})", "행 제거", "완전 중복 행 자동 제거"))
            else:
                uf.issues.append(Issue(uf.name, sheet.name, cell, "충돌", "1단계", ERROR,
                                       f"키 '{key}'가 같으나 값이 상이한 행 존재", sheet.headers[key_idx]))


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
    "당신은 공공기관 실적자료 검토 담당자입니다. 각 항목의 '작성기준'과 '값'을 보고 값이 기준에 "
    "부합하는지 판정하세요. 형식 오류는 이미 별도 규칙으로 검증되었으므로, 문맥상 부적합(질문과 "
    "무관한 서술, 문맥상 맞지 않는 분류, 다른 행 대비 비합리적인 수치)만 '부적합'으로 판정합니다. "
    '반드시 {"results":[{"id":"...","verdict":"적합"|"부적합","reason":"..."}]} 형태의 JSON만 출력하세요.'
)


def _ai_items(uf: UploadedFile) -> list[dict]:
    """AI 재검증 대상 항목 생성. 전송되는 모든 값에 마스킹 적용 (§13.2)."""
    items = []
    for sheet in uf.sheets:
        for idx, header in enumerate(sheet.headers):
            ctype = sheet.col_types[idx] if idx < len(sheet.col_types) else "freeform"
            if ctype == "image" or not header:
                continue
            letter = get_column_letter(idx + 1)
            stats = ""
            if ctype == "numeric":
                numbers = [n for n in (_to_number(r[idx]) for r in sheet.rows) if n is not None]
                if numbers:
                    mean = sum(numbers) / len(numbers)
                    stats = (f"같은 컬럼 분포: 건수 {len(numbers)}, 평균 {mean:.1f}, "
                             f"최소 {min(numbers)}, 최대 {max(numbers)}")
            for i, row in enumerate(sheet.rows):
                if _is_blank(row[idx]):
                    continue
                item = {
                    "id": f"{sheet.name}!{letter}{sheet.row_numbers[i]}",
                    "sheet": sheet.name,
                    "cell": f"{letter}{sheet.row_numbers[i]}",
                    "작성기준": header,
                    "값": mask(row[idx]),
                }
                if ctype == "numeric" and stats:
                    item["컬럼분포"] = stats
                elif ctype in ("code_list", "freeform"):
                    # 인접 컬럼 컨텍스트 (패턴 기반은 값 자체로 판정 가능하므로 제외, §4.2)
                    item["같은행참고"] = {
                        h: mask(row[k]) for k, h in enumerate(sheet.headers)
                        if k != idx and h and not _is_blank(row[k])
                    }
                    if ctype == "freeform":
                        text = str(row[idx])
                        item["특징추출"] = {
                            "길이": len(text),
                            "헤더키워드포함": any(t and t in text for t in re.split(r"[\s·/]", header)),
                        }
                items.append(item)
    return items


def _ai_call(client, model: str, items: list[dict]) -> dict[str, dict]:
    payload = [{k: v for k, v in it.items() if k not in ("sheet", "cell")} for it in items]
    response = client.chat.completions.create(
        model=model,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": AI_SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
    )
    parsed = json.loads(response.choices[0].message.content)
    return {str(r.get("id")): r for r in parsed.get("results", []) if r.get("id")}


def _ai_client():
    """OpenAI 클라이언트. 만들 수 없으면 None (§13.3 서비스 전체 이용 불가)."""
    try:
        from openai import OpenAI
        return OpenAI()
    except Exception as exc:
        print(f"  · AI 재검증 생략(서비스 이용 불가: {exc}) → 정상(AI 미검증)", file=sys.stderr)
        return None


def _apply_verdicts(uf: UploadedFile, items: list[dict], verdicts: dict[str, dict]) -> None:
    for it in items:
        result = verdicts.get(it["id"])
        if not result or result.get("verdict") not in ("적합", "부적합"):
            # §4.2.1 서비스는 정상이나 개별 항목 판정 불가 → 이상(AI 재검증 실패), 오류 등급
            uf.issues.append(Issue(uf.name, it["sheet"], it["cell"], "AI 재검증 실패", "2단계(실패)", ERROR,
                                   "AI 재검증 실패 — 사유: AI 응답을 해석할 수 없어 최종 판정을 내리지 못했습니다."
                                   "(1단계 규칙기반 검증은 통과했으나, 2단계에서 예상된 '적합/부적합' 형식이 아닌 "
                                   "응답을 받아 처리가 중단되었습니다.) 담당자가 직접 값을 확인해주세요.",
                                   it["작성기준"]))
        elif result["verdict"] == "부적합":
            # §6.2 LLM 판단은 확정적이지 않으므로 경고 등급. 자동교정 대상에서 제외(§8)
            uf.issues.append(Issue(uf.name, it["sheet"], it["cell"], "정합성", "2단계", WARN,
                                   f"AI 재검증 부적합: {result.get('reason', '(사유 없음)')}", it["작성기준"]))


def review_stage2_many(files: list[UploadedFile], model: str, batch_size: int = 100,
                       workers: int | None = None, on_done=None) -> None:
    """여러 파일의 2단계를 한 번에 병렬로 돌린다.

    §4.2 게이팅이 파일 단위이므로 파일 간 호출은 서로의 판정에 영향을 주지 않고,
    한 파일 안의 배치도 독립적이다. 그래서 (파일, 배치)를 하나의 평평한 작업 목록으로
    만들어 단일 동시성 상한으로 돌린다 — 파일별로 풀을 중첩하면 동시 호출 수를
    통제할 수 없다.

    실패 격리는 파일 단위로 유지한다(§13.3). 어느 배치가 터지면 그 파일만
    `정상(AI 미검증)`이 되고 다른 파일의 판정은 그대로 살아 있다.

    on_done(uf)는 파일 하나의 판정이 끝날 때마다 호출된다(진행률 표시용).
    """
    targets = [uf for uf in files if uf.readable and not uf.issues]
    if not targets:
        return
    client = _ai_client()
    if client is None:
        for uf in targets:
            uf.ai_unverified = True
        return

    items_of = {id(uf): _ai_items(uf) for uf in targets}
    jobs = [(uf, items_of[id(uf)][s:s + batch_size])
            for uf in targets
            for s in range(0, len(items_of[id(uf)]), batch_size)]
    if not jobs:
        return

    if workers is None:
        workers = int(os.environ.get("AI_CONCURRENCY", "6"))
    workers = max(1, min(workers, len(jobs)))

    verdicts: dict[int, dict[str, dict]] = {id(uf): {} for uf in targets}
    failed: set[int] = set()
    lock = threading.Lock()

    def run(job) -> None:
        uf, batch = job
        try:
            got = _ai_call(client, model, batch)
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
                    verdicts[id(uf)].update(_ai_call(client, model, missing))
                except Exception:
                    pass
            _apply_verdicts(uf, items, verdicts[id(uf)])
        if on_done:
            on_done(uf)


def review_stage2(uf: UploadedFile, model: str, batch_size: int = 100) -> None:
    """파일 하나의 2단계. 파일 전체가 1단계를 통과한 경우에만 호출된다(§4.2)."""
    review_stage2_many([uf], model, batch_size=batch_size)


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


def synthesize(files: list[UploadedFile], mode: str, group_map: dict, summary_cols: list[str]):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    notes: list[str] = []

    if mode == "A":
        # 원본 보존형: {부서명}_{원본시트명} 시트 n×m개
        for uf in files:
            for sheet in uf.sheets:
                ws = wb.create_sheet(_safe_sheet_name(f"{uf.dept}_{sheet.name}", used))
                _write_block(ws, sheet.headers, sheet.rows, 1)
                # 헤더가 1행으로 당겨지므로 제목행이 있던 만큼 이미지도 위로 올라간다
                _add_images(ws, sheet, _row_offset(sheet, 2))
        return wb, notes

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

    extra = ["구분", "부서"] if mode == "C" else ["부서"]
    dept_rows: dict[str, dict[str, tuple[int, int]]] = {}   # sheet → dept → (첫행, 마지막행)

    for key, entries in buckets.items():
        base_headers = entries[0][1].headers
        title = _safe_sheet_name(key, used)
        ws = wb.create_sheet(title)
        err_ws = None
        cursor = _write_block(ws, extra + base_headers, [], 1)
        for uf, sheet in entries:
            if sheet.headers != base_headers:
                # §9 모드 B 예외: 부서 간 헤더 불일치 → 오류 시트로 분리
                if err_ws is None:
                    err_ws = wb.create_sheet(_safe_sheet_name(f"{key}_헤더불일치", used))
                    _write_block(err_ws, ["부서", "시트"] + sheet.headers, [], 1)
                notes.append(f"'{uf.dept}'의 '{sheet.name}' 시트 헤더가 기준과 달라 '{err_ws.title}' 시트로 분리했습니다.")
                rows = [[uf.dept, sheet.name] + list(r) for r in sheet.rows]
                _write_block(err_ws, [], rows, err_ws.max_row + 1)
                continue
            prefix = [sheet.name, uf.dept] if mode == "C" else [uf.dept]
            rows = [prefix + list(r) for r in sheet.rows]
            first = cursor
            cursor = _write_block(ws, [], rows, cursor)
            if rows:
                dept_rows.setdefault(title, {})[uf.dept] = (first, cursor - 1)
            # extra(구분·부서)가 앞에 끼므로 이미지도 그만큼 오른쪽으로 밀어야 한다
            _add_images(ws, sheet, _row_offset(sheet, first), col_offset=len(extra))

    if mode == "D":
        _add_summary_sheet(wb, files, summary_cols, dept_rows, extra, notes)
    return wb, notes


def _add_summary_sheet(wb, files, summary_cols, dept_rows, extra, notes) -> None:
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
                headers = [target.cell(row=1, column=i).value for i in range(1, target.max_column + 1)]
                if col not in headers or dept not in ranges:
                    continue
                value_letter = get_column_letter(headers.index(col) + 1)
                dept_letter = get_column_letter(headers.index("부서") + 1)
                lo, hi = ranges[dept]
                terms.append(f"SUMIF('{title}'!{dept_letter}{lo}:{dept_letter}{hi},"
                             f"$A{r},'{title}'!{value_letter}{lo}:{value_letter}{hi})")
            if terms:
                ws.cell(row=r, column=c, value="=" + "+".join(terms))
            else:
                ws.cell(row=r, column=c, value="N/A")     # §9 지정 컬럼 부재 시
                notes.append(f"요약 지표 '{col}'을(를) 찾을 수 없어 '{dept}' 행에 N/A로 표기했습니다.")


# ── 10.2 오류 리포트 ─────────────────────────────────────────────────────────
def write_report(files: list[UploadedFile], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "오류 목록"
    ws.append(["파일명", "시트", "셀 위치", "오류 유형", "검증 단계", "사유", "작성 기준"])
    for uf in files:
        for issue in uf.issues:
            ws.append([issue.file, issue.sheet, issue.cell, issue.kind, issue.stage,
                       f"{issue.tag} {issue.reason}", issue.attr])
    ws2 = wb.create_sheet("자동교정 이력")
    ws2.append(["파일명", "시트", "셀 위치", "원본값", "교정값", "적용 규칙"])
    for uf in files:
        for fix in uf.fixes:
            ws2.append([fix.file, fix.sheet, fix.cell, fix.original, fix.corrected, fix.rule])
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
    args = parser.parse_args(argv)

    rules = json.loads(Path(args.rules).read_text(encoding="utf-8")) if args.rules else {}
    group_map = json.loads(Path(args.group_map).read_text(encoding="utf-8")) if args.group_map else {}
    summary_cols = [c.strip() for c in args.summary_cols.split(",") if c.strip()]

    paths = sorted(p for p in Path(args.input_dir).iterdir()
                   if p.is_file() and not p.name.startswith("~$"))
    if not paths:
        print(f"'{args.input_dir}'에 파일이 없습니다.", file=sys.stderr)
        return 1

    files = []
    for path in paths:
        print(f"· {path.name} 검토 중")
        uf = read_file(path, args.include_hidden)
        if uf.readable:
            review_stage1(uf, rules)
            if uf.issues:
                # 파일 단위 게이팅: 1단계 위반이 하나라도 있으면 2단계로 진입하지 않는다 (§4.2)
                print("  · 1단계 위반 발견 → 2단계 AI 재검증 미진입(API 호출 절약)")
        files.append(uf)

    if not args.no_ai:
        # 2단계는 파일마다 API 호출이라 순차로 돌리면 파일 수에 비례해 늘어난다.
        # 파일 간 호출은 서로 독립적이므로 병렬로 묶는다.
        pending = [uf for uf in files if uf.readable and not uf.issues]
        if pending:
            print(f"· 2단계 AI 재검증 {len(pending)}건 (동시 호출 "
                  f"{os.environ.get('AI_CONCURRENCY', '6')}개)")
            review_stage2_many(files, args.model)

    print_review(files)

    selected = [uf for uf in files if uf.readable and uf.sheets
                and (uf.grade != ERROR or args.include_anomalous)]
    forced = [uf for uf in selected if uf.grade == ERROR]
    print(f"\n취합 대상: {len(selected)}/{len(files)}건")
    if not selected:
        write_report(files, Path(args.report))
        print(f"취합 가능한 파일이 없습니다. 오류 리포트: {args.report}")
        return 2

    for uf in selected:
        preprocess(uf)

    wb, notes = synthesize(selected, args.mode, group_map, summary_cols)
    wb.save(args.out)
    write_report(files, Path(args.report))

    print(f"\n결과: {args.out} (시트 {len(wb.sheetnames)}개)")
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
