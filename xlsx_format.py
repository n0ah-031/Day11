#!/usr/bin/env python3
"""취합 결과에 원본 양식의 서식을 입힌다.

aggregate.py의 `_write_block`은 값만 쓴다. 이 모듈이 원본 xlsx에서 서식을 읽고
(capture) 여러 회신본의 공통값을 고르고(consensus) 출력 시트에 입힌다
(open_sheet → 값 쓰기 → close_sheet).

설계·근거: docs/superpowers/specs/2026-08-07-취합문서-서식상속-design.md

서식은 부가물이다. 공개 함수는 예외를 밖으로 내지 않는다 — 서식을 못 입혀도
값은 반드시 나와야 한다.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MAX_WIDTH = 60.0                    # 파일명처럼 긴 값이 열을 화면 밖으로 밀지 않게
MIN_WIDTH = 8.43                    # 엑셀 기본 열너비
DEFAULT_DATE_FORMAT = "yyyy-mm-dd"
_WIDTH_PER_CHAR = 1.4               # 한글이 섞이면 글자당 1자 너비로는 좁다
_FIT_SCAN_ROWS = 200                # 너비 계산에 훑는 최대 행수(전량 훑을 이유가 없다)

_GREY = PatternFill("solid", fgColor="DDDDDD")
_THIN = Side(style="thin")
_BOX = Border(bottom=_THIN, top=_THIN, left=_THIN, right=_THIN)


@dataclass
class CellStyle:
    """셀 하나의 서식. openpyxl 스타일 객체는 공유되므로 넣고 뺄 때 복사한다."""
    font: object = None
    fill: object = None
    border: object = None
    alignment: object = None

    @classmethod
    def of(cls, cell) -> "CellStyle":
        return cls(copy(cell.font), copy(cell.fill), copy(cell.border), copy(cell.alignment))

    def put(self, cell) -> None:
        if self.font is not None:
            cell.font = copy(self.font)
        if self.fill is not None:
            cell.fill = copy(self.fill)
        if self.border is not None:
            cell.border = copy(self.border)
        if self.alignment is not None:
            cell.alignment = copy(self.alignment)

    def key(self) -> str:
        """다수결에서 세기 위한 값 비교용 키."""
        return f"{self.font}|{self.fill}|{self.border}|{self.alignment}"


@dataclass
class TitleRow:
    cells: list = field(default_factory=list)   # [(열번호 1-base, 값, CellStyle)]
    height: float | None = None


@dataclass
class FormatTemplate:
    header_row: int = 1
    title_rows: list = field(default_factory=list)      # 1 ~ header_row-1
    title_merges: list = field(default_factory=list)    # (min_row, min_col, max_row, max_col)
    header_style: CellStyle | None = None
    data_style: CellStyle | None = None
    widths: dict = field(default_factory=dict)          # 헤더명 → 너비
    number_formats: dict = field(default_factory=dict)  # 헤더명 → 표시형식
    orientation: str = "portrait"
    fit_to_page: bool = False
    print_title_rows: str | None = None
    date_format: str = DEFAULT_DATE_FORMAT
    source_label: str = ""


def _is_date_format(nf: str) -> bool:
    low = (nf or "").lower()
    return "yy" in low or "mm-dd" in low or "m/d" in low


def capture(path, sheet_name: str, header_row: int) -> FormatTemplate | None:
    """원본 xlsx 한 시트에서 서식만 읽는다. 못 읽으면 None(예외를 내지 않는다)."""
    try:
        if header_row < 1:
            return None
        wb = openpyxl.load_workbook(Path(path))
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        if header_row > ws.max_row:
            return None

        # 헤더 이름 ← 컬럼 번호. 이 대응이 열너비·숫자서식 매칭의 기준이다
        headers: dict[int, str] = {}
        for c in range(1, ws.max_column + 1):
            value = ws.cell(row=header_row, column=c).value
            if value is not None and str(value).strip():
                headers[c] = str(value)
        if not headers:
            return None

        first = min(headers)
        tpl = FormatTemplate(
            header_row=header_row,
            header_style=CellStyle.of(ws.cell(row=header_row, column=first)),
            data_style=CellStyle.of(ws.cell(row=header_row + 1, column=first)),
            orientation=ws.page_setup.orientation or "portrait",
            fit_to_page=bool(ws.sheet_properties.pageSetUpPr
                             and ws.sheet_properties.pageSetUpPr.fitToPage),
            print_title_rows=ws.print_title_rows,
            source_label=Path(path).name,
        )

        for c, name in headers.items():
            # dict 인덱싱([])은 openpyxl에서 미설정 컬럼에도 기본폭(13.0)을 만들어
            # 낸다 — 실제로 폭을 지정한 컬럼만 담기 위해 .get()으로 존재 여부를 본다
            col_dim = ws.column_dimensions.get(get_column_letter(c))
            width = col_dim.width if col_dim is not None else None
            if width:
                tpl.widths[name] = round(float(width), 1)
            fmt = ws.cell(row=header_row + 1, column=c).number_format
            if fmt and fmt != "General":
                tpl.number_formats[name] = fmt
                if _is_date_format(fmt):
                    # 기본서식 폴백에서 쓸 날짜 표시형식 — 지어내지 않고 원본에서 본 것을 쓴다
                    tpl.date_format = fmt

        last_col = max(headers)
        for r in range(1, header_row):
            cells = [(c, ws.cell(row=r, column=c).value, CellStyle.of(ws.cell(row=r, column=c)))
                     for c in range(1, last_col + 1)]
            tpl.title_rows.append(TitleRow(cells, ws.row_dimensions[r].height))
        for rng in ws.merged_cells.ranges:
            if rng.max_row < header_row:
                tpl.title_merges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
        tpl.title_merges.sort()
        return tpl
    except Exception:
        return None


def basic(date_format: str = DEFAULT_DATE_FORMAT) -> FormatTemplate:
    """상속할 원본이 없을 때의 기본서식(설계 §6). 지어낸 스타일이 아니라 최소한의 가독성."""
    return FormatTemplate(
        header_row=1,
        header_style=CellStyle(Font(bold=True), _GREY, _BOX,
                               Alignment(horizontal="center", vertical="center")),
        data_style=None,
        date_format=date_format,
        source_label="기본서식",
    )


def _mode(values: list, key=None):
    """최빈값과 (득표수, 전체수). 동수면 먼저 들어온 것 — 호출자가 파일명 순으로 넣는다.

    dict는 삽입 순서를 지키고 max는 동점일 때 먼저 만난 키를 돌려주므로
    별도 tie-break 없이 '첫 파일 우선'이 된다.
    """
    counts: dict = {}
    first: dict = {}
    for value in values:
        k = key(value) if key else str(value)
        counts[k] = counts.get(k, 0) + 1
        first.setdefault(k, value)
    if not counts:
        return None, 0, 0
    best = max(counts, key=lambda k: counts[k])
    return first[best], counts[best], len(values)


def _title_key(tpl: FormatTemplate) -> str:
    rows = [(tr.height, [(c, v) for c, v, _ in tr.cells]) for tr in tpl.title_rows]
    return f"{rows}|{tpl.title_merges}"


def consensus(templates: list) -> FormatTemplate | None:
    """항목별 최빈값으로 기준 서식을 만든다 (설계 §3 규칙 2).

    파일 단위로 대표 1개를 뽑지 않는 이유: 지사마다 손댄 곳이 다르면 '완전히
    일치하는 그룹'이 잘게 부서진다. 항목별로 세면 각 항목이 원래 값으로 수렴한다.
    """
    tpls = [t for t in templates if t is not None]
    if not tpls:
        return None
    if len(tpls) == 1:
        return tpls[0]

    out = FormatTemplate()
    out.header_row = _mode([t.header_row for t in tpls])[0]
    out.header_style = _mode([t.header_style for t in tpls], key=lambda s: s.key())[0]
    out.data_style = _mode([t.data_style for t in tpls], key=lambda s: s.key())[0]
    out.orientation = _mode([t.orientation for t in tpls])[0]
    out.fit_to_page = _mode([t.fit_to_page for t in tpls])[0]
    out.print_title_rows = _mode([t.print_title_rows for t in tpls])[0]
    out.date_format = _mode([t.date_format for t in tpls])[0]

    # 제목 블록은 통째로 한 벌을 고른다 — 행마다 섞으면 어느 파일에도 없던 제목이 된다
    picked = _mode(tpls, key=_title_key)[0]
    out.title_rows, out.title_merges = picked.title_rows, picked.title_merges

    minority = 0
    for name in {n for t in tpls for n in t.widths}:
        value, hit, total = _mode([t.widths[name] for t in tpls if name in t.widths])
        out.widths[name] = value
        if hit < total:
            minority += 1
    for name in {n for t in tpls for n in t.number_formats}:
        out.number_formats[name] = _mode(
            [t.number_formats[name] for t in tpls if name in t.number_formats])[0]

    out.source_label = f"취합 파일 {len(tpls)}개의 공통 서식"
    if minority:
        out.source_label += f" (열너비 {minority}개는 다수값 채택)"
    return out


def open_sheet(ws, tpl: FormatTemplate, col_offset: int, out_headers: list) -> int:
    """제목 블록과 헤더 값을 쓰고 데이터 시작 행을 돌려준다.

    제목 블록은 열 오프셋을 적용하지 않는다 — 문서 제목은 표의 컬럼이 아니라
    시트 전체에 걸린 것이라, 앞에 부서 컬럼이 끼었다고 오른쪽으로 밀 이유가 없다.
    대신 병합 범위는 출력 열수까지 넓혀 제목이 표 전체를 덮게 한다.
    """
    try:
        for r, title in enumerate(tpl.title_rows, start=1):
            for c, value, style in title.cells:
                cell = ws.cell(row=r, column=c)
                if value is not None:
                    cell.value = value
                style.put(cell)
            if title.height:
                ws.row_dimensions[r].height = title.height
        for r0, c0, r1, c1 in tpl.title_merges:
            ws.merge_cells(start_row=r0, start_column=c0,
                           end_row=r1, end_column=max(c1, len(out_headers)))
        header_row = tpl.header_row
        for c, name in enumerate(out_headers, start=1):
            ws.cell(row=header_row, column=c, value=name)
        return header_row + 1
    except Exception:
        # 제목 블록을 못 쓰면 헤더만 1행에 놓고 진행한다. 값은 나와야 한다
        for c, name in enumerate(out_headers, start=1):
            ws.cell(row=1, column=c, value=name)
        return 2


def _has_datetime(ws, col: int, first_row: int, last_row: int) -> bool:
    """그 컬럼에 날짜/시각 값이 하나라도 있는가. 표시형식을 상속받지 못한 날짜 컬럼을
    찾는 데만 쓴다 — 값을 바꾸지 않고 보이는 형식만 정한다."""
    for r in range(first_row, min(last_row, first_row + _FIT_SCAN_ROWS) + 1):
        if isinstance(ws.cell(row=r, column=col).value, (datetime, date)):
            return True
    return False


def _fit_width(ws, col: int, header_row: int, last_row: int) -> float:
    longest = 0
    for r in range(header_row, min(last_row, header_row + _FIT_SCAN_ROWS) + 1):
        value = ws.cell(row=r, column=col).value
        if value is not None:
            longest = max(longest, len(str(value)))
    return max(MIN_WIDTH, min(MAX_WIDTH, longest * _WIDTH_PER_CHAR + 2))


def close_sheet(ws, tpl: FormatTemplate, out_headers: list, first_data_row: int,
                last_row: int, row_heights: dict | None = None) -> None:
    """헤더·데이터 서식, 열너비, 숫자서식, 틀고정·필터·인쇄설정을 마무리한다.

    자동필터 범위와 인쇄영역은 원본 값을 쓰지 않고 **결과 행수로 재계산**한다.
    회신본마다 데이터 행수가 달라 원본 값(A4:H63 등)은 결과와 무관하기 때문이다.
    """
    try:
        header_row = tpl.header_row
        last_row = max(last_row, header_row)

        if tpl.header_style is not None:
            for c in range(1, len(out_headers) + 1):
                tpl.header_style.put(ws.cell(row=header_row, column=c))
        if tpl.data_style is not None:
            for r in range(first_data_row, last_row + 1):
                for c in range(1, len(out_headers) + 1):
                    tpl.data_style.put(ws.cell(row=r, column=c))

        for c, name in enumerate(out_headers, start=1):
            fmt = tpl.number_formats.get(name)
            if not fmt and _has_datetime(ws, c, first_data_row, last_row):
                # 상속할 표시형식이 없는 날짜 컬럼 — 그냥 두면 '2025-01-08 00:00:00'으로
                # 보인다(이 작업이 고치려는 증상 그 자체다). 지어내지 않고 기준 서식에서
                # 본 날짜 형식을 쓰고, 그것도 없으면 DEFAULT_DATE_FORMAT
                fmt = tpl.date_format
            if fmt:
                for r in range(first_data_row, last_row + 1):
                    ws.cell(row=r, column=c).number_format = fmt
            width = tpl.widths.get(name)
            if width is None:
                width = _fit_width(ws, c, header_row, last_row)
            ws.column_dimensions[get_column_letter(c)].width = width

        last_col = get_column_letter(max(len(out_headers), 1))
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
        ws.auto_filter.ref = f"A{header_row}:{last_col}{last_row}"
        ws.print_area = f"A1:{last_col}{last_row}"
        if tpl.print_title_rows:
            ws.print_title_rows = tpl.print_title_rows
        ws.page_setup.orientation = tpl.orientation
        if tpl.fit_to_page:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0

        # 사진을 옮겨 붙인 행만 원본 높이를 따른다. 나머지까지 원본 높이(85.5 등)를
        # 쓰면 사진 없는 표가 화면을 넘긴다
        for row, height in (row_heights or {}).items():
            if height:
                ws.row_dimensions[row].height = height
    except Exception:
        pass
