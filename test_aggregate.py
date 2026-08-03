#!/usr/bin/env python3
"""aggregate.py 자체 점검. 실행: python3 test_aggregate.py (프레임워크 없음)"""

import json
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

import aggregate as ag


def make_book(path: Path, sheets: dict, title_row: bool = True) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        if title_row:
            # §2.2 예시 구조: 1행 병합 제목 → 2행 공백 → 3행 헤더 → 4행 데이터
            ws["A1"] = f"{name} 실적 보고"
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
            ws.append([None, None, None, None])
        for row in rows:
            ws.append(row)
    wb.save(path)


def test_structure_and_stage1(tmp: Path):
    """제목행 skip + 누락/정합성/충돌을 전부(첫 위반에서 멈추지 않고) 수집한다."""
    path = tmp / "기획부.xlsx"
    make_book(path, {"실적": [
        ["사번", "부서", "예산액", "비고"],
        ["A01", "기획부", 1000, "정상 집행"],
        ["A02", "기획부", None, "누락 케이스"],       # 누락 → 오류
        ["A03", "기획부", "abc", "숫자아님"],          # 정합성 → 오류
        ["A04", "기획부", "1,200", "콤마"],            # 정합성 → 경고(자동교정)
        ["A01", "기획부", 9999, "키충돌"],             # 충돌 → 오류
    ]})
    uf = ag.read_file(path)
    assert uf.readable, "xlsx를 읽어야 한다"
    assert len(uf.sheets) == 1, uf.issues
    sheet = uf.sheets[0]
    assert sheet.header_row == 3, f"제목행(1)·공백 아님을 넘어 3행이 헤더여야 함, got {sheet.header_row}"
    assert sheet.headers == ["사번", "부서", "예산액", "비고"], sheet.headers
    assert sheet.row_numbers[0] == 4, sheet.row_numbers

    ag.review_stage1(uf, rules={})
    kinds = sorted({i.kind for i in uf.issues})
    assert "누락" in kinds and "정합성" in kinds and "충돌" in kinds, f"{kinds} / {[i.line() for i in uf.issues]}"
    assert uf.grade == ag.ERROR and uf.status == "이상"
    # 예산액 컬럼은 값 타입 분포로 numeric 판정되어야 한다
    assert sheet.col_types[2] == "numeric", sheet.col_types
    # 자동교정 이력에 원본값이 보존되어야 한다 (§8 감사 추적)
    assert any(f.original == "1,200" for f in uf.fixes), uf.fixes
    print("  ✓ 구조 인식 + 1단계 규칙 검증(전량 스캔)")


def test_clean_file_and_gating(tmp: Path):
    """위반 없는 파일은 정상, 위반 파일은 2단계 미진입(파일 단위 게이팅)."""
    clean = tmp / "총무부.xlsx"
    make_book(clean, {"실적": [
        ["사번", "부서", "예산액"],
        ["B01", "총무부", 500],
        ["B02", "총무부", 700],
    ]})
    uf = ag.read_file(clean)
    ag.review_stage1(uf, rules={})
    assert not uf.issues, [i.line() for i in uf.issues]
    assert uf.status == "정상"

    # AI 서비스 이용 불가(키 없음) → §13.3 폴백: 정상(AI 미검증)으로 남고 이상이 되지 않는다
    ag.review_stage2(uf, model="nonexistent-model")
    assert uf.status == "정상", [i.line() for i in uf.issues]
    assert uf.ai_unverified, "장애 시 AI 미검증 플래그가 서야 한다"
    print("  ✓ 정상 파일 판정 + AI 장애 폴백(정상(AI 미검증))")


def test_rules_and_masking(tmp: Path):
    """작성기준(§5.4)이 적용되고, 외부 전송 값은 마스킹된다(§13.2)."""
    path = tmp / "인사부.xlsx"
    make_book(path, {"명단": [
        ["사번", "부서", "이메일", "연봉"],
        ["C01", "없는부서", "a@b.com", 100],
        ["C02", "인사부", "bad-email", 100],
    ]})
    rules = {
        "부서": {"type": "code_list", "master_list": ["인사부", "총무부"]},
        "이메일": {"type": "pattern"},
        "연봉": {"type": "numeric", "min": 0, "max": 1000},
    }
    uf = ag.read_file(path)
    ag.review_stage1(uf, rules)
    reasons = " ".join(i.reason for i in uf.issues)
    assert "코드표" in reasons, reasons
    assert "허용 형식" in reasons, reasons

    assert ag.mask("문의: hong@corp.kr, 010-1234-5678, 900101-1234567") == \
        "문의: [이메일], [전화번호], [주민번호]", ag.mask("hong@corp.kr")
    # AI 전송 payload에도 마스킹이 적용되어야 한다
    uf2 = ag.read_file(path)
    ag.review_stage1(uf2, rules={})
    items = ag._ai_items(uf2)
    assert all("@" not in json.dumps(it, ensure_ascii=False) for it in items), items
    print("  ✓ 작성기준 적용 + 개인정보 마스킹(전송 payload 포함)")


def test_synthesis_modes(tmp: Path):
    """모드 A/B/C/D 결과 시트 구성 (§9)."""
    a = tmp / "기획부.xlsx"
    b = tmp / "총무부.xlsx"
    for path, dept in ((a, "기획부"), (b, "총무부")):
        make_book(path, {
            "예산": [["사번", "부서", "예산액"], [f"{dept[0]}1", dept, 100], [f"{dept[0]}2", dept, 200]],
            "인원": [["사번", "부서", "인원수"], [f"{dept[0]}1", dept, 3]],
        })
    files = []
    for path in (a, b):
        uf = ag.read_file(path)
        ag.review_stage1(uf, rules={})
        files.append(uf)

    wb, _ = ag.synthesize(files, "A", {}, [])
    assert len(wb.sheetnames) == 4, wb.sheetnames                 # n×m = 2×2
    assert "기획부_예산" in wb.sheetnames, wb.sheetnames

    wb, _ = ag.synthesize(files, "B", {}, [])
    assert sorted(wb.sheetnames) == ["예산", "인원"], wb.sheetnames  # m = 2
    ws = wb["예산"]
    assert ws.cell(row=1, column=1).value == "부서", "부서 구분 컬럼이 추가돼야 한다"
    assert ws.max_row == 5, ws.max_row                            # 헤더 1 + (2+2) 행

    wb, notes = ag.synthesize(files, "C", {"예산": "재정"}, [])
    assert sorted(wb.sheetnames) == ["미분류", "재정"], wb.sheetnames  # k = 2, 인원→미분류
    assert any("미분류" in n for n in notes), notes
    assert wb["재정"].cell(row=1, column=1).value == "구분", "원 시트명 컬럼이 추가돼야 한다"

    wb, _ = ag.synthesize(files, "D", {}, ["예산액"])
    assert wb.sheetnames[0] == "종합요약", wb.sheetnames           # m+1, 최상단
    formula = wb["종합요약"].cell(row=2, column=2).value
    assert isinstance(formula, str) and formula.startswith("=SUMIF"), formula  # 하드코딩 금지
    print("  ✓ 합성 모드 A/B/C/D 시트 구성 + 모드 D 수식 기반 요약")


def test_cli_end_to_end(tmp: Path):
    """CLI 전체 경로: 이상 파일은 기본 제외되고 결과·리포트가 생성된다."""
    work = tmp / "e2e"
    work.mkdir()
    make_book(work / "기획부.xlsx", {"예산": [
        ["사번", "부서", "예산액"], ["A1", "기획부", 100], ["A2", "기획부", 200]]})
    make_book(work / "총무부.xlsx", {"예산": [
        ["사번", "부서", "예산액"], ["B1", "총무부", None]]})     # 누락 → 오류 → 기본 제외
    out, report = tmp / "merged.xlsx", tmp / "report.xlsx"
    code = ag.main([str(work), "--mode", "B", "--no-ai",
                    "--out", str(out), "--report", str(report)])
    assert code == 0, code
    assert out.exists() and report.exists()
    ws = openpyxl.load_workbook(out)["예산"]
    depts = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    assert depts == {"기획부"}, f"오류 파일은 기본 제외돼야 한다: {depts}"
    rep = openpyxl.load_workbook(report)
    assert rep.sheetnames == ["오류 목록", "자동교정 이력"], rep.sheetnames
    assert rep["오류 목록"].max_row >= 2, "누락 오류가 리포트에 기록돼야 한다"

    # 강제 포함 시에는 오류 파일도 취합에 들어간다 (§7, §10.1)
    code = ag.main([str(work), "--mode", "B", "--no-ai", "--include-anomalous",
                    "--out", str(out), "--report", str(report)])
    assert code == 0
    ws = openpyxl.load_workbook(out)["예산"]
    depts = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}
    assert depts == {"기획부", "총무부"}, depts
    print("  ✓ CLI end-to-end(기본 제외 / 강제 포함)")


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="agg-test-"))
    try:
        for fn in (test_structure_and_stage1, test_clean_file_and_gating,
                   test_rules_and_masking, test_synthesis_modes, test_cli_end_to_end):
            sub = tmp / fn.__name__
            sub.mkdir()
            fn(sub)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
