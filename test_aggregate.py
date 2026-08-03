#!/usr/bin/env python3
"""aggregate.py 자체 점검. 실행: python3 test_aggregate.py (프레임워크 없음)"""

import io
import json
import threading
import time
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.units import EMU_to_pixels
from PIL import Image as PILImage

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


def test_image_anchor_relocation(tmp: Path):
    """합성 시 이미지가 원래 붙어 있던 '그 행 · 그 컬럼'을 따라가야 한다 (§9)."""
    def make_with_image(path: Path, dept: str):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "실적"
        ws["A1"] = f"{dept} 실적 보고"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
        for c, h in enumerate(["사번", "부서", "예산액", "증빙사진"], start=1):
            ws.cell(row=3, column=c, value=h)
        for i in range(2):
            ws.cell(row=4 + i, column=1, value=f"{dept[0]}{i + 1}")
            ws.cell(row=4 + i, column=2, value=dept)
            ws.cell(row=4 + i, column=3, value=100 * (i + 1))
        buf = io.BytesIO()
        PILImage.new("RGB", (300, 300), (10, 20, 30)).save(buf, format="PNG")
        buf.seek(0)
        picture = XLImage(buf)
        picture.width, picture.height = 96, 72      # 엑셀에서 셀 크기에 맞춰 줄여 놓은 상태
        ws.add_image(picture, "D4")                 # 데이터 첫 행의 증빙사진 칸
        wb.save(path)

    def placed(wb, tag):
        """anchor는 저장 시점에 확정되므로 저장 후 다시 읽어 실제 위치를 본다."""
        out = tmp / f"{tag}.xlsx"
        wb.save(out)
        found = {}
        for ws in openpyxl.load_workbook(out).worksheets:
            for im in getattr(ws, "_images", []):
                frm = im.anchor._from
                size = getattr(im.anchor, "ext", None)
                found.setdefault(ws.title, []).append((
                    frm.row + 1, frm.col + 1,
                    (EMU_to_pixels(size.cx), EMU_to_pixels(size.cy)) if size else None))
        return found

    files = []
    for dept in ("기획부", "총무부"):
        path = tmp / f"{dept}.xlsx"
        make_with_image(path, dept)
        uf = ag.read_file(path)
        ag.review_stage1(uf, rules={})
        files.append(uf)
    assert files[0].sheets[0].images, "원본 이미지를 읽어야 한다"

    # 모드 A: 헤더가 1행으로 당겨지므로 데이터 첫 행은 2행. 컬럼은 그대로 D(4).
    wb, _ = ag.synthesize(files, "A", {}, [])
    got = placed(wb, "A")
    assert got["기획부_실적"] == [(2, 4, (96, 72))], got

    # 모드 B: '부서' 컬럼이 앞에 끼므로 증빙사진은 D→E(5). 행은 파일별 시작 행.
    wb, _ = ag.synthesize(files, "B", {}, [])
    assert placed(wb, "B")["실적"] == [(2, 5, (96, 72)), (4, 5, (96, 72))], placed(wb, "B")

    # 모드 C: '구분'+'부서' 2개가 끼므로 D→F(6).
    wb, _ = ag.synthesize(files, "C", {"실적": "상반기"}, [])
    assert placed(wb, "C")["상반기"] == [(2, 6, (96, 72)), (4, 6, (96, 72))], placed(wb, "C")
    print("  ✓ 이미지 anchor 재배치(행·열 추종 + 표시 크기 보존)")


def test_real_form_structure(tmp: Path):
    """실제 업무 양식 구조에서 오탐이 없어야 한다.

    실제 회신 파일(지사별 자체점검 리스트)에서 확인된 네 가지를 그대로 재현한다.
    이 구조에서 종전 엔진은 8개 파일 전부를 '오류'로 판정해 아무것도 취합되지
    않았다.
      ① 첫 시트가 '작성 주의사항' 안내문 — 표가 아님
      ② 제목 행(병합) + 안내 행 뒤 4행에 헤더
      ③ 지사·개소가 세로 병합 — 파일에는 첫 칸만 값이 있다
      ④ 표 끝에 '합 계' 행과 '*' 각주 행
    """
    path = tmp / "판교지사.xlsx"
    wb = openpyxl.Workbook()

    guide = wb.active
    guide.title = "작성 주의사항"
    guide["A1"] = "집중안전점검 사업장 자체점검 작성 안내"
    guide.merge_cells("A1:L1")
    for r, text in enumerate(["ㅇ 해당지사 : 전 지사", "ㅇ 주의사항", "1. 매달 취합합니다"], start=3):
        guide.cell(row=r, column=1, value=text)

    ws = wb.create_sheet("2025년 대상 리스트")
    ws["A1"] = "2025년 집중안전점검 지열지점 사업장 자체점검 리스트"
    ws.merge_cells("A1:H1")
    ws["A3"] = "  ※ 2024.11.01.~2025.04.30. 기간 기준"
    for c, h in enumerate(["지사", "개소", "위치", "관경(A)", "준공연도", "온도차(℃)", "보수여부"], start=1):
        ws.cell(row=4, column=c, value=h)
    ws["A5"], ws["B5"] = "판교지사", 2
    ws.merge_cells("A5:A7")          # 지사가 세 행에 걸쳐 병합
    ws.merge_cells("B5:B7")
    for i, (loc, dia, year, temp, fix) in enumerate([
        ("백현동 555-2", 200, 2015, 24.7, "보수완료"),
        ("삼평동 672-6", 150, 2015, 3.1, "보수완료"),
        ("판교동 12-3", 300, 2016, 5.5, "보수중"),
    ]):
        ws.cell(row=5 + i, column=3, value=loc)
        ws.cell(row=5 + i, column=4, value=dia)
        ws.cell(row=5 + i, column=5, value=year)
        ws.cell(row=5 + i, column=6, value=temp)
        ws.cell(row=5 + i, column=7, value=fix)
    ws["A9"], ws["B9"] = "합 계", 2                       # 합계 행
    ws["A10"] = " *  해당기간 내 점검 완료분만 기재"        # 각주 행
    # 양식은 아래쪽까지 미리 병합해 둔다 — 여기에 값이 새면 빈 행이 데이터로 살아난다
    ws.merge_cells("A12:A30")
    ws.merge_cells("B12:B30")
    wb.save(path)

    uf = ag.read_file(path)
    assert uf.readable, uf.issues
    # 안내 시트는 취합 대상에서 빠지고, 데이터 시트만 남는다
    assert [s.name for s in uf.sheets] == ["2025년 대상 리스트"], [s.name for s in uf.sheets]
    sheet = uf.sheets[0]
    assert sheet.header_row == 4, sheet.header_row
    assert sheet.headers[:3] == ["지사", "개소", "위치"], sheet.headers
    # 합계·각주 행은 제외, 미리 병합해 둔 빈 행은 되살아나지 않는다
    assert sheet.row_numbers == [5, 6, 7], sheet.row_numbers
    # 세로 병합된 지사·개소가 세 행 모두 채워진다
    assert [r[0] for r in sheet.rows] == ["판교지사"] * 3, [r[0] for r in sheet.rows]
    assert [r[1] for r in sheet.rows] == [2, 2, 2], [r[1] for r in sheet.rows]

    ag.review_stage1(uf, rules={})
    errors = [i for i in uf.issues if i.grade == ag.ERROR]
    assert not errors, [i.line() for i in errors]
    assert uf.grade == ag.WARN, [i.line() for i in uf.issues]
    # 등급이 오류가 아니므로 취합에 기본 포함된다 (§7)
    reasons = " ".join(i.reason for i in uf.issues)
    assert "표 구조가 아니어서" in reasons and "합계·각주" in reasons, reasons

    # 첫 컬럼(지사)이 반복되어도 키 충돌로 오탐하지 않는다 — 고유한 '위치'를 키로 잡는다
    assert ag._pick_key_column(sheet.headers, sheet.rows) == 2, sheet.headers

    # 합계 행이 취합에 섞이면 모드 B/D에서 이중 계상된다
    result, _ = ag.synthesize([uf], "B", {}, [])
    body = result["2025년 대상 리스트"]
    assert body.max_row == 4, body.max_row          # 헤더 1 + 데이터 3
    assert all(str(body.cell(row=r, column=2).value).replace(" ", "") not in ("합계", "계")
               for r in range(2, body.max_row + 1))
    print("  ✓ 실제 업무 양식 구조(안내 시트·세로 병합·합계·각주 행)")


def test_stage2_parallel(tmp: Path):
    """2단계는 파일 간 병렬로 돌고, 한 파일의 실패가 다른 파일을 오염시키지 않는다.

    실제 API를 부르지 않는다 — _ai_call을 지연·실패가 흉내나는 스텁으로 바꿔
    호출 스케줄만 본다(네트워크·비용 없이 결정적으로 검증).
    """
    files = []
    for i in range(6):
        path = tmp / f"제{i}부서.xlsx"
        make_book(path, {"예산": [
            ["사번", "사업내용", "예산액"],
            [f"{i}-1", "교육 운영", 100],
            [f"{i}-2", "비품 구매", 200],
        ]})
        uf = ag.read_file(path)
        ag.review_stage1(uf, rules={})
        assert not uf.issues, [x.line() for x in uf.issues]   # 전부 게이팅 통과
        files.append(uf)

    DELAY = 0.15
    calls, peak, active = [], [0], [0]
    lock = threading.Lock()

    def fake_call(client, model, items):
        with lock:
            active[0] += 1
            peak[0] = max(peak[0], active[0])
            calls.append(items[0]["id"])
        time.sleep(DELAY)
        with lock:
            active[0] -= 1
        # '제3부서'는 서비스 장애를 흉내낸다
        if any("제3부서" in str(it.get("작성기준", "")) for it in items):
            raise RuntimeError("모의 장애")
        return {it["id"]: {"verdict": "부적합", "reason": "모의 판정"} for it in items}

    real_call, real_client = ag._ai_call, ag._ai_client
    ag._ai_call = fake_call
    ag._ai_client = lambda: object()          # 클라이언트 생성은 성공한 것으로 둔다
    try:
        # 파일명이 payload에 남지 않으므로, 장애 대상 식별용으로 작성기준에 심는다
        for uf in files:
            for sheet in uf.sheets:
                sheet.headers = [f"{h}({uf.dept})" for h in sheet.headers]
                sheet.col_types = []

        t0 = time.perf_counter()
        ag.review_stage2_many(files, model="stub", workers=6)
        elapsed = time.perf_counter() - t0
    finally:
        ag._ai_call, ag._ai_client = real_call, real_client

    assert len(calls) == 6, calls
    assert peak[0] > 1, f"동시에 호출되지 않았다(최대 동시 {peak[0]})"
    # 6개를 순차로 돌면 6*DELAY, 병렬이면 그보다 확실히 짧다
    assert elapsed < 6 * DELAY * 0.7, f"병렬 이득이 없다: {elapsed:.2f}s"

    failed = [uf for uf in files if uf.ai_unverified]
    assert [uf.dept for uf in failed] == ["제3부서"], [uf.dept for uf in failed]
    assert not failed[0].issues, "장애 파일은 2단계 판정을 남기지 않는다"
    for uf in files:
        if uf.ai_unverified:
            continue
        # 정상 파일은 부적합 판정이 경고 등급으로 붙는다 (§6.2)
        assert uf.issues and all(i.grade == ag.WARN and i.stage == "2단계" for i in uf.issues), \
            [i.line() for i in uf.issues]
    print("  ✓ 2단계 병렬 호출 + 파일 단위 실패 격리")


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
                   test_rules_and_masking, test_synthesis_modes,
                   test_image_anchor_relocation, test_real_form_structure,
                   test_stage2_parallel,
                   test_cli_end_to_end):
            sub = tmp / fn.__name__
            sub.mkdir()
            fn(sub)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
