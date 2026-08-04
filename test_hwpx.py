"""hwpx 병합 회귀 테스트.

실제 업무 파일(sampledata/hwpx)은 커밋할 수 없으므로, 그 파일에서 확인한 구조를 그대로
재현한 최소 fixture로 고정한다. 재현하는 것:

- 두 문서가 같은 id를 쓴다(charPr·paraPr·style 전부 0부터)
- borderFill의 id는 0이 아니라 1부터 시작한다 → 오프셋 산술이 아니라 매핑이어야 맞는다
- charPrIDRef="4294967295"('없음' sentinel)는 id가 아니라서 옮기면 안 된다
- font id 공간이 언어별로 나뉘어 있다(hh:fontRef의 hangul·latin…)
- BinData 이름이 두 문서에서 겹친다(image1)
- header에 정의가 없는 참조(linkListIDRef)는 손대면 없는 곳을 가리키게 된다
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import hwpx_merge as H

HEADER = """<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" \
xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" secCnt="1">\
<hh:beginNum page="1"/><hh:refList>\
<hh:fontfaces itemCnt="7">\
<hh:fontface lang="HANGUL" fontCnt="2">\
<hh:font id="0" face="굴림" type="TTF"/><hh:font id="1" face="{f}" type="TTF"/>\
</hh:fontface>\
<hh:fontface lang="LATIN" fontCnt="1"><hh:font id="0" face="Arial" type="TTF"/></hh:fontface>\
<hh:fontface lang="HANJA" fontCnt="1"><hh:font id="0" face="굴림" type="TTF"/></hh:fontface>\
<hh:fontface lang="JAPANESE" fontCnt="1"><hh:font id="0" face="굴림" type="TTF"/></hh:fontface>\
<hh:fontface lang="OTHER" fontCnt="1"><hh:font id="0" face="굴림" type="TTF"/></hh:fontface>\
<hh:fontface lang="SYMBOL" fontCnt="1"><hh:font id="0" face="굴림" type="TTF"/></hh:fontface>\
<hh:fontface lang="USER" fontCnt="1"><hh:font id="0" face="굴림" type="TTF"/></hh:fontface>\
</hh:fontfaces>\
<hh:borderFills itemCnt="2">\
<hh:borderFill id="1" threeD="0"/><hh:borderFill id="2" threeD="0"/>\
</hh:borderFills>\
<hh:charProperties itemCnt="2">\
<hh:charPr id="0" height="1000" borderFillIDRef="1">\
<hh:fontRef hangul="1" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/></hh:charPr>\
<hh:charPr id="1" height="1200" borderFillIDRef="2">\
<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/></hh:charPr>\
</hh:charProperties>\
<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0"/></hh:tabProperties>\
<hh:numberings itemCnt="1"><hh:numbering id="1" start="0">\
<hh:paraHead level="1" charPrIDRef="4294967295">^1.</hh:paraHead></hh:numbering></hh:numberings>\
<hh:paraProperties itemCnt="2">\
<hh:paraPr id="0" tabPrIDRef="0"><hh:heading type="NONE" idRef="0" level="0"/></hh:paraPr>\
<hh:paraPr id="1" tabPrIDRef="0"><hh:heading type="NUMBER" idRef="1" level="0"/></hh:paraPr>\
</hh:paraProperties>\
<hh:styles itemCnt="2">\
<hh:style id="0" type="PARA" name="바탕글" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0"/>\
<hh:style id="1" type="PARA" name="개요" paraPrIDRef="1" charPrIDRef="1" nextStyleIDRef="1"/>\
</hh:styles>\
</hh:refList></hh:head>"""

SECTION = """<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" \
xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\
<hp:p paraPrIDRef="1" styleIDRef="1"><hp:run charPrIDRef="1"><hp:t>{t}</hp:t></hp:run></hp:p>\
<hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0">\
<hp:t>{t}-본문</hp:t><hp:pic><hp:img binaryItemIDRef="image1"/></hp:pic></hp:run>\
<hp:linesegarray linkListIDRef="0" linkListNextIDRef="0"/></hp:p>\
<hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="4294967295">\
<hp:t>sentinel</hp:t></hp:run></hp:p>\
</hs:sec>"""

HPF = """<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\
<opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:metadata><opf:title/>\
</opf:metadata><opf:manifest>\
<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>\
<opf:item id="image1" href="BinData/image1.png" media-type="image/png" isEmbeded="1"/>\
<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>\
</opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/>\
<opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>"""


def _fixture(path: Path, tag: str, image: bytes) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(zipfile.ZipInfo("mimetype"), H.MIMETYPE, zipfile.ZIP_STORED)
        z.writestr("version.xml", '<?xml version="1.0"?><hv:HCFVersion major="5"/>')
        z.writestr("settings.xml", '<?xml version="1.0"?><ha:HWPApplicationSetting name="%s"/>' % tag)
        z.writestr("META-INF/container.xml", '<?xml version="1.0"?><ocf:container/>')
        z.writestr("Contents/header.xml", HEADER.format(f="돋움" if tag == "A" else "맑은고딕"))
        z.writestr("Contents/section0.xml", SECTION.format(t=tag))
        z.writestr("Contents/content.hpf", HPF)
        z.writestr("BinData/image1.png", image)


def test_merge(tmp: Path) -> None:
    a, b, out = tmp / "a.hwpx", tmp / "b.hwpx", tmp / "merged.hwpx"
    _fixture(a, "A", b"\x89PNG-A")
    _fixture(b, "B", b"\x89PNG-B")

    r = H.merge([a, b], out)
    assert r["결과"] == "성공", r["dangling"]
    assert not r["dangling"]
    assert not r["itemCnt불일치"], r["itemCnt불일치"]

    # 수량 대조: 두 문서 몫이 그대로 살아 있다
    assert r["수량대조"]["문단"] == {"원본합계": 6, "병합결과": 6}, r["수량대조"]
    assert r["수량대조"]["그림"] == {"원본합계": 2, "병합결과": 2}, r["수량대조"]

    # ID 공간이 겹치지 않게 합쳐졌다
    assert r["정의수"]["charPr"] == 4, r["정의수"]
    assert r["정의수"]["style"] == 4, r["정의수"]
    assert r["정의수"]["borderFill"] == 4, r["정의수"]
    assert r["정의수"]["font"] == {"HANGUL": 4, "LATIN": 2, "HANJA": 2, "JAPANESE": 2,
                                  "OTHER": 2, "SYMBOL": 2, "USER": 2}, r["정의수"]["font"]
    print("  ✓ ID 공간 병합(정의 수·해석 불가 참조 0건·수량 대조)")

    z = zipfile.ZipFile(out)
    header = z.read("Contents/header.xml").decode()
    s0 = z.read("Contents/section0.xml").decode()
    s1 = z.read("Contents/section1.xml").decode()

    # 첫 문서는 손대지 않고, 두 번째 문서만 뒤 번호를 받는다
    assert 'charPrIDRef="1"' in s0 and 'styleIDRef="1"' in s0
    assert 'charPrIDRef="3"' in s1 and 'styleIDRef="3"' in s1, s1
    assert 'paraPrIDRef="3"' in s1, s1
    # borderFill은 id 1부터 시작 → 두 번째 문서는 3·4를 받아야 한다(오프셋 산술이면 3·4가 아니다)
    assert 'id="3" threeD' in header and 'id="4" threeD' in header, header
    print("  ✓ 첫 문서 id 보존 + 뒤 문서 재번호(borderFill이 1부터 시작해도 맞음)")

    # '없음' sentinel은 옮기지 않는다
    assert s0.count('charPrIDRef="4294967295"') == 1
    assert s1.count('charPrIDRef="4294967295"') == 1, s1
    assert header.count('charPrIDRef="4294967295"') == 2, header
    print("  ✓ '없음' sentinel(4294967295) 미변경")

    # 언어별 font 공간 — 두 번째 문서의 hangul=1 은 3이 되고 latin=0 은 1이 된다
    assert 'hangul="3"' in header and 'latin="1"' in header, header
    assert 'face="맑은고딕"' in header, "두 번째 문서 글꼴이 누락됨"
    print("  ✓ 언어별 font 공간 분리 재매핑")

    # heading idRef는 NUMBER일 때만 옮긴다
    assert '<hh:heading type="NUMBER" idRef="2"' in header, header
    assert '<hh:heading type="NONE" idRef="0"' in header, header
    print("  ✓ heading idRef는 NUMBER/BULLET만 이동")

    # BinData 이름 충돌 회피 + 참조 동시 갱신
    assert z.read("BinData/image1.png") == b"\x89PNG-A"
    assert z.read("BinData/f2_image1.png") == b"\x89PNG-B"
    assert 'binaryItemIDRef="image1"' in s0 and 'binaryItemIDRef="f2_image1"' in s1
    assert 'href="BinData/f2_image1.png"' in z.read("Contents/content.hpf").decode()
    print("  ✓ BinData 충돌 회피 + 참조·manifest 동시 갱신")

    # header에 정의가 없는 참조는 그대로 둔다
    assert 'linkListIDRef="0"' in s1, s1
    assert r["미대응"] == {"linkListIDRef": 2, "linkListNextIDRef": 2}, r["미대응"]
    print("  ✓ 정의 없는 참조(linkList) 미변경 + 건수 보고")

    # hh:head secCnt — 안 고치면 뷰어가 첫 구역만 보여준다(실측으로 잡힌 결함)
    assert 'secCnt="2"' in header, header[:400]

    # 패키지 골격
    first = z.infolist()[0]
    assert first.filename == "mimetype" and first.compress_type == zipfile.ZIP_STORED
    hpf = z.read("Contents/content.hpf").decode()
    assert 'id="section1"' in hpf and '<opf:itemref idref="section1"' in hpf
    print("  ✓ mimetype 무압축 첫 엔트리 + hpf manifest·spine 재생성")

    # 조용히 버려지는 것은 리포트에 남는다
    assert any("settings.xml" in l for l in r["한계"]), r["한계"]
    print("  ✓ 첫 파일만 유지되는 문서 설정을 한계로 보고")


def test_rejects(tmp: Path) -> None:
    bad = tmp / "not.hwpx"
    bad.write_bytes(b"not a zip")
    try:
        H.merge([bad, bad], tmp / "x.hwpx")
        raise AssertionError("zip이 아닌 입력을 받아들였다")
    except H.MergeError as e:
        assert "hwp" in str(e)

    xlsx = tmp / "sheet.hwpx"
    with zipfile.ZipFile(xlsx, "w") as z:
        z.writestr("mimetype", b"application/vnd.ms-excel")
    try:
        H.merge([xlsx, xlsx], tmp / "x.hwpx")
        raise AssertionError("hwpx가 아닌 입력을 받아들였다")
    except H.MergeError as e:
        assert "hwpx가 아닙니다" in str(e)

    a = tmp / "a.hwpx"
    _fixture(a, "A", b"\x89PNG")
    try:
        H.merge([a], tmp / "x.hwpx")
        raise AssertionError("1개 입력을 받아들였다")
    except H.MergeError:
        pass
    print("  ✓ 입력 거부(zip 아님 / hwpx 아님 / 1개)")


def test_dangling_blocks_output(tmp: Path) -> None:
    """해석 불가 참조가 있으면 결과 파일을 쓰지 않는다."""
    a, b, out = tmp / "a.hwpx", tmp / "b.hwpx", tmp / "merged.hwpx"
    _fixture(a, "A", b"\x89PNG-A")
    _fixture(b, "B", b"\x89PNG-B")
    # 두 번째 문서 본문이 없는 style을 가리키게 만든다
    data = {n: zipfile.ZipFile(b).read(n) for n in zipfile.ZipFile(b).namelist()}
    data["Contents/section0.xml"] = data["Contents/section0.xml"].replace(
        b'styleIDRef="1"', b'styleIDRef="99"')
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(zipfile.ZipInfo("mimetype"), H.MIMETYPE, zipfile.ZIP_STORED)
        for n, d in data.items():
            if n != "mimetype":
                z.writestr(n, d)

    r = H.merge([a, b], out)
    assert r["결과"] == "실패", r
    assert any("styleIDRef=99" in d for d in r["dangling"]), r["dangling"]
    assert not out.exists(), "실패인데 결과 파일을 썼다"
    print("  ✓ 해석 불가 참조 시 결과 파일 미작성")


# 실제 파일로 확인하지 못한 계열 — 사용자가 예시를 주지 못해 구조를 재현해서 본다.
# ① 원래 여러 구역인 문서 ② 머리말·꼬리말이 있는 문서 ③ 우리가 모르는 id 공간(메모 등)

SECTION_WITH_HEADER_FOOTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" \
xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\
<hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0">\
<hp:ctrl><hp:colPr/></hp:ctrl>\
<hp:secPr><hp:pagePr landscape="WIDELY"/>\
<hp:header applyPageType="BOTH"><hp:subList>\
<hp:p paraPrIDRef="1" styleIDRef="1"><hp:run charPrIDRef="1"><hp:t>{t} 머리말</hp:t></hp:run></hp:p>\
</hp:subList></hp:header>\
<hp:footer applyPageType="BOTH"><hp:subList>\
<hp:p paraPrIDRef="1" styleIDRef="1"><hp:run charPrIDRef="1"><hp:t>{t} 꼬리말</hp:t></hp:run></hp:p>\
</hp:subList></hp:footer>\
</hp:secPr><hp:t>{t}-본문</hp:t></hp:run></hp:p>\
</hs:sec>"""

HPF_TWO_SECTIONS = HPF.replace(
    '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>',
    '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
    '<opf:item id="section1" href="Contents/section1.xml" media-type="application/xml"/>'
).replace(
    '<opf:itemref idref="section0" linear="yes"/>',
    '<opf:itemref idref="section0" linear="yes"/><opf:itemref idref="section1" linear="yes"/>')


def _fixture2(path: Path, tag: str, *, sections: int = 1,
              header_footer: bool = False, extra_box: str = "") -> None:
    """구역 수·머리말·모르는 id 공간을 골라 만드는 fixture."""
    header = HEADER.format(f="돋움" if tag == "A" else "맑은고딕")
    if extra_box:
        header = header.replace("</hh:refList>", extra_box + "</hh:refList>")
    if sections > 1:
        header = header.replace('secCnt="1"', 'secCnt="%d"' % sections)
    body = SECTION_WITH_HEADER_FOOTER if header_footer else SECTION
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(zipfile.ZipInfo("mimetype"), H.MIMETYPE, zipfile.ZIP_STORED)
        z.writestr("version.xml", '<?xml version="1.0"?><hv:HCFVersion major="5"/>')
        z.writestr("settings.xml", '<?xml version="1.0"?><ha:HWPApplicationSetting/>')
        z.writestr("Contents/header.xml", header)
        for i in range(sections):
            z.writestr("Contents/section%d.xml" % i, body.format(t="%s-%d" % (tag, i)))
        z.writestr("Contents/content.hpf", HPF_TWO_SECTIONS if sections > 1 else HPF)
        z.writestr("BinData/image1.png", b"\x89PNG-" + tag.encode())


def test_multi_section_and_header_footer(tmp: Path) -> None:
    """원래 여러 구역인 문서와 머리말·꼬리말이 있는 문서.

    실제 파일로 확인하지 못한 계열이다(감사 처분요구서 3개는 모두 1구역·머리말 없음).
    구역은 파일 경계가 아니라 **구역 수만큼** 늘어나야 하고, 머리말·꼬리말 안의 서식
    참조도 본문과 똑같이 옮겨져야 한다 — 안 옮기면 머리말만 남의 서식으로 렌더된다.
    """
    a, b, out = tmp / "a.hwpx", tmp / "b.hwpx", tmp / "merged.hwpx"
    _fixture2(a, "A", sections=2, header_footer=True)      # 2구역 + 머리말·꼬리말
    _fixture2(b, "B", sections=1, header_footer=True)

    r = H.merge([a, b], out)
    assert r["결과"] == "성공", (r["dangling"], r["itemCnt불일치"], r["옮길수없는서식"])
    # 구역은 파일 수(2)가 아니라 구역 수(2+1)만큼이어야 한다
    assert r["구역"] == {"a.hwpx": 2, "b.hwpx": 1}, r["구역"]
    z = zipfile.ZipFile(out)
    assert "Contents/section2.xml" in z.namelist(), z.namelist()
    assert 'secCnt="3"' in z.read("Contents/header.xml").decode()
    hpf = z.read("Contents/content.hpf").decode()
    assert '<opf:itemref idref="section2"' in hpf, hpf

    # 머리말·꼬리말이 살아 있고, 그 안의 참조가 뒤 문서 몫으로 옮겨졌다
    s2 = z.read("Contents/section2.xml").decode()
    assert "B-0 머리말" in s2 and "B-0 꼬리말" in s2, s2
    assert 'charPrIDRef="3"' in s2 and 'paraPrIDRef="3"' in s2, s2
    s0 = z.read("Contents/section0.xml").decode()
    assert "A-0 머리말" in s0 and 'charPrIDRef="1"' in s0, s0
    print("  ✓ 여러 구역 문서 + 머리말·꼬리말 서식 참조 이동")

    # 15개 규모 — 실제 파일로는 3개까지만 돌려봤다. id 공간과 BinData 접두사가
    # 파일 수만큼 늘어나도 참조가 어긋나지 않아야 한다
    many = []
    for i in range(15):
        p = tmp / f"m{i}.hwpx"
        _fixture2(p, "M%d" % i)
        many.append(p)
    big = tmp / "big.hwpx"
    r = H.merge(many, big)
    assert r["결과"] == "성공", (r["dangling"][:5], r["옮길수없는서식"])
    assert sum(r["구역"].values()) == 15, r["구역"]
    zb = zipfile.ZipFile(big)
    assert 'secCnt="15"' in zb.read("Contents/header.xml").decode()
    assert r["정의수"]["charPr"] == 30, r["정의수"]          # 파일당 2개 × 15
    assert len([n for n in zb.namelist() if n.startswith("BinData/")]) == 15, zb.namelist()
    assert r["수량대조"]["문단"]["원본합계"] == r["수량대조"]["문단"]["병합결과"], r["수량대조"]
    print("  ✓ 15개 규모 병합(구역 15·id 공간 30·BinData 15·참조 어긋남 0)")


def test_unmergeable_id_space_blocks_output(tmp: Path) -> None:
    """우리가 옮길 수 없는 id 공간이 있으면 조용히 섞지 않고 실패한다.

    메모·개체가 든 문서는 header에 그 공간을 정의한다. 실측한 3개 파일에는 정의가 없어
    참조를 그대로 뒀는데(`UNTOUCHED`), 정의가 **있는** 문서를 그대로 병합하면 2번째
    문서의 참조가 첫 문서의 정의를 가리켜 서식이 조용히 바뀐다. 예시 파일이 없어
    구조를 재현해 확인한다.
    """
    memo = ('<hh:memoProperties itemCnt="1">'
            '<hh:memoPr id="0" width="1000" lineType="SOLID"/></hh:memoProperties>')
    a, b, out = tmp / "a.hwpx", tmp / "b.hwpx", tmp / "merged.hwpx"
    _fixture2(a, "A", extra_box=memo)
    _fixture2(b, "B", extra_box=memo)

    r = H.merge([a, b], out)
    assert r["결과"] == "실패", r
    assert any("memoProperties" in x for x in r["옮길수없는서식"]), r["옮길수없는서식"]
    assert not out.exists(), "실패인데 결과 파일을 썼다"

    # 첫 문서만 갖고 있으면 문제가 없다 — 그 정의는 그대로 남는다
    _fixture2(b, "B")
    r = H.merge([a, b], out)
    assert r["결과"] == "성공", (r["옮길수없는서식"], r["dangling"])
    assert not r["옮길수없는서식"], r["옮길수없는서식"]
    print("  ✓ 옮길 수 없는 id 공간(메모 등) 감지 → 실패, 첫 문서만이면 통과")


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="hwpx-test-"))
    try:
        for fn in (test_merge, test_rejects, test_dangling_blocks_output,
                   test_multi_section_and_header_footer,
                   test_unmergeable_id_space_blocks_output):
            sub = tmp / fn.__name__
            sub.mkdir()
            fn(sub)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
