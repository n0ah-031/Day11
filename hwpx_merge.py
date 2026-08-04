"""hwpx 병합 (F3) — OWPML 직접 병합.

hwpx는 ZIP + XML(OWPML)이라 표준 라이브러리만으로 병합된다. 어려운 곳은 XML을 이어
붙이는 게 아니라 **ID 공간**이다: 모든 문서의 header.xml이 style·charPr·paraPr 등을
자기 문서 안에서만 유효한 번호로 매기기 때문에 그대로 합치면 서식 참조가 뒤섞인다.
그래서 문서별로 `{옛 id: 새 id}` 매핑을 만들어 항목의 id와 본문의 모든 참조를 함께 옮긴다.

매핑을 쓰는 이유(오프셋 산술이 아니라): 실측 파일의 borderFill은 id 0이 아니라 1부터
시작한다. "앞 문서의 항목 수만큼 더한다"는 방식은 id가 빈틈없이 같은 번호에서 시작할 때만
맞다. 매핑은 그 가정이 없고, 매핑에 없는 값(예: '없음'을 뜻하는 4294967295, 기본값 0)은
건드리지 않고 통과시킨다.

XML은 ElementTree가 아니라 정규식으로 다룬다. hwpx는 접두사(hh:·hp:·hs:…)가 많고
ElementTree로 다시 직렬화하면 접두사·빈 요소 표기가 바뀌어 한글이 못 읽을 위험이 있다.
읽은 문자열을 최소한만 고쳐 그대로 내보내는 편이 안전하고 짧다.
"""
from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path

MIMETYPE = b"application/hwp+zip"
NONE_ID = 4294967295  # OWPML의 '없음' sentinel(0xFFFFFFFF). id가 아니므로 옮기지 않는다.

# refList 컨테이너 → 항목 태그
SPACES = {
    "borderFills": "borderFill",
    "charProperties": "charPr",
    "tabProperties": "tabPr",
    "numberings": "numbering",
    "bullets": "bullet",
    "paraProperties": "paraPr",
    "styles": "style",
}

# 참조 속성 → 가리키는 공간
REFS = {
    "charPrIDRef": "charPr",
    "paraPrIDRef": "paraPr",
    "styleIDRef": "style",
    "nextStyleIDRef": "style",
    "borderFillIDRef": "borderFill",
    "tabPrIDRef": "tabPr",
    "numberingIDRef": "numbering",
    "bulletIDRef": "bullet",
}

# header에 정의가 없어(실측 3파일 확인) 손대지 않는 참조. 옮기면 없는 id를 가리킨다.
UNTOUCHED = ("outlineShapeIDRef", "memoShapeIDRef", "linkListIDRef", "linkListNextIDRef")

FONT_LANGS = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")

MEDIA = {".jpg": "image/jpg", ".jpeg": "image/jpg", ".png": "image/png",
         ".gif": "image/gif", ".bmp": "image/bmp", ".tif": "image/tif",
         ".tiff": "image/tif", ".wmf": "image/wmf", ".emf": "image/emf",
         ".ole": "application/x-ole-storage"}


class MergeError(Exception):
    """병합할 수 없는 입력."""


# --------------------------------------------------------------------- 읽기

def _read(path: Path) -> dict:
    if not zipfile.is_zipfile(path):
        raise MergeError(f"{path.name}: zip이 아닙니다. 구버전 .hwp(바이너리)는 지원하지 않습니다")
    z = zipfile.ZipFile(path)
    names = z.namelist()
    mt = z.read("mimetype") if "mimetype" in names else b""
    if mt != MIMETYPE:
        raise MergeError(f"{path.name}: hwpx가 아닙니다 (mimetype={mt!r})")
    for need in ("Contents/header.xml", "Contents/content.hpf"):
        if need not in names:
            raise MergeError(f"{path.name}: {need} 없음 — 손상된 파일입니다")
    secs = sorted((n for n in names if re.fullmatch(r"Contents/section\d+\.xml", n)),
                  key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    if not secs:
        raise MergeError(f"{path.name}: 본문(section)이 없습니다")
    return {
        "path": path,
        "header": z.read("Contents/header.xml").decode("utf-8"),
        "sections": [z.read(n).decode("utf-8") for n in secs],
        "hpf": z.read("Contents/content.hpf").decode("utf-8"),
        "bin": {n: z.read(n) for n in names if n.startswith("BinData/")},
        "extra": {n: z.read(n) for n in names
                  if not n.startswith(("Contents/", "BinData/")) and n != "mimetype"},
    }


# ------------------------------------------------------------------ ID 공간

def _inner(xml: str, tag: str) -> tuple[str, int, int] | None:
    """`<hh:tag ...>inner</hh:tag>` 의 inner와 그 범위. 자기닫음 컨테이너는 빈 inner."""
    m = re.search(r"<hh:%s\b[^>]*?/>" % tag, xml)
    if m:
        return "", m.start(), m.end()
    m = re.search(r"<hh:%s\b[^>]*?>(.*?)</hh:%s>" % (tag, tag), xml, re.S)
    return (m.group(1), m.start(), m.end()) if m else None


def _items(xml: str, tag: str) -> list[str]:
    """컨테이너 inner의 항목 XML 문자열들(항목은 같은 태그로 중첩되지 않는다)."""
    return re.findall(r"<hh:%s\b(?:[^>]*/>|[^>]*>.*?</hh:%s>)" % (tag, tag), xml, re.S)


def _own_id(item: str) -> int:
    return int(re.search(r'\bid="(\d+)"', item).group(1))


def _set_own_id(item: str, new: int) -> str:
    """항목 자신의 id만 교체(첫 id 속성. 하위 요소에는 id 속성이 없다)."""
    return re.sub(r'\bid="\d+"', 'id="%d"' % new, item, count=1)


def _box_items(header: str, box: str, tag: str) -> list[str]:
    found = _inner(header, box)
    return _items(found[0], tag) if found else []


def _fonts(header: str) -> dict[str, list[str]]:
    """언어별 font 항목. font id 공간은 전역이 아니라 언어별로 나뉜다."""
    got = {}
    block = _inner(header, "fontfaces")
    if block:
        for m in re.finditer(r'<hh:fontface\b[^>]*lang="(\w+)"[^>]*>(.*?)</hh:fontface>',
                             block[0], re.S):
            got[m.group(1)] = _items(m.group(2), "font")
    return got


def _assign(id_lists: list[list[int]]) -> list[dict[int, int]]:
    """문서별 id 목록 → 문서별 {옛 id: 새 id}. 첫 문서는 그대로 두고 뒤 문서만 이어 번호."""
    maps, nxt = [], 0
    for i, ids in enumerate(id_lists):
        if i == 0:
            maps.append({x: x for x in ids})
            nxt = max(ids, default=-1) + 1
        else:
            maps.append({x: nxt + k for k, x in enumerate(ids)})
            nxt += len(ids)
    return maps


def _build_maps(docs: list[dict]) -> list[dict]:
    """문서별 매핑 {공간: {옛:새}, 'font': {lang: {옛:새}}}."""
    maps = [{"font": {}} for _ in docs]
    for box, tag in SPACES.items():
        per = _assign([[_own_id(it) for it in _box_items(d["header"], box, tag)] for d in docs])
        for m, p in zip(maps, per):
            m[tag] = p
    fonts = [_fonts(d["header"]) for d in docs]
    for lang in FONT_LANGS:
        per = _assign([[_own_id(it) for it in f.get(lang, [])] for f in fonts])
        for m, p in zip(maps, per):
            m["font"][lang] = p
    return maps


# ------------------------------------------------------------------- 재매핑

def _remap(xml: str, maps: dict, binmap: dict[str, str]) -> str:
    """이 문서의 모든 참조를 새 id로 옮긴다. 매핑에 없는 값은 그대로 둔다."""

    def ref(m):
        attr, val = m.group(1), int(m.group(2))
        new = maps[REFS[attr]].get(val)
        return m.group(0) if new is None else '%s="%d"' % (attr, new)

    xml = re.sub(r"\b(%s)=\"(\d+)\"" % "|".join(REFS), ref, xml)

    def fontref(m):
        s = m.group(0)
        for lang in FONT_LANGS:
            fm = maps["font"][lang]
            s = re.sub(r'\b%s="(\d+)"' % lang.lower(),
                       lambda mm, fm=fm, lang=lang: (
                           mm.group(0) if int(mm.group(1)) not in fm
                           else '%s="%d"' % (lang.lower(), fm[int(mm.group(1))])), s)
        return s

    xml = re.sub(r"<hh:fontRef\b[^>]*/>", fontref, xml)

    def heading(m):
        # idRef는 type이 NUMBER/BULLET일 때만 numbering·bullet을 가리킨다.
        s = m.group(0)
        kind = (re.search(r'type="(\w+)"', s) or [None, ""])[1]
        space = {"NUMBER": "numbering", "BULLET": "bullet"}.get(kind)
        if not space:
            return s
        return re.sub(r'idRef="(\d+)"',
                      lambda mm: (mm.group(0) if int(mm.group(1)) not in maps[space]
                                  else 'idRef="%d"' % maps[space][int(mm.group(1))]), s)

    xml = re.sub(r"<hh:heading\b[^>]*/>", heading, xml)

    for old, new in binmap.items():
        xml = xml.replace('binaryItemIDRef="%s"' % old, 'binaryItemIDRef="%s"' % new)
    return xml


def _replace_box(header: str, box: str, items: list[str]) -> str:
    """컨테이너 내용을 병합 항목으로 갈아끼우고 itemCnt를 다시 센다."""
    found = _inner(header, box)
    if not found:
        return header
    _, start, end = found
    open_tag = re.match(r"<hh:%s\b[^>]*?/?>" % box, header[start:end]).group(0)
    open_tag = re.sub(r'itemCnt="\d+"', 'itemCnt="%d"' % len(items), open_tag)
    open_tag = open_tag.rstrip(">").rstrip("/") + ">"
    return header[:start] + open_tag + "".join(items) + "</hh:%s>" % box + header[end:]


# -------------------------------------------------------------------- 병합

def merge(paths: list[Path], out: Path, progress=None) -> dict:
    """hwpx 여러 개를 파일당 별도 구역으로 이어 붙인다. 리포트를 돌려준다.

    progress(phase, done, total)를 주면 단계를 알린다. 읽기가 파일 수에 비례하는
    유일한 구간이라 거기서만 파일별로 보고하고, 나머지는 단계 이름만 넘긴다.
    """
    if len(paths) < 2:
        raise MergeError("병합할 파일이 2개 이상 필요합니다")
    step = progress or (lambda *a: None)

    docs = []
    for i, p in enumerate(paths):
        step(f"{p.name} 읽는 중", i, len(paths))
        docs.append(_read(p))
    base = docs[0]
    step("서식 정보를 합치는 중", len(paths), len(paths))
    maps = _build_maps(docs)

    # BinData 이름 충돌 회피 — 2번째 문서부터 접두사
    binmaps, binout = [{}], dict(base["bin"])
    for i, d in enumerate(docs[1:], start=2):
        m = {}
        for name, data in d["bin"].items():
            stem = name[len("BinData/"):]
            new = "f%d_%s" % (i, stem)
            binout["BinData/" + new] = data
            m[Path(stem).stem] = Path(new).stem  # binaryItemIDRef는 확장자 없는 id
        binmaps.append(m)

    # header — 항목의 id와 항목 안의 참조를 함께 옮겨 이어 붙인다
    header = base["header"]
    for box, tag in SPACES.items():
        merged = []
        for d, mp, bm in zip(docs, maps, binmaps):
            for it in _box_items(d["header"], box, tag):
                merged.append(_set_own_id(_remap(it, mp, bm), mp[tag][_own_id(it)]))
        if merged:
            header = _replace_box(header, box, merged)

    # fontfaces는 언어별로 따로 이어 붙인다
    fonts = [_fonts(d["header"]) for d in docs]
    ff = _inner(header, "fontfaces")
    if ff:
        block = ff[0]
        for lang in FONT_LANGS:
            fm = re.search(r'(<hh:fontface\b[^>]*lang="%s"[^>]*>)(.*?)(</hh:fontface>)' % lang,
                           block, re.S)
            if not fm:
                continue
            items = []
            for f, mp, bm in zip(fonts, maps, binmaps):
                for it in f.get(lang, []):
                    items.append(_set_own_id(_remap(it, mp, bm),
                                             mp["font"][lang][_own_id(it)]))
            open_tag = re.sub(r'fontCnt="\d+"', 'fontCnt="%d"' % len(items), fm.group(1))
            block = block[:fm.start()] + open_tag + "".join(items) + fm.group(3) + block[fm.end():]
        header = _replace_raw(header, "fontfaces", block)

    # 본문 — 파일 순서대로 section0..N
    step("본문을 이어 붙이는 중", 0, 0)
    sections = []
    for d, mp, bm in zip(docs, maps, binmaps):
        sections += [_remap(s, mp, bm) for s in d["sections"]]

    # 한글은 hh:head의 secCnt로 읽을 구역 수를 정한다. 여기를 안 고치면 spine에 등록해도
    # 첫 구역만 보인다(실측: 3개 병합 결과가 뷰어에서 1/1구역·첫 파일 몫만 표시됨).
    header = re.sub(r'(<hh:head\b[^>]*?)\bsecCnt="\d+"', r'\1secCnt="%d"' % len(sections),
                    header, count=1)

    hpf = _rebuild_hpf(base["hpf"], len(sections), sorted(binout))

    # 검증 — 통과하지 못하면 파일을 쓰지 않는다
    step("검증 중", 0, 0)
    report = _verify(header, sections, binout, docs)
    report["파일"] = [p.name for p in paths]
    report["구역"] = {p.name: len(d["sections"]) for p, d in zip(paths, docs)}
    report["한계"] = _limits(docs)
    if report["dangling"] or report["itemCnt불일치"] or any(
            v["원본합계"] != v["병합결과"] for v in report["수량대조"].values()):
        report["결과"] = "실패"
        return report

    step("결과 파일을 저장하는 중", 0, 0)
    _write(out, header, sections, hpf, binout, base["extra"])
    report["결과"] = "성공"
    report["출력"] = str(out)
    return report


def _replace_raw(header: str, box: str, new_inner: str) -> str:
    _, start, end = _inner(header, box)
    open_tag = re.match(r"<hh:%s\b[^>]*?>" % box, header[start:end]).group(0)
    return header[:start] + open_tag + new_inner + "</hh:%s>" % box + header[end:]


def _rebuild_hpf(hpf: str, n_sections: int, bin_names: list[str]) -> str:
    items = ['<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>']
    for n in bin_names:
        mt = MEDIA.get(Path(n).suffix.lower(), "application/octet-stream")
        items.append('<opf:item id="%s" href="%s" media-type="%s" isEmbeded="1"/>'
                     % (Path(n).stem, n, mt))
    for i in range(n_sections):
        items.append('<opf:item id="section%d" href="Contents/section%d.xml" '
                     'media-type="application/xml"/>' % (i, i))
    items.append('<opf:item id="settings" href="settings.xml" media-type="application/xml"/>')
    spine = ['<opf:itemref idref="header" linear="yes"/>']
    spine += ['<opf:itemref idref="section%d" linear="yes"/>' % i for i in range(n_sections)]
    body = ("<opf:manifest>" + "".join(items) + "</opf:manifest>"
            "<opf:spine>" + "".join(spine) + "</opf:spine>")
    return re.sub(r"<opf:manifest>.*</opf:spine>", body, hpf, flags=re.S)


def _write(out: Path, header: str, sections: list[str], hpf: str,
           binout: dict[str, bytes], extra: dict[str, bytes]) -> None:
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        # mimetype은 무압축 첫 엔트리여야 한다(ODF 관례, hwpx도 동일).
        z.writestr(zipfile.ZipInfo("mimetype"), MIMETYPE, zipfile.ZIP_STORED)
        z.writestr("Contents/header.xml", header)
        for i, s in enumerate(sections):
            z.writestr("Contents/section%d.xml" % i, s)
        z.writestr("Contents/content.hpf", hpf)
        for name, data in binout.items():
            z.writestr(name, data)
        for name, data in extra.items():
            z.writestr(name, data)


# -------------------------------------------------------------------- 검증

def _defined(header: str) -> dict[str, set[int]]:
    return {tag: {_own_id(it) for it in _box_items(header, box, tag)}
            for box, tag in SPACES.items()}


def _verify(header: str, sections: list[str], binout: dict, docs: list[dict]) -> dict:
    defined = _defined(header)
    fonts = {lang: {_own_id(it) for it in v} for lang, v in _fonts(header).items()}
    whole = header + "".join(sections)
    dangling = []

    for attr, space in REFS.items():
        for val in sorted(set(re.findall(r'\b%s="(\d+)"' % attr, whole))):
            if int(val) != NONE_ID and int(val) not in defined[space]:
                dangling.append("%s=%s (%s 미정의)" % (attr, val, space))
    for m in re.finditer(r"<hh:fontRef\b[^>]*/>", whole):
        for lang in FONT_LANGS:
            v = re.search(r'\b%s="(\d+)"' % lang.lower(), m.group(0))
            if v and int(v.group(1)) not in fonts.get(lang, set()):
                dangling.append("fontRef %s=%s (font %s 미정의)" % (lang.lower(), v.group(1), lang))
    refs = set(re.findall(r'binaryItemIDRef="([^"]+)"', whole))
    stems = {Path(n).stem for n in binout}
    dangling += ["binaryItemIDRef=%s (BinData 없음)" % r for r in sorted(refs - stems)]

    src = [s for d in docs for s in d["sections"]]
    tally = {}
    for label, pat in (("문단", r"<hp:p\b"), ("표", r"<hp:tbl\b"), ("그림", r"<hp:pic\b")):
        tally[label] = {"원본합계": sum(len(re.findall(pat, x)) for x in src),
                        "병합결과": sum(len(re.findall(pat, x)) for x in sections)}
    text = lambda xs: sum(len("".join(re.findall(r"<hp:t>(.*?)</hp:t>", x, re.S))) for x in xs)
    tally["글자수"] = {"원본합계": text(src), "병합결과": text(sections)}

    itemcnt = []
    sec = re.search(r'<hh:head\b[^>]*?\bsecCnt="(\d+)"', header)
    if sec and int(sec.group(1)) != len(sections):
        itemcnt.append("hh:head secCnt=%s 실제 구역=%d" % (sec.group(1), len(sections)))
    for box, tag in SPACES.items():
        found = _inner(header, box)
        if not found:
            continue
        declared = re.search(r'itemCnt="(\d+)"', header[found[1]:found[2]])
        actual = len(_items(found[0], tag))
        if declared and int(declared.group(1)) != actual:
            itemcnt.append("%s: itemCnt=%s 실제=%d" % (box, declared.group(1), actual))
    for lang, ids in fonts.items():
        m = re.search(r'<hh:fontface\b[^>]*lang="%s"[^>]*fontCnt="(\d+)"' % lang, header)
        if m and int(m.group(1)) != len(ids):
            itemcnt.append("fontface %s: fontCnt=%s 실제=%d" % (lang, m.group(1), len(ids)))

    return {
        "dangling": dangling,
        "수량대조": tally,
        "itemCnt불일치": itemcnt,
        "고아BinData": sorted(stems - refs),
        "미대응": {a: len(re.findall(r'\b%s="' % a, whole)) for a in UNTOUCHED
                   if re.search(r'\b%s="' % a, whole)},
        "정의수": {k: len(v) for k, v in defined.items()} | {
            "font": {lang: len(v) for lang, v in fonts.items()}},
    }


def _limits(docs: list[dict]) -> list[str]:
    """조용히 버려지는 것들을 건수로 남긴다."""
    out = ["settings.xml·version.xml·Preview는 첫 파일 것만 유지됩니다"]
    diff = [d["path"].name for d in docs[1:]
            if d["extra"].get("settings.xml") != docs[0]["extra"].get("settings.xml")]
    if diff:
        out.append("문서 설정(settings.xml)이 첫 파일과 다른 파일 %d개: %s"
                   % (len(diff), ", ".join(diff)))
    return out


# ---------------------------------------------------------------------- CLI

def main() -> int:
    ap = argparse.ArgumentParser(description="hwpx 병합 (파일당 별도 구역으로 이어 붙임)")
    ap.add_argument("files", nargs="+", type=Path, help="병합할 hwpx (지정한 순서대로)")
    ap.add_argument("-o", "--out", type=Path, default=Path("merged.hwpx"))
    a = ap.parse_args()
    try:
        r = merge(a.files, a.out)
    except MergeError as e:
        print("병합 불가:", e)
        return 1

    print("입력 %d개 → 구역 %d개" % (len(r["파일"]), sum(r["구역"].values())))
    for name, n in r["구역"].items():
        print("  · %s (구역 %d)" % (name, n))
    print("\n[수량 대조] 원본 합계 = 병합 결과")
    for k, v in r["수량대조"].items():
        print("  %-5s %8d = %8d  %s" % (k, v["원본합계"], v["병합결과"],
                                        "일치" if v["원본합계"] == v["병합결과"] else "불일치"))
    print("\n[ID 공간] " + ", ".join("%s %d" % (k, v) for k, v in r["정의수"].items()
                                     if k != "font"))
    print("[font] " + ", ".join("%s %d" % (k, v) for k, v in r["정의수"]["font"].items()))
    if r["미대응"]:
        print("[미대응 참조] header에 정의가 없어 그대로 둔 참조: "
              + ", ".join("%s %d건" % (k, v) for k, v in r["미대응"].items()))
    if r["itemCnt불일치"]:
        print("[itemCnt 불일치]", "; ".join(r["itemCnt불일치"]))
    if r["고아BinData"]:
        print("[고아 BinData]", ", ".join(r["고아BinData"]))
    for l in r["한계"]:
        print("[한계]", l)
    if r["결과"] != "성공":
        print("\n실패 — 결과 파일을 쓰지 않았습니다.")
        for d in r["dangling"][:20]:
            print("  ·", d)
        if len(r["dangling"]) > 20:
            print("  … 그 외 %d건" % (len(r["dangling"]) - 20))
        return 1
    print("\n결과:", r["출력"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
