#!/usr/bin/env python3
"""설교 자막 PPT 빌더.

spec.json을 읽어 template.pptx의 예시 슬라이드를 복제·텍스트 교체하여
설교 자막 PPT를 생성한다.

사용법:
    python build.py spec.json output.pptx [--template template.pptx]

spec.json 형식:
{
  "date": "2026년 7월 19일 주일말씀",
  "title": "설교 제목",
  "slides": [
    {"type": "caption", "lines": ["하단 자막 첫 줄", "둘째 줄"]},
    {"type": "word",    "reveal": true, "lines": ["말씀 1줄", "말씀 2줄", "", "강조 줄"]},
    {"type": "verse",   "ref": "(전 1:2~3)", "lines": ["“전도자가 가로되", "...유익한고”"]},
    {"type": "keyverse","ref": "잠 13:12", "lines": ["“소망이 더디...", "...생명 나무니라”"]}
  ]
}

- type:
    caption  = 아래자막(녹색 크로마 배경 + 하단 크림 박스). 원고의 파란색 글자.
    word     = 말씀 슬라이드(예수님 배경, 상단 텍스트). 원고의 녹색 글자 중 말씀 인용.
    verse    = 성경 본문(액자 배경, 장절 표기 + 본문). 원고의 녹색 글자 중 성구.
    keyverse = 본문 요절 강조(보라색 명조체 + 우하단 장절).
- reveal: true 이면 lines를 한 줄씩 누적하며 슬라이드를 여러 장 생성 (word/verse에서 사용).
- lines 안의 "" (빈 문자열)은 빈 줄(단락 간격)로 렌더링되고, reveal 시 다음 줄과 함께 나타난다.
- 표지(cover)는 date/title로 자동 생성되어 항상 첫 장이 된다.
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

PPTX_SCRIPTS = Path("/mnt/skills/public/pptx/scripts")

# 예시 슬라이드 (template.pptx 내 파일명)
EXEMPLAR = {
    "passage": "slide1.xml",   # 본문 성구(검정 큰 글씨 + 노란 장절 박스), 맨 앞
    "cover": "slide5.xml",     # 표지(녹색 크로마)
    "word": "slide16.xml",     # 말씀/성구 인용(예수님 배경, 상단)
    "caption": "slide8.xml",   # 아래자막(녹색 크로마 하단 박스)
    "verse": "slide11.xml",    # 인용 성구(액자, 파란 장절)
    "emphasis": "slide13.xml", # 강조(흰 배경 큰 글씨)
    "emphasis2": "slide54.xml",# 강조2(별도 배경 큰 글씨)
    "notice": "slide67.xml",   # 공지(광고/일정)
    "closing": "slide71.xml",  # 마무리(예배명 + 제목)
    "photo": "slide5.xml",     # 전체화면 사진 (표지 슬라이드를 복제해 도형 제거 후 그림 삽입)
}

# title 하나에 lines를 넣는 단순 유형 (word와 동일 처리)
SIMPLE_TITLE_TYPES = {"word", "caption", "emphasis", "emphasis2", "notice"}

# 파란색(이어짐 표시 >>, <b> 마커) 색상. 템플릿의 파란 장절 색과 동일.
BLUE_HEX = "2683C6"

# 이어짐 표시. 앞장 자막 박스 오른쪽 끝에 파란 글씨로 얹는다.
CONT_MARK = ">>"

# 이어짐 표시를 자막 박스 오른쪽 끝에서 얼마나 띄울지 (EMU). 분홍 테두리 폭을 감안한 값.
CONT_RIGHT_MARGIN = 100000

# caption 한 장 기본 줄 수. 4줄 이상이면 이 수로 쪼갠다 (3줄까지는 그대로 둔다).
CAPTION_LINES_PER_PAGE = 2
CAPTION_SPLIT_MIN = 4

# 누적 공개 자막에서 줄 높이만 차지하는 보이지 않는 줄 (줄바꿈 없는 공백)
PAD_LINE = " "

# word 슬라이드 화자별 배경 레이아웃. speaker 값 → 레이아웃 파일명.
SPEAKER_LAYOUT = {
    "성령": "slideLayout83.xml",   # 2_성령님말씀
}


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f"명령 실패: {' '.join(map(str, cmd))}\n{r.stdout}\n{r.stderr}")
    return r.stdout


def first_rpr(block, must_contain=None):
    """블록 안에서 (조건에 맞는) 첫 <a:rPr>...</a:rPr> 전체 요소를 반환."""
    for m in re.finditer(r"<a:rPr\b[^>]*?/>|<a:rPr\b[^>]*>.*?</a:rPr>", block, re.S):
        if must_contain is None or must_contain in m.group(0):
            return m.group(0)
    return None


def first_ppr(block):
    m = re.search(r"<a:pPr\b[^>]*?/>|<a:pPr\b[^>]*>.*?</a:pPr>", block, re.S)
    return m.group(0) if m else ""


def make_run(rpr, text):
    t = escape(text)
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f"<a:r>{rpr}<a:t{space}>{t}</a:t></a:r>"


def recolor_rpr(rpr, hex_color):
    """주어진 rPr의 글자색을 hex_color로 바꾼 새 rPr을 반환.

    schemeClr/srgbClr 형태의 solidFill을 해당 색 srgbClr로 치환하고,
    solidFill이 없으면 하나 삽입한다.
    """
    fill = f'<a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
    if rpr is None:
        return f'<a:rPr lang="ko-KR" altLang="en-US" dirty="0">{fill}</a:rPr>'
    if "<a:solidFill>" in rpr:
        return re.sub(r"<a:solidFill>.*?</a:solidFill>", fill, rpr, count=1, flags=re.S)
    # solidFill 없음: 여는 태그 뒤에 삽입
    m = re.match(r"(<a:rPr\b[^>]*?)(/?)>", rpr)
    if not m:
        return rpr
    if m.group(2) == "/":  # self-closing <a:rPr .../>
        return m.group(1) + ">" + fill + "</a:rPr>"
    return rpr[: m.end()] + fill + rpr[m.end():]


# 인라인 색상 마커: <r>…</r> = 빨강, <b>…</b> = 파랑
COLOR_PATTERN = re.compile(r"<(r|b)>(.*?)</\1>", re.S)
COLOR_HEX = {"r": "FF0000", "b": BLUE_HEX}


def line_to_runs(line, base_rpr):
    """한 줄 문자열을 (rpr, text) run 목록으로 변환.

    line 안의 <r>…</r>는 빨간색, <b>…</b>는 파란색 rPr을 적용한다.
    마커가 없으면 [(base_rpr, line)] 하나를 반환.
    """
    if "<r>" not in line and "<b>" not in line:
        return [(base_rpr, line)]
    runs = []
    pos = 0
    for m in COLOR_PATTERN.finditer(line):
        if m.start() > pos:
            runs.append((base_rpr, line[pos:m.start()]))
        runs.append((recolor_rpr(base_rpr, COLOR_HEX[m.group(1)]), m.group(2)))
        pos = m.end()
    if pos < len(line):
        runs.append((base_rpr, line[pos:]))
    return runs


def make_br(rpr):
    return f"<a:br>{rpr}</a:br>"


def build_paragraph(ppr, segments):
    """segments: (rpr, text|None) 목록. text가 None이면 줄바꿈(<a:br>)."""
    parts = [ppr]
    for rpr, text in segments:
        parts.append(make_br(rpr) if text is None else make_run(rpr, text))
    return "<a:p>" + "".join(parts) + "</a:p>"


def lines_to_segments(lines, base_rpr, br_rpr=None):
    """여러 줄을 build_paragraph용 segments로 변환. 줄 사이에 <a:br> 삽입.

    각 줄은 line_to_runs로 처리되어 <r>…</r> 빨강 구간을 지원한다.
    """
    if br_rpr is None:
        br_rpr = base_rpr
    segs = []
    for i, line in enumerate(lines):
        if i > 0:
            segs.append((br_rpr, None))
        if line:
            segs.extend(line_to_runs(line, base_rpr))
    return segs


def replace_txbody_paragraphs(txbody, new_paragraphs_xml):
    """txBody의 bodyPr/lstStyle은 유지하고 <a:p>들만 교체."""
    head = re.split(r"<a:p>", txbody, maxsplit=1)[0]
    return head + new_paragraphs_xml + "</p:txBody>"


def get_sp_blocks(xml):
    return re.findall(r"<p:sp>.*?</p:sp>", xml, re.S)


def insert_full_photo(workdir, slide_path, image_path):
    """슬라이드의 모든 도형을 지우고 이미지를 화면에 꽉 차게(contain) 배치.

    - 배경(녹색 크로마 등)은 유지하고 그 위에 그림을 얹는다.
    - 이미지 종횡비를 유지하며 슬라이드 안에 최대로 맞추고 가운데 정렬한다.
    """
    from PIL import Image as PILImage

    slide_path = Path(slide_path)
    image_path = Path(image_path)
    n = re.search(r"slide(\d+)\.xml", slide_path.name).group(1)

    # 슬라이드 크기
    pres = (workdir / "ppt" / "presentation.xml").read_text(encoding="utf-8")
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    SW, SH = int(m.group(1)), int(m.group(2))

    # 이미지 크기 → contain 배치 계산
    with PILImage.open(image_path) as im:
        iw, ih = im.size
    scale = min(SW / iw, SH / ih)
    w, h = int(iw * scale), int(ih * scale)
    x, y = (SW - w) // 2, (SH - h) // 2

    # 미디어 파일 복사
    media_dir = workdir / "ppt" / "media"
    media_dir.mkdir(exist_ok=True)
    ext = image_path.suffix.lower().lstrip(".")
    if ext == "jpg":
        ext = "jpeg"
    existing = [p.name for p in media_dir.glob("image*")]
    idx = 1
    while f"image_photo{idx}.{ext}" in existing:
        idx += 1
    media_name = f"image_photo{idx}.{ext}"
    shutil.copy(image_path, media_dir / media_name)

    # 슬라이드 rels에 관계 추가
    rels_path = workdir / "ppt" / "slides" / "_rels" / f"slide{n}.xml.rels"
    rels = rels_path.read_text(encoding="utf-8")
    rids = [int(r) for r in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{max(rids) + 1 if rids else 1}"
    new_rel = (f'<Relationship Id="{rid}" '
               f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
               f'Target="../media/{media_name}"/>')
    rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
    rels_path.write_text(rels, encoding="utf-8")

    # content types에 확장자 등록 보장
    ct_path = workdir / "[Content_Types].xml"
    ct = ct_path.read_text(encoding="utf-8")
    if f'Extension="{ext}"' not in ct:
        ct = ct.replace("</Types>",
                        f'<Default Extension="{ext}" ContentType="image/{ext}"/></Types>')
        ct_path.write_text(ct, encoding="utf-8")

    # 슬라이드의 모든 <p:sp> 제거하고 검정 배경 + <p:pic> 삽입
    xml = slide_path.read_text(encoding="utf-8")
    xml = re.sub(r"<p:sp>.*?</p:sp>", "", xml, flags=re.S)
    # 기존 배경 제거 후 검정 배경 지정 (세로/가로 사진의 여백을 검정으로)
    xml = re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)
    black_bg = ('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
                '<a:effectLst/></p:bgPr></p:bg>')
    xml = re.sub(r"(<p:cSld[^>]*>)", r"\1" + black_bg, xml, count=1)
    pic = (
        f'<p:pic><p:nvPicPr>'
        f'<p:cNvPr id="10" name="Photo"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    xml = re.sub(r"(</p:grpSpPr>)", r"\1" + pic, xml, count=1)
    slide_path.write_text(xml, encoding="utf-8")


def set_slide_text(slide_path, slide_type, entry, styles):
    if slide_type == "photo":
        return  # 사진 슬라이드는 main에서 이미지 삽입으로 별도 처리
    xml = Path(slide_path).read_text(encoding="utf-8")
    sps = get_sp_blocks(xml)

    def txbody_of(sp):
        return re.search(r"<p:txBody>.*?</p:txBody>", sp, re.S).group(0)

    if slide_type == "cover":
        # sp[0] = 제목 placeholder, sp[1] = 날짜 박스 (템플릿 고정 구조)
        title_sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
        date_sp = next(sp for sp in sps if '<p:ph type="title"' not in sp)
        st = styles["cover"]
        new_title = build_paragraph(st["title_ppr"], [(st["title_rpr"], entry["title"])])
        new_date = build_paragraph(st["date_ppr"], [(st["date_rpr"], entry["date"])])
        xml = xml.replace(txbody_of(title_sp), replace_txbody_paragraphs(txbody_of(title_sp), new_title))
        xml = xml.replace(txbody_of(date_sp), replace_txbody_paragraphs(txbody_of(date_sp), new_date))

    elif slide_type in SIMPLE_TITLE_TYPES:
        st = styles[slide_type]
        segs = lines_to_segments(entry["lines"], st["rpr"])
        sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
        tb = txbody_of(sp)
        xml = xml.replace(tb, replace_txbody_paragraphs(tb, build_paragraph(st["ppr"], segs)))
        if entry.get("cont"):
            xml = xml.replace("</p:spTree>",
                              cont_shape_xml(st, len(entry["lines"])) + "</p:spTree>", 1)

    elif slide_type == "passage":
        # 본문 성구: title(검정 큰 글씨 본문) + 노란 장절 박스
        st = styles["passage"]
        body_sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
        ref_sp = next(sp for sp in sps if '<p:ph type="title"' not in sp)
        segs = lines_to_segments(entry["lines"], st["rpr"])
        tb = txbody_of(body_sp)
        xml = xml.replace(tb, replace_txbody_paragraphs(tb, build_paragraph(st["ppr"], segs)))
        tb2 = txbody_of(ref_sp)
        xml = xml.replace(tb2, replace_txbody_paragraphs(
            tb2, build_paragraph(st["ref_ppr"], [(st["ref_rpr"], entry["ref"])])))

    elif slide_type == "closing":
        # 마무리: 첫 문단=예배명, 마지막 비지 않은 문단=제목. 서식은 원본 유지하고 텍스트만 교체.
        sp = sps[0]
        tb = txbody_of(sp)
        paras = re.findall(r"<a:p>.*?</a:p>", tb, re.S)

        def set_para_text(para, new_text):
            runs = re.findall(r"<a:r>.*?</a:r>", para, re.S)
            if not runs:
                return para
            rpr_m = re.search(r"<a:rPr\b[^>]*?/>|<a:rPr\b[^>]*>.*?</a:rPr>", runs[0], re.S)
            rpr = rpr_m.group(0) if rpr_m else ""
            new_run = make_run(rpr, new_text)
            # 첫 run 위치에 새 run 하나만, 나머지 run 제거
            first_start = para.index(runs[0])
            last_end = para.index(runs[-1]) + len(runs[-1])
            return para[:first_start] + new_run + para[last_end:]

        new_paras = list(paras)
        new_paras[0] = set_para_text(paras[0], entry["worship"])
        for i in range(len(new_paras) - 1, -1, -1):
            if re.findall(r"<a:t>(.*?)</a:t>", paras[i]):
                if i != 0:
                    new_paras[i] = set_para_text(paras[i], entry["title"])
                    break
        new_tb = tb
        for old, new in zip(paras, new_paras):
            if old != new:
                new_tb = new_tb.replace(old, new, 1)
        xml = xml.replace(tb, new_tb)

    elif slide_type == "verse":
        st = styles["verse"]
        segs = []
        if entry.get("ref"):
            segs = [(st["ref_rpr"], entry["ref"]), (st["ref_rpr"], None), (st["body_rpr"], None)]
        segs.extend(lines_to_segments(entry["lines"], st["body_rpr"]))
        sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
        tb = txbody_of(sp)
        xml = xml.replace(tb, replace_txbody_paragraphs(tb, build_paragraph(st["ppr"], segs)))

    Path(slide_path).write_text(xml, encoding="utf-8")


def extract_styles(workdir):
    """템플릿 예시 슬라이드에서 pPr/rPr 서식을 추출."""
    styles = {}
    slides = workdir / "ppt" / "slides"

    # 단순 title 유형: 첫 pPr/rPr만 추출
    for t in SIMPLE_TITLE_TYPES:
        x = (slides / EXEMPLAR[t]).read_text(encoding="utf-8")
        title_sp = next(sp for sp in get_sp_blocks(x) if '<p:ph type="title"' in sp)
        styles[t] = {"ppr": first_ppr(title_sp), "rpr": first_rpr(title_sp)}

    # caption: 이어짐 표시 상자를 얹으려면 레이아웃의 자막 글상자 좌표·bodyPr이 필요하다
    rels = (slides / "_rels" / (EXEMPLAR["caption"] + ".rels")).read_text(encoding="utf-8")
    layout_name = re.search(r"slideLayouts/(slideLayout\d+\.xml)", rels).group(1)
    lx = (workdir / "ppt" / "slideLayouts" / layout_name).read_text(encoding="utf-8")
    lsp = next(sp for sp in get_sp_blocks(lx) if '<p:ph type="title"' in sp)
    g = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', lsp)
    styles["caption"]["cont_geom"] = (g.group(1), g.group(2), g.group(4))
    # 자막 크림 박스(레이아웃의 그림)의 오른쪽 끝까지 이어짐 표시 상자를 늘린다
    box_right = int(g.group(1)) + int(g.group(3))
    for m in re.finditer(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', lx):
        if int(m.group(4)) < 300000:
            continue
        box_right = max(box_right, int(m.group(1)) + int(m.group(3)))
    styles["caption"]["cont_right"] = box_right - CONT_RIGHT_MARGIN
    # bodyPr는 마스터의 제목 자리표시자에서 가져온다. 레이아웃 쪽에는 anchor가 없어
    # 그대로 쓰면 일반 텍스트 상자가 위쪽 정렬이 되어 >> 가 본문보다 위로 뜬다.
    master_name = re.search(r"slideMasters/(slideMaster\d+\.xml)",
                            (workdir / "ppt" / "slideLayouts" / "_rels" /
                             (layout_name + ".rels")).read_text(encoding="utf-8")).group(1)
    mx = (workdir / "ppt" / "slideMasters" / master_name).read_text(encoding="utf-8")
    msp = next(sp for sp in get_sp_blocks(mx) if '<p:ph type="title"' in sp)
    bp = re.search(r"<a:bodyPr\b[^>]*?/>|<a:bodyPr\b[^>]*>.*?</a:bodyPr>", msp, re.S)
    styles["caption"]["cont_bodypr"] = bp.group(0) if bp else '<a:bodyPr anchor="ctr"/>'

    # passage: 본문(title) + 노란 장절 박스
    x = (slides / EXEMPLAR["passage"]).read_text(encoding="utf-8")
    sps = get_sp_blocks(x)
    body_sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
    ref_sp = next(sp for sp in sps if '<p:ph type="title"' not in sp)
    styles["passage"] = {
        "ppr": first_ppr(body_sp),
        "rpr": first_rpr(body_sp),
        "ref_ppr": first_ppr(ref_sp),
        "ref_rpr": first_rpr(ref_sp),
    }

    x = (slides / EXEMPLAR["verse"]).read_text(encoding="utf-8")
    styles["verse"] = {
        "ppr": first_ppr(x),
        "ref_rpr": first_rpr(x, must_contain='val="accent2"'),
        "body_rpr": first_rpr(x, must_contain='val="tx1"'),
    }

    x = (slides / EXEMPLAR["cover"]).read_text(encoding="utf-8")
    sps = get_sp_blocks(x)
    title_sp = next(sp for sp in sps if '<p:ph type="title"' in sp)
    date_sp = next(sp for sp in sps if '<p:ph type="title"' not in sp)
    styles["cover"] = {
        "title_ppr": first_ppr(title_sp),
        "title_rpr": first_rpr(title_sp),
        "date_ppr": first_ppr(date_sp),
        "date_rpr": first_rpr(date_sp),
    }

    return styles


def _pages_from_lines(lines):
    """줄 목록을 페이지(줄 목록의 목록)로 나눈다.

    3줄 이하는 그대로 한 장. 4줄 이상이면 2줄씩 나누고,
    마지막에 1줄만 남으면 앞 페이지에 붙여 3줄로 만든다.
    """
    if len(lines) < CAPTION_SPLIT_MIN:
        return [list(lines)]
    n = CAPTION_LINES_PER_PAGE
    pages = [list(lines[i:i + n]) for i in range(0, len(lines), n)]
    if len(pages) > 1 and len(pages[-1]) == 1:
        last = pages.pop()
        pages[-1] = pages[-1] + last
    return pages


def _pages_from_chunks(chunks):
    """청크(문장) 목록을 페이지로 나눈다. 문장 중간은 자르지 않는다.

    반환값은 페이지별 청크 목록. 총 줄 수가 3줄 이하면 나누지 않는다.
    """
    total = sum(len(c) for c in chunks)
    if total < CAPTION_SPLIT_MIN:
        return [list(chunks)]
    pages, cur, cur_n = [], [], 0
    for c in chunks:
        if cur and cur_n + len(c) > CAPTION_LINES_PER_PAGE:
            pages.append(cur)
            cur, cur_n = [], 0
        cur.append(list(c))
        cur_n += len(c)
    if cur:
        pages.append(cur)
    return pages


def split_caption(entry):
    """caption이 4줄 이상이면 여러 장으로 나누고, 앞장 끝에 파란 >> 를 붙인다."""
    if entry.get("chunks"):
        pages = _pages_from_chunks(entry["chunks"])
        out = []
        for i, page in enumerate(pages):
            e = {k: v for k, v in entry.items() if k != "chunks"}
            e["chunks"] = [list(c) for c in page]
            if i < len(pages) - 1:
                e["cont"] = True
            out.append(e)
        return out
    pages = _pages_from_lines(entry.get("lines", []))
    out = []
    for i, page in enumerate(pages):
        e = dict(entry)
        e["lines"] = page
        if i < len(pages) - 1:
            e["cont"] = True
        out.append(e)
    return out


def cont_shape_xml(st, n_lines):
    """이어짐 표시(>>)를 자막 박스 오른쪽 끝에 놓는 텍스트 상자 XML을 만든다.

    자막 글상자와 같은 좌표·같은 줄 수를 쓰되 오른쪽 정렬이라,
    >>가 마지막 줄과 같은 높이의 오른쪽 끝에 놓인다.
    """
    x, y, cy = st["cont_geom"]
    cx = st["cont_right"] - int(x)
    ppr = re.sub(r'algn="[a-z]+"', 'algn="r"', st["ppr"])
    if 'algn=' not in ppr:
        ppr = ppr.replace("<a:pPr", '<a:pPr algn="r"', 1)
    rpr = recolor_rpr(st["rpr"], BLUE_HEX)
    bodypr = st["cont_bodypr"]
    if 'rIns=' in bodypr:
        bodypr = re.sub(r'rIns="-?\d+"', 'rIns="0"', bodypr)
    else:
        bodypr = re.sub(r"<a:bodyPr\b", '<a:bodyPr rIns="0"', bodypr, count=1)
    segs = [(rpr, None)] * max(0, n_lines - 1) + [(rpr, CONT_MARK)]
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="90" name="이어짐표시"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody>{bodypr}<a:lstStyle/>{build_paragraph(ppr, segs)}</p:txBody></p:sp>'
    )


def apply_speaker_layout(workdir, slide_path, speaker):
    """word 슬라이드의 배경 레이아웃을 화자에 맞게 바꾼다.

    슬라이드에 박혀 있는 좌표(xfrm)를 지워 레이아웃의 자리표시자 위치를
    상속하게 한다. 성령님 배경은 인물이 왼쪽에 있어 좌표를 그대로 두면 겹친다.
    """
    layout = SPEAKER_LAYOUT.get(speaker)
    if not layout:
        sys.exit(f"알 수 없는 speaker: {speaker} (가능: {', '.join(SPEAKER_LAYOUT)})")
    slide_path = Path(slide_path)
    n = re.search(r"slide(\d+)\.xml", slide_path.name).group(1)
    rels_p = workdir / "ppt" / "slides" / "_rels" / f"slide{n}.xml.rels"
    rels = rels_p.read_text(encoding="utf-8")
    rels = re.sub(r'Target="\.\./slideLayouts/slideLayout\d+\.xml"',
                  f'Target="../slideLayouts/{layout}"', rels, count=1)
    rels_p.write_text(rels, encoding="utf-8")
    xml = slide_path.read_text(encoding="utf-8")
    xml = re.sub(r"<p:spPr><a:xfrm>.*?</a:xfrm></p:spPr>", "<p:spPr/>", xml, flags=re.S)
    slide_path.write_text(xml, encoding="utf-8")


def expand_reveal(entry):
    """reveal 항목을 누적 슬라이드 목록으로 확장.

    - "chunks": [[줄,...], [줄,...]] 가 있으면 청크(문장) 단위로 누적 (권장).
    - 없으면 "lines"를 한 줄씩 누적 (기존 방식).
    """
    if entry.get("chunks"):
        out = []
        acc = []
        total = sum(len(c) for c in entry["chunks"])
        chunks = entry["chunks"]
        for i, chunk in enumerate(chunks):
            acc = acc + list(chunk)
            e = {k: v for k, v in entry.items() if k not in ("chunks", "reveal")}
            e["lines"] = list(acc)
            if entry.get("type") == "caption":
                # 자막 박스는 세로 가운데 정렬이라 줄 수가 늘면 앞 줄이 위로 밀린다.
                # 보이지 않는 줄을 채워 최종 줄 수에 맞추면 앞 줄 위치가 고정된다.
                e["lines"] += [PAD_LINE] * (total - len(acc))
                # 이어짐 표시는 그 장의 마지막 누적 슬라이드에만 붙인다.
                if i < len(chunks) - 1:
                    e.pop("cont", None)
            out.append(e)
        return out
    if not entry.get("reveal"):
        return [entry]
    out = []
    lines = entry["lines"]
    for i in range(len(lines)):
        if lines[i] == "":
            continue  # 빈 줄 단독으로는 슬라이드를 만들지 않음
        e = dict(entry)
        e["lines"] = lines[: i + 1]
        e.pop("reveal", None)
        out.append(e)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("output")
    ap.add_argument("--template", default=str(Path(__file__).parent.parent / "assets" / "template.pptx"))
    args = ap.parse_args()

    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    workdir = Path(tempfile.mkdtemp(prefix="jamak_"))
    zipfile.ZipFile(args.template).extractall(workdir)

    styles = extract_styles(workdir)

    # 표지 + 본문 슬라이드 평탄화
    # slides 안에 cover 항목이 있으면 그 순서를 그대로 따르고 자동 표지는 만들지 않음.
    # spec에 title/date가 비어 있으면(사용자가 cover를 직접 배치) 자동 표지를 만들지 않음.
    has_explicit_cover = any(e.get("type") == "cover" for e in spec["slides"])
    auto_cover = (not has_explicit_cover) and spec.get("title") and spec.get("date")
    flat = []
    if auto_cover:
        flat.append(("cover", {"title": spec["title"], "date": spec["date"]}))
    for entry in spec["slides"]:
        t = entry["type"]
        if t not in EXEMPLAR:
            sys.exit(f"알 수 없는 type: {t}")
        if t == "cover":
            flat.append(("cover", entry))
            continue
        parts = split_caption(entry) if t == "caption" else [entry]
        for part in parts:
            for e in expand_reveal(part):
                flat.append((t, e))

    # 예시 슬라이드 복제 (sldIdLst 끝에 순서대로 추가됨)
    new_slides = []
    for t, e in flat:
        out = run([sys.executable, str(PPTX_SCRIPTS / "add_slide.py"), str(workdir), EXEMPLAR[t]])
        m = re.search(r"(ppt/slides/slide\d+\.xml)", out)
        if not m:
            sys.exit(f"add_slide 결과 해석 실패:\n{out}")
        new_slides.append((t, e, workdir / m.group(1)))

    # 텍스트 교체
    for t, e, path in new_slides:
        set_slide_text(path, t, e, styles)

    # 사진 슬라이드: 전체화면 이미지 삽입
    for t, e, path in new_slides:
        if t == "photo":
            insert_full_photo(workdir, path, e["image"])

    # word 슬라이드: 화자별 배경 레이아웃 교체
    for t, e, path in new_slides:
        if t == "word" and e.get("speaker"):
            apply_speaker_layout(workdir, path, e["speaker"])

    # sldIdLst를 새 슬라이드만 순서대로 남기고 재작성
    pres_p = workdir / "ppt" / "presentation.xml"
    rels_p = workdir / "ppt" / "_rels" / "presentation.xml.rels"
    pres = pres_p.read_text(encoding="utf-8")
    rels = rels_p.read_text(encoding="utf-8")
    file2rid = {v: k for k, v in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+)\.xml"', rels)}
    entries = {rid: full for full, rid in re.findall(r'(<p:sldId id="\d+" r:id="(rId\d+)"/>)', pres)}
    order = "".join(entries[file2rid[p.stem]] for _, _, p in new_slides)
    pres = re.sub(r"(<p:sldIdLst>).*?(</p:sldIdLst>)", lambda m: m.group(1) + order + m.group(2), pres, flags=re.S)
    pres_p.write_text(pres, encoding="utf-8")

    run([sys.executable, str(PPTX_SCRIPTS / "clean.py"), str(workdir)])

    out = Path(args.output).resolve()
    out.unlink(missing_ok=True)
    subprocess.run(["zip", "-qXr", str(out), "."], cwd=workdir, check=True)
    shutil.rmtree(workdir)
    print(f"완료: {out} ({len(flat)}장)")


if __name__ == "__main__":
    main()
