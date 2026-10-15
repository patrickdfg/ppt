#!/usr/bin/env python3
# hwpx 설교 원고에서 글자 색상별 텍스트를 뽑아 hwp_colors.py와 같은 형식으로 출력한다
"""
사용법 (설치 필요 없음):
    python hwpx_colors.py 원고.hwpx

출력: 단락별로 [BLUE]/[GREEN]/[RED]/[BLACK] 라벨이 붙은 텍스트, 단락 끝에 ---.
  BLUE(#0000ff)  → caption (아래자막)
  GREEN(#008000) → word / verse / plain (말씀·성경본문)
  RED(#ff0000)   → 빨간 글씨 구간
  BLACK          → 자막에 넣지 않는 서술 (문맥 참고용)
정확히 같은 색이 아니면 가장 가까운 계열(파랑/초록/빨강/검정)로 분류하고,
어느 계열도 아니면 색 코드(#RRGGBB)를 그대로 라벨로 쓴다.
"""
import re
import sys
import zipfile
from defusedxml import ElementTree as ET

LABEL = {"#0000ff": "BLUE", "#008000": "GREEN", "#000000": "BLACK", "#ff0000": "RED"}


def local(tag):
    return tag.rsplit("}", 1)[-1]


def classify(color):
    c = (color or "#000000").lower()
    if c in LABEL:
        return LABEL[c]
    m = re.fullmatch(r"#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})", c)
    if not m:
        return "BLACK"
    r, g, b = (int(x, 16) for x in m.groups())
    if max(r, g, b) < 80:
        return "BLACK"
    if b > r + 60 and b > g + 40:
        return "BLUE"
    if g > r + 50 and g > b + 30:
        return "GREEN"
    if r > g + 80 and r > b + 80:
        return "RED"
    return c.upper()


def run_text(run):
    """hp:run 안의 hp:t 글자. 탭은 공백, 줄바꿈은 그대로 이어 붙인다."""
    out = []
    for t in run:
        if local(t.tag) != "t":
            continue
        if t.text:
            out.append(t.text)
        for child in t:
            if local(child.tag) == "tab":
                out.append(" ")
            if child.tail:
                out.append(child.tail)
    return "".join(out)


def section_names(z):
    names = [n for n in z.namelist() if re.fullmatch(r"Contents/section\d+\.xml", n)]
    return sorted(names, key=lambda n: int(re.search(r"(\d+)", n).group(1)))


def main(path):
    z = zipfile.ZipFile(path)
    header = ET.fromstring(z.read("Contents/header.xml"))
    id2color = {}
    for el in header.iter():
        if local(el.tag) == "charPr":
            id2color[el.get("id")] = el.get("textColor", "#000000")

    for sec in section_names(z):
        root = ET.fromstring(z.read(sec))
        # 표 안의 단락도 문서 순서대로 나온다 (iter는 깊이 우선 = 문서 순서)
        for p in root.iter():
            if local(p.tag) != "p":
                continue
            runs = []
            for run in p:                      # 이 단락에 직접 속한 글자만
                if local(run.tag) != "run":
                    continue
                txt = run_text(run)
                if txt.strip():
                    color = id2color.get(run.get("charPrIDRef"), "#000000")
                    runs.append((classify(color), txt))
            if not runs:
                continue
            merged = []
            for c, t in runs:                  # 같은 색 연속 런 병합
                if merged and merged[-1][0] == c:
                    merged[-1][1] += t
                else:
                    merged.append([c, t])
            for c, t in merged:
                print(f"[{c}] {t}")
            print("---")


if __name__ == "__main__":
    main(sys.argv[1])
