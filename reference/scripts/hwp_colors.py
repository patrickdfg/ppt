#!/usr/bin/env python3
"""hwp 설교 원고에서 글자 색상별 텍스트를 추출한다.

사용법:
    pip install pyhwp --break-system-packages   # 최초 1회
    hwp5proc xml 원고.hwp > 원고.xml
    python hwp_colors.py 원고.xml

출력: 단락별로 [BLUE]/[GREEN]/[BLACK] 라벨이 붙은 텍스트.
  BLUE(#0000ff)  → caption (아래자막)
  GREEN(#008000) → word / verse / plain (말씀·성경본문)
  BLACK          → 자막에 넣지 않는 서술 (문맥 참고용)
"""
import sys
from defusedxml import minidom

LABEL = {"#0000ff": "BLUE", "#008000": "GREEN", "#000000": "BLACK", "#ff0000": "RED"}


def main(path):
    doc = minidom.parse(path)
    cs_list = doc.getElementsByTagName("CharShape")
    id2color = {}
    for i, cs in enumerate(cs_list):
        attrs = dict(cs.attributes.items())
        id2color[i] = attrs.get("text-color", "#000000")

    for p in doc.getElementsByTagName("Paragraph"):
        runs = []
        for t in p.getElementsByTagName("Text"):
            cid = int(t.getAttribute("charshape-id") or 0)
            txt = t.firstChild.data if t.firstChild else ""
            if txt.strip():
                color = id2color.get(cid, "#000000")
                runs.append((LABEL.get(color, color), txt))
        if not runs:
            continue
        # 같은 색 연속 런 병합
        merged = []
        for c, t in runs:
            if merged and merged[-1][0] == c:
                merged[-1][1] += t
            else:
                merged.append([c, t])
        for c, t in merged:
            print(f"[{c}] {t}")
        print("---")


if __name__ == "__main__":
    main(sys.argv[1])
