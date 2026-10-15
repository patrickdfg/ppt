#!/usr/bin/env python3
# hwpx 원고에 들어 있는 그림을 꺼내고, 원고 속 등장 순서와 바로 앞 글(맥락)을 출력한다
"""
사용법 (설치 필요 없음):
    python hwpx_images.py 원고.hwpx 출력폴더

각 그림을 출력폴더에 원래 이름(image1.jpg 등)으로 저장하고,
본문에 나오는 순서대로 "순번. 파일 ← 직전 글" 을 출력한다.
이 맥락을 보고 어느 위치에 photo 슬라이드를 넣을지 판단한다.
"""
import re
import sys
import zipfile
from pathlib import Path
from defusedxml import ElementTree as ET


def local(tag):
    return tag.rsplit("}", 1)[-1]


def main(path, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    z = zipfile.ZipFile(path)

    # 그림 id → 압축 안 경로 (content.hpf 목록)
    id2href = {}
    try:
        hpf = ET.fromstring(z.read("Contents/content.hpf"))
        for el in hpf.iter():
            if local(el.tag) == "item" and el.get("href", "").startswith("BinData/"):
                id2href[el.get("id")] = el.get("href")
    except KeyError:
        pass
    for n in z.namelist():                     # 목록에 없더라도 BinData 파일은 모두 포함
        if n.startswith("BinData/") and not n.endswith("/"):
            id2href.setdefault(Path(n).stem, n)

    saved = {}
    for iid, href in id2href.items():
        name = Path(href).name
        if name.lower().endswith(".jpeg"):
            name = name[:-5] + ".jpg"
        (out / name).write_bytes(z.read(href))
        saved[iid] = name
    print(f"# {len(saved)}개 이미지 추출 → {out}")

    # 본문 등장 순서와 직전 글
    secs = sorted([n for n in z.namelist() if re.fullmatch(r"Contents/section\d+\.xml", n)],
                  key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    order, last_text = [], ""
    for sec in secs:
        root = ET.fromstring(z.read(sec))
        for el in root.iter():
            tag = local(el.tag)
            if tag == "t":
                txt = "".join(el.itertext()).strip()
                if txt:
                    last_text = txt
            elif tag == "img" and el.get("binaryItemIDRef"):
                order.append((saved.get(el.get("binaryItemIDRef"), el.get("binaryItemIDRef")), last_text))
    print("# 본문 등장 순서 (파일 ← 직전 글):")
    for i, (f, ctx) in enumerate(order, 1):
        print(f"{i}. {f} ← {ctx[:60]}")
    unused = set(saved.values()) - {f for f, _ in order}
    if unused:
        print("# 본문에서 쓰이지 않는 그림:", sorted(unused))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
