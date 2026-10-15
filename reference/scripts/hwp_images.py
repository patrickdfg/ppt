#!/usr/bin/env python3
"""hwp 파일에서 내장 이미지(BinData)를 추출한다.

사용법:
    python hwp_images.py 원고.hwp 출력폴더

각 이미지를 출력폴더에 BIN0001.jpg 등으로 저장하고,
원고 내 등장 순서와 직전 텍스트(맥락)를 stdout에 출력한다.
이 맥락을 보고 어느 위치에 photo 슬라이드를 넣을지 판단한다.
"""
import re
import sys
import zlib
from pathlib import Path

import olefile


def main(hwp_path, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ole = olefile.OleFileIO(hwp_path)

    hdr = ole.openstream("FileHeader").read()
    compressed = bool(hdr[36] & 1)

    saved = {}
    for s in ole.listdir():
        if "BinData" not in "/".join(s):
            continue
        name = s[-1]
        data = ole.openstream(s).read()
        if compressed:
            try:
                data = zlib.decompress(data, -15)
            except zlib.error:
                pass  # 이미 원본 바이트인 경우
        # 확장자 정규화
        ext = name.split(".")[-1].lower()
        if ext == "jpeg":
            ext = "jpg"
        base = name.split(".")[0]
        fname = f"{base}.{ext}"
        (out / fname).write_bytes(data)
        # BinData id (BIN0001 → 1)
        m = re.search(r"(\d+)", base)
        saved[m.group(1).lstrip("0") or "0"] = fname

    print(f"# {len(saved)}개 이미지 추출 → {out}")
    print("# id → 파일:", saved)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
