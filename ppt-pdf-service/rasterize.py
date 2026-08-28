# -*- coding: utf-8 -*-
"""把 blank.pdf 栅格化为每页 JPG 背景图（一次性预处理）。"""
import fitz  # PyMuPDF
from pathlib import Path

BASE = Path(__file__).resolve().parent / "php" / "assets"
doc = fitz.open(str(BASE / "blank.pdf"))
out_dir = BASE / "slides"
out_dir.mkdir(exist_ok=True)

ZOOM = 2.0  # 960x540pt -> 1920x1080px
for i, page in enumerate(doc, 1):
    pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
    out = out_dir / f"slide-{i}.jpg"
    pix.save(str(out), jpg_quality=88)
    print(f"{out.name}: {pix.width}x{pix.height}")
print("done")
