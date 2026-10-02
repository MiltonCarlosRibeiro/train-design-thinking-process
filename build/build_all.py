"""Gera todos os arquivos do treinamento: slides (.pptx) e PDFs.

Uso:  python build_all.py
(Edite videos.json antes para incluir os links dos vídeos.)
"""
import os
import runpy
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # console cp932 não imprime "Ã"

BUILD = Path(__file__).resolve().parent
os.chdir(BUILD)
sys.path.insert(0, str(BUILD))

import brand as B  # noqa: E402

if not (B.ASSETS / "logo_full_t.png").exists():
    runpy.run_path(str(BUILD / "make_assets.py"))

import build_slides  # noqa: E402
from build_pdfs import build_all_pdfs  # noqa: E402

pptx = B.ROOT / "DT_na_Pratica_Slides.pptx"
print("slides:", build_slides.build(str(pptx)), "->", pptx)
build_all_pdfs()
