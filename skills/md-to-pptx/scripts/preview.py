#!/usr/bin/env python3
"""Render a .pptx to per-slide PNGs + one contact-sheet grid, for visual QA.

Uses PowerPoint (COM, Windows) when available, else LibreOffice (soffice → PDF → PyMuPDF).

  preview.py deck.pptx [--out DIR] [--cols 4]
"""
import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image


def via_powerpoint(pptx: Path, out: Path) -> bool:
    if sys.platform != "win32":
        return False
    ps = (
        "$pp = New-Object -ComObject PowerPoint.Application; "
        f"$p = $pp.Presentations.Open('{pptx}', $true, $false, $false); "
        f"$p.SaveAs('{out}', 18); $p.Close(); $pp.Quit()"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    return r.returncode == 0 and any(out.glob("*.PNG"))


def via_libreoffice(pptx: Path, out: Path) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx)],
                   capture_output=True)
    pdf = out / (pptx.stem + ".pdf")
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("PDF 已產生，但缺 PyMuPDF（pip install pymupdf）無法轉 PNG：", pdf)
        return False
    for n, page in enumerate(fitz.open(pdf), 1):
        page.get_pixmap(dpi=96).save(out / f"slide{n}.PNG")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out")
    ap.add_argument("--cols", type=int, default=4)
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    pptx = Path(a.pptx).resolve()
    out = Path(a.out).resolve() if a.out else pptx.parent / f"{pptx.stem}_preview"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    if not (via_powerpoint(pptx, out) or via_libreoffice(pptx, out)):
        sys.exit("無法渲染：需要 PowerPoint（Windows）或 LibreOffice。")
    files = sorted(out.glob("*.PNG"), key=lambda f: int(re.findall(r"(\d+)", f.stem)[-1]))
    w, h, gap = 480, 270, 10
    rows = -(-len(files) // a.cols)
    grid = Image.new("RGB", (a.cols * (w + gap) + gap, rows * (h + gap) + gap), "#888888")
    for k, f in enumerate(files):
        with Image.open(f) as im:
            grid.paste(im.convert("RGB").resize((w, h)), (gap + (k % a.cols) * (w + gap), gap + (k // a.cols) * (h + gap)))
    grid_path = out / "grid.png"
    grid.save(grid_path)
    print(f"✔ {len(files)} slides → {out}\n  grid: {grid_path}")


if __name__ == "__main__":
    main()
