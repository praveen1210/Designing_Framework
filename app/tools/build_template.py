"""
Rebuild the drawing template (web/sst_body_template.js or web/sst_leg_template.js) from the Excel sheet.

Run this when the drawing or the cell layout in the Excel changes
(or copy it for the SST-LEG sheet later).

What it does:
 1. Copies the workbook and blanks the label cells (yellow inputs + calculated
    labels + the table) at XML level, so pictures and shapes are kept.
 2. Copies it again with each label cell filled with its own marker colour.
 3. Prints both to PDF with LibreOffice and rasterises the sheet's page at 200 dpi.
 4. Finds each marker colour -> pixel box of that cell on the page.
 5. Removes leftover yellow from the background and writes the template JS.

Needs: LibreOffice (soffice on PATH or --soffice), pip install openpyxl pymupdf pillow numpy

    python tools/build_template.py "../PROP_Calculator(Square) .xlsx" --sheet SST-BODY
    python tools/build_template.py "../PROP_Calculator(Square) .xlsx" --sheet SST-LEG
"""
import argparse
import base64
import io
import json
import re
import shutil
import subprocess
import tempfile
import warnings
import zipfile
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:  # older versions
    import fitz
import numpy as np
import openpyxl
from PIL import Image

warnings.filterwarnings("ignore")

SHEETS = {"SST-BODY": "PROP TYPE SST-BODY", "SST-LEG": "PROP TYPE SST-LEG"}
LABEL_CELLS = ("C18 C25 B30 B31 B32 Y30 W31 W32 AD32 AE37 S37 T35 N33 C40 "
               "AG31 AF32 AH34 S35 U35").split() + [f"{c}{r}" for r in range(74, 83) for c in "EFGHIJ"]
DPI = 200


def sheet_xml_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb).group(1)
    target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return "xl/" + target.group(1).lstrip("/").replace("xl/", "")


def rewrite(src: Path, dst: Path, sheet_xml: str, edit):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            data = zi.read(item.filename)
            if item.filename == sheet_xml:
                data = edit(data.decode("utf8")).encode("utf8")
            elif item.filename == "[Content_Types].xml":
                data = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", data)
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"", data)
            zo.writestr(item, data)


def blank_cells(xml: str, style_for=None) -> str:
    for c in LABEL_CELLS:
        def repl(m):
            s = style_for(c) if style_for else (m.group(1) or "")
            return f'<c r="{c}"{s}/>'
        xml, n = re.subn(r'<c r="%s"( s="\d+")?[^>]*?(?:/>|>.*?</c>)' % c, repl, xml, count=1)
        if n != 1:
            raise SystemExit(f"cell {c} not found in sheet XML")
    return xml


def add_marker_styles(src: Path, dst: Path):
    """Append one solid-fill cell style per label cell; return {cell: style index}."""
    with zipfile.ZipFile(src) as z:
        styles = z.read("xl/styles.xml").decode("utf8")
    fills_m = re.search(r'<fills count="(\d+)">(.*?)</fills>', styles, re.S)
    xfs_m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    nf, nx = int(fills_m.group(1)), int(xfs_m.group(1))
    colors, new_fills, new_xfs, style_idx = {}, "", "", {}
    for i, c in enumerate(LABEL_CELLS):
        rgb = "%02X7BE7" % (i * 3 + 3)
        colors[c] = rgb
        new_fills += f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/><bgColor indexed="64"/></patternFill></fill>'
        new_xfs += f'<xf numFmtId="0" fontId="0" fillId="{nf + i}" borderId="0" xfId="0" applyFill="1"/>'
        style_idx[c] = nx + i
    styles = styles.replace(fills_m.group(0), f'<fills count="{nf + len(LABEL_CELLS)}">{fills_m.group(2)}{new_fills}</fills>')
    styles = styles.replace(xfs_m.group(0), f'<cellXfs count="{nx + len(LABEL_CELLS)}">{xfs_m.group(2)}{new_xfs}</cellXfs>')
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            zo.writestr(item, styles.encode("utf8") if item.filename == "xl/styles.xml" else zi.read(item.filename))
    return colors, style_idx


def merged_cells(ws, cell):
    for m in ws.merged_cells.ranges:
        if cell in m:
            return [ws.cell(r, c).coordinate for r in range(m.min_row, m.max_row + 1)
                    for c in range(m.min_col, m.max_col + 1)]
    return [cell]


def to_png(xlsx: Path, soffice: str, out_dir: Path, page_no: int, n_sheets: int) -> np.ndarray:
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(xlsx)],
                   check=True, capture_output=True, timeout=300)
    pdf = fitz.open(out_dir / (xlsx.stem + ".pdf"))
    # each sheet is set to fit one A3 page, so page N = sheet N
    if len(pdf) != n_sheets:
        raise SystemExit(f"expected {n_sheets} pages (one per sheet), got {len(pdf)} - check print areas")
    pix = pdf[page_no].get_pixmap(dpi=DPI)
    return np.array(Image.frombytes("RGB", (pix.width, pix.height), pix.samples)).astype(int)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--sheet", default="SST-BODY", choices=list(SHEETS))
    ap.add_argument("--out", help="default: web/sst_body_template.js or web/sst_leg_template.js")
    ap.add_argument("--soffice", default=shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe")
    a = ap.parse_args()

    SHEET = SHEETS[a.sheet]
    out = a.out or str(Path(__file__).resolve().parents[1] / "web" / f"{a.sheet.lower().replace('-', '_')}_template.js")
    src = Path(a.xlsx)
    wb = openpyxl.load_workbook(src)
    ws = wb[SHEET]
    page_no, n_sheets = wb.sheetnames.index(SHEET), len(wb.sheetnames)
    tmp = Path(tempfile.mkdtemp())

    with zipfile.ZipFile(src) as z:
        sx = sheet_xml_path(z, SHEET)

    # 1. clean background
    clean = tmp / "clean.xlsx"
    rewrite(src, clean, sx, blank_cells)

    # 2. marker copy (merged ranges get the colour on every cell)
    styled = tmp / "styled.xlsx"
    colors, style_idx = add_marker_styles(src, styled)
    marker = tmp / "marker.xlsx"

    def mark(xml):
        xml = blank_cells(xml, lambda c: f' s="{style_idx[c]}"')
        for c in LABEL_CELLS:
            for mc in merged_cells(ws, c)[1:]:
                xml = re.sub(r'<c r="%s"( s="\d+")?' % mc, f'<c r="{mc}" s="{style_idx[c]}"', xml, count=1)
        return xml
    rewrite(styled, marker, sx, mark)

    bg = to_png(clean, a.soffice, tmp, page_no, n_sheets)
    mk = to_png(marker, a.soffice, tmp, page_no, n_sheets)

    # 3. boxes
    boxes = {}
    for c, hexc in colors.items():
        r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
        ys, xs = np.nonzero((abs(mk[..., 0] - r) <= 1) & (mk[..., 1] == g) & (mk[..., 2] == b))
        if len(xs) == 0:
            print("WARNING: cell not visible on page:", c)
            continue
        boxes[c] = [int(xs.min()), int(ys.min()), int(xs.max() + 1), int(ys.max() + 1)]

    # 4. remove leftover yellow fill around label cells
    R, G, B = bg[..., 0], bg[..., 1], bg[..., 2]
    for x0, y0, x1, y1 in boxes.values():
        sl = (slice(max(0, y0 - 3), y1 + 3), slice(max(0, x0 - 3), x1 + 3))
        yel = (R[sl] > 180) & (G[sl] > 180) & (B[sl] < 200) & ((R[sl] - B[sl]) > 40)
        bg[sl][yel] = 255

    # 5. font sizes: pixels per point from the table row height
    row_pt = ws.row_dimensions[74].height or 15
    px_per_pt = (boxes["E74"][3] - boxes["E74"][1]) / row_pt
    fields = {}
    for c, box in boxes.items():
        cell = ws[c]
        fields[c] = {"box": box, "rot": cell.alignment.textRotation or 0,
                     "align": cell.alignment.horizontal or "general",
                     "size": round((cell.font.sz or 11) * px_per_pt, 1),
                     "bold": bool(cell.font.b), "font": cell.font.name}

    img = Image.fromarray(bg.astype("uint8")).quantize(colors=48, method=Image.Quantize.MEDIANCUT,
                                                       dither=Image.Dither.NONE)
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    tpl = {"name": SHEET, "W": int(bg.shape[1]), "H": int(bg.shape[0]),
           "bg": "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(), "fields": fields}
    Path(out).write_text(
        "// Generated by tools/build_template.py from the Excel sheet. Background = sheet page with the\n"
        "// yellow/table cells blanked; fields = pixel box of each cell on that page (200 dpi, A3).\n"
        "window.TEMPLATES = window.TEMPLATES || {};\n"
        f'window.TEMPLATES["{a.sheet}"] = {json.dumps(tpl)};\n', encoding="utf8")
    print(f"written {out}  ({len(fields)} fields)")


if __name__ == "__main__":
    main()
