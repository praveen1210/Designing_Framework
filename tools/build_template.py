"""
Rebuild the drawing template (web/sst_body_template.js or web/sst_leg_template.js) from the Excel sheet.

Run this when the drawing or the cell layout in the Excel changes
(or copy it for the SST-LEG sheet later).

What it does:
 1. Copies the workbook and blanks the label cells (yellow inputs, calculated
    labels, the table, title-block fields) at XML level, so pictures and shapes are kept.
 2. Copies it again with each label cell filled with its own marker colour.
    Shapes that get a text field (the empty balloon circle) are filled with a
    marker colour too.
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
LABEL_CELLS = (
    # drawing labels: yellow inputs + values calculated from them
    "C18 C25 B30 B31 B32 Y30 W31 W32 AD32 AE37 S37 T35 N33 C40 AG31 AF32 AH34 S35 U35 AD40 "
    # title block: revision row, client, project, names/dates
    "L60 P60 R60 AA60 AC60 AE60 AG60 AI60 L65 N71 N75 N76 N77 N78 P75 P76 P77 P78"
).split() + [f"{c}{r}" for r in range(74, 83) for c in "EFGHIJ"] \
  + [f"{c}80" for c in "N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG".split()]   # doc no., 1 char per cell

# Rotated dimension values stacked along one vertical dimension line, top to bottom.
# Excel top-aligns them in their cells; on the drawing each value belongs centred
# between two dimension ticks, so the ticks are found in the drawing itself.
VERTICAL_DIM_STACKS = [["W31", "W32"]]   # plate dims 70 / 50

# Shapes that get a text field: field name -> preset geometry of the shape in the drawing
SHAPE_FIELDS = {"BALLOON": "flowChartConnector"}   # empty circle next to the angle size (item no. 51)
DPI = 200


def sheet_xml_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb).group(1)
    target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return "xl/" + target.group(1).lstrip("/").replace("xl/", "")


def drawing_xml_path(z: zipfile.ZipFile, sheet_xml: str):
    rels_path = sheet_xml.replace("worksheets/", "worksheets/_rels/") + ".rels"
    try:
        rels = z.read(rels_path).decode("utf8")
    except KeyError:
        return None
    m = re.search(r'Target="([^"]*drawings/[^"]+)"', rels)
    return "xl/drawings/" + m.group(1).split("drawings/")[1] if m else None


def marker_rgb(i: int) -> str:
    return "%02X%02XE7" % (3 + 3 * (i % 80), 0x7B + 8 * (i // 80))


def rewrite(src: Path, dst: Path, edits: dict):
    """Copy the xlsx, applying edits {part path: fn(xml) -> xml}."""
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            data = zi.read(item.filename)
            if item.filename in edits:
                data = edits[item.filename](data.decode("utf8")).encode("utf8")
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
        rgb = marker_rgb(i)
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


def vertical_dim_spans(bg: np.ndarray, boxes: dict, cells: list) -> dict:
    """Find the ticks of the vertical dimension line just left of the stacked cells and
    return {cell: [y_top, y_bottom]} = the gap between consecutive ticks."""
    if not all(c in boxes for c in cells):
        return {}
    dark = bg.mean(axis=2) < 140
    x0 = boxes[cells[0]][0]
    y_top, y_bot = boxes[cells[0]][1], boxes[cells[-1]][3]
    line_x = max(range(x0 - 30, x0), key=lambda x: dark[y_top:y_bot, x].sum())
    rows = [y for y in range(y_top - 15, y_bot + 15) if dark[y, line_x - 20:line_x - 2].mean() > 0.9]
    ticks = []
    for y in rows:                       # merge neighbouring rows of one thick line
        if ticks and y - ticks[-1][-1] <= 2:
            ticks[-1].append(y)
        else:
            ticks.append([y])
    ticks = [sum(t) / len(t) for t in ticks]
    if len(ticks) != len(cells) + 1:
        print(f"WARNING: expected {len(cells) + 1} dimension ticks for {cells}, found {len(ticks)}; keeping Excel position")
        return {}
    return {c: [round(ticks[i]), round(ticks[i + 1])] for i, c in enumerate(cells)}


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
        dx = drawing_xml_path(z, sx)

    # 1. clean background
    clean = tmp / "clean.xlsx"
    rewrite(src, clean, {sx: blank_cells})

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
    shape_colors = {name: marker_rgb(len(LABEL_CELLS) + i) for i, name in enumerate(SHAPE_FIELDS)}

    def mark_shapes(xml):
        for name, prst in SHAPE_FIELDS.items():
            i = xml.find(f'prst="{prst}"')
            if i < 0:
                print("WARNING: shape not found:", name)
                continue
            end = xml.find("</a:prstGeom>", i) + len("</a:prstGeom>")
            xml = xml[:end] + f'<a:solidFill><a:srgbClr val="{shape_colors[name]}"/></a:solidFill>' + xml[end:]
        return xml
    edits = {sx: mark}
    if dx:
        edits[dx] = mark_shapes
    rewrite(styled, marker, edits)

    bg = to_png(clean, a.soffice, tmp, page_no, n_sheets)
    mk = to_png(marker, a.soffice, tmp, page_no, n_sheets)

    # 3. boxes
    boxes = {}
    for c, hexc in {**colors, **shape_colors}.items():
        r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
        ys, xs = np.nonzero((abs(mk[..., 0] - r) <= 1) & (mk[..., 1] == g) & (mk[..., 2] == b))
        if len(xs) == 0:
            print("WARNING: cell not visible on page:", c)
            continue
        boxes[c] = [int(xs.min()), int(ys.min()), int(xs.max() + 1), int(ys.max() + 1)]

    # 4. remove leftover yellow fill around the yellow (input) cells only - other
    #    boxes (e.g. the client name) sit next to coloured logos
    R, G, B = bg[..., 0], bg[..., 1], bg[..., 2]
    yellow_cells = [c for c in LABEL_CELLS if c in boxes and ws[c].fill.fill_type
                    and str(ws[c].fill.fgColor.rgb) == "FFFFFF00"]
    for x0, y0, x1, y1 in (boxes[c] for c in yellow_cells):
        sl = (slice(max(0, y0 - 3), y1 + 3), slice(max(0, x0 - 3), x1 + 3))
        yel = (R[sl] > 180) & (G[sl] > 180) & (B[sl] < 200) & ((R[sl] - B[sl]) > 40)
        bg[sl][yel] = 255

    # 5. font sizes: pixels per point from the table row height
    row_pt = ws.row_dimensions[74].height or 15
    px_per_pt = (boxes["E74"][3] - boxes["E74"][1]) / row_pt
    fields = {}
    for c, box in boxes.items():
        if c in SHAPE_FIELDS:
            h = box[3] - box[1]
            fields[c] = {"box": box, "rot": 0, "align": "center", "size": round(h * 0.5, 1),
                         "bold": False, "font": "Arial Narrow", "wrap": False}
            continue
        cell = ws[c]
        fields[c] = {"box": box, "rot": cell.alignment.textRotation or 0,
                     "align": cell.alignment.horizontal or "general",
                     "size": round((cell.font.sz or 11) * px_per_pt, 1),
                     "bold": bool(cell.font.b), "font": cell.font.name,
                     "wrap": bool(cell.alignment.wrapText)}

    for stack in VERTICAL_DIM_STACKS:
        for c, span in vertical_dim_spans(bg, boxes, stack).items():
            fields[c]["span"] = span      # page centres the value between these two ticks

    img = Image.fromarray(bg.astype("uint8")).quantize(colors=48, method=Image.Quantize.MEDIANCUT,
                                                       dither=Image.Dither.NONE)
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    tpl = {"name": SHEET, "W": int(bg.shape[1]), "H": int(bg.shape[0]),
           "bg": "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(), "fields": fields}
    Path(out).write_text(
        "// Generated by tools/build_template.py from the Excel sheet. Background = sheet page with the\n"
        "// label cells blanked; fields = pixel box of each cell (or shape) on that page (200 dpi, A3).\n"
        "window.TEMPLATES = window.TEMPLATES || {};\n"
        f'window.TEMPLATES["{a.sheet}"] = {json.dumps(tpl)};\n', encoding="utf8")
    print(f"written {out}  ({len(fields)} fields)")


if __name__ == "__main__":
    main()
