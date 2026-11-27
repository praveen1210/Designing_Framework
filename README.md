# Design Framework – PROP TYPE SST (BODY + LEG)

Fill in the yellow values and get the full GA sheet (drawing + body extension
table) as PNG/PDF, or your own Excel workbook filled with the same values.
Both sheets of `PROP_Calculator(Square) .xlsx` are covered: **SST-BODY** and **SST-LEG**.

```
app/
  web/index.html            <- open this in Chrome/Edge (double-click, no install)
  web/prop_engine.js        calculation engine (browser), both sheets
  web/sst_body_template.js  BODY sheet background + label positions (generated)
  web/sst_leg_template.js   LEG sheet background + label positions (generated)
  web/excel_export.js       Export Excel: fills the workbook in the browser
  web/excel_template.js     the workbook itself, embedded (generated)
  engine/prop_sst.py        same engine in Python (for batch / DXF later)
  lisp/PROP1_AUTO.LSP       PROP1 without manual prompts (parked: the page's AutoCAD
                            button is removed until the AutoCAD requirement is final)
  tests/                    checks both engines against Excel-calculated values
  tools/export_expected.py  refresh the golden test values from the Excel
  tools/build_template.py   rebuild a sheet template when the Excel layout changes
  tools/build_excel_template.py  re-embed the workbook for Export Excel
```

## Hosted

https://praveen1210.github.io/Designing_Framework/ (GitHub Pages, branch `pkr-sample`, folder `/ (root)`;
the root `index.html` opens `app/web/`). Every push to `pkr-sample` updates the site in about a minute.

## Daily use

1. Open `web/index.html`. Pick **SST-BODY** or **SST-LEG** at the top.
   Each sheet keeps its own values (remembered in the browser).
2. Type the yellow values in the left panel. The panel stays fixed; the drawing,
   table and calculation summary on the right update straight away.
3. Drawing controls (top-right of the drawing):
   - `−` / `+` zoom (or Ctrl + mouse wheel), `Fit`, `1:1`; drag with the mouse to pan
   - `▁` minimize / restore the drawing
   - `▢` maximize the drawing next to the input panel (double-click the drawing does the same)
   - `⛶` full window (hides the input panel)
   - `Esc` restores from maximize / full window
4. **Export PDF** → print dialog → *Save as PDF*, paper **A3**, landscape, margins *None*.
   **Export PNG** → full-size A3 image (200 dpi).
5. Title block and labels are editable too: revision row, client, project, names and
   dates, doc. no. (one character per box, 20 boxes), the TOWER SLOPE label and the
   balloon number in the circle next to the angle size (51 — this was empty in the Excel).
   Grey fields follow the Excel formula (e.g. revision "Designed" = DSGN name,
   CHKD/APPD dates = DSGN date); type to override, clear to follow again.
6. **Save inputs / Load inputs** keep a .json per tower type (the file remembers the sheet).
7. **Export Excel** → `PROP_Calculator_<date>.xlsx`: your workbook with both sheets
   filled (each with its own tab's values). Formulas stay; Excel recalculates when
   the file opens (it may ask to save on close — that's normal). Pictures, shapes and
   print setup are unchanged, and the balloon circle gets its number (51).
   Title-block grey fields left blank keep their Excel formula.

## BODY vs LEG

Same inputs and calculation block; only F91 and the table formulas differ:

| | BODY | LEG |
|---|---|---|
| F91 saved in Excel | 25365.76 | 455055 |
| L2 | F91 + 2tan·ext | (F91 + 2tan·ext) / 2 |
| L1 | L2 − 2tan·H1 | (F91 + 2tan·ext − 2tan·H1) / 2 |
| L3 | L2 + 2tan·H3 | (F91 + 2tan·ext + 2tan·H3) / 2 |

## Tests

```
pip install -r requirements.txt
pytest -v                                 # Python engine vs Excel, both sheets
node --test tests/engine_parity.test.js   # browser engine vs Excel, both sheets
```

For each sheet: the workbook as saved, plus different inputs that a
spreadsheet engine recalculated. If the workbook sits next to `app/`, extra
tests read it live, so a formula change in Excel fails the tests until the
engines are updated.

When the Excel formulas change: update both engines, then
`python tools/export_expected.py "../PROP_Calculator(Square) .xlsx" tests/expected_sst.json`.

When the workbook changes in any way, re-embed it for Export Excel:
`python tools/build_excel_template.py "../PROP_Calculator(Square) .xlsx"`

When the drawing or cell layout changes (needs LibreOffice):
`python tools/build_template.py "../PROP_Calculator(Square) .xlsx" --sheet SST-BODY`
(and again with `--sheet SST-LEG`).

## Notes

- The drawing is the Excel picture with new labels on it — labels change, geometry
  doesn't. When the GA drawing is available as DWG/DXF, step 3 replaces this with
  real DXF output from `engine/prop_sst.py`.
- Excel quirks kept as-is (override them in the page if they are wrong):
  revision "CHECKED" (AC60) is `=L77`, the label "CHKD", not a name;
  revision "DRAWN" (AE60) is `=N77`, the CHKD name, not the DRWN name (N76).
- Hole gauge uses `150 - T35` exactly like Excel N88; the 150 is fixed in the sheet,
  it does not follow the angle size text (S37).
