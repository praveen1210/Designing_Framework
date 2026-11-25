# Design Framework – PROP TYPE SST (BODY + LEG)

Fill in the yellow values and get the full GA sheet (drawing + body extension
table) as PNG/PDF, plus an AutoCAD data file so PROP1 runs without prompts.
Both sheets of `PROP_Calculator(Square) .xlsx` are covered: **SST-BODY** and **SST-LEG**.

```
app/
  web/index.html            <- open this in Chrome/Edge (double-click, no install)
  web/prop_engine.js        calculation engine (browser), both sheets
  web/sst_body_template.js  BODY sheet background + label positions (generated)
  web/sst_leg_template.js   LEG sheet background + label positions (generated)
  engine/prop_sst.py        same engine in Python (for batch / DXF later)
  lisp/PROP1_AUTO.LSP       PROP1 without manual prompts (BODY and LEG)
  tests/                    checks both engines against Excel-calculated values
  tools/export_expected.py  refresh the golden test values from the Excel
  tools/build_template.py   rebuild a sheet template when the Excel layout changes
```

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
5. **Save inputs / Load inputs** keep a .json per tower type (the file remembers the sheet).
6. **AutoCAD data** → in AutoCAD: `APPLOAD` `lisp/PROP1_AUTO.LSP` once,
   load the downloaded `PROP1_DATA_*.lsp`, then type `PROP1AUTO`
   (table at the X,Y you entered) or `PROP1PICK` (pick the point).
   The data file says which sheet it is, so the LEG table is drawn with the LEG formulas.

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

When the drawing or cell layout changes (needs LibreOffice):
`python tools/build_template.py "../PROP_Calculator(Square) .xlsx" --sheet SST-BODY`
(and again with `--sheet SST-LEG`).

## Notes

- The drawing is the Excel picture with new labels on it — labels change, geometry
  doesn't. When the GA drawing is available as DWG/DXF, step 3 replaces this with
  real DXF output from `engine/prop_sst.py`.
- Hole gauge uses `150 - T35` exactly like Excel N88; the 150 is fixed in the sheet,
  it does not follow the angle size text (S37).
