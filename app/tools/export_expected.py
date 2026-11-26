"""
Read the inputs and calculated results that Excel saved in the workbook and
write them as a golden test case (JSON).

    python tools/export_expected.py "../PROP_Calculator(Square) .xlsx" tests/expected_sst.json

Run it again whenever the Excel formulas change, then run the tests.
"""
import json
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")  # openpyxl warns about the embedded WMF picture

INPUT_CELLS = {
    "top_of_stub_height": "C18", "conc_to_gl": "C25", "stub_below_gl": "C40",
    "leg_width": "B32", "leg_size_text_2": "B31", "leg_size_text_3": "B30",
    "hole_text": "Y30", "slotted_hole_text": "N33", "plate_dim_1": "W31",
    "plate_dim_2": "W32", "hole_gap": "T35", "angle_size_text": "S37",
    "slope_vertical": "AD32", "slope_horizontal": "AE37",
    "base_width_conc_lvl": "F91",
}
RESULT_CELLS = {
    "slope_face_length": "AF32", "slope_develop_length": "AG31",
    "slope_horizontal_2": "AH34", "gauge_total": "N88", "gauge_half": "N89",
    "conc_to_bottom_of_stub": "I87", "bottom_vertical_diff": "I88",
    "tan": "I90", "angle_deg": "J90", "two_tan": "I91", "face": "I92",
    "face_x2200": "J92", "develop": "I93", "develop_x2200": "J93",
    "second_slope_tan": "I95", "second_slope_deg": "J95",
    "second_slope_cos": "I97", "l1_height": "F93", "l3_height": "F94",
}
LABEL_CELLS = ("AD40 L60 P60 R60 AA60 AC60 AE60 AG60 AI60 L65 N71 N75 N76 N77 N78 P75 P76 P77 P78".split()
               + [f"{c}80" for c in "N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG".split()])
ROW_COLS = {"l1": "E", "l1_diag": "F", "l2": "G", "l2_diag": "H", "l3": "I", "l3_diag": "J"}


SHEETS = {"SST-BODY": "PROP TYPE SST-BODY", "SST-LEG": "PROP TYPE SST-LEG"}


def export(xlsx_path: str, sheet: str = "SST-BODY") -> dict:
    ws = openpyxl.load_workbook(xlsx_path, data_only=True)[SHEETS[sheet]]
    case = {
        "sheet": sheet,
        "inputs": {k: ws[c].value for k, c in INPUT_CELLS.items()},
        "results": {k: ws[c].value for k, c in RESULT_CELLS.items()},
        "rows": [
            {"label": ws[f"D{r}"].value, **{k: ws[f"{col}{r}"].value for k, col in ROW_COLS.items()}}
            for r in range(74, 83)
        ],
        # text Excel shows in the title block (numbers as Excel prints them, e.g. 0 -> "0")
        "labels": {c: ("" if ws[c].value is None else str(ws[c].value)) for c in LABEL_CELLS},
    }
    return case


if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "tests/expected_sst.json"
    cases = []
    for sh in SHEETS:
        c = export(src, sh)
        c["name"] = f"{sh}: Excel as saved"
        cases.append(c)
    json.dump({"source": src, "cases": cases}, open(out, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    print("written", out)
