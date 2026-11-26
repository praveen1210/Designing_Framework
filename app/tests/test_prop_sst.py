"""
Checks the Python engine against values Excel itself calculated, for both
sheets (SST-BODY and SST-LEG).

    pip install pytest openpyxl
    pytest -v

For each sheet: case 1 = the workbook as saved, case 2 = different yellow
inputs, recalculated by a spreadsheet engine. If the workbook is found next to
the app folder, extra tests read it live, so a formula change in Excel shows up here.
"""
import json
import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "engine"))
sys.path.insert(0, str(ROOT / "tools"))

from prop_sst import SstInputs, calculate, default_inputs, extn_label  # noqa: E402

GOLDEN = json.loads((ROOT / "tests" / "expected_sst.json").read_text(encoding="utf-8"))
WORKBOOK = ROOT.parent / "PROP_Calculator(Square) .xlsx"
TOL = 1e-6


def _check(case):
    sheet = case["sheet"]
    res = calculate(SstInputs.from_dict(case["inputs"], sheet), sheet)
    for key, expected in case["results"].items():
        assert getattr(res, key) == pytest.approx(expected, abs=TOL), key
    assert len(res.rows) == len(case["rows"]) == 9
    for got, exp in zip(res.rows, case["rows"]):
        assert got.label == exp["label"]
        for key in ("l1", "l1_diag", "l2", "l2_diag", "l3", "l3_diag"):
            assert getattr(got, key) == pytest.approx(exp[key], abs=TOL), f"{exp['label']} {key}"


@pytest.mark.parametrize("case", GOLDEN["cases"], ids=[c["name"] for c in GOLDEN["cases"]])
def test_matches_excel_golden(case):
    _check(case)


@pytest.mark.skipif(not WORKBOOK.exists(), reason="workbook not next to app folder")
@pytest.mark.parametrize("sheet", ["SST-BODY", "SST-LEG"])
def test_matches_live_workbook(sheet):
    from export_expected import export
    _check(export(str(WORKBOOK), sheet))


@pytest.mark.parametrize("sheet,f91", [("SST-BODY", 25365.76), ("SST-LEG", 455055)])
def test_defaults_match_excel(sheet, f91):
    assert default_inputs(sheet).base_width_conc_lvl == f91


def test_body_table_shown_to_2_decimals():
    res = calculate(SstInputs())
    assert f"{res.rows[0].l1:.2f}" == "23331.48"
    assert f"{res.rows[2].l2:.2f}" == "25365.76"   # +0M at conc level = F91
    assert f"{res.rows[8].l3_diag:.2f}" == "45373.28"


def test_leg_is_half_of_full_width():
    inp = default_inputs("SST-LEG")
    leg = calculate(inp, "SST-LEG")
    assert leg.rows[2].l2 == pytest.approx(inp.base_width_conc_lvl / 2)
    assert f"{leg.rows[0].l1:.2f}" == "226510.36"


def test_diagonal_is_side_times_root2():
    for sheet in ("SST-BODY", "SST-LEG"):
        for r in calculate(default_inputs(sheet), sheet).rows:
            assert r.l2_diag == pytest.approx(r.l2 * math.sqrt(2))


def test_bad_input_rejected():
    with pytest.raises(ValueError):
        calculate(SstInputs(slope_vertical=0))
    with pytest.raises(ValueError):
        calculate(SstInputs(), "SST-ARM")


@pytest.mark.parametrize("ext,label", [(-3000, "-3.0M BODY EXTN"), (0, "+0.0M BODY EXTN"), (9000, "+9.0M BODY EXTN")])
def test_extension_labels(ext, label):
    assert extn_label(ext) == label


# ---------------- title block / drawing labels ----------------
from prop_sst import title_block  # noqa: E402


@pytest.mark.parametrize("case", GOLDEN["cases"], ids=[c["name"] for c in GOLDEN["cases"]])
def test_title_block_defaults_match_excel(case):
    labels = title_block(SstInputs.from_dict(case["inputs"], case["sheet"]))
    for cell, expected in case["labels"].items():
        assert labels[cell] == expected, cell


def test_title_block_blank_fields_follow_excel_links():
    tb = title_block(SstInputs(dsgn_name="ABC", chkd_name="XYZ", appd_name="PQR", dsgn_date="01.01.27"))
    assert tb["AA60"] == "ABC"           # =N75
    assert tb["AE60"] == "XYZ"           # =N77
    assert tb["AI60"] == "PQR"           # =N78
    assert tb["P77"] == tb["P78"] == "01.01.27"   # =P75, =P77


def test_title_block_overrides_win():
    tb = title_block(SstInputs(rev_checked="KMR", chkd_date="05.01.27", appd_date="06.01.27"))
    assert tb["AC60"] == "KMR"
    assert (tb["P77"], tb["P78"]) == ("05.01.27", "06.01.27")


def test_doc_no_one_char_per_cell():
    tb = title_block(SstInputs(doc_no="ABC"))
    assert (tb["N80"], tb["O80"], tb["P80"], tb["Q80"], tb["AG80"]) == ("A", "B", "C", "", "")


def test_balloon_default_is_51():
    assert title_block(SstInputs())["BALLOON"] == "51"
