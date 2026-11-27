"""
Checks the Python engine against values Excel itself calculated.

    pip install pytest openpyxl
    pytest -v

Case 1 = the workbook as saved. Case 2 = different yellow inputs, recalculated
by a spreadsheet engine. If the workbook is found next to the app folder, an
extra test also reads it live, so a formula change in Excel shows up here.
"""
import json
import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "engine"))
sys.path.insert(0, str(ROOT / "tools"))

from prop_sst_body import SstBodyInputs, calculate, extn_label  # noqa: E402

GOLDEN = json.loads((ROOT / "tests" / "expected_sst_body.json").read_text(encoding="utf-8"))
WORKBOOK = ROOT.parent / "PROP_Calculator(Square) .xlsx"
TOL = 1e-6


def _check(case):
    res = calculate(SstBodyInputs.from_dict(case["inputs"]))
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
def test_matches_live_workbook():
    from export_expected import export
    _check(export(str(WORKBOOK)))


def test_table_shown_to_2_decimals_matches_sheet():
    """The drawing shows the table rounded to 2 decimals (Excel format 0.00)."""
    res = calculate(SstBodyInputs())
    assert f"{res.rows[0].l1:.2f}" == "23331.48"
    assert f"{res.rows[2].l2:.2f}" == "25365.76"   # +0M at conc level = F91
    assert f"{res.rows[8].l3_diag:.2f}" == "45373.28"


def test_zero_extension_l2_equals_base_width():
    inp = SstBodyInputs(base_width_conc_lvl=19999.99)
    assert calculate(inp).rows[2].l2 == pytest.approx(19999.99)


def test_diagonal_is_side_times_root2():
    for r in calculate(SstBodyInputs()).rows:
        assert r.l2_diag == pytest.approx(r.l2 * math.sqrt(2))


def test_zero_slope_vertical_rejected():
    with pytest.raises(ValueError):
        calculate(SstBodyInputs(slope_vertical=0))


@pytest.mark.parametrize("ext,label", [(-3000, "-3.0M BODY EXTN"), (0, "+0.0M BODY EXTN"), (9000, "+9.0M BODY EXTN")])
def test_extension_labels(ext, label):
    assert extn_label(ext) == label
