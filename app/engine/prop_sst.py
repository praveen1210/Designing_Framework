"""
PROP TYPE SST calculation engine for both sheets of
"PROP_Calculator(Square) .xlsx":

  SST-BODY  ('PROP TYPE SST-BODY')
  SST-LEG   ('PROP TYPE SST-LEG')

Both sheets have the same yellow inputs and the same calculation block
(rows 86-97). Only the body-extension table differs:

  BODY:  L2 = F91 + 2tan*ext        L1 = L2 - 2tan*H1        L3 = L2 + 2tan*H3
  LEG :  D  = F91 + 2tan*ext        L1 = (D - 2tan*H1)/2     L2 = D/2     L3 = (D + 2tan*H3)/2

Keep this file and web/prop_engine.js in step: both are checked against the
same golden values (tests/expected_sst.json).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, asdict, field, replace

SHEETS = {"SST-BODY": "PROP TYPE SST-BODY", "SST-LEG": "PROP TYPE SST-LEG"}
BODY_EXTENSIONS = [-3000, -1500, 0, 1500, 3000, 4500, 6000, 7500, 9000]

# Excel N88 is "=150-T35". 150 is hard-coded in the sheet (the angle flange
# width of L150X150X12). Kept identical to the Excel on purpose.
ANGLE_FLANGE_FOR_GAUGE = 150


@dataclass
class SstInputs:
    top_of_stub_height: float = 774.29     # C18  -> F93 (L1 height)
    conc_to_gl: float = 225                # C25
    stub_below_gl: float = 3300            # C40
    leg_width: float = 250                 # B32  (numeric, used in I86)
    leg_size_text_2: str = "X250"          # B31  (label only)
    leg_size_text_3: str = "X18(H)"        # B30  (label only)
    hole_text: str = "17.5Ø HOLE"          # Y30  (label only)
    slotted_hole_text: str = "17.5Ø HOLE\nSLOTTED HOLE"  # N33 (label only)
    plate_dim_1: float = 70                # W31  (label only)
    plate_dim_2: float = 50                # W32  (label only)
    hole_gap: float = 30                   # T35  -> N88/N89
    angle_size_text: str = "L150X150X12"   # S37  (label only)
    slope_vertical: float = 40595          # AD32
    slope_horizontal: float = 10940        # AE37
    base_width_conc_lvl: float = 25365.76  # F91  (+0M body extn at conc. level)

    @classmethod
    def from_dict(cls, d: dict, sheet: str = "SST-BODY") -> "SstInputs":
        base = default_inputs(sheet)
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        return replace(base, **known)


def default_inputs(sheet: str = "SST-BODY") -> SstInputs:
    """Values saved in the Excel for each sheet (only F91 differs)."""
    if sheet == "SST-LEG":
        return SstInputs(base_width_conc_lvl=455055)
    return SstInputs()


@dataclass
class SstRow:
    body_extn: int
    label: str
    l1: float
    l1_diag: float
    l2: float
    l2_diag: float
    l3: float
    l3_diag: float


@dataclass
class SstResult:
    sheet: str
    slope_face_length: float      # AF32 = ROUND(SQRT(AE37^2 + AD32^2), 3)
    slope_develop_length: float   # AG31 = ROUND(SQRT(AF32^2 + AH34^2), 3)
    slope_horizontal_2: float     # AH34 = AE37
    gauge_total: float            # N88 = 150 - T35
    gauge_half: float             # N89 = N88 / 2  (shown in S35 and U35)
    conc_to_bottom_of_stub: float  # I87 = C40 + C25
    bottom_vertical_diff: float    # I88 = (I95 * I86) / I93
    tan: float                     # I90 = AE37 / AD32
    angle_deg: float               # J90
    two_tan: float                 # I91 / F92
    face: float                    # I92 = AF32 / AD32
    face_x2200: float              # J92
    develop: float                 # I93 = AG31 / AD32
    develop_x2200: float           # J93
    second_slope_tan: float        # I95 = AH34 / AF32
    second_slope_deg: float        # J95
    second_slope_cos: float        # I97 = AG31 / AF32
    l1_height: float               # F93 = C18
    l3_height: float               # F94 = I87 - I88
    rows: list[SstRow] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


def extn_label(ext: int) -> str:
    sign = "-" if ext < 0 else "+"
    return f"{sign}{abs(ext) / 1000:.1f}M BODY EXTN"


def calculate(inp: SstInputs, sheet: str = "SST-BODY") -> SstResult:
    if sheet not in SHEETS:
        raise ValueError(f"Unknown sheet {sheet!r}, use one of {list(SHEETS)}")
    ad32 = float(inp.slope_vertical)
    ae37 = float(inp.slope_horizontal)
    if ad32 == 0:
        raise ValueError("Slope vertical (AD32) cannot be 0")

    ah34 = ae37
    af32 = round(math.sqrt(ae37 ** 2 + ad32 ** 2), 3)
    ag31 = round(math.sqrt(af32 ** 2 + ah34 ** 2), 3)

    n88 = ANGLE_FLANGE_FOR_GAUGE - float(inp.hole_gap)
    n89 = n88 / 2

    i86 = float(inp.leg_width)
    i87 = float(inp.stub_below_gl) + float(inp.conc_to_gl)
    i90 = ae37 / ad32
    i91 = i90 * 2
    i92 = af32 / ad32
    i93 = ag31 / ad32
    i95 = ah34 / af32
    i97 = ag31 / af32
    i88 = (i95 * i86) / i93

    f91 = float(inp.base_width_conc_lvl)
    f92 = i91
    f93 = float(inp.top_of_stub_height)
    f94 = i87 - i88

    rows = []
    for ext in BODY_EXTENSIONS:
        d = f91 + f92 * ext
        if sheet == "SST-LEG":
            e, g, i = (d - f92 * f93) / 2, d / 2, (d + f92 * f94) / 2
        else:
            e, g, i = d - f92 * f93, d, d + f92 * f94
        rows.append(SstRow(
            body_extn=ext, label=extn_label(ext),
            l1=e, l1_diag=math.sqrt(2 * e * e),
            l2=g, l2_diag=math.sqrt(2 * g * g),
            l3=i, l3_diag=math.sqrt(2 * i * i),
        ))

    return SstResult(
        sheet=sheet,
        slope_face_length=af32, slope_develop_length=ag31, slope_horizontal_2=ah34,
        gauge_total=n88, gauge_half=n89,
        conc_to_bottom_of_stub=i87, bottom_vertical_diff=i88,
        tan=i90, angle_deg=math.degrees(math.atan(i90)), two_tan=i91,
        face=i92, face_x2200=2200 * i92, develop=i93, develop_x2200=2200 * i93,
        second_slope_tan=i95, second_slope_deg=math.degrees(math.atan(i95)),
        second_slope_cos=i97, l1_height=f93, l3_height=f94, rows=rows,
    )


if __name__ == "__main__":
    import argparse
    import json

    ap = argparse.ArgumentParser(description="Print the body extension table")
    ap.add_argument("inputs", nargs="?", help="inputs .json (as saved from the web page)")
    ap.add_argument("--sheet", choices=list(SHEETS))
    a = ap.parse_args()
    data = json.load(open(a.inputs, encoding="utf-8")) if a.inputs else {}
    sheet = a.sheet or data.get("sheet", "SST-BODY")
    sheet = "SST-LEG" if "LEG" in sheet else "SST-BODY"
    res = calculate(SstInputs.from_dict(data.get("inputs", data), sheet), sheet)
    print(SHEETS[sheet])
    print(f"{'BODY EXTN':<18}{'L1':>12}{'L1 DIAG':>12}{'L2':>12}{'L2 DIAG':>12}{'L3':>12}{'L3 DIAG':>12}")
    for r in res.rows:
        print(f"{r.label:<18}{r.l1:>12.2f}{r.l1_diag:>12.2f}{r.l2:>12.2f}{r.l2_diag:>12.2f}{r.l3:>12.2f}{r.l3_diag:>12.2f}")
