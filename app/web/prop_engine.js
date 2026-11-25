/*
 * PROP TYPE SST calculation engine (browser + Node), sheets SST-BODY and SST-LEG.
 * Same formulas as engine/prop_sst.py — both are tested against
 * tests/expected_sst.json (values calculated by Excel).
 *
 *   BODY: L2 = F91 + 2tan*ext       L1 = L2 - 2tan*H1       L3 = L2 + 2tan*H3
 *   LEG : D  = F91 + 2tan*ext       L1 = (D - 2tan*H1)/2    L2 = D/2    L3 = (D + 2tan*H3)/2
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PropEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const BODY_EXTENSIONS = [-3000, -1500, 0, 1500, 3000, 4500, 6000, 7500, 9000];
  const ANGLE_FLANGE_FOR_GAUGE = 150; // Excel N88 "=150-T35" (hard-coded in the sheet)

  const SHEETS = { "SST-BODY": "PROP TYPE SST-BODY", "SST-LEG": "PROP TYPE SST-LEG" };
  const DEFAULT_INPUTS = {
    top_of_stub_height: 774.29,   // C18
    conc_to_gl: 225,              // C25
    stub_below_gl: 3300,          // C40
    leg_width: 250,               // B32
    leg_size_text_2: "X250",      // B31
    leg_size_text_3: "X18(H)",    // B30
    hole_text: "17.5Ø HOLE",      // Y30
    slotted_hole_text: "17.5Ø HOLE\nSLOTTED HOLE", // N33
    plate_dim_1: 70,              // W31
    plate_dim_2: 50,              // W32
    hole_gap: 30,                 // T35
    angle_size_text: "L150X150X12", // S37
    slope_vertical: 40595,        // AD32
    slope_horizontal: 10940,      // AE37
    base_width_conc_lvl: 25365.76 // F91
  };

  // Values saved in the Excel for each sheet (only F91 differs)
  function defaultInputs(sheet) {
    const d = Object.assign({}, DEFAULT_INPUTS);
    if (sheet === "SST-LEG") d.base_width_conc_lvl = 455055;
    return d;
  }

  const round = (v, d) => { const f = Math.pow(10, d); return Math.round(v * f) / f; };
  const deg = (r) => (r * 180) / Math.PI;

  function extnLabel(ext) {
    return (ext < 0 ? "-" : "+") + (Math.abs(ext) / 1000).toFixed(1) + "M BODY EXTN";
  }

  function calculate(input, sheet) {
    sheet = sheet || "SST-BODY";
    if (!SHEETS[sheet]) throw new Error("Unknown sheet " + sheet);
    const p = Object.assign(defaultInputs(sheet), input || {});
    const num = (k) => {
      const v = Number(p[k]);
      if (!isFinite(v)) throw new Error(k + " must be a number");
      return v;
    };
    const ad32 = num("slope_vertical"), ae37 = num("slope_horizontal");
    if (ad32 === 0) throw new Error("Slope vertical (AD32) cannot be 0");

    const ah34 = ae37;
    const af32 = round(Math.sqrt(ae37 * ae37 + ad32 * ad32), 3);
    const ag31 = round(Math.sqrt(af32 * af32 + ah34 * ah34), 3);
    const n88 = ANGLE_FLANGE_FOR_GAUGE - num("hole_gap");
    const n89 = n88 / 2;

    const i86 = num("leg_width");
    const i87 = num("stub_below_gl") + num("conc_to_gl");
    const i90 = ae37 / ad32, i91 = i90 * 2;
    const i92 = af32 / ad32, i93 = ag31 / ad32;
    const i95 = ah34 / af32, i97 = ag31 / af32;
    const i88 = (i95 * i86) / i93;

    const f91 = num("base_width_conc_lvl"), f92 = i91;
    const f93 = num("top_of_stub_height"), f94 = i87 - i88;

    const rows = BODY_EXTENSIONS.map((ext) => {
      const d = f91 + f92 * ext;
      let e, g, i;
      if (sheet === "SST-LEG") { e = (d - f92 * f93) / 2; g = d / 2; i = (d + f92 * f94) / 2; }
      else { e = d - f92 * f93; g = d; i = d + f92 * f94; }
      return {
        body_extn: ext, label: extnLabel(ext),
        l1: e, l1_diag: Math.sqrt(2 * e * e),
        l2: g, l2_diag: Math.sqrt(2 * g * g),
        l3: i, l3_diag: Math.sqrt(2 * i * i)
      };
    });

    return {
      sheet, inputs: p,
      slope_face_length: af32, slope_develop_length: ag31, slope_horizontal_2: ah34,
      gauge_total: n88, gauge_half: n89,
      conc_to_bottom_of_stub: i87, bottom_vertical_diff: i88,
      tan: i90, angle_deg: deg(Math.atan(i90)), two_tan: i91,
      face: i92, face_x2200: 2200 * i92, develop: i93, develop_x2200: 2200 * i93,
      second_slope_tan: i95, second_slope_deg: deg(Math.atan(i95)),
      second_slope_cos: i97, l1_height: f93, l3_height: f94, rows
    };
  }

  return { SHEETS, BODY_EXTENSIONS, DEFAULT_INPUTS, defaultInputs, calculate, extnLabel };
});
