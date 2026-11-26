// Checks the browser engine against the same Excel golden values as the Python tests.
//   node --test tests/
const test = require("node:test");
const assert = require("node:assert");
const path = require("path");
const { calculate } = require(path.join(__dirname, "..", "web", "prop_engine.js"));
const golden = require("./expected_sst.json");

const close = (a, b, msg) => assert.ok(Math.abs(a - b) <= 1e-6, `${msg}: got ${a}, expected ${b}`);

for (const c of golden.cases) {
  test(`JS engine matches Excel: ${c.name}`, () => {
    const r = calculate(c.inputs, c.sheet);
    for (const [k, v] of Object.entries(c.results)) close(r[k], v, k);
    c.rows.forEach((exp, i) => {
      assert.strictEqual(r.rows[i].label, exp.label);
      for (const k of ["l1", "l1_diag", "l2", "l2_diag", "l3", "l3_diag"]) close(r.rows[i][k], exp[k], `${exp.label} ${k}`);
    });
  });
}

for (const c of golden.cases) {
  test(`JS title block matches Excel: ${c.name}`, () => {
    const r = calculate(c.inputs, c.sheet);
    for (const [cell, v] of Object.entries(c.labels)) assert.strictEqual(r.labels[cell], v, cell);
  });
}

test("title block: blank follows Excel link, override wins, doc no split", () => {
  const { titleBlock } = require(path.join(__dirname, "..", "web", "prop_engine.js"));
  let tb = titleBlock({ dsgn_name: "ABC", chkd_name: "XYZ", dsgn_date: "01.01.27" });
  assert.strictEqual(tb.AA60, "ABC"); assert.strictEqual(tb.AE60, "XYZ");
  assert.strictEqual(tb.P77, "01.01.27"); assert.strictEqual(tb.P78, "01.01.27");
  tb = titleBlock({ rev_checked: "KMR", appd_date: "06.01.27", doc_no: "ABC" });
  assert.strictEqual(tb.AC60, "KMR"); assert.strictEqual(tb.P78, "06.01.27");
  assert.deepStrictEqual([tb.N80, tb.O80, tb.P80, tb.Q80], ["A", "B", "C", ""]);
  assert.strictEqual(tb.BALLOON, "51");
});

test("rejects bad input", () => {
  assert.throws(() => calculate({ slope_vertical: 0 }));
  assert.throws(() => calculate({}, "SST-ARM"));
});

test("LEG defaults use the LEG sheet F91", () => {
  assert.strictEqual(calculate({}, "SST-LEG").rows[2].l2, 455055 / 2);
});
