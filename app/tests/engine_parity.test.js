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

test("rejects bad input", () => {
  assert.throws(() => calculate({ slope_vertical: 0 }));
  assert.throws(() => calculate({}, "SST-ARM"));
});

test("LEG defaults use the LEG sheet F91", () => {
  assert.strictEqual(calculate({}, "SST-LEG").rows[2].l2, 455055 / 2);
});
