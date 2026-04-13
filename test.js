const assert = require("node:assert/strict");
const { calc, validateCombination } = require("./app.js");

function testD113TenByTenExample() {
  const room = {
    width: 1000,
    length: 1000,
    area: 100,
    systemType: "D113",
    loadClass: "0.50",
    fireProtection: false,
    boardType: "12.5_or_2x12.5",
    a: 700,
    b: 500,
    c: 800,
    offset: 30,
    udAnchorSpacing: 625,
    overrides: { area: false, a: true, b: true, c: true, offset: true, udAnchorSpacing: true },
  };

  const result = calc(room);

  assert.equal(validateCombination(room), true);
  assert.equal(result.bearingCount, 14);
  assert.equal(result.mountingCount, 21);
  assert.equal(result.crossConnectors, 294);
  assert.equal(result.hangersPerBearing, 15);
  assert.equal(result.hangersTotal, 210);
  assert.equal(result.bearingLengthTotal, 140);
  assert.equal(result.mountingLengthTotal, 210);
}

testD113TenByTenExample();
console.log("All tests passed");
