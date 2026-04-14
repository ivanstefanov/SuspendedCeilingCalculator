import { describe, expect, it } from "vitest";
import { buildPositions, calc, DEFAULT_CONSTANTS, getTableValue, Room, validateCombination } from "./calculator";

function makeRoom(patch: Partial<Room> = {}): Room {
  return {
    id: "test-room",
    name: "Test room",
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
    ...patch,
  };
}

describe("calculator", () => {
  it("keeps the known D113 10 x 10 m example aligned with the Knauf material example", () => {
    const room = makeRoom();
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(validateCombination(room)).toBe(true);
    expect(result.bearingCount).toBe(14);
    expect(result.mountingCount).toBe(21);
    expect(result.crossConnectors).toBe(294);
    expect(result.hangersPerBearing).toBe(15);
    expect(result.hangersTotal).toBe(210);
    expect(result.bearingLengthTotal).toBe(140);
    expect(result.mountingLengthTotal).toBe(210);
  });

  it.each([
    ["0.15", 1150],
    ["0.30", 900],
    ["0.40", 800],
    ["0.50", 750],
    ["0.65", 700],
  ] as const)("reads D113 c = 600 table value for load class %s", (loadClass, expectedA) => {
    const room = makeRoom({ loadClass, c: 600, a: expectedA });

    expect(getTableValue(room)).toBe(expectedA);
    expect(validateCombination(room)).toBe(true);
  });

  it("does not add a redundant final bearing profile for a 900 x 420 cm room with c = 500 mm", () => {
    const room = makeRoom({
      width: 900,
      length: 420,
      loadClass: "0.50",
      a: 800,
      b: 500,
      c: 500,
    });
    const result = calc(room, DEFAULT_CONSTANTS);
    const positions = buildPositions(result.W, room.c, room.offset).map(Math.round);

    expect(result.bearingCount).toBe(9);
    expect(positions).toEqual([30, 75, 120, 165, 210, 255, 300, 345, 390]);
  });

  it("does not return negative extensions for rooms shorter than a profile length", () => {
    const room = makeRoom({
      width: 300,
      length: 250,
      area: 7.5,
      a: 800,
      b: 500,
      c: 500,
    });
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(result.extensionsTotal).toBe(0);
  });
});
