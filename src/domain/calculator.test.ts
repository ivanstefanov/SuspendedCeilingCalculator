import { describe, expect, it } from "vitest";
import { buildPositions, calc, DEFAULT_CONSTANTS, getAllowedAValues, getAllowedBValues, getAutoABC, getBoardOptions, getTableValue, getValidationWarnings, Room, validateCombination } from "./calculator";

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
    boardType: "knauf_a_12.5",
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

  it("provides conservative custom defaults for manually tuned suspended ceilings", () => {
    expect(getAutoABC("0.30", false, "knauf_a_12.5", "CUSTOM")).toEqual({
      a: 900,
      b: 500,
      c: 1000,
      offset: 30,
      udAnchorSpacing: 1000,
    });
  });

  it("validates custom rooms by positive editable spacing values instead of Knauf table lookup", () => {
    const room = makeRoom({
      systemType: "CUSTOM",
      a: 875,
      b: 480,
      c: 950,
      udAnchorSpacing: 900,
    });

    expect(validateCombination(room)).toBe(true);
  });

  it("keeps legacy saved board type values compatible", () => {
    const room = makeRoom({ boardType: "12.5_or_2x12.5" as Room["boardType"] });

    expect(validateCombination(room)).toBe(true);
    expect(room.boardType).toBe("knauf_a_12.5");
  });

  it("shows only Knauf board options for Knauf systems and all options for Custom", () => {
    expect(getBoardOptions("D113").every((option) => option.value.startsWith("knauf_"))).toBe(true);
    expect(getBoardOptions("CUSTOM").some((option) => option.value.startsWith("generic_"))).toBe(true);
    expect(getBoardOptions("CUSTOM").some((option) => option.value === "custom")).toBe(true);
  });

  it("limits Knauf distance dropdown options while Custom keeps editable current values", () => {
    const room = makeRoom({ c: 600, loadClass: "0.50" });

    expect(getAllowedAValues(room).at(-1)).toBe(750);
    expect(getAllowedAValues(room)).not.toContain(800);
    expect(getAllowedBValues("D113")).toEqual([400, 500, 550, 625, 800]);
    expect(getAllowedAValues(makeRoom({ systemType: "CUSTOM", a: 875 }))).toEqual([875]);
  });

  it("reports an error when hanger spacing exceeds the Knauf table value", () => {
    const room = makeRoom({ c: 600, loadClass: "0.50", a: 800 });
    const warnings = getValidationWarnings(room);

    expect(warnings.some((warning) => warning.code === "hanger-spacing-too-large" && warning.severity === "error")).toBe(true);
    expect(validateCombination(room)).toBe(false);
  });

  it("reports the D113 b = 800 footnote restriction", () => {
    const room = makeRoom({ c: 700, loadClass: "0.65", b: 800, a: 650 });
    const warnings = getValidationWarnings(room);

    expect(warnings.some((warning) => warning.code === "d113-b800-footnote" && warning.severity === "error")).toBe(true);
  });

  it("warns when fire protection uses UD anchor spacing above 625 mm", () => {
    const room = makeRoom({ fireProtection: true, loadClass: "0.50", c: 600, a: 650, udAnchorSpacing: 1000 });
    const warnings = getValidationWarnings(room);

    expect(warnings.some((warning) => warning.code === "fire-ud-anchor-spacing" && warning.severity === "warning")).toBe(true);
  });
});
