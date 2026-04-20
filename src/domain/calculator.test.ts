import { describe, expect, it } from "vitest";
import { buildMaterialTakeoff, buildPositions, calc, DEFAULT_CONSTANTS, estimateLoadKgPerM2, getAllowedAValues, getAllowedBValues, getAutoABC, getAutomaticLoadClass, getBoardOptions, getEffectiveBoardLayers, getFireCertificationCheck, getHangerOptions, getTableValue, getUdAnchoringRule, getValidationWarnings, Room, validateCombination } from "./calculator";

function makeRoom(patch: Partial<Room> = {}): Room {
  return {
    id: "test-room",
    name: "Test room",
    width: 1000,
    length: 1000,
    area: 100,
    systemType: "D113",
    customBearingProfile: "cd_60_27",
    customMountingProfile: "cd_60_27",
    customPerimeterProfile: "ud_28_27",
    loadInputMode: "manual",
    additionalLoadKgPerM2: 0,
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
  it("counts D113 rows inside the edge offsets without adding an extra redistributed row", () => {
    const room = makeRoom();
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(validateCombination(room)).toBe(true);
    expect(result.bearingCount).toBe(13);
    expect(result.mountingCount).toBe(20);
    expect(result.crossConnectors).toBe(260);
    expect(result.hangersPerBearing).toBe(15);
    expect(result.hangersTotal).toBe(195);
    expect(result.bearingLengthTotal).toBe(130);
    expect(result.mountingLengthTotal).toBe(200);
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
    expect(positions).toEqual([30, 80, 130, 180, 230, 280, 330, 380, 390]);
  });

  it("places D113 profiles in a 200 x 200 cm room at the selected spacing instead of redistributing to 35 cm", () => {
    const room = makeRoom({
      width: 200,
      length: 200,
      area: 4,
      systemType: "D113",
      loadClass: "0.50",
      a: 800,
      b: 500,
      c: 500,
      offset: 30,
    });
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(result.bearingCount).toBe(4);
    expect(result.mountingCount).toBe(4);
    expect(result.hangersPerBearing).toBe(3);
    expect(result.bearingProfiles).toBe(4);
    expect(result.mountingProfiles).toBe(4);
    expect(result.cdTotalProfiles).toBe(8);
    expect(result.udProfiles).toBe(4);
    expect(result.anchorsUd).toBe(16);
    expect(buildPositions(result.W, room.c, room.offset)).toEqual([30, 80, 130, 170]);
    expect(buildPositions(result.L, room.a, room.offset)).toEqual([30, 110, 170]);
  });

  it("counts profile pieces per line and UD pieces per side for practical purchasing", () => {
    const room = makeRoom({
      width: 420,
      length: 900,
      area: 37.8,
      a: 800,
      b: 500,
      c: 500,
      udAnchorSpacing: 625,
    });
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(result.bearingProfiles).toBe(result.bearingCount * 3);
    expect(result.mountingProfiles).toBe(result.mountingCount * 2);
    expect(result.udProfiles).toBe(10);
    expect(result.anchorsUd).toBe(44);
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

  it("does not apply Knauf footnote limits to custom rooms", () => {
    const warnings = getValidationWarnings(makeRoom({
      systemType: "CUSTOM",
      loadClass: "0.65",
      hangerType: "generic",
      a: 2000,
      b: 800,
      c: 1000,
      udAnchorSpacing: 1200,
    }));

    expect(warnings).toEqual([]);
  });

  it("allows every hanger option for custom rooms", () => {
    const options = getHangerOptions("CUSTOM").map((option) => option.value);

    expect(options).toEqual(expect.arrayContaining(["generic", "direct", "nonius", "anchorfix", "acoustic", "ua_m8"]));
  });

  it("uses selected custom profiles for material takeoff", () => {
    const rows = buildMaterialTakeoff([makeRoom({
      systemType: "CUSTOM",
      customBearingProfile: "ua_50_40",
      customMountingProfile: "wood_batten",
      customPerimeterProfile: "uw_50_40",
      hangerType: "nonius",
      a: 875,
      b: 480,
      c: 950,
      udAnchorSpacing: 900,
    })], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.some((row) => row.key === "custom-bearing-ua_50_40" && row.label.includes("UA 50/40"))).toBe(true);
    expect(rows.some((row) => row.key === "custom-mounting-wood_batten" && row.label.includes("Дървена летва"))).toBe(true);
    expect(rows.some((row) => row.key === "custom-perimeter-uw_50_40" && row.label.includes("UW 50/40"))).toBe(true);
    expect(rows.some((row) => row.key === "custom-hangers-nonius")).toBe(true);
  });

  it("estimates an automatic load class from boards, fire protection and extra load", () => {
    const room = makeRoom({
      boardType: "knauf_25_18",
      fireRating: "ei90_bottom",
      fireProtection: true,
      additionalLoadKgPerM2: 8,
    });

    expect(estimateLoadKgPerM2(room)).toBe(46);
    expect(getAutomaticLoadClass(room)).toBe("0.50");
  });

  it("marks fire certification as incomplete when a concrete EI system is not selected", () => {
    const check = getFireCertificationCheck(makeRoom({ fireRating: "fire_table", fireProtection: true }));

    expect(check.status).toBe("incomplete");
    expect(check.issues.some((issue) => issue.code === "specific-fire-rating-missing")).toBe(true);
  });

  it("marks fire certification as invalid and reports allowed values", () => {
    const check = getFireCertificationCheck(makeRoom({
      fireRating: "ei90_bottom",
      fireProtection: true,
      boardType: "knauf_a_12.5",
      hangerType: "direct",
      b: 625,
      c: 600,
      a: 650,
      udAnchorSpacing: 1000,
      loadClass: "0.50",
    }));

    expect(check.status).toBe("invalid");
    expect(check.issues.find((issue) => issue.code === "fire-board-type-invalid")?.allowedValues).toContain("Knauf 2x15 mm");
    expect(check.issues.find((issue) => issue.code === "fire-b-spacing-invalid")?.allowedValues).toContain("500 mm");
    expect(check.issues.find((issue) => issue.code === "fire-hanger-capacity-invalid")?.allowedValues).toContain("Нониус окачвач");
  });

  it("marks a modeled EI combination as complete", () => {
    const check = getFireCertificationCheck(makeRoom({
      systemType: "D113",
      fireRating: "ei60_bottom",
      fireProtection: true,
      boardType: "knauf_2x12.5",
      hangerType: "nonius",
      loadClass: "0.50",
      b: 500,
      c: 600,
      a: 650,
      udAnchorSpacing: 625,
    }));

    expect(check.status).toBe("complete");
    expect(check.issues).toEqual([]);
  });

  it("warns when high load uses a hanger type without 0.40 kN capacity", () => {
    const warnings = getValidationWarnings(makeRoom({
      systemType: "D113",
      loadClass: "0.50",
      hangerType: "direct",
      c: 600,
      a: 750,
    }));

    expect(warnings.some((warning) => warning.code === "hanger-type-capacity-review")).toBe(true);
  });

  it("does not warn for high load when a 0.40 kN hanger type is selected", () => {
    const warnings = getValidationWarnings(makeRoom({
      systemType: "D113",
      loadClass: "0.50",
      hangerType: "nonius",
      c: 600,
      a: 750,
    }));

    expect(warnings.some((warning) => warning.code === "hanger-type-capacity-review")).toBe(false);
  });

  it("uses the selected hanger type in material takeoff rows", () => {
    const rows = buildMaterialTakeoff([makeRoom({
      systemType: "D113",
      loadClass: "0.30",
      hangerType: "anchorfix",
      c: 600,
      a: 900,
    })], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.some((row) => row.key === "d113-hangers-anchorfix" && row.label.includes("Анкерфикс"))).toBe(true);
  });

  it("adds board, joint and finish material rows to the takeoff", () => {
    const rows = buildMaterialTakeoff([makeRoom({
      systemType: "D113",
      area: 12,
      loadClass: "0.30",
      boardType: "knauf_2x12.5",
      c: 600,
      a: 900,
    })], {
      ...DEFAULT_CONSTANTS,
      wastePercent: 0,
      boardWidth: 1.2,
      boardLength: 2,
      boardLayers: 1,
    });

    expect(rows.find((row) => row.key === "d113-boards")?.quantity).toBe(10);
    expect(rows.find((row) => row.key === "d113-boards-area")?.quantity).toBe(24);
    expect(rows.find((row) => row.key === "d113-boards")?.label).toContain("Knauf 2x12.5 mm");
    expect(rows.some((row) => row.key === "d113-joint-tape")).toBe(true);
    expect(rows.some((row) => row.key === "d113-joint-compound")).toBe(true);
    expect(rows.some((row) => row.key === "d113-trenn-fix")).toBe(true);
  });

  it("scales TN drywall screws by board layers", () => {
    const room = makeRoom({ area: 12, loadClass: "0.30", boardType: "knauf_2x12.5", c: 600, a: 900 });
    const result = calc(room, { ...DEFAULT_CONSTANTS, boardLayers: 1, drywallScrewsPerM2: 25 });

    expect(result.drywallScrews).toBe(600);
  });

  it("uses manual board layers only for custom board types", () => {
    const customRoom = makeRoom({ boardType: "custom", area: 12, b: 480 });
    const knaufRoom = makeRoom({ boardType: "knauf_2x15", area: 12, b: 550 });
    const constants = { ...DEFAULT_CONSTANTS, boardLayers: 3, drywallScrewsPerM2: 25 };

    expect(getEffectiveBoardLayers(customRoom, constants)).toBe(3);
    expect(getEffectiveBoardLayers(knaufRoom, constants)).toBe(2);
    expect(calc(customRoom, constants).drywallScrews).toBe(900);
    expect(calc(knaufRoom, constants).drywallScrews).toBe(600);
  });

  it("adds mineral wool when fire protection is selected", () => {
    const rows = buildMaterialTakeoff([makeRoom({
      systemType: "D113",
      fireRating: "fire_table",
      fireProtection: true,
      loadClass: "0.50",
      c: 600,
      a: 650,
      area: 12,
    })], { ...DEFAULT_CONSTANTS, wastePercent: 0, mineralWoolEnabled: false });

    expect(rows.find((row) => row.key === "d113-mineral-wool")?.quantity).toBe(12);
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

  it("reads the D112 double CD table separately from the direct CD variant", () => {
    const doubleCd = makeRoom({
      systemType: "D112",
      d112Variant: "double_cd",
      loadClass: "0.30",
      fireProtection: false,
      c: 500,
      a: 900,
    });
    const directCd = makeRoom({
      systemType: "D112",
      d112Variant: "direct_cd",
      loadClass: "0.30",
      fireProtection: false,
      c: 500,
      a: 950,
    });

    expect(getTableValue(doubleCd)).toBe(900);
    expect(getTableValue(directCd)).toBe(950);
  });

  it("supports the 0.40 load class for D112 direct CD only", () => {
    const directCd = makeRoom({
      systemType: "D112",
      d112Variant: "direct_cd",
      loadClass: "0.40",
      fireProtection: false,
      c: 500,
      a: 850,
    });
    const doubleCd = makeRoom({
      systemType: "D112",
      d112Variant: "double_cd",
      loadClass: "0.40",
      fireProtection: false,
      c: 500,
      a: 850,
    });

    expect(getTableValue(directCd)).toBe(850);
    expect(getTableValue(doubleCd)).toBeNull();
  });

  it("normalizes D112 direct CD fire protection because that extracted table has no fire rows", () => {
    const room = makeRoom({
      systemType: "D112",
      d112Variant: "direct_cd",
      fireProtection: true,
      loadClass: "0.50",
      c: 500,
      a: 800,
    });

    expect(validateCombination(room)).toBe(true);
    expect(room.fireProtection).toBe(false);
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

    expect(warnings.some((warning) => warning.code === "ud-anchor-spacing" && warning.severity === "warning")).toBe(true);
  });

  it("warns that generic table fire protection still needs a specific EI review", () => {
    const room = makeRoom({ fireRating: "fire_table", fireProtection: true, loadClass: "0.50", c: 600, a: 650 });
    const warnings = getValidationWarnings(room);

    expect(warnings.some((warning) => warning.code === "fire-ei-not-specific")).toBe(true);
  });

  it("normalizes legacy fireProtection to fireRating for saved rooms", () => {
    const room = makeRoom({ fireRating: undefined, fireProtection: true, loadClass: "0.50", c: 600, a: 650 });

    expect(validateCombination(room)).toBe(true);
    expect(room.fireRating).toBe("fire_table");
    expect(room.fireProtection).toBe(true);
  });

  it("warns for wide board spacing combinations that need footnote review", () => {
    const room = makeRoom({ loadClass: "0.65", b: 625, c: 600, a: 700, boardType: "knauf_18" });
    const warnings = getValidationWarnings(room);

    expect(warnings.some((warning) => warning.code === "b500-footnote-review")).toBe(true);
  });

  it("keeps known b = 800 restrictions as errors and other b = 800 cases as review warnings", () => {
    const knownRestriction = getValidationWarnings(makeRoom({ c: 700, loadClass: "0.65", b: 800, a: 650 }));
    const reviewOnly = getValidationWarnings(makeRoom({ c: 600, loadClass: "0.30", b: 800, a: 900, boardType: "knauf_25" }));

    expect(knownRestriction.some((warning) => warning.code === "d113-b800-footnote" && warning.severity === "error")).toBe(true);
    expect(reviewOnly.some((warning) => warning.code === "b800-footnote-review" && warning.severity === "warning")).toBe(true);
  });

  it("uses a stricter UD anchoring rule for metal D11 systems than for D111", () => {
    expect(getUdAnchoringRule(makeRoom({ systemType: "D113", fireProtection: false })).maxSpacing).toBe(625);
    expect(getUdAnchoringRule(makeRoom({ systemType: "D111", fireProtection: false })).maxSpacing).toBe(1000);
    expect(getUdAnchoringRule(makeRoom({ systemType: "D113", fireProtection: true })).mode).toBe("огнезащита");
  });

  it.each([
    ["0.15", 1400],
    ["0.30", 1150],
    ["0.40", 1000],
    ["0.50", 950],
    ["0.65", 850],
  ] as const)("reads D116 UA/CD c = 500 table value for load class %s", (loadClass, expectedA) => {
    const room = makeRoom({ systemType: "D116", loadClass, fireProtection: false, c: 500, a: expectedA });

    expect(getTableValue(room)).toBe(expectedA);
    expect(validateCombination(room)).toBe(true);
  });

  it("reports a D116 b <= 500 footnote from the UA/CD table", () => {
    const warnings = getValidationWarnings(makeRoom({
      systemType: "D116",
      d116Variant: "ua_cd",
      loadClass: "0.30",
      fireProtection: false,
      c: 1000,
      a: 900,
      b: 625,
      boardType: "knauf_18",
    }));

    expect(warnings.some((warning) => warning.code === "b500-footnote-review")).toBe(true);
  });

  it("reads the D116 wide span table separately from UA/CD", () => {
    const room = makeRoom({
      systemType: "D116",
      d116Variant: "wide_span",
      loadClass: "0.30",
      fireProtection: false,
      c: 500,
      a: 1700,
    });

    expect(getTableValue(room)).toBe(2050);
    expect(getAllowedAValues(room).at(-1)).toBe(2050);
  });

  it("reports the D116 wide span a = 1700 mm footnote only for that variant", () => {
    const wideSpan = getValidationWarnings(makeRoom({
      systemType: "D116",
      d116Variant: "wide_span",
      loadClass: "0.30",
      fireProtection: false,
      c: 500,
      a: 2050,
    }));
    const uaCd = getValidationWarnings(makeRoom({
      systemType: "D116",
      d116Variant: "ua_cd",
      loadClass: "0.30",
      fireProtection: false,
      c: 500,
      a: 1150,
    }));

    expect(wideSpan.some((warning) => warning.code === "d116-wide-span-a1700-footnote" && warning.severity === "error")).toBe(true);
    expect(uaCd.some((warning) => warning.code === "d116-wide-span-a1700-footnote")).toBe(false);
  });

  it("splits D116 material takeoff into UA, CD and UW rows", () => {
    const room = makeRoom({
      systemType: "D116",
      loadClass: "0.30",
      fireProtection: false,
      c: 600,
      a: 1050,
    });
    const rows = buildMaterialTakeoff([room], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.some((row) => row.key === "d116-ua-50-40" && row.label.includes("UA 50/40"))).toBe(true);
    expect(rows.some((row) => row.key === "d116-cd-60-27" && row.label.includes("CD 60/27"))).toBe(true);
    expect(rows.some((row) => row.key === "d116-uw-50-40" && row.source === "estimate")).toBe(true);
  });

  it("uses wood batten rows for D111 instead of a generic CD material row", () => {
    const room = makeRoom({
      systemType: "D111",
      loadClass: "0.30",
      fireProtection: false,
      c: 600,
      a: 900,
    });
    const rows = buildMaterialTakeoff([room], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.some((row) => row.key === "d111-bearing-battens")).toBe(true);
    expect(rows.some((row) => row.key === "d111-mounting-battens")).toBe(true);
    expect(rows.some((row) => row.key === "d111-cd-60-27")).toBe(false);
  });
});
