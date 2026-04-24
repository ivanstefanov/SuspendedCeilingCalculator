import { describe, expect, it } from "vitest";
import { buildCutOptimizationInput, buildLayoutHangerPositions, buildLinearPositions, buildMaterialTakeoff, buildPositions, buildStaggeredExtensionLayout, buildSuspendedCeilingLayout, calc, DEFAULT_CONSTANTS, estimateLoadKgPerM2, getAllowedAValues, getAllowedBValues, getAutoABC, getAutomaticLoadClass, getBoardOptions, getEffectiveBoardLayers, getFireCertificationCheck, getHangerOptions, getTableValue, getUdAnchoringRule, getValidationWarnings, optimizeSuspendedCeilingCuts, Room, validateCombination } from "./calculator";

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
    udAnchorSpacing: 625,
    overrides: { area: false, a: true, b: true, c: true, udAnchorSpacing: true },
    ...patch,
  };
}

describe("calculator", () => {
  it("counts D113 rows inside the edge offsets without adding an extra redistributed row", () => {
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
    const positions = buildPositions(result.W, room.c, DEFAULT_CONSTANTS.profileEdgeOffsetCm).map(Math.round);

    expect(result.bearingCount).toBe(9);
    expect(positions).toEqual([10, 60, 110, 160, 210, 260, 310, 360, 410]);
  });

  it("places D113 profiles in a 200 x 200 cm room without exceeding selected spacing", () => {
    const room = makeRoom({
      width: 200,
      length: 200,
      area: 4,
      systemType: "D113",
      loadClass: "0.50",
      a: 800,
      b: 500,
      c: 500,
    });
    const result = calc(room, DEFAULT_CONSTANTS);

    expect(result.bearingCount).toBe(5);
    expect(result.mountingCount).toBe(5);
    expect(result.hangersPerBearing).toBe(4);
    expect(result.bearingProfiles).toBe(5);
    expect(result.mountingProfiles).toBe(5);
    expect(result.cdTotalProfiles).toBe(10);
    expect(result.udProfiles).toBe(4);
    expect(result.anchorsUd).toBe(16);
    expect(buildPositions(result.W, room.c, DEFAULT_CONSTANTS.profileEdgeOffsetCm)).toEqual([10, 55, 100, 145, 190]);
    expect(buildPositions(result.L, room.a, DEFAULT_CONSTANTS.profileEdgeOffsetCm)).toEqual([10, 70, 130, 190]);
  });

  it("redistributes short-room hanger positions while keeping Knauf maximum spacing", () => {
    const positions = buildPositions(226, 800, DEFAULT_CONSTANTS.profileEdgeOffsetCm);
    const gaps = positions.slice(1).map((position, index) => position - positions[index]);

    expect(positions.map(Math.round)).toEqual([10, 79, 147, 216]);
    expect(Math.max(...gaps)).toBeLessThanOrEqual(80);
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

  it("does not generate joints for a 200 cm row with 400 cm stock length", () => {
    expect(buildStaggeredExtensionLayout(1, 200, 4, [10, 90, 170])).toEqual([{ lineIndex: 0, pointsCm: [] }]);
  });

  it("does not generate joints for a 400 cm row with 400 cm stock length", () => {
    expect(buildStaggeredExtensionLayout(1, 400, 4, [10, 90, 170, 250, 330])).toEqual([{ lineIndex: 0, pointsCm: [] }]);
  });

  it("creates one balanced joint when the row is just over one stock length", () => {
    const layout = buildStaggeredExtensionLayout(1, 401, 4, [10, 90, 170, 250, 330]);
    const joint = layout[0]?.pointsCm[0] ?? 0;

    expect(layout).toEqual([{ lineIndex: 0, pointsCm: [170] }]);
    expect(joint).toBeGreaterThan(30);
    expect(401 - joint).toBeGreaterThan(30);
  });

  it("uses the requested four-row stagger pattern for carrier profile joints", () => {
    const hangers = [10, 90, 170, 250, 330, 410, 490, 570, 650, 730, 810, 890];
    const layout = buildStaggeredExtensionLayout(4, 900, 4, hangers);

    expect(layout.map((line) => line.pointsCm)).toEqual([
      [400, 800],
      [240, 640],
      [320, 720],
      [160, 560],
    ]);
  });

  it("keeps the regular hanger grid when joints can snap to it", () => {
    const layout = buildSuspendedCeilingLayout({
      roomLengthCm: 900,
      roomWidthCm: 200,
      profileLengthCm: 400,
      carrierRowSpacingCm: 50,
      hangerSpacingCm: 80,
      firstHangerOffsetCm: 10,
      hangerNearJointMaxDistanceCm: 15,
    });

    expect(layout.hangerPositionsCm).toEqual([10, 90, 170, 250, 330, 410, 490, 570, 650, 730, 810, 890]);
    expect(layout.carrierExtensions[0]?.pointsCm).toEqual([400, 800]);
    expect(layout.carrierExtensions[1]?.pointsCm).toEqual([240, 640]);
    expect(layout.carrierExtensions.flatMap((line) => line.pointsCm).every((joint) => (
      layout.hangerPositionsCm.some((hanger) => Math.abs(hanger - joint) <= 15)
    ))).toBe(true);
    expect(layout.hangerPositionsCm).toHaveLength(12);
  });

  it("does not generate carrier segments longer than the stock profile length", () => {
    const layout = buildSuspendedCeilingLayout({
      roomLengthCm: 900,
      roomWidthCm: 200,
      profileLengthCm: 400,
      carrierRowSpacingCm: 50,
      hangerSpacingCm: 80,
      firstHangerOffsetCm: 10,
      hangerNearJointMaxDistanceCm: 15,
    });

    const segmentLengths = layout.carrierExtensions.flatMap((line) => {
      const boundaries = [0, ...line.pointsCm, 900];
      return boundaries.slice(1).map((point, index) => point - boundaries[index]);
    });

    expect(segmentLengths.every((lengthCm) => lengthCm <= 400)).toBe(true);
  });

  it("adds extra hangers only as a fallback when the regular grid cannot support the joints", () => {
    const regularGrid = buildLinearPositions(900, 120, 10);
    const hangers = buildLayoutHangerPositions(900, 120, 10, [{ lineIndex: 0, pointsCm: [400, 800] }], 15);

    expect(hangers.length).toBeGreaterThan(regularGrid.length);
    expect([400, 800].every((joint) => hangers.some((hanger) => Math.abs(hanger - joint) <= 15))).toBe(true);
  });

  it("keeps multiple splices ordered on each row", () => {
    const layout = buildStaggeredExtensionLayout(3, 1400, 4);

    expect(layout.every((line) => line.pointsCm.length === 3)).toBe(true);
    expect(layout.every((line) => line.pointsCm.every((point, index, points) => index === 0 || point > points[index - 1]))).toBe(true);
  });

  it("optimizes cuts with isolated full-length bars and packed shorter pieces", () => {
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: [
        { rowIndex: 0, segments: [{ fromCm: 0, toCm: 390, lengthCm: 390 }, { fromCm: 390, toCm: 790, lengthCm: 400 }, { fromCm: 790, toCm: 900, lengthCm: 110 }] },
        { rowIndex: 1, segments: [{ fromCm: 0, toCm: 250, lengthCm: 250 }, { fromCm: 250, toCm: 650, lengthCm: 400 }, { fromCm: 650, toCm: 900, lengthCm: 250 }] },
      ],
      mountingRows: [],
      udProfiles: { segments: [] },
    }, {
      stockLengthCm: 400,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
    });

    expect(plan.bars.filter((bar) => bar.pieces.length === 1 && bar.pieces[0]?.lengthCm === 400)).toHaveLength(2);
    expect(plan.bars.every((bar) => bar.usedCm <= bar.stockLengthCm)).toBe(true);
    expect(plan.totalBars).toBeGreaterThan(0);
  });

  it("packs carrier, mounting and ud pieces together in a shared pool", () => {
    const twoHundredSegment = { fromCm: 0, toCm: 200, lengthCm: 200 };
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      mountingRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      udProfiles: { segments: Array.from({ length: 4 }, () => twoHundredSegment) },
    }, {
      stockLengthCm: 400,
      kerfCm: 0,
      minReusableOffcutCm: 20,
    });

    expect(plan.totalBars).toBe(6);
    expect(plan.bars.every((bar) => bar.pieces.length === 2)).toBe(true);
    expect(plan.efficiencyPercent).toBe(100);
    expect(plan.carrierStats.totalUsedCm).toBe(800);
    expect(plan.mountingStats.totalUsedCm).toBe(800);
    expect(plan.udStats.totalUsedCm).toBe(800);
  });

  it("fits 200 + 200 into a 400 cm stock bar when kerf is excluded from fit check", () => {
    const twoHundredSegment = { fromCm: 0, toCm: 200, lengthCm: 200 };
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      mountingRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      udProfiles: { segments: Array.from({ length: 4 }, () => twoHundredSegment) },
    }, {
      stockLengthCm: 400,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
      includeKerfInFitCheck: false,
    });

    expect(plan.totalBars).toBe(6);
    expect(plan.bars.every((bar) => bar.usedCm === 400)).toBe(true);
    expect(plan.totalUsedCm).toBe(2400);
    expect(plan.totalWasteCm).toBe(0);
    expect(plan.efficiencyPercent).toBe(100);
  });

  it("uses the provided 300 cm stock length and requires more bars", () => {
    const twoHundredSegment = { fromCm: 0, toCm: 200, lengthCm: 200 };
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      mountingRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      udProfiles: { segments: Array.from({ length: 4 }, () => twoHundredSegment) },
    }, {
      stockLengthCm: 300,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
      includeKerfInFitCheck: false,
    });

    expect(plan.totalBars).toBe(12);
    expect(plan.totalUsedCm).toBe(2400);
    expect(plan.totalWasteCm).toBe(1200);
    expect(plan.efficiencyPercent).toBe(66.67);
    expect(plan.bars.every((bar) => bar.stockLengthCm === 300)).toBe(true);
  });

  it("uses the provided 600 cm stock length and requires fewer bars", () => {
    const twoHundredSegment = { fromCm: 0, toCm: 200, lengthCm: 200 };
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      mountingRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [twoHundredSegment] })),
      udProfiles: { segments: Array.from({ length: 4 }, () => twoHundredSegment) },
    }, {
      stockLengthCm: 600,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
      includeKerfInFitCheck: false,
    });

    expect(plan.totalBars).toBe(4);
    expect(plan.totalUsedCm).toBe(2400);
    expect(plan.totalWasteCm).toBe(0);
    expect(plan.efficiencyPercent).toBe(100);
    expect(plan.bars.every((bar) => bar.stockLengthCm === 600)).toBe(true);
  });

  it("prioritizes perfect pairs before general packing for mixed piece sizes", () => {
    const piece226 = { fromCm: 0, toCm: 226, lengthCm: 226 };
    const piece200 = { fromCm: 0, toCm: 200, lengthCm: 200 };
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: Array.from({ length: 4 }, (_, rowIndex) => ({ rowIndex, segments: [piece226] })),
      mountingRows: Array.from({ length: 3 }, (_, rowIndex) => ({ rowIndex, segments: [piece200] })),
      udProfiles: { segments: Array.from({ length: 3 }, () => piece200) },
    }, {
      stockLengthCm: 400,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
      includeKerfInFitCheck: false,
    });

    const fullTwoHundredBars = plan.bars.filter((bar) => (
      bar.pieces.length === 2
      && bar.pieces.every((piece) => piece.lengthCm === 200)
      && bar.usedCm === 400
    ));

    expect(fullTwoHundredBars).toHaveLength(3);
    expect(plan.bars.filter((bar) => bar.pieces.length === 1 && bar.pieces[0]?.lengthCm === 226)).toHaveLength(4);
  });

  it("fills existing gaps with smaller pieces after the initial FFD pass", () => {
    const plan = optimizeSuspendedCeilingCuts({
      carrierRows: [
        { rowIndex: 0, segments: [{ fromCm: 0, toCm: 280, lengthCm: 280 }] },
        { rowIndex: 1, segments: [{ fromCm: 0, toCm: 270, lengthCm: 270 }] },
      ],
      mountingRows: [
        { rowIndex: 0, segments: [{ fromCm: 0, toCm: 70, lengthCm: 70 }] },
        { rowIndex: 1, segments: [{ fromCm: 0, toCm: 60, lengthCm: 60 }] },
        { rowIndex: 2, segments: [{ fromCm: 0, toCm: 60, lengthCm: 60 }] },
        { rowIndex: 3, segments: [{ fromCm: 0, toCm: 60, lengthCm: 60 }] },
      ],
      udProfiles: { segments: [] },
    }, {
      stockLengthCm: 400,
      kerfCm: 0.3,
      minReusableOffcutCm: 20,
      includeKerfInFitCheck: false,
    });

    expect(plan.totalBars).toBe(2);
    expect(plan.totalWasteCm).toBe(0);
    expect(plan.efficiencyPercent).toBe(100);
    expect(plan.bars.every((bar) => bar.usedCm === 400)).toBe(true);
  });

  it("builds a true global cut plan for a 2 x 2 m room", () => {
    const room = makeRoom({ width: 200, length: 200, a: 900, b: 600, c: 600 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);
    const plan = optimizeSuspendedCeilingCuts(input, {
      stockLengthCm: 400,
      kerfCm: 0,
      minReusableOffcutCm: 20,
    });

    const allPieces = [
      ...input.carrierRows.flatMap((row) => row.segments),
      ...input.mountingRows.flatMap((row) => row.segments),
      ...input.udProfiles.segments,
    ];

    expect(allPieces).toHaveLength(12);
    expect(allPieces.every((segment) => segment.lengthCm === 200)).toBe(true);
    expect(plan.totalBars).toBe(6);
    expect(plan.totalUsedCm).toBe(2400);
    expect(plan.totalWasteCm).toBe(0);
    expect(plan.efficiencyPercent).toBe(100);
    expect(plan.bars.every((bar) => bar.pieces.length === 2)).toBe(true);
  });

  it("builds cut optimization input from the generated layout without changing segment lengths", () => {
    const room = makeRoom({ width: 420, length: 900, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);
    const expectedMountingRows = buildLinearPositions(result.L, room.b / 10, DEFAULT_CONSTANTS.profileEdgeOffsetCm);

    expect(input.carrierRows).toHaveLength(result.bearingCount);
    expect(input.mountingRows).toHaveLength(expectedMountingRows.length);
    expect(input.carrierRows.every((row) => row.segments.reduce((sum, segment) => sum + segment.lengthCm, 0) === result.L)).toBe(true);
    expect(input.mountingRows.every((row) => row.segments.reduce((sum, segment) => sum + segment.lengthCm, 0) === result.W)).toBe(true);
    expect(input.mountingRows.every((row) => row.segments.every((segment) => segment.lengthCm <= DEFAULT_CONSTANTS.cdLength * 100))).toBe(true);
    expect(input.udProfiles.segments.every((segment) => segment.lengthCm <= DEFAULT_CONSTANTS.udLength * 100)).toBe(true);
  });

  it("splits a 420 cm mounting row into balanced cut segments", () => {
    const room = makeRoom({ width: 420, length: 900, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);

    expect(input.mountingRows[0]?.segments.map((segment) => segment.lengthCm)).toEqual([210, 210]);
  });

  it("keeps a short mounting row as one cut segment", () => {
    const room = makeRoom({ width: 200, length: 900, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);

    expect(input.mountingRows[0]?.segments.map((segment) => segment.lengthCm)).toEqual([200]);
  });

  it("splits long mounting rows so that no cut segment exceeds stock length", () => {
    const room = makeRoom({ width: 900, length: 900, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);

    expect(input.mountingRows[0]?.segments.map((segment) => segment.lengthCm)).toEqual([300, 300, 300]);
    expect(input.mountingRows.every((row) => row.segments.every((segment) => segment.lengthCm <= DEFAULT_CONSTANTS.cdLength * 100))).toBe(true);
  });

  it("builds cut optimization input without any piece longer than stock length", () => {
    const room = makeRoom({ width: 420, length: 900, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);

    const allLengths = [
      ...input.carrierRows.flatMap((row) => row.segments.map((segment) => segment.lengthCm)),
      ...input.mountingRows.flatMap((row) => row.segments.map((segment) => segment.lengthCm)),
      ...input.udProfiles.segments.map((segment) => segment.lengthCm),
    ];

    expect(allLengths.every((lengthCm) => lengthCm <= DEFAULT_CONSTANTS.cdLength * 100)).toBe(true);
    expect(() => optimizeSuspendedCeilingCuts(input, { stockLengthCm: DEFAULT_CONSTANTS.cdLength * 100 })).not.toThrow();
  });

  it("builds a single carrier segment for short rooms without joints or fallback hangers", () => {
    const room = makeRoom({ width: 200, length: 200, a: 800, b: 500, c: 500 });
    const result = calc(room, DEFAULT_CONSTANTS);
    const input = buildCutOptimizationInput(room, result, DEFAULT_CONSTANTS);
    const layout = buildSuspendedCeilingLayout({
      roomWidthCm: result.W,
      roomLengthCm: result.L,
      profileLengthCm: DEFAULT_CONSTANTS.cdLength * 100,
      carrierRowSpacingCm: room.c / 10,
      hangerSpacingCm: room.a / 10,
      firstHangerOffsetCm: DEFAULT_CONSTANTS.profileEdgeOffsetCm,
    });

    expect(layout.carrierExtensions.every((line) => line.pointsCm.length === 0)).toBe(true);
    expect(input.carrierRows.every((row) => row.segments.length === 1 && row.segments[0]?.fromCm === 0 && row.segments[0]?.toCm === result.L)).toBe(true);
    expect(layout.hangerPositionsCm).toEqual(buildLinearPositions(result.L, room.a / 10, DEFAULT_CONSTANTS.profileEdgeOffsetCm));
  });

  it("provides conservative custom defaults for manually tuned suspended ceilings", () => {
    expect(getAutoABC("0.30", false, "knauf_a_12.5", "CUSTOM")).toEqual({
      a: 900,
      b: 500,
      c: 1000,
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

  it("does not require an EI configuration when fire protection is disabled", () => {
    const check = getFireCertificationCheck(makeRoom({ fireRating: "none", fireProtection: false }));

    expect(check.status).toBe("none");
    expect(check.issues).toEqual([]);
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
    expect(warnings.some((warning) => warning.code === "hanger-load-class-040")).toBe(false);
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

  it("omits the system prefix from material labels when all rooms use one system", () => {
    const rows = buildMaterialTakeoff([makeRoom({
      systemType: "D113",
      loadClass: "0.30",
      c: 600,
      a: 900,
    })], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.find((row) => row.key === "d113-cd-60-27")?.label).toBe("CD 60/27 профили");
    expect(rows.find((row) => row.key === "d113-ud-28-27")?.label).toBe("UD 28/27 профили");
  });

  it("keeps the system prefix from material labels when rooms use mixed systems", () => {
    const rows = buildMaterialTakeoff([
      makeRoom({ systemType: "D113", loadClass: "0.30", c: 600, a: 900 }),
      makeRoom({ systemType: "D116", loadClass: "0.30", c: 600, a: 1050 }),
    ], { ...DEFAULT_CONSTANTS, wastePercent: 0 });

    expect(rows.find((row) => row.key === "d113-cd-60-27")?.label).toBe("D113 CD 60/27 профили");
    expect(rows.find((row) => row.key === "d116-ua-50-40")?.label).toBe("D116 UA 50/40 носещи профили");
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
