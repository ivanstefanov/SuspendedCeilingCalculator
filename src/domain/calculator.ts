export type SystemType = "D111" | "D112" | "D113" | "D116" | "CUSTOM";
export type LoadClass = "0.15" | "0.30" | "0.40" | "0.50" | "0.65";
export type BoardType =
  | "knauf_a_12.5"
  | "knauf_h2_12.5"
  | "knauf_df_12.5"
  | "knauf_diamant_12.5"
  | "knauf_silentboard_12.5"
  | "knauf_2x12.5"
  | "knauf_15"
  | "knauf_2x15"
  | "knauf_18"
  | "knauf_20"
  | "knauf_25"
  | "knauf_25_18"
  | "generic_12.5"
  | "generic_2x12.5"
  | "generic_15"
  | "generic_18_or_20"
  | "generic_25"
  | "custom";

export type OverrideKey = "area" | "a" | "b" | "c" | "offset" | "udAnchorSpacing";

export interface Room {
  id: string;
  name: string;
  width: number;
  length: number;
  area: number;
  systemType: SystemType;
  loadClass: LoadClass;
  fireProtection: boolean;
  boardType: BoardType;
  a: number;
  b: number;
  c: number;
  offset: number;
  udAnchorSpacing: number;
  overrides: Record<OverrideKey, boolean>;
}

export interface CalculatorConstants {
  cdLength: number;
  udLength: number;
  metalScrewsPerCrossConnector: number;
  metalScrewsPerDirectHanger: number;
  drywallScrewsPerM2: number;
  anchorsPerDirectHanger: number;
  wastePercent: number;
}

export interface ConstructionType {
  label: string;
  materialHint: string;
  loadClasses: LoadClass[];
  fireLoadClasses: LoadClass[];
  defaultLoadClass: LoadClass;
  defaultFireProtection: boolean;
  defaultOffset: number;
  defaultUdAnchorSpacing: number;
  table: Record<"false" | "true", Partial<Record<number, Array<number | null>>>>;
}

export interface CalcResult {
  W: number;
  L: number;
  bearingCount: number;
  mountingCount: number;
  bearingLengthTotal: number;
  mountingLengthTotal: number;
  bearingProfiles: number;
  mountingProfiles: number;
  cdTotalLength: number;
  cdTotalProfiles: number;
  crossConnectors: number;
  hangersPerBearing: number;
  hangersTotal: number;
  udTotalLength: number;
  udProfiles: number;
  anchorsUd: number;
  anchorsHangers: number;
  anchorsTotal: number;
  metalScrews: number;
  drywallScrews: number;
  extensionsTotal: number;
}

export type ValidationSeverity = "error" | "warning";

export interface ValidationWarning {
  code: string;
  severity: ValidationSeverity;
  message: string;
}

export const LOAD_CLASSES: LoadClass[] = ["0.15", "0.30", "0.40", "0.50", "0.65"];
export const FIRE_LOAD_CLASSES: LoadClass[] = ["0.30", "0.40", "0.50", "0.65"];

export const BOARD_TYPE_TO_B: Record<Exclude<BoardType, "custom">, number> = {
  "knauf_a_12.5": 500,
  "knauf_h2_12.5": 500,
  "knauf_df_12.5": 500,
  "knauf_diamant_12.5": 500,
  "knauf_silentboard_12.5": 400,
  "knauf_2x12.5": 500,
  "knauf_15": 550,
  "knauf_2x15": 550,
  "knauf_18": 625,
  "knauf_20": 625,
  "knauf_25": 800,
  "knauf_25_18": 625,
  "generic_12.5": 500,
  "generic_2x12.5": 500,
  "generic_15": 550,
  "generic_18_or_20": 625,
  "generic_25": 800,
};

export const BOARD_OPTIONS: Array<{ value: BoardType; label: string }> = [
  { value: "knauf_a_12.5", label: "Knauf A 12.5 mm - b 500" },
  { value: "knauf_h2_12.5", label: "Knauf H2 12.5 mm - b 500" },
  { value: "knauf_df_12.5", label: "Knauf DF/GKF 12.5 mm - b 500" },
  { value: "knauf_diamant_12.5", label: "Knauf Diamant 12.5 mm - b 500" },
  { value: "knauf_silentboard_12.5", label: "Knauf Silentboard 12.5 mm - b 400" },
  { value: "knauf_2x12.5", label: "Knauf 2x12.5 mm - b 500" },
  { value: "knauf_15", label: "Knauf 15 mm - b 550" },
  { value: "knauf_2x15", label: "Knauf 2x15 mm - b 550" },
  { value: "knauf_18", label: "Knauf 18 mm - b 625" },
  { value: "knauf_20", label: "Knauf 20 mm - b 625" },
  { value: "knauf_25", label: "Knauf 25 mm - b 800" },
  { value: "knauf_25_18", label: "Knauf 25+18 mm - b 625" },
  { value: "generic_12.5", label: "Generic 12.5 mm - b 500" },
  { value: "generic_2x12.5", label: "Generic 2x12.5 mm - b 500" },
  { value: "generic_15", label: "Generic 15 mm - b 550" },
  { value: "generic_18_or_20", label: "Generic 18/20 mm - b 625" },
  { value: "generic_25", label: "Generic 25 mm - b 800" },
  { value: "custom", label: "Custom / ръчно b" },
];

export function getBoardOptions(systemType: SystemType): Array<{ value: BoardType; label: string }> {
  if (systemType === "CUSTOM") return BOARD_OPTIONS;
  return BOARD_OPTIONS.filter((option) => option.value.startsWith("knauf_"));
}

export function getAllowedBValues(systemType: SystemType): number[] {
  return Array.from(new Set(
    getBoardOptions(systemType)
      .filter((option) => option.value !== "custom")
      .map((option) => BOARD_TYPE_TO_B[option.value as Exclude<BoardType, "custom">]),
  )).sort((a, b) => a - b);
}

const LEGACY_BOARD_TYPE_MAP: Record<string, BoardType> = {
  "12.5_silent": "knauf_silentboard_12.5",
  "12.5_or_2x12.5": "knauf_a_12.5",
  "15_or_2x15": "knauf_15",
  "18_or_25_18": "knauf_18",
  "20_or_2x20": "knauf_20",
  "25": "knauf_25",
};

function normalizeBoardType(value: unknown): BoardType {
  if (typeof value === "string" && value in LEGACY_BOARD_TYPE_MAP) return LEGACY_BOARD_TYPE_MAP[value];
  if (typeof value === "string" && value in BOARD_TYPE_TO_B) return value as BoardType;
  if (value === "custom") return "custom";
  return "knauf_a_12.5";
}

export const CONSTRUCTION_TYPES: Record<SystemType, ConstructionType> = {
  D111: {
    label: "D111.bg дървена подконструкция",
    materialHint: "Дървени летви; текущият разчет е геометричен и не е пълна покупна листа.",
    loadClasses: ["0.15", "0.30", "0.50"],
    fireLoadClasses: [],
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultOffset: 30,
    defaultUdAnchorSpacing: 1000,
    table: {
      false: {
        500: [1200, 950, 800],
        600: [1150, 900, 750],
        700: [1050, 850, 700],
        800: [1050, 800, null],
        900: [1000, 800, null],
        1000: [950, null, null],
        1100: [900, null, null],
        1200: [900, null, null],
      },
      true: {},
    },
  },
  D112: {
    label: "D112.bg метална подконструкция",
    materialHint: "CD 60/27 двоен грид; липсват варианти само монтажен профил/W профил.",
    loadClasses: ["0.15", "0.30", "0.50", "0.65"],
    fireLoadClasses: FIRE_LOAD_CLASSES,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultOffset: 30,
    defaultUdAnchorSpacing: 625,
    table: {
      false: {
        500: [1200, 950, 800, 750],
        600: [1150, 900, 750, 700],
        700: [1100, 850, 700, 650],
        800: [1050, 800, 700, null],
        900: [1000, 800, null, null],
        1000: [950, 750, null, null],
        1100: [900, 750, null, null],
        1200: [900, null, null, null],
      },
      true: {
        500: [950, 850, 800, 700],
        600: [900, 800, 700, 700],
        700: [850, 750, 700, 650],
        800: [800, null, null, null],
      },
    },
  },
  D113: {
    label: "D113.bg метална подконструкция на едно ниво",
    materialHint: "CD 60/27 на едно ниво; текущият модел е най-близо до примера на Knauf.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: FIRE_LOAD_CLASSES,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultOffset: 30,
    defaultUdAnchorSpacing: 625,
    table: {
      false: {
        500: [1200, 950, 850, 800, 750],
        600: [1150, 900, 800, 750, 700],
        700: [1100, 850, 750, 700, 650],
        800: [1050, 800, 750, 700, null],
        900: [1000, 800, 700, null, null],
        1000: [950, 750, 700, null, null],
        1100: [900, 750, null, null, null],
        1200: [900, 700, null, null, null],
        1250: [900, 650, null, null, null],
      },
      true: {
        500: [850, 750, 700, 600],
        600: [800, 700, 650, 550],
        700: [750, 650, 600, 550],
        800: [700, 650, 600, null],
        900: [700, 600, 550, null],
        1000: [650, 600, 550, null],
        1100: [650, 600, null, null],
        1200: [600, 550, null, null],
        1250: [600, null, null, null],
      },
    },
  },
  D116: {
    label: "D116.bg UA/CD за големи отвори",
    materialHint: "UA 50/40 + CD 60/27; текущата таблица показва профили общо, не отделя UA/CD/UW.",
    loadClasses: ["0.15", "0.30", "0.50", "0.65"],
    fireLoadClasses: FIRE_LOAD_CLASSES,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultOffset: 30,
    defaultUdAnchorSpacing: 625,
    table: {
      false: {
        500: [2600, 2050, 1600, 1200],
        600: [2450, 1950, 1300, 1000],
        700: [2300, 1850, 1100, 850],
        800: [2200, 1650, 1000, null],
        900: [2150, 1450, null, null],
        1000: [2050, 1300, null, null],
        1100: [2000, 1200, null, null],
        1200: [1950, null, null, null],
        1300: [1900, null, null, null],
        1400: [1850, null, null, null],
        1500: [1750, null, null, null],
      },
      true: {
        500: [1150, 1000, 950, 850],
        600: [1050, 950, 900, 800],
        700: [1000, 900, 850, 750],
        800: [950, 850, 800, null],
        900: [900, 800, null, null],
        1000: [900, null, null, null],
      },
    },
  },
  CUSTOM: {
    label: "Custom - ръчни стойности",
    materialHint: "Ръчна конструкция с начални стойности за стандартен CD грид: a 900 mm, b 500 mm, c 1000 mm.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: [],
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultOffset: 30,
    defaultUdAnchorSpacing: 1000,
    table: {
      false: {
        1000: [900, 900, 900, 900, 900],
      },
      true: {},
    },
  },
};

export const DEFAULT_CONSTANTS: CalculatorConstants = {
  cdLength: 4,
  udLength: 4,
  metalScrewsPerCrossConnector: 4,
  metalScrewsPerDirectHanger: 2,
  drywallScrewsPerM2: 25,
  anchorsPerDirectHanger: 1,
  wastePercent: 10,
};

export function getConstruction(roomOrType: Room | SystemType = "D113"): ConstructionType {
  const type = typeof roomOrType === "string" ? roomOrType : roomOrType.systemType;
  return CONSTRUCTION_TYPES[type] ?? CONSTRUCTION_TYPES.D113;
}

export function getLoadClasses(systemType: SystemType, fireProtection: boolean): LoadClass[] {
  const construction = getConstruction(systemType);
  return fireProtection ? construction.fireLoadClasses : construction.loadClasses;
}

export function getTableValue(room: Room, c = room.c, loadClass = room.loadClass): number | null {
  const classes = getLoadClasses(room.systemType, room.fireProtection);
  const idx = classes.indexOf(loadClass);
  if (idx === -1) return null;
  return getConstruction(room).table[String(room.fireProtection) as "false" | "true"]?.[c]?.[idx] ?? null;
}

export function getValidCValues(room: Room): number[] {
  if (room.systemType === "CUSTOM") return [room.c || 1000];

  const classes = getLoadClasses(room.systemType, room.fireProtection);
  const idx = classes.indexOf(room.loadClass);
  const table = getConstruction(room).table[String(room.fireProtection) as "false" | "true"] ?? {};
  return Object.keys(table)
    .map(Number)
    .filter((c) => table[c]?.[idx] != null)
    .sort((a, b) => a - b);
}

export function getAllowedAValues(room: Room): number[] {
  if (room.systemType === "CUSTOM") return [room.a || 900];

  const maxA = getTableValue(room);
  if (!maxA) return [];

  const values: number[] = [];
  for (let value = 100; value <= maxA; value += 50) {
    values.push(value);
  }
  if (!values.includes(maxA)) values.push(maxA);
  return values.sort((a, b) => a - b);
}

export function createRoom(name = "Стая"): Room {
  const construction = getConstruction("D113");
  const room: Room = {
    id: crypto.randomUUID(),
    name,
    width: 400,
    length: 300,
    area: 12,
    systemType: "D113",
    loadClass: construction.defaultLoadClass,
    fireProtection: construction.defaultFireProtection,
    boardType: "knauf_a_12.5",
    a: 900,
    b: 500,
    c: 600,
    offset: construction.defaultOffset,
    udAnchorSpacing: construction.defaultUdAnchorSpacing,
    overrides: { area: false, a: false, b: false, c: false, offset: false, udAnchorSpacing: false },
  };
  applyAutoABC(room);
  return room;
}

export function normalizeRoom(room: Room): Room {
  const construction = getConstruction(room.systemType);
  room.boardType = normalizeBoardType(room.boardType);
  room.fireProtection = Boolean(room.fireProtection && construction.fireLoadClasses.length);
  const incomingOverrides = room.overrides;
  room.overrides = {
    area: Boolean(incomingOverrides?.area),
    a: Boolean(incomingOverrides?.a),
    b: Boolean(incomingOverrides?.b),
    c: Boolean(incomingOverrides?.c),
    offset: Boolean(incomingOverrides?.offset),
    udAnchorSpacing: Boolean(incomingOverrides?.udAnchorSpacing),
  };
  if (!getLoadClasses(room.systemType, room.fireProtection).includes(room.loadClass)) {
    room.loadClass = construction.defaultLoadClass;
  }
  if (!Number(room.offset)) room.offset = construction.defaultOffset;
  if (!Number(room.udAnchorSpacing)) room.udAnchorSpacing = construction.defaultUdAnchorSpacing;
  return room;
}

export function getAutoABC(loadClass: LoadClass, fireProtection: boolean, boardType: BoardType, systemType: SystemType = "D113") {
  const construction = getConstruction(systemType);
  if (systemType === "CUSTOM") {
    return {
      a: 900,
      c: 1000,
      b: boardType === "custom" ? 500 : BOARD_TYPE_TO_B[boardType] ?? 500,
      offset: construction.defaultOffset,
      udAnchorSpacing: construction.defaultUdAnchorSpacing,
    };
  }

  const table = construction.table[String(fireProtection) as "false" | "true"];
  const classes = getLoadClasses(systemType, fireProtection);
  const idx = Math.max(0, classes.indexOf(loadClass));
  const sortedC = Object.keys(table).map(Number).sort((a, b) => a - b);
  const firstValid = sortedC.find((c) => table[c]?.[idx] != null) ?? 600;
  return {
    a: table[firstValid]?.[idx] ?? 900,
    c: firstValid,
    b: boardType === "custom" ? 500 : BOARD_TYPE_TO_B[boardType] ?? 500,
    offset: construction.defaultOffset,
    udAnchorSpacing: construction.defaultUdAnchorSpacing,
  };
}

export function applyAutoABC(room: Room): void {
  normalizeRoom(room);
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType);
  if (!room.overrides.c) room.c = auto.c;
  if (!room.overrides.a) room.a = auto.a;
  if (!room.overrides.b) room.b = auto.b;
  if (!room.overrides.offset) room.offset = auto.offset;
  if (!room.overrides.udAnchorSpacing) room.udAnchorSpacing = auto.udAnchorSpacing;
}

export function syncSpacingFromKnaufTable(room: Room, { keepC = true } = {}): void {
  if (room.systemType === "CUSTOM") {
    const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType);
    if (!room.overrides.c) room.c = auto.c;
    if (!room.overrides.a) room.a = auto.a;
    if (!room.overrides.b) room.b = auto.b;
    if (!room.overrides.offset) room.offset = auto.offset;
    if (!room.overrides.udAnchorSpacing) room.udAnchorSpacing = auto.udAnchorSpacing;
    return;
  }

  const validCValues = getValidCValues(room);
  if (!keepC || !validCValues.includes(Number(room.c))) {
    room.c = validCValues[0] ?? room.c;
  }

  const a = getTableValue(room);
  if (a) room.a = a;

  if (!room.overrides.b && room.boardType !== "custom") {
    room.b = BOARD_TYPE_TO_B[room.boardType] ?? room.b;
  }
}

export function validateCombination(room: Room): boolean {
  return !getValidationWarnings(room).some((warning) => warning.severity === "error");
}

export function getValidationWarnings(room: Room): ValidationWarning[] {
  normalizeRoom(room);
  if (room.systemType === "CUSTOM") {
    return [room.a, room.b, room.c, room.offset, room.udAnchorSpacing].every((value) => Number.isFinite(value) && value > 0)
      ? []
      : [{
        code: "custom-positive-spacing",
        severity: "error",
        message: "Custom конструкцията изисква положителни стойности за a, b, c, начално отстояние и UD стъпка.",
      }];
  }

  const warnings: ValidationWarning[] = [];
  const aExpected = getTableValue(room);
  if (!aExpected) {
    warnings.push({
      code: "missing-table-value",
      severity: "error",
      message: "Няма допустима стойност по Knauf за тази комбинация от система, натоварване, огнезащита и c.",
    });
  } else if (room.a > aExpected) {
    warnings.push({
      code: "hanger-spacing-too-large",
      severity: "error",
      message: `Разстоянието a е ${room.a} mm, а максимумът по Knauf за тази комбинация е ${aExpected} mm.`,
    });
  }

  if (!Object.values(BOARD_TYPE_TO_B).includes(room.b)) {
    warnings.push({
      code: "unsupported-board-spacing",
      severity: "error",
      message: "Стойността b не съответства на избран тип/дебелина гипсокартон. Използвай Custom, ако искаш ръчна стойност.",
    });
  }

  if (room.systemType === "D113" && !room.fireProtection && room.c === 700 && room.loadClass === "0.65" && room.b === 800) {
    warnings.push({
      code: "d113-b800-footnote",
      severity: "error",
      message: "За D113 без огнезащита при c = 700 и товар 0.65 kN/m2 стойността не важи за b = 800 mm.",
    });
  }

  const highLoadClasses: LoadClass[] = ["0.40", "0.50", "0.65"];
  if (highLoadClasses.includes(room.loadClass)) {
    warnings.push({
      code: "hanger-load-class-040",
      severity: "warning",
      message: "За този клас натоварване документацията отбелязва използване на окачвачи с клас носимоспособност 0.40 kN.",
    });
  }

  if (room.fireProtection && room.udAnchorSpacing > 625) {
    warnings.push({
      code: "fire-ud-anchor-spacing",
      severity: "warning",
      message: "При огнезащита/носеща връзка закрепването на UD профила трябва да е до 625 mm.",
    });
  } else if (!room.fireProtection && room.udAnchorSpacing > 1000) {
    warnings.push({
      code: "nonbearing-ud-anchor-spacing",
      severity: "warning",
      message: "При неносеща връзка документацията дава закрепване на UD профила до около 1000 mm.",
    });
  }

  if (room.systemType === "D116") {
    warnings.push({
      code: "d116-material-split-missing",
      severity: "warning",
      message: "D116 използва UA + CD; текущият материален разчет още не отделя UA, CD и UW по отделни редове.",
    });
  }

  if (room.systemType === "D111") {
    warnings.push({
      code: "d111-wood-material-split-missing",
      severity: "warning",
      message: "D111 е дървена подконструкция; текущият материален разчет е геометричен и не отделя дървени летви по система.",
    });
  }

  return warnings;
}

export function countBySpacing(lengthCm: number, spacingMm: number, edgeCm: number): number {
  if (!lengthCm || !spacingMm) return 0;
  const effectiveLength = Math.max(0, lengthCm - edgeCm);
  return Math.ceil(effectiveLength / (spacingMm / 10)) + 1;
}

export function buildPositions(limitCm: number, spacingMm: number, edgeCm = 30): number[] {
  if (!limitCm || !spacingMm) return [];
  const count = countBySpacing(limitCm, spacingMm, edgeCm);
  if (count === 1) return [Math.min(edgeCm, limitCm / 2)];

  const start = Math.min(edgeCm, limitCm / 2);
  const finalInside = Math.max(start, limitCm - edgeCm);
  const actualSpacing = (finalInside - start) / (count - 1);
  return Array.from({ length: count }, (_, idx) => start + idx * actualSpacing);
}

export function calc(room: Room, constants: CalculatorConstants = DEFAULT_CONSTANTS): CalcResult {
  normalizeRoom(room);
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = Number(room.offset);

  const bearingCount = countBySpacing(W, room.c, offset);
  const bearingLengthTotal = bearingCount * (L / 100);
  const bearingProfiles = Math.ceil(bearingLengthTotal / constants.cdLength);

  const mountingCount = countBySpacing(L, room.b, offset);
  const mountingLengthTotal = mountingCount * (W / 100);
  const mountingProfiles = Math.ceil(mountingLengthTotal / constants.cdLength);

  const cdTotalLength = bearingLengthTotal + mountingLengthTotal;
  const cdTotalProfiles = bearingProfiles + mountingProfiles;
  const crossConnectors = bearingCount * mountingCount;
  const hangersPerBearing = countBySpacing(L, room.a, offset);
  const hangersTotal = bearingCount * hangersPerBearing;
  const udTotalLength = (2 * (X + Y)) / 100;
  const udProfiles = Math.ceil(udTotalLength / constants.udLength);
  const anchorsUd = Math.ceil(udTotalLength / (room.udAnchorSpacing / 1000));
  const anchorsHangers = hangersTotal * constants.anchorsPerDirectHanger;
  const anchorsTotal = anchorsUd + anchorsHangers;
  const metalScrews = Math.ceil(
    (crossConnectors * constants.metalScrewsPerCrossConnector)
    + (hangersTotal * constants.metalScrewsPerDirectHanger),
  );
  const drywallScrews = Math.ceil(Number(room.area) * constants.drywallScrewsPerM2);
  const extBearing = bearingCount * Math.max(0, Math.ceil((L / 100) / constants.cdLength) - 1);
  const extMounting = mountingCount * Math.max(0, Math.ceil((W / 100) / constants.cdLength) - 1);

  return {
    W,
    L,
    bearingCount,
    mountingCount,
    bearingLengthTotal,
    mountingLengthTotal,
    bearingProfiles,
    mountingProfiles,
    cdTotalLength,
    cdTotalProfiles,
    crossConnectors,
    hangersPerBearing,
    hangersTotal,
    udTotalLength,
    udProfiles,
    anchorsUd,
    anchorsHangers,
    anchorsTotal,
    metalScrews,
    drywallScrews,
    extensionsTotal: extBearing + extMounting,
  };
}
