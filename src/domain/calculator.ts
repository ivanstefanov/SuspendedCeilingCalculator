export type SystemType = "D111" | "D112" | "D113" | "D116" | "CUSTOM";
export type LoadClass = "0.15" | "0.30" | "0.40" | "0.50" | "0.65";
export type D112Variant = "double_cd" | "direct_cd";
export type D116Variant = "ua_cd" | "wide_span";
export type FireRating = "none" | "fire_table" | "ei30_bottom" | "ei60_bottom" | "ei90_bottom" | "ei90_top" | "ei120_bottom";
export type HangerType = "generic" | "direct" | "nonius" | "anchorfix" | "acoustic" | "ua_m8";
export type CustomProfileType = "cd_60_27" | "ud_28_27" | "ua_50_40" | "uw_50_40" | "wood_batten" | "custom_profile";
export type LoadInputMode = "manual" | "auto";
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

export type OverrideKey = "area" | "a" | "b" | "c" | "udAnchorSpacing";

export interface Room {
  id: string;
  name: string;
  width: number;
  length: number;
  area: number;
  systemType: SystemType;
  d112Variant?: D112Variant;
  d116Variant?: D116Variant;
  customBearingProfile?: CustomProfileType;
  customMountingProfile?: CustomProfileType;
  customPerimeterProfile?: CustomProfileType;
  loadInputMode?: LoadInputMode;
  additionalLoadKgPerM2?: number;
  loadClass: LoadClass;
  fireRating?: FireRating;
  fireProtection: boolean;
  hangerType?: HangerType;
  boardType: BoardType;
  a: number;
  b: number;
  c: number;
  udAnchorSpacing: number;
  overrides: Record<OverrideKey, boolean>;
}

export interface CalculatorConstants {
  cdLength: number;
  udLength: number;
  uaLength: number;
  woodBattenLength: number;
  metalScrewsPerCrossConnector: number;
  metalScrewsPerDirectHanger: number;
  drywallScrewsPerM2: number;
  boardWidth: number;
  boardLength: number;
  boardLayers: number;
  profileEdgeOffsetCm: number;
  jointTapePerM2: number;
  jointCompoundKgPerM2: number;
  trennFixPerimeterMultiplier: number;
  mineralWoolEnabled: boolean;
  mineralWoolThickness: number;
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

export interface ExtensionLineLayout {
  lineIndex: number;
  pointsCm: number[];
}

export interface SuspendedCeilingLayoutInput {
  roomLengthCm: number;
  roomWidthCm: number;
  profileLengthCm?: number;
  carrierRowSpacingCm: number;
  hangerSpacingCm: number;
  firstHangerOffsetCm?: number;
  minDistanceBetweenAlignedJointsCm?: number;
  hangerNearJointMaxDistanceCm?: number;
}

export interface SuspendedCeilingLayout {
  carrierRowsYcm: number[];
  mountingRowsXcm: number[];
  hangerPositionsCm: number[];
  carrierExtensions: ExtensionLineLayout[];
}

export interface CeilingSegment {
  fromCm: number;
  toCm: number;
  lengthCm: number;
}

export interface CeilingRowSegments {
  rowIndex: number;
  segments: CeilingSegment[];
}

export interface CutOptimizationInput {
  carrierRows: CeilingRowSegments[];
  mountingRows: CeilingRowSegments[];
  udProfiles: {
    segments: CeilingSegment[];
  };
}

export type CutPieceType = "carrier" | "mounting" | "ud";

export interface CutPiece {
  id: string;
  type: CutPieceType;
  lengthCm: number;
  roomId?: string;
  roomName?: string;
}

export interface CutBarSegment {
  pieceId: string;
  type: CutPieceType;
  lengthCm: number;
  startCm: number;
  endCm: number;
}

export type CutBarType = CutPieceType | "cd" | "mixed";

export interface CutBar {
  id: string;
  type: CutBarType;
  stockLengthCm: number;
  pieces: CutPiece[];
  segments: CutBarSegment[];
  usedCm: number;
  wasteCm: number;
  estimatedKerfLossCm: number;
}

export interface CutPlanStats {
  totalBars: number;
  totalUsedCm: number;
  totalWasteCm: number;
  efficiencyPercent: number;
  lowerBoundBars: number;
  estimatedKerfLossCm: number;
}

export interface CutOptimizationResult extends CutPlanStats {
  bars: CutBar[];
  carrierStats: CutPlanStats;
  mountingStats: CutPlanStats;
  udStats: CutPlanStats;
}

export interface CutOptimizationConfig {
  stockLengthCm?: number;
  kerfCm?: number;
  minReusableOffcutCm?: number;
  perTypeStockLengthsCm?: Partial<Record<CutPieceType, number>>;
  includeKerfInFitCheck?: boolean;
  strictProfileSeparation?: boolean;
}

export type ValidationSeverity = "error" | "warning";

export interface ValidationWarning {
  code: string;
  severity: ValidationSeverity;
  message: string;
}

export type FireCertificationStatus = "none" | "complete" | "incomplete" | "invalid";
export type FireCertificationIssueType = "missing" | "invalid";

export interface FireCertificationIssue {
  code: string;
  type: FireCertificationIssueType;
  message: string;
  allowedValues?: string[];
}

export interface FireCertificationCheck {
  status: FireCertificationStatus;
  label: string;
  summary: string;
  issues: FireCertificationIssue[];
}

export type QuantitySource = "knauf-table" | "geometry" | "estimate" | "manual";

export interface MaterialTakeoffItem {
  key: string;
  label: string;
  quantity: number;
  unit: "бр." | "m" | "m2" | "kg";
  note: string;
  source: QuantitySource;
  optimizedQuantity?: number;
  optimizedExplanation?: string;
}

export interface UdAnchoringRule {
  mode: "неносеща връзка" | "носеща връзка" | "огнезащита";
  maxSpacing: number;
  note: string;
}

export interface HangerOption {
  value: HangerType;
  label: string;
  capacityKn: 0.25 | 0.4 | null;
  systems: SystemType[];
  description: string;
  useWhen: string;
}

export interface CustomProfileOption {
  value: CustomProfileType;
  label: string;
  defaultLengthKey: keyof Pick<CalculatorConstants, "cdLength" | "udLength" | "uaLength" | "woodBattenLength">;
}

export const LOAD_CLASSES: LoadClass[] = ["0.15", "0.30", "0.40", "0.50", "0.65"];
export const FIRE_LOAD_CLASSES: LoadClass[] = ["0.30", "0.40", "0.50", "0.65"];
export const FIRE_RATING_OPTIONS: Array<{ value: FireRating; label: string }> = [
  { value: "none", label: "Без огнезащита" },
  { value: "fire_table", label: "Огнезащита по D11 таблица" },
  { value: "ei30_bottom", label: "EI30 отдолу" },
  { value: "ei60_bottom", label: "EI60 отдолу" },
  { value: "ei90_bottom", label: "EI90 отдолу" },
  { value: "ei90_top", label: "EI90 отгоре" },
  { value: "ei120_bottom", label: "EI120 отдолу" },
];

const SPECIFIC_FIRE_RATINGS: FireRating[] = ["ei30_bottom", "ei60_bottom", "ei90_bottom", "ei90_top", "ei120_bottom"];
const FIRE_CERT_BOARD_OPTIONS: Record<Exclude<FireRating, "none" | "fire_table">, BoardType[]> = {
  ei30_bottom: ["knauf_df_12.5", "knauf_2x12.5", "knauf_15", "knauf_2x15", "knauf_18", "knauf_20", "knauf_25", "knauf_25_18"],
  ei60_bottom: ["knauf_2x12.5", "knauf_2x15", "knauf_18", "knauf_20", "knauf_25", "knauf_25_18"],
  ei90_bottom: ["knauf_2x15", "knauf_25", "knauf_25_18"],
  ei90_top: ["knauf_2x15", "knauf_25_18"],
  ei120_bottom: ["knauf_25_18"],
};

const FIRE_CERT_MAX_B: Record<Exclude<FireRating, "none" | "fire_table">, number> = {
  ei30_bottom: 625,
  ei60_bottom: 625,
  ei90_bottom: 500,
  ei90_top: 500,
  ei120_bottom: 500,
};

const FIRE_CERT_ALLOWED_SYSTEMS: Record<Exclude<FireRating, "none" | "fire_table">, SystemType[]> = {
  ei30_bottom: ["D112", "D113", "D116"],
  ei60_bottom: ["D112", "D113", "D116"],
  ei90_bottom: ["D112", "D113", "D116"],
  ei90_top: ["D112", "D113"],
  ei120_bottom: ["D113"],
};

const FIRE_CERT_REQUIRED_HANGER_CAPACITY: Record<Exclude<FireRating, "none" | "fire_table">, 0.25 | 0.4> = {
  ei30_bottom: 0.25,
  ei60_bottom: 0.4,
  ei90_bottom: 0.4,
  ei90_top: 0.4,
  ei120_bottom: 0.4,
};
export const DEFAULT_HANGER_TYPE: HangerType = "generic";
export const DEFAULT_D112_VARIANT: D112Variant = "double_cd";
export const DEFAULT_D116_VARIANT: D116Variant = "ua_cd";
export const DEFAULT_CUSTOM_BEARING_PROFILE: CustomProfileType = "cd_60_27";
export const DEFAULT_CUSTOM_MOUNTING_PROFILE: CustomProfileType = "cd_60_27";
export const DEFAULT_CUSTOM_PERIMETER_PROFILE: CustomProfileType = "ud_28_27";

export const CUSTOM_PROFILE_OPTIONS: Record<CustomProfileType, CustomProfileOption> = {
  cd_60_27: { value: "cd_60_27", label: "CD 60/27", defaultLengthKey: "cdLength" },
  ud_28_27: { value: "ud_28_27", label: "UD 28/27", defaultLengthKey: "udLength" },
  ua_50_40: { value: "ua_50_40", label: "UA 50/40", defaultLengthKey: "uaLength" },
  uw_50_40: { value: "uw_50_40", label: "UW 50/40", defaultLengthKey: "uaLength" },
  wood_batten: { value: "wood_batten", label: "Дървена летва", defaultLengthKey: "woodBattenLength" },
  custom_profile: { value: "custom_profile", label: "Друг профил", defaultLengthKey: "cdLength" },
};

export const HANGER_OPTIONS: Record<HangerType, HangerOption> = {
  generic: {
    value: "generic",
    label: "Общ/неуточнен окачвач",
    capacityKn: null,
    systems: ["D111", "D112", "D113", "D116", "CUSTOM"],
    description: "Не задава конкретен системен окачвач.",
    useWhen: "Използвай при Custom или когато детайлът още не е избран. За Knauf системи приложението ще показва предупреждения при високи товари.",
  },
  direct: {
    value: "direct",
    label: "Директен окачвач за CD",
    capacityKn: 0.25,
    systems: ["D112", "D113"],
    description: "Стандартен директен окачвач за CD 60/27 при малки до средни окачвания.",
    useWhen: "Подходящ за обичайни D112/D113 решения без специални изисквания за по-висока носимоспособност.",
  },
  nonius: {
    value: "nonius",
    label: "Нониус окачвач",
    capacityKn: 0.4,
    systems: ["D112", "D113", "D116"],
    description: "Регулируемо окачване с по-висока носимоспособност според системния комплект.",
    useWhen: "Използвай при по-високи товари, по-големи височини на окачване или когато footnote изисква 0.40 kN.",
  },
  anchorfix: {
    value: "anchorfix",
    label: "Анкерфикс за CD",
    capacityKn: 0.4,
    systems: ["D112", "D113"],
    description: "Системен окачващ елемент за CD профили с носимоспособност 0.40 kN при правилен монтаж.",
    useWhen: "Използвай при CD системи, когато таблицата/бележките изискват окачвач клас 0.40 kN.",
  },
  acoustic: {
    value: "acoustic",
    label: "Акустичен окачвач",
    capacityKn: 0.25,
    systems: ["D112", "D113"],
    description: "Окачвач с еластичен/акустичен елемент за намаляване на вибрации и шум.",
    useWhen: "Използвай при акустични изисквания; носимоспособността трябва да се провери по конкретния продукт.",
  },
  ua_m8: {
    value: "ua_m8",
    label: "UA/M8 окачване",
    capacityKn: 0.4,
    systems: ["D116"],
    description: "Окачване за UA профили при D116, включително M8 решенията в системните детайли.",
    useWhen: "Използвай при D116, особено за варианти с UA профили и по-големи отвори.",
  },
};

export const D112_VARIANTS: Record<D112Variant, {
  label: string;
  materialHint: string;
  loadClasses: LoadClass[];
  fireLoadClasses: LoadClass[];
  table: Record<"false" | "true", Partial<Record<number, Array<number | null>>>>;
}> = {
  double_cd: {
    label: "D112 двоен CD грид",
    materialHint: "Основна D112 таблица от стр. 10: носещи + монтажни CD 60/27.",
    loadClasses: ["0.15", "0.30", "0.50", "0.65"],
    fireLoadClasses: FIRE_LOAD_CLASSES,
    table: {
      false: {
        400: [1200, 950, 800, 750],
        500: [1150, 900, 750, 700],
        600: [1100, 850, 700, 650],
        700: [1050, 800, 700, null],
        900: [1000, 800, null, null],
        1000: [950, 750, null, null],
        1100: [900, 750, null, null],
        1200: [900, null, null, null],
      },
      true: {
        400: [950, 850, 800, 700],
        500: [900, 800, 700, 700],
        600: [850, 750, 700, 650],
        700: [800, null, null, null],
      },
    },
  },
  direct_cd: {
    label: "D112 директен CD вариант",
    materialHint: "D112 таблица от стр. 13: директен CD вариант; W-профилната таблица изисква отделен модел.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: [],
    table: {
      false: {
        500: [1200, 950, 850, 800, 700],
        600: [1100, 900, 800, 700, 700],
        700: [1000, 850, 750, 700, 650],
        800: [1000, 800, null, null, null],
        900: [1000, null, null, null, null],
      },
      true: {},
    },
  },
};

export const D116_VARIANTS: Record<D116Variant, {
  label: string;
  materialHint: string;
  loadClasses: LoadClass[];
  fireLoadClasses: LoadClass[];
  table: Record<"false" | "true", Partial<Record<number, Array<number | null>>>>;
}> = {
  ua_cd: {
    label: "D116 UA + CD",
    materialHint: "Таблица D116 UA + CD от стр. 14: UA 50/40 + CD 60/27.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: FIRE_LOAD_CLASSES,
    table: {
      false: {
        500: [1400, 1150, 1000, 950, 850],
        600: [1350, 1050, 950, 900, 800],
        700: [1250, 1000, 900, 850, 750],
        800: [1200, 950, 850, 800, null],
        900: [1150, 900, 800, null, null],
        1000: [1100, 900, null, null, null],
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
  wide_span: {
    label: "D116 големи отвори",
    materialHint: "Таблица D116 от стр. 12 за по-големи отвори; използва отделни footnotes за a, b = 800 и b <= 500.",
    loadClasses: ["0.15", "0.30", "0.50", "0.65"],
    fireLoadClasses: FIRE_LOAD_CLASSES,
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
};

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

export const BOARD_TYPE_TO_LAYERS: Record<Exclude<BoardType, "custom">, number> = {
  "knauf_a_12.5": 1,
  "knauf_h2_12.5": 1,
  "knauf_df_12.5": 1,
  "knauf_diamant_12.5": 1,
  "knauf_silentboard_12.5": 1,
  "knauf_2x12.5": 2,
  "knauf_15": 1,
  "knauf_2x15": 2,
  "knauf_18": 1,
  "knauf_20": 1,
  "knauf_25": 1,
  "knauf_25_18": 2,
  "generic_12.5": 1,
  "generic_2x12.5": 2,
  "generic_15": 1,
  "generic_18_or_20": 1,
  "generic_25": 1,
};

export const BOARD_TYPE_MATERIAL_LABELS: Record<BoardType, string> = {
  "knauf_a_12.5": "Knauf A 12.5 mm",
  "knauf_h2_12.5": "Knauf H2 12.5 mm",
  "knauf_df_12.5": "Knauf DF/GKF 12.5 mm",
  "knauf_diamant_12.5": "Knauf Diamant 12.5 mm",
  "knauf_silentboard_12.5": "Knauf Silentboard 12.5 mm",
  "knauf_2x12.5": "Knauf 2x12.5 mm",
  "knauf_15": "Knauf 15 mm",
  "knauf_2x15": "Knauf 2x15 mm",
  "knauf_18": "Knauf 18 mm",
  "knauf_20": "Knauf 20 mm",
  "knauf_25": "Knauf 25 mm",
  "knauf_25_18": "Knauf 25+18 mm",
  "generic_12.5": "Generic 12.5 mm",
  "generic_2x12.5": "Generic 2x12.5 mm",
  "generic_15": "Generic 15 mm",
  "generic_18_or_20": "Generic 18/20 mm",
  "generic_25": "Generic 25 mm",
  custom: "Custom гипсокартон",
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

export function getHangerOptions(systemType: SystemType): HangerOption[] {
  if (systemType === "CUSTOM") return Object.values(HANGER_OPTIONS);
  return Object.values(HANGER_OPTIONS).filter((option) => option.systems.includes(systemType));
}

export function getDefaultHangerType(systemType: SystemType): HangerType {
  if (systemType === "D116") return "ua_m8";
  if (systemType === "D112" || systemType === "D113") return "direct";
  return "generic";
}

export function getAllowedBValues(systemType: SystemType): number[] {
  return Array.from(new Set(
    getBoardOptions(systemType)
      .filter((option) => option.value !== "custom")
      .map((option) => BOARD_TYPE_TO_B[option.value as Exclude<BoardType, "custom">]),
  )).sort((a, b) => a - b);
}

export function getBoardMaterialLabel(boardType: BoardType): string {
  return BOARD_TYPE_MATERIAL_LABELS[normalizeBoardType(boardType)];
}

export function getEffectiveBoardLayers(room: Room, constants: CalculatorConstants): number {
  const boardType = normalizeBoardType(room.boardType);
  if (boardType === "custom") return Math.max(1, Math.round(Number(constants.boardLayers) || 1));
  return BOARD_TYPE_TO_LAYERS[boardType] ?? 1;
}

const BOARD_TYPE_WEIGHT_KG_M2: Record<BoardType, number> = {
  "knauf_a_12.5": 8.5,
  "knauf_h2_12.5": 9,
  "knauf_df_12.5": 10.5,
  "knauf_diamant_12.5": 12.8,
  "knauf_silentboard_12.5": 17.5,
  "knauf_2x12.5": 17,
  "knauf_15": 12,
  "knauf_2x15": 24,
  "knauf_18": 15,
  "knauf_20": 17,
  "knauf_25": 21,
  "knauf_25_18": 36,
  "generic_12.5": 8.5,
  "generic_2x12.5": 17,
  "generic_15": 12,
  "generic_18_or_20": 16,
  "generic_25": 21,
  custom: 8.5,
};

export function estimateLoadKgPerM2(room: Room): number {
  const boardType = normalizeBoardType(room.boardType);
  const boardWeight = BOARD_TYPE_WEIGHT_KG_M2[boardType];
  const insulationWeight = room.fireProtection ? 2 : 0;
  return Number((boardWeight + insulationWeight + (Number(room.additionalLoadKgPerM2) || 0)).toFixed(2));
}

export function getAutomaticLoadClass(room: Room): LoadClass {
  const knPerM2 = estimateLoadKgPerM2(room) * 0.00981;
  if (knPerM2 <= 0.15) return "0.15";
  if (knPerM2 <= 0.30) return "0.30";
  if (knPerM2 <= 0.40) return "0.40";
  if (knPerM2 <= 0.50) return "0.50";
  return "0.65";
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
    loadClasses: D112_VARIANTS.double_cd.loadClasses,
    fireLoadClasses: D112_VARIANTS.double_cd.fireLoadClasses,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultUdAnchorSpacing: 625,
    table: D112_VARIANTS.double_cd.table,
  },
  D113: {
    label: "D113.bg метална подконструкция на едно ниво",
    materialHint: "CD 60/27 на едно ниво; текущият модел е най-близо до примера на Knauf.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: FIRE_LOAD_CLASSES,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
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
    loadClasses: D116_VARIANTS.ua_cd.loadClasses,
    fireLoadClasses: FIRE_LOAD_CLASSES,
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
    defaultUdAnchorSpacing: 625,
    table: D116_VARIANTS.ua_cd.table,
  },
  CUSTOM: {
    label: "Custom - ръчни стойности",
    materialHint: "Ръчна конструкция с начални стойности за стандартен CD грид: a 900 mm, b 500 mm, c 1000 mm.",
    loadClasses: LOAD_CLASSES,
    fireLoadClasses: [],
    defaultLoadClass: "0.30",
    defaultFireProtection: false,
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
  uaLength: 4,
  woodBattenLength: 4,
  metalScrewsPerCrossConnector: 4,
  metalScrewsPerDirectHanger: 2,
  drywallScrewsPerM2: 25,
  boardWidth: 1.2,
  boardLength: 2,
  boardLayers: 1,
  profileEdgeOffsetCm: 10,
  jointTapePerM2: 1.4,
  jointCompoundKgPerM2: 0.3,
  trennFixPerimeterMultiplier: 1,
  mineralWoolEnabled: false,
  mineralWoolThickness: 50,
  anchorsPerDirectHanger: 1,
  wastePercent: 10,
};

export function getConstruction(roomOrType: Room | SystemType = "D113"): ConstructionType {
  const type = typeof roomOrType === "string" ? roomOrType : roomOrType.systemType;
  return CONSTRUCTION_TYPES[type] ?? CONSTRUCTION_TYPES.D113;
}

function withDefaultConstants(constants: CalculatorConstants): CalculatorConstants {
  return { ...DEFAULT_CONSTANTS, ...constants };
}

function getD116Variant(roomOrVariant?: Room | D116Variant): D116Variant {
  const value = typeof roomOrVariant === "string" ? roomOrVariant : roomOrVariant?.d116Variant;
  return value && value in D116_VARIANTS ? value : DEFAULT_D116_VARIANT;
}

function getD112Variant(roomOrVariant?: Room | D112Variant): D112Variant {
  const value = typeof roomOrVariant === "string" ? roomOrVariant : roomOrVariant?.d112Variant;
  return value && value in D112_VARIANTS ? value : DEFAULT_D112_VARIANT;
}

function getHangerType(room: Room): HangerType {
  const value = room.hangerType;
  if (room.systemType === "CUSTOM" && value && value in HANGER_OPTIONS) return value;
  if (value && value in HANGER_OPTIONS && HANGER_OPTIONS[value].systems.includes(room.systemType)) return value;
  return getDefaultHangerType(room.systemType);
}

function hasFireTable(room: Room): boolean {
  return getLoadClasses(room.systemType, true, room.d116Variant, room.d112Variant).length > 0;
}

export function normalizeFireRating(room: Room): FireRating {
  const requested = room.fireRating ?? (room.fireProtection ? "fire_table" : "none");
  if (requested === "none") return "none";
  return hasFireTable(room) ? requested : "none";
}

function getConstructionTable(room: Room): Partial<Record<number, Array<number | null>>> {
  if (room.systemType === "D112") {
    return D112_VARIANTS[getD112Variant(room)].table[String(room.fireProtection) as "false" | "true"] ?? {};
  }
  if (room.systemType === "D116") {
    return D116_VARIANTS[getD116Variant(room)].table[String(room.fireProtection) as "false" | "true"] ?? {};
  }
  return getConstruction(room).table[String(room.fireProtection) as "false" | "true"] ?? {};
}

export function getLoadClasses(systemType: SystemType, fireProtection: boolean, d116Variant: D116Variant = DEFAULT_D116_VARIANT, d112Variant: D112Variant = DEFAULT_D112_VARIANT): LoadClass[] {
  if (systemType === "D112") {
    const variant = D112_VARIANTS[d112Variant] ?? D112_VARIANTS[DEFAULT_D112_VARIANT];
    return fireProtection ? variant.fireLoadClasses : variant.loadClasses;
  }
  if (systemType === "D116") {
    const variant = D116_VARIANTS[d116Variant] ?? D116_VARIANTS[DEFAULT_D116_VARIANT];
    return fireProtection ? variant.fireLoadClasses : variant.loadClasses;
  }
  const construction = getConstruction(systemType);
  return fireProtection ? construction.fireLoadClasses : construction.loadClasses;
}

export function getTableValue(room: Room, c = room.c, loadClass = room.loadClass): number | null {
  const classes = getLoadClasses(room.systemType, room.fireProtection, getD116Variant(room), getD112Variant(room));
  const idx = classes.indexOf(loadClass);
  if (idx === -1) return null;
  return getConstructionTable(room)?.[c]?.[idx] ?? null;
}

export function getValidCValues(room: Room): number[] {
  if (room.systemType === "CUSTOM") return [room.c || 1000];

  const classes = getLoadClasses(room.systemType, room.fireProtection, getD116Variant(room), getD112Variant(room));
  const idx = classes.indexOf(room.loadClass);
  const table = getConstructionTable(room);
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
    d112Variant: DEFAULT_D112_VARIANT,
    d116Variant: DEFAULT_D116_VARIANT,
    customBearingProfile: DEFAULT_CUSTOM_BEARING_PROFILE,
    customMountingProfile: DEFAULT_CUSTOM_MOUNTING_PROFILE,
    customPerimeterProfile: DEFAULT_CUSTOM_PERIMETER_PROFILE,
    loadInputMode: "manual",
    additionalLoadKgPerM2: 0,
    loadClass: construction.defaultLoadClass,
    fireRating: "none",
    fireProtection: construction.defaultFireProtection,
    hangerType: getDefaultHangerType("D113"),
    boardType: "knauf_a_12.5",
    a: 900,
    b: 500,
    c: 600,
    udAnchorSpacing: construction.defaultUdAnchorSpacing,
    overrides: { area: false, a: false, b: false, c: false, udAnchorSpacing: false },
  };
  applyAutoABC(room);
  return room;
}

export function normalizeRoom(room: Room): Room {
  const construction = getConstruction(room.systemType);
  room.d112Variant = getD112Variant(room);
  room.d116Variant = getD116Variant(room);
  room.customBearingProfile = normalizeCustomProfile(room.customBearingProfile, DEFAULT_CUSTOM_BEARING_PROFILE);
  room.customMountingProfile = normalizeCustomProfile(room.customMountingProfile, DEFAULT_CUSTOM_MOUNTING_PROFILE);
  room.customPerimeterProfile = normalizeCustomProfile(room.customPerimeterProfile, DEFAULT_CUSTOM_PERIMETER_PROFILE);
  room.loadInputMode = room.loadInputMode === "auto" ? "auto" : "manual";
  room.additionalLoadKgPerM2 = Math.max(0, Number(room.additionalLoadKgPerM2) || 0);
  room.boardType = normalizeBoardType(room.boardType);
  room.hangerType = getHangerType(room);
  room.fireRating = normalizeFireRating(room);
  room.fireProtection = room.fireRating !== "none";
  const incomingOverrides = room.overrides;
  room.overrides = {
    area: Boolean(incomingOverrides?.area),
    a: Boolean(incomingOverrides?.a),
    b: Boolean(incomingOverrides?.b),
    c: Boolean(incomingOverrides?.c),
    udAnchorSpacing: Boolean(incomingOverrides?.udAnchorSpacing),
  };
  const loadClasses = getLoadClasses(room.systemType, room.fireProtection, room.d116Variant, room.d112Variant);
  if (room.loadInputMode === "auto") {
    room.loadClass = coerceLoadClassToAvailable(getAutomaticLoadClass(room), loadClasses);
  }
  if (!loadClasses.includes(room.loadClass)) {
    room.loadClass = loadClasses[0] ?? construction.defaultLoadClass;
  }
  if (!Number(room.udAnchorSpacing)) room.udAnchorSpacing = construction.defaultUdAnchorSpacing;
  return room;
}

function coerceLoadClassToAvailable(loadClass: LoadClass, available: LoadClass[]): LoadClass {
  if (!available.length) return loadClass;
  if (available.includes(loadClass)) return loadClass;
  const requested = Number(loadClass);
  return available.find((value) => Number(value) >= requested) ?? available[available.length - 1];
}

function normalizeCustomProfile(value: unknown, fallback: CustomProfileType): CustomProfileType {
  return typeof value === "string" && value in CUSTOM_PROFILE_OPTIONS ? value as CustomProfileType : fallback;
}

function getTableForAuto(systemType: SystemType, fireProtection: boolean, d116Variant = DEFAULT_D116_VARIANT, d112Variant = DEFAULT_D112_VARIANT) {
  if (systemType === "D112") {
    return D112_VARIANTS[d112Variant].table[String(fireProtection) as "false" | "true"] ?? {};
  }
  if (systemType === "D116") {
    return D116_VARIANTS[d116Variant].table[String(fireProtection) as "false" | "true"] ?? {};
  }
  return getConstruction(systemType).table[String(fireProtection) as "false" | "true"] ?? {};
}

export function getAutoABC(loadClass: LoadClass, fireProtection: boolean, boardType: BoardType, systemType: SystemType = "D113", d116Variant: D116Variant = DEFAULT_D116_VARIANT, d112Variant: D112Variant = DEFAULT_D112_VARIANT) {
  const construction = getConstruction(systemType);
  if (systemType === "CUSTOM") {
    return {
      a: 900,
      c: 1000,
      b: boardType === "custom" ? 500 : BOARD_TYPE_TO_B[boardType] ?? 500,
      udAnchorSpacing: construction.defaultUdAnchorSpacing,
    };
  }

  const table = getTableForAuto(systemType, fireProtection, d116Variant, d112Variant);
  const classes = getLoadClasses(systemType, fireProtection, d116Variant, d112Variant);
  const idx = Math.max(0, classes.indexOf(loadClass));
  const sortedC = Object.keys(table).map(Number).sort((a, b) => a - b);
  const firstValid = sortedC.find((c) => table[c]?.[idx] != null) ?? 600;
  return {
    a: table[firstValid]?.[idx] ?? 900,
    c: firstValid,
    b: boardType === "custom" ? 500 : BOARD_TYPE_TO_B[boardType] ?? 500,
    udAnchorSpacing: construction.defaultUdAnchorSpacing,
  };
}

export function applyAutoABC(room: Room): void {
  normalizeRoom(room);
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType, getD116Variant(room), getD112Variant(room));
  if (!room.overrides.c) room.c = auto.c;
  if (!room.overrides.a) room.a = auto.a;
  if (!room.overrides.b) room.b = auto.b;
  if (!room.overrides.udAnchorSpacing) room.udAnchorSpacing = auto.udAnchorSpacing;
}

export function syncSpacingFromKnaufTable(room: Room, { keepC = true } = {}): void {
  if (room.systemType === "CUSTOM") {
    const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType, getD116Variant(room), getD112Variant(room));
    if (!room.overrides.c) room.c = auto.c;
    if (!room.overrides.a) room.a = auto.a;
    if (!room.overrides.b) room.b = auto.b;
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

export function getFireCertificationCheck(room: Room): FireCertificationCheck {
  normalizeRoom(room);
  const issues: FireCertificationIssue[] = [];
  const fireRating = room.fireRating ?? "none";

  if (fireRating === "none") {
    return {
      status: "none",
      label: "Без огнезащита",
      summary: "EI проверка не се прилага за тази стая.",
      issues,
    };
  }

  if (fireRating === "fire_table") {
    issues.push({
      code: "specific-fire-rating-missing",
      type: "missing",
      message: "Избрана е обща fire таблица, но не е избран конкретен EI клас и посока.",
      allowedValues: FIRE_RATING_OPTIONS.filter((option) => SPECIFIC_FIRE_RATINGS.includes(option.value)).map((option) => option.label),
    });
  }

  if (room.systemType === "CUSTOM") {
    issues.push({
      code: "custom-system-not-certifiable",
      type: "invalid",
      message: "Custom конструкция не може да бъде потвърдена като пълна Knauf EI система. Избери D112, D113 или D116 системен вариант.",
      allowedValues: ["D112", "D113", "D116"],
    });
  }

  if (fireRating !== "fire_table") {
    const specificRating = fireRating as Exclude<FireRating, "none" | "fire_table">;
    const allowedSystems = FIRE_CERT_ALLOWED_SYSTEMS[specificRating];
    if (!allowedSystems.includes(room.systemType)) {
      issues.push({
        code: "fire-system-not-allowed",
        type: "invalid",
        message: `${getFireRatingLabel(specificRating)} не е моделирана като пълна EI система за ${room.systemType}.`,
        allowedValues: allowedSystems,
      });
    }

    const tableValue = getTableValue(room);
    if (!tableValue) {
      issues.push({
        code: "fire-table-value-missing",
        type: "invalid",
        message: "Няма допустима fire таблица за избраните система, товар и c.",
        allowedValues: getValidCValues(room).map((value) => `c = ${value} mm`),
      });
    } else if (room.a > tableValue) {
      issues.push({
        code: "fire-hanger-spacing-invalid",
        type: "invalid",
        message: `a = ${room.a} mm е по-голямо от допустимото за избраната fire таблица.`,
        allowedValues: [`a <= ${tableValue} mm`],
      });
    }

    const validCValues = getValidCValues(room);
    if (!validCValues.includes(room.c)) {
      issues.push({
        code: "fire-c-spacing-invalid",
        type: "invalid",
        message: `c = ${room.c} mm не е допустимо за избраната fire комбинация.`,
        allowedValues: validCValues.map((value) => `${value} mm`),
      });
    }

    const maxB = FIRE_CERT_MAX_B[specificRating];
    if (room.b > maxB) {
      issues.push({
        code: "fire-b-spacing-invalid",
        type: "invalid",
        message: `b = ${room.b} mm е прекалено голямо за ${getFireRatingLabel(specificRating)}.`,
        allowedValues: getAllowedBValues(room.systemType).filter((value) => value <= maxB).map((value) => `${value} mm`),
      });
    }

    const allowedBoards = FIRE_CERT_BOARD_OPTIONS[specificRating];
    if (!allowedBoards.includes(room.boardType)) {
      issues.push({
        code: "fire-board-type-invalid",
        type: "invalid",
        message: `Типът плоскост "${getBoardMaterialLabel(room.boardType)}" не покрива моделираните изисквания за ${getFireRatingLabel(specificRating)}.`,
        allowedValues: allowedBoards.map(getBoardMaterialLabel),
      });
    }

    const requiredCapacity = FIRE_CERT_REQUIRED_HANGER_CAPACITY[specificRating];
    const hanger = HANGER_OPTIONS[getHangerType(room)];
    if (hanger.capacityKn == null) {
      issues.push({
        code: "fire-hanger-missing",
        type: "missing",
        message: "Не е избран конкретен системен окачвач за EI проверката.",
        allowedValues: getHangerOptions(room.systemType).filter((option) => option.value !== "generic").map((option) => option.label),
      });
    } else if (hanger.capacityKn < requiredCapacity) {
      issues.push({
        code: "fire-hanger-capacity-invalid",
        type: "invalid",
        message: `Избраният окачвач е ${hanger.capacityKn.toFixed(2)} kN, а тази EI конфигурация изисква минимум ${requiredCapacity.toFixed(2)} kN.`,
        allowedValues: getHangerOptions(room.systemType).filter((option) => (option.capacityKn ?? 0) >= requiredCapacity).map((option) => option.label),
      });
    }

    const udRule = getUdAnchoringRule(room);
    if (room.udAnchorSpacing > udRule.maxSpacing) {
      issues.push({
        code: "fire-ud-spacing-invalid",
        type: "invalid",
        message: `Периферните дюбели са през ${room.udAnchorSpacing} mm, а EI/носещата връзка допуска до ${udRule.maxSpacing} mm.`,
        allowedValues: [`<= ${udRule.maxSpacing} mm`],
      });
    }

    if (room.systemType === "D116" && !["ua_m8", "nonius"].includes(getHangerType(room))) {
      issues.push({
        code: "fire-d116-hanger-invalid",
        type: "invalid",
        message: "За D116 EI проверката изисква UA/M8 или нониус окачване.",
        allowedValues: ["UA/M8 окачване", "Нониус окачвач"],
      });
    }
  }

  const status: FireCertificationStatus = issues.some((issue) => issue.type === "invalid")
    ? "invalid"
    : issues.length
      ? "incomplete"
      : "complete";

  return {
    status,
    label: status === "complete" ? "Пълна EI система" : status === "incomplete" ? "EI непълна" : "EI грешна",
    summary: status === "complete"
      ? "Всички моделирани EI условия са избрани и съвместими."
      : status === "incomplete"
        ? "Липсват избори, преди системата да може да се потвърди."
        : "Има стойности, които не отговарят на моделираните EI ограничения.",
    issues,
  };
}

export function getValidationWarnings(room: Room): ValidationWarning[] {
  normalizeRoom(room);
  if (room.systemType === "CUSTOM") {
    return [room.a, room.b, room.c, room.udAnchorSpacing].every((value) => Number.isFinite(value) && value > 0)
      ? []
      : [{
        code: "custom-positive-spacing",
        severity: "error",
        message: "Custom конструкцията изисква положителни стойности за a, b, c и UD стъпка.",
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

  warnings.push(...getFootnoteWarnings(room));

  const udRule = getUdAnchoringRule(room);
  if (room.udAnchorSpacing > udRule.maxSpacing) {
    warnings.push({
      code: "ud-anchor-spacing",
      severity: "warning",
      message: `За ${udRule.mode} закрепването на UD профила трябва да е до ${udRule.maxSpacing} mm. Текущата стойност е ${room.udAnchorSpacing} mm.`,
    });
  }

  if (room.fireProtection && room.b > 625) {
    warnings.push({
      code: "fire-board-spacing-review",
      severity: "warning",
      message: "При огнезащитни решения провери дали избраното b е допустимо за конкретния слой плоскости и EI конфигурация.",
    });
  }

  if (room.systemType === "D116") {
    warnings.push({
      code: "d116-ua-system-review",
      severity: "warning",
      message: "D116 използва UA + CD; UW елементите за удължаване и окачването трябва да се проверят спрямо избрания системен детайл.",
    });
  }

  if (room.systemType === "D111") {
    warnings.push({
      code: "d111-wood-system-review",
      severity: "warning",
      message: "D111 е дървена подконструкция; размерът и класът на дървените летви трябва да се потвърдят спрямо конкретния системен вариант.",
    });
  }

  return warnings;
}

function getFootnoteWarnings(room: Room): ValidationWarning[] {
  const warnings: ValidationWarning[] = [];
  const highLoadClasses: LoadClass[] = ["0.40", "0.50", "0.65"];
  const hanger = HANGER_OPTIONS[getHangerType(room)];

  if (highLoadClasses.includes(room.loadClass) && hanger.capacityKn !== 0.4) {
    warnings.push({
      code: "hanger-load-class-040",
      severity: "warning",
      message: "За този клас натоварване документацията отбелязва използване на окачвачи с клас носимоспособност 0.40 kN.",
    });
  }

  if (requiresB500FootnoteReview(room)) {
    warnings.push({
      code: "b500-footnote-review",
      severity: "warning",
      message: "Бележките към таблицата ограничават тази комбинация до b <= 500 mm. Провери избрания тип плоскост или намали b.",
    });
  }

  if (requiresB800FootnoteReview(room)) {
    warnings.push({
      code: "b800-footnote-review",
      severity: "warning",
      message: "Бележките към таблицата маркират тази стойност с ограничение за b = 800 mm. Провери дали избраният слой плоскости отговаря.",
    });
  }

  if (room.fireRating && room.fireRating !== "none") {
    const eiNote = room.fireRating === "fire_table"
      ? "Избрана е огнезащита по D11 таблица. Конкретният EI клас и посока отдолу/отгоре още трябва да се проверят отделно."
      : `Избран е ${FIRE_RATING_OPTIONS.find((option) => option.value === room.fireRating)?.label}. Приложението използва fire таблиците, но това не замества проверка на пълния системен детайл.`;
    warnings.push({
      code: "fire-ei-not-specific",
      severity: "warning",
      message: eiNote,
    });
  }

  if (room.systemType === "D116") {
    warnings.push({
      code: "d116-hanger-type-not-modeled",
      severity: "warning",
      message: "За D116 допустимите разстояния зависят от конкретния системен детайл. Провери избрания окачвач спрямо проекта.",
    });

    if (getD116Variant(room) === "wide_span" && !room.fireProtection && room.loadClass === "0.30" && room.c <= 600 && room.a > 1700) {
      warnings.push({
        code: "d116-wide-span-a1700-footnote",
        severity: "error",
        message: "За този D116 вариант бележката към таблицата ограничава a до 1700 mm.",
      });
    }
  }

  warnings.push(...getHangerWarnings(room));

  return warnings;
}

function getHangerWarnings(room: Room): ValidationWarning[] {
  if (room.systemType === "CUSTOM" || room.systemType === "D111") return [];

  const hangerType = getHangerType(room);
  const hanger = HANGER_OPTIONS[hangerType];
  const highLoadClasses: LoadClass[] = ["0.40", "0.50", "0.65"];
  const warnings: ValidationWarning[] = [];

  if (highLoadClasses.includes(room.loadClass) && hanger.capacityKn !== 0.4) {
    warnings.push({
      code: "hanger-type-capacity-review",
      severity: "warning",
      message: `Товар ${room.loadClass} kN/m2 обикновено изисква окачвач/закрепване 0.40 kN. Избрано е: ${hanger.label}.`,
    });
  }

  if (room.systemType === "D116" && hangerType !== "ua_m8" && hangerType !== "nonius") {
    warnings.push({
      code: "d116-hanger-type-review",
      severity: "warning",
      message: "За D116 обичайно се проверява UA/M8 или нониус окачване според системния детайл.",
    });
  }

  return warnings;
}

function getFireRatingLabel(fireRating: FireRating): string {
  return FIRE_RATING_OPTIONS.find((option) => option.value === fireRating)?.label ?? fireRating;
}

function requiresB500FootnoteReview(room: Room): boolean {
  if (room.b <= 500) return false;

  if (room.systemType === "D113") {
    return room.loadClass === "0.65" || (room.fireProtection && room.loadClass === "0.50");
  }

  if (room.systemType === "D112") {
    return room.loadClass === "0.50" || room.loadClass === "0.65";
  }

  if (room.systemType === "D116") {
    return (room.fireProtection && (room.loadClass === "0.50" || room.loadClass === "0.65"))
      || (!room.fireProtection && getD116Variant(room) === "ua_cd" && room.c === 1000 && room.loadClass === "0.30")
      || (!room.fireProtection && getD116Variant(room) === "wide_span" && room.b > 500 && room.loadClass === "0.30" && room.c <= 600);
  }

  return false;
}

function requiresB800FootnoteReview(room: Room): boolean {
  if (room.b !== 800) return false;
  if (room.systemType === "D113" && !room.fireProtection && room.c === 700 && room.loadClass === "0.65") return false;

  if (room.systemType === "D112") {
    return !room.fireProtection && (room.loadClass === "0.40" || room.loadClass === "0.50" || room.loadClass === "0.65");
  }

  if (room.systemType === "D116") {
    return !room.fireProtection;
  }

  return room.systemType === "D113";
}

export function getUdAnchoringRule(room: Room): UdAnchoringRule {
  if (room.fireProtection) {
    return {
      mode: "огнезащита",
      maxSpacing: 625,
      note: "При огнезащита и носещи връзки използвай стъпка до 625 mm.",
    };
  }

  if (room.systemType === "D112" || room.systemType === "D113" || room.systemType === "D116") {
    return {
      mode: "носеща връзка",
      maxSpacing: 625,
      note: "За металните D11 системи приложението приема носеща периферна връзка до 625 mm.",
    };
  }

  return {
    mode: "неносеща връзка",
    maxSpacing: 1000,
    note: "При неносеща периферна връзка ориентировъчната стъпка е до около 1000 mm.",
  };
}

export function countBySpacing(lengthCm: number, spacingMm: number, edgeCm: number): number {
  if (!lengthCm || !spacingMm) return 0;
  const start = Math.min(edgeCm, lengthCm / 2);
  const end = Math.max(start, lengthCm - edgeCm);
  const effectiveLength = Math.max(0, end - start);
  return Math.ceil(effectiveLength / (spacingMm / 10)) + 1;
}

export function buildPositions(limitCm: number, spacingMm: number, edgeCm = DEFAULT_CONSTANTS.profileEdgeOffsetCm): number[] {
  if (!limitCm || !spacingMm) return [];
  const count = countBySpacing(limitCm, spacingMm, edgeCm);
  const start = Math.min(edgeCm, limitCm / 2);
  if (count === 1) return [start];

  const finalInside = Math.max(start, limitCm - edgeCm);
  const actualSpacing = (finalInside - start) / (count - 1);
  return Array.from({ length: count }, (_, idx) => start + idx * actualSpacing);
}

export function buildLinearPositions(limitCm: number, spacingCm: number, firstOffsetCm = DEFAULT_CONSTANTS.profileEdgeOffsetCm): number[] {
  if (!limitCm || !spacingCm) return [];
  const max = Math.max(0, limitCm - firstOffsetCm);
  const positions: number[] = [];
  for (let current = firstOffsetCm; current <= max; current += spacingCm) positions.push(current);
  return positions;
}

function getStaggerOffsetCm(rowIndex: number, profileLengthCm: number): number {
  const offsets = [0, profileLengthCm / 2, profileLengthCm / 4, (profileLengthCm * 3) / 4];
  return offsets[rowIndex % offsets.length] ?? 0;
}

function hasNearbyPosition(positionsCm: number[], targetCm: number, maxDistanceCm: number): boolean {
  return positionsCm.some((position) => Math.abs(position - targetCm) <= maxDistanceCm);
}

function canPlaceHanger(positionsCm: number[], candidateCm: number, roomLengthCm: number, minHangerDistanceCm: number): boolean {
  if (candidateCm <= 0 || candidateCm >= roomLengthCm) return false;
  return positionsCm.every((position) => Math.abs(position - candidateCm) >= minHangerDistanceCm);
}

function getGridStepCm(positionsCm: number[]): number | null {
  if (positionsCm.length < 2) return null;
  const gaps = positionsCm.slice(1).map((position, index) => position - positionsCm[index]).filter((gap) => gap > 0);
  if (!gaps.length) return null;
  return gaps[0] ?? null;
}

function getJointCandidatePositions(
  jointCm: number,
  previousBoundaryCm: number,
  nextBoundaryCm: number,
  roomLengthCm: number,
): number[] {
  const leftSegmentLength = jointCm - previousBoundaryCm;
  const rightSegmentLength = nextBoundaryCm - jointCm;
  const preferLeft = leftSegmentLength >= rightSegmentLength;
  const preferred10 = preferLeft ? jointCm - 10 : jointCm + 10;
  const secondary10 = preferLeft ? jointCm + 10 : jointCm - 10;
  const preferred15 = preferLeft ? jointCm - 15 : jointCm + 15;
  const secondary15 = preferLeft ? jointCm + 15 : jointCm - 15;
  return [preferred10, secondary10, preferred15, secondary15]
    .filter((value, index, values) => value > 0 && value < roomLengthCm && values.indexOf(value) === index);
}

function getJointSnapCandidates(supportCm: number, roomLengthCm: number): number[] {
  return [supportCm - 10, supportCm + 10, supportCm - 15, supportCm + 15]
    .filter((value, index, values) => value > 0 && value < roomLengthCm && values.indexOf(value) === index);
}

function sortNumericAscending(values: number[]): number[] {
  return [...values].sort((left, right) => left - right);
}

function isValidJointCandidate(
  jointCm: number,
  previousBoundaryCm: number,
  lineLengthCm: number,
  profileLengthCm: number,
  remainingSplices: number,
): boolean {
  const currentSegmentLength = jointCm - previousBoundaryCm;
  const remainingLength = lineLengthCm - jointCm;
  const remainingSegments = remainingSplices + 1;

  return currentSegmentLength > 0
    && currentSegmentLength <= profileLengthCm
    && remainingLength > 0
    && remainingLength <= remainingSegments * profileLengthCm;
}

function selectSingleJointSupport(
  supports: number[],
  lineLengthCm: number,
  profileLengthCm: number,
  previousRowPoints: number[],
): number | null {
  const midpoint = lineLengthCm / 2;
  const candidates = supports.filter((jointCm) => (
    jointCm > 30
    && lineLengthCm - jointCm > 30
    && jointCm <= profileLengthCm
    && lineLengthCm - jointCm <= profileLengthCm
    && previousRowPoints.every((point) => Math.abs(point - jointCm) >= 60)
  ));
  if (!candidates.length) return null;

  return candidates
    .map((jointCm) => ({
      jointCm,
      balanceScore: Math.min(jointCm, lineLengthCm - jointCm),
      midpointDistance: Math.abs(midpoint - jointCm),
    }))
    .sort((left, right) => (
      right.balanceScore - left.balanceScore
      || left.midpointDistance - right.midpointDistance
      || left.jointCm - right.jointCm
    ))[0]?.jointCm ?? null;
}

function allJointsAreNearHangers(
  carrierExtensions: ExtensionLineLayout[],
  hangersCm: number[],
  hangerNearJointMaxDistanceCm: number,
): boolean {
  return carrierExtensions.every((line) => (
    line.pointsCm.every((jointCm) => hasNearbyPosition(hangersCm, jointCm, hangerNearJointMaxDistanceCm))
  ));
}

export function buildStaggeredExtensionLayout(
  lineCount: number,
  lineLengthCm: number,
  profileLengthM: number,
  supportPositionsCm: number[] = [],
  minDistanceBetweenAlignedJointsCm = 80,
  hangerNearJointMaxDistanceCm = 15,
): ExtensionLineLayout[] {
  const normalizedLineCount = Math.max(0, Math.floor(lineCount));
  const profileLengthCm = Math.max(0, profileLengthM * 100);
  if (!normalizedLineCount || !profileLengthCm || !lineLengthCm) {
    return Array.from({ length: normalizedLineCount }, (_, lineIndex) => ({ lineIndex, pointsCm: [] }));
  }
  if (lineLengthCm <= profileLengthCm) {
    return Array.from({ length: normalizedLineCount }, (_, lineIndex) => ({ lineIndex, pointsCm: [] }));
  }

  const supports = sortNumericAscending(
    supportPositionsCm.filter((position) => Number.isFinite(position) && position > 0 && position < lineLengthCm),
  );
  const supportStepCm = getGridStepCm(supports);
  const stepHangers = supportStepCm ? Math.max(1, Math.round(profileLengthCm / supportStepCm)) : 0;
  const rowStartPattern = stepHangers
    ? [stepHangers, Math.max(1, stepHangers - 2), Math.max(1, stepHangers - 1), Math.max(1, stepHangers - 3)]
    : [];
  const spliceCount = Math.max(0, Math.ceil(lineLengthCm / profileLengthCm) - 1);

  const layout: ExtensionLineLayout[] = [];

  for (let lineIndex = 0; lineIndex < normalizedLineCount; lineIndex += 1) {
    const previousRowPoints = layout[lineIndex - 1]?.pointsCm ?? [];

    if (spliceCount === 1 && supports.length) {
      const jointCm = selectSingleJointSupport(supports, lineLengthCm, profileLengthCm, previousRowPoints);
      layout.push({ lineIndex, pointsCm: jointCm == null ? [] : [jointCm] });
      continue;
    }

    if (supports.length && stepHangers > 0) {
      const pointsCm: number[] = [];
      let supportIndex = rowStartPattern[lineIndex % rowStartPattern.length] ?? stepHangers;

      while (supportIndex < supports.length) {
        const supportCm = supports[supportIndex] ?? Number.NaN;
        const previousBoundaryCm = pointsCm[pointsCm.length - 1] ?? 0;
        const remainingSplices = Math.max(0, spliceCount - pointsCm.length - 1);
        const jointCm = Number.isFinite(supportCm)
          ? getJointSnapCandidates(supportCm, lineLengthCm).find((candidateCm) => (
            candidateCm > 30
            && lineLengthCm - candidateCm > 30
            && hasNearbyPosition(supports, candidateCm, hangerNearJointMaxDistanceCm)
            && pointsCm.every((point) => Math.abs(point - candidateCm) >= minDistanceBetweenAlignedJointsCm)
            && previousRowPoints.every((point) => Math.abs(point - candidateCm) >= 60)
            && isValidJointCandidate(candidateCm, previousBoundaryCm, lineLengthCm, profileLengthCm, remainingSplices)
          )) ?? Number.NaN
          : Number.NaN;
        const isFarFromAdjacentRow = previousRowPoints.every((point) => Math.abs(point - jointCm) >= 60);
        if (
          Number.isFinite(jointCm)
          && isFarFromAdjacentRow
        ) {
          pointsCm.push(jointCm);
        }
        supportIndex += stepHangers;
      }

      layout.push({ lineIndex, pointsCm });
      continue;
    }

    const pointsCm: number[] = [];
    const offsetCm = getStaggerOffsetCm(lineIndex, profileLengthCm);
    let jointCm = profileLengthCm - offsetCm;

    while (jointCm < lineLengthCm) {
      const distanceToNearestExisting = pointsCm.length
        ? Math.min(...pointsCm.map((point) => Math.abs(point - jointCm)))
        : Number.POSITIVE_INFINITY;
      if (
        jointCm > 30
        && lineLengthCm - jointCm > 30
        && distanceToNearestExisting >= minDistanceBetweenAlignedJointsCm
        && previousRowPoints.every((point) => Math.abs(point - jointCm) >= 60)
      ) {
        pointsCm.push(jointCm);
      }
      jointCm += profileLengthCm;
    }

    layout.push({ lineIndex, pointsCm });
  }

  return layout;
}

export function buildLayoutHangerPositions(
  roomLengthCm: number,
  hangerSpacingCm: number,
  firstHangerOffsetCm: number,
  carrierExtensions: ExtensionLineLayout[],
  hangerNearJointMaxDistanceCm = 15,
  minHangerDistanceCm = 20,
): number[] {
  const hangers = buildLinearPositions(roomLengthCm, hangerSpacingCm, firstHangerOffsetCm);
  if (allJointsAreNearHangers(carrierExtensions, hangers, hangerNearJointMaxDistanceCm)) {
    return hangers;
  }

  for (const line of carrierExtensions) {
    const boundaries = [0, ...line.pointsCm, roomLengthCm];
    line.pointsCm.forEach((jointCm, jointIndex) => {
      if (hasNearbyPosition(hangers, jointCm, hangerNearJointMaxDistanceCm)) return;

      const candidatePositions = getJointCandidatePositions(
        jointCm,
        boundaries[jointIndex] ?? 0,
        boundaries[jointIndex + 2] ?? roomLengthCm,
        roomLengthCm,
      );

      const candidate = candidatePositions.find((value) => canPlaceHanger(hangers, value, roomLengthCm, minHangerDistanceCm));
      if (candidate != null) hangers.push(candidate);
    });
  }

  return sortNumericAscending(hangers);
}

export function buildSuspendedCeilingLayout(input: SuspendedCeilingLayoutInput): SuspendedCeilingLayout {
  const {
    roomLengthCm,
    roomWidthCm,
    profileLengthCm = 400,
    carrierRowSpacingCm,
    hangerSpacingCm,
    firstHangerOffsetCm = 10,
    minDistanceBetweenAlignedJointsCm = 80,
    hangerNearJointMaxDistanceCm = 15,
  } = input;

  const carrierRowsYcm = buildLinearPositions(roomWidthCm, carrierRowSpacingCm, firstHangerOffsetCm);
  const mountingRowsXcm = buildLinearPositions(roomLengthCm, carrierRowSpacingCm, firstHangerOffsetCm);
  const regularHangerPositions = buildLinearPositions(roomLengthCm, hangerSpacingCm, firstHangerOffsetCm);
  const carrierExtensions = buildStaggeredExtensionLayout(
    carrierRowsYcm.length,
    roomLengthCm,
    profileLengthCm / 100,
    regularHangerPositions,
    minDistanceBetweenAlignedJointsCm,
    hangerNearJointMaxDistanceCm,
  );
  const hangerPositionsCm = buildLayoutHangerPositions(
    roomLengthCm,
    hangerSpacingCm,
    firstHangerOffsetCm,
    carrierExtensions,
    hangerNearJointMaxDistanceCm,
  );

  return {
    carrierRowsYcm,
    mountingRowsXcm,
    hangerPositionsCm,
    carrierExtensions,
  };
}

function buildSegmentsFromBreakpoints(totalLengthCm: number, breakpointsCm: number[]): CeilingSegment[] {
  const points = [0, ...breakpointsCm, totalLengthCm]
    .filter((value, index, values) => value >= 0 && value <= totalLengthCm && values.indexOf(value) === index)
    .sort((left, right) => left - right);
  const segments: CeilingSegment[] = [];
  for (let index = 1; index < points.length; index += 1) {
    const fromCm = points[index - 1] ?? 0;
    const toCm = points[index] ?? totalLengthCm;
    if (toCm <= fromCm) continue;
    segments.push({ fromCm, toCm, lengthCm: Number((toCm - fromCm).toFixed(2)) });
  }
  return segments;
}

function splitSegmentForCut(segment: CeilingSegment, maxLengthCm: number): CeilingSegment[] {
  if (!Number.isFinite(segment.lengthCm) || segment.lengthCm <= 0) return [];
  if (!Number.isFinite(maxLengthCm) || maxLengthCm <= 0 || segment.lengthCm <= maxLengthCm + 0.0001) {
    return [{ ...segment, lengthCm: Number(segment.lengthCm.toFixed(2)) }];
  }

  const partCount = Math.ceil(segment.lengthCm / maxLengthCm);
  const baseLengthCm = Number((segment.lengthCm / partCount).toFixed(2));
  const parts: CeilingSegment[] = [];
  let fromCm = segment.fromCm;

  for (let index = 0; index < partCount; index += 1) {
    const remainingLengthCm = Number((segment.toCm - fromCm).toFixed(2));
    const isLast = index === partCount - 1;
    const lengthCm = isLast ? remainingLengthCm : Math.min(baseLengthCm, remainingLengthCm);
    const toCm = Number((fromCm + lengthCm).toFixed(2));
    parts.push({
      fromCm: Number(fromCm.toFixed(2)),
      toCm,
      lengthCm: Number(lengthCm.toFixed(2)),
    });
    fromCm = toCm;
  }

  return parts;
}

function splitSegmentsForCut(segments: CeilingSegment[], maxLengthCm: number): CeilingSegment[] {
  return segments.flatMap((segment) => splitSegmentForCut(segment, maxLengthCm));
}

function buildMountingCutSegments(room: Room, result: CalcResult, carrierPositionsCm: number[], maxLengthCm: number): CeilingSegment[] {
  const segments = room.systemType === "D113"
    ? buildSegmentsFromBreakpoints(result.W, carrierPositionsCm)
    : [{ fromCm: 0, toCm: result.W, lengthCm: Number(result.W.toFixed(2)) }];

  return splitSegmentsForCut(segments, maxLengthCm);
}

export function buildCutOptimizationInput(room: Room, result: CalcResult, constants: CalculatorConstants = DEFAULT_CONSTANTS): CutOptimizationInput {
  constants = withDefaultConstants(constants);
  const cdStockLengthCm = constants.cdLength * 100;
  const udStockLengthCm = constants.udLength * 100;
  const layout = buildSuspendedCeilingLayout({
    roomWidthCm: result.W,
    roomLengthCm: result.L,
    profileLengthCm: cdStockLengthCm,
    carrierRowSpacingCm: room.c / 10,
    hangerSpacingCm: room.a / 10,
    firstHangerOffsetCm: constants.profileEdgeOffsetCm,
  });
  const mountingRowsXcm = buildLinearPositions(result.L, room.b / 10, constants.profileEdgeOffsetCm);

  return {
    carrierRows: layout.carrierRowsYcm.map((_, rowIndex) => ({
      rowIndex,
      segments: splitSegmentsForCut(
        buildSegmentsFromBreakpoints(result.L, layout.carrierExtensions.find((line) => line.lineIndex === rowIndex)?.pointsCm ?? []),
        cdStockLengthCm,
      ),
    })),
    mountingRows: mountingRowsXcm.map((_, rowIndex) => ({
      rowIndex,
      segments: buildMountingCutSegments(room, result, layout.carrierRowsYcm, cdStockLengthCm),
    })),
    udProfiles: {
      segments: splitSegmentsForCut([
        { fromCm: 0, toCm: result.L, lengthCm: Number(result.L.toFixed(2)) },
        { fromCm: 0, toCm: result.L, lengthCm: Number(result.L.toFixed(2)) },
        { fromCm: 0, toCm: result.W, lengthCm: Number(result.W.toFixed(2)) },
        { fromCm: 0, toCm: result.W, lengthCm: Number(result.W.toFixed(2)) },
      ], udStockLengthCm),
    },
  };
}

function buildCutPieces(input: CutOptimizationInput): CutPiece[] {
  const carrierPieces = input.carrierRows.flatMap((row) => row.segments.map((segment, index) => ({
    id: `carrier-${row.rowIndex}-${index}`,
    type: "carrier" as const,
    lengthCm: segment.lengthCm,
  })));
  const mountingPieces = input.mountingRows.flatMap((row) => row.segments.map((segment, index) => ({
    id: `mounting-${row.rowIndex}-${index}`,
    type: "mounting" as const,
    lengthCm: segment.lengthCm,
  })));
  const udPieces = input.udProfiles.segments.map((segment, index) => ({
    id: `ud-${index}`,
    type: "ud" as const,
    lengthCm: segment.lengthCm,
  }));
  return [...carrierPieces, ...mountingPieces, ...udPieces];
}

export function buildRoomCutPieces(
  room: Room,
  result: CalcResult,
  constants: CalculatorConstants = DEFAULT_CONSTANTS,
): CutPiece[] {
  const pieces = buildCutPieces(buildCutOptimizationInput(room, result, constants));
  return pieces.map((piece) => ({
    ...piece,
    id: `${room.id}-${piece.id}`,
    roomId: room.id,
    roomName: room.name,
  }));
}

function finalizeCutBar(
  bar: Omit<CutBar, "segments" | "usedCm" | "wasteCm" | "estimatedKerfLossCm">,
  kerfCm: number,
  includeKerfInFitCheck: boolean,
): CutBar {
  let cursor = 0;
  const segments = bar.pieces.map((piece, index) => {
    const startCm = cursor;
    const endCm = startCm + piece.lengthCm;
    cursor = endCm + (includeKerfInFitCheck && index < bar.pieces.length - 1 ? kerfCm : 0);
    return { pieceId: piece.id, type: piece.type, lengthCm: piece.lengthCm, startCm, endCm };
  });
  const pieceLengthCm = bar.pieces.reduce((sum, piece) => sum + piece.lengthCm, 0);
  const estimatedKerfLossCm = Number((Math.max(0, bar.pieces.length - 1) * kerfCm).toFixed(2));
  const usedCm = Number((pieceLengthCm + (includeKerfInFitCheck ? estimatedKerfLossCm : 0)).toFixed(2));
  const uniqueTypes = [...new Set(bar.pieces.map((piece) => piece.type))];
  const hasCd = bar.pieces.some((piece) => piece.type === "carrier" || piece.type === "mounting");
  const hasUd = bar.pieces.some((piece) => piece.type === "ud");
  const type = hasCd && hasUd
    ? "mixed"
    : hasCd && uniqueTypes.length > 1
      ? "cd"
      : uniqueTypes[0] ?? "mixed";
  return {
    ...bar,
    type,
    segments,
    usedCm,
    wasteCm: Number((bar.stockLengthCm - usedCm).toFixed(2)),
    estimatedKerfLossCm,
  };
}

function getBarUsedCm(pieces: CutPiece[], kerfCm: number, includeKerfInFitCheck: boolean): number {
  const pieceLengthCm = pieces.reduce((sum, piece) => sum + piece.lengthCm, 0);
  const kerfLossCm = includeKerfInFitCheck ? Math.max(0, pieces.length - 1) * kerfCm : 0;
  return Number((pieceLengthCm + kerfLossCm).toFixed(2));
}

function canFitPieceInBar(pieces: CutPiece[], piece: CutPiece, stockLengthCm: number, kerfCm: number, includeKerfInFitCheck: boolean): boolean {
  const nextUsed = getBarUsedCm([...pieces, piece], kerfCm, includeKerfInFitCheck);
  return nextUsed <= stockLengthCm + 0.0001;
}

function extractPerfectPairBars(
  pieces: CutPiece[],
  stockLengthCm: number,
  kerfCm: number,
  includeKerfInFitCheck: boolean,
): { pairedBars: Array<Omit<CutBar, "segments" | "usedCm" | "wasteCm" | "estimatedKerfLossCm">>; remainingPieces: CutPiece[] } {
  const groups = new Map<number, CutPiece[]>();
  for (const piece of pieces) {
    const key = Number(piece.lengthCm.toFixed(2));
    const bucket = groups.get(key) ?? [];
    bucket.push(piece);
    groups.set(key, bucket);
  }

  const pairedBars: Array<Omit<CutBar, "segments" | "usedCm" | "wasteCm" | "estimatedKerfLossCm">> = [];
  const remainingPieces: CutPiece[] = [];

  for (const [lengthCm, groupPieces] of groups.entries()) {
    const canPerfectPair = Math.abs(lengthCm * 2 - stockLengthCm) < 0.0001
      && canFitPieceInBar([groupPieces[0]!], groupPieces[0]!, stockLengthCm, kerfCm, includeKerfInFitCheck);

    if (!canPerfectPair) {
      remainingPieces.push(...groupPieces);
      continue;
    }

    const pairCount = Math.floor(groupPieces.length / 2);
    for (let index = 0; index < pairCount; index += 1) {
      const firstPiece = groupPieces[index * 2]!;
      const secondPiece = groupPieces[index * 2 + 1]!;
      pairedBars.push({
        id: `bar-pair-${pairedBars.length + 1}`,
        type: firstPiece.type,
        stockLengthCm,
        pieces: [firstPiece, secondPiece],
      });
    }

    if (groupPieces.length % 2 === 1) {
      remainingPieces.push(groupPieces[groupPieces.length - 1]!);
    }
  }

  return { pairedBars, remainingPieces };
}

function packPiecesFfd(pieces: CutPiece[], stockLengthCm: number, kerfCm: number, includeKerfInFitCheck: boolean): CutBar[] {
  const directBars = pieces
    .filter((piece) => Math.abs(piece.lengthCm - stockLengthCm) < 0.0001)
    .map((piece, index) => finalizeCutBar({
      id: `bar-full-${index + 1}`,
      type: piece.type,
      stockLengthCm,
      pieces: [piece],
    }, kerfCm, includeKerfInFitCheck));

  const shortPieces = pieces.filter((piece) => piece.lengthCm < stockLengthCm - 0.0001);
  const perfectPairs = extractPerfectPairBars(shortPieces, stockLengthCm, kerfCm, includeKerfInFitCheck);
  const remainingPieces = perfectPairs.remainingPieces
    .sort((left, right) => right.lengthCm - left.lengthCm);

  const packedBars: Array<Omit<CutBar, "segments" | "usedCm" | "wasteCm" | "estimatedKerfLossCm">> = [];
  for (const piece of remainingPieces) {
    const targetBar = packedBars.find((bar) => canFitPieceInBar(bar.pieces, piece, stockLengthCm, kerfCm, includeKerfInFitCheck));
    if (targetBar) {
      targetBar.pieces.push(piece);
      continue;
    }
    packedBars.push({
      id: `bar-${packedBars.length + 1}`,
      type: piece.type,
      stockLengthCm,
      pieces: [piece],
    });
  }

  return [
    ...directBars,
    ...perfectPairs.pairedBars.map((bar) => finalizeCutBar(bar, kerfCm, includeKerfInFitCheck)),
    ...packedBars.map((bar) => finalizeCutBar(bar, kerfCm, includeKerfInFitCheck)),
  ];
}

function optimizeBarsWithOffcuts(bars: CutBar[], kerfCm: number, minReusableOffcutCm: number, includeKerfInFitCheck: boolean): CutBar[] {
  const mutableBars = bars.map((bar) => ({ ...bar, pieces: [...bar.pieces] }));
  let changed = true;

  while (changed) {
    changed = false;
    const sourceCandidates = [...mutableBars]
      .filter((bar) => bar.pieces.length > 0 && bar.wasteCm > 0)
      .sort((left, right) => left.usedCm - right.usedCm);

    for (const source of sourceCandidates) {
      const sourceIndex = mutableBars.findIndex((bar) => bar.id === source.id);
      if (sourceIndex < 0) continue;

      const targetIndices = mutableBars
        .map((bar, index) => ({ bar, index }))
        .filter(({ index, bar }) => index !== sourceIndex && bar.wasteCm >= minReusableOffcutCm);

      const simulated = mutableBars.map((bar) => ({ ...bar, pieces: [...bar.pieces] }));
      let canMoveAll = true;

      for (const piece of [...source.pieces].sort((left, right) => left.lengthCm - right.lengthCm)) {
        const target = targetIndices.find(({ index }) => canFitPieceInBar(simulated[index]?.pieces ?? [], piece, simulated[index]?.stockLengthCm ?? 0, kerfCm, includeKerfInFitCheck));
        if (!target) {
          canMoveAll = false;
          break;
        }
        simulated[target.index]!.pieces.push(piece);
      }

      if (!canMoveAll) continue;

      targetIndices.forEach(({ index }) => {
        mutableBars[index] = finalizeCutBar({
          id: simulated[index]!.id,
          type: simulated[index]!.type,
          stockLengthCm: simulated[index]!.stockLengthCm,
          pieces: simulated[index]!.pieces,
        }, kerfCm, includeKerfInFitCheck);
      });
      mutableBars.splice(sourceIndex, 1);
      changed = true;
      break;
    }
  }

  changed = true;
  while (changed) {
    changed = false;
    const candidates = mutableBars
      .map((bar, index) => ({ bar, index }))
      .sort((left, right) => right.bar.wasteCm - left.bar.wasteCm);

    for (let start = 0; start < candidates.length - 1; start += 1) {
      const subset = candidates.slice(start, start + 3);
      if (subset.length < 2) continue;

      const repacked = tryRepackBarSubset(
        subset.map(({ bar }) => bar.pieces),
        subset[0]!.bar.stockLengthCm,
        kerfCm,
        includeKerfInFitCheck,
      );
      if (!repacked || repacked.length >= subset.length) continue;

      subset
        .map(({ index }) => index)
        .sort((left, right) => right - left)
        .forEach((index) => { mutableBars.splice(index, 1); });

      repacked.forEach((pieces, index) => {
        mutableBars.push(finalizeCutBar({
          id: `bar-repack-${start + 1}-${index + 1}`,
          type: pieces[0]?.type ?? "mixed",
          stockLengthCm: subset[0]!.bar.stockLengthCm,
          pieces,
        }, kerfCm, includeKerfInFitCheck));
      });

      changed = true;
      break;
    }
  }

  return mutableBars.map((bar) => finalizeCutBar({
    id: bar.id,
    type: bar.type,
    stockLengthCm: bar.stockLengthCm,
    pieces: bar.pieces,
  }, kerfCm, includeKerfInFitCheck));
}

function tryRepackBarSubset(
  pieceGroups: CutPiece[][],
  stockLengthCm: number,
  kerfCm: number,
  includeKerfInFitCheck: boolean,
): CutPiece[][] | null {
  const pieces = pieceGroups.flat().sort((left, right) => right.lengthCm - left.lengthCm);
  const originalBarCount = pieceGroups.length;
  let best: CutPiece[][] | null = null;

  function search(pieceIndex: number, bars: CutPiece[][]): void {
    if (bars.length >= originalBarCount) return;
    if (best && bars.length >= best.length) return;
    if (pieceIndex >= pieces.length) {
      best = bars.map((bar) => [...bar]);
      return;
    }

    const piece = pieces[pieceIndex]!;
    const seenUsages = new Set<number>();
    for (let barIndex = 0; barIndex < bars.length; barIndex += 1) {
      const usage = getBarUsedCm(bars[barIndex]!, kerfCm, includeKerfInFitCheck);
      if (seenUsages.has(usage)) continue;
      if (!canFitPieceInBar(bars[barIndex]!, piece, stockLengthCm, kerfCm, includeKerfInFitCheck)) continue;
      seenUsages.add(usage);
      bars[barIndex]!.push(piece);
      search(pieceIndex + 1, bars);
      bars[barIndex]!.pop();
    }

    if (bars.length + 1 < originalBarCount) {
      bars.push([piece]);
      search(pieceIndex + 1, bars);
      bars.pop();
    }
  }

  search(0, []);
  return best;
}

function buildCutPlanStats(bars: CutBar[], stockLengthCm?: number): CutPlanStats {
  const totalBars = bars.length;
  const totalUsedCm = Number(bars.reduce((sum, bar) => sum + bar.pieces.reduce((pieceSum, piece) => pieceSum + piece.lengthCm, 0), 0).toFixed(2));
  const totalPurchasedCm = stockLengthCm == null
    ? bars.reduce((sum, bar) => sum + bar.stockLengthCm, 0)
    : totalBars * stockLengthCm;
  const totalWasteCm = Number((totalPurchasedCm - totalUsedCm).toFixed(2));
  const efficiencyPercent = totalPurchasedCm > 0 ? Number(((totalUsedCm / totalPurchasedCm) * 100).toFixed(2)) : 0;
  const totalPieceLength = Number(bars.reduce((sum, bar) => sum + bar.pieces.reduce((pieceSum, piece) => pieceSum + piece.lengthCm, 0), 0).toFixed(2));
  const estimatedKerfLossCm = Number(bars.reduce((sum, bar) => sum + bar.estimatedKerfLossCm, 0).toFixed(2));
  const lowerBoundBars = stockLengthCm && stockLengthCm > 0
    ? Math.ceil(totalPieceLength / stockLengthCm)
    : Math.ceil(totalPieceLength / Math.max(...bars.map((bar) => bar.stockLengthCm), 1));
  return { totalBars, totalUsedCm, totalWasteCm, efficiencyPercent, lowerBoundBars, estimatedKerfLossCm };
}

function resolveSharedCutStockLength(config: Required<CutOptimizationConfig>): number {
  return config.stockLengthCm;
}

function getProfileGroupStockLength(pieces: CutPiece[], config: Required<CutOptimizationConfig>): number {
  const firstType = pieces[0]?.type;
  if (!firstType) return config.stockLengthCm;
  if (firstType === "ud") return config.perTypeStockLengthsCm.ud ?? config.stockLengthCm;
  return config.perTypeStockLengthsCm.carrier ?? config.perTypeStockLengthsCm.mounting ?? config.stockLengthCm;
}

function optimizePieceGroup(pieces: CutPiece[], stockLengthCm: number, config: Required<CutOptimizationConfig>, idPrefix: string): CutBar[] {
  const invalidPiece = pieces.find((piece) => piece.lengthCm > stockLengthCm + 0.0001);
  if (invalidPiece) {
    throw new Error(`Парче ${invalidPiece.id} (${invalidPiece.lengthCm} cm) е по-дълго от наличния прът ${stockLengthCm} cm.`);
  }

  const packed = packPiecesFfd(pieces, stockLengthCm, config.kerfCm, config.includeKerfInFitCheck);
  const optimized = optimizeBarsWithOffcuts(packed, config.kerfCm, config.minReusableOffcutCm, config.includeKerfInFitCheck);
  return optimized.map((bar, index) => finalizeCutBar({
    id: `${idPrefix}-${index + 1}`,
    type: bar.type,
    stockLengthCm: bar.stockLengthCm,
    pieces: bar.pieces,
  }, config.kerfCm, config.includeKerfInFitCheck));
}

function buildCutPlanStatsForType(type: CutPieceType, bars: CutBar[], stockLengthCm: number): CutPlanStats {
  const relevantBars = bars.filter((bar) => bar.pieces.some((piece) => piece.type === type));
  const totalBars = relevantBars.length;
  const totalPieceLength = Number(relevantBars.reduce((sum, bar) => sum + bar.pieces
    .filter((piece) => piece.type === type)
    .reduce((pieceSum, piece) => pieceSum + piece.lengthCm, 0), 0).toFixed(2));
  const totalUsedCm = totalPieceLength;
  const capacityCm = relevantBars.reduce((sum, bar) => sum + bar.stockLengthCm, 0);
  const totalWasteCm = Number((capacityCm - totalUsedCm).toFixed(2));
  const efficiencyPercent = capacityCm > 0 ? Number(((totalUsedCm / capacityCm) * 100).toFixed(2)) : 0;
  const lowerBoundBars = stockLengthCm > 0 ? Math.ceil(totalPieceLength / stockLengthCm) : 0;
  const estimatedKerfLossCm = Number(relevantBars.reduce((sum, bar) => {
    const barPieceLength = bar.pieces.reduce((pieceSum, piece) => pieceSum + piece.lengthCm, 0);
    const relevantPieceLength = bar.pieces
      .filter((piece) => piece.type === type)
      .reduce((pieceSum, piece) => pieceSum + piece.lengthCm, 0);
    const share = barPieceLength > 0 ? relevantPieceLength / barPieceLength : 0;
    return sum + bar.estimatedKerfLossCm * share;
  }, 0).toFixed(2));
  return { totalBars, totalUsedCm, totalWasteCm, efficiencyPercent, lowerBoundBars, estimatedKerfLossCm };
}

export function optimizeSuspendedCeilingCuts(input: CutOptimizationInput, config: CutOptimizationConfig = {}): CutOptimizationResult {
  const resolvedConfig: Required<CutOptimizationConfig> = {
    stockLengthCm: config.stockLengthCm ?? 400,
    kerfCm: config.kerfCm ?? 0.3,
    minReusableOffcutCm: config.minReusableOffcutCm ?? 20,
    perTypeStockLengthsCm: config.perTypeStockLengthsCm ?? {},
    includeKerfInFitCheck: config.includeKerfInFitCheck ?? false,
    strictProfileSeparation: config.strictProfileSeparation ?? true,
  };

  const pieces = buildCutPieces(input);
  const stockLengthCm = resolveSharedCutStockLength(resolvedConfig);
  const bars = resolvedConfig.strictProfileSeparation
    ? [
      ...optimizePieceGroup(
        pieces.filter((piece) => piece.type === "carrier" || piece.type === "mounting"),
        getProfileGroupStockLength(pieces.filter((piece) => piece.type === "carrier" || piece.type === "mounting"), resolvedConfig),
        resolvedConfig,
        "cd-bar",
      ),
      ...optimizePieceGroup(
        pieces.filter((piece) => piece.type === "ud"),
        getProfileGroupStockLength(pieces.filter((piece) => piece.type === "ud"), resolvedConfig),
        resolvedConfig,
        "ud-bar",
      ),
    ]
    : optimizePieceGroup(pieces, stockLengthCm, resolvedConfig, "bar");
  const overallStats = buildCutPlanStats(bars);
  const carrierStats = buildCutPlanStatsForType("carrier", bars, resolvedConfig.perTypeStockLengthsCm.carrier ?? stockLengthCm);
  const mountingStats = buildCutPlanStatsForType("mounting", bars, resolvedConfig.perTypeStockLengthsCm.mounting ?? stockLengthCm);
  const udStats = buildCutPlanStatsForType("ud", bars, resolvedConfig.perTypeStockLengthsCm.ud ?? stockLengthCm);

  return {
    ...overallStats,
    bars,
    carrierStats,
    mountingStats,
    udStats,
  };
}

export function optimizeAllRoomsSuspendedCeilingCuts(
  rooms: Room[],
  constants: CalculatorConstants = DEFAULT_CONSTANTS,
  config: CutOptimizationConfig = {},
): CutOptimizationResult {
  constants = withDefaultConstants(constants);
  const resolvedConfig: Required<CutOptimizationConfig> = {
    stockLengthCm: config.stockLengthCm ?? Math.max(constants.cdLength, constants.udLength) * 100,
    kerfCm: config.kerfCm ?? 0.3,
    minReusableOffcutCm: config.minReusableOffcutCm ?? 20,
    perTypeStockLengthsCm: {
      carrier: constants.cdLength * 100,
      mounting: constants.cdLength * 100,
      ud: constants.udLength * 100,
      ...config.perTypeStockLengthsCm,
    },
    includeKerfInFitCheck: config.includeKerfInFitCheck ?? false,
    strictProfileSeparation: true,
  };

  const pieces = rooms.flatMap((room) => buildRoomCutPieces(room, calc(room, constants), constants));
  const cdPieces = pieces.filter((piece) => piece.type === "carrier" || piece.type === "mounting");
  const udPieces = pieces.filter((piece) => piece.type === "ud");
  const stockLengthCm = resolveSharedCutStockLength(resolvedConfig);
  const bars = [
    ...optimizePieceGroup(
      cdPieces,
      getProfileGroupStockLength(cdPieces, resolvedConfig),
      resolvedConfig,
      "global-cd-bar",
    ),
    ...optimizePieceGroup(
      udPieces,
      getProfileGroupStockLength(udPieces, resolvedConfig),
      resolvedConfig,
      "global-ud-bar",
    ),
  ];
  const overallStats = buildCutPlanStats(bars);
  const carrierStats = buildCutPlanStatsForType("carrier", bars, resolvedConfig.perTypeStockLengthsCm.carrier ?? stockLengthCm);
  const mountingStats = buildCutPlanStatsForType("mounting", bars, resolvedConfig.perTypeStockLengthsCm.mounting ?? stockLengthCm);
  const udStats = buildCutPlanStatsForType("ud", bars, resolvedConfig.perTypeStockLengthsCm.ud ?? stockLengthCm);

  return {
    ...overallStats,
    bars,
    carrierStats,
    mountingStats,
    udStats,
  };
}

export function calc(room: Room, constants: CalculatorConstants = DEFAULT_CONSTANTS): CalcResult {
  constants = withDefaultConstants(constants);
  normalizeRoom(room);
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = Number(constants.profileEdgeOffsetCm) || DEFAULT_CONSTANTS.profileEdgeOffsetCm;

  const actualCarrierRows = buildLinearPositions(W, room.c / 10, offset);
  const actualMountingRows = buildLinearPositions(L, room.b / 10, offset);
  const bearingCount = actualCarrierRows.length;
  const bearingLengthTotal = bearingCount * (L / 100);
  const bearingProfiles = bearingCount * Math.ceil((L / 100) / constants.cdLength);

  const mountingCount = actualMountingRows.length;
  const mountingLengthTotal = mountingCount * (W / 100);
  const mountingProfiles = mountingCount * Math.ceil((W / 100) / constants.cdLength);

  const cdTotalLength = bearingLengthTotal + mountingLengthTotal;
  const cdTotalProfiles = bearingProfiles + mountingProfiles;
  const crossConnectors = bearingCount * mountingCount;
  const hangersPerBearing = countBySpacing(L, room.a, offset);
  const hangersTotal = bearingCount * hangersPerBearing;
  const udTotalLength = (2 * (X + Y)) / 100;
  const udProfiles = 2 * Math.ceil((X / 100) / constants.udLength) + 2 * Math.ceil((Y / 100) / constants.udLength);
  const anchorsUd = countPerimeterAnchors(X, Y, room.udAnchorSpacing);
  const anchorsHangers = hangersTotal * constants.anchorsPerDirectHanger;
  const anchorsTotal = anchorsUd + anchorsHangers;
  const metalScrews = Math.ceil(
    (crossConnectors * constants.metalScrewsPerCrossConnector)
    + (hangersTotal * constants.metalScrewsPerDirectHanger),
  );
  const boardLayers = getEffectiveBoardLayers(room, constants);
  const drywallScrews = Math.ceil(Number(room.area) * boardLayers * constants.drywallScrewsPerM2);
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

function countProfiles(lengthM: number, profileLengthM: number): number {
  return Math.ceil(lengthM / profileLengthM);
}

function countLineProfiles(lineCount: number, lineLengthCm: number, profileLengthM: number): number {
  return lineCount * Math.ceil((lineLengthCm / 100) / profileLengthM);
}

function countPerimeterProfiles(widthCm: number, lengthCm: number, profileLengthM: number): number {
  return 2 * Math.ceil((widthCm / 100) / profileLengthM) + 2 * Math.ceil((lengthCm / 100) / profileLengthM);
}

function countSideAnchors(lengthCm: number, spacingMm: number): number {
  if (!lengthCm || !spacingMm) return 0;
  return Math.ceil((lengthCm / 100) / (spacingMm / 1000)) + 1;
}

function countPerimeterAnchors(widthCm: number, lengthCm: number, spacingMm: number): number {
  const anchors = 2 * countSideAnchors(widthCm, spacingMm) + 2 * countSideAnchors(lengthCm, spacingMm);
  return Math.max(0, anchors - 4);
}

function getProfileLength(profile: CustomProfileType, constants: CalculatorConstants): number {
  return Number(constants[CUSTOM_PROFILE_OPTIONS[profile].defaultLengthKey]) || constants.cdLength;
}

function getProfileLabel(profile: CustomProfileType): string {
  return CUSTOM_PROFILE_OPTIONS[profile].label;
}

function countExtensions(lineCount: number, lineLengthCm: number, profileLengthM: number): number {
  return lineCount * Math.max(0, Math.ceil((lineLengthCm / 100) / profileLengthM) - 1);
}

function formatTakeoffLength(value: number): number {
  return Number(value.toFixed(2));
}

function formatTakeoffWeight(value: number): number {
  return Number(value.toFixed(2));
}

function applyReserve(value: number, unit: MaterialTakeoffItem["unit"], constants: CalculatorConstants): number {
  const reserveMultiplier = 1 + (Number(constants.wastePercent) || 0) / 100;
  const next = value * reserveMultiplier;
  if (unit === "m" || unit === "m2") return formatTakeoffLength(next);
  if (unit === "kg") return formatTakeoffWeight(next);
  return Math.ceil(next);
}

function addMaterial(
  map: Map<string, MaterialTakeoffItem>,
  item: Omit<MaterialTakeoffItem, "quantity" | "optimizedQuantity" | "optimizedExplanation"> & { quantity: number },
): void {
  const existing = map.get(item.key);
  if (!existing) {
    map.set(item.key, item);
    return;
  }

  existing.quantity = item.unit === "m" || item.unit === "m2" || item.unit === "kg"
    ? formatTakeoffLength(existing.quantity + item.quantity)
    : existing.quantity + item.quantity;
}

function getMaterialRoomsForKey(key: string, rooms: Room[]): Room[] {
  const prefix = key.split("-")[0].toUpperCase();
  return prefix === "CUSTOM"
    ? rooms.filter((room) => room.systemType === "CUSTOM")
    : rooms.filter((room) => room.systemType === prefix);
}

function isOptimizedCdMaterial(item: MaterialTakeoffItem): boolean {
  return item.unit === "бр." && item.key.includes("cd-60-27") && !item.key.includes("connectors");
}

function isOptimizedUdMaterial(item: MaterialTakeoffItem): boolean {
  return item.unit === "бр." && item.key.includes("ud-28-27");
}

interface OptimizedProfileRoomData {
  room: Room;
  result: CalcResult;
  input?: CutOptimizationInput;
}

function getOptimizedProfileRoomData(sourceRoom: Room, constants: CalculatorConstants, cache: Map<string, OptimizedProfileRoomData>): OptimizedProfileRoomData {
  const existing = cache.get(sourceRoom.id);
  if (existing) return existing;
  const room = { ...sourceRoom, overrides: { ...sourceRoom.overrides } };
  const data = { room, result: calc(room, constants) };
  cache.set(room.id, data);
  return data;
}

function getOptimizedProfileInput(data: OptimizedProfileRoomData, constants: CalculatorConstants): CutOptimizationInput {
  if (!data.input) data.input = buildCutOptimizationInput(data.room, data.result, constants);
  return data.input;
}

function countOptimizedProfileBars(
  item: MaterialTakeoffItem,
  rooms: Room[],
  constants: CalculatorConstants,
  cache: Map<string, OptimizedProfileRoomData>,
): number | undefined {
  if (!isOptimizedCdMaterial(item) && !isOptimizedUdMaterial(item)) return undefined;
  const materialRooms = getMaterialRoomsForKey(item.key, rooms);
  if (!materialRooms.length) return undefined;

  let rowIndex = 0;
  const input: CutOptimizationInput = {
    carrierRows: [],
    mountingRows: [],
    udProfiles: { segments: [] },
  };

  materialRooms.forEach((sourceRoom) => {
    const data = getOptimizedProfileRoomData(sourceRoom, constants, cache);
    const roomInput = getOptimizedProfileInput(data, constants);
    if (isOptimizedCdMaterial(item)) {
      if (data.room.systemType !== "D116") {
        input.carrierRows.push(...roomInput.carrierRows.map((row) => ({ ...row, rowIndex: rowIndex++ })));
      }
      input.mountingRows.push(...roomInput.mountingRows.map((row) => ({ ...row, rowIndex: rowIndex++ })));
    } else {
      input.udProfiles.segments.push(...roomInput.udProfiles.segments);
    }
  });

  const plan = optimizeSuspendedCeilingCuts(input, {
    stockLengthCm: Math.max(constants.cdLength, constants.udLength) * 100,
    kerfCm: 0.3,
    minReusableOffcutCm: 20,
    perTypeStockLengthsCm: {
      carrier: constants.cdLength * 100,
      mounting: constants.cdLength * 100,
      ud: constants.udLength * 100,
    },
  });
  return isOptimizedCdMaterial(item)
    ? plan.bars.filter((bar) => bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "carrier" || piece.type === "mounting")).length
    : plan.bars.filter((bar) => bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "ud")).length;
}

function getOptimizedProfileExplanation(item: MaterialTakeoffItem): string | undefined {
  if (isOptimizedCdMaterial(item)) {
    return "Стандартната бройка е изчислена чрез закръгляне на общата дължина към цели CD профили. Бройката след разкрой отчита оптимизирано комбиниране на носещи и монтажни CD парчета, без смесване с UD профили.";
  }
  if (isOptimizedUdMaterial(item)) {
    return "Стандартната бройка е изчислена чрез закръгляне на общия периметър към цели UD профили. Бройката след разкрой отчита оптимизирано използване на остатъците от UD профили по периферията.";
  }
  return undefined;
}

export function buildMaterialTakeoff(
  rooms: Room[],
  constants: CalculatorConstants = DEFAULT_CONSTANTS,
  options: { includeOptimizedQuantities?: boolean } = {},
): MaterialTakeoffItem[] {
  constants = withDefaultConstants(constants);
  const includeOptimizedQuantities = options.includeOptimizedQuantities ?? true;
  const map = new Map<string, MaterialTakeoffItem>();
  const systems = new Set<SystemType>();
  const optimizedProfileCache = new Map<string, OptimizedProfileRoomData>();

  rooms.forEach((sourceRoom) => {
    const room = { ...sourceRoom, overrides: { ...sourceRoom.overrides } };
    const result = calc(room, constants);
    optimizedProfileCache.set(room.id, { room, result });
    systems.add(room.systemType);
    const system = room.systemType.toLowerCase();
    const udRule = getUdAnchoringRule(room);
    const hanger = HANGER_OPTIONS[getHangerType(room)];
    const boardLayers = getEffectiveBoardLayers(room, constants);
    const boardLabel = getBoardMaterialLabel(room.boardType);
    const bearingExtensions = countExtensions(result.bearingCount, result.L, constants.cdLength);
    const uaExtensions = countExtensions(result.bearingCount, result.L, constants.uaLength);
    const mountingExtensions = countExtensions(result.mountingCount, result.W, constants.cdLength);

    const add = (item: Omit<MaterialTakeoffItem, "quantity" | "optimizedQuantity" | "optimizedExplanation"> & { quantity: number }) => addMaterial(map, item);

    if (room.systemType === "D111") {
      add({
        key: "d111-bearing-battens",
        label: "D111 носещи дървени летви",
        quantity: countLineProfiles(result.bearingCount, result.L, constants.woodBattenLength),
        unit: "бр.",
        note: `${formatTakeoffLength(result.bearingLengthTotal)} m; геометрично по носещите редове`,
        source: "estimate",
      });
      add({
        key: "d111-mounting-battens",
        label: "D111 монтажни дървени летви",
        quantity: countLineProfiles(result.mountingCount, result.W, constants.woodBattenLength),
        unit: "бр.",
        note: `${formatTakeoffLength(result.mountingLengthTotal)} m; геометрично по монтажните редове`,
        source: "estimate",
      });
      add({
        key: "d111-wood-cross-points",
        label: "D111 точки на кръстосване/закрепване",
        quantity: result.crossConnectors,
        unit: "бр.",
        note: "брой пресичания; типът крепеж зависи от избрания D111 вариант",
        source: "estimate",
      });
      add({
        key: "d111-perimeter-anchors",
        label: "D111 периферни дюбели",
        quantity: result.anchorsUd,
        unit: "бр.",
        note: udRule.note,
        source: "geometry",
      });
    } else if (room.systemType === "D116") {
      add({
        key: "d116-ua-50-40",
        label: "D116 UA 50/40 носещи профили",
        quantity: countLineProfiles(result.bearingCount, result.L, constants.uaLength),
        unit: "бр.",
        note: `${formatTakeoffLength(result.bearingLengthTotal)} m носещи UA`,
        source: "geometry",
      });
      add({
        key: "d116-cd-60-27",
        label: "D116 CD 60/27 монтажни профили",
        quantity: result.mountingProfiles,
        unit: "бр.",
        note: `${formatTakeoffLength(result.mountingLengthTotal)} m монтажни CD`,
        source: "geometry",
      });
      add({
        key: "d116-uw-50-40",
        label: "D116 UW 50/40 за UA удължаване",
        quantity: uaExtensions * 2,
        unit: "бр.",
        note: "оценка: 2 UW елемента за всяко удължаване на UA профил",
        source: "estimate",
      });
      add({
        key: "d116-ua-cd-connectors",
        label: "D116 кръстати връзки UA/CD",
        quantity: result.crossConnectors,
        unit: "бр.",
        note: "пресичания между UA носещи и CD монтажни профили",
        source: "geometry",
      });
      add({
        key: `d116-hangers-${hanger.value}`,
        label: `D116 ${hanger.label}`,
        quantity: result.hangersTotal,
        unit: "бр.",
        note: hanger.useWhen,
        source: "estimate",
      });
    } else if (room.systemType === "CUSTOM") {
      const bearingProfile = normalizeCustomProfile(room.customBearingProfile, DEFAULT_CUSTOM_BEARING_PROFILE);
      const mountingProfile = normalizeCustomProfile(room.customMountingProfile, DEFAULT_CUSTOM_MOUNTING_PROFILE);
      const perimeterProfile = normalizeCustomProfile(room.customPerimeterProfile, DEFAULT_CUSTOM_PERIMETER_PROFILE);
      const bearingProfileLength = getProfileLength(bearingProfile, constants);
      const mountingProfileLength = getProfileLength(mountingProfile, constants);
      const bearingProfileLabel = getProfileLabel(bearingProfile);
      const mountingProfileLabel = getProfileLabel(mountingProfile);

      add({
        key: `custom-bearing-${bearingProfile}`,
        label: `Custom носещи ${bearingProfileLabel}`,
        quantity: countLineProfiles(result.bearingCount, result.L, bearingProfileLength),
        unit: "бр.",
        note: `${formatTakeoffLength(result.bearingLengthTotal)} m носещи редове; дължина ${bearingProfileLength} m`,
        source: "manual",
      });
      add({
        key: `custom-mounting-${mountingProfile}`,
        label: `Custom монтажни ${mountingProfileLabel}`,
        quantity: countLineProfiles(result.mountingCount, result.W, mountingProfileLength),
        unit: "бр.",
        note: `${formatTakeoffLength(result.mountingLengthTotal)} m монтажни редове; дължина ${mountingProfileLength} m`,
        source: "manual",
      });
      add({
        key: "custom-connectors",
        label: "Custom връзки",
        quantity: result.crossConnectors,
        unit: "бр.",
        note: `пресичания ${bearingProfileLabel} x ${mountingProfileLabel}`,
        source: "manual",
      });
      add({
        key: `custom-hangers-${hanger.value}`,
        label: `Custom ${hanger.label}`,
        quantity: result.hangersTotal,
        unit: "бр.",
        note: `${result.hangersPerBearing} бр. на носещ ред; ${hanger.useWhen}`,
        source: "manual",
      });
      add({
        key: `custom-perimeter-${perimeterProfile}`,
        label: `Custom периферни ${getProfileLabel(perimeterProfile)}`,
        quantity: countPerimeterProfiles(room.width, room.length, getProfileLength(perimeterProfile, constants)),
        unit: "бр.",
        note: `${formatTakeoffLength(result.udTotalLength)} m по периметъра; ${udRule.mode}`,
        source: "manual",
      });
    } else {
      const connectorLabel = room.systemType === "D113"
        ? "D113 връзки на едно ниво за CD"
        : room.systemType === "D112"
          ? "D112 кръстати връзки за CD"
          : "Custom връзки";
      add({
        key: `${system}-cd-60-27`,
        label: `${room.systemType} CD 60/27 профили`,
        quantity: result.cdTotalProfiles,
        unit: "бр.",
        note: `${formatTakeoffLength(result.cdTotalLength)} m носещи + монтажни профили`,
        source: "geometry",
      });
      add({
        key: `${system}-connectors`,
        label: connectorLabel,
        quantity: result.crossConnectors,
        unit: "бр.",
        note: "брой пресичания носещи x монтажни редове",
        source: "geometry",
      });
      add({
        key: `${system}-hangers-${hanger.value}`,
        label: `${room.systemType} ${hanger.label}`,
        quantity: result.hangersTotal,
        unit: "бр.",
        note: `${result.hangersPerBearing} бр. на носещ ред; ${hanger.useWhen}`,
        source: "geometry",
      });
    }

    if (room.systemType !== "D111" && room.systemType !== "CUSTOM") {
      add({
        key: `${system}-ud-28-27`,
        label: `${room.systemType} UD 28/27 профили`,
        quantity: result.udProfiles,
        unit: "бр.",
        note: `${formatTakeoffLength(result.udTotalLength)} m по периметъра; ${udRule.mode}`,
        source: "geometry",
      });
    }

    add({
      key: `${system}-anchors-ud`,
      label: `${room.systemType} дюбели за периферия`,
      quantity: result.anchorsUd,
      unit: "бр.",
      note: `стъпка ${room.udAnchorSpacing} mm; ${udRule.note}`,
      source: room.overrides.udAnchorSpacing ? "manual" : "geometry",
    });
    add({
      key: `${system}-anchors-hangers`,
      label: `${room.systemType} дюбели за окачвачи`,
      quantity: result.anchorsHangers,
      unit: "бр.",
      note: `геометрично по броя окачвачи; избран тип: ${hanger.label}`,
      source: "geometry",
    });
    add({
      key: `${system}-metal-screws`,
      label: `${room.systemType} LN/LB винтове за метал/връзки`,
      quantity: result.metalScrews,
      unit: "бр.",
      note: "по глобалните настройки за винтове/връзка и винтове/окачвач",
      source: "estimate",
    });
    add({
      key: `${system}-tn-screws`,
      label: `${room.systemType} TN винтове за гипсокартон`,
      quantity: result.drywallScrews,
      unit: "бр.",
      note: `${constants.drywallScrewsPerM2} бр./m2 x ${boardLayers} слой/слоя`,
      source: "estimate",
    });
    const boardArea = Number(room.area) * boardLayers;
    const boardSize = Math.max(0.01, constants.boardWidth * constants.boardLength);
    add({
      key: `${system}-boards`,
      label: `${room.systemType} ${boardLabel} плоскости`,
      quantity: Math.ceil(boardArea / boardSize),
      unit: "бр.",
      note: `${formatTakeoffLength(boardArea)} m2 обшивка; лист ${constants.boardWidth} x ${constants.boardLength} m`,
      source: "estimate",
    });
    add({
      key: `${system}-boards-area`,
      label: `${room.systemType} площ гипсокартон`,
      quantity: boardArea,
      unit: "m2",
      note: `${formatTakeoffLength(Number(room.area))} m2 x ${boardLayers} слой/слоя`,
      source: "estimate",
    });
    add({
      key: `${system}-joint-tape`,
      label: `${room.systemType} фугопокривна лента`,
      quantity: Number(room.area) * constants.jointTapePerM2,
      unit: "m",
      note: `${constants.jointTapePerM2} m/m2 ориентировъчно`,
      source: "estimate",
    });
    add({
      key: `${system}-joint-compound`,
      label: `${room.systemType} шпакловка за фуги`,
      quantity: Number(room.area) * constants.jointCompoundKgPerM2,
      unit: "kg",
      note: `${constants.jointCompoundKgPerM2} kg/m2 ориентировъчно`,
      source: "estimate",
    });
    add({
      key: `${system}-trenn-fix`,
      label: `${room.systemType} Trenn-Fix / разделителна лента`,
      quantity: result.udTotalLength * constants.trennFixPerimeterMultiplier,
      unit: "m",
      note: "по периметъра на помещението",
      source: "estimate",
    });
    if (constants.mineralWoolEnabled || room.fireProtection) {
      add({
        key: `${system}-mineral-wool`,
        label: `${room.systemType} минерална вата`,
        quantity: Number(room.area),
        unit: "m2",
        note: `${constants.mineralWoolThickness} mm; провери изискването по EI/акустичния детайл`,
        source: "estimate",
      });
    }
    add({
      key: `${system}-extensions`,
      label: `${room.systemType} надлъжни удължители`,
      quantity: room.systemType === "D116"
        ? mountingExtensions
        : room.systemType === "CUSTOM"
          ? countExtensions(result.bearingCount, result.L, getProfileLength(normalizeCustomProfile(room.customBearingProfile, DEFAULT_CUSTOM_BEARING_PROFILE), constants))
            + countExtensions(result.mountingCount, result.W, getProfileLength(normalizeCustomProfile(room.customMountingProfile, DEFAULT_CUSTOM_MOUNTING_PROFILE), constants))
          : bearingExtensions + mountingExtensions,
      unit: "бр.",
      note: room.systemType === "D116" ? "за CD профили; UA удължаването е отделено като UW оценка" : "геометрична оценка според дължината на профила",
      source: "estimate",
    });
  });

  const singleSystemPrefix = systems.size === 1 ? `${Array.from(systems)[0]} ` : "";

  return Array.from(map.values()).map((item) => {
    const optimizedQuantity = includeOptimizedQuantities
      ? countOptimizedProfileBars(item, rooms, constants, optimizedProfileCache)
      : undefined;
    return {
      ...item,
      label: singleSystemPrefix && item.label.startsWith(singleSystemPrefix)
        ? item.label.slice(singleSystemPrefix.length)
        : item.label,
      quantity: applyReserve(item.quantity, item.unit, constants),
      optimizedQuantity,
      optimizedExplanation: optimizedQuantity == null ? undefined : getOptimizedProfileExplanation(item),
    };
  });
}
