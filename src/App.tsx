import { ChangeEvent, ReactNode, useState } from "react";
import {
  CalcResult,
  CalculatorConstants,
  CONSTRUCTION_TYPES,
  CUSTOM_PROFILE_OPTIONS,
  CustomProfileType,
  D112_VARIANTS,
  D112Variant,
  D116_VARIANTS,
  D116Variant,
  DEFAULT_CONSTANTS,
  FIRE_RATING_OPTIONS,
  FireRating,
  HANGER_OPTIONS,
  HangerType,
  LoadInputMode,
  LoadClass,
  Room,
  SystemType,
  applyAutoABC,
  buildMaterialTakeoff,
  buildPositions,
  calc,
  createRoom,
  getConstruction,
  getAutoABC,
  getAllowedAValues,
  getAllowedBValues,
  getAutomaticLoadClass,
  getBoardOptions,
  getDefaultHangerType,
  getEffectiveBoardLayers,
  getFireCertificationCheck,
  getHangerOptions,
  getLoadClasses,
  getTableValue,
  getUdAnchoringRule,
  getValidCValues,
  getValidationWarnings,
  estimateLoadKgPerM2,
  syncSpacingFromKnaufTable,
  validateCombination,
} from "./domain/calculator";

const STORAGE_KEY = "d113-calculator-v2";
const BGN_PER_EUR = 1.95583;

interface AppState {
  rooms: Room[];
  draftRoom: Room;
  constants: CalculatorConstants;
  materialPrices: MaterialPrices;
  activeRoomId: string;
}

type MaterialPrices = Record<string, number>;

interface CatalogPriceEstimate {
  price: number;
  source: string;
}

interface CatalogPriceSource {
  url: string;
  source: string;
}

interface KnaufCatalogProduct {
  id: string;
  knaufName: string;
  matchTerms: string[];
  estimate?: CatalogPriceEstimate;
  priceSources: CatalogPriceSource[];
}

const KNAUF_PRODUCT_CATALOG: KnaufCatalogProduct[] = [
  {
    id: "cd_60_27",
    knaufName: "Профил Кнауф CD 60/27",
    matchTerms: ["cd 60/27", "cd-60-27", "cd_60_27"],
    estimate: { price: 5.92, source: "Baustoff + Metall, CD 60/27 4 m" },
    priceSources: [{ url: "https://shop.baustoff-metall.bg/product/1435/profil-knauf-cd-60-27-0-6-mm-1-broy-4-m.html", source: "Baustoff + Metall, CD 60/27 4 m" }],
  },
  {
    id: "ud_28_27",
    knaufName: "Профил Кнауф UD 28/27",
    matchTerms: ["ud 28/27", "ud-28-27", "ud_28_27", "периферни ud"],
    estimate: { price: 3.00, source: "Baustoff + Metall, UD 28/27 3 m" },
    priceSources: [{ url: "https://shop.baustoff-metall.bg/product/1438/profil-knauf-ud-28-27-0-6-mm-1-broy-3-m.html", source: "Baustoff + Metall, UD 28/27 3 m" }],
  },
  {
    id: "ua_50_40",
    knaufName: "UA профил Кнауф 50/40",
    matchTerms: ["ua 50/40", "ua-50-40", "ua_50_40"],
    estimate: { price: 26.07, source: "Praktiker, UA 50/40 3 m" },
    priceSources: [{ url: "https://praktiker.bg/bg/Profili-za-suho-stroitelstvo/UA-PROFIL-KNAUF-50-40/p/116447", source: "Praktiker, UA 50/40 3 m" }],
  },
  {
    id: "direct_hanger_cd",
    knaufName: "Директен окачвач за CD 60/27",
    matchTerms: ["директен окачвач"],
    estimate: { price: 0.33, source: "DS Home, директен окачвач 125 mm" },
    priceSources: [{ url: "https://www.dshome.bg/drugi-krepezhi/direkten-okachvach-za-cd-6027-125-mm-knauf", source: "DS Home, директен окачвач 125 mm" }],
  },
  {
    id: "acoustic_hanger_cd",
    knaufName: "Акустичен регулируем директен окачвач за CD 60/27",
    matchTerms: ["акустичен окачвач"],
    estimate: { price: 4.23, source: "Polaris, акустичен окачвач" },
    priceSources: [{ url: "https://polaris-bg.com/bg/akustichen-direkten-okachvach-knauf-60-27/p382.html", source: "Polaris, акустичен окачвач" }],
  },
  {
    id: "single_level_connector_cd",
    knaufName: "Кръстата връзка за едно ниво за CD 60/27",
    matchTerms: ["връзки на едно ниво"],
    estimate: { price: 1.84, source: "Polaris, връзка за едно ниво CD 60/27" },
    priceSources: [{ url: "https://polaris-bg.com/bg/krystata-vryzka-cd-60-27-edno-nivo-knauf/p301.html", source: "Polaris, връзка едно ниво CD" }],
  },
  {
    id: "cross_connector_cd",
    knaufName: "Кръстата връзка за CD 60/27",
    matchTerms: ["кръстати връзки", "връзки ua/cd", "custom връзки"],
    estimate: { price: 0.34, source: "DS Home, кръстата връзка CD 60/27" },
    priceSources: [{ url: "https://www.dshome.bg/drugi-krepezhi/krstata-vrzka-za-cd-6027-knauf", source: "DS Home, кръстата връзка CD" }],
  },
  {
    id: "board_a13_12_5",
    knaufName: "Стандартна гипскартонена плоскост A13 12.5 / 1200 / 2000 Knauf",
    matchTerms: ["knauf a 12.5 mm плоскости", "knauf 2x12.5 mm плоскости"],
    estimate: { price: 8.03, source: "Maxxmart, Knauf A13 12.5 1200/2000" },
    priceSources: [{ url: "https://www.maxxmart.eu/product/standartna-gipskartonena-ploskost-a13-125-1200-2000-knauf", source: "Maxxmart, Knauf A13 12.5" }],
  },
  {
    id: "board_h13_12_5",
    knaufName: "Гипсокартон влагоустойчив KNAUF H13 12.5 / 1200 / 2000",
    matchTerms: ["knauf h2 12.5 mm плоскости"],
    estimate: { price: 10.08, source: "SKK, Knauf H13 12.5 1200/2000" },
    priceSources: [{ url: "https://www.skk.bg/product/gipsokarton-zelen-vlagoustoychiv-knauf-h13-12-5-1200-2000", source: "SKK, Knauf H13 12.5" }],
  },
  {
    id: "board_diamant_12_5",
    knaufName: "Гипсова плоскост KNAUF Diamant 12.5 / 1200 / 2000",
    matchTerms: ["knauf diamant 12.5 mm плоскости"],
    estimate: { price: 20.47, source: "SKK, Knauf Diamant 12.5 1200/2000" },
    priceSources: [{ url: "https://skk.bg/en/product/gipsova-ploskost-knauf-diamant-gkfi-12-5-12-5-1200-2000", source: "SKK, Knauf Diamant 12.5" }],
  },
  {
    id: "joint_tape",
    knaufName: "Фугопокривна лента Knauf",
    matchTerms: ["фугопокривна лента"],
    estimate: { price: 1.75, source: "Строймаркет, стъклофазерна лента 25 m" },
    priceSources: [{ url: "https://stroymarket.bg/produkt/staklofazerna-lenta-knauf/", source: "Строймаркет, стъклофазерна лента 25 m" }],
  },
];

function loadState(): AppState {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") as Partial<AppState>;
    const rooms = Array.isArray(raw.rooms) ? raw.rooms : [];
    const draftRoom = raw.draftRoom
      ? { ...createRoom(), ...raw.draftRoom, overrides: { ...createRoom().overrides, ...raw.draftRoom.overrides } }
      : rooms[0]
        ? cloneRoom(rooms[0])
        : createRoom();
    return {
      rooms,
      draftRoom,
      constants: { ...DEFAULT_CONSTANTS, ...raw.constants },
      materialPrices: raw.materialPrices ?? {},
      activeRoomId: draftRoom.id,
    };
  } catch {
    const room = createRoom();
    return { rooms: [], draftRoom: room, constants: DEFAULT_CONSTANTS, materialPrices: {}, activeRoomId: room.id };
  }
}

function saveState(next: AppState): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}

function cloneRoom(room: Room): Room {
  return { ...room, overrides: { ...room.overrides } };
}

function formatNumber(value: number, digits = 2): string {
  return Number(value).toFixed(digits);
}

function formatCurrency(value: number): string {
  return `${formatNumber(value)} €`;
}

function getMaterialUnitPrice(prices: MaterialPrices, key: string): number {
  return Math.max(0, Number(prices[key]) || 0);
}

function getCatalogProductForMaterial(label: string, key: string): KnaufCatalogProduct | null {
  const haystack = normalizeCatalogText(`${key} ${label}`);
  return KNAUF_PRODUCT_CATALOG.find((product) => product.matchTerms.some((term) => haystack.includes(normalizeCatalogText(term)))) ?? null;
}

function getEstimatedMaterialPrice(label: string, key: string): CatalogPriceEstimate | null {
  return getCatalogProductForMaterial(label, key)?.estimate ?? null;
}

function normalizeCatalogText(value: string): string {
  return value.toLowerCase().replaceAll("_", "-").replace(/\s+/g, " ").trim();
}

function parseOnlinePriceToEur(html: string): number | null {
  const compact = html.replace(/\s+/g, " ");
  const eurMatch = compact.match(/(?:€|EUR)\s*([0-9]+(?:[,.][0-9]{1,2})?)|([0-9]+(?:[,.][0-9]{1,2})?)\s*(?:€|EUR)/i);
  if (eurMatch) return normalizePrice(eurMatch[1] ?? eurMatch[2]);

  const bgnMatch = compact.match(/([0-9]+(?:[,.][0-9]{1,2})?)\s*(?:лв\.?|лева|BGN)/i);
  const bgn = bgnMatch ? normalizePrice(bgnMatch[1]) : null;
  return bgn == null ? null : Number((bgn / BGN_PER_EUR).toFixed(2));
}

function normalizePrice(value: string): number | null {
  const price = Number(value.replace(",", "."));
  return Number.isFinite(price) && price > 0 ? Number(price.toFixed(2)) : null;
}

async function fetchOnlinePrice(source: CatalogPriceSource): Promise<number | null> {
  try {
    const response = await fetch(source.url, { mode: "cors" });
    if (!response.ok) return null;
    return parseOnlinePriceToEur(await response.text());
  } catch {
    return null;
  }
}

const CUSTOM_A_OPTIONS = buildRange(100, 3000, 50);
const CUSTOM_B_OPTIONS = buildRange(100, 1250, 25);
const CUSTOM_C_OPTIONS = buildRange(100, 3000, 50);
const CUSTOM_OFFSET_OPTIONS = buildRange(0, 100, 5);
const CUSTOM_UD_ANCHOR_OPTIONS = buildRange(100, 1500, 25);

function buildRange(min: number, max: number, step: number): number[] {
  const values: number[] = [];
  for (let value = min; value <= max; value += step) values.push(value);
  return values;
}

function buildCalculationFormulas(room: Room, result: CalcResult, constants: CalculatorConstants) {
  const X = Number(room.width);
  const Y = Number(room.length);
  const offset = Number(room.offset);
  const boardLayers = getEffectiveBoardLayers(room, constants);
  return [
    {
      key: "dimensions",
      label: "Размери",
      formula: "W = min(X, Y), L = max(X, Y)",
      steps: [
        `X = ${X} cm, Y = ${Y} cm.`,
        `По-късата страна се приема за W: ${result.W} cm.`,
        `По-дългата страна се приема за L: ${result.L} cm.`,
      ],
    },
    {
      key: "bearingCount",
      label: "Носещи редове",
      formula: "ceil((W - 2 x offset) / (c / 10)) + 1",
      steps: [
        `c = ${room.c} mm = ${room.c / 10} cm.`,
        `ceil((${result.W} - 2 x ${offset}) / ${room.c / 10}) + 1 = ${result.bearingCount} бр.`,
      ],
    },
    {
      key: "mountingCount",
      label: "Монтажни редове",
      formula: "ceil((L - 2 x offset) / (b / 10)) + 1",
      steps: [
        `b = ${room.b} mm = ${room.b / 10} cm.`,
        `ceil((${result.L} - 2 x ${offset}) / ${room.b / 10}) + 1 = ${result.mountingCount} бр.`,
      ],
    },
    {
      key: "profileLengths",
      label: "Метри профили",
      formula: "носещи m = носещи редове x L / 100; монтажни m = монтажни редове x W / 100",
      steps: [
        `${result.bearingCount} x ${result.L} / 100 = ${formatNumber(result.bearingLengthTotal)} m носещи.`,
        `${result.mountingCount} x ${result.W} / 100 = ${formatNumber(result.mountingLengthTotal)} m монтажни.`,
        `Общо = ${formatNumber(result.cdTotalLength)} m.`,
      ],
    },
    {
      key: "profilePieces",
      label: "Профили общо",
      formula: "носещи редове x ceil(L / дължина профил) + монтажни редове x ceil(W / дължина профил)",
      steps: [
        `Дължина профил = ${constants.cdLength} m.`,
        `${result.bearingCount} x ceil(${formatNumber(result.L / 100)} / ${constants.cdLength}) + ${result.mountingCount} x ceil(${formatNumber(result.W / 100)} / ${constants.cdLength}) = ${result.cdTotalProfiles} бр.`,
      ],
    },
    {
      key: "connectors",
      label: "Връзки",
      formula: "носещи редове x монтажни редове",
      steps: [`${result.bearingCount} x ${result.mountingCount} = ${result.crossConnectors} бр.`],
    },
    {
      key: "hangers",
      label: "Окачвачи",
      formula: "носещи редове x (ceil((L - 2 x offset) / (a / 10)) + 1)",
      steps: [
        `a = ${room.a} mm = ${room.a / 10} cm.`,
        `Окачвачи на носещ ред = ceil((${result.L} - 2 x ${offset}) / ${room.a / 10}) + 1 = ${result.hangersPerBearing} бр.`,
        `Общо = ${result.bearingCount} x ${result.hangersPerBearing} = ${result.hangersTotal} бр.`,
      ],
    },
    {
      key: "ud",
      label: "UD профили и дюбели",
      formula: "UD m = 2 x (X + Y) / 100; UD бр. = 2 x ceil(X / UD) + 2 x ceil(Y / UD); дюбели UD = по страни с недублирани ъгли",
      steps: [
        `UD дължина = 2 x (${X} + ${Y}) / 100 = ${formatNumber(result.udTotalLength)} m.`,
        `UD профили = 2 x ceil(${formatNumber(X / 100)} / ${constants.udLength}) + 2 x ceil(${formatNumber(Y / 100)} / ${constants.udLength}) = ${result.udProfiles} бр.`,
        `Дюбели UD = 2 x (ceil(${formatNumber(X / 100)} / ${room.udAnchorSpacing / 1000}) + 1) + 2 x (ceil(${formatNumber(Y / 100)} / ${room.udAnchorSpacing / 1000}) + 1) - 4 = ${result.anchorsUd} бр.`,
      ],
    },
    {
      key: "screws",
      label: "Винтове",
      formula: "метал = връзки x винтове/връзка + окачвачи x винтове/окачвач; гипсокартон = площ x слоеве x винтове/m2",
      steps: [
        `Метал = ${result.crossConnectors} x ${constants.metalScrewsPerCrossConnector} + ${result.hangersTotal} x ${constants.metalScrewsPerDirectHanger} = ${result.metalScrews} бр.`,
        `Гипсокартон = ceil(${formatNumber(room.area)} x ${boardLayers} x ${constants.drywallScrewsPerM2}) = ${result.drywallScrews} бр.`,
      ],
    },
    {
      key: "extensions",
      label: "Удължители",
      formula: "редове x max(0, ceil(дължина ред / дължина профил) - 1)",
      steps: [`Общо удължители = ${result.extensionsTotal} бр.`],
    },
  ];
}

function syncDistanceOverride(room: Room, key: "a" | "b" | "c" | "offset" | "udAnchorSpacing"): void {
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType, room.d116Variant, room.d112Variant);
  const expected = {
    a: getTableValue(room) ?? auto.a,
    b: room.boardType === "custom" ? Number.NaN : auto.b,
    c: auto.c,
    offset: auto.offset,
    udAnchorSpacing: auto.udAnchorSpacing,
  }[key];

  room.overrides[key] = Number(room[key]) !== Number(expected);
}

function App() {
  const [state, setState] = useState<AppState>(() => loadState());
  const [zoom, setZoom] = useState(1);
  const [saveStatus, setSaveStatus] = useState("");
  const [priceSearchStatus, setPriceSearchStatus] = useState("");

  const activeRoom = state.draftRoom;
  const activeResult = calc(activeRoom, state.constants);
  const activeWarnings = getValidationWarnings(cloneRoom(activeRoom));
  const isValid = !activeWarnings.some((warning) => warning.severity === "error");

  function commit(updater: (draft: AppState) => void): void {
    setState((current) => {
      const next = {
        rooms: current.rooms.map(cloneRoom),
        draftRoom: cloneRoom(current.draftRoom),
        constants: { ...current.constants },
        materialPrices: { ...current.materialPrices },
        activeRoomId: current.activeRoomId,
      };
      updater(next);
      saveState(next);
      return next;
    });
  }

  function saveCurrentState(): void {
    commit((draft) => {
      const room = cloneRoom(draft.draftRoom);
      const index = draft.rooms.findIndex((item) => item.id === room.id);
      if (index >= 0) {
        draft.rooms[index] = room;
      } else {
        draft.rooms.push(room);
      }
      draft.activeRoomId = room.id;
      draft.draftRoom = cloneRoom(room);
    });
    setSaveStatus("Запазено");
    window.setTimeout(() => setSaveStatus(""), 1600);
  }

  function updateActiveRoom(updater: (room: Room) => void): void {
    commit((draft) => {
      const room = draft.draftRoom;
      updater(room);
      applyAutoABC(room);
      if (!room.overrides.area && room.width && room.length) {
        room.area = (room.width * room.length) / 10000;
      }
    });
  }

  function addRoom(): void {
    commit((draft) => {
      const room = createRoom(`Стая ${draft.rooms.length + 1}`);
      draft.draftRoom = room;
      draft.activeRoomId = room.id;
    });
  }

  function deleteRoom(roomId: string): void {
    commit((draft) => {
      draft.rooms = draft.rooms.filter((room) => room.id !== roomId);
      if (draft.activeRoomId === roomId) {
        draft.activeRoomId = draft.draftRoom.id;
      }
    });
  }

  function deleteAllRooms(): void {
    commit((draft) => {
      draft.rooms = [];
      draft.activeRoomId = draft.draftRoom.id;
    });
  }

  function changeSystem(systemType: SystemType): void {
    updateActiveRoom((room) => {
      const construction = getConstruction(systemType);
      room.systemType = systemType;
      if (systemType === "D112") room.d112Variant = "double_cd";
      if (systemType === "D116") room.d116Variant = "ua_cd";
      room.hangerType = getDefaultHangerType(systemType);
      if (systemType !== "CUSTOM" && !room.boardType.startsWith("knauf_")) {
        room.boardType = "knauf_a_12.5";
      }
      room.fireRating = construction.defaultFireProtection ? "fire_table" : "none";
      room.fireProtection = room.fireRating !== "none";
      room.loadClass = getLoadClasses(systemType, room.fireProtection, room.d116Variant, room.d112Variant)[0] ?? construction.defaultLoadClass;
      room.loadInputMode = "manual";
      room.overrides.a = false;
      room.overrides.b = false;
      room.overrides.c = false;
      room.overrides.offset = false;
      room.overrides.udAnchorSpacing = false;
      syncSpacingFromKnaufTable(room, { keepC: false });
      if (systemType === "CUSTOM") {
        room.customBearingProfile = "cd_60_27";
        room.customMountingProfile = "cd_60_27";
        room.customPerimeterProfile = "ud_28_27";
        room.fireRating = "none";
        room.fireProtection = false;
        room.overrides.a = true;
        room.overrides.b = true;
        room.overrides.c = true;
        room.overrides.offset = true;
        room.overrides.udAnchorSpacing = true;
      }
    });
  }

  function resetAutoSpacing(): void {
    updateActiveRoom((room) => {
      room.overrides.a = false;
      room.overrides.b = false;
      room.overrides.c = false;
      room.overrides.offset = false;
      room.overrides.udAnchorSpacing = false;
      syncSpacingFromKnaufTable(room, { keepC: false });
    });
  }

  function exportJson(): void {
    downloadJson(state, "suspended-ceiling-calculator.json");
  }

  function downloadJson(payload: unknown, filename: string): void {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function exportActiveCalculation(): void {
    const room = cloneRoom(activeRoom);
    const result = calc(room, state.constants);
    const payload = {
      exportedAt: new Date().toISOString(),
      scope: "active-room-calculation",
      room,
      constants: state.constants,
      validation: {
        isValid: validateCombination(cloneRoom(room)),
        constructionLabel: getConstruction(room).label,
        warnings: getValidationWarnings(cloneRoom(room)),
      },
      result,
      formulas: buildCalculationFormulas(room, result, state.constants),
      pricing: {
        currency: "EUR",
        materials: buildMaterialTakeoff([room], state.constants).map((row) => ({
          ...row,
          unitPrice: getMaterialUnitPrice(state.materialPrices, row.key),
          totalPrice: row.quantity * getMaterialUnitPrice(state.materialPrices, row.key),
        })),
      },
      positions: {
        bearingProfilesCm: buildPositions(result.W, room.c, room.offset),
        mountingProfilesCm: buildPositions(result.L, room.b, room.offset),
        hangersCm: buildPositions(result.L, room.a, room.offset),
        extensionsCm: Array.from(
          { length: Math.max(0, Math.ceil(result.L / (state.constants.cdLength * 100)) - 1) },
          (_, index) => (index + 1) * state.constants.cdLength * 100,
        ),
      },
      materials: buildMaterialTakeoff([room], state.constants),
    };
    downloadJson(payload, `calculation-${room.name || room.systemType}.json`);
  }

  async function importJson(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const raw = JSON.parse(await file.text()) as Partial<AppState>;
      if (!Array.isArray(raw.rooms)) throw new Error("Очаква се масив rooms.");
      const rooms = raw.rooms.map((room) => ({ ...createRoom(), ...room, id: crypto.randomUUID() }));
      const draftRoom = raw.draftRoom
        ? { ...createRoom(), ...raw.draftRoom, id: crypto.randomUUID() }
        : rooms[0]
          ? cloneRoom(rooms[0])
          : createRoom();
      const next = {
        rooms,
        draftRoom,
        constants: { ...DEFAULT_CONSTANTS, ...raw.constants },
        materialPrices: raw.materialPrices ?? {},
        activeRoomId: draftRoom.id,
      };
      setState(next);
      saveState(next);
    } finally {
      event.target.value = "";
    }
  }

  function exportExcel(): void {
    const headers = [
      "Стая", "Конструкция", "X", "Y", "Площ", "Носещи реда", "Монтажни реда",
      "Носещи m", "Монтажни m", "Профили общо", "UD бр.", "Връзки", "Окачвачи",
      "Дюбели UD", "Дюбели окачвачи", "Дюбели общо", "Винтове метал", "Винтове гипсокартон", "Удължители",
    ];
    const escapeXml = (value: unknown) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    const rows = state.rooms.map((room) => {
      const result = calc(cloneRoom(room), state.constants);
      return [
        room.name, room.systemType, room.width, room.length, formatNumber(room.area),
        result.bearingCount, result.mountingCount, formatNumber(result.bearingLengthTotal),
        formatNumber(result.mountingLengthTotal), result.cdTotalProfiles, result.udProfiles,
        result.crossConnectors, result.hangersTotal, result.anchorsUd, result.anchorsHangers,
        result.anchorsTotal, result.metalScrews, result.drywallScrews, result.extensionsTotal,
      ];
    });
    const sheetRows = [headers, ...rows]
      .map((row) => `<Row>${row.map((cell) => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join("")}</Row>`)
      .join("");
    const materialHeaders = ["Материал", "Количество", "Ед.", "Ед. цена", "Общо цена", "Точност", "Бележка"];
    const sourceLabel = {
      "knauf-table": "по Knauf таблица",
      geometry: "геометрично",
      estimate: "оценка",
      manual: "ръчно",
    };
    const materialRows = buildMaterialTakeoff(state.rooms, state.constants).map((row) => [
      row.label,
      row.quantity,
      row.unit,
      formatNumber(getMaterialUnitPrice(state.materialPrices, row.key)),
      formatNumber(row.quantity * getMaterialUnitPrice(state.materialPrices, row.key)),
      sourceLabel[row.source],
      row.note,
    ]);
    const materialTotal = buildMaterialTakeoff(state.rooms, state.constants)
      .reduce((sum, row) => sum + row.quantity * getMaterialUnitPrice(state.materialPrices, row.key), 0);
    materialRows.push(["Общо предварителна цена", "", "", "", formatNumber(materialTotal), "", ""]);
    const materialSheetRows = [materialHeaders, ...materialRows]
      .map((row) => `<Row>${row.map((cell) => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join("")}</Row>`)
      .join("");
    const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Стаи"><Table>${sheetRows}</Table></Worksheet><Worksheet ss:Name="Материали"><Table>${materialSheetRows}</Table></Worksheet></Workbook>`;
    const blob = new Blob([workbook], { type: "application/vnd.ms-excel" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "suspended-ceiling-materials.xls";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  async function searchOnlinePrices(): Promise<void> {
    const rows = buildMaterialTakeoff(state.rooms, state.constants)
      .filter((row) => getMaterialUnitPrice(state.materialPrices, row.key) === 0);
    setPriceSearchStatus("Търся цени...");

    const foundPrices: MaterialPrices = {};
    for (const row of rows) {
      const product = getCatalogProductForMaterial(row.label, row.key);
      if (!product) continue;

      for (const source of product.priceSources) {
        const price = await fetchOnlinePrice(source);
        if (price == null) continue;
        foundPrices[row.key] = price;
        break;
      }
    }

    if (Object.keys(foundPrices).length) {
      commit((draft) => {
        Object.entries(foundPrices).forEach(([key, price]) => {
          if (getMaterialUnitPrice(draft.materialPrices, key) === 0) draft.materialPrices[key] = price;
        });
      });
      setPriceSearchStatus(`Намерени цени: ${Object.keys(foundPrices).length}`);
    } else {
      setPriceSearchStatus("Няма намерени нови цени");
    }
    window.setTimeout(() => setPriceSearchStatus(""), 2400);
  }

  const loadClasses = getLoadClasses(activeRoom.systemType, activeRoom.fireProtection, activeRoom.d116Variant, activeRoom.d112Variant);

  return (
    <main className="app-shell">
      <section className="workbench">
        <div className="topbar">
          <div>
            <p className="eyebrow">Knauf D11 calculator</p>
            <h1>Калкулатор за окачени тавани</h1>
          </div>
          <div className="topbar-actions">
            <button type="button" onClick={addRoom}>Нова стая</button>
            <button type="button" className="ghost" onClick={exportJson}>Експорт JSON</button>
            <button type="button" className="ghost" onClick={exportActiveCalculation}>Експорт изчисление</button>
            <button type="button" className="ghost" onClick={exportExcel}>Експорт Excel</button>
            <label className="file-button">Импорт JSON<input type="file" accept="application/json" onChange={importJson} /></label>
          </div>
        </div>

        <div className="workspace-grid">
          <aside className="side-rail">
            <RoomEditor
              room={activeRoom}
              loadClasses={loadClasses}
              isValid={isValid}
              warnings={activeWarnings}
              onSystemChange={changeSystem}
              onResetAuto={resetAutoSpacing}
              onSave={saveCurrentState}
              saveStatus={saveStatus}
              onRoomChange={updateActiveRoom}
            />
          </aside>

          <section className="visual-workspace">
            <ResultCards result={activeResult} room={activeRoom} />
            <Visualization
              room={activeRoom}
              result={activeResult}
              zoom={zoom}
              onZoomChange={setZoom}
            />
          </section>
        </div>

        <ConstantsEditor
          constants={state.constants}
          onChange={(patch) => commit((draft) => { draft.constants = { ...draft.constants, ...patch }; })}
        />
      </section>

      <section className="results-section">
        <RoomsTable
          rooms={state.rooms}
          constants={state.constants}
          activeRoomId={state.activeRoomId}
          onSelect={(roomId) => commit((draft) => {
            const room = draft.rooms.find((item) => item.id === roomId);
            if (!room) return;
            draft.draftRoom = cloneRoom(room);
            draft.activeRoomId = room.id;
          })}
          onDelete={deleteRoom}
          onDeleteAll={deleteAllRooms}
        />
        <MaterialsPanel
          rooms={state.rooms}
          constants={state.constants}
          prices={state.materialPrices}
          onReserveChange={(wastePercent) => commit((draft) => { draft.constants.wastePercent = wastePercent; })}
          onPriceChange={(key, price) => commit((draft) => { draft.materialPrices[key] = price; })}
          onSearchOnlinePrices={searchOnlinePrices}
          priceSearchStatus={priceSearchStatus}
          onAutofillPrices={() => commit((draft) => {
            buildMaterialTakeoff(draft.rooms, draft.constants).forEach((row) => {
              if (getMaterialUnitPrice(draft.materialPrices, row.key) > 0) return;
              const estimate = getEstimatedMaterialPrice(row.label, row.key);
              if (estimate && estimate.price > 0) draft.materialPrices[row.key] = estimate.price;
            });
          })}
        />
      </section>
      <HelpPanel />
    </main>
  );
}

interface RoomEditorProps {
  room: Room;
  loadClasses: LoadClass[];
  isValid: boolean;
  warnings: ReturnType<typeof getValidationWarnings>;
  onSystemChange: (systemType: SystemType) => void;
  onResetAuto: () => void;
  onSave: () => void;
  saveStatus: string;
  onRoomChange: (updater: (room: Room) => void) => void;
}

function RoomEditor({ room, loadClasses, isValid, warnings, onSystemChange, onResetAuto, onSave, saveStatus, onRoomChange }: RoomEditorProps) {
  const construction = CONSTRUCTION_TYPES[room.systemType];
  const d112Variant = room.d112Variant ?? "double_cd";
  const d116Variant = room.d116Variant ?? "ua_cd";
  const boardOptions = getBoardOptions(room.systemType);
  const hangerOptions = getHangerOptions(room.systemType);
  const selectedHanger = hangerOptions.find((option) => option.value === room.hangerType) ?? HANGER_OPTIONS[getDefaultHangerType(room.systemType)];
  const fireLoadClasses = getLoadClasses(room.systemType, true, room.d116Variant, room.d112Variant);
  const aOptions = getAllowedAValues(cloneRoom(room));
  const bOptions = getAllowedBValues(room.systemType);
  const cOptions = getValidCValues(cloneRoom(room));
  const udRule = getUdAnchoringRule(room);
  const automaticLoadClass = getAutomaticLoadClass(cloneRoom(room));
  const estimatedLoadKg = estimateLoadKgPerM2(cloneRoom(room));
  const fireCertification = getFireCertificationCheck(cloneRoom(room));
  const errorCount = warnings.filter((warning) => warning.severity === "error").length;
  const warningCount = warnings.filter((warning) => warning.severity === "warning").length;
  const statusLabel = errorCount > 0 ? "провери" : warningCount > 0 ? `${warningCount} предупрежд.` : "валидна";
  const variantHint = room.systemType === "D112"
    ? D112_VARIANTS[d112Variant].materialHint
    : room.systemType === "D116"
      ? D116_VARIANTS[d116Variant].materialHint
      : construction.materialHint;
  const fireRating = room.fireRating ?? (room.fireProtection ? "fire_table" : "none");
  const isCustom = room.systemType === "CUSTOM";
  return (
    <section className="panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Активна стая</p>
          <h2>{room.name}</h2>
        </div>
        <div className="room-actions">
          {saveStatus && <span className="save-status">{saveStatus}</span>}
          <span className={isValid && !warningCount ? "status ok" : "status warn"}>{statusLabel}</span>
          <button type="button" className="small" onClick={onSave}>Save</button>
        </div>
      </div>

      <div className="field-grid">
        <label>Име
          <input value={room.name} onChange={(event) => onRoomChange((draft) => { draft.name = event.target.value || "Стая"; })} />
        </label>
        <label>Ширина X (cm)
          <input type="number" value={room.width} onChange={(event) => onRoomChange((draft) => { draft.width = Number(event.target.value); })} />
        </label>
        <label>Дължина Y (cm)
          <input type="number" value={room.length} onChange={(event) => onRoomChange((draft) => { draft.length = Number(event.target.value); })} />
        </label>
        <label>Площ (m2)
          <input type="number" step="0.01" value={room.area} onChange={(event) => onRoomChange((draft) => {
            draft.area = Number(event.target.value);
            draft.overrides.area = true;
          })} />
        </label>
      </div>

      <div className="divider" />
      <FireCertificationPanel check={fireCertification} />
      <div className="system-card">
        <label className="system-select-label">Конструкция
          <select value={room.systemType} onChange={(event) => onSystemChange(event.target.value as SystemType)}>
            {Object.entries(CONSTRUCTION_TYPES).map(([value, item]) => (
              <option key={value} value={value}>{item.label}</option>
            ))}
          </select>
        </label>
        {room.systemType === "D112" && (
          <label className="system-select-label">Вариант D112
            <select value={d112Variant} onChange={(event) => onRoomChange((draft) => {
              draft.d112Variant = event.target.value as D112Variant;
              draft.fireRating = draft.fireRating !== "none" && getLoadClasses(draft.systemType, true, draft.d116Variant, draft.d112Variant).length ? draft.fireRating : "none";
              draft.fireProtection = draft.fireRating !== "none";
              draft.loadClass = getLoadClasses(draft.systemType, draft.fireProtection, draft.d116Variant, draft.d112Variant)[0] ?? draft.loadClass;
              draft.overrides.a = false;
              draft.overrides.c = false;
              syncSpacingFromKnaufTable(draft, { keepC: false });
            })}>
              {Object.entries(D112_VARIANTS).map(([value, item]) => (
                <option key={value} value={value}>{item.label}</option>
              ))}
            </select>
          </label>
        )}
        {room.systemType === "D116" && (
          <label className="system-select-label">Вариант D116
            <select value={d116Variant} onChange={(event) => onRoomChange((draft) => {
              draft.d116Variant = event.target.value as D116Variant;
              draft.loadClass = getLoadClasses(draft.systemType, draft.fireProtection, draft.d116Variant, draft.d112Variant)[0] ?? draft.loadClass;
              draft.overrides.a = false;
              draft.overrides.c = false;
              syncSpacingFromKnaufTable(draft, { keepC: false });
            })}>
              {Object.entries(D116_VARIANTS).map(([value, item]) => (
                <option key={value} value={value}>{item.label}</option>
              ))}
            </select>
          </label>
        )}
        {room.systemType === "CUSTOM" && (
          <div className="custom-profile-grid">
            <label>Носещ профил
              <select value={room.customBearingProfile ?? "cd_60_27"} onChange={(event) => onRoomChange((draft) => {
                draft.customBearingProfile = event.target.value as CustomProfileType;
              })}>
                {Object.values(CUSTOM_PROFILE_OPTIONS).map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
            <label>Монтажен профил
              <select value={room.customMountingProfile ?? "cd_60_27"} onChange={(event) => onRoomChange((draft) => {
                draft.customMountingProfile = event.target.value as CustomProfileType;
              })}>
                {Object.values(CUSTOM_PROFILE_OPTIONS).map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
            <label>Периферен профил
              <select value={room.customPerimeterProfile ?? "ud_28_27"} onChange={(event) => onRoomChange((draft) => {
                draft.customPerimeterProfile = event.target.value as CustomProfileType;
              })}>
                {Object.values(CUSTOM_PROFILE_OPTIONS).map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
          </div>
        )}
      </div>

      {warnings.length > 0 && (
        <div className="validation-list">
          {warnings.map((warning) => (
            <div key={warning.code} className={`validation-item ${warning.severity}`}>
              {warning.message}
            </div>
          ))}
        </div>
      )}

      <div className="field-grid">
        <label className="span-2">Режим товар
          <select value={room.loadInputMode ?? "manual"} onChange={(event) => onRoomChange((draft) => {
            draft.loadInputMode = event.target.value as LoadInputMode;
            if (draft.loadInputMode === "auto") draft.loadClass = getAutomaticLoadClass(draft);
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            <option value="manual">Ръчно</option>
            <option value="auto">Автоматично от слоеве</option>
          </select>
          <span className="field-note">Оценка {estimatedLoadKg} kg/m2 {"->"} до {automaticLoadClass} kN/m2</span>
        </label>
        <label className="load-select-label span-2">Натоварване
          <select value={room.loadClass} disabled={room.loadInputMode === "auto"} onChange={(event) => onRoomChange((draft) => {
            draft.loadClass = event.target.value as LoadClass;
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            {loadClasses.map((value) => <option key={value} value={value}>до {value} kN/m2</option>)}
          </select>
        </label>
        <label>Огнезащита/EI
          <select value={fireRating} disabled={!fireLoadClasses.length} onChange={(event) => onRoomChange((draft) => {
            draft.fireRating = event.target.value as FireRating;
            draft.fireProtection = draft.fireRating !== "none";
            if (draft.loadInputMode === "auto") draft.loadClass = getAutomaticLoadClass(draft);
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            {FIRE_RATING_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>{option.label}</option>
            ))}
          </select>
        </label>
        <label className="span-2">Тип/дебелина гипсокартон
          <select value={room.boardType} onChange={(event) => onRoomChange((draft) => {
            draft.boardType = event.target.value as Room["boardType"];
            draft.overrides.b = draft.boardType === "custom";
            if (draft.loadInputMode === "auto") draft.loadClass = getAutomaticLoadClass(draft);
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            {boardOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
        </label>
        <label>Доп. товар (kg/m2)
          <input type="number" min="0" step="0.1" value={room.additionalLoadKgPerM2 ?? 0} onChange={(event) => onRoomChange((draft) => {
            draft.additionalLoadKgPerM2 = Math.max(0, Number(event.target.value) || 0);
            if (draft.loadInputMode === "auto") draft.loadClass = getAutomaticLoadClass(draft);
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })} />
        </label>
        <label className="span-2">Тип окачвач
          <select value={selectedHanger.value} onChange={(event) => onRoomChange((draft) => {
            draft.hangerType = event.target.value as HangerType;
          })}>
            {hangerOptions.map((option) => (
              <option key={option.value} value={option.value}>{option.label}</option>
            ))}
          </select>
          <span className="field-note">{selectedHanger.useWhen}</span>
        </label>
      </div>

      <div className="spacing-head">
        <strong>Разстояния</strong>
        <button type="button" className="ghost small" onClick={onResetAuto}>{isCustom ? "Върни стандартни" : "Върни по Knauf"}</button>
      </div>
      <div className="field-grid spacing-card">
        {isCustom ? (
          <>
            <SelectNumberField label="a Разстояние между окачвачи (mm)" value={room.a} manual={room.overrides.a} options={CUSTOM_A_OPTIONS} onChange={(value) => onRoomChange((draft) => {
              draft.a = value;
              syncDistanceOverride(draft, "a");
            })} />
            <SelectNumberField label="b Разстояние между монтажни профили (mm)" value={room.b} manual={room.overrides.b} options={CUSTOM_B_OPTIONS} onChange={(value) => onRoomChange((draft) => {
              draft.b = value;
              syncDistanceOverride(draft, "b");
            })} />
            <SelectNumberField label="c Разстояние между носещи профили (mm)" value={room.c} manual={room.overrides.c} options={CUSTOM_C_OPTIONS} onChange={(value) => onRoomChange((draft) => {
              draft.c = value;
              syncDistanceOverride(draft, "c");
            })} />
          </>
        ) : (
          <>
            <SelectNumberField label="a Разстояние между окачвачи (mm)" value={room.a} manual={room.overrides.a} options={aOptions} onChange={(value) => onRoomChange((draft) => {
              draft.a = value;
              syncDistanceOverride(draft, "a");
            })} />
            <SelectNumberField label="b Разстояние между монтажни CD профили (mm)" value={room.b} manual={room.overrides.b} options={bOptions} onChange={(value) => onRoomChange((draft) => {
              draft.b = value;
              syncDistanceOverride(draft, "b");
            })} />
            <SelectNumberField label="c Разстояние между носещи CD/UA профили (mm)" value={room.c} manual={room.overrides.c} options={cOptions} onChange={(value) => onRoomChange((draft) => {
              draft.c = value;
              syncDistanceOverride(draft, "c");
              draft.overrides.a = false;
              syncSpacingFromKnaufTable(draft, { keepC: true });
            })} />
          </>
        )}
        {isCustom ? (
          <>
            <SelectNumberField label="Начално отстояние (cm)" value={room.offset} manual={room.overrides.offset} options={CUSTOM_OFFSET_OPTIONS} onChange={(value) => onRoomChange((draft) => {
              draft.offset = value;
              syncDistanceOverride(draft, "offset");
            })} />
            <SelectNumberField label={`Дюбели периферия (${udRule.mode}, mm)`} value={room.udAnchorSpacing} manual={room.overrides.udAnchorSpacing} options={CUSTOM_UD_ANCHOR_OPTIONS} onChange={(value) => onRoomChange((draft) => {
              draft.udAnchorSpacing = value;
              syncDistanceOverride(draft, "udAnchorSpacing");
            })} />
          </>
        ) : (
          <>
            <NumberField label="Начално отстояние (cm)" value={room.offset} manual={room.overrides.offset} onChange={(value) => onRoomChange((draft) => {
              draft.offset = value;
              syncDistanceOverride(draft, "offset");
            })} />
            <NumberField label={`UD дюбели (${udRule.mode}, mm)`} value={room.udAnchorSpacing} manual={room.overrides.udAnchorSpacing} onChange={(value) => onRoomChange((draft) => {
              draft.udAnchorSpacing = value;
              syncDistanceOverride(draft, "udAnchorSpacing");
            })} />
          </>
        )}
      </div>
      <p className="hint">{variantHint}</p>
      <p className="hint">{udRule.note}</p>
    </section>
  );
}

function SelectNumberField({ label, value, manual, options, onChange }: {
  label: string;
  value: number;
  manual: boolean;
  options: number[];
  onChange: (value: number) => void;
}) {
  const allOptions = options.includes(value) ? options : [...options, value].sort((a, b) => a - b);
  return (
    <label>{label}
      <select value={value} onChange={(event) => onChange(Number(event.target.value))}>
        {allOptions.map((option) => <option key={option} value={option}>{option}</option>)}
      </select>
      <span className={manual ? "auto-state manual" : "auto-state"}>{manual ? "ръчно" : "по Knauf"}</span>
    </label>
  );
}

function NumberField({ label, value, manual, onChange }: { label: string; value: number; manual: boolean; onChange: (value: number) => void }) {
  return (
    <label>{label}
      <input type="number" value={value} onChange={(event) => onChange(Number(event.target.value))} />
      <span className={manual ? "auto-state manual" : "auto-state"}>{manual ? "ръчно" : "по Knauf"}</span>
    </label>
  );
}

function FireCertificationPanel({ check }: { check: ReturnType<typeof getFireCertificationCheck> }) {
  return (
    <section className={`fire-cert-card ${check.status} certification-block`}>
      <div className="fire-cert-head">
        <span className="fire-cert-dot" />
        <div>
          <strong>{check.label}</strong>
          <p>{check.summary}</p>
        </div>
      </div>
      {check.issues.length > 0 && (
        <div className="fire-cert-list">
          {check.issues.map((issue) => (
            <div key={issue.code} className={`fire-cert-issue ${issue.type}`}>
              <span>{issue.type === "missing" ? "Липсва" : "Грешно"}</span>
              <p>{issue.message}</p>
              {issue.allowedValues?.length ? <small>Допустимо: {issue.allowedValues.join(", ")}</small> : null}
            </div>
          ))}
        </div>
      )}
    </section>
  );
}

function ConstantsEditor({ constants, onChange }: { constants: CalculatorConstants; onChange: (patch: Partial<CalculatorConstants>) => void }) {
  return (
    <section className="panel compact settings-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Глобални константи</p>
          <h2>Настройки</h2>
        </div>
      </div>
      <div className="settings-groups">
        <SettingsGroup title="Профили и летви" note="Дължини на покупните елементи за профили/летви. Използват се за броя цели бройки.">
          <NumberInput label="CD профил (m)" value={constants.cdLength} step={0.1} onChange={(value) => onChange({ cdLength: value })} />
          <NumberInput label="UD профил (m)" value={constants.udLength} step={0.1} onChange={(value) => onChange({ udLength: value })} />
          <NumberInput label="UA профил (m)" value={constants.uaLength} step={0.1} onChange={(value) => onChange({ uaLength: value })} />
          <NumberInput label="Дървена летва (m)" value={constants.woodBattenLength} step={0.1} onChange={(value) => onChange({ woodBattenLength: value })} />
        </SettingsGroup>

        <SettingsGroup title="Крепежи" note="Норми за винтове. Дюбелите за UD и окачвачи се изчисляват отделно по геометрията.">
          <NumberInput label="Винтове/връзка" value={constants.metalScrewsPerCrossConnector} onChange={(value) => onChange({ metalScrewsPerCrossConnector: value })} />
          <NumberInput label="Винтове/окачвач" value={constants.metalScrewsPerDirectHanger} onChange={(value) => onChange({ metalScrewsPerDirectHanger: value })} />
          <NumberInput label="TN винтове/m2/слой" value={constants.drywallScrewsPerM2} onChange={(value) => onChange({ drywallScrewsPerM2: value })} />
        </SettingsGroup>

        <SettingsGroup title="Плоскости" note="Размер на листа за превръщане на площта в бройки. Слоевете идват от типа плоскост, освен при Custom.">
          <NumberInput label="Ширина плоскост (m)" value={constants.boardWidth} step={0.1} onChange={(value) => onChange({ boardWidth: value })} />
          <NumberInput label="Дължина плоскост (m)" value={constants.boardLength} step={0.1} onChange={(value) => onChange({ boardLength: value })} />
          <NumberInput label="Custom слоеве плоскости" value={constants.boardLayers} onChange={(value) => onChange({ boardLayers: Math.max(1, Math.round(value)) })} />
        </SettingsGroup>

        <SettingsGroup title="Фуги и периметър" note="Ориентировъчни разходни норми за довършителни материали.">
          <NumberInput label="Фуголента m/m2" value={constants.jointTapePerM2} step={0.1} onChange={(value) => onChange({ jointTapePerM2: value })} />
          <NumberInput label="Шпакловка kg/m2" value={constants.jointCompoundKgPerM2} step={0.05} onChange={(value) => onChange({ jointCompoundKgPerM2: value })} />
          <NumberInput label="Trenn-Fix x периметър" value={constants.trennFixPerimeterMultiplier} step={0.1} onChange={(value) => onChange({ trennFixPerimeterMultiplier: value })} />
        </SettingsGroup>

        <SettingsGroup title="Изолация" note="Минералната вата се добавя автоматично при огнезащита или ръчно оттук.">
          <label>Минерална вата
            <select value={String(constants.mineralWoolEnabled)} onChange={(event) => onChange({ mineralWoolEnabled: event.target.value === "true" })}>
              <option value="false">Само при огнезащита</option>
              <option value="true">Винаги включи</option>
            </select>
          </label>
          <NumberInput label="Вата дебелина (mm)" value={constants.mineralWoolThickness} onChange={(value) => onChange({ mineralWoolThickness: value })} />
        </SettingsGroup>
      </div>
    </section>
  );
}

function SettingsGroup({ title, note, children }: { title: string; note: string; children: ReactNode }) {
  return (
    <section className="settings-group">
      <div className="settings-group-head">
        <h3>{title}</h3>
        <p>{note}</p>
      </div>
      <div className="field-grid">{children}</div>
    </section>
  );
}

function NumberInput({ label, value, step = 1, onChange }: { label: string; value: number; step?: number; onChange: (value: number) => void }) {
  return (
    <label>{label}
      <input type="number" step={step} value={value} onChange={(event) => onChange(Number(event.target.value))} />
    </label>
  );
}

function ResultCards({ result, room }: { result: CalcResult; room: Room }) {
  return (
    <div className="metric-strip">
      <Metric label="Площ" value={Number(formatNumber(room.area))} suffix="m2" note={`${room.width} x ${room.length} cm`} />
      <Metric label="Носещи редове" value={result.bearingCount} suffix="бр." note={`${formatNumber(result.bearingLengthTotal)} m носещи профили`} />
      <Metric label="Монтажни редове" value={result.mountingCount} suffix="бр." note={`${formatNumber(result.mountingLengthTotal)} m монтажни профили`} />
      <Metric label="Профили общо" value={result.cdTotalProfiles} suffix="бр." note={`${formatNumber(result.cdTotalLength)} m носещи + монтажни`} />
      <Metric label="UD профили" value={result.udProfiles} suffix="бр." note={`${formatNumber(result.udTotalLength)} m по периметъра`} />
      <Metric label="Връзки" value={result.crossConnectors} suffix="бр." note="пресичания носещи x монтажни" />
      <Metric label="Окачвачи" value={result.hangersTotal} suffix="бр." note={`${result.hangersPerBearing} бр. на носещ ред`} />
      <Metric label="Дюбели общо" value={result.anchorsTotal} suffix="бр." note={`UD ${result.anchorsUd}, окачвачи ${result.anchorsHangers}`} />
      <Metric label="Винтове" value={result.metalScrews + result.drywallScrews} suffix="бр." note={`метал ${result.metalScrews}, гипсокартон ${result.drywallScrews}`} />
      <Metric label="Удължители" value={result.extensionsTotal} suffix="бр." note="само геометрична оценка" />
    </div>
  );
}

function Metric({ label, value, suffix, note }: { label: string; value: number; suffix: string; note: string }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{suffix}</small>
      <em>{note}</em>
    </div>
  );
}

function Visualization({ room, result, zoom, onZoomChange }: { room: Room; result: CalcResult; zoom: number; onZoomChange: (zoom: number) => void }) {
  const bearingPositions = buildPositions(result.W, room.c, room.offset);
  const mountingPositions = buildPositions(result.L, room.b, room.offset);
  const hangerPositions = buildPositions(result.L, room.a, room.offset);
  const width = 980;
  const height = 560;
  const pad = 86;
  const gridW = width - pad * 2;
  const gridH = height - pad * 2;
  const xScale = gridW / result.L;
  const yScale = gridH / result.W;
  const segmentLengthCm = 100 * 4;
  const extensionPoints: number[] = [];
  for (let pos = segmentLengthCm; pos < result.L; pos += segmentLengthCm) extensionPoints.push(pos);

  return (
    <section className="visual-panel">
      <div className="visual-toolbar">
        <div>
          <p className="eyebrow">Работна схема</p>
          <h2>{room.name} - {room.systemType}</h2>
        </div>
        <div className="zoom-group">
          <button type="button" className="ghost small" onClick={() => onZoomChange(Math.max(0.7, zoom - 0.1))}>-</button>
          <span>{Math.round(zoom * 100)}%</span>
          <button type="button" className="ghost small" onClick={() => onZoomChange(Math.min(1.8, zoom + 0.1))}>+</button>
          <button type="button" className="ghost small" onClick={() => onZoomChange(1)}>100%</button>
        </div>
      </div>
      <div className="drawing-scroll">
        <svg className="ceiling-svg" style={{ transform: `scale(${zoom})` }} viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Схема на профилите и окачвачите">
          <defs>
            <pattern id="grid" width="32" height="32" patternUnits="userSpaceOnUse">
              <path d="M 32 0 L 0 0 0 32" fill="none" stroke="#d7e0e3" strokeWidth="1" />
            </pattern>
          </defs>
          <rect x="0" y="0" width={width} height={height} fill="#f6f8f4" />
          <rect x={pad} y={pad} width={gridW} height={gridH} rx="6" fill="url(#grid)" stroke="#1d6f68" strokeWidth="3" />
          <text x={pad + gridW - 8} y={pad + gridH + 18} className="axis-label" textAnchor="end">L {result.L.toFixed(0)} cm</text>
          <text x={pad + gridW - 8} y={pad + gridH + 36} className="axis-label" textAnchor="end">W {result.W.toFixed(0)} cm</text>

          {hangerPositions.map((hanger, index) => {
            const x = pad + hanger * xScale;
            const y1 = pad - 18;
            const y2 = pad;
            const labelY = pad - 26 - ((index % 2) * 13);
            return (
              <g key={`hanger-dim-${hanger}`}>
                <line x1={x} y1={y1} x2={x} y2={y2} className="hanger-dimension-line" />
                <text x={x} y={labelY} className="hanger-label" textAnchor="middle">{Math.round(hanger)} cm</text>
              </g>
            );
          })}

          {extensionPoints.map((extension, index) => {
            const x = pad + extension * xScale;
            const y1 = pad + gridH;
            const y2 = pad + gridH + 24;
            const labelY = pad + gridH + 39 + ((index % 2) * 13);
            return (
              <g key={`extension-dim-${extension}`}>
                <line x1={x} y1={y1} x2={x} y2={y2} className="extension-dimension-line" />
                <text x={x} y={labelY} className="extension-label" textAnchor="middle">{Math.round(extension)} cm</text>
              </g>
            );
          })}

          {mountingPositions.map((position) => {
            const x = pad + position * xScale;
            return (
              <g key={`mount-${position}`}>
                <line x1={x} y1={pad} x2={x} y2={pad + gridH} className="mounting-line" />
                <text x={x + 6} y={pad + 18} className="position-label">{Math.round(position)}</text>
              </g>
            );
          })}

          {bearingPositions.map((position) => {
            const y = pad + position * yScale;
            return (
              <g key={`bearing-${position}`}>
                <line x1={pad} y1={y} x2={pad + gridW} y2={y} className="bearing-line" />
                <text x={pad + 8} y={y - 7} className="position-label bearing">{Math.round(position)}</text>
                {hangerPositions.map((hanger) => {
                  const hx = pad + hanger * xScale;
                  return (
                    <g key={`hanger-${position}-${hanger}`}>
                      <circle cx={hx} cy={y} r="5" className="hanger-dot" />
                    </g>
                  );
                })}
                {extensionPoints.map((extension) => (
                  <line key={`extension-${position}-${extension}`} x1={pad + extension * xScale} y1={y - 12} x2={pad + extension * xScale} y2={y + 12} className="extension-mark" />
                ))}
              </g>
            );
          })}
        </svg>
      </div>
      <div className="legend-grid">
        <span><i className="legend bearing-line-swatch" /> Носещ профил</span>
        <span><i className="legend mounting-line-swatch" /> Монтажен профил</span>
        <span><i className="legend hanger-dot-swatch" /> Окачвач с разстояние от стената</span>
        <span><i className="legend extension-swatch" /> Удължител</span>
      </div>
    </section>
  );
}

function RoomsTable({ rooms, constants, activeRoomId, onSelect, onDelete, onDeleteAll }: {
  rooms: Room[];
  constants: CalculatorConstants;
  activeRoomId: string;
  onSelect: (roomId: string) => void;
  onDelete: (roomId: string) => void;
  onDeleteAll: () => void;
}) {
  return (
    <section className="panel table-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Стаи</p>
          <h2>Количества по стаи</h2>
        </div>
        <button type="button" className="danger small" disabled={!rooms.length} onClick={onDeleteAll}>Изтрий всички</button>
      </div>
      {!rooms.length ? (
        <p className="empty-state">Няма запазени стаи. Активната стая ще се появи тук след Save.</p>
      ) : (
        <div className="table-scroll">
          <table>
            <thead>
              <tr>
                <th>Стая</th><th>Система</th><th>X</th><th>Y</th><th>m2</th><th>Носещи</th><th>Монтажни</th><th>Профили</th><th>UD</th><th>Връзки</th><th>Окачвачи</th><th>Дюбели</th><th>Винтове</th><th></th>
              </tr>
            </thead>
            <tbody>
              {rooms.map((room) => {
                const result = calc(cloneRoom(room), constants);
                return (
                  <tr key={room.id} className={room.id === activeRoomId ? "selected-row" : ""}>
                    <td><button type="button" className="link-button" onClick={() => onSelect(room.id)}>{room.name}</button></td>
                    <td>{room.systemType}</td>
                    <td>{room.width}</td>
                    <td>{room.length}</td>
                    <td>{formatNumber(room.area)}</td>
                    <td>{result.bearingCount}</td>
                    <td>{result.mountingCount}</td>
                    <td>{result.cdTotalProfiles}</td>
                    <td>{result.udProfiles}</td>
                    <td>{result.crossConnectors}</td>
                    <td>{result.hangersTotal}</td>
                    <td>{result.anchorsTotal}</td>
                    <td>{result.metalScrews + result.drywallScrews}</td>
                    <td><button type="button" className="danger small" onClick={() => onDelete(room.id)}>Изтрий</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

function MaterialsPanel({ rooms, constants, prices, priceSearchStatus, onReserveChange, onPriceChange, onSearchOnlinePrices, onAutofillPrices }: {
  rooms: Room[];
  constants: CalculatorConstants;
  prices: MaterialPrices;
  priceSearchStatus: string;
  onReserveChange: (wastePercent: number) => void;
  onPriceChange: (key: string, price: number) => void;
  onSearchOnlinePrices: () => void;
  onAutofillPrices: () => void;
}) {
  const rows = buildMaterialTakeoff(rooms, constants);
  const totalPrice = rows.reduce((sum, row) => sum + row.quantity * getMaterialUnitPrice(prices, row.key), 0);
  const sourceLabel = {
    "knauf-table": "по Knauf таблица",
    geometry: "геометрично",
    estimate: "оценка",
    manual: "ръчно",
  };
  return (
    <section className="panel materials-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Материали</p>
          <h2>Общо за всички стаи</h2>
        </div>
      </div>
      <div className="reserve-field">
        <NumberInput label="Резерв (%)" value={constants.wastePercent} step={0.1} onChange={onReserveChange} />
      </div>
      <div className="price-summary">
        <span>Предварителна цена</span>
        <strong>{formatCurrency(totalPrice)}</strong>
        <small>по въведените единични цени в евро, с включен резерв в количествата</small>
        <div className="price-actions">
          <button type="button" className="ghost small" onClick={onSearchOnlinePrices}>Потърси цени</button>
          <button type="button" className="ghost small" onClick={onAutofillPrices}>Попълни ориентири</button>
        </div>
        {priceSearchStatus && <em>{priceSearchStatus}</em>}
      </div>
      <div className="material-list">
        {rows.map((row) => {
          const unitPrice = getMaterialUnitPrice(prices, row.key);
          const rowTotal = row.quantity * unitPrice;
          const estimate = getEstimatedMaterialPrice(row.label, row.key);
          return (
            <div key={row.key} className="material-row priced">
              <span>{row.label}</span>
              <strong>{row.quantity} {row.unit}</strong>
              <label>Ед. цена
                <input type="number" min="0" step="0.01" value={unitPrice} onChange={(event) => onPriceChange(row.key, Math.max(0, Number(event.target.value) || 0))} />
              </label>
              <b>{formatCurrency(rowTotal)}</b>
              <small>{sourceLabel[row.source]} - {row.note}{estimate ? `; каталог: ${getCatalogProductForMaterial(row.label, row.key)?.knaufName}; ориентир: ${formatCurrency(estimate.price)} (${estimate.source})` : ""}</small>
            </div>
          );
        })}
      </div>
    </section>
  );
}

function HelpPanel() {
  return (
    <section className="panel help-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Help</p>
          <h2>Теория и означения</h2>
        </div>
      </div>
      <div className="help-grid">
        <article>
          <h3>Окачвачи</h3>
          <dl>
            {Object.values(HANGER_OPTIONS).map((option) => (
              <div key={option.value}>
                <dt>{option.label}</dt>
                <dd>{option.description} {option.useWhen}</dd>
              </div>
            ))}
          </dl>
        </article>
        <article>
          <h3>EI и огнезащита</h3>
          <p><strong>EI30/EI60/EI90/EI120</strong> означава, че системата запазва цялост и изолационна способност съответно 30, 60, 90 или 120 минути при стандартно пожарно изпитване.</p>
          <p>Текущият избор "Огнезащита по D11 таблица" използва fire таблиците, но конкретният EI клас и посоката отдолу/отгоре трябва да се проверят по системния детайл.</p>
        </article>
        <article>
          <h3>Разстояния</h3>
          <p><strong>a</strong> е разстояние между окачвачите. <strong>b</strong> е разстояние между монтажните профили или летви. <strong>c</strong> е разстояние между носещите профили или летви.</p>
          <p><strong>kN/m2</strong> е клас натоварване. Бележката <strong>0.40 kN</strong> се отнася до носимоспособност на окачвач/закрепване.</p>
        </article>
        <article>
          <h3>Профили</h3>
          <p><strong>CD 60/27</strong> е таванен профил. <strong>UD 28/27</strong> е периферен профил. <strong>UA 50/40</strong> е усилен профил за D116. <strong>UW 50/40</strong> се използва в някои UA детайли.</p>
        </article>
        <article>
          <h3>Обшивка и фуги</h3>
          <p><strong>TN</strong> са винтове за гипсокартон към метал. <strong>LN/LB</strong> са винтове за метални връзки и профили.</p>
          <p>Плоскостите се смятат по площ, размер на листа и брой слоеве. Фуголента, шпакловка и Trenn-Fix са ориентировъчни количества и трябва да се сверят с избраната технология.</p>
        </article>
      </div>
    </section>
  );
}

export default App;
