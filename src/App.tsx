import { ChangeEvent, ReactNode, useState } from "react";
import {
  CalcResult,
  CalculatorConstants,
  CutBar,
  CutOptimizationResult,
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
  MaterialTakeoffItem,
  Room,
  SystemType,
  applyAutoABC,
  buildCutOptimizationInput,
  buildLinearPositions,
  buildMaterialTakeoff,
  buildSuspendedCeilingLayout,
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
  optimizeSuspendedCeilingCuts,
  estimateLoadKgPerM2,
  syncSpacingFromKnaufTable,
} from "./domain/calculator";

const STORAGE_KEY = "d113-calculator-v2";
const BGN_PER_EUR = 1.95583;
const DEFAULT_EXPORT_DIRECTORY_NAME = "Knauf calculator";

declare global {
  interface Window {
    showDirectoryPicker?: () => Promise<FileSystemDirectoryHandle>;
  }

  interface FileSystemDirectoryHandle {
    readonly name: string;
    getDirectoryHandle(name: string, options?: { create?: boolean }): Promise<FileSystemDirectoryHandle>;
    getFileHandle(name: string, options?: { create?: boolean }): Promise<FileSystemFileHandle>;
  }

  interface FileSystemFileHandle {
    createWritable(): Promise<FileSystemWritableFileStream>;
  }

  interface FileSystemWritableFileStream {
    write(data: Blob | string): Promise<void>;
    close(): Promise<void>;
  }
}

interface AppState {
  rooms: Room[];
  draftRoom: Room;
  constants: CalculatorConstants;
  materialPrices: MaterialPrices;
  activeRoomId: string;
}

type MaterialPrices = Record<string, number>;
type CutPlanState = { plan: CutOptimizationResult; error: null } | { plan: null; error: string };
type AppSection = "room" | "settings" | "materials" | "help";
type RoomWorkspacePanel = "visual" | "cut";
type ExportContentType = "room-cards" | "table";
type RoomCardExportFileType = "pdf" | "png" | "html";
type TableExportFileType = "excel" | "json" | "html";
type ExportFileType = RoomCardExportFileType | TableExportFileType;

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

function formatHangerOptionLabel(option: (typeof HANGER_OPTIONS)[HangerType]): string {
  return option.capacityKn == null
    ? `${option.label} - без зададена носимоспособност`
    : `${option.label} - ${option.capacityKn.toFixed(2)} kN`;
}

function recalculateRoomWithCurrentLogic(source: Room, constants: CalculatorConstants): Room {
  const room = {
    ...createRoom(source.name || "Стая"),
    ...source,
    overrides: { ...createRoom().overrides, ...source.overrides },
  };
  const manual = {
    area: room.area,
    a: room.a,
    b: room.b,
    c: room.c,
    udAnchorSpacing: room.udAnchorSpacing,
  };

  calc(room, constants);
  if (room.loadInputMode === "auto") room.loadClass = getAutomaticLoadClass(room);
  syncSpacingFromKnaufTable(room, { keepC: room.overrides.c });

  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType, room.d116Variant, room.d112Variant);
  if (!room.overrides.udAnchorSpacing) room.udAnchorSpacing = auto.udAnchorSpacing;
  if (!room.overrides.area && room.width && room.length) room.area = (room.width * room.length) / 10000;

  if (room.overrides.area) room.area = manual.area;
  if (room.overrides.a) room.a = manual.a;
  if (room.overrides.b) room.b = manual.b;
  if (room.overrides.c) room.c = manual.c;
  if (room.overrides.udAnchorSpacing) room.udAnchorSpacing = manual.udAnchorSpacing;

  calc(room, constants);
  return room;
}

function formatNumber(value: number, digits = 2): string {
  return Number(value).toFixed(digits);
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

function buildSafeCutPlan(room: Room, result: CalcResult, constants: CalculatorConstants): CutPlanState {
  try {
    return {
      plan: optimizeSuspendedCeilingCuts(buildCutOptimizationInput(room, result, constants), {
        stockLengthCm: Math.max(constants.cdLength, constants.udLength) * 100,
        kerfCm: 0.3,
        minReusableOffcutCm: 20,
        perTypeStockLengthsCm: {
          carrier: constants.cdLength * 100,
          mounting: constants.cdLength * 100,
          ud: constants.udLength * 100,
        },
      }),
      error: null,
    };
  } catch (error) {
    return {
      plan: null,
      error: error instanceof Error ? error.message : "Неуспешен разкрой.",
    };
  }
}

function escapeHtml(value: unknown): string {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function sanitizeFilename(value: string): string {
  const normalized = value.trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-").replace(/\s+/g, " ");
  return (normalized || "Стая").slice(0, 80);
}

function svgText(x: number, y: number, value: unknown, className = "", maxChars = 90, lineHeight = 18): { markup: string; height: number } {
  const words = String(value).split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let line = "";
  words.forEach((word) => {
    const next = line ? `${line} ${word}` : word;
    if (next.length > maxChars && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
  });
  if (line) lines.push(line);
  const safeLines = lines.length ? lines : [""];
  const markup = `<text x="${x}" y="${y}" class="${className}">${safeLines.map((item, index) => `<tspan x="${x}" dy="${index === 0 ? 0 : lineHeight}">${escapeHtml(item)}</tspan>`).join("")}</text>`;
  return { markup, height: safeLines.length * lineHeight };
}

function svgDataUrl(svg: string): string {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function renderWorkingSchemeSvg(room: Room, result: CalcResult, constants: CalculatorConstants): string {
  const layout = buildSuspendedCeilingLayout({
    roomWidthCm: result.W,
    roomLengthCm: result.L,
    profileLengthCm: constants.cdLength * 100,
    carrierRowSpacingCm: room.c / 10,
    hangerSpacingCm: room.a / 10,
    firstHangerOffsetCm: constants.profileEdgeOffsetCm,
  });
  const bearingPositions = layout.carrierRowsYcm;
  const mountingPositions = buildLinearPositions(result.L, room.b / 10, constants.profileEdgeOffsetCm);
  const hangerPositions = layout.hangerPositionsCm;
  const extensionLayout = layout.carrierExtensions;
  const extensionPositions = Array.from(new Set(extensionLayout.flatMap((line) => line.pointsCm))).sort((left, right) => left - right);
  const width = 980;
  const height = 560;
  const pad = 86;
  const gridW = width - pad * 2;
  const gridH = height - pad * 2;
  const xScale = gridW / result.L;
  const yScale = gridH / result.W;

  return `<svg class="ceiling-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Схема на профилите и окачвачите">
    <defs>
      <pattern id="report-grid" width="32" height="32" patternUnits="userSpaceOnUse">
        <path d="M 32 0 L 0 0 0 32" fill="none" stroke="#d7e0e3" stroke-width="1" />
      </pattern>
    </defs>
    <rect x="0" y="0" width="${width}" height="${height}" fill="#f6f8f4" />
    <rect x="${pad}" y="${pad}" width="${gridW}" height="${gridH}" rx="6" fill="url(#report-grid)" stroke="#1d6f68" stroke-width="3" />
    <text x="${pad + gridW - 8}" y="${pad + gridH + 18}" class="axis-label" text-anchor="end">L ${result.L.toFixed(0)} cm</text>
    <text x="${pad + gridW - 8}" y="${pad + gridH + 36}" class="axis-label" text-anchor="end">W ${result.W.toFixed(0)} cm</text>
    ${hangerPositions.map((hanger, index) => {
      const x = pad + hanger * xScale;
      const labelY = pad - 26 - ((index % 2) * 13);
      return `<g><line x1="${x}" y1="${pad - 18}" x2="${x}" y2="${pad}" class="hanger-dimension-line" /><text x="${x}" y="${labelY}" class="hanger-label" text-anchor="middle">${Math.round(hanger)} cm</text></g>`;
    }).join("")}
    ${extensionPositions.map((extension, index) => {
      const x = pad + extension * xScale;
      const labelY = pad + gridH + 39 + ((index % 2) * 13);
      return `<g><line x1="${x}" y1="${pad + gridH}" x2="${x}" y2="${pad + gridH + 24}" class="extension-dimension-line" /><text x="${x}" y="${labelY}" class="extension-label" text-anchor="middle">${Math.round(extension)} cm</text></g>`;
    }).join("")}
    ${mountingPositions.map((position) => {
      const x = pad + position * xScale;
      return `<g><line x1="${x}" y1="${pad}" x2="${x}" y2="${pad + gridH}" class="mounting-line" /><text x="${x + 6}" y="${pad + 18}" class="position-label">${Math.round(position)}</text></g>`;
    }).join("")}
    ${bearingPositions.map((position, lineIndex) => {
      const y = pad + position * yScale;
      const lineExtensions = extensionLayout.find((line) => line.lineIndex === lineIndex)?.pointsCm ?? [];
      return `<g>
        <line x1="${pad}" y1="${y}" x2="${pad + gridW}" y2="${y}" class="bearing-line" />
        <text x="${pad + 8}" y="${y - 7}" class="position-label bearing">${Math.round(position)}</text>
        ${hangerPositions.map((hanger) => `<circle cx="${pad + hanger * xScale}" cy="${y}" r="5" class="hanger-dot" />`).join("")}
        ${lineExtensions.map((extension) => `<line x1="${pad + extension * xScale}" y1="${y - 12}" x2="${pad + extension * xScale}" y2="${y + 12}" class="extension-mark" />`).join("")}
      </g>`;
    }).join("")}
  </svg>`;
}

function getMaterialRooms(row: MaterialTakeoffItem, rooms: Room[]): Room[] {
  const prefix = row.key.split("-")[0].toUpperCase();
  const systemRooms = prefix === "CUSTOM"
    ? rooms.filter((room) => room.systemType === "CUSTOM")
    : rooms.filter((room) => room.systemType === prefix);
  const hangerMatch = row.key.match(/-hangers-(.+)$/);
  if (!hangerMatch) return systemRooms;
  const hangerType = hangerMatch[1];
  return systemRooms.filter((room) => getEffectiveRoomHangerType(room) === hangerType);
}

function getEffectiveRoomHangerType(room: Room): HangerType {
  const value = room.hangerType;
  if (room.systemType === "CUSTOM" && value && value in HANGER_OPTIONS) return value;
  if (value && value in HANGER_OPTIONS && HANGER_OPTIONS[value].systems.includes(room.systemType)) return value;
  return getDefaultHangerType(room.systemType);
}

function formatReserveSuffix(constants: CalculatorConstants): string {
  return constants.wastePercent > 0 ? ` Крайното количество включва ${formatNumber(constants.wastePercent, 1)}% резерв.` : "";
}

function sumRoomResults(rooms: Room[], constants: CalculatorConstants) {
  return rooms.reduce((total, sourceRoom) => {
    const room = cloneRoom(sourceRoom);
    const result = calc(room, constants);
    const boardLayers = getEffectiveBoardLayers(room, constants);
    total.area += Number(room.area) || 0;
    total.boardArea += (Number(room.area) || 0) * boardLayers;
    total.bearingLength += result.bearingLengthTotal;
    total.mountingLength += result.mountingLengthTotal;
    total.cdLength += result.cdTotalLength;
    total.udLength += result.udTotalLength;
    total.bearingRows += result.bearingCount;
    total.mountingRows += result.mountingCount;
    total.crossConnectors += result.crossConnectors;
    total.hangers += result.hangersTotal;
    total.anchorsUd += result.anchorsUd;
    total.anchorsHangers += result.anchorsHangers;
    total.metalScrews += result.metalScrews;
    total.drywallScrews += result.drywallScrews;
    total.extensions += result.extensionsTotal;
    total.boardLayers += boardLayers;
    return total;
  }, {
    area: 0,
    boardArea: 0,
    bearingLength: 0,
    mountingLength: 0,
    cdLength: 0,
    udLength: 0,
    bearingRows: 0,
    mountingRows: 0,
    crossConnectors: 0,
    hangers: 0,
    anchorsUd: 0,
    anchorsHangers: 0,
    metalScrews: 0,
    drywallScrews: 0,
    extensions: 0,
    boardLayers: 0,
  });
}

function buildMaterialCalculationInfo(row: MaterialTakeoffItem, rooms: Room[], constants: CalculatorConstants): string {
  const materialRooms = getMaterialRooms(row, rooms);
  const totals = sumRoomResults(materialRooms, constants);
  const roomText = materialRooms.length === 1 ? "1 стая" : `${materialRooms.length} стаи`;
  const reserve = formatReserveSuffix(constants);
  const boardSize = constants.boardWidth * constants.boardLength;
  const averageUdSpacing = materialRooms.length
    ? Math.round(materialRooms.reduce((sum, room) => sum + room.udAnchorSpacing, 0) / materialRooms.length)
    : 0;

  if (row.key.includes("cd-60-27") && !row.key.includes("connectors")) {
    if (row.key.startsWith("d116-")) {
      return `${roomText}: ${formatNumber(totals.mountingLength)} m монтажни CD. Бройките са по дължина на CD профил ${constants.cdLength} m.${reserve}`;
    }
    return `${roomText}: общо ${formatNumber(totals.cdLength)} m CD профили - ${formatNumber(totals.bearingLength)} m носещи и ${formatNumber(totals.mountingLength)} m монтажни. Бройките са закръглени към цели профили по ${constants.cdLength} m.${reserve}`;
  }

  if (row.key.includes("ua-50-40")) {
    return `${roomText}: ${formatNumber(totals.bearingLength)} m носещи UA профили. Бройките са закръглени към цели профили по ${constants.uaLength} m.${reserve}`;
  }

  if (row.key.includes("bearing-battens")) {
    return `${roomText}: ${formatNumber(totals.bearingLength)} m носещи летви в ${totals.bearingRows} реда. Бройките са закръглени към цели летви по ${constants.woodBattenLength} m.${reserve}`;
  }

  if (row.key.includes("mounting-battens")) {
    return `${roomText}: ${formatNumber(totals.mountingLength)} m монтажни летви в ${totals.mountingRows} реда. Бройките са закръглени към цели летви по ${constants.woodBattenLength} m.${reserve}`;
  }

  if (row.key.includes("custom-bearing")) {
    return `${roomText}: ${formatNumber(totals.bearingLength)} m носещи редове. Бройките са закръглени според избраната дължина на носещия профил.${reserve}`;
  }

  if (row.key.includes("custom-mounting")) {
    return `${roomText}: ${formatNumber(totals.mountingLength)} m монтажни редове. Бройките са закръглени според избраната дължина на монтажния профил.${reserve}`;
  }

  if (row.key.includes("ud-28-27") || row.key.includes("custom-perimeter")) {
    return `${roomText}: ${formatNumber(totals.udLength)} m по периметъра на помещенията. Бройките са закръглени към цели UD профили по ${constants.udLength} m.${reserve}`;
  }

  if (row.key.includes("connectors") || row.key.includes("cross-points")) {
    return `${roomText}: ${totals.bearingRows} носещи реда и ${totals.mountingRows} монтажни реда. Връзките са пресичанията между тях: общо ${totals.crossConnectors} бр.${reserve}`;
  }

  if (row.key.includes("hangers")) {
    return `${roomText}: окачвачите са по носещите редове през разстояние a. Общо по геометрия: ${totals.hangers} бр.${reserve}`;
  }

  if (row.key.includes("anchors-ud") || row.key.includes("perimeter-anchors")) {
    return `${roomText}: дюбели по периметъра ${formatNumber(totals.udLength)} m, средна стъпка около ${averageUdSpacing} mm. Ъглите не се броят двойно. Общо по геометрия: ${totals.anchorsUd} бр.${reserve}`;
  }

  if (row.key.includes("anchors-hangers")) {
    return `${roomText}: по един дюбел за всеки окачвач. Окачвачи общо: ${totals.hangers} бр.${reserve}`;
  }

  if (row.key.includes("metal-screws")) {
    return `${roomText}: винтове за връзки и окачвачи - ${totals.crossConnectors} връзки x ${constants.metalScrewsPerCrossConnector} бр. плюс ${totals.hangers} окачвачи x ${constants.metalScrewsPerDirectHanger} бр.${reserve}`;
  }

  if (row.key.includes("tn-screws")) {
    return `${roomText}: ${formatNumber(totals.boardArea)} m2 обшивка x ${constants.drywallScrewsPerM2} винта/m2.${reserve}`;
  }

  if (row.key.includes("boards-area")) {
    return `${roomText}: площта е сбор от площта на стаите по броя слоеве гипсокартон: ${formatNumber(totals.boardArea)} m2.${reserve}`;
  }

  if (row.key.includes("boards")) {
    return `${roomText}: ${formatNumber(totals.boardArea)} m2 обшивка. Един лист е ${constants.boardWidth} x ${constants.boardLength} m = ${formatNumber(boardSize)} m2, после се закръгля към цели листове.${reserve}`;
  }

  if (row.key.includes("joint-tape")) {
    return `${roomText}: ${formatNumber(totals.area)} m2 площ x ${constants.jointTapePerM2} m фугопокривна лента/m2.${reserve}`;
  }

  if (row.key.includes("joint-compound")) {
    return `${roomText}: ${formatNumber(totals.area)} m2 площ x ${constants.jointCompoundKgPerM2} kg шпакловка/m2.${reserve}`;
  }

  if (row.key.includes("trenn-fix")) {
    return `${roomText}: ${formatNumber(totals.udLength)} m периметър x коефициент ${constants.trennFixPerimeterMultiplier}.${reserve}`;
  }

  if (row.key.includes("mineral-wool")) {
    return `${roomText}: ватата следва площта на тавана: ${formatNumber(totals.area)} m2, дебелина ${constants.mineralWoolThickness} mm.${reserve}`;
  }

  if (row.key.includes("uw-50-40")) {
    const uaExtensionPoints = materialRooms.reduce((sum, sourceRoom) => {
      const room = cloneRoom(sourceRoom);
      const result = calc(room, constants);
      return sum + result.bearingCount * Math.max(0, Math.ceil((result.L / 100) / constants.uaLength) - 1);
    }, 0);
    return `${roomText}: ${uaExtensionPoints} места за удължаване на UA носещи профили. Приема се по 2 UW елемента на всяко място.${reserve}`;
  }

  if (row.key.includes("extensions")) {
    return `${roomText}: удължители се добавят там, където редът е по-дълъг от стандартния профил. Общо по геометрия: ${totals.extensions} бр.${reserve}`;
  }

  return `${roomText}: количеството е изчислено от геометрията на запазените стаи и избраните настройки.${reserve}`;
}

function renderSvgKeyValueRows(rows: Array<[string, unknown]>, startY: number): { markup: string; height: number } {
  const x = 48;
  const labelW = 210;
  const valueW = 790;
  const rowH = 30;
  const markup = rows.map(([label, value], index) => {
    const y = startY + index * rowH;
    return `<g>
      <rect x="${x}" y="${y}" width="${labelW}" height="${rowH}" class="table-head" />
      <rect x="${x + labelW}" y="${y}" width="${valueW}" height="${rowH}" class="table-cell" />
      <text x="${x + 10}" y="${y + 20}" class="table-label">${escapeHtml(label)}</text>
      <text x="${x + labelW + 10}" y="${y + 20}" class="table-text">${escapeHtml(value)}</text>
    </g>`;
  }).join("");
  return { markup, height: rows.length * rowH };
}

function renderSvgMaterialsRows(rows: MaterialTakeoffItem[], room: Room, constants: CalculatorConstants, startY: number): { markup: string; height: number } {
  const x = 48;
  const nameW = 280;
  const qtyW = 115;
  const noteW = 605;
  const headerH = 32;
  let y = startY;
  let markup = `<g>
    <rect x="${x}" y="${y}" width="${nameW}" height="${headerH}" class="table-head" />
    <rect x="${x + nameW}" y="${y}" width="${qtyW}" height="${headerH}" class="table-head" />
    <rect x="${x + nameW + qtyW}" y="${y}" width="${noteW}" height="${headerH}" class="table-head" />
    <text x="${x + 10}" y="${y + 21}" class="table-label">Материал</text>
    <text x="${x + nameW + 10}" y="${y + 21}" class="table-label">Количество</text>
    <text x="${x + nameW + qtyW + 10}" y="${y + 21}" class="table-label">Как е сметнато</text>
  </g>`;
  y += headerH;

  rows.forEach((row) => {
    const explanation = buildMaterialCalculationInfo(row, [room], constants);
    const name = svgText(x + 10, y + 20, row.label, "table-text", 32, 16);
    const qty = svgText(x + nameW + 10, y + 20, `${row.quantity} ${row.unit}`, "table-text", 14, 16);
    const note = svgText(x + nameW + qtyW + 10, y + 20, explanation, "table-text", 76, 16);
    const rowH = Math.max(34, name.height, qty.height, note.height) + 14;
    markup += `<g>
      <rect x="${x}" y="${y}" width="${nameW}" height="${rowH}" class="table-cell" />
      <rect x="${x + nameW}" y="${y}" width="${qtyW}" height="${rowH}" class="table-cell" />
      <rect x="${x + nameW + qtyW}" y="${y}" width="${noteW}" height="${rowH}" class="table-cell" />
      ${name.markup}${qty.markup}${note.markup}
    </g>`;
    y += rowH;
  });

  return { markup, height: y - startY };
}

function renderHtmlMaterialsRows(rows: MaterialTakeoffItem[], room: Room, constants: CalculatorConstants): string {
  return rows.map((row) => `<tr>
    <td>${escapeHtml(row.label)}</td>
    <td>${escapeHtml(row.quantity)} ${escapeHtml(row.unit)}</td>
    <td>${escapeHtml(buildMaterialCalculationInfo(row, [room], constants))}</td>
  </tr>`).join("");
}

function renderSvgCutPlan(cutPlan: CutPlanState, result: CalcResult, startY: number): { markup: string; height: number } {
  const x = 48;
  const width = 1000;
  if (cutPlan.error || !cutPlan.plan) {
    const message = svgText(x + 12, startY + 24, `Разкроят не може да се изчисли: ${cutPlan.error || "неизвестна грешка"}.`, "warning-text", 110, 18);
    return {
      markup: `<rect x="${x}" y="${startY}" width="${width}" height="${message.height + 22}" class="warning-box" />${message.markup}`,
      height: message.height + 22,
    };
  }

  const plan = cutPlan.plan;
  const currentProfiles = result.cdTotalProfiles + result.udProfiles;
  let y = startY;
  let markup = "";
  const summary = [
    `Профили за покупка: ${plan.totalBars} бр.`,
    `Просто броене: ${currentProfiles} бр.`,
    `Използвана дължина: ${formatNumber(plan.totalUsedCm)} cm`,
    `Отпадък: ${formatNumber(plan.totalWasteCm)} cm`,
    `Ефективност: ${formatNumber(plan.efficiencyPercent)}%`,
  ];
  markup += `<text x="${x}" y="${y + 18}" class="table-text">${escapeHtml(summary.join("  |  "))}</text>`;
  y += 34;

  plan.bars.forEach((bar) => {
    const barH = 42;
    markup += `<g>
      <text x="${x}" y="${y + 15}" class="table-label">${escapeHtml(bar.id)} - ${escapeHtml(getCutPieceLabel(bar.type))} - ${formatNumber(bar.usedCm)} / ${formatNumber(bar.stockLengthCm)} cm</text>
      <rect x="${x}" y="${y + 21}" width="${width}" height="18" class="cut-bg" />
      ${bar.segments.map((segment) => {
        const sx = x + (segment.startCm / bar.stockLengthCm) * width;
        const sw = Math.max(1, (segment.lengthCm / bar.stockLengthCm) * width);
        const label = Math.round(segment.lengthCm);
        return `<g>
          <rect x="${sx}" y="${y + 21}" width="${sw}" height="18" class="cut-${segment.type}" />
          ${sw >= 28 ? `<text x="${sx + sw / 2}" y="${y + 35}" class="cut-piece-label" text-anchor="middle">${label}</text>` : ""}
        </g>`;
      }).join("")}
      ${bar.wasteCm > 0 ? (() => {
        const wx = x + (bar.usedCm / bar.stockLengthCm) * width;
        const ww = Math.max(1, (bar.wasteCm / bar.stockLengthCm) * width);
        return `<g><rect x="${wx}" y="${y + 21}" width="${ww}" height="18" class="cut-waste" />${ww >= 28 ? `<text x="${wx + ww / 2}" y="${y + 35}" class="cut-waste-label" text-anchor="middle">${Math.round(bar.wasteCm)}</text>` : ""}</g>`;
      })() : ""}
    </g>`;
    y += barH;
  });

  return { markup, height: y - startY };
}

function renderHtmlCutPlan(cutPlan: CutPlanState, result: CalcResult): string {
  if (cutPlan.error || !cutPlan.plan) {
    return `<p class="warning">Разкроят не може да се изчисли: ${escapeHtml(cutPlan.error || "неизвестна грешка")}.</p>`;
  }
  const plan = cutPlan.plan;
  const currentProfiles = result.cdTotalProfiles + result.udProfiles;
  const summaryRows = [
    ["Профили за покупка", plan.totalBars],
    ["Просто броене преди оптимизация", currentProfiles],
    ["Използвана дължина", `${formatNumber(plan.totalUsedCm)} cm`],
    ["Отпадък", `${formatNumber(plan.totalWasteCm)} cm`],
    ["Ефективност", `${formatNumber(plan.efficiencyPercent)} %`],
    ["Теоретичен минимум", `${plan.lowerBoundBars} профила`],
  ];
  return `
    <table><tbody>${summaryRows.map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table>
    <div class="cut-legend"><span class="carrier">Носещ CD</span><span class="mounting">Монтажен CD</span><span class="ud">UD</span><span class="waste">Остатък</span></div>
    <div class="cut-bar-list">
      ${plan.bars.map((bar) => `
        <article class="cut-bar-card">
          <div class="cut-bar-head"><strong>${escapeHtml(bar.id)}</strong><span>${escapeHtml(getCutPieceLabel(bar.type))}</span><small>${formatNumber(bar.usedCm)} / ${formatNumber(bar.stockLengthCm)} cm</small></div>
          <div class="cut-strip">
            ${bar.segments.map((segment) => `<div class="cut-piece ${segment.type}" style="width:${(segment.lengthCm / bar.stockLengthCm) * 100}%">${Math.round(segment.lengthCm)}</div>`).join("")}
            ${bar.wasteCm > 0 ? `<div class="cut-piece waste" style="width:${(bar.wasteCm / bar.stockLengthCm) * 100}%">${Math.round(bar.wasteCm)}</div>` : ""}
          </div>
        </article>`).join("")}
    </div>`;
}

function buildRoomReportSvg(room: Room, constants: CalculatorConstants): string {
  const safeRoom = cloneRoom(room);
  const result = calc(safeRoom, constants);
  const materials = buildMaterialTakeoff([safeRoom], constants);
  const cutPlan = buildSafeCutPlan(safeRoom, result, constants);
  const generatedAt = new Date().toLocaleString("bg-BG");
  const pageW = 1100;
  let y = 42;
  let body = "";

  body += `<text x="48" y="${y}" class="title">${escapeHtml(safeRoom.name)}</text>`;
  y += 28;
  body += `<text x="48" y="${y}" class="meta">Knauf D11 calculator · ${escapeHtml(generatedAt)}</text>`;
  y += 34;

  body += `<text x="48" y="${y}" class="section-title">Данни за стаята</text>`;
  y += 14;
  const roomRows = renderSvgKeyValueRows([
    ["Система", `${safeRoom.systemType} - ${getConstruction(safeRoom).label}`],
    ["Размер X", `${safeRoom.width} cm`],
    ["Размер Y", `${safeRoom.length} cm`],
    ["Площ", `${formatNumber(safeRoom.area)} m2`],
    ["Натоварване", `${safeRoom.loadClass} kN/m2`],
    ["Разстояния", `a ${safeRoom.a} mm, b ${safeRoom.b} mm, c ${safeRoom.c} mm`],
    ["Носещи / монтажни редове", `${result.bearingCount} / ${result.mountingCount}`],
    ["Окачвачи / дюбели / винтове", `${result.hangersTotal} / ${result.anchorsTotal} / ${result.metalScrews + result.drywallScrews}`],
  ], y);
  body += roomRows.markup;
  y += roomRows.height + 34;

  body += `<text x="48" y="${y}" class="section-title">Необходими материали</text>`;
  y += 14;
  const materialRows = renderSvgMaterialsRows(materials, safeRoom, constants, y);
  body += materialRows.markup;
  y += materialRows.height + 34;

  body += `<text x="48" y="${y}" class="section-title">Работна схема</text>`;
  y += 16;
  const scheme = renderWorkingSchemeSvg(safeRoom, result, constants).replace("<svg ", `<svg x="48" y="${y}" width="1000" height="572" `);
  body += scheme;
  y += 602;

  body += `<text x="48" y="${y}" class="section-title">Разкрой за стаята</text>`;
  y += 16;
  const cut = renderSvgCutPlan(cutPlan, result, y);
  body += cut.markup;
  y += cut.height + 48;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${pageW}" height="${y}" viewBox="0 0 ${pageW} ${y}">
    <style>
      .page-bg { fill: #ffffff; }
      .title { font: 700 28px Arial, sans-serif; fill: #17211f; }
      .meta { font: 13px Arial, sans-serif; fill: #5f6f6a; }
      .section-title { font: 700 18px Arial, sans-serif; fill: #0f766e; }
      .table-head { fill: #edf5f1; stroke: #d7e0e3; }
      .table-cell { fill: #ffffff; stroke: #d7e0e3; }
      .table-label { font: 700 13px Arial, sans-serif; fill: #34413c; }
      .table-text { font: 13px Arial, sans-serif; fill: #17211f; }
      .axis-label, .position-label, .hanger-label, .extension-label { font: 12px Arial, sans-serif; fill: #263633; }
      .bearing-line { stroke: #0f766e; stroke-width: 3; }
      .mounting-line { stroke: #f59e0b; stroke-width: 2; }
      .hanger-dot { fill: #111827; }
      .hanger-dimension-line { stroke: #111827; stroke-width: 1; stroke-dasharray: 4 4; }
      .extension-dimension-line, .extension-mark { stroke: #dc2626; stroke-width: 2; }
      .extension-label { fill: #991b1b; }
      .cut-bg, .cut-waste { fill: #d1d5db; }
      .cut-carrier { fill: #0f766e; }
      .cut-mounting { fill: #f59e0b; }
      .cut-ud { fill: #2563eb; }
      .cut-piece-label { font: 700 11px Arial, sans-serif; fill: #ffffff; }
      .cut-mounting + .cut-piece-label, .cut-waste-label { font: 700 11px Arial, sans-serif; fill: #374151; }
      .warning-box { fill: #fff4f4; stroke: #f3c5c5; }
      .warning-text { font: 13px Arial, sans-serif; fill: #991b1b; }
    </style>
    <rect x="0" y="0" width="${pageW}" height="${y}" class="page-bg" />
    ${body}
  </svg>`;
}

function buildRoomReportHtml(room: Room, constants: CalculatorConstants): string {
  const safeRoom = cloneRoom(room);
  const result = calc(safeRoom, constants);
  const materials = buildMaterialTakeoff([safeRoom], constants);
  const cutPlan = buildSafeCutPlan(safeRoom, result, constants);
  const generatedAt = new Date().toLocaleString("bg-BG");
  const roomRows: Array<[string, unknown]> = [
    ["Система", `${safeRoom.systemType} - ${getConstruction(safeRoom).label}`],
    ["Размер X", `${safeRoom.width} cm`],
    ["Размер Y", `${safeRoom.length} cm`],
    ["Площ", `${formatNumber(safeRoom.area)} m2`],
    ["Натоварване", `${safeRoom.loadClass} kN/m2`],
    ["Разстояния", `a ${safeRoom.a} mm, b ${safeRoom.b} mm, c ${safeRoom.c} mm`],
    ["Носещи / монтажни редове", `${result.bearingCount} / ${result.mountingCount}`],
    ["Окачвачи / дюбели / винтове", `${result.hangersTotal} / ${result.anchorsTotal} / ${result.metalScrews + result.drywallScrews}`],
  ];

  return `<!doctype html>
<html lang="bg">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(safeRoom.name)} - Knauf calculator</title>
  <style>
    body { margin: 0; padding: 28px; font-family: Arial, sans-serif; color: #17211f; background: #ffffff; }
    main { max-width: 1120px; margin: 0 auto; }
    h1 { margin: 0 0 6px; font-size: 26px; }
    h2 { margin: 28px 0 10px; font-size: 18px; color: #0f766e; }
    .meta { color: #5f6f6a; margin-bottom: 22px; }
    table { width: 100%; border-collapse: collapse; margin: 8px 0 18px; font-size: 13px; }
    th, td { border: 1px solid #d7e0e3; padding: 7px 8px; text-align: left; vertical-align: top; }
    th { width: 28%; background: #f6f8f4; }
    thead th { width: auto; background: #edf5f1; }
    .ceiling-svg { width: 100%; height: auto; border: 1px solid #d7e0e3; background: #f6f8f4; }
    .axis-label, .position-label, .hanger-label, .extension-label { font: 12px Arial, sans-serif; fill: #263633; }
    .bearing-line { stroke: #0f766e; stroke-width: 3; }
    .mounting-line { stroke: #f59e0b; stroke-width: 2; }
    .hanger-dot { fill: #111827; }
    .hanger-dimension-line { stroke: #111827; stroke-width: 1; stroke-dasharray: 4 4; }
    .extension-dimension-line, .extension-mark { stroke: #dc2626; stroke-width: 2; }
    .extension-label { fill: #991b1b; }
    .cut-legend { display: flex; flex-wrap: wrap; gap: 8px; margin: 10px 0 12px; font-size: 12px; }
    .cut-legend span { padding: 4px 8px; border-radius: 999px; color: #fff; }
    .cut-legend .carrier, .cut-piece.carrier { background: #0f766e; }
    .cut-legend .mounting, .cut-piece.mounting { background: #f59e0b; color: #17211f; }
    .cut-legend .ud, .cut-piece.ud { background: #2563eb; }
    .cut-legend .waste, .cut-piece.waste { background: #d1d5db; color: #374151; }
    .cut-bar-list { display: grid; gap: 9px; }
    .cut-bar-card { break-inside: avoid; border: 1px solid #d7e0e3; border-radius: 6px; padding: 8px; }
    .cut-bar-head { display: flex; gap: 10px; justify-content: space-between; font-size: 12px; margin-bottom: 6px; }
    .cut-strip { display: flex; height: 28px; overflow: hidden; border-radius: 4px; background: #eef2f7; }
    .cut-piece { min-width: 24px; display: grid; place-items: center; color: #fff; font-size: 11px; border-right: 1px solid rgba(255,255,255,0.7); }
    .warning { padding: 10px; border: 1px solid #f3c5c5; background: #fff4f4; color: #991b1b; }
    @media print { body { padding: 12mm; } h2 { break-after: avoid; } .ceiling-svg, .cut-bar-card { break-inside: avoid; } }
  </style>
</head>
<body>
  <main>
    <h1>${escapeHtml(safeRoom.name)}</h1>
    <div class="meta">Knauf D11 calculator · ${escapeHtml(generatedAt)}</div>
    <h2>Данни за стаята</h2>
    <table><tbody>${roomRows.map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table>
    <h2>Необходими материали</h2>
    <table>
      <thead><tr><th>Материал</th><th>Количество</th><th>Как е сметнато</th></tr></thead>
      <tbody>${renderHtmlMaterialsRows(materials, safeRoom, constants)}</tbody>
    </table>
    <h2>Работна схема</h2>
    ${renderWorkingSchemeSvg(safeRoom, result, constants)}
    <h2>Разкрой за стаята</h2>
    ${renderHtmlCutPlan(cutPlan, result)}
  </main>
</body>
</html>`;
}

async function buildRoomReportCanvas(room: Room, constants: CalculatorConstants): Promise<HTMLCanvasElement> {
  const svg = buildRoomReportSvg(room, constants);
  const image = new Image();
  image.decoding = "async";
  const loaded = new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error("PNG export image failed to load."));
  });
  image.src = svgDataUrl(svg);
  await loaded;
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth || 1100;
  canvas.height = image.naturalHeight || 1600;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas export is not supported.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0);
  return canvas;
}

function canvasToBlob(canvas: HTMLCanvasElement, type: string, quality?: number): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("Export failed."));
    }, type, quality);
  });
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function escapePdfText(value: string): string {
  return value.replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function buildPdfFromCanvas(canvas: HTMLCanvasElement, title: string): Blob {
  const jpeg = base64ToBytes(canvas.toDataURL("image/jpeg", 0.92).split(",")[1] ?? "");
  const width = canvas.width;
  const height = canvas.height;
  const content = `q\n${width} 0 0 ${height} 0 0 cm\n/Im0 Do\nQ\n`;
  const encoder = new TextEncoder();
  const parts: BlobPart[] = [];
  const offsets: number[] = [];
  let byteLength = 0;
  const pushString = (value: string) => {
    parts.push(value);
    byteLength += encoder.encode(value).length;
  };
  const pushBytes = (value: Uint8Array) => {
    parts.push(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength) as ArrayBuffer);
    byteLength += value.length;
  };
  const beginObject = (id: number) => {
    offsets[id] = byteLength;
    pushString(`${id} 0 obj\n`);
  };

  pushString("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
  beginObject(1);
  pushString(`<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`);
  beginObject(2);
  pushString(`<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n`);
  beginObject(3);
  pushString(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`);
  beginObject(4);
  pushString(`<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg.length} >>\nstream\n`);
  pushBytes(jpeg);
  pushString("\nendstream\nendobj\n");
  beginObject(5);
  pushString(`<< /Length ${encoder.encode(content).length} >>\nstream\n${content}endstream\nendobj\n`);
  const xrefOffset = byteLength;
  pushString(`xref\n0 6\n0000000000 65535 f \n`);
  for (let id = 1; id <= 5; id += 1) pushString(`${String(offsets[id]).padStart(10, "0")} 00000 n \n`);
  pushString(`trailer\n<< /Size 6 /Root 1 0 R /Info << /Title (${escapePdfText(title)}) >> >>\nstartxref\n${xrefOffset}\n%%EOF`);
  return new Blob(parts, { type: "application/pdf" });
}

async function buildRoomReportBlob(room: Room, constants: CalculatorConstants, fileType: RoomCardExportFileType): Promise<Blob> {
  if (fileType === "html") {
    return new Blob([buildRoomReportHtml(room, constants)], { type: "text/html;charset=utf-8" });
  }
  const canvas = await buildRoomReportCanvas(room, constants);
  return fileType === "pdf"
    ? buildPdfFromCanvas(canvas, room.name || "Knauf calculator")
    : canvasToBlob(canvas, "image/png");
}

function getExportTypeDescription(contentType: ExportContentType, fileType: ExportFileType): string {
  if (fileType === "pdf") {
    return "PDF export създава отделен файл за всяка стая. Всеки файл съдържа данни за стаята, необходимите материали с обяснение, работната схема и разкроя с дължините на отделните парчета.";
  }
  if (fileType === "png") {
    return "PNG export създава отделна картинка за всяка стая със същия отчет: данни, материали, работна схема и разкрой. Подходящо е за бързо споделяне като изображение.";
  }
  if (fileType === "html") {
    return contentType === "room-cards"
      ? "HTML export създава отделен отваряем файл за всяка стая. Файлът може да се преглежда в браузър и да се печата, като съдържа таблици, SVG работна схема и подробен разкрой по пръти."
      : "HTML export създава един файл rooms.html с таблица на всички стаи. Подходящ е за преглед в браузър или бърз печат.";
  }
  if (fileType === "json") {
    return "JSON export създава един файл rooms.json със запазените стаи, активната стая, глобалните настройки и настройките на проекта. Този файл може после да се върне през Импорт JSON.";
  }
  return "Excel export създава един файл rooms.xls с таблица на всички стаи и основните изчислени стойности от grid-а.";
}

function getExportContentDescription(contentType: ExportContentType): string {
  return contentType === "room-cards"
    ? "Ще бъдат създадени отделни файлове за всяка стая с размери, материали, схема и разкрой."
    : "Ще бъде създаден един файл с всички стаи и техните данни.";
}

function getExportFormatOptions(contentType: ExportContentType): Array<{ value: ExportFileType; label: string }> {
  return contentType === "room-cards"
    ? [
      { value: "pdf", label: "PDF" },
      { value: "png", label: "PNG" },
      { value: "html", label: "HTML" },
    ]
    : [
      { value: "excel", label: "Excel" },
      { value: "json", label: "JSON" },
      { value: "html", label: "HTML" },
    ];
}

function isRoomCardFileType(fileType: ExportFileType): fileType is RoomCardExportFileType {
  return fileType === "pdf" || fileType === "png" || fileType === "html";
}

const CUSTOM_A_OPTIONS = buildRange(100, 3000, 50);
const CUSTOM_B_OPTIONS = buildRange(100, 1250, 25);
const CUSTOM_C_OPTIONS = buildRange(100, 3000, 50);
const CUSTOM_UD_ANCHOR_OPTIONS = buildRange(100, 1500, 25);

function buildRange(min: number, max: number, step: number): number[] {
  const values: number[] = [];
  for (let value = min; value <= max; value += step) values.push(value);
  return values;
}

function syncDistanceOverride(room: Room, key: "a" | "b" | "c" | "udAnchorSpacing"): void {
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType, room.d116Variant, room.d112Variant);
  const expected = {
    a: getTableValue(room) ?? auto.a,
    b: room.boardType === "custom" ? Number.NaN : auto.b,
    c: auto.c,
    udAnchorSpacing: auto.udAnchorSpacing,
  }[key];

  room.overrides[key] = Number(room[key]) !== Number(expected);
}

function App() {
  const [state, setState] = useState<AppState>(() => loadState());
  const [zoom, setZoom] = useState(1);
  const [saveStatus, setSaveStatus] = useState("");
  const [exportDirectoryName, setExportDirectoryName] = useState(DEFAULT_EXPORT_DIRECTORY_NAME);
  const [exportContentType, setExportContentType] = useState<ExportContentType>("room-cards");
  const [exportFileType, setExportFileType] = useState<ExportFileType>("pdf");
  const [exportDirectoryHandle, setExportDirectoryHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [isDeleteAllModalOpen, setIsDeleteAllModalOpen] = useState(false);
  const [exportStatus, setExportStatus] = useState("");
  const [activeSection, setActiveSection] = useState<AppSection>("room");
  const [roomWorkspacePanel, setRoomWorkspacePanel] = useState<RoomWorkspacePanel>("visual");

  const activeRoom = state.draftRoom;
  const activeResult = calc(activeRoom, state.constants);
  const activeCutPlan = buildSafeCutPlan(activeRoom, activeResult, state.constants);
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

  function confirmDeleteAllRooms(): void {
    deleteAllRooms();
    setIsDeleteAllModalOpen(false);
  }

  function recalculateSavedRooms(): void {
    commit((draft) => {
      draft.rooms = draft.rooms.map((room) => recalculateRoomWithCurrentLogic(room, draft.constants));
      const activeRoom = draft.rooms.find((room) => room.id === draft.activeRoomId);
      if (activeRoom) {
        draft.draftRoom = cloneRoom(activeRoom);
      }
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
        room.overrides.udAnchorSpacing = true;
      }
    });
  }

  function resetAutoSpacing(): void {
    updateActiveRoom((room) => {
      room.overrides.a = false;
      room.overrides.b = false;
      room.overrides.c = false;
      room.overrides.udAnchorSpacing = false;
      syncSpacingFromKnaufTable(room, { keepC: false });
    });
  }

  function downloadBlob(blob: Blob, filename: string): void {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function buildAppJsonBlob(): Blob {
    return new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  }

  function buildRoomsTableRows(): unknown[][] {
    const headers = [
      "Стая", "Конструкция", "X", "Y", "Площ", "Носещи реда", "Монтажни реда",
      "Носещи m", "Монтажни m", "Профили общо", "UD бр.", "Връзки", "Окачвачи",
      "Дюбели UD", "Дюбели окачвачи", "Дюбели общо", "Винтове метал", "Винтове гипсокартон", "Удължители",
    ];
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
    return [headers, ...rows];
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

  function buildRoomsExcelBlob(): Blob {
    const escapeXml = (value: unknown) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    const sheetRows = buildRoomsTableRows()
      .map((row) => `<Row>${row.map((cell) => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join("")}</Row>`)
      .join("");
    const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Стаи"><Table>${sheetRows}</Table></Worksheet></Workbook>`;
    return new Blob([workbook], { type: "application/vnd.ms-excel" });
  }

  function buildRoomsHtmlBlob(): Blob {
    const rows = buildRoomsTableRows();
    const [headers, ...bodyRows] = rows;
    const html = `<!doctype html><html lang="bg"><head><meta charset="utf-8" /><meta name="viewport" content="width=device-width, initial-scale=1" /><title>rooms</title><style>body{font-family:Arial,sans-serif;margin:0;padding:24px;color:#17211f}table{border-collapse:collapse;width:100%;font-size:13px}th,td{border:1px solid #d7e0e3;padding:7px 8px;text-align:left;white-space:nowrap}th{background:#edf5f1}h1{font-size:24px;margin:0 0 16px}</style></head><body><h1>Таблица със стаи</h1><table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${bodyRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
    return new Blob([html], { type: "text/html;charset=utf-8" });
  }

  function exportMaterialsExcel(): void {
    const escapeXml = (value: unknown) => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    const headers = ["Материал", "Количество", "Ед.", "Как е сметнато"];
    const rows = buildMaterialTakeoff(state.rooms, state.constants).map((row) => [
      row.label,
      row.quantity,
      row.unit,
      buildMaterialCalculationInfo(row, state.rooms, state.constants),
    ]);
    const sheetRows = [headers, ...rows]
      .map((row) => `<Row>${row.map((cell) => `<Cell><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join("")}</Row>`)
      .join("");
    const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Материали"><Table>${sheetRows}</Table></Worksheet></Workbook>`;
    downloadBlob(new Blob([workbook], { type: "application/vnd.ms-excel" }), "suspended-ceiling-materials-only.xls");
  }

  async function browseExportDestination(): Promise<void> {
    if (!window.showDirectoryPicker) {
      setExportStatus("Браузърът не поддържа избор на директория. При експорт файловете ще се свалят поотделно.");
      return;
    }
    try {
      const handle = await window.showDirectoryPicker();
      setExportDirectoryHandle(handle);
      setExportStatus(`Избрана дестинация: ${handle.name}`);
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") setExportStatus("Изборът на директория е отказан.");
      else setExportStatus("Неуспешен избор на директория.");
    }
  }

  function changeExportContentType(contentType: ExportContentType): void {
    setExportContentType(contentType);
    setExportFileType(contentType === "room-cards" ? "pdf" : "excel");
  }

  async function exportRoomReports(): Promise<void> {
    if (!state.rooms.length) return;
    if (window.showDirectoryPicker && !exportDirectoryHandle) {
      setExportStatus("Първо избери дестинация с Browse.");
      return;
    }
    const directoryName = exportDirectoryName.trim() || DEFAULT_EXPORT_DIRECTORY_NAME;
    const filenameCounts = new Map<string, number>();
    const reports = exportContentType === "table"
      ? [{
        filename: exportFileType === "json" ? "rooms.json" : exportFileType === "html" ? "rooms.html" : "rooms.xls",
        blob: exportFileType === "json" ? buildAppJsonBlob() : exportFileType === "html" ? buildRoomsHtmlBlob() : buildRoomsExcelBlob(),
      }]
      : await Promise.all(state.rooms.map(async (room) => {
          const roomCardFileType = isRoomCardFileType(exportFileType) ? exportFileType : "pdf";
          const baseName = sanitizeFilename(room.name || room.systemType);
          const count = filenameCounts.get(baseName) ?? 0;
          filenameCounts.set(baseName, count + 1);
          const filename = `${baseName}${count ? `-${count + 1}` : ""}.${roomCardFileType}`;
          return {
            filename,
            blob: await buildRoomReportBlob(room, state.constants, roomCardFileType),
          };
        }));

    try {
      if (window.showDirectoryPicker && exportDirectoryHandle) {
        const target = await exportDirectoryHandle.getDirectoryHandle(directoryName, { create: true });
        for (const report of reports) {
          const file = await target.getFileHandle(report.filename, { create: true });
          const writable = await file.createWritable();
          await writable.write(report.blob);
          await writable.close();
        }
        setExportStatus(`Експортирани ${reports.length} файла в "${directoryName}".`);
        setIsExportModalOpen(false);
      } else {
        reports.forEach((report) => {
          downloadBlob(report.blob, report.filename);
        });
        setExportStatus("Браузърът не поддържа избор на директория. Файловете са свалени поотделно.");
        setIsExportModalOpen(false);
      }
      window.setTimeout(() => setExportStatus(""), 4000);
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        setExportStatus("Експортът е отказан.");
      } else {
        setExportStatus("Неуспешен експорт на файловете.");
      }
      window.setTimeout(() => setExportStatus(""), 4000);
    }
  }

  const loadClasses = getLoadClasses(activeRoom.systemType, activeRoom.fireProtection, activeRoom.d116Variant, activeRoom.d112Variant);
  const menuItems: Array<{ key: AppSection; label: string }> = [
    { key: "room", label: "Стаи" },
    { key: "settings", label: "Глобални настройки" },
    { key: "materials", label: "Общо Материали" },
    { key: "help", label: "Help" },
  ];

  return (
    <main className="app-shell">
      <section className="workbench">
        <div className="topbar">
          <div>
            <p className="eyebrow">Knauf D11 calculator</p>
            <h1>Калкулатор за окачени тавани</h1>
          </div>
          <div className="topbar-actions">
            {menuItems.map((item) => (
              <button
                key={item.key}
                type="button"
                className={activeSection === item.key ? "menu-button active" : "menu-button"}
                onClick={() => setActiveSection(item.key)}
              >
                {item.label}
              </button>
            ))}
          </div>
        </div>

        {activeSection === "room" && (
          <section className="section-stack">
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
                  onAddRoom={() => {
                    addRoom();
                    setActiveSection("room");
                    setRoomWorkspacePanel("visual");
                  }}
                  onRoomChange={updateActiveRoom}
                />
              </aside>

              <section className="visual-workspace">
                <ResultCards result={activeResult} room={activeRoom} />
                {roomWorkspacePanel === "visual" ? (
                  <Visualization
                    room={activeRoom}
                    result={activeResult}
                    constants={state.constants}
                    zoom={zoom}
                    onZoomChange={setZoom}
                    onOpenCutOptimization={() => setRoomWorkspacePanel("cut")}
                  />
                ) : (
                  <CutOptimizationPanel
                    room={activeRoom}
                    result={activeResult}
                    cutPlan={activeCutPlan}
                    constants={state.constants}
                    onBackToVisualization={() => setRoomWorkspacePanel("visual")}
                  />
                )}
              </section>
            </div>

            <RoomsTable
              rooms={state.rooms}
              constants={state.constants}
              activeRoomId={state.activeRoomId}
              onAddRoom={() => {
                addRoom();
                setActiveSection("room");
                setRoomWorkspacePanel("visual");
              }}
              onSelect={(roomId) => commit((draft) => {
                const room = draft.rooms.find((item) => item.id === roomId);
                if (!room) return;
                draft.draftRoom = cloneRoom(room);
                draft.activeRoomId = room.id;
                setActiveSection("room");
                setRoomWorkspacePanel("visual");
              })}
              onDelete={deleteRoom}
              onDeleteAll={() => setIsDeleteAllModalOpen(true)}
              onRecalculate={recalculateSavedRooms}
              onOpenExport={() => setIsExportModalOpen(true)}
              onImportJson={importJson}
            />
          </section>
        )}

        {activeSection === "settings" && (
          <ConstantsEditor
            constants={state.constants}
            onChange={(patch) => commit((draft) => { draft.constants = { ...draft.constants, ...patch }; })}
          />
        )}

        {activeSection === "materials" && (
          <MaterialsPanel
            rooms={state.rooms}
            constants={state.constants}
            onReserveChange={(wastePercent) => commit((draft) => { draft.constants.wastePercent = wastePercent; })}
            onExportExcel={exportMaterialsExcel}
          />
        )}

        {activeSection === "help" && <HelpPanel />}
      </section>
      {isExportModalOpen && (
        <ExportModal
          directoryName={exportDirectoryName}
          contentType={exportContentType}
          fileType={exportFileType}
          destinationName={exportDirectoryHandle?.name ?? ""}
          status={exportStatus}
          canBrowse={Boolean(window.showDirectoryPicker)}
          canExport={Boolean(state.rooms.length)}
          onDirectoryNameChange={setExportDirectoryName}
          onContentTypeChange={changeExportContentType}
          onFileTypeChange={setExportFileType}
          onBrowse={browseExportDestination}
          onCancel={() => setIsExportModalOpen(false)}
          onConfirm={exportRoomReports}
        />
      )}
      {isDeleteAllModalOpen && (
        <ConfirmDeleteAllModal
          roomCount={state.rooms.length}
          onCancel={() => setIsDeleteAllModalOpen(false)}
          onConfirm={confirmDeleteAllRooms}
        />
      )}
    </main>
  );
}

function ExportModal({ directoryName, contentType, fileType, destinationName, status, canBrowse, canExport, onDirectoryNameChange, onContentTypeChange, onFileTypeChange, onBrowse, onCancel, onConfirm }: {
  directoryName: string;
  contentType: ExportContentType;
  fileType: ExportFileType;
  destinationName: string;
  status: string;
  canBrowse: boolean;
  canExport: boolean;
  onDirectoryNameChange: (value: string) => void;
  onContentTypeChange: (value: ExportContentType) => void;
  onFileTypeChange: (value: ExportFileType) => void;
  onBrowse: () => void;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  const contentDescription = getExportContentDescription(contentType);
  const exportDescription = getExportTypeDescription(contentType, fileType);
  const formatOptions = getExportFormatOptions(contentType);
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="panel export-modal" role="dialog" aria-modal="true" aria-labelledby="export-modal-title">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Експорт по стаи</p>
            <h2 id="export-modal-title">Настройки за експорт</h2>
          </div>
          <button type="button" className="ghost small" onClick={onCancel}>Затвори</button>
        </div>
        <div className="field-grid">
          <label>Какво искаш да експортираш?
            <select value={contentType} onChange={(event) => onContentTypeChange(event.target.value as ExportContentType)}>
              <option value="room-cards">Карти за стаи (схема + материали + разкрой)</option>
              <option value="table">Таблица със стаи (данни / grid)</option>
            </select>
          </label>
          <label>Тип файлове
            <select value={fileType} onChange={(event) => onFileTypeChange(event.target.value as ExportFileType)}>
              {formatOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
            </select>
          </label>
        </div>
        <p className="export-description">{contentDescription}</p>
        <p className="export-description">{exportDescription}</p>
        <div className="destination-row">
          <button type="button" className="ghost" onClick={onBrowse} disabled={!canBrowse}>Browse</button>
          <span>Папка за експорт: {canBrowse ? destinationName || "не е избрана" : "браузърът ще свали файловете поотделно"}</span>
        </div>
        <label>Име на подпапка
          <input value={directoryName} onChange={(event) => onDirectoryNameChange(event.target.value)} />
        </label>
        {status ? <small className="export-status">{status}</small> : null}
        <div className="modal-actions">
          <button type="button" className="ghost" onClick={onCancel}>Отказ</button>
          <button type="button" className="workspace-toggle-button" disabled={!canExport} onClick={onConfirm}>Експорт</button>
        </div>
      </section>
    </div>
  );
}

function ConfirmDeleteAllModal({ roomCount, onCancel, onConfirm }: {
  roomCount: number;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <div className="modal-backdrop" role="presentation">
      <section className="panel confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-all-title">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Потвърждение</p>
            <h2 id="delete-all-title">Изтриване на всички стаи</h2>
          </div>
        </div>
        <div className="danger-note">
          Ще бъдат изтрити всички запазени стаи ({roomCount} бр.). Увери се, че наистина искаш да продължиш.
          Добре е преди това да направиш export, особено JSON export, за да можеш при нужда да ги импортираш отново.
        </div>
        <div className="modal-actions">
          <button type="button" className="ghost" onClick={onCancel}>Отказ</button>
          <button type="button" className="danger" onClick={onConfirm}>Изтрий всички</button>
        </div>
      </section>
    </div>
  );
}

function CutOptimizationPanel({ room, result, cutPlan, constants, onBackToVisualization }: {
  room: Room;
  result: CalcResult;
  cutPlan: CutPlanState;
  constants: CalculatorConstants;
  onBackToVisualization?: () => void;
}) {
  if (cutPlan.error || !cutPlan.plan) {
    return (
      <section className="panel cut-plan-panel">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Оптимизация на разкроя</p>
            <h2>Разкрой за {room.name}</h2>
          </div>
          {onBackToVisualization ? <button type="button" className="workspace-toggle-button" onClick={onBackToVisualization}>Работна схема</button> : null}
        </div>
        <div className="validation-item error">
          Разкроят не може да се изчисли с текущия layout и зададените дължини на профилите.
          {" "}
          {cutPlan.error}
        </div>
      </section>
    );
  }

  const plan = cutPlan.plan;
  const currentProfiles = result.cdTotalProfiles + result.udProfiles;
  const savedBars = Math.max(0, currentProfiles - plan.totalBars);
  const savedPercent = currentProfiles > 0
    ? Number(((savedBars / currentProfiles) * 100).toFixed(2))
    : 0;
  const totalPurchasedLengthCm = plan.bars.reduce((sum, bar) => sum + bar.stockLengthCm, 0);
  const stockLengthCm = plan.bars[0]?.stockLengthCm ?? Math.max(constants.cdLength, constants.udLength) * 100;
  const summaryText = savedBars > 0
    ? `С този разкрой за ${room.name} спестяваш ${savedBars} стандартни профила (${formatNumber(savedPercent)} %) спрямо простото броене ${currentProfiles} бр., защото парчетата от носещи CD, монтажни CD и UD се комбинират по-ефективно в едни и същи пръти.`
    : `За ${room.name} този разкрой не намалява броя нужни профили спрямо простото броене ${currentProfiles} бр., но подрежда парчетата в реален план за рязане и показва точно какъв остатък остава след разкроя.`;

  return (
    <section className="panel cut-plan-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Оптимизация на разкроя</p>
          <h2>Разкрой за {room.name}</h2>
        </div>
        {onBackToVisualization ? <button type="button" className="workspace-toggle-button" onClick={onBackToVisualization}>Работна схема</button> : null}
        <div className="cut-plan-summary-chip">
          <strong>{plan.totalBars} профила</strong>
          <small>стандартен прът {formatNumber(stockLengthCm)} cm, срез {0.3} cm, остатък над {20} cm се счита за използваем</small>
        </div>
      </div>

      <div className="cut-plan-metrics">
        <div className="cut-metric">
          <span>Общо профили</span>
          <strong>{plan.totalBars}</strong>
          <small>Това е колко стандартни пръта трябва да купиш за този разкрой.</small>
          <small>Сравнение с текущото просто броене: {currentProfiles} профила.</small>
        </div>
        <div className="cut-metric">
          <span>Отпадък</span>
          <strong>{formatNumber(plan.totalWasteCm)} cm</strong>
          <small>Това е общият неизползван материал след разкроя на всички закупени пръти.</small>
          <small>{formatNumber(100 - plan.efficiencyPercent)} % от общо закупената дължина остава неизползвана.</small>
        </div>
        <div className="cut-metric">
          <span>Ефективност</span>
          <strong>{formatNumber(plan.efficiencyPercent)} %</strong>
          <small>Показва колко добре използваш общата дължина на закупените пръти.</small>
          <small>Теоретичен минимум: {plan.lowerBoundBars} профила при идеален разкрой без практически ограничения.</small>
        </div>
      </div>

      <div className="cut-plan-type-grid">
        <TypeStatCard label="Носещи CD" stats={plan.carrierStats} />
        <TypeStatCard label="Монтажни CD" stats={plan.mountingStats} />
        <TypeStatCard label="UD" stats={plan.udStats} />
      </div>

      <div className="cut-plan-summary-note">
        <strong>Обобщение за стаята</strong>
        <p>{summaryText}</p>
        <small>
          Общата закупена дължина тук е {formatNumber(totalPurchasedLengthCm)} cm, от които {formatNumber(plan.totalUsedCm)} cm
          {" "}се използват, а {formatNumber(plan.totalWasteCm)} cm остават като остатък.
        </small>
      </div>

      <div className="cut-plan-legend">
        <div className="cut-plan-legend-head">
          <h3>Легенда</h3>
          <small>Оцветяването показва вида на всяко парче в пръта</small>
        </div>
        <div className="cut-plan-legend-grid">
          <LegendItem type="carrier" label="Носещ CD профил" />
          <LegendItem type="mounting" label="Монтажен CD профил" />
          <LegendItem type="ud" label="UD периферен профил" />
          <LegendItem type="waste" label="Остатък / отпадък" />
        </div>
      </div>

      <div className="cut-bar-groups">
        <CutBarGroup title="Разкрой по пръти" bars={plan.bars} />
      </div>
    </section>
  );
}

function getCutPieceLabel(type: CutBar["type"] | "waste"): string {
  if (type === "carrier") return "носещ CD";
  if (type === "mounting") return "монтажен CD";
  if (type === "ud") return "UD";
  if (type === "mixed") return "смесен прът";
  return "остатък";
}

function TypeStatCard({ label, stats }: { label: string; stats: CutOptimizationResult["carrierStats"] }) {
  return (
    <div className="cut-type-card">
      <span>{label}</span>
      <strong>{formatNumber(stats.totalUsedCm)} cm</strong>
      <small>Това е общата използвана дължина от този вид профил в разкроя.</small>
      <small>Участва в {stats.totalBars} пръта; разпределен остатък {formatNumber(stats.totalWasteCm)} cm.</small>
      <em>Ефективност за този тип: {formatNumber(stats.efficiencyPercent)} %.</em>
    </div>
  );
}

function LegendItem({ type, label }: { type: "carrier" | "mounting" | "ud" | "waste"; label: string }) {
  return (
    <div className="cut-legend-item">
      <span className={`cut-legend-swatch ${type}`} />
      <span>{label}</span>
    </div>
  );
}

function CutBarGroup({ title, bars }: { title: string; bars: CutBar[] }) {
  if (!bars.length) return null;
  return (
    <section className="cut-bar-group">
      <div className="cut-bar-group-head">
        <h3>{title}</h3>
        <small>{bars.length} пръта</small>
      </div>
      <div className="cut-bar-list">
        {bars.map((bar) => <CutBarStrip key={bar.id} bar={bar} />)}
      </div>
    </section>
  );
}

function CutBarStrip({ bar }: { bar: CutBar }) {
  const barLabel = getCutPieceLabel(bar.type);
  return (
    <div className="cut-bar-card">
      <div className="cut-bar-head">
        <strong>{bar.id}</strong>
        <span>{barLabel}</span>
        <small>използвани {formatNumber(bar.usedCm)} / {formatNumber(bar.stockLengthCm)} cm</small>
      </div>
      <div className="cut-strip">
        {bar.segments.map((segment) => (
          <div
            key={segment.pieceId}
            className={`cut-piece ${segment.type}`}
            style={{ width: `${(segment.lengthCm / bar.stockLengthCm) * 100}%` }}
            title={`${getCutPieceLabel(segment.type)}: ${segment.lengthCm} cm`}
          >
            {Math.round(segment.lengthCm)}
          </div>
        ))}
        {bar.wasteCm > 0 && (
          <div
            className="cut-piece waste"
            style={{ width: `${(bar.wasteCm / bar.stockLengthCm) * 100}%` }}
            title={`Отпадък ${bar.wasteCm} cm`}
          >
            {Math.round(bar.wasteCm)}
          </div>
        )}
      </div>
    </div>
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
  onAddRoom: () => void;
  onRoomChange: (updater: (room: Room) => void) => void;
}

function RoomEditor({ room, loadClasses, isValid, warnings, onSystemChange, onResetAuto, onSave, saveStatus, onAddRoom, onRoomChange }: RoomEditorProps) {
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
          <button type="button" className="ghost small" onClick={onAddRoom}>Нова стая</button>
          <span className={isValid && !warningCount ? "status ok" : "status warn"}>{statusLabel}</span>
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
              <option key={option.value} value={option.value}>{formatHangerOptionLabel(option)}</option>
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
          <SelectNumberField label={`Дюбели периферия (${udRule.mode}, mm)`} value={room.udAnchorSpacing} manual={room.overrides.udAnchorSpacing} options={CUSTOM_UD_ANCHOR_OPTIONS} onChange={(value) => onRoomChange((draft) => {
            draft.udAnchorSpacing = value;
            syncDistanceOverride(draft, "udAnchorSpacing");
          })} />
        ) : (
          <NumberField label={`UD дюбели (${udRule.mode}, mm)`} value={room.udAnchorSpacing} manual={room.overrides.udAnchorSpacing} onChange={(value) => onRoomChange((draft) => {
            draft.udAnchorSpacing = value;
            syncDistanceOverride(draft, "udAnchorSpacing");
          })} />
        )}
      </div>
      <p className="hint">{variantHint}</p>
      <p className="hint">{udRule.note}</p>
      <div className="room-save-row">
        {saveStatus && <span className="save-status">{saveStatus}</span>}
        <button type="button" className="workspace-toggle-button" onClick={onSave}>Save</button>
      </div>
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
  if (check.status === "none") return null;

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
          <NumberInput label="Първи профил от стена (cm)" value={constants.profileEdgeOffsetCm} step={1} onChange={(value) => onChange({ profileEdgeOffsetCm: Math.max(0, value) })} />
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

function Visualization({ room, result, constants, zoom, onZoomChange, onOpenCutOptimization }: {
  room: Room;
  result: CalcResult;
  constants: CalculatorConstants;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  onOpenCutOptimization: () => void;
}) {
  const layout = buildSuspendedCeilingLayout({
    roomWidthCm: result.W,
    roomLengthCm: result.L,
    profileLengthCm: constants.cdLength * 100,
    carrierRowSpacingCm: room.c / 10,
    hangerSpacingCm: room.a / 10,
    firstHangerOffsetCm: constants.profileEdgeOffsetCm,
  });
  const bearingPositions = layout.carrierRowsYcm;
  const mountingPositions = buildLinearPositions(result.L, room.b / 10, constants.profileEdgeOffsetCm);
  const hangerPositions = layout.hangerPositionsCm;
  const extensionLayout = layout.carrierExtensions;
  const extensionPositions = Array.from(new Set(extensionLayout.flatMap((line) => line.pointsCm))).sort((left, right) => left - right);
  const width = 980;
  const height = 560;
  const pad = 86;
  const gridW = width - pad * 2;
  const gridH = height - pad * 2;
  const xScale = gridW / result.L;
  const yScale = gridH / result.W;

  return (
    <section className="visual-panel">
      <div className="visual-toolbar">
        <div>
          <p className="eyebrow">Работна схема</p>
          <h2>{room.name} - {room.systemType}</h2>
        </div>
        <div className="zoom-group">
          <button type="button" className="workspace-toggle-button" onClick={onOpenCutOptimization}>Оптимизация на разкроя</button>
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
          {extensionPositions.map((extension, index) => {
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

          {bearingPositions.map((position, lineIndex) => {
            const y = pad + position * yScale;
            const lineExtensions = extensionLayout.find((line) => line.lineIndex === lineIndex)?.pointsCm ?? [];
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
                {lineExtensions.map((extension) => (
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

function RoomsTable({ rooms, constants, activeRoomId, onAddRoom, onSelect, onDelete, onDeleteAll, onRecalculate, onOpenExport, onImportJson }: {
  rooms: Room[];
  constants: CalculatorConstants;
  activeRoomId: string;
  onAddRoom: () => void;
  onSelect: (roomId: string) => void;
  onDelete: (roomId: string) => void;
  onDeleteAll: () => void;
  onRecalculate: () => void;
  onOpenExport: () => void;
  onImportJson: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <section className="panel table-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Стаи</p>
          <h2>Количества по стаи</h2>
        </div>
        <div className="room-table-actions">
          <label className="file-button small">Импорт JSON<input type="file" accept="application/json" onChange={onImportJson} /></label>
          <button type="button" className="ghost small" disabled={!rooms.length} onClick={onOpenExport}>Експорт</button>
          <button type="button" className="ghost small" disabled={!rooms.length} onClick={onRecalculate}>Прекалкулирай</button>
          <button type="button" className="danger small" disabled={!rooms.length} onClick={onDeleteAll}>Изтрий всички</button>
        </div>
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

function MaterialsPanel({ rooms, constants, onReserveChange, onExportExcel }: {
  rooms: Room[];
  constants: CalculatorConstants;
  onReserveChange: (wastePercent: number) => void;
  onExportExcel: () => void;
}) {
  const rows = buildMaterialTakeoff(rooms, constants);
  return (
    <section className="panel materials-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Материали</p>
          <h2>Общо за всички стаи</h2>
        </div>
        <button type="button" className="ghost small" disabled={!rows.length} onClick={onExportExcel}>Експорт</button>
      </div>
      <div className="reserve-field">
        <NumberInput label="Резерв (%)" value={constants.wastePercent} step={0.1} onChange={onReserveChange} />
      </div>
      <div className="material-list">
        {rows.map((row) => {
          const catalogProduct = getCatalogProductForMaterial(row.label, row.key);
          const explanation = buildMaterialCalculationInfo(row, rooms, constants);
          return (
            <div key={row.key} className="material-row">
              <span>{row.label}</span>
              <strong>{row.quantity} {row.unit}</strong>
              <small>{explanation}{catalogProduct ? ` Каталожен артикул: ${catalogProduct.knaufName}.` : ""}</small>
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
