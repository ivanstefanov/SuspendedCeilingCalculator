import { ChangeEvent, ReactNode, useMemo, useState } from "react";
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
  optimizeAllRoomsSuspendedCeilingCuts,
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
type RoomWorkspacePanel = "visual" | "materials" | "cut" | "installation" | "validations";
type CutOptimizationMode = "room" | "global";
type ExportContentType = "room-cards" | "table" | "installation-guide";
type RoomCardExportFileType = "pdf" | "png" | "html";
type InstallationGuideExportFileType = "pdf" | "html";
type TableExportFileType = "excel" | "json" | "html";
type ExportFileType = RoomCardExportFileType | InstallationGuideExportFileType | TableExportFileType;

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

function buildSafeGlobalCutPlan(rooms: Room[], constants: CalculatorConstants): CutPlanState {
  if (!rooms.length) {
    return {
      plan: null,
      error: "Няма запазени стаи за общ разкрой.",
    };
  }

  try {
    return {
      plan: optimizeAllRoomsSuspendedCeilingCuts(rooms, constants, {
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
      error: error instanceof Error ? error.message : "Неуспешен общ разкрой.",
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
    if (materialRooms.length === 1) {
      return `${roomText}: Връзките са пресичанията между реално генерираните носещи и монтажни CD редове: ${totals.bearingRows} x ${totals.mountingRows} = ${totals.crossConnectors} бр.${reserve}`;
    }
    return `${roomText}: Връзките са пресичанията между реално генерираните носещи и монтажни CD редове във всяка стая. Общ сбор: ${totals.crossConnectors} бр.${reserve}`;
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

function formatMaterialQuantity(row: MaterialTakeoffItem): string {
  const base = `${formatNumber(row.quantity)} ${row.unit}`;
  if (row.optimizedQuantity == null) return base;
  return `${base} (след разкрой: ${formatNumber(row.optimizedQuantity)} ${row.unit})`;
}

function buildMaterialExplanation(row: MaterialTakeoffItem, rooms: Room[], constants: CalculatorConstants): string {
  const parts = [buildMaterialCalculationInfo(row, rooms, constants)];
  if (row.optimizedExplanation) parts.push(row.optimizedExplanation);
  if (row.optimizedQuantity != null && row.optimizedQuantity > row.quantity) {
    parts.push("Оптимизираният разкрой изисква повече профили заради реалните дължини на отделните парчета.");
  }
  return parts.join(" ");
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

function renderSvgMaterialsRows(rows: MaterialTakeoffItem[], rooms: Room[], constants: CalculatorConstants, startY: number): { markup: string; height: number } {
  const x = 48;
  const nameW = 280;
  const qtyW = 205;
  const noteW = 515;
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
    const explanation = buildMaterialExplanation(row, rooms, constants);
    const name = svgText(x + 10, y + 20, row.label, "table-text", 32, 16);
    const qty = svgText(x + nameW + 10, y + 20, formatMaterialQuantity(row), "table-text", 22, 16);
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

function renderHtmlMaterialsRows(rows: MaterialTakeoffItem[], rooms: Room[], constants: CalculatorConstants): string {
  return rows.map((row) => `<tr>
    <td>${escapeHtml(row.label)}</td>
    <td>${escapeHtml(formatMaterialQuantity(row))}</td>
    <td>${escapeHtml(buildMaterialExplanation(row, rooms, constants))}</td>
  </tr>`).join("");
}

function getCutSegmentDisplayLabel(bar: CutBar, pieceId: string, lengthCm: number, showRoomNames: boolean): string {
  const piece = bar.pieces.find((item) => item.id === pieceId);
  return showRoomNames && piece?.roomName
    ? `${Math.round(lengthCm)} cm · ${piece.roomName}`
    : String(Math.round(lengthCm));
}

function renderSvgCutPlan(cutPlan: CutPlanState, result: CalcResult, startY: number, showRoomNames = false): { markup: string; height: number } {
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
  const geometryNote = svgText(
    x,
    y + 16,
    "Разкроят само подрежда вече генерираните конструктивни парчета в покупни пръти. Късите монтажни CD сегменти при D113 са реални парчета между носещи редове/периметър, не грешка в оптимизацията.",
    "table-text",
    120,
    18,
  );
  markup += geometryNote.markup;
  y += geometryNote.height + 12;

  const groups = [
    { title: "Разкрой CD профили", bars: plan.bars.filter(isCdCutBar) },
    { title: "Разкрой UD профили", bars: plan.bars.filter(isUdCutBar) },
  ];

  groups.forEach((group) => {
    if (!group.bars.length) return;
    markup += `<text x="${x}" y="${y + 16}" class="table-label">${escapeHtml(group.title)}</text>`;
    y += 24;
    group.bars.forEach((bar) => {
    const barH = 42;
    markup += `<g>
      <text x="${x}" y="${y + 15}" class="table-label">${escapeHtml(bar.id)} - ${escapeHtml(getCutPieceLabel(bar.type))} - ${formatNumber(bar.usedCm)} / ${formatNumber(bar.stockLengthCm)} cm</text>
      <rect x="${x}" y="${y + 21}" width="${width}" height="18" class="cut-bg" />
      ${bar.segments.map((segment) => {
        const sx = x + (segment.startCm / bar.stockLengthCm) * width;
        const sw = Math.max(1, (segment.lengthCm / bar.stockLengthCm) * width);
        const label = getCutSegmentDisplayLabel(bar, segment.pieceId, segment.lengthCm, showRoomNames);
        return `<g>
          <rect x="${sx}" y="${y + 21}" width="${sw}" height="18" class="cut-${segment.type}" />
          ${sw >= (showRoomNames ? 90 : 28) ? `<text x="${sx + sw / 2}" y="${y + 35}" class="cut-piece-label" text-anchor="middle">${escapeHtml(label)}</text>` : ""}
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
  });

  return { markup, height: y - startY };
}

function renderHtmlCutPlan(cutPlan: CutPlanState, result: CalcResult, showRoomNames = false): string {
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
    <p class="cut-note">Разкроят само подрежда вече генерираните конструктивни парчета в покупни пръти. Късите монтажни CD сегменти при D113 са реални парчета между носещи редове/периметър, не грешка в оптимизацията.</p>
    <div class="cut-legend"><span class="carrier">Носещ CD</span><span class="mounting">Монтажен CD</span><span class="ud">UD</span><span class="waste">Остатък</span></div>
    ${[
      ["Разкрой CD профили", plan.bars.filter(isCdCutBar)] as const,
      ["Разкрой UD профили", plan.bars.filter(isUdCutBar)] as const,
    ].map(([title, bars]) => bars.length ? `<h3>${escapeHtml(title)}</h3><div class="cut-bar-list">
      ${bars.map((bar) => `
        <article class="cut-bar-card">
          <div class="cut-bar-head"><strong>${escapeHtml(bar.id)}</strong><span>${escapeHtml(getCutPieceLabel(bar.type))}</span><small>${formatNumber(bar.usedCm)} / ${formatNumber(bar.stockLengthCm)} cm</small></div>
          <div class="cut-strip">
            ${bar.segments.map((segment) => {
              const label = getCutSegmentDisplayLabel(bar, segment.pieceId, segment.lengthCm, showRoomNames);
              return `<div class="cut-piece ${segment.type}" title="${escapeHtml(label)}" style="width:${(segment.lengthCm / bar.stockLengthCm) * 100}%">${escapeHtml(label)}</div>`;
            }).join("")}
            ${bar.wasteCm > 0 ? `<div class="cut-piece waste" style="width:${(bar.wasteCm / bar.stockLengthCm) * 100}%">${Math.round(bar.wasteCm)}</div>` : ""}
          </div>
        </article>`).join("")}
    </div>` : "").join("")}`;
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
  const materialRows = renderSvgMaterialsRows(materials, [safeRoom], constants, y);
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

function buildAggregateCutResult(rooms: Room[], constants: CalculatorConstants): CalcResult {
  return rooms.reduce((total, sourceRoom) => {
    const result = calc(cloneRoom(sourceRoom), constants);
    total.W += result.W;
    total.L += result.L;
    total.bearingCount += result.bearingCount;
    total.mountingCount += result.mountingCount;
    total.bearingLengthTotal += result.bearingLengthTotal;
    total.mountingLengthTotal += result.mountingLengthTotal;
    total.bearingProfiles += result.bearingProfiles;
    total.mountingProfiles += result.mountingProfiles;
    total.cdTotalLength += result.cdTotalLength;
    total.cdTotalProfiles += result.cdTotalProfiles;
    total.crossConnectors += result.crossConnectors;
    total.hangersPerBearing += result.hangersPerBearing;
    total.hangersTotal += result.hangersTotal;
    total.udTotalLength += result.udTotalLength;
    total.udProfiles += result.udProfiles;
    total.anchorsUd += result.anchorsUd;
    total.anchorsHangers += result.anchorsHangers;
    total.anchorsTotal += result.anchorsTotal;
    total.metalScrews += result.metalScrews;
    total.drywallScrews += result.drywallScrews;
    total.extensionsTotal += result.extensionsTotal;
    return total;
  }, {
    W: 0,
    L: 0,
    bearingCount: 0,
    mountingCount: 0,
    bearingLengthTotal: 0,
    mountingLengthTotal: 0,
    bearingProfiles: 0,
    mountingProfiles: 0,
    cdTotalLength: 0,
    cdTotalProfiles: 0,
    crossConnectors: 0,
    hangersPerBearing: 0,
    hangersTotal: 0,
    udTotalLength: 0,
    udProfiles: 0,
    anchorsUd: 0,
    anchorsHangers: 0,
    anchorsTotal: 0,
    metalScrews: 0,
    drywallScrews: 0,
    extensionsTotal: 0,
  });
}

function buildGlobalCutReportSvg(rooms: Room[], constants: CalculatorConstants): string {
  const cutPlan = buildSafeGlobalCutPlan(rooms, constants);
  const result = buildAggregateCutResult(rooms, constants);
  const materials = buildMaterialTakeoff(rooms, constants);
  const generatedAt = new Date().toLocaleString("bg-BG");
  const pageW = 1100;
  let y = 42;
  let body = "";

  body += `<text x="48" y="${y}" class="title">Общ разкрой</text>`;
  y += 28;
  body += `<text x="48" y="${y}" class="meta">Knauf D11 calculator · ${escapeHtml(generatedAt)} · ${rooms.length} стаи</text>`;
  y += 34;

  body += `<text x="48" y="${y}" class="section-title">Общо материали</text>`;
  y += 14;
  const materialRows = renderSvgMaterialsRows(materials, rooms, constants, y);
  body += materialRows.markup;
  y += materialRows.height + 34;

  body += `<text x="48" y="${y}" class="section-title">Общ разкрой за всички стаи</text>`;
  y += 16;
  const cut = renderSvgCutPlan(cutPlan, result, y, true);
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
    * { box-sizing: border-box; }
    body { margin: 0; padding: 12px; font-family: Arial, sans-serif; color: #17211f; background: #ffffff; overflow-x: hidden; }
    main { width: 100%; max-width: none; margin: 0; }
    h1 { margin: 0 0 6px; font-size: 26px; }
    h2 { margin: 28px 0 10px; font-size: 18px; color: #0f766e; }
    .meta { color: #5f6f6a; margin-bottom: 22px; }
    table { width: 100%; max-width: 100%; table-layout: fixed; border-collapse: collapse; margin: 8px 0 18px; font-size: 13px; }
    th, td { border: 1px solid #d7e0e3; padding: 7px 8px; text-align: left; vertical-align: top; overflow-wrap: anywhere; }
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
    .cut-note { color: #5f6f6a; font-size: 13px; line-height: 1.45; }
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
      <tbody>${renderHtmlMaterialsRows(materials, [safeRoom], constants)}</tbody>
    </table>
    <h2>Работна схема</h2>
    ${renderWorkingSchemeSvg(safeRoom, result, constants)}
    <h2>Разкрой за стаята</h2>
    ${renderHtmlCutPlan(cutPlan, result)}
  </main>
</body>
</html>`;
}

function buildGlobalCutReportHtml(rooms: Room[], constants: CalculatorConstants): string {
  const cutPlan = buildSafeGlobalCutPlan(rooms, constants);
  const result = buildAggregateCutResult(rooms, constants);
  const materials = buildMaterialTakeoff(rooms, constants);
  const generatedAt = new Date().toLocaleString("bg-BG");
  return `<!doctype html>
<html lang="bg">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Общ разкрой - Knauf calculator</title>
  <style>
    body { margin: 0; padding: 28px; font-family: Arial, sans-serif; color: #17211f; background: #ffffff; }
    main { max-width: 1120px; margin: 0 auto; }
    h1 { margin: 0 0 6px; font-size: 26px; }
    h2 { margin: 28px 0 10px; font-size: 18px; color: #0f766e; }
    .meta { color: #5f6f6a; margin-bottom: 22px; }
    table { width: 100%; border-collapse: collapse; margin: 8px 0 18px; font-size: 13px; }
    th, td { border: 1px solid #d7e0e3; padding: 7px 8px; text-align: left; vertical-align: top; }
    th { width: 28%; background: #f6f8f4; }
    .cut-legend { display: flex; flex-wrap: wrap; gap: 8px; margin: 10px 0 12px; font-size: 12px; }
    .cut-note { color: #5f6f6a; font-size: 13px; line-height: 1.45; }
    .cut-legend span { padding: 4px 8px; border-radius: 999px; color: #fff; }
    .cut-legend .carrier, .cut-piece.carrier { background: #0f766e; }
    .cut-legend .mounting, .cut-piece.mounting { background: #f59e0b; color: #17211f; }
    .cut-legend .ud, .cut-piece.ud { background: #2563eb; }
    .cut-legend .waste, .cut-piece.waste { background: #d1d5db; color: #374151; }
    .cut-bar-list { display: grid; gap: 9px; max-width: 100%; }
    .cut-bar-card { max-width: 100%; overflow: hidden; break-inside: avoid; border: 1px solid #d7e0e3; border-radius: 6px; padding: 8px; }
    .cut-bar-head { display: flex; flex-wrap: wrap; gap: 8px; justify-content: space-between; font-size: 12px; margin-bottom: 6px; }
    .cut-strip { display: flex; width: 100%; height: 28px; overflow: hidden; border-radius: 4px; background: #eef2f7; }
    .cut-piece { min-width: 0; padding: 0 2px; display: grid; place-items: center; color: #fff; font-size: 11px; border-right: 1px solid rgba(255,255,255,0.7); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .warning { padding: 10px; border: 1px solid #f3c5c5; background: #fff4f4; color: #991b1b; }
    @media (max-width: 720px) { body { padding: 8px; } }
    @media print { body { padding: 8mm; } h2, .cut-bar-card { break-inside: avoid; } }
  </style>
</head>
<body>
  <main>
    <h1>Общ разкрой</h1>
    <div class="meta">Knauf D11 calculator · ${escapeHtml(generatedAt)} · ${rooms.length} стаи</div>
    <h2>Общо материали</h2>
    <table>
      <thead><tr><th>Материал</th><th>Количество</th><th>Как е сметнато</th></tr></thead>
      <tbody>${renderHtmlMaterialsRows(materials, rooms, constants)}</tbody>
    </table>
    <h2>Общ разкрой за всички стаи</h2>
    ${renderHtmlCutPlan(cutPlan, result, true)}
  </main>
</body>
</html>`;
}

interface InstallationGuideStep {
  title: string;
  items: string[];
}

interface InstallationCutPreparationStep {
  enabled: boolean;
  cd?: {
    totalProfiles: number;
    bars: Array<{
      id: string;
      segments: Array<{
        lengthCm: number;
        type: "carrier" | "mounting";
      }>;
    }>;
  };
  ud?: {
    totalProfiles: number;
    bars: Array<{
      id: string;
      segments: Array<{
        lengthCm: number;
      }>;
    }>;
  };
  notes: string[];
}

interface InstallationGuide {
  header: Array<[string, unknown]>;
  summary: Array<[string, unknown]>;
  cutPreparationStep?: InstallationCutPreparationStep;
  steps: InstallationGuideStep[];
  notes: string[];
}

interface InstallationGuideOptions {
  showCutPreparation?: boolean;
}

function formatPositions(positions: number[], maxItems = 12): string {
  if (!positions.length) return "няма данни";
  const shown = positions.slice(0, maxItems).map((value) => `${Math.round(value)} cm`).join(", ");
  return positions.length > maxItems ? `${shown} ...` : shown;
}

function formatSegmentLengths(lengths: number[], maxItems = 10): string {
  if (!lengths.length) return "няма снадки/сегменти";
  const shown = lengths.slice(0, maxItems).map((value) => `${formatNumber(value)} cm`).join(", ");
  return lengths.length > maxItems ? `${shown} ...` : shown;
}

function formatCutBarLengths(lengths: number[]): string {
  return lengths.map((lengthCm) => formatNumber(lengthCm)).join(" + ");
}

function buildCutPreparationStep(cutOptimizationResult: CutOptimizationResult | null | undefined, showCutPreparation: boolean): InstallationCutPreparationStep | undefined {
  if (!showCutPreparation || !cutOptimizationResult) return undefined;
  const cdBars = cutOptimizationResult.bars.filter(isCdCutBar).map((bar) => ({
    id: bar.id,
    segments: bar.segments
      .filter((segment) => segment.type === "carrier" || segment.type === "mounting")
      .map((segment) => ({
        lengthCm: segment.lengthCm,
        type: segment.type as "carrier" | "mounting",
      })),
  })).filter((bar) => bar.segments.length > 0);
  const udBars = cutOptimizationResult.bars.filter(isUdCutBar).map((bar) => ({
    id: bar.id,
    segments: bar.segments
      .filter((segment) => segment.type === "ud")
      .map((segment) => ({ lengthCm: segment.lengthCm })),
  })).filter((bar) => bar.segments.length > 0);

  return {
    enabled: true,
    cd: cdBars.length ? { totalProfiles: cdBars.length, bars: cdBars } : undefined,
    ud: udBars.length ? { totalProfiles: udBars.length, bars: udBars } : undefined,
    notes: [
      "CD и UD профилите не се смесват при рязане.",
      "Снадките при CD да са близо до окачвач.",
      "UD може да се съставя от няколко по-къси парчета.",
    ],
  };
}

function buildInstallationGuide(
  room: Room,
  result: CalcResult,
  constants: CalculatorConstants,
  cutOptimizationResult?: CutOptimizationResult | null,
  options: InstallationGuideOptions = {},
): InstallationGuide {
  const boardLayers = getEffectiveBoardLayers(room, constants);
  const boardArea = Number(room.area) * boardLayers;
  const boardSize = Math.max(0.01, constants.boardWidth * constants.boardLength);
  const boardCount = Math.ceil(boardArea / boardSize);
  const reserveMultiplier = 1 + constants.wastePercent / 100;
  const jointTapeLength = Number((Number(room.area) * constants.jointTapePerM2 * reserveMultiplier).toFixed(2));
  const jointCompoundKg = Number((Number(room.area) * constants.jointCompoundKgPerM2 * reserveMultiplier).toFixed(2));
  const trennFixLength = Number((result.udTotalLength * constants.trennFixPerimeterMultiplier * reserveMultiplier).toFixed(2));
  const layout = buildSuspendedCeilingLayout({
    roomWidthCm: result.W,
    roomLengthCm: result.L,
    profileLengthCm: constants.cdLength * 100,
    carrierRowSpacingCm: room.c / 10,
    hangerSpacingCm: room.a / 10,
    firstHangerOffsetCm: constants.profileEdgeOffsetCm,
  });
  const mountingPositions = buildLinearPositions(result.L, room.b / 10, constants.profileEdgeOffsetCm);
  const cutInput = buildCutOptimizationInput(room, result, constants);
  const carrierSegmentLengths = cutInput.carrierRows.flatMap((row) => row.segments.map((segment) => segment.lengthCm));
  const mountingSegmentLengths = cutInput.mountingRows.flatMap((row) => row.segments.map((segment) => segment.lengthCm));
  const cdOptimizedQuantity = cutOptimizationResult
    ? cutOptimizationResult.bars.filter(isCdCutBar).length
    : undefined;
  const cutPreparationStep = buildCutPreparationStep(cutOptimizationResult, options.showCutPreparation ?? true);

  return {
    header: [
      ["Стая", room.name],
      ["Система", room.systemType],
      ["Размери", `${room.width} x ${room.length} cm`],
      ["Площ", `${formatNumber(room.area)} m2`],
    ],
    summary: [
      ["Носещи CD редове", result.bearingCount],
      ["Монтажни CD редове", result.mountingCount],
      ["Окачвачи", result.hangersTotal],
      ["Връзки CD", result.crossConnectors],
      ["UD периметър", `${formatNumber(result.udTotalLength)} m`],
      ["CD след разкрой", cdOptimizedQuantity == null ? "няма данни" : `${cdOptimizedQuantity} бр.`],
    ],
    cutPreparationStep,
    steps: [
      {
        title: "Монтаж на UD по периметъра",
        items: [`UD дължина: ${formatNumber(result.udTotalLength)} m`, `Дюбели: ${result.anchorsUd} бр. през ${room.udAnchorSpacing} mm`],
      },
      {
        title: "Разчертаване на носещите CD редове",
        items: [`Носещи редове: ${result.bearingCount} бр.`, `Позиции: ${formatPositions(layout.carrierRowsYcm)}`],
      },
      {
        title: "Монтаж на директните окачвачи",
        items: [`Окачвачи: ${result.hangersTotal} бр.`, `Разстояние a: ${room.a} mm`, `X позиции: ${formatPositions(layout.hangerPositionsCm)}`],
      },
      {
        title: "Монтаж на носещите CD профили",
        items: [`Носещи редове: ${result.bearingCount} бр.`, `Сегменти/снадки: ${formatSegmentLengths(carrierSegmentLengths)}`],
      },
      {
        title: "Монтаж на монтажните CD профили",
        items: [`Монтажни редове: ${result.mountingCount} бр.`, `Разстояние b: ${room.b} mm`, `Позиции: ${formatPositions(mountingPositions)}`, `Връзки CD: ${result.crossConnectors} бр.`],
      },
      {
        title: "Проверка на нивелация и геометрия",
        items: ["UD профилите са нивелирани", "CD профилите са в една равнина", "Окачвачите са фиксирани стабилно", "Снадките са близо до окачвач"],
      },
      {
        title: "Монтаж на гипсокартон",
        items: [`Плоскости: ${boardCount} бр.`, `Площ обшивка: ${formatNumber(boardArea)} m2`, `TN винтове: ${result.drywallScrews} бр.`],
      },
      {
        title: "Фугиране и довършване",
        items: [`Фуголента: ${jointTapeLength} m`, `Шпакловка: ${jointCompoundKg} kg`, `Trenn-Fix: ${trennFixLength} m`],
      },
    ],
    notes: [
      "Провери основата и нивото преди пробиване.",
      "Не смесвай CD и UD профили при рязане и монтаж.",
      "Снадките на CD профилите не трябва да се подреждат в една линия.",
      "Преди затваряне с плоскости провери всички връзки, окачвачи и периферия.",
    ],
  };
}

function getCutSummaryRows(cutPlan: CutPlanState): Array<[string, unknown]> {
  if (cutPlan.error || !cutPlan.plan) return [["Разкрой", cutPlan.error || "неуспешно изчисление"]];
  const plan = cutPlan.plan;
  return [
    ["Профили след разкрой", `${plan.totalBars} бр.`],
    ["CD пръти", `${plan.bars.filter(isCdCutBar).length} бр.`],
    ["UD пръти", `${plan.bars.filter(isUdCutBar).length} бр.`],
    ["Отпадък", `${formatNumber(plan.totalWasteCm)} cm`],
    ["Ефективност", `${formatNumber(plan.efficiencyPercent)} %`],
  ];
}

function renderInstallationCutPreparationHtml(step?: InstallationCutPreparationStep): string {
  if (!step?.enabled) return "";
  const cd = step.cd ? `
    <h3>CD профили</h3>
    <p>Общо профили CD за покупка: ${step.cd.totalProfiles} бр. Профилите са оптимизирани така, че остатъците да се използват максимално ефективно.</p>
    <ul>${step.cd.bars.map((bar, index) => `<li>Профил ${index + 1}: ${escapeHtml(formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm)))}</li>`).join("")}</ul>
  ` : "";
  const ud = step.ud ? `
    <h3>UD профили (периметър)</h3>
    <p>UD профилите могат да се комбинират свободно по периферията.</p>
    <ul>${step.ud.bars.map((bar, index) => `<li>Профил ${index + 1}: ${escapeHtml(formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm)))}</li>`).join("")}</ul>
  ` : "";
  return `<section class="cut-prep">
    <h2>Разкрой на профилите (подготовка)</h2>
    ${cd}
    ${ud}
    <ul>${step.notes.map((note) => `<li>${escapeHtml(note)}</li>`).join("")}</ul>
  </section>`;
}

function appendInstallationCutPreparationSvg(body: string, y: number, step?: InstallationCutPreparationStep): { body: string; y: number } {
  if (!step?.enabled) return { body, y };
  let nextBody = body;
  let nextY = y;
  nextBody += `<rect x="30" y="${nextY - 16}" width="734" height="34" class="cut-prep-bg" />`;
  nextBody += `<text x="42" y="${nextY + 6}" class="section-title">Разкрой на профилите (подготовка)</text>`;
  nextY += 28;
  if (step.cd) {
    nextBody += `<text x="42" y="${nextY}" class="step-title">CD профили: ${step.cd.totalProfiles} бр.</text>`;
    nextY += 16;
    step.cd.bars.slice(0, 8).forEach((bar, index) => {
      nextBody += `<text x="54" y="${nextY}" class="step-data">Профил ${index + 1}: ${escapeHtml(formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm)))}</text>`;
      nextY += 14;
    });
  }
  if (step.ud) {
    nextBody += `<text x="42" y="${nextY}" class="step-title">UD профили: ${step.ud.totalProfiles} бр.</text>`;
    nextY += 16;
    step.ud.bars.slice(0, 6).forEach((bar, index) => {
      nextBody += `<text x="54" y="${nextY}" class="step-data">Профил ${index + 1}: ${escapeHtml(formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm)))}</text>`;
      nextY += 14;
    });
  }
  step.notes.forEach((note) => {
    nextBody += `<text x="42" y="${nextY}" class="step-data">• ${escapeHtml(note)}</text>`;
    nextY += 14;
  });
  return { body: nextBody, y: nextY + 10 };
}

function buildInstallationGuideHtml(room: Room, constants: CalculatorConstants): string {
  const safeRoom = cloneRoom(room);
  const result = calc(safeRoom, constants);
  const cutPlan = buildSafeCutPlan(safeRoom, result, constants);
  const guide = buildInstallationGuide(safeRoom, result, constants, cutPlan.plan);
  const generatedAt = new Date().toLocaleString("bg-BG");

  return `<!doctype html>
<html lang="bg">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(safeRoom.name)} - Монтажни етапи</title>
  <style>
    @page { size: A4; margin: 10mm; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 14px; font-family: Arial, sans-serif; color: #17211f; background: #fff; }
    main { max-width: 980px; margin: 0 auto; }
    h1 { margin: 0 0 4px; font-size: 25px; }
    h2 { margin: 18px 0 8px; font-size: 16px; color: #0f766e; }
    .meta { color: #5f6f6a; font-size: 12px; margin-bottom: 12px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed; }
    th, td { border: 1px solid #d7e0e3; padding: 6px 7px; text-align: left; vertical-align: top; overflow-wrap: anywhere; }
    th { width: 42%; background: #f6f8f4; }
    .steps { display: grid; gap: 6px; }
    .step { display: grid; grid-template-columns: 24px 1fr; gap: 8px; align-items: start; border: 1px solid #d7e0e3; padding: 7px; break-inside: avoid; }
    .cut-prep { margin: 12px 0; border: 1px solid #cde7e2; border-radius: 6px; padding: 10px; background: #f5fbf8; break-inside: avoid; }
    .cut-prep h2 { margin-top: 0; }
    .cut-prep h3 { margin: 10px 0 4px; font-size: 13px; }
    .cut-prep p, .cut-prep li { font-size: 12px; color: #34413c; }
    .cut-prep ul { margin: 4px 0 0; padding-left: 18px; }
    .box { width: 15px; height: 15px; border: 2px solid #17211f; margin-top: 1px; }
    .step strong { display: block; font-size: 13px; }
    .step span, .notes li { font-size: 12px; color: #34413c; }
    .ceiling-svg { width: 100%; height: auto; max-height: 330px; border: 1px solid #d7e0e3; background: #f6f8f4; }
    .axis-label, .position-label, .hanger-label, .extension-label { font: 12px Arial, sans-serif; fill: #263633; }
    .bearing-line { stroke: #0f766e; stroke-width: 3; }
    .mounting-line { stroke: #f59e0b; stroke-width: 2; }
    .hanger-dot { fill: #111827; }
    .hanger-dimension-line { stroke: #111827; stroke-width: 1; stroke-dasharray: 4 4; }
    .extension-dimension-line, .extension-mark { stroke: #dc2626; stroke-width: 2; }
    .extension-label { fill: #991b1b; }
    .notes { margin: 0; padding-left: 18px; }
    @media print { body { padding: 0; } h2, .step, .ceiling-svg { break-inside: avoid; } }
  </style>
</head>
<body>
  <main>
    <h1>Монтажни етапи</h1>
    <div class="meta">${escapeHtml(safeRoom.name)} · ${escapeHtml(safeRoom.systemType)} · ${escapeHtml(generatedAt)}</div>
    <div class="grid">
      <section><h2>Данни</h2><table><tbody>${guide.header.map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table></section>
      <section><h2>Обобщение</h2><table><tbody>${guide.summary.map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table></section>
    </div>
    <h2>Checklist</h2>
    ${renderInstallationCutPreparationHtml(guide.cutPreparationStep)}
    <section class="steps">${guide.steps.map((step) => `<article class="step"><span class="box"></span><div><strong>${escapeHtml(step.title)}</strong><span>${step.items.map(escapeHtml).join(" · ")}</span></div></article>`).join("")}</section>
    <h2>Схема</h2>
    ${renderWorkingSchemeSvg(safeRoom, result, constants)}
    <div class="grid">
      <section><h2>Кратък разкрой</h2><table><tbody>${getCutSummaryRows(cutPlan).map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}</tbody></table></section>
      <section><h2>Бележки</h2><ul class="notes">${guide.notes.map((note) => `<li>${escapeHtml(note)}</li>`).join("")}</ul></section>
    </div>
  </main>
</body>
</html>`;
}

function buildInstallationGuideSvg(room: Room, constants: CalculatorConstants): string {
  const safeRoom = cloneRoom(room);
  const result = calc(safeRoom, constants);
  const cutPlan = buildSafeCutPlan(safeRoom, result, constants);
  const guide = buildInstallationGuide(safeRoom, result, constants, cutPlan.plan);
  const generatedAt = new Date().toLocaleString("bg-BG");
  const pageW = 794;
  const pageH = 1123;
  let y = 34;
  let body = "";

  body += `<text x="36" y="${y}" class="title">Монтажни етапи</text>`;
  y += 22;
  body += `<text x="36" y="${y}" class="meta">${escapeHtml(safeRoom.name)} · ${escapeHtml(safeRoom.systemType)} · ${escapeHtml(generatedAt)}</text>`;
  y += 22;
  const headerRows = renderSvgKeyValueRows([...guide.header, ...guide.summary.slice(0, 5)], 0);
  body += `<g transform="translate(0 ${y}) scale(0.72)">${headerRows.markup}</g>`;
  y += headerRows.height * 0.72 + 24;

  body += `<text x="36" y="${y}" class="section-title">Checklist</text>`;
  y += 14;
  const cutPrepSvg = appendInstallationCutPreparationSvg(body, y, guide.cutPreparationStep);
  body = cutPrepSvg.body;
  y = cutPrepSvg.y;
  guide.steps.forEach((step) => {
    body += `<rect x="36" y="${y - 10}" width="12" height="12" class="checkbox" />
      <text x="58" y="${y}" class="step-title">${escapeHtml(step.title)}</text>
      <text x="58" y="${y + 15}" class="step-data">${escapeHtml(step.items.join(" · "))}</text>`;
    y += 34;
  });

  body += `<text x="36" y="${y}" class="section-title">Схема</text>`;
  y += 10;
  body += renderWorkingSchemeSvg(safeRoom, result, constants).replace("<svg ", `<svg x="36" y="${y}" width="722" height="412" `);
  y += 430;

  body += `<text x="36" y="${y}" class="section-title">Кратък разкрой</text>`;
  y += 12;
  const cutRows = renderSvgKeyValueRows(getCutSummaryRows(cutPlan), 0);
  body += `<g transform="translate(0 ${y}) scale(0.72)">${cutRows.markup}</g>`;
  y += cutRows.height * 0.72 + 28;

  body += `<text x="36" y="${y}" class="section-title">Бележки</text>`;
  y += 18;
  guide.notes.slice(0, 3).forEach((note) => {
    body += `<text x="48" y="${y}" class="step-data">• ${escapeHtml(note)}</text>`;
    y += 16;
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${pageW}" height="${pageH}" viewBox="0 0 ${pageW} ${pageH}">
    <style>
      .page-bg { fill: #ffffff; }
      .title { font: 700 26px Arial, sans-serif; fill: #17211f; }
      .meta { font: 12px Arial, sans-serif; fill: #5f6f6a; }
      .section-title { font: 700 16px Arial, sans-serif; fill: #0f766e; }
      .table-head { fill: #edf5f1; stroke: #d7e0e3; }
      .table-cell { fill: #ffffff; stroke: #d7e0e3; }
      .table-label { font: 700 13px Arial, sans-serif; fill: #34413c; }
      .table-text { font: 13px Arial, sans-serif; fill: #17211f; }
      .checkbox { fill: #ffffff; stroke: #17211f; stroke-width: 2; }
      .cut-prep-bg { fill: #f5fbf8; stroke: #cde7e2; }
      .step-title { font: 700 13px Arial, sans-serif; fill: #17211f; }
      .step-data { font: 12px Arial, sans-serif; fill: #34413c; }
      .axis-label, .position-label, .hanger-label, .extension-label { font: 12px Arial, sans-serif; fill: #263633; }
      .bearing-line { stroke: #0f766e; stroke-width: 3; }
      .mounting-line { stroke: #f59e0b; stroke-width: 2; }
      .hanger-dot { fill: #111827; }
      .hanger-dimension-line { stroke: #111827; stroke-width: 1; stroke-dasharray: 4 4; }
      .extension-dimension-line, .extension-mark { stroke: #dc2626; stroke-width: 2; }
      .extension-label { fill: #991b1b; }
    </style>
    <rect x="0" y="0" width="${pageW}" height="${pageH}" class="page-bg" />
    ${body}
  </svg>`;
}

async function buildInstallationGuideBlob(room: Room, constants: CalculatorConstants, fileType: InstallationGuideExportFileType): Promise<Blob> {
  if (fileType === "html") {
    return new Blob([buildInstallationGuideHtml(room, constants)], { type: "text/html;charset=utf-8" });
  }
  const canvas = await buildCanvasFromSvg(buildInstallationGuideSvg(room, constants));
  return buildPdfFromCanvas(canvas, `${room.name} - Монтажни етапи`);
}

async function buildCanvasFromSvg(svg: string): Promise<HTMLCanvasElement> {
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

async function buildRoomReportCanvas(room: Room, constants: CalculatorConstants): Promise<HTMLCanvasElement> {
  return buildCanvasFromSvg(buildRoomReportSvg(room, constants));
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

async function buildGlobalCutReportBlob(rooms: Room[], constants: CalculatorConstants, fileType: RoomCardExportFileType): Promise<Blob> {
  if (fileType === "html") {
    return new Blob([buildGlobalCutReportHtml(rooms, constants)], { type: "text/html;charset=utf-8" });
  }
  const canvas = await buildCanvasFromSvg(buildGlobalCutReportSvg(rooms, constants));
  return fileType === "pdf"
    ? buildPdfFromCanvas(canvas, "Общ разкрой")
    : canvasToBlob(canvas, "image/png");
}

function getExportTypeDescription(contentType: ExportContentType, fileType: ExportFileType): string {
  if (contentType === "installation-guide") {
    return fileType === "pdf"
      ? "PDF export създава printable A4 монтажен checklist за всяка стая: данни, обобщение, етапи, схема, кратък разкрой и бележки."
      : "HTML export създава printable монтажен checklist за всяка стая, който може да се отвори в браузър и да се печата на A4.";
  }
  if (fileType === "pdf") {
    return "PDF export създава отделен файл за всяка стая плюс файл \"общ разкрой\". Файловете по стаи съдържат данни, материали, работна схема и разкрой, а общият файл събира оптимизирания разкрой за всички стаи.";
  }
  if (fileType === "png") {
    return "PNG export създава отделна картинка за всяка стая плюс картинка \"общ разкрой\". Подходящо е за бързо споделяне като изображение.";
  }
  if (fileType === "html") {
    return contentType === "room-cards"
      ? "HTML export създава отделен отваряем файл за всяка стая плюс файл \"общ разкрой\". Файловете могат да се преглеждат в браузър и да се печатат."
      : "HTML export създава един файл rooms.html с таблица на всички стаи. Подходящ е за преглед в браузър или бърз печат.";
  }
  if (fileType === "json") {
    return "JSON export създава един файл rooms.json със запазените стаи, активната стая, глобалните настройки и настройките на проекта. Този файл може после да се върне през Импорт JSON.";
  }
  return "Excel export създава един файл rooms.xls с таблица на всички стаи и основните изчислени стойности от grid-а.";
}

function getExportContentDescription(contentType: ExportContentType): string {
  if (contentType === "room-cards") {
    return "Ще бъдат създадени отделни файлове за всяка стая с размери, материали, схема и разкрой, както и файл \"общ разкрой\" за всички стаи.";
  }
  if (contentType === "installation-guide") {
    return "Ще бъдат създадени отделни файлове за всяка стая с кратки монтажни етапи, чекбоксове, схема и обобщен разкрой.";
  }
  return "Ще бъде създаден един файл с всички стаи и техните данни.";
}

function getExportFormatOptions(contentType: ExportContentType): Array<{ value: ExportFileType; label: string }> {
  if (contentType === "room-cards") {
    return [
      { value: "pdf", label: "PDF" },
      { value: "png", label: "PNG" },
      { value: "html", label: "HTML" },
    ];
  }
  if (contentType === "installation-guide") {
    return [
      { value: "pdf", label: "PDF" },
      { value: "html", label: "HTML" },
    ];
  }
  return [
    { value: "excel", label: "Excel" },
    { value: "json", label: "JSON" },
    { value: "html", label: "HTML" },
  ];
}

function isRoomCardFileType(fileType: ExportFileType): fileType is RoomCardExportFileType {
  return fileType === "pdf" || fileType === "png" || fileType === "html";
}

function isInstallationGuideFileType(fileType: ExportFileType): fileType is InstallationGuideExportFileType {
  return fileType === "pdf" || fileType === "html";
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
  const [cutOptimizationMode, setCutOptimizationMode] = useState<CutOptimizationMode>("room");
  const [installationGuideRoomId, setInstallationGuideRoomId] = useState<string | null>(null);

  const activeRoom = state.draftRoom;
  const activeResult = useMemo(() => calc(cloneRoom(activeRoom), state.constants), [activeRoom, state.constants]);
  const activeCutPlan = useMemo(() => buildSafeCutPlan(activeRoom, activeResult, state.constants), [activeRoom, activeResult, state.constants]);
  const globalCutPlan = useMemo(() => buildSafeGlobalCutPlan(state.rooms, state.constants), [state.rooms, state.constants]);
  const activeWarnings = useMemo(() => getValidationWarnings(cloneRoom(activeRoom)), [activeRoom]);
  const isValid = !activeWarnings.some((warning) => warning.severity === "error");
  const installationGuideRoom = installationGuideRoomId
    ? state.rooms.find((room) => room.id === installationGuideRoomId) ?? null
    : null;

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
      formatMaterialQuantity(row),
      row.unit,
      buildMaterialExplanation(row, state.rooms, state.constants),
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
    setExportFileType(contentType === "table" ? "excel" : "pdf");
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
      : exportContentType === "installation-guide"
        ? await Promise.all(state.rooms.map(async (room) => {
            const guideFileType = isInstallationGuideFileType(exportFileType) ? exportFileType : "pdf";
            const baseName = sanitizeFilename(`${room.name || room.systemType} - монтажни етапи`);
            const count = filenameCounts.get(baseName) ?? 0;
            filenameCounts.set(baseName, count + 1);
            const filename = `${baseName}${count ? `-${count + 1}` : ""}.${guideFileType}`;
            return {
              filename,
              blob: await buildInstallationGuideBlob(room, state.constants, guideFileType),
            };
          }))
        : await (async () => {
          const roomCardFileType = isRoomCardFileType(exportFileType) ? exportFileType : "pdf";
          const roomReports = await Promise.all(state.rooms.map(async (room) => {
            const baseName = sanitizeFilename(room.name || room.systemType);
            const count = filenameCounts.get(baseName) ?? 0;
            filenameCounts.set(baseName, count + 1);
            const filename = `${baseName}${count ? `-${count + 1}` : ""}.${roomCardFileType}`;
            return {
              filename,
              blob: await buildRoomReportBlob(room, state.constants, roomCardFileType),
            };
          }));
          return [
            ...roomReports,
            {
              filename: `${sanitizeFilename("общ разкрой")}.${roomCardFileType}`,
              blob: await buildGlobalCutReportBlob(state.rooms, state.constants, roomCardFileType),
            },
          ];
        })();

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
    { key: "room", label: "Активна стая" },
    { key: "materials", label: "Общи материали" },
    { key: "settings", label: "Настройки" },
    { key: "help", label: "Помощ" },
  ];

  return (
    <main className="app-shell">
      <section className="workbench">
        <div className="topbar">
          <div>
            <p className="eyebrow">Knauf D11 Calculator</p>
            <h1>Калкулатор за окачени тавани</h1>
            <small className="topbar-subtitle">{state.rooms.length} запазени стаи · {activeRoom.systemType}</small>
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
                <ResultTabs activeTab={roomWorkspacePanel} onChange={setRoomWorkspacePanel} warningCount={activeWarnings.length} />
                <RoomWorkspacePanelView
                  activeTab={roomWorkspacePanel}
                  room={activeRoom}
                  result={activeResult}
                  constants={state.constants}
                  rooms={state.rooms}
                  zoom={zoom}
                  warnings={activeWarnings}
                  activeCutPlan={activeCutPlan}
                  globalCutPlan={globalCutPlan}
                  cutOptimizationMode={cutOptimizationMode}
                  onZoomChange={setZoom}
                  onCutOptimizationModeChange={setCutOptimizationMode}
                />
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
              onOpenInstallationGuide={setInstallationGuideRoomId}
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
      {installationGuideRoom && (
        <InstallationGuideModal
          room={installationGuideRoom}
          constants={state.constants}
          onClose={() => setInstallationGuideRoomId(null)}
        />
      )}
      {activeSection === "room" && (
        <MobileActionBar
          onSave={saveCurrentState}
          onCut={() => setRoomWorkspacePanel("cut")}
          onMaterials={() => setRoomWorkspacePanel("materials")}
          onExport={() => setIsExportModalOpen(true)}
          canExport={Boolean(state.rooms.length)}
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
              <option value="installation-guide">Монтажни етапи (printable checklist)</option>
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

function InstallationGuideModal({ room, constants, onClose }: {
  room: Room;
  constants: CalculatorConstants;
  onClose: () => void;
}) {
  function printGuide(): void {
    const printWindow = window.open("", "_blank");
    if (!printWindow) return;
    printWindow.document.open();
    printWindow.document.write(buildInstallationGuideHtml(room, constants));
    printWindow.document.close();
    printWindow.onload = () => {
      printWindow.focus();
      printWindow.print();
    };
    window.setTimeout(() => {
      if (!printWindow.closed) {
        printWindow.focus();
        printWindow.print();
      }
    }, 500);
  }

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="panel installation-modal" role="dialog" aria-modal="true" aria-labelledby="installation-guide-title">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Монтажни етапи</p>
            <h2 id="installation-guide-title">Монтажни етапи - {room.name}</h2>
          </div>
          <button type="button" className="ghost small" onClick={onClose}>Затвори</button>
        </div>
        <div className="installation-modal-body">
          <InstallationGuideContent room={room} constants={constants} mode="modal" />
        </div>
        <div className="modal-actions">
          <button type="button" className="ghost" onClick={printGuide}>Print</button>
          <button type="button" className="ghost" onClick={onClose}>Затвори</button>
        </div>
      </section>
    </div>
  );
}

function InstallationGuideContent({ room, constants, mode }: {
  room: Room;
  constants: CalculatorConstants;
  mode: "modal" | "print";
}) {
  const safeRoom = useMemo(() => cloneRoom(room), [room]);
  const result = useMemo(() => calc(safeRoom, constants), [safeRoom, constants]);
  const cutPlan = useMemo(() => buildSafeCutPlan(safeRoom, result, constants), [safeRoom, result, constants]);
  const guide = useMemo(() => buildInstallationGuide(safeRoom, result, constants, cutPlan.plan), [safeRoom, result, constants, cutPlan.plan]);

  return (
    <div className={`installation-guide ${mode}`}>
      <div className="installation-guide-grid">
        <InstallationSummary title="Данни" rows={guide.header} />
        <InstallationSummary title="Обобщение" rows={guide.summary} />
        <InstallationSummary title="Кратък разкрой" rows={getCutSummaryRows(cutPlan)} />
        <section className="installation-notes">
          <h3>Бележки</h3>
          <ul>
            {guide.notes.map((note) => <li key={note}>{note}</li>)}
          </ul>
        </section>
      </div>
      <section>
        <h3>Схема</h3>
        <ExistingRoomScheme room={safeRoom} result={result} constants={constants} />
      </section>
      <InstallationRoomCutPlan cutPlan={cutPlan} />
      <section>
        <h3>Checklist</h3>
        <InstallationCutPreparation step={guide.cutPreparationStep} />
        <InstallationStepList steps={guide.steps} />
      </section>
    </div>
  );
}

function InstallationSummary({ title, rows }: { title: string; rows: Array<[string, unknown]> }) {
  return (
    <section className="installation-summary">
      <h3>{title}</h3>
      <dl>
        {rows.map(([label, value]) => (
          <div key={String(label)}>
            <dt>{label}</dt>
            <dd>{String(value)}</dd>
          </div>
        ))}
      </dl>
    </section>
  );
}

function InstallationRoomCutPlan({ cutPlan }: { cutPlan: CutPlanState }) {
  if (cutPlan.error || !cutPlan.plan) {
    return (
      <section className="installation-cut-plan">
        <h3>Разкрой за стаята</h3>
        <div className="validation-item error">{cutPlan.error || "Разкроят не може да се изчисли."}</div>
      </section>
    );
  }

  return (
    <section className="installation-cut-plan">
      <h3>Разкрой за стаята</h3>
      <div className="cut-bar-groups">
        <CutBarGroup title="Разкрой CD профили" bars={cutPlan.plan.bars.filter(isCdCutBar)} />
        <CutBarGroup title="Разкрой UD профили" bars={cutPlan.plan.bars.filter(isUdCutBar)} />
      </div>
    </section>
  );
}

function InstallationCutPreparation({ step }: { step?: InstallationCutPreparationStep }) {
  if (!step?.enabled) return null;
  return (
    <details className="installation-cut-prep" open>
      <summary>Разкрой на профилите</summary>
      {step.cd ? (
        <section>
          <h4>CD профили</h4>
          <p>Общо профили CD за покупка: {step.cd.totalProfiles} бр.</p>
          <p>Профилите са оптимизирани така, че остатъците да се използват максимално ефективно.</p>
          <ul>
            {step.cd.bars.map((bar, index) => (
              <li key={bar.id}>Профил {index + 1}: {formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm))}</li>
            ))}
          </ul>
        </section>
      ) : null}
      {step.ud ? (
        <section>
          <h4>UD профили (периметър)</h4>
          <p>UD профилите могат да се комбинират свободно по периферията.</p>
          <ul>
            {step.ud.bars.map((bar, index) => (
              <li key={bar.id}>Профил {index + 1}: {formatCutBarLengths(bar.segments.map((segment) => segment.lengthCm))}</li>
            ))}
          </ul>
        </section>
      ) : null}
      <ul className="installation-cut-notes">
        {step.notes.map((note) => <li key={note}>{note}</li>)}
      </ul>
    </details>
  );
}

function InstallationStepList({ steps }: { steps: InstallationGuideStep[] }) {
  return (
    <div className="installation-step-list">
      {steps.map((step, index) => (
        <article key={step.title} className="installation-step">
          <span className="installation-checkbox" />
          <div>
            <strong>{index + 1}. {step.title}</strong>
            <ul>
              {step.items.map((item) => <li key={item}>{item}</li>)}
            </ul>
          </div>
        </article>
      ))}
    </div>
  );
}

function ExistingRoomScheme({ room, result, constants }: { room: Room; result: CalcResult; constants: CalculatorConstants }) {
  return (
    <div
      className="installation-scheme"
      dangerouslySetInnerHTML={{ __html: renderWorkingSchemeSvg(room, result, constants) }}
    />
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

function ResultTabs({ activeTab, onChange, warningCount }: {
  activeTab: RoomWorkspacePanel;
  onChange: (tab: RoomWorkspacePanel) => void;
  warningCount: number;
}) {
  const tabs: Array<{ key: RoomWorkspacePanel; label: string }> = [
    { key: "visual", label: "Схема" },
    { key: "materials", label: "Материали" },
    { key: "cut", label: "Разкрой" },
    { key: "installation", label: "Монтажни етапи" },
    { key: "validations", label: warningCount ? `Валидации (${warningCount})` : "Валидации" },
  ];
  return (
    <div className="result-tabs" role="tablist" aria-label="Резултати за активната стая">
      {tabs.map((tab) => (
        <button
          key={tab.key}
          type="button"
          role="tab"
          aria-selected={activeTab === tab.key}
          className={activeTab === tab.key ? "active" : ""}
          onClick={() => onChange(tab.key)}
        >
          {tab.label}
        </button>
      ))}
    </div>
  );
}

function RoomWorkspacePanelView({ activeTab, room, result, constants, rooms, zoom, warnings, activeCutPlan, globalCutPlan, cutOptimizationMode, onZoomChange, onCutOptimizationModeChange }: {
  activeTab: RoomWorkspacePanel;
  room: Room;
  result: CalcResult;
  constants: CalculatorConstants;
  rooms: Room[];
  zoom: number;
  warnings: ReturnType<typeof getValidationWarnings>;
  activeCutPlan: CutPlanState;
  globalCutPlan: CutPlanState;
  cutOptimizationMode: CutOptimizationMode;
  onZoomChange: (zoom: number) => void;
  onCutOptimizationModeChange: (mode: CutOptimizationMode) => void;
}) {
  if (activeTab === "materials") {
    return <RoomMaterialsPanel room={room} constants={constants} />;
  }
  if (activeTab === "cut") {
    return (
      <CutOptimizationPanel
        room={room}
        result={result}
        cutPlan={cutOptimizationMode === "room" ? activeCutPlan : globalCutPlan}
        constants={constants}
        mode={cutOptimizationMode}
        rooms={rooms}
        onModeChange={onCutOptimizationModeChange}
      />
    );
  }
  if (activeTab === "installation") {
    return (
      <section className="panel result-panel">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Монтаж</p>
            <h2>Монтажни етапи за {room.name}</h2>
          </div>
        </div>
        <InstallationGuideContent room={room} constants={constants} mode="modal" />
      </section>
    );
  }
  if (activeTab === "validations") {
    return <RoomValidationPanel warnings={warnings} room={room} />;
  }
  return (
    <Visualization
      room={room}
      result={result}
      constants={constants}
      zoom={zoom}
      onZoomChange={onZoomChange}
    />
  );
}

function RoomMaterialsPanel({ room, constants }: { room: Room; constants: CalculatorConstants }) {
  const rows = useMemo(() => buildMaterialTakeoff([room], constants), [room, constants]);
  return (
    <section className="panel result-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Материали</p>
          <h2>Материали за {room.name}</h2>
        </div>
      </div>
      <MaterialCardList rows={rows} rooms={[room]} constants={constants} />
    </section>
  );
}

function MaterialCardList({ rows, rooms, constants }: { rows: MaterialTakeoffItem[]; rooms: Room[]; constants: CalculatorConstants }) {
  const groups = [
    { title: "Профили", rows: rows.filter((row) => row.key.includes("cd-") || row.key.includes("ud-") || row.key.includes("ua-") || row.key.includes("uw-") || row.key.includes("battens")) },
    { title: "Крепежи", rows: rows.filter((row) => row.key.includes("anchors") || row.key.includes("hangers") || row.key.includes("connectors") || row.key.includes("screws") || row.key.includes("extensions")) },
    { title: "Гипсокартон", rows: rows.filter((row) => row.key.includes("boards")) },
    { title: "Фугиране", rows: rows.filter((row) => row.key.includes("joint") || row.key.includes("trenn")) },
    { title: "Изолация", rows: rows.filter((row) => row.key.includes("mineral")) },
  ].filter((group) => group.rows.length);

  return (
    <div className="material-groups">
      {groups.map((group) => (
        <section key={group.title} className="material-group-card">
          <h3>{group.title}</h3>
          <div className="material-card-list">
            {group.rows.map((row) => (
              <details key={row.key} className="material-card">
                <summary>
                  <span>{row.label}</span>
                  <strong>{formatMaterialQuantity(row)}</strong>
                </summary>
                <p>
                  {buildMaterialExplanation(row, rooms, constants)}
                  {(() => {
                    const catalogProduct = getCatalogProductForMaterial(row.label, row.key);
                    return catalogProduct ? ` Каталожен артикул: ${catalogProduct.knaufName}.` : "";
                  })()}
                </p>
              </details>
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}

function RoomValidationPanel({ warnings, room }: { warnings: ReturnType<typeof getValidationWarnings>; room: Room }) {
  return (
    <section className="panel result-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Валидации</p>
          <h2>Бележки за {room.name}</h2>
        </div>
      </div>
      {!warnings.length ? (
        <p className="empty-state">Няма активни предупреждения за тази стая.</p>
      ) : (
        <div className="validation-list visible-list">
          {warnings.map((warning) => (
            <div key={warning.code} className={`validation-item ${warning.severity}`}>
              <strong>{warning.severity === "error" ? "Грешка" : "Предупреждение"}</strong>
              <span>{warning.message}</span>
            </div>
          ))}
        </div>
      )}
    </section>
  );
}

function MobileActionBar({ onSave, onCut, onMaterials, onExport, canExport }: {
  onSave: () => void;
  onCut: () => void;
  onMaterials: () => void;
  onExport: () => void;
  canExport: boolean;
}) {
  return (
    <nav className="mobile-action-bar" aria-label="Бързи действия">
      <button type="button" onClick={onSave}>Запази</button>
      <button type="button" className="ghost" onClick={onCut}>Разкрой</button>
      <button type="button" className="ghost" onClick={onMaterials}>Материали</button>
      <button type="button" className="ghost" disabled={!canExport} onClick={onExport}>Експорт</button>
    </nav>
  );
}

function CutOptimizationPanel({ room, result, cutPlan, constants, mode, rooms, onModeChange }: {
  room: Room;
  result: CalcResult;
  cutPlan: CutPlanState;
  constants: CalculatorConstants;
  mode: CutOptimizationMode;
  rooms: Room[];
  onModeChange: (mode: CutOptimizationMode) => void;
}) {
  const isGlobalMode = mode === "global";
  const title = isGlobalMode ? "Общ разкрой за всички стаи" : `Разкрой за ${room.name}`;
  const savedRoomsProfileCounts = rooms.reduce((sum, savedRoom) => {
    const savedResult = calc(savedRoom, constants);
    sum.cd += savedResult.cdTotalProfiles;
    sum.ud += savedResult.udProfiles;
    return sum;
  }, { cd: 0, ud: 0 });
  const currentCdProfiles = isGlobalMode ? savedRoomsProfileCounts.cd : result.cdTotalProfiles;
  const currentUdProfiles = isGlobalMode ? savedRoomsProfileCounts.ud : result.udProfiles;
  const currentProfiles = currentCdProfiles + currentUdProfiles;
  const modeToggle = (
    <div className="cut-mode-toggle" role="group" aria-label="Режим на разкрой">
      <button
        type="button"
        className={mode === "room" ? "active" : ""}
        onClick={() => onModeChange("room")}
      >
        Разкрой по стая
      </button>
      <button
        type="button"
        className={mode === "global" ? "active" : ""}
        onClick={() => onModeChange("global")}
      >
        Общ разкрой за всички стаи
      </button>
    </div>
  );

  if (cutPlan.error || !cutPlan.plan) {
    return (
      <section className="panel cut-plan-panel">
        <div className="panel-title">
          <div>
            <p className="eyebrow">Оптимизация на разкроя</p>
            <h2>{title}</h2>
          </div>
        </div>
        {modeToggle}
        <div className="validation-item error">
          Разкроят не може да се изчисли с текущия layout и зададените дължини на профилите.
          {" "}
          {cutPlan.error}
        </div>
      </section>
    );
  }

  const plan = cutPlan.plan;
  const savedBars = Math.max(0, currentProfiles - plan.totalBars);
  const savedPercent = currentProfiles > 0
    ? Number(((savedBars / currentProfiles) * 100).toFixed(2))
    : 0;
  const totalPurchasedLengthCm = plan.bars.reduce((sum, bar) => sum + bar.stockLengthCm, 0);
  const stockLengthCm = plan.bars[0]?.stockLengthCm ?? Math.max(constants.cdLength, constants.udLength) * 100;
  const optimizedCdBars = plan.bars.filter(isCdCutBar).length;
  const optimizedUdBars = plan.bars.filter(isUdCutBar).length;
  const summarySubject = isGlobalMode ? "всички запазени стаи" : room.name;
  const summaryText = savedBars > 0
    ? `С този разкрой за ${summarySubject} спестяваш ${savedBars} стандартни профила (${formatNumber(savedPercent)} %) спрямо простото броене ${currentProfiles} бр., като подрежда CD профилите отделно от UD профилите.`
    : `За ${summarySubject} този разкрой не намалява броя нужни профили спрямо простото броене ${currentProfiles} бр., но подрежда CD и UD парчетата в отделни реални планове за рязане и показва остатъците.`;

  return (
    <section className="panel cut-plan-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Оптимизация на разкроя</p>
          <h2>{title}</h2>
        </div>
        <div className="cut-plan-summary-chip">
          <strong>{plan.totalBars} профила</strong>
          <small>стандартен прът {formatNumber(stockLengthCm)} cm, срез {0.3} cm, остатък над {20} cm се счита за използваем</small>
        </div>
      </div>
      {modeToggle}

      <div className="cut-plan-metrics">
        <div className="cut-metric">
          <span>Общо профили</span>
          <strong>{plan.totalBars} <small>(CD {optimizedCdBars} бр., UD {optimizedUdBars} бр.)</small></strong>
          <small>Това е колко стандартни пръта трябва да купиш за този разкрой.</small>
          <small>Сравнение с текущото просто броене: {currentProfiles} профила - CD {currentCdProfiles} бр., UD {currentUdProfiles} бр.</small>
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
        <small>
          Разкроят не променя конструктивната геометрия. Носещите CD редове остават възможно най-цели; снадки има само когато редът е по-дълъг от профила.
          {!isGlobalMode && room.systemType === "D113" ? " Късите монтажни CD сегменти при D113 са реални парчета между носещи редове/периметър." : ""}
          {isGlobalMode ? " При общия разкрой всяко парче пази името на стаята, от която идва." : ""}
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
        <CutBarGroup title="Разкрой CD профили" bars={plan.bars.filter(isCdCutBar)} showRoomNames={isGlobalMode} />
        <CutBarGroup title="Разкрой UD профили" bars={plan.bars.filter(isUdCutBar)} showRoomNames={isGlobalMode} />
      </div>
    </section>
  );
}

function getCutPieceLabel(type: CutBar["type"] | "waste"): string {
  if (type === "cd") return "CD профил";
  if (type === "carrier") return "носещ CD";
  if (type === "mounting") return "монтажен CD";
  if (type === "ud") return "UD";
  if (type === "mixed") return "смесен прът";
  return "остатък";
}

function isCdCutBar(bar: CutBar): boolean {
  return bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "carrier" || piece.type === "mounting");
}

function isUdCutBar(bar: CutBar): boolean {
  return bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "ud");
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

function CutBarGroup({ title, bars, showRoomNames = false }: { title: string; bars: CutBar[]; showRoomNames?: boolean }) {
  if (!bars.length) return null;
  return (
    <section className="cut-bar-group">
      <div className="cut-bar-group-head">
        <h3>{title}</h3>
        <small>{bars.length} пръта</small>
      </div>
      <div className="cut-bar-list">
        {bars.map((bar) => <CutBarStrip key={bar.id} bar={bar} showRoomNames={showRoomNames} />)}
      </div>
    </section>
  );
}

function CutBarStrip({ bar, showRoomNames = false }: { bar: CutBar; showRoomNames?: boolean }) {
  const barLabel = getCutPieceLabel(bar.type);
  return (
    <div className="cut-bar-card">
      <div className="cut-bar-head">
        <strong>{bar.id}</strong>
        <span>{barLabel}</span>
        <small>използвани {formatNumber(bar.usedCm)} / {formatNumber(bar.stockLengthCm)} cm</small>
      </div>
      <div className="cut-strip">
        {bar.segments.map((segment) => {
          const piece = bar.pieces.find((item) => item.id === segment.pieceId);
          const roomName = piece?.roomName;
          const segmentLabel = showRoomNames && roomName
            ? `${Math.round(segment.lengthCm)} cm · ${roomName}`
            : String(Math.round(segment.lengthCm));
          const title = showRoomNames && roomName
            ? `${getCutPieceLabel(segment.type)}: ${segment.lengthCm} cm · ${roomName}`
            : `${getCutPieceLabel(segment.type)}: ${segment.lengthCm} cm`;
          return (
            <div
              key={segment.pieceId}
              className={`cut-piece ${segment.type}`}
              style={{ width: `${(segment.lengthCm / bar.stockLengthCm) * 100}%` }}
              title={title}
            >
              {segmentLabel}
            </div>
          );
        })}
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
    <section className="panel room-editor-panel">
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

      <details className="settings-accordion" open>
        <summary>
          <span>Основни параметри</span>
          <small>Име, размери, площ и конструкция</small>
        </summary>
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
      </details>

      {warnings.length > 0 && (
        <div className="validation-list">
          {warnings.map((warning) => (
            <div key={warning.code} className={`validation-item ${warning.severity}`}>
              {warning.message}
            </div>
          ))}
        </div>
      )}

      <details className="settings-accordion" open>
        <summary>
          <span>Натоварване</span>
          <small>Товар, огнезащита, плоскости и окачвачи</small>
        </summary>
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
      </details>

      <details className="settings-accordion" open>
        <summary>
          <span>Разстояния</span>
          <small>a, b, c и периферни дюбели</small>
        </summary>
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
      </details>

      <details className="settings-accordion">
        <summary>
          <span>Допълнително</span>
          <small>Бележки и състояние на записа</small>
        </summary>
        <p className="hint">{variantHint}</p>
        <p className="hint">{udRule.note}</p>
      </details>
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
    <details className="settings-group" open>
      <summary className="settings-group-head">
        <h3>{title}</h3>
        <p>{note}</p>
      </summary>
      <div className="field-grid">{children}</div>
    </details>
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

function Visualization({ room, result, constants, zoom, onZoomChange }: {
  room: Room;
  result: CalcResult;
  constants: CalculatorConstants;
  zoom: number;
  onZoomChange: (zoom: number) => void;
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

function RoomsTable({ rooms, constants, activeRoomId, onAddRoom, onSelect, onDelete, onDeleteAll, onRecalculate, onOpenExport, onImportJson, onOpenInstallationGuide }: {
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
  onOpenInstallationGuide: (roomId: string) => void;
}) {
  const totalArea = rooms.reduce((sum, room) => sum + (Number(room.area) || 0), 0);

  return (
    <section className="panel table-panel">
      <div className="panel-title rooms-grid-toolbar">
        <div>
          <p className="eyebrow">Стаи</p>
          <h2>Количества по стаи</h2>
          <small>{rooms.length} стаи · {formatNumber(totalArea)} m2</small>
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
        <>
        <div className="table-scroll rooms-table-scroll">
          <table className="rooms-table">
            <colgroup>
              <col className="col-room" />
              <col className="col-system" />
              <col className="col-dimensions" />
              <col className="col-area" />
              <col className="col-small" />
              <col className="col-small" />
              <col className="col-small" />
              <col className="col-small" />
              <col className="col-compact" />
              <col className="col-small" />
              <col className="col-compact" />
              <col className="col-small" />
              <col className="col-compact" />
              <col className="col-small" />
              <col className="col-actions" />
            </colgroup>
            <thead>
              <tr>
                <th scope="col" className="sticky-room-col">Стая</th>
                <th scope="col">Система</th>
                <th scope="col" className="group-dimensions">Размери</th>
                <th scope="col" className="group-dimensions">m2</th>
                <th scope="col" className="group-structure">Носещи</th>
                <th scope="col" className="group-structure">Монтажни</th>
                <th scope="col" className="group-structure">CD</th>
                <th scope="col" className="group-structure">CD Опт</th>
                <th scope="col" className="group-structure">UD</th>
                <th scope="col" className="group-structure">UD Опт</th>
                <th scope="col" className="group-fasteners">Връзки</th>
                <th scope="col" className="group-fasteners">Окачвачи</th>
                <th scope="col" className="group-fasteners">Дюбели</th>
                <th scope="col" className="group-fasteners">Винтове</th>
                <th scope="col" className="group-actions">Действия</th>
              </tr>
            </thead>
            <tbody>
              {rooms.map((room) => {
                const result = calc(cloneRoom(room), constants);
                const cutPlan = buildSafeCutPlan(room, result, constants);
                const optimizedCdProfiles = cutPlan.plan ? cutPlan.plan.bars.filter(isCdCutBar).length : null;
                const optimizedUdProfiles = cutPlan.plan ? cutPlan.plan.bars.filter(isUdCutBar).length : null;
                return (
                    <tr key={room.id} className={room.id === activeRoomId ? "selected-row" : ""}>
                      <td className="sticky-room-col">
                        <button type="button" className="link-button room-name-button" onClick={() => onSelect(room.id)}>{room.name}</button>
                      </td>
                      <td><span className="system-pill">{room.systemType}</span></td>
                      <td>{room.width} x {room.length} cm</td>
                      <td>{formatNumber(room.area)}</td>
                      <td>{result.bearingCount}</td>
                      <td>{result.mountingCount}</td>
                      <td>{result.cdTotalProfiles}</td>
                      <td>{optimizedCdProfiles ?? "-"}</td>
                      <td>{result.udProfiles}</td>
                      <td>{optimizedUdProfiles ?? "-"}</td>
                      <td>{result.crossConnectors}</td>
                      <td>{result.hangersTotal}</td>
                      <td>{result.anchorsTotal}</td>
                      <td>{result.metalScrews + result.drywallScrews}</td>
                      <td>
                        <div className="row-actions">
                          <button type="button" className="ghost small" title="Монтажни етапи" aria-label={`Монтажни етапи за ${room.name}`} onClick={() => onOpenInstallationGuide(room.id)}>Етапи</button>
                          <button type="button" className="danger small" aria-label={`Изтрий ${room.name}`} onClick={() => onDelete(room.id)}>Изтрий</button>
                        </div>
                      </td>
                    </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="rooms-card-list">
          {rooms.map((room) => {
                const result = calc(cloneRoom(room), constants);
                const cutPlan = buildSafeCutPlan(room, result, constants);
                const optimizedCdProfiles = cutPlan.plan ? cutPlan.plan.bars.filter(isCdCutBar).length : null;
                const optimizedUdProfiles = cutPlan.plan ? cutPlan.plan.bars.filter(isUdCutBar).length : null;
                return (
              <article key={room.id} className={room.id === activeRoomId ? "room-card selected" : "room-card"}>
                <button type="button" className="link-button room-card-title" onClick={() => onSelect(room.id)}>{room.name}</button>
                <p>{room.systemType} · {room.width} x {room.length} cm · {formatNumber(room.area)} m2</p>
                <div className="room-card-metrics">
                  <span>CD: <strong>{result.cdTotalProfiles}</strong>{optimizedCdProfiles != null ? <small>опт. {optimizedCdProfiles}</small> : null}</span>
                  <span>Окачвачи: <strong>{result.hangersTotal}</strong></span>
                  <span>UD: <strong>{result.udProfiles}</strong>{optimizedUdProfiles != null ? <small>опт. {optimizedUdProfiles}</small> : null}</span>
                </div>
                <div className="row-actions">
                  <button type="button" className="ghost small" aria-label={`Монтажни етапи за ${room.name}`} onClick={() => onOpenInstallationGuide(room.id)}>Етапи</button>
                  <button type="button" className="danger small" aria-label={`Изтрий ${room.name}`} onClick={() => onDelete(room.id)}>Изтрий</button>
                </div>
              </article>
            );
          })}
        </div>
        </>
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
  const rows = useMemo(() => buildMaterialTakeoff(rooms, constants), [rooms, constants]);
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
      <MaterialCardList rows={rows} rooms={rooms} constants={constants} />
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
