const STORAGE_KEY = "d113-calculator-v2";
const isBrowser = typeof document !== "undefined";

const LOAD_CLASSES = ["0.15", "0.30", "0.40", "0.50", "0.65"];
const FIRE_LOAD_CLASSES = ["0.30", "0.40", "0.50", "0.65"];
const BOARD_TYPE_TO_B = {
  "12.5_silent": 400,
  "12.5_or_2x12.5": 500,
  "15_or_2x15": 550,
  "18_or_25_18": 625,
  "20_or_2x20": 625,
  "25": 800,
};

const CONSTRUCTION_TYPES = {
  D111: {
    label: "D111.bg дървена подконструкция",
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
};

function getConstruction(roomOrType = "D113") {
  const type = typeof roomOrType === "string" ? roomOrType : roomOrType.systemType;
  return CONSTRUCTION_TYPES[type] || CONSTRUCTION_TYPES.D113;
}

function getLoadClasses(systemType, fireProtection) {
  const construction = getConstruction(systemType);
  return fireProtection ? construction.fireLoadClasses : construction.loadClasses;
}

function getTableValue(room, c = room.c, loadClass = room.loadClass) {
  const classes = getLoadClasses(room.systemType, room.fireProtection);
  const idx = classes.indexOf(loadClass);
  if (idx === -1) return null;
  return getConstruction(room).table[String(room.fireProtection)]?.[c]?.[idx] || null;
}

function getValidCValues(room) {
  const classes = getLoadClasses(room.systemType, room.fireProtection);
  const idx = classes.indexOf(room.loadClass);
  const table = getConstruction(room).table[String(room.fireProtection)] || {};
  return Object.keys(table)
    .map(Number)
    .filter((c) => table[c]?.[idx] != null)
    .sort((a, b) => a - b);
}

function syncSpacingFromKnaufTable(room, { keepC = true } = {}) {
  const validCValues = getValidCValues(room);
  if (!keepC || !validCValues.includes(Number(room.c))) {
    room.c = validCValues[0] || room.c;
  }

  const a = getTableValue(room);
  if (a) room.a = a;

  if (!room.overrides?.b) {
    room.b = BOARD_TYPE_TO_B[room.boardType] || room.b;
  }
}

function countBySpacing(lengthCm, spacingMm, edgeCm) {
  if (!lengthCm || !spacingMm) return 0;
  const effectiveLength = Math.max(0, lengthCm - edgeCm);
  return Math.ceil(effectiveLength / (spacingMm / 10)) + 1;
}

const state = isBrowser ? loadState() : {
  rooms: [],
  constants: {
    cdLength: 4,
    udLength: 4,
    offset: 30,
    udAnchorSpacing: 625,
    metalScrewsPerCrossConnector: 4,
    metalScrewsPerDirectHanger: 2,
    drywallScrewsPerM2: 25,
    anchorsPerDirectHanger: 1,
    wastePercent: 10,
    roundingMode: "up",
  },
  activeRoomId: "",
};
if (isBrowser && !state.rooms.length) state.rooms.push(createRoom());
if (isBrowser && !state.activeRoomId) state.activeRoomId = state.rooms[0].id;

const el = isBrowser ? {
  constructionForm: document.getElementById("construction-form"),
  constantsForm: document.getElementById("constants-form"),
  cdLength: document.getElementById("const-cd-length"),
  udLength: document.getElementById("const-ud-length"),
  offset: document.getElementById("const-offset"),
  udAnchor: document.getElementById("const-ud-anchor"),
  metalCross: document.getElementById("const-metal-cross"),
  metalHanger: document.getElementById("const-metal-hanger"),
  drywallPerM2: document.getElementById("const-drywall-per-m2"),
  anchorsPerHanger: document.getElementById("const-anchors-per-hanger"),

  formTitle: document.getElementById("form-title"),
  roomId: document.getElementById("room-id"),
  name: document.getElementById("room-name"),
  width: document.getElementById("room-width"),
  length: document.getElementById("room-length"),
  area: document.getElementById("room-area"),
  load: document.getElementById("room-load"),
  fire: document.getElementById("room-fire"),
  system: document.getElementById("room-system"),
  board: document.getElementById("room-board"),
  a: document.getElementById("room-a"),
  b: document.getElementById("room-b"),
  c: document.getElementById("room-c"),
  constructionValidation: document.getElementById("construction-validation"),

  saveRoom: document.getElementById("save-room"),
  cancelRoom: document.getElementById("cancel-room"),
  tbody: document.querySelector("#rooms-table tbody"),
  tableFoot: document.getElementById("rooms-table-foot"),
  totalsPanel: document.getElementById("totals-panel"),
  reservePercent: document.getElementById("reserve-percent"),
  formulas: document.getElementById("formulas"),
  scheme: document.getElementById("scheme"),
  schemeLegend: document.getElementById("scheme-legend"),
  zoomIn: document.getElementById("zoom-in"),
  zoomOut: document.getElementById("zoom-out"),
  zoomReset: document.getElementById("zoom-reset"),
  openFullscreen: document.getElementById("open-fullscreen"),
  zoomIndicator: document.getElementById("zoom-indicator"),
} : {};

let areaDirty = false;
let schemeZoom = 1;

function loadState() {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    return {
      rooms: Array.isArray(raw.rooms) ? raw.rooms : [],
      constants: {
        cdLength: Number(raw.constants?.cdLength) || 4,
        udLength: Number(raw.constants?.udLength) || 4,
        offset: Number(raw.constants?.offset) || 30,
        udAnchorSpacing: Number(raw.constants?.udAnchorSpacing) || 625,
        metalScrewsPerCrossConnector: Number(raw.constants?.metalScrewsPerCrossConnector) || 4,
        metalScrewsPerDirectHanger: Number(raw.constants?.metalScrewsPerDirectHanger) || 2,
        drywallScrewsPerM2: Number(raw.constants?.drywallScrewsPerM2) || 25,
        anchorsPerDirectHanger: Number(raw.constants?.anchorsPerDirectHanger) || 1,
        wastePercent: Number(raw.constants?.wastePercent) || 10,
        roundingMode: raw.constants?.roundingMode || "up",
      },
      activeRoomId: raw.activeRoomId || "",
    };
  } catch {
    return {
      rooms: [],
      constants: {
        cdLength: 4,
        udLength: 4,
        offset: 30,
        udAnchorSpacing: 625,
        metalScrewsPerCrossConnector: 4,
        metalScrewsPerDirectHanger: 2,
        drywallScrewsPerM2: 25,
        anchorsPerDirectHanger: 1,
        wastePercent: 10,
        roundingMode: "up",
      },
      activeRoomId: "",
    };
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function createRoom() {
  const id = crypto.randomUUID();
  const construction = getConstruction("D113");
  const room = {
    id,
    name: "Стая",
    width: 400,
    length: 300,
    area: 12,
    systemType: "D113",
    loadClass: construction.defaultLoadClass,
    fireProtection: construction.defaultFireProtection,
    boardType: "12.5_or_2x12.5",
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

function normalizeRoom(room) {
  const construction = getConstruction(room.systemType || "D113");
  room.systemType = room.systemType || "D113";
  room.fireProtection = Boolean(room.fireProtection && construction.fireLoadClasses.length);
  room.overrides = { area: false, a: false, b: false, c: false, offset: false, udAnchorSpacing: false, ...room.overrides };
  if (!getLoadClasses(room.systemType, room.fireProtection).includes(room.loadClass)) {
    room.loadClass = construction.defaultLoadClass;
  }
  if (!Number(room.offset)) room.offset = Number(state.constants?.offset) || construction.defaultOffset;
  if (!Number(room.udAnchorSpacing)) room.udAnchorSpacing = Number(state.constants?.udAnchorSpacing) || construction.defaultUdAnchorSpacing;
  return room;
}

function applyAutoABC(room) {
  normalizeRoom(room);
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType);
  if (!room.overrides?.c) room.c = auto.c;
  if (!room.overrides?.a) room.a = auto.a;
  if (!room.overrides?.b) room.b = auto.b;
  if (!room.overrides?.offset) room.offset = auto.offset;
  if (!room.overrides?.udAnchorSpacing) room.udAnchorSpacing = auto.udAnchorSpacing;
}

function getAutoABC(loadClass, fireProtection, boardType, systemType = "D113") {
  const construction = getConstruction(systemType);
  const table = construction.table[String(fireProtection)];
  const classes = getLoadClasses(systemType, fireProtection);
  const idx = Math.max(0, classes.indexOf(loadClass));
  const sortedC = Object.keys(table).map(Number).sort((a, b) => a - b);
  const firstValid = sortedC.find((c) => table[c][idx] != null) || 600;
  return {
    a: table[firstValid][idx] || 900,
    c: firstValid,
    b: BOARD_TYPE_TO_B[boardType] || 500,
    offset: construction.defaultOffset,
    udAnchorSpacing: construction.defaultUdAnchorSpacing,
  };
}

function validateCombination(room) {
  normalizeRoom(room);
  const aExpected = getTableValue(room);
  const validB = Object.values(BOARD_TYPE_TO_B).includes(room.b);
  const validBForLoad = !(room.systemType === "D113"
    && room.fireProtection === false
    && room.c === 700
    && room.loadClass === "0.65"
    && room.b === 800);
  const validA = aExpected && room.a <= aExpected;
  return Boolean(validA && validB && validBForLoad);
}

function calc(room) {
  normalizeRoom(room);
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = Number(room.offset);

  const bearingCount = countBySpacing(W, room.c, offset);
  const bearingLengthTotal = bearingCount * (L / 100);
  const bearingProfiles = Math.ceil(bearingLengthTotal / state.constants.cdLength);

  const mountingCount = countBySpacing(L, room.b, offset);
  const mountingLengthTotal = mountingCount * (W / 100);
  const mountingProfiles = Math.ceil(mountingLengthTotal / state.constants.cdLength);

  const cdTotalLength = bearingLengthTotal + mountingLengthTotal;
  const cdTotalProfiles = bearingProfiles + mountingProfiles;

  const crossConnectors = bearingCount * mountingCount;
  const hangersPerBearing = countBySpacing(L, room.a, offset);
  const hangersTotal = bearingCount * hangersPerBearing;

  const udTotalLength = (2 * (X + Y)) / 100;
  const udProfiles = Math.ceil(udTotalLength / state.constants.udLength);

  const anchorsUd = Math.ceil(udTotalLength / (room.udAnchorSpacing / 1000));
  const anchorsHangers = hangersTotal * state.constants.anchorsPerDirectHanger;
  const anchorsTotal = anchorsUd + anchorsHangers;
  const metalScrews = Math.ceil(
    (crossConnectors * state.constants.metalScrewsPerCrossConnector)
    + (hangersTotal * state.constants.metalScrewsPerDirectHanger),
  );
  const drywallScrews = Math.ceil(Number(room.area) * state.constants.drywallScrewsPerM2);

  const extBearing = bearingCount * (Math.ceil((L / 100) / state.constants.cdLength) - 1);
  const extMounting = mountingCount * (Math.ceil((W / 100) / state.constants.cdLength) - 1);
  const extensionsTotal = extBearing + extMounting;

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
    extensionsTotal,
  };
}

function buildPositions(limitCm, spacingMm, edgeCm = 30) {
  const count = countBySpacing(limitCm, spacingMm, edgeCm);
  if (!count) return [];
  if (count === 1) return [Math.min(edgeCm, limitCm / 2)];

  const start = Math.min(edgeCm, limitCm / 2);
  const end = Math.max(start, limitCm - edgeCm);
  const actualSpacing = (end - start) / (count - 1);
  return Array.from({ length: count }, (_, idx) => start + idx * actualSpacing);
}

function updateZoomUI() {
  el.scheme.style.setProperty("--scheme-zoom", String(schemeZoom));
  el.zoomIndicator.textContent = `${Math.round(schemeZoom * 100)}%`;
}

function setZoom(nextZoom) {
  schemeZoom = Math.min(2.5, Math.max(0.6, nextZoom));
  updateZoomUI();
}

function render() {
  el.cdLength.value = state.constants.cdLength;
  el.udLength.value = state.constants.udLength;
  el.metalCross.value = state.constants.metalScrewsPerCrossConnector;
  el.metalHanger.value = state.constants.metalScrewsPerDirectHanger;
  el.drywallPerM2.value = state.constants.drywallScrewsPerM2;
  el.anchorsPerHanger.value = state.constants.anchorsPerDirectHanger;
  el.reservePercent.value = state.constants.wastePercent;
  renderTable();
  renderTotals();
  const active = state.rooms.find((r) => r.id === state.activeRoomId);
  if (active) {
    normalizeRoom(active);
    bindRoomToForm(active);
    renderScheme(active);
    renderFormulas(active);
  }
  updateZoomUI();
  saveState();
}

function bindRoomToForm(room) {
  normalizeRoom(room);
  el.formTitle.textContent = "Стаи";
  el.roomId.value = room.id;
  el.name.value = room.name;
  el.width.value = room.width;
  el.length.value = room.length;
  el.area.value = room.area;
  el.load.value = room.loadClass;
  el.fire.value = String(room.fireProtection);
  el.system.value = room.systemType;
  el.board.value = room.boardType;
  el.offset.value = room.offset;
  el.udAnchor.value = room.udAnchorSpacing;
  renderLoadOptions(room);
  el.a.value = room.a;
  el.b.value = room.b;
  el.c.value = room.c;
  const validationMessage = validateCombination(room) ? "" : `Невалидна комбинация според Knauf ${room.systemType}`;
  el.constructionValidation.textContent = validationMessage;
}

function renderLoadOptions(room) {
  const classes = getLoadClasses(room.systemType, room.fireProtection);
  el.load.innerHTML = classes
    .map((value) => `<option value="${value}">до ${value} kN/m²</option>`)
    .join("");
  el.load.value = room.loadClass;
  el.fire.disabled = !getConstruction(room).fireLoadClasses.length;
}

function renderTable() {
  el.tbody.innerHTML = "";
  const totals = {
    area: 0,
    bearingCount: 0,
    mountingCount: 0,
    bearingLengthTotal: 0,
    mountingLengthTotal: 0,
    cdTotalProfiles: 0,
    udProfiles: 0,
    crossConnectors: 0,
    hangersTotal: 0,
    anchorsUd: 0,
    anchorsHangers: 0,
    anchorsTotal: 0,
    metalScrews: 0,
    drywallScrews: 0,
    extensionsTotal: 0,
  };
  state.rooms.forEach((room) => {
    const r = calc(room);
    totals.area += Number(room.area);
    totals.bearingCount += r.bearingCount;
    totals.mountingCount += r.mountingCount;
    totals.bearingLengthTotal += r.bearingLengthTotal;
    totals.mountingLengthTotal += r.mountingLengthTotal;
    totals.cdTotalProfiles += r.cdTotalProfiles;
    totals.udProfiles += r.udProfiles;
    totals.crossConnectors += r.crossConnectors;
    totals.hangersTotal += r.hangersTotal;
    totals.anchorsUd += r.anchorsUd;
    totals.anchorsHangers += r.anchorsHangers;
    totals.anchorsTotal += r.anchorsTotal;
    totals.metalScrews += r.metalScrews;
    totals.drywallScrews += r.drywallScrews;
    totals.extensionsTotal += r.extensionsTotal;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><a href="#" class="room-link" data-id="${room.id}" data-action="edit">${room.name}</a></td><td>${room.systemType}</td><td>${room.width}</td><td>${room.length}</td><td>${Number(room.area).toFixed(2)}</td><td>${r.bearingCount}</td><td>${r.mountingCount}</td><td>${r.bearingLengthTotal.toFixed(2)}</td><td>${r.mountingLengthTotal.toFixed(2)}</td><td>${r.cdTotalProfiles}</td><td>${r.udProfiles}</td><td>${r.crossConnectors}</td><td>${r.hangersTotal}</td><td>${r.anchorsUd}</td><td>${r.anchorsHangers}</td><td>${r.anchorsTotal}</td><td>${r.metalScrews}</td><td>${r.drywallScrews}</td><td>${r.extensionsTotal}</td><td class="actions"><button data-id="${room.id}" data-action="edit" title="Редакция">✏️</button><button data-id="${room.id}" data-action="del" class="danger" title="Изтрий">🗑️</button></td>`;
    el.tbody.appendChild(tr);
  });
  el.tableFoot.innerHTML = `
    <tr class="total-row">
      <td><strong>Total</strong></td>
      <td>—</td>
      <td>—</td>
      <td>—</td>
      <td><strong>${totals.area.toFixed(2)}</strong></td>
      <td><strong>${totals.bearingCount}</strong></td>
      <td><strong>${totals.mountingCount}</strong></td>
      <td><strong>${totals.bearingLengthTotal.toFixed(2)}</strong></td>
      <td><strong>${totals.mountingLengthTotal.toFixed(2)}</strong></td>
      <td><strong>${totals.cdTotalProfiles}</strong></td>
      <td><strong>${totals.udProfiles}</strong></td>
      <td><strong>${totals.crossConnectors}</strong></td>
      <td><strong>${totals.hangersTotal}</strong></td>
      <td><strong>${totals.anchorsUd}</strong></td>
      <td><strong>${totals.anchorsHangers}</strong></td>
      <td><strong>${totals.anchorsTotal}</strong></td>
      <td><strong>${totals.metalScrews}</strong></td>
      <td><strong>${totals.drywallScrews}</strong></td>
      <td><strong>${totals.extensionsTotal}</strong></td>
      <td>—</td>
    </tr>
  `;
}

function renderTotals() {
  const reservePercent = Number(state.constants.wastePercent) || 0;
  const multiplier = 1 + reservePercent / 100;
  const totals = state.rooms.reduce((acc, room) => {
    const r = calc(room);
    acc.cd += r.cdTotalProfiles;
    acc.ud += r.udProfiles;
    acc.connectors += r.crossConnectors;
    acc.hangers += r.hangersTotal;
    acc.anchors += r.anchorsTotal;
    acc.screws += r.metalScrews + r.drywallScrews;
    acc.extensions += r.extensionsTotal;
    return acc;
  }, {
    cd: 0, ud: 0, connectors: 0, hangers: 0, anchors: 0, screws: 0, extensions: 0,
  });

  const withReserve = (value) => Math.ceil(value * multiplier);
  el.totalsPanel.innerHTML = `
    <div class="totals-grid">
      <div><strong>Изчислен резерв CD бр.:</strong> ${withReserve(totals.cd)}</div>
      <div><strong>Изчислен резерв UD бр.:</strong> ${withReserve(totals.ud)}</div>
      <div><strong>Изчислен резерв Връзки:</strong> ${withReserve(totals.connectors)}</div>
      <div><strong>Изчислен резерв Окачвачи:</strong> ${withReserve(totals.hangers)}</div>
      <div><strong>Изчислен резерв Дюбели:</strong> ${withReserve(totals.anchors)}</div>
      <div><strong>Изчислен резерв Винтове:</strong> ${withReserve(totals.screws)}</div>
      <div><strong>Изчислен резерв Удължители:</strong> ${withReserve(totals.extensions)}</div>
    </div>
  `;
}

function renderScheme(room) {
  normalizeRoom(room);
  const r = calc(room);
  const bearingLinePositionsCm = buildPositions(r.W, room.c, room.offset);
  const mountingLinePositionsCm = buildPositions(r.L, room.b, room.offset);
  const hangerPositionsCm = buildPositions(r.L, room.a, room.offset);
  const extensionPoints = [];
  const pad = 40;
  const w = 760;
  const h = 300;
  const xScale = w / r.L;
  const yScale = h / r.W;
  const bearingStroke = 2;
  const hangerRadius = 4.5;

  let svg = `<rect x="${pad}" y="${pad}" width="${w}" height="${h}" fill="#eef6ff" stroke="#13588f" stroke-width="2" />`;

  // Размери в краищата на стаята
  svg += `<text x="${pad}" y="${pad - 10}" fill="#17324d" font-size="11">0 см</text>`;
  svg += `<text x="${pad + w - 48}" y="${pad - 10}" fill="#17324d" font-size="11">${r.L.toFixed(0)} см</text>`;
  svg += `<text x="${pad - 34}" y="${pad + 4}" fill="#17324d" font-size="11">0 см</text>`;
  svg += `<text x="${pad - 42}" y="${pad + h}" fill="#17324d" font-size="11">${r.W.toFixed(0)} см</text>`;

  mountingLinePositionsCm.forEach((p) => {
    const x = pad + p * xScale;
    svg += `<line x1="${x}" y1="${pad}" x2="${x}" y2="${pad + h}" stroke="#2b9a42" stroke-width="1.5"/>`;
    svg += `<text x="${x + 2}" y="${pad + 12}" fill="#2b9a42" font-size="10">${p.toFixed(0)}см</text>`;
  });

  bearingLinePositionsCm.forEach((p) => {
    const y = pad + p * yScale;
    svg += `<line x1="${pad}" y1="${y}" x2="${pad + w}" y2="${y}" stroke="#1f5e93" stroke-width="${bearingStroke}"/>`;
    svg += `<text x="${pad + 2}" y="${y - 3}" fill="#1f5e93" font-size="10">${p.toFixed(0)}см</text>`;

    // Позиции на окачвачите върху носещите профили
    hangerPositionsCm.forEach((hp, idx) => {
      const hx = pad + hp * xScale;
      const labelYOffset = idx % 2 === 0 ? -9 : 15;
      svg += `<circle cx="${hx}" cy="${y}" r="${hangerRadius}" fill="#1b1b1b"/>`;
      svg += `<text x="${hx}" y="${y + labelYOffset}" fill="#1b1b1b" font-size="8.5" text-anchor="middle" stroke="#f7fbff" stroke-width="2" paint-order="stroke">${hp.toFixed(0)}см</text>`;
    });
  });



  const segmentLengthCm = state.constants.cdLength * 100;
  if (r.L > segmentLengthCm) {
    for (let pos = segmentLengthCm; pos < r.L; pos += segmentLengthCm) extensionPoints.push(pos);

    bearingLinePositionsCm.forEach((p) => {
      const y = pad + p * yScale;
      extensionPoints.forEach((ep, idx) => {
        const x = pad + ep * xScale;
        const textOffset = idx % 2 === 0 ? -11 : 16;
        svg += `<line x1="${x}" y1="${y - 8}" x2="${x}" y2="${y + 8}" stroke="#d14a00" stroke-width="2"/>`;
        svg += `<text x="${x}" y="${y + textOffset}" fill="#d14a00" font-size="8.5" text-anchor="middle" stroke="#f7fbff" stroke-width="2" paint-order="stroke">${ep.toFixed(0)}см</text>`;
      });
    });
  }

  hangerPositionsCm.forEach((hp) => {
    const hx = pad + hp * xScale;
    svg += `<text x="${hx - 10}" y="${pad + h + 14}" fill="#1b1b1b" font-size="10">${hp.toFixed(0)}</text>`;
  });

  el.scheme.innerHTML = svg;
  const legendRows = [
    ["W (широчина)", `${r.W.toFixed(0)} cm`],
    ["L (дължина)", `${r.L.toFixed(0)} cm`],
    ["Конструкция", `${room.systemType} - ${getConstruction(room).label}`],
    ["Клас на натоварване", `${room.loadClass} kN/m²`],
    ["a (Разстояние между окачвачите)", `${room.a} mm`],
    ["b (Разстояние между монтажните профили)", `${room.b} mm`],
    ["c (Разстояние между носещите профили)", `${room.c} mm`],
    ["o (първоначален offset)", `${room.offset} cm`],
    ["Елементи в схемата", `
      <div class="legend-swatch-row"><span class="legend-swatch line mounting"></span><span>Монтажен CD профил</span></div>
      <div class="legend-swatch-row"><span class="legend-swatch line bearing"></span><span>Носещ CD профил</span></div>
      <div class="legend-swatch-row"><span class="legend-swatch point hanger"></span><span>Окачвач (позиция x см от началото на стената)</span></div>
      <div class="legend-swatch-row"><span class="legend-swatch line extension"></span><span>Удължител (вертикална позиция x см от стената)</span></div>
    `],
    ["Носещи", formatLegendPositions(bearingLinePositionsCm)],
    ["Монтажни", formatLegendPositions(mountingLinePositionsCm)],
    ["Окачвачи", formatLegendPositions(hangerPositionsCm)],
    ["Удължители", formatLegendPositions(extensionPoints)],
  ];
  el.schemeLegend.innerHTML = legendRows
    .map(([title, value]) => `<div class="legend-row"><span class="legend-title">${title}:</span> <span>${value}</span></div>`)
    .join("");
}

function f2(value) {
  return Number(value).toFixed(2);
}

function formatLegendPositions(positions) {
  if (!positions.length) return "няма";
  return positions.map((p) => p.toFixed(0)).join(", ");
}

function withBoldResult(text) {
  const [left, ...rightParts] = text.split(" = ");
  if (!rightParts.length) return text;
  return `<span class="formula-result">${left}</span> = ${rightParts.join(" = ")}`;
}

function renderFormulas(room) {
  normalizeRoom(room);
  const r = calc(room);
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = Number(room.offset);
  const udAnchorSpacingMm = Number(room.udAnchorSpacing);

  const rows = [
    {
      title: "Изчисления за количествата",
      items: [
        `W = min(X, Y) = min(${X}, ${Y}) = ${W} cm`,
        `L = max(X, Y) = max(${X}, ${Y}) = ${L} cm`,
        `Носещи CD редове = ceil((W - o) / (c / 10)) + 1 = ceil((${W} - ${offset}) / (${room.c} / 10)) + 1 = ${r.bearingCount}`,
        `Монтажни CD редове = ceil((L - o) / (b / 10)) + 1 = ceil((${L} - ${offset}) / (${room.b} / 10)) + 1 = ${r.mountingCount}`,
        `Носещи CD метри = Носещи редове × (L / 100) = ${r.bearingCount} × (${L} / 100) = ${f2(r.bearingLengthTotal)} m`,
        `Монтажни CD метри = Монтажни редове × (W / 100) = ${r.mountingCount} × (${W} / 100) = ${f2(r.mountingLengthTotal)} m`,
        `CD бр. = ceil(Носещи m / CD дължина) + ceil(Монтажни m / CD дължина) = ceil(${f2(r.bearingLengthTotal)} / ${state.constants.cdLength}) + ceil(${f2(r.mountingLengthTotal)} / ${state.constants.cdLength}) = ${r.cdTotalProfiles}`,
        `UD дължина = 2 × (X + Y) / 100 = 2 × (${X} + ${Y}) / 100 = ${f2(r.udTotalLength)} m`,
        `UD бр. = ceil(UD дължина / UD дължина профил) = ceil(${f2(r.udTotalLength)} / ${state.constants.udLength}) = ${r.udProfiles}`,
        `Връзки = Носещи редове × Монтажни редове = ${r.bearingCount} × ${r.mountingCount} = ${r.crossConnectors}`,
        `Окачвачи/носещ = ceil((L - o) / (a / 10)) + 1 = ceil((${L} - ${offset}) / (${room.a} / 10)) + 1 = ${r.hangersPerBearing}`,
        `Окачвачи общо = Носещи редове × Окачвачи/носещ = ${r.bearingCount} × ${r.hangersPerBearing} = ${r.hangersTotal}`,
        `udAnchorSpacingMm = ${udAnchorSpacingMm}`,
        `udAnchors = ceil(udTotalLength / (udAnchorSpacingMm / 1000)) = ceil(${f2(r.udTotalLength)} / (${udAnchorSpacingMm} / 1000)) = ${r.anchorsUd}`,
        `Дюбели окачвачи = Окачвачи × дюбели/окачвач = ${r.hangersTotal} × ${state.constants.anchorsPerDirectHanger} = ${r.anchorsHangers}`,
        `Дюбели общо = Дюбели UD + Дюбели окачвачи = ${r.anchorsUd} + ${r.anchorsHangers} = ${r.anchorsTotal}`,
        `Винтове метал = ceil((Връзки × ${state.constants.metalScrewsPerCrossConnector}) + (Окачвачи × ${state.constants.metalScrewsPerDirectHanger})) = ${r.metalScrews}`,
        `Винтове гипсокартон = ceil(Площ × ${state.constants.drywallScrewsPerM2}) = ceil(${f2(room.area)} × ${state.constants.drywallScrewsPerM2}) = ${r.drywallScrews}`,
        `Удължители = Носещи удълж. + Монтажни удълж. = ${r.extensionsTotal}`,
      ],
    },
    {
      title: "Формули за изчертаване на схемата",
      items: [
        `Позиции на носещи линии (cm) = ${r.bearingCount} позиции, разпределени от ${offset} cm до ${W - offset} cm`,
        `Позиции на монтажни линии (cm) = ${r.mountingCount} позиции, разпределени от ${offset} cm до ${L - offset} cm`,
        `Позиции на окачвачи (cm) = ${r.hangersPerBearing} позиции, разпределени от ${offset} cm до ${L - offset} cm`,
        `Хоризонтален мащаб = 760 / L = 760 / ${L} = ${f2(760 / L)}`,
        `Вертикален мащаб = 300 / W = 300 / ${W} = ${f2(300 / W)}`,
        `Координати в SVG = x = pad + позиция × хоризонтален мащаб, y = pad + позиция × вертикален мащаб (pad = 40)`,
      ],
    },
  ];

  el.formulas.innerHTML = `
    <h3>Формули за активната стая: ${room.name}</h3>
    ${rows.map((group) => `
      <div class="formulas-group">
        <h4>${group.title}</h4>
        ${group.items.map((item) => `<div class="formula-row">${withBoldResult(item)}</div>`).join("")}
      </div>
    `).join("")}
  `;
}

function updateRoomFromForm() {
  const room = state.rooms.find((r) => r.id === el.roomId.value);
  if (!room) return;
  room.name = el.name.value.trim() || "Стая";
  room.width = Number(el.width.value);
  room.length = Number(el.length.value);
  room.area = Number(el.area.value);
  room.systemType = el.system.value;
  room.loadClass = el.load.value;
  room.fireProtection = el.fire.value === "true";
  room.boardType = el.board.value;
  room.a = Number(el.a.value);
  room.b = Number(el.b.value);
  room.c = Number(el.c.value);
  room.offset = Number(el.offset.value);
  room.udAnchorSpacing = Number(el.udAnchor.value);

  applyAutoABC(room);
  if (!room.overrides.area && room.width && room.length) room.area = (room.width * room.length) / 10000;
  const validationMessage = validateCombination(room) ? "" : `Невалидна комбинация според Knauf ${room.systemType}`;
  el.constructionValidation.textContent = validationMessage;
}

function handleLoadClassChange() {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (!room) return;
  room.loadClass = el.load.value;
  if (!getLoadClasses(room.systemType, room.fireProtection).includes(room.loadClass)) {
    room.loadClass = getLoadClasses(room.systemType, room.fireProtection)[0];
    el.load.value = room.loadClass;
  }
  room.overrides.a = false;
  syncSpacingFromKnaufTable(room, { keepC: true });
  render();
}

if (isBrowser) {
[el.width, el.length].forEach((input) => input.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (!room) return;
  room.width = Number(el.width.value);
  room.length = Number(el.length.value);
  if (!areaDirty && room.width && room.length) {
    room.area = (room.width * room.length) / 10000;
    el.area.value = room.area.toFixed(2);
  }
  render();
}));

el.area.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (!room) return;
  room.overrides.area = true;
  areaDirty = true;
  room.area = Number(el.area.value);
  render();
});

el.load.addEventListener("change", handleLoadClassChange);
el.load.addEventListener("input", handleLoadClassChange);

el.system.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (!room) return;
  room.systemType = el.system.value;
  const construction = getConstruction(room);
  room.fireProtection = construction.defaultFireProtection;
  room.loadClass = construction.defaultLoadClass;
  room.overrides.a = false;
  room.overrides.b = false;
  room.overrides.c = false;
  room.overrides.offset = false;
  room.overrides.udAnchorSpacing = false;
  applyAutoABC(room);
  syncSpacingFromKnaufTable(room, { keepC: false });
  render();
});

el.fire.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.fireProtection = el.fire.value === "true";
  if (!getConstruction(room).fireLoadClasses.length) room.fireProtection = false;
  if (!getLoadClasses(room.systemType, room.fireProtection).includes(room.loadClass)) {
    room.loadClass = getLoadClasses(room.systemType, room.fireProtection)[0];
    el.load.value = room.loadClass;
  }
  room.overrides.a = false;
  syncSpacingFromKnaufTable(room, { keepC: true });
  render();
});

el.board.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.boardType = el.board.value;
  room.overrides.b = el.board.value === "custom";
  syncSpacingFromKnaufTable(room, { keepC: true });
  render();
});

el.a.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.a = true;
  room.a = Number(el.a.value);
  render();
});
el.a.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.a = true;
  room.a = Number(el.a.value);
  render();
});
el.b.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.b = true;
  room.b = Number(el.b.value);
  render();
});
el.b.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.b = true;
  room.b = Number(el.b.value);
  render();
});
el.c.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.c = true;
  room.c = Number(el.c.value);
  room.overrides.a = false;
  syncSpacingFromKnaufTable(room, { keepC: true });
  render();
});
el.c.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.c = true;
  room.c = Number(el.c.value);
  room.overrides.a = false;
  syncSpacingFromKnaufTable(room, { keepC: true });
  render();
});

el.offset.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.offset = true;
  room.offset = Number(el.offset.value);
  render();
});

el.udAnchor.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.udAnchorSpacing = true;
  room.udAnchorSpacing = Number(el.udAnchor.value);
  render();
});

el.constantsForm.addEventListener("input", () => {
  state.constants.cdLength = Number(el.cdLength.value);
  state.constants.udLength = Number(el.udLength.value);
  state.constants.metalScrewsPerCrossConnector = Number(el.metalCross.value);
  state.constants.metalScrewsPerDirectHanger = Number(el.metalHanger.value);
  state.constants.drywallScrewsPerM2 = Number(el.drywallPerM2.value);
  state.constants.anchorsPerDirectHanger = Number(el.anchorsPerHanger.value);
  state.constants.roundingMode = "up";
  render();
});

document.getElementById("new-room").addEventListener("click", () => {
  const room = createRoom();
  room.name = `Стая ${state.rooms.length + 1}`;
  state.rooms.push(room);
  state.activeRoomId = room.id;
  areaDirty = false;
  render();
});

document.getElementById("save-room").addEventListener("click", () => {
  updateRoomFromForm();
  render();
});

document.getElementById("cancel-room").addEventListener("click", () => {
  el.name.value = "";
  el.width.value = "";
  el.length.value = "";
  el.area.value = "";
  el.constructionValidation.textContent = "";
  areaDirty = false;
});

el.tbody.addEventListener("click", (event) => {
  const actionElement = event.target.closest("button, a[data-action]");
  if (!actionElement) return;
  if (actionElement.tagName === "A") event.preventDefault();
  const id = actionElement.dataset.id;
  const room = state.rooms.find((item) => item.id === id);
  if (!room) return;
  if (actionElement.dataset.action === "edit") {
    state.activeRoomId = id;
    areaDirty = !!room.overrides?.area;
    render();
    return;
  }
  state.rooms = state.rooms.filter((item) => item.id !== id);
  if (!state.rooms.length) state.rooms.push(createRoom());
  state.activeRoomId = state.rooms[0].id;
  render();
});

document.getElementById("export-json").addEventListener("click", () => {
  const payload = {
    rooms: state.rooms.map((room) => ({
      name: room.name,
      width: room.width,
      length: room.length,
      area: room.area,
      systemType: room.systemType,
      a: room.a,
      b: room.b,
      c: room.c,
      offset: room.offset,
      udAnchorSpacing: room.udAnchorSpacing,
      loadClass: room.loadClass,
      fireProtection: room.fireProtection,
      boardType: room.boardType,
      overrides: room.overrides,
      results: calc(room),
    })),
    constants: state.constants,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "knauf-d113-export.json";
  link.click();
  URL.revokeObjectURL(link.href);
});



function xmlEscape(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function exportRoomsToExcel() {
  const headers = [
    "Стая", "Конструкция", "X", "Y", "Площ", "Носещи CD (бр реда)", "Монтажни CD (бр реда)",
    "Носещи m", "Монтажни m", "CD бр.", "UD бр.", "Връзки", "Окачвачи",
    "Дюбели UD", "Дюбели окачвачи", "Дюбели общо", "Винтове метал", "Винтове гипсокартон", "Удължители",
  ];

  const rows = state.rooms.map((room) => {
    const r = calc(room);
    return [
      room.name,
      room.systemType,
      room.width,
      room.length,
      Number(room.area).toFixed(2),
      r.bearingCount,
      r.mountingCount,
      r.bearingLengthTotal.toFixed(2),
      r.mountingLengthTotal.toFixed(2),
      r.cdTotalProfiles,
      r.udProfiles,
      r.crossConnectors,
      r.hangersTotal,
      r.anchorsUd,
      r.anchorsHangers,
      r.anchorsTotal,
      r.metalScrews,
      r.drywallScrews,
      r.extensionsTotal,
    ];
  });

  const sheetRows = [headers, ...rows]
    .map((row) => `<Row>${row.map((cell) => `<Cell><Data ss:Type="String">${xmlEscape(cell)}</Data></Cell>`).join("")}</Row>`)
    .join("");

  const workbook = `<?xml version="1.0" encoding="UTF-8"?>
  <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
    <Worksheet ss:Name="Стаи">
      <Table>${sheetRows}</Table>
    </Worksheet>
  </Workbook>`;

  const blob = new Blob([workbook], { type: "application/vnd.ms-excel" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "knauf-d113-stai.xls";
  link.click();
  URL.revokeObjectURL(link.href);
}


document.getElementById("export-excel").addEventListener("click", exportRoomsToExcel);

document.getElementById("import-json").addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const raw = JSON.parse(await file.text());
    if (!Array.isArray(raw.rooms)) throw new Error("Очаква се масив rooms.");
    state.rooms = raw.rooms.map((room) => ({ ...createRoom(), ...room, id: crypto.randomUUID() }));
    if (raw.constants) state.constants = { ...state.constants, ...raw.constants };
    state.activeRoomId = state.rooms[0]?.id || "";
    render();
    alert("Импортът е успешен.");
  } catch (error) {
    alert(`Грешка: ${error.message}`);
  } finally {
    event.target.value = "";
  }
});

el.zoomIn.addEventListener("click", () => setZoom(schemeZoom + 0.2));
el.zoomOut.addEventListener("click", () => setZoom(schemeZoom - 0.2));
el.zoomReset.addEventListener("click", () => setZoom(1));
el.openFullscreen.addEventListener("click", () => {
  const svgMarkup = el.scheme.outerHTML;
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.write(`<!doctype html>
<html lang="bg">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Схема на цял екран</title>
  <style>
    body { margin: 0; background: #0b1a2b; display: flex; justify-content: center; align-items: center; min-height: 100vh; }
    svg { width: 100vw; height: 100vh; background: #f7fbff; }
  </style>
</head>
<body>${svgMarkup}</body>
</html>`);
  win.document.close();
});

el.scheme.addEventListener("wheel", (event) => {
  if (!event.ctrlKey) return;
  event.preventDefault();
  const delta = event.deltaY < 0 ? 0.1 : -0.1;
  setZoom(schemeZoom + delta);
}, { passive: false });

document.getElementById("clear-all").addEventListener("click", () => {
  state.rooms = [createRoom()];
  state.activeRoomId = state.rooms[0].id;
  render();
});

el.reservePercent.addEventListener("input", () => {
  state.constants.wastePercent = Number(el.reservePercent.value) || 0;
  render();
});

render();
}

if (typeof module !== "undefined") {
  module.exports = {
    BOARD_TYPE_TO_B,
    CONSTRUCTION_TYPES,
    calc,
    countBySpacing,
    getTableValue,
    getValidCValues,
    syncSpacingFromKnaufTable,
    validateCombination,
  };
}
