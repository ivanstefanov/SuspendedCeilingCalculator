const STORAGE_KEY = "d113-calculator-v2";

const LOAD_CLASSES = ["0.15", "0.30", "0.40", "0.50", "0.65"];
const BOARD_TYPE_TO_B = {
  "12.5_silent": 400,
  "12.5_or_2x12.5": 500,
  "15_or_2x15": 550,
  "18_or_25_18": 625,
  "20_or_2x20": 625,
  "25": 800,
};

const KNAUF_TABLE = {
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
    500: [850, 750, 700, 600, null],
    600: [800, 700, 650, 550, null],
    700: [750, 650, 600, 550, null],
    800: [700, 650, 600, null, null],
    900: [700, 600, 550, null, null],
    1000: [650, 600, 550, null, null],
    1100: [650, 600, 550, null, null],
    1200: [600, 550, null, null, null],
    1250: [600, null, null, null, null],
  },
};

const state = loadState();
if (!state.rooms.length) state.rooms.push(createRoom());
if (!state.activeRoomId) state.activeRoomId = state.rooms[0].id;

const el = {
  constantsForm: document.getElementById("constants-form"),
  cdLength: document.getElementById("const-cd-length"),
  udLength: document.getElementById("const-ud-length"),
  offset: document.getElementById("const-offset"),
  udAnchor: document.getElementById("const-ud-anchor"),

  formTitle: document.getElementById("form-title"),
  roomId: document.getElementById("room-id"),
  name: document.getElementById("room-name"),
  width: document.getElementById("room-width"),
  length: document.getElementById("room-length"),
  area: document.getElementById("room-area"),
  load: document.getElementById("room-load"),
  fire: document.getElementById("room-fire"),
  board: document.getElementById("room-board"),
  a: document.getElementById("room-a"),
  b: document.getElementById("room-b"),
  c: document.getElementById("room-c"),
  validation: document.getElementById("validation"),

  saveRoom: document.getElementById("save-room"),
  cancelRoom: document.getElementById("cancel-room"),
  roomTabs: document.getElementById("room-tabs"),
  tbody: document.querySelector("#rooms-table tbody"),
  formulas: document.getElementById("formulas"),
  scheme: document.getElementById("scheme"),
  schemeLegend: document.getElementById("scheme-legend"),
};

let areaDirty = false;

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
      },
      activeRoomId: raw.activeRoomId || "",
    };
  } catch {
    return {
      rooms: [],
      constants: { cdLength: 4, udLength: 4, offset: 30, udAnchorSpacing: 625 },
      activeRoomId: "",
    };
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function createRoom() {
  const id = crypto.randomUUID();
  const room = {
    id,
    name: "Стая",
    width: 400,
    length: 300,
    area: 12,
    loadClass: "0.30",
    fireProtection: false,
    boardType: "12.5_or_2x12.5",
    a: 900,
    b: 500,
    c: 600,
    overrides: { area: false, a: false, b: false, c: false },
  };
  applyAutoABC(room);
  return room;
}

function applyAutoABC(room) {
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType);
  if (!room.overrides?.a) room.a = auto.a;
  if (!room.overrides?.c) room.c = auto.c;
  if (!room.overrides?.b) room.b = auto.b;
}

function getAutoABC(loadClass, fireProtection, boardType) {
  const idx = LOAD_CLASSES.indexOf(loadClass);
  const table = KNAUF_TABLE[String(fireProtection)];
  const sortedC = Object.keys(table).map(Number).sort((a, b) => a - b);
  const firstValid = sortedC.find((c) => table[c][idx] != null) || 600;
  return {
    a: table[firstValid][idx] || 900,
    c: firstValid,
    b: BOARD_TYPE_TO_B[boardType] || 500,
  };
}

function validateCombination(room) {
  const idx = LOAD_CLASSES.indexOf(room.loadClass);
  const table = KNAUF_TABLE[String(room.fireProtection)];
  const aExpected = table[room.c]?.[idx];
  const validB = Object.values(BOARD_TYPE_TO_B).includes(room.b);
  return Boolean(aExpected && aExpected === room.a && validB);
}

function calc(room) {
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = state.constants.offset;

  const bearingCount = Math.floor((W - offset) / (room.c / 10)) + 1;
  const bearingLengthTotal = bearingCount * (L / 100);
  const bearingProfiles = Math.ceil(bearingLengthTotal / state.constants.cdLength);

  const mountingCount = Math.floor((L - offset) / (room.b / 10)) + 1;
  const mountingLengthTotal = mountingCount * (W / 100);
  const mountingProfiles = Math.ceil(mountingLengthTotal / state.constants.cdLength);

  const cdTotalLength = bearingLengthTotal + mountingLengthTotal;
  const cdTotalProfiles = bearingProfiles + mountingProfiles;

  const crossConnectors = bearingCount * mountingCount;
  const hangersPerBearing = Math.floor((L - offset) / (room.a / 10)) + 1;
  const hangersTotal = bearingCount * hangersPerBearing;

  const udTotalLength = (2 * (X + Y)) / 100;
  const udProfiles = Math.ceil(udTotalLength / state.constants.udLength);

  const anchorsUd = Math.ceil(udTotalLength / (state.constants.udAnchorSpacing / 1000));
  const anchorsTotal = anchorsUd + hangersTotal;

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
    anchorsTotal,
    extensionsTotal,
  };
}

function buildPositions(limitCm, spacingMm) {
  const out = [];
  for (let pos = state.constants.offset; pos <= limitCm; pos += spacingMm / 10) out.push(pos);
  return out;
}

function render() {
  el.cdLength.value = state.constants.cdLength;
  el.udLength.value = state.constants.udLength;
  el.offset.value = state.constants.offset;
  el.udAnchor.value = state.constants.udAnchorSpacing;

  renderTabs();
  renderTable();
  const active = state.rooms.find((r) => r.id === state.activeRoomId);
  if (active) {
    bindRoomToForm(active);
    renderScheme(active);
    renderFormulas(active);
  }
  saveState();
}

function renderTabs() {
  el.roomTabs.innerHTML = "";
  state.rooms.forEach((room) => {
    const b = document.createElement("button");
    b.className = room.id === state.activeRoomId ? "tab active" : "tab";
    b.textContent = room.name;
    b.onclick = () => { state.activeRoomId = room.id; areaDirty = !!room.overrides?.area; render(); };
    el.roomTabs.appendChild(b);
  });
  const add = document.createElement("button");
  add.className = "tab add";
  add.textContent = "+ Нова стая";
  add.onclick = () => {
    const room = createRoom();
    room.name = `Стая ${state.rooms.length + 1}`;
    state.rooms.push(room);
    state.activeRoomId = room.id;
    areaDirty = false;
    render();
  };
  el.roomTabs.appendChild(add);
}

function bindRoomToForm(room) {
  el.formTitle.textContent = `Редакция: ${room.name}`;
  el.roomId.value = room.id;
  el.name.value = room.name;
  el.width.value = room.width;
  el.length.value = room.length;
  el.area.value = room.area;
  el.load.value = room.loadClass;
  el.fire.value = String(room.fireProtection);
  el.board.value = room.boardType;
  el.a.value = room.a;
  el.b.value = room.b;
  el.c.value = room.c;
  el.validation.textContent = validateCombination(room) ? "" : "Невалидна комбинация според Knauf D113";
}

function renderTable() {
  el.tbody.innerHTML = "";
  state.rooms.forEach((room) => {
    const r = calc(room);
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${room.name}</td><td>${room.width}</td><td>${room.length}</td><td>${Number(room.area).toFixed(2)}</td><td>${r.bearingCount}</td><td>${r.mountingCount}</td><td>${r.bearingLengthTotal.toFixed(2)}</td><td>${r.mountingLengthTotal.toFixed(2)}</td><td>${r.cdTotalProfiles}</td><td>${r.udProfiles}</td><td>${r.crossConnectors}</td><td>${r.hangersTotal}</td><td>${r.anchorsTotal}</td><td>${r.extensionsTotal}</td><td class="actions"><button data-id="${room.id}" data-action="edit">Редакция</button><button data-id="${room.id}" data-action="del" class="danger">Изтрий</button></td>`;
    el.tbody.appendChild(tr);
  });
}

function renderScheme(room) {
  const r = calc(room);
  const bearingPos = buildPositions(r.W, room.c);
  const mountingPos = buildPositions(r.L, room.b);
  const pad = 40;
  const w = 760;
  const h = 300;
  const xScale = w / r.L;
  const yScale = h / r.W;

  let svg = `<rect x="${pad}" y="${pad}" width="${w}" height="${h}" fill="#eef6ff" stroke="#13588f" stroke-width="2" />`;

  mountingPos.forEach((p) => {
    const x = pad + p * xScale;
    svg += `<line x1="${x}" y1="${pad}" x2="${x}" y2="${pad + h}" stroke="#2b9a42" stroke-width="1.5"/>`;
    svg += `<text x="${x + 2}" y="${pad + 12}" fill="#2b9a42" font-size="10">${p.toFixed(0)}см</text>`;
  });

  bearingPos.forEach((p) => {
    const y = pad + p * yScale;
    svg += `<line x1="${pad}" y1="${y}" x2="${pad + w}" y2="${y}" stroke="#1f5e93" stroke-width="2"/>`;
    svg += `<text x="${pad + 2}" y="${y - 3}" fill="#1f5e93" font-size="10">${p.toFixed(0)}см</text>`;
  });

  el.scheme.innerHTML = svg;
  const legendRows = [
    ["W (широчина)", `${r.W.toFixed(0)} cm`],
    ["L (дължина)", `${r.L.toFixed(0)} cm`],
    ["a (Разстояние между окачвачите)", `${room.a} mm`],
    ["b (Разстояние между монтажните профили)", `${room.b} mm`],
    ["c (Разстояние между носещите профили)", `${room.c} mm`],
    ["o (първоначален offset)", `${state.constants.offset} cm`],
    ["Клас на натоварване", `${room.loadClass} kN/m²`],
  ];
  el.schemeLegend.innerHTML = legendRows
    .map(([title, value]) => `<div class="legend-row"><span class="legend-title">${title}:</span> <span>${value}</span></div>`)
    .join("");
}

function f2(value) {
  return Number(value).toFixed(2);
}

function renderFormulas(room) {
  const r = calc(room);
  const X = Number(room.width);
  const Y = Number(room.length);
  const W = Math.min(X, Y);
  const L = Math.max(X, Y);
  const offset = state.constants.offset;
  const anchorStep = state.constants.udAnchorSpacing / 1000;

  const rows = [
    {
      title: "Изчисления за количествата",
      items: [
        `W = min(X, Y) = min(${X}, ${Y}) = ${W} cm`,
        `L = max(X, Y) = max(${X}, ${Y}) = ${L} cm`,
        `Носещи CD редове = floor((W - o) / (c / 10)) + 1 = floor((${W} - ${offset}) / (${room.c} / 10)) + 1 = ${r.bearingCount}`,
        `Монтажни CD редове = floor((L - o) / (b / 10)) + 1 = floor((${L} - ${offset}) / (${room.b} / 10)) + 1 = ${r.mountingCount}`,
        `Носещи CD метри = Носещи редове × (L / 100) = ${r.bearingCount} × (${L} / 100) = ${f2(r.bearingLengthTotal)} m`,
        `Монтажни CD метри = Монтажни редове × (W / 100) = ${r.mountingCount} × (${W} / 100) = ${f2(r.mountingLengthTotal)} m`,
        `CD бр. = ceil(Носещи m / CD дължина) + ceil(Монтажни m / CD дължина) = ceil(${f2(r.bearingLengthTotal)} / ${state.constants.cdLength}) + ceil(${f2(r.mountingLengthTotal)} / ${state.constants.cdLength}) = ${r.cdTotalProfiles}`,
        `UD дължина = 2 × (X + Y) / 100 = 2 × (${X} + ${Y}) / 100 = ${f2(r.udTotalLength)} m`,
        `UD бр. = ceil(UD дължина / UD дължина профил) = ceil(${f2(r.udTotalLength)} / ${state.constants.udLength}) = ${r.udProfiles}`,
        `Връзки = Носещи редове × Монтажни редове = ${r.bearingCount} × ${r.mountingCount} = ${r.crossConnectors}`,
        `Окачвачи/носещ = floor((L - o) / (a / 10)) + 1 = floor((${L} - ${offset}) / (${room.a} / 10)) + 1 = ${r.hangersPerBearing}`,
        `Окачвачи общо = Носещи редове × Окачвачи/носещ = ${r.bearingCount} × ${r.hangersPerBearing} = ${r.hangersTotal}`,
        `Дюбели UD = ceil(UD дължина / стъпка UD) = ceil(${f2(r.udTotalLength)} / ${f2(anchorStep)}) = ${r.anchorsUd}`,
        `Дюбели общо = Дюбели UD + Окачвачи = ${r.anchorsUd} + ${r.hangersTotal} = ${r.anchorsTotal}`,
        `Удължители = Носещи удълж. + Монтажни удълж. = ${r.extensionsTotal}`,
      ],
    },
    {
      title: "Формули за изчертаване на схемата",
      items: [
        `bearingPos = [o, o + c/10, o + 2c/10, ... ≤ W] = [${offset}, ${offset + room.c / 10}, ... ≤ ${W}]`,
        `mountingPos = [o, o + b/10, o + 2b/10, ... ≤ L] = [${offset}, ${offset + room.b / 10}, ... ≤ ${L}]`,
        `xScale = 760 / L = 760 / ${L} = ${f2(760 / L)}`,
        `yScale = 300 / W = 300 / ${W} = ${f2(300 / W)}`,
        `x(line) = pad + p × xScale, y(line) = pad + p × yScale (pad = 40)`,
      ],
    },
  ];

  el.formulas.innerHTML = `
    <h3>Формули за активната стая: ${room.name}</h3>
    ${rows.map((group) => `
      <div class="formulas-group">
        <h4>${group.title}</h4>
        ${group.items.map((item) => `<div class="formula-row">${item}</div>`).join("")}
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
  room.loadClass = el.load.value;
  room.fireProtection = el.fire.value === "true";
  room.boardType = el.board.value;
  room.a = Number(el.a.value);
  room.b = Number(el.b.value);
  room.c = Number(el.c.value);

  applyAutoABC(room);
  if (!room.overrides.area && room.width && room.length) room.area = (room.width * room.length) / 10000;
  el.validation.textContent = validateCombination(room) ? "" : "Невалидна комбинация според Knauf D113";
}

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

function handleLoadClassChange() {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (!room) return;
  room.loadClass = el.load.value;
  room.overrides.a = false;
  room.overrides.c = false;
  applyAutoABC(room);
  render();
}

el.load.addEventListener("change", handleLoadClassChange);
el.load.addEventListener("input", handleLoadClassChange);

el.fire.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.fireProtection = el.fire.value === "true";
  room.overrides.a = false;
  room.overrides.c = false;
  applyAutoABC(room);
  render();
});

el.board.addEventListener("change", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.boardType = el.board.value;
  room.overrides.b = el.board.value === "custom";
  applyAutoABC(room);
  render();
});

el.a.addEventListener("input", () => {
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
el.c.addEventListener("input", () => {
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  room.overrides.c = true;
  room.c = Number(el.c.value);
  render();
});

el.constantsForm.addEventListener("input", () => {
  state.constants.cdLength = Number(el.cdLength.value);
  state.constants.udLength = Number(el.udLength.value);
  state.constants.offset = Number(el.offset.value);
  state.constants.udAnchorSpacing = Number(el.udAnchor.value);
  render();
});

document.getElementById("save-room").addEventListener("click", () => {
  updateRoomFromForm();
  render();
});

document.getElementById("cancel-room").addEventListener("click", () => {
  areaDirty = false;
  const room = state.rooms.find((r) => r.id === state.activeRoomId);
  if (room) {
    room.overrides.area = false;
    room.area = (room.width * room.length) / 10000;
  }
  render();
});

el.tbody.addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button) return;
  const id = button.dataset.id;
  if (button.dataset.action === "edit") {
    state.activeRoomId = id;
    render();
    return;
  }
  state.rooms = state.rooms.filter((room) => room.id !== id);
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
      a: room.a,
      b: room.b,
      c: room.c,
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

document.getElementById("clear-all").addEventListener("click", () => {
  state.rooms = [createRoom()];
  state.activeRoomId = state.rooms[0].id;
  render();
});

render();
