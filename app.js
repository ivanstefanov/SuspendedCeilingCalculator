const STORAGE_KEY = "suspended-ceiling-rooms-v1";
const LOAD_CLASSES = ["0.15", "0.30", "0.40", "0.50", "0.65"];
const BOARD_TYPE_TO_B_SPACING = {
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

const form = document.getElementById("room-form");
const roomIdInput = document.getElementById("room-id");
const xCmInput = document.getElementById("xCm");
const yCmInput = document.getElementById("yCm");
const areaInput = document.getElementById("areaM2");
const nameInput = document.getElementById("name");
const loadInput = document.getElementById("loadClass");
const fireInput = document.getElementById("fireProtection");
const cSpacingInput = document.getElementById("cSpacing");
const hangerSpacingInput = document.getElementById("hangerSpacing");
const cSpacingOptions = document.getElementById("cSpacingOptions");
const hangerSpacingOptions = document.getElementById("hangerSpacingOptions");
const mountSpacingInput = document.getElementById("mountSpacing");
const boardTypeInput = document.getElementById("boardType");
const udDowelInput = document.getElementById("udDowelSpacing");
const cdProfileLengthInput = document.getElementById("cdProfileLengthM");
const udProfileLengthInput = document.getElementById("udProfileLengthM");
const tbody = document.querySelector("#rooms-table tbody");
const tableNote = document.getElementById("table-note");
const constantsTbody = document.querySelector("#constants-table tbody");
const scheme = document.getElementById("scheme");

let areaDirty = false;
let rooms = loadRooms();

function loadRooms() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveRooms() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rooms));
}

function recalcArea() {
  if (areaDirty) return;
  const x = Number(xCmInput.value);
  const y = Number(yCmInput.value);
  if (!x || !y) return;
  areaInput.value = ((x * y) / 10000).toFixed(2);
}

function pickKnaufParams(loadClass, fireProtection) {
  const classIndex = LOAD_CLASSES.indexOf(loadClass);
  const table = KNAUF_TABLE[String(fireProtection)];
  const sortedC = Object.keys(table).map(Number).sort((a, b) => b - a);
  for (const c of sortedC) {
    const a = table[c][classIndex];
    if (a) return { cSpacingMm: c, hangerSpacingMm: a };
  }
  return { cSpacingMm: 500, hangerSpacingMm: 600 };
}

function buildPositionsMm(spanMm, spacingMm, firstOffsetMm = 300) {
  if (spanMm <= firstOffsetMm * 2) return [spanMm / 2];
  const positions = [];
  const endLimit = spanMm - firstOffsetMm;
  for (let pos = firstOffsetMm; pos <= endLimit; pos += spacingMm) {
    positions.push(pos);
  }
  return positions.length ? positions : [spanMm / 2];
}

function getKnaufOptions(loadClass, fireProtection) {
  const classIndex = LOAD_CLASSES.indexOf(loadClass);
  const table = KNAUF_TABLE[String(fireProtection)];
  const options = Object.entries(table)
    .map(([c, values]) => ({ c: Number(c), a: values[classIndex] }))
    .filter((row) => row.a)
    .sort((left, right) => left.c - right.c);
  return options;
}

function setDatalistOptions(el, values) {
  el.innerHTML = values.map((value) => `<option value="${value}"></option>`).join("");
}

function refreshSpacingPresets(force = false) {
  const options = getKnaufOptions(loadInput.value, fireInput.value === "true");
  if (!options.length) return;

  setDatalistOptions(cSpacingOptions, options.map((row) => row.c));
  setDatalistOptions(hangerSpacingOptions, options.map((row) => row.a));

  if (force || !cSpacingInput.value) cSpacingInput.value = options[0].c;
  if (force || !hangerSpacingInput.value) hangerSpacingInput.value = options[0].a;
}

function ceilDiv(a, b) {
  return Math.ceil(a / b);
}

function calcRoomMetrics(room) {
  const xCm = Number(room.xCm);
  const yCm = Number(room.yCm);
  const areaM2 = Number(room.areaM2);
  const W = Math.min(xCm, yCm) / 100;
  const L = Math.max(xCm, yCm) / 100;

  const fallback = pickKnaufParams(room.loadClass, room.fireProtection);
  const cSpacingMm = Number(room.cSpacingMm || fallback.cSpacingMm);
  const hangerSpacingMm = Number(room.hangerSpacingMm || fallback.hangerSpacingMm);
  const bSpacingMm = Number(room.mountSpacingMm || 500);
  const udDowelSpacing = Number(room.udDowelSpacingMm || 500);
  const cdProfileLengthM = Number(room.cdProfileLengthM || 4);
  const udProfileLengthM = Number(room.udProfileLengthM || 4);

  const udNeededM = 2 * (W + L);
  const udProfiles = ceilDiv(udNeededM, udProfileLengthM);

  const firstCarrierOffsetMm = Math.max(0, cSpacingMm - 30);
  const primaryPositionsMm = buildPositionsMm(W * 1000, cSpacingMm, firstCarrierOffsetMm);
  const primaryRows = primaryPositionsMm.length;
  const primaryTotalM = primaryRows * L;
  const primaryProfiles = ceilDiv(primaryTotalM, cdProfileLengthM);

  const secondaryPositionsMm = buildPositionsMm(L * 1000, bSpacingMm, 300);
  const secondaryRows = secondaryPositionsMm.length;
  const secondaryTotalM = secondaryRows * W;
  const secondaryProfiles = ceilDiv(secondaryTotalM, cdProfileLengthM);

  const totalCdM = primaryTotalM + secondaryTotalM;
  const totalCdProfiles = primaryProfiles + secondaryProfiles;

  const crossConnectors = primaryRows * secondaryRows;
  const hangerPositionsMm = buildPositionsMm(L * 1000, hangerSpacingMm, 300);
  const directPerPrimary = hangerPositionsMm.length;
  const directTotal = directPerPrimary * primaryRows;

  const udDowels = Math.ceil((udNeededM * 1000) / udDowelSpacing);
  const ceilingDowels = directTotal;
  const totalMetalDowels = udDowels + ceilingDowels;

  const extPrimary = Math.max(0, ceilDiv(L, cdProfileLengthM) - 1) * primaryRows;
  const extSecondary = Math.max(0, ceilDiv(W, cdProfileLengthM) - 1) * secondaryRows;
  const totalExt = extPrimary + extSecondary;

  return {
    xCm, yCm, areaM2,
    wShortM: W,
    lLongM: L,
    udNeededM,
    udProfiles,
    primaryRows,
    primaryTotalM,
    primaryProfiles,
    secondaryRows,
    secondaryTotalM,
    secondaryProfiles,
    totalCdM,
    totalCdProfiles,
    crossConnectors,
    directPerPrimary,
    directTotal,
    udDowels,
    ceilingDowels,
    totalMetalDowels,
    extPrimary,
    extSecondary,
    totalExt,
    cSpacingMm,
    hangerSpacingMm,
    bSpacingMm,
    cdProfileLengthM,
    udProfileLengthM,
    firstCarrierOffsetMm,
    primaryPositionsMm,
    secondaryPositionsMm,
    hangerPositionsMm,
  };
}

function format(n, digits = 2) {
  return Number(n).toFixed(digits);
}

function renderTable() {
  tbody.innerHTML = "";
  rooms.forEach((room) => {
    const m = calcRoomMetrics(room);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${room.name}</td><td>${m.xCm}</td><td>${m.yCm}</td><td>${format(m.areaM2)}</td><td>${format(m.wShortM)}</td><td>${format(m.lLongM)}</td>
      <td>${format(m.udNeededM)}</td><td>${m.udProfiles}</td><td>${m.primaryRows}</td><td>${format(m.primaryTotalM)}</td>
      <td>${m.primaryProfiles}</td><td>${m.secondaryRows}</td><td>${format(m.secondaryTotalM)}</td><td>${m.secondaryProfiles}</td>
      <td>${format(m.totalCdM)}</td><td>${m.totalCdProfiles}</td><td>${m.crossConnectors}</td><td>${m.directPerPrimary}</td>
      <td>${m.directTotal}</td><td>${m.udDowels}</td><td>${m.ceilingDowels}</td><td>${m.totalMetalDowels}</td>
      <td>${m.extPrimary}</td><td>${m.extSecondary}</td><td>${m.totalExt}</td>
      <td>до ${room.loadClass}</td><td>${room.fireProtection ? "Да" : "Не"}</td>
      <td class="row-actions">
        <button data-action="edit" data-id="${room.id}">Редакция</button>
        <button data-action="delete" data-id="${room.id}" class="danger">Изтрий</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tableNote.textContent = "Бележка: c и a се префилват по таблица Knauf според натоварване/пожарозащита, b се предлага според типа плоскост. За D113: първи и последен носещ CD са на (c - 30) мм от стените. Броят профили се закръгля винаги нагоре.";

  const selected = rooms[0];
  if (selected) renderScheme(calcRoomMetrics(selected), selected.name);
  else scheme.innerHTML = "<text x='20' y='40' fill='#3f5972'>Добавете стая, за да видите схема.</text>";
}

function renderScheme(m, roomName) {
  const w = 800;
  const h = 420;
  const pad = 40;
  const rw = w - pad * 2;
  const rh = h - pad * 2;
  const primaryCount = m.primaryRows;
  const secondaryCount = m.secondaryRows;
  let svg = `
    <rect x="${pad}" y="${pad}" width="${rw}" height="${rh}" fill="#d9ecff" stroke="#0f4f88" stroke-width="3" />
    <text x="${pad}" y="26" fill="#0f4f88" font-size="14">Стая: ${roomName}</text>
    <text x="${pad + 220}" y="26" fill="#0f4f88" font-size="14">W=${format(m.wShortM)} m, L=${format(m.lLongM)} m</text>
    <text x="${pad + 500}" y="26" fill="#0f4f88" font-size="14">c=${m.cSpacingMm} мм, b=${m.bSpacingMm} мм, a=${m.hangerSpacingMm} мм</text>
    <text x="${pad}" y="${h - 10}" fill="#365b7f" font-size="13">Отстояние на първи/последен носещ CD: c - 30 = ${m.firstCarrierOffsetMm} мм</text>
  `;

  const primaryScale = rh / (m.wShortM * 1000);
  const secondaryScale = rw / (m.lLongM * 1000);

  for (let i = 0; i < primaryCount; i++) {
    const positionMm = m.primaryPositionsMm[i];
    const y = pad + positionMm * primaryScale;
    svg += `<line x1="${pad}" y1="${y}" x2="${pad + rw}" y2="${y}" stroke="#284e7a" stroke-width="2.5" />`;
    svg += `<text x="${pad + 8}" y="${y - 6}" fill="#1f3b5c" font-size="11">${Math.round(positionMm / 10)} см</text>`;
  }

  for (let i = 0; i < secondaryCount; i++) {
    const positionMm = m.secondaryPositionsMm[i];
    const x = pad + positionMm * secondaryScale;
    svg += `<line x1="${x}" y1="${pad}" x2="${x}" y2="${pad + rh}" stroke="#7a99bd" stroke-width="1.7" />`;
    svg += `<text x="${x + 4}" y="${pad + 14}" fill="#476483" font-size="10">${Math.round(positionMm / 10)} см</text>`;
  }

  svg += `
    <line x1="${pad + 6}" y1="${pad + 30}" x2="${pad + 6}" y2="${h - pad - 30}" stroke="#de8f00" stroke-dasharray="5,4" stroke-width="1.5"/>
    <line x1="${pad + rw - 6}" y1="${pad + 30}" x2="${pad + rw - 6}" y2="${h - pad - 30}" stroke="#de8f00" stroke-dasharray="5,4" stroke-width="1.5"/>
  `;
  scheme.innerHTML = svg;
}

function resetForm() {
  form.reset();
  roomIdInput.value = "";
  areaDirty = false;
  document.getElementById("form-title").textContent = "Нова стая";
  mountSpacingInput.value = 500;
  boardTypeInput.value = "12.5_or_2x12.5";
  udDowelInput.value = 500;
  cdProfileLengthInput.value = 4;
  udProfileLengthInput.value = 4;
  refreshSpacingPresets(true);
}

form.addEventListener("submit", (e) => {
  e.preventDefault();
  const room = {
    id: roomIdInput.value || crypto.randomUUID(),
    name: nameInput.value.trim(),
    xCm: Number(xCmInput.value),
    yCm: Number(yCmInput.value),
    areaM2: Number(areaInput.value),
    loadClass: loadInput.value,
    fireProtection: fireInput.value === "true",
    cSpacingMm: Number(cSpacingInput.value),
    hangerSpacingMm: Number(hangerSpacingInput.value),
    boardType: boardTypeInput.value,
    mountSpacingMm: Number(mountSpacingInput.value),
    udDowelSpacingMm: Number(udDowelInput.value),
    cdProfileLengthM: Number(cdProfileLengthInput.value),
    udProfileLengthM: Number(udProfileLengthInput.value),
  };

  const idx = rooms.findIndex((r) => r.id === room.id);
  if (idx >= 0) rooms[idx] = room;
  else rooms.push(room);

  saveRooms();
  renderTable();
  resetForm();
});

document.getElementById("cancel-edit").addEventListener("click", resetForm);

tbody.addEventListener("click", (e) => {
  const button = e.target.closest("button");
  if (!button) return;
  const id = button.dataset.id;
  const action = button.dataset.action;
  const room = rooms.find((r) => r.id === id);
  if (!room) return;

  if (action === "delete") {
    rooms = rooms.filter((r) => r.id !== id);
    saveRooms();
    renderTable();
    return;
  }

  roomIdInput.value = room.id;
  nameInput.value = room.name;
  xCmInput.value = room.xCm;
  yCmInput.value = room.yCm;
  areaInput.value = room.areaM2;
  loadInput.value = room.loadClass;
  fireInput.value = String(room.fireProtection);
  refreshSpacingPresets(true);
  cSpacingInput.value = room.cSpacingMm || pickKnaufParams(room.loadClass, room.fireProtection).cSpacingMm;
  hangerSpacingInput.value = room.hangerSpacingMm || pickKnaufParams(room.loadClass, room.fireProtection).hangerSpacingMm;
  boardTypeInput.value = room.boardType || "12.5_or_2x12.5";
  mountSpacingInput.value = room.mountSpacingMm || 500;
  udDowelInput.value = room.udDowelSpacingMm || 500;
  cdProfileLengthInput.value = room.cdProfileLengthM || 4;
  udProfileLengthInput.value = room.udProfileLengthM || 4;
  areaDirty = true;
  document.getElementById("form-title").textContent = `Редакция: ${room.name}`;
});

xCmInput.addEventListener("input", recalcArea);
yCmInput.addEventListener("input", recalcArea);
areaInput.addEventListener("input", () => { areaDirty = true; });
loadInput.addEventListener("change", () => refreshSpacingPresets(true));
fireInput.addEventListener("change", () => refreshSpacingPresets(true));
boardTypeInput.addEventListener("change", () => {
  if (boardTypeInput.value === "custom") return;
  mountSpacingInput.value = BOARD_TYPE_TO_B_SPACING[boardTypeInput.value] || 500;
});

// Export / import

document.getElementById("export-json").addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(rooms, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "suspended-ceiling-rooms.json";
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById("import-json").addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  const text = await file.text();
  try {
    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed)) throw new Error("Очаква се масив от стаи");
    rooms = parsed;
    saveRooms();
    renderTable();
    resetForm();
    alert("Импортът е успешен.");
  } catch (err) {
    alert(`Грешка при импорт: ${err.message}`);
  } finally {
    e.target.value = "";
  }
});

document.getElementById("clear-all").addEventListener("click", () => {
  if (!confirm("Сигурни ли сте, че искате да изтриете всички стаи?")) return;
  rooms = [];
  saveRooms();
  renderTable();
  resetForm();
});

function renderConstantsTable() {
  const constants = [
    { key: "b (монтажни CD)", value: "400 / 500 / 550 / 625 / 800 мм", description: "Според тип и дебелина на плоскостта." },
    { key: "Първи/последен носещ CD", value: "c - 30 мм", description: "Оста е симетрично разположена от двете крайни стени." },
    { key: "UD анкериране", value: "≤ 625 мм", description: "Закрепване на UD профила към периметъра." },
    { key: "Влизане в UD", value: "≥ 20 мм", description: "Носещи/монтажни профили влизат минимум 20 мм в UD." },
    { key: "Винтове към UD", value: "≤ 170 мм", description: "При носеща връзка по периметъра (вариант 2)." },
    { key: "Макс. конзолно издаване", value: "≈ 100 мм", description: "Максимално издаване на облицовката към периметъра." },
    { key: "Монтажни CD при мазилка ≥6 мм", value: "≤ 312.5 мм", description: "По-гъста подконструкция при допълнителен товар от мазилка." },
    { key: "Мин. разст. окачвания по CD", value: "≥ 500 мм", description: "Разстоянието между две окачвания по един CD." },
    { key: "Разпределен товар (без огнезащита)", value: "≤ 20 kg/m²", description: "По-големи товари се окачват към основния таван/помощна конструкция." },
    { key: "Единичен товар към стоманена конструкция", value: "≤ 10 kg", description: "Максимум на точков товар към подконструкцията." },
  ];
  constantsTbody.innerHTML = constants
    .map((item) => `<tr><td>${item.key}</td><td>${item.value}</td><td>${item.description}</td></tr>`)
    .join("");
}

renderTable();
refreshSpacingPresets(true);
renderConstantsTable();
