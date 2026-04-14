import { ChangeEvent, useState } from "react";
import {
  CalcResult,
  CalculatorConstants,
  CONSTRUCTION_TYPES,
  DEFAULT_CONSTANTS,
  LoadClass,
  Room,
  SystemType,
  applyAutoABC,
  buildPositions,
  calc,
  createRoom,
  getConstruction,
  getAutoABC,
  getAllowedAValues,
  getAllowedBValues,
  getBoardOptions,
  getLoadClasses,
  getTableValue,
  getValidCValues,
  getValidationWarnings,
  syncSpacingFromKnaufTable,
  validateCombination,
} from "./domain/calculator";

const STORAGE_KEY = "d113-calculator-v2";

interface AppState {
  rooms: Room[];
  constants: CalculatorConstants;
  activeRoomId: string;
}

function loadState(): AppState {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") as Partial<AppState>;
    const rooms = Array.isArray(raw.rooms) && raw.rooms.length ? raw.rooms : [createRoom()];
    return {
      rooms,
      constants: { ...DEFAULT_CONSTANTS, ...raw.constants },
      activeRoomId: raw.activeRoomId || rooms[0].id,
    };
  } catch {
    const room = createRoom();
    return { rooms: [room], constants: DEFAULT_CONSTANTS, activeRoomId: room.id };
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

function buildCalculationFormulas(room: Room, result: CalcResult, constants: CalculatorConstants) {
  const X = Number(room.width);
  const Y = Number(room.length);
  const offset = Number(room.offset);
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
      formula: "ceil((W - offset) / (c / 10)) + 1",
      steps: [
        `c = ${room.c} mm = ${room.c / 10} cm.`,
        `ceil((${result.W} - ${offset}) / ${room.c / 10}) + 1 = ${result.bearingCount} бр.`,
      ],
    },
    {
      key: "mountingCount",
      label: "Монтажни редове",
      formula: "ceil((L - offset) / (b / 10)) + 1",
      steps: [
        `b = ${room.b} mm = ${room.b / 10} cm.`,
        `ceil((${result.L} - ${offset}) / ${room.b / 10}) + 1 = ${result.mountingCount} бр.`,
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
      formula: "ceil(носещи m / дължина профил) + ceil(монтажни m / дължина профил)",
      steps: [
        `Дължина профил = ${constants.cdLength} m.`,
        `ceil(${formatNumber(result.bearingLengthTotal)} / ${constants.cdLength}) + ceil(${formatNumber(result.mountingLengthTotal)} / ${constants.cdLength}) = ${result.cdTotalProfiles} бр.`,
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
      formula: "носещи редове x (ceil((L - offset) / (a / 10)) + 1)",
      steps: [
        `a = ${room.a} mm = ${room.a / 10} cm.`,
        `Окачвачи на носещ ред = ceil((${result.L} - ${offset}) / ${room.a / 10}) + 1 = ${result.hangersPerBearing} бр.`,
        `Общо = ${result.bearingCount} x ${result.hangersPerBearing} = ${result.hangersTotal} бр.`,
      ],
    },
    {
      key: "ud",
      label: "UD профили и дюбели",
      formula: "UD m = 2 x (X + Y) / 100; UD бр. = ceil(UD m / дължина UD); дюбели UD = ceil(UD m / стъпка UD)",
      steps: [
        `UD дължина = 2 x (${X} + ${Y}) / 100 = ${formatNumber(result.udTotalLength)} m.`,
        `UD профили = ceil(${formatNumber(result.udTotalLength)} / ${constants.udLength}) = ${result.udProfiles} бр.`,
        `Дюбели UD = ceil(${formatNumber(result.udTotalLength)} / (${room.udAnchorSpacing} / 1000)) = ${result.anchorsUd} бр.`,
      ],
    },
    {
      key: "screws",
      label: "Винтове",
      formula: "метал = връзки x винтове/връзка + окачвачи x винтове/окачвач; гипсокартон = площ x винтове/m2",
      steps: [
        `Метал = ${result.crossConnectors} x ${constants.metalScrewsPerCrossConnector} + ${result.hangersTotal} x ${constants.metalScrewsPerDirectHanger} = ${result.metalScrews} бр.`,
        `Гипсокартон = ceil(${formatNumber(room.area)} x ${constants.drywallScrewsPerM2}) = ${result.drywallScrews} бр.`,
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
  const auto = getAutoABC(room.loadClass, room.fireProtection, room.boardType, room.systemType);
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

  const activeRoom = state.rooms.find((room) => room.id === state.activeRoomId) ?? state.rooms[0];
  const activeResult = calc(activeRoom, state.constants);
  const activeWarnings = getValidationWarnings(cloneRoom(activeRoom));
  const isValid = !activeWarnings.some((warning) => warning.severity === "error");

  function commit(updater: (draft: AppState) => void): void {
    setState((current) => {
      const next = {
        rooms: current.rooms.map(cloneRoom),
        constants: { ...current.constants },
        activeRoomId: current.activeRoomId,
      };
      updater(next);
      saveState(next);
      return next;
    });
  }

  function saveCurrentState(): void {
    saveState(state);
    setSaveStatus("Запазено");
    window.setTimeout(() => setSaveStatus(""), 1600);
  }

  function updateActiveRoom(updater: (room: Room) => void): void {
    commit((draft) => {
      const room = draft.rooms.find((item) => item.id === draft.activeRoomId);
      if (!room) return;
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
      draft.rooms.push(room);
      draft.activeRoomId = room.id;
    });
  }

  function deleteRoom(roomId: string): void {
    commit((draft) => {
      draft.rooms = draft.rooms.filter((room) => room.id !== roomId);
      if (!draft.rooms.length) draft.rooms.push(createRoom());
      draft.activeRoomId = draft.rooms[0].id;
    });
  }

  function changeSystem(systemType: SystemType): void {
    updateActiveRoom((room) => {
      const construction = getConstruction(systemType);
      room.systemType = systemType;
      if (systemType !== "CUSTOM" && !room.boardType.startsWith("knauf_")) {
        room.boardType = "knauf_a_12.5";
      }
      room.fireProtection = construction.defaultFireProtection;
      room.loadClass = construction.defaultLoadClass;
      room.overrides.a = false;
      room.overrides.b = false;
      room.overrides.c = false;
      room.overrides.offset = false;
      room.overrides.udAnchorSpacing = false;
      syncSpacingFromKnaufTable(room, { keepC: false });
      if (systemType === "CUSTOM") {
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
      positions: {
        bearingProfilesCm: buildPositions(result.W, room.c, room.offset),
        mountingProfilesCm: buildPositions(result.L, room.b, room.offset),
        hangersCm: buildPositions(result.L, room.a, room.offset),
        extensionsCm: Array.from(
          { length: Math.max(0, Math.ceil(result.L / (state.constants.cdLength * 100)) - 1) },
          (_, index) => (index + 1) * state.constants.cdLength * 100,
        ),
      },
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
      const next = {
        rooms,
        constants: { ...DEFAULT_CONSTANTS, ...raw.constants },
        activeRoomId: rooms[0]?.id ?? createRoom().id,
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
    const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Стаи"><Table>${sheetRows}</Table></Worksheet></Workbook>`;
    const blob = new Blob([workbook], { type: "application/vnd.ms-excel" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "suspended-ceiling-materials.xls";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  const loadClasses = getLoadClasses(activeRoom.systemType, activeRoom.fireProtection);

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
            <ConstantsEditor
              constants={state.constants}
              onChange={(patch) => commit((draft) => { draft.constants = { ...draft.constants, ...patch }; })}
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
      </section>

      <section className="results-section">
        <RoomsTable
          rooms={state.rooms}
          constants={state.constants}
          activeRoomId={state.activeRoomId}
          onSelect={(roomId) => commit((draft) => {
            draft.activeRoomId = roomId;
          })}
          onDelete={deleteRoom}
        />
        <MaterialsPanel
          rooms={state.rooms}
          constants={state.constants}
          onReserveChange={(wastePercent) => commit((draft) => { draft.constants.wastePercent = wastePercent; })}
        />
      </section>
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
  const boardOptions = getBoardOptions(room.systemType);
  const aOptions = getAllowedAValues(cloneRoom(room));
  const bOptions = getAllowedBValues(room.systemType);
  const cOptions = getValidCValues(cloneRoom(room));
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
          <span className={isValid ? "status ok" : "status warn"}>{isValid ? "валидна" : "провери"}</span>
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
      <div className="system-card">
        <label className="system-select-label">Конструкция
          <select value={room.systemType} onChange={(event) => onSystemChange(event.target.value as SystemType)}>
            {Object.entries(CONSTRUCTION_TYPES).map(([value, item]) => (
              <option key={value} value={value}>{item.label}</option>
            ))}
          </select>
        </label>
      </div>

      <div className="field-grid">
        <label className="load-select-label">Натоварване
          <select value={room.loadClass} onChange={(event) => onRoomChange((draft) => {
            draft.loadClass = event.target.value as LoadClass;
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            {loadClasses.map((value) => <option key={value} value={value}>до {value} kN/m2</option>)}
          </select>
        </label>
        <label>Огнезащита
          <select value={String(room.fireProtection)} disabled={!construction.fireLoadClasses.length} onChange={(event) => onRoomChange((draft) => {
            draft.fireProtection = event.target.value === "true";
            draft.overrides.a = false;
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            <option value="false">Без</option>
            <option value="true">Да</option>
          </select>
        </label>
        <label className="span-2">Тип/дебелина гипсокартон
          <select value={room.boardType} onChange={(event) => onRoomChange((draft) => {
            draft.boardType = event.target.value as Room["boardType"];
            draft.overrides.b = draft.boardType === "custom";
            syncSpacingFromKnaufTable(draft, { keepC: true });
          })}>
            {boardOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
        </label>
      </div>

      <div className="spacing-head">
        <strong>Разстояния</strong>
        <button type="button" className="ghost small" onClick={onResetAuto}>Върни по Knauf</button>
      </div>
      <div className="field-grid spacing-card">
        {isCustom ? (
          <>
            <NumberField label="a Разстояние между окачвачи (mm)" value={room.a} manual={room.overrides.a} onChange={(value) => onRoomChange((draft) => {
              draft.a = value;
              syncDistanceOverride(draft, "a");
            })} />
            <NumberField label="b Разстояние между монтажни CD профили (mm)" value={room.b} manual={room.overrides.b} onChange={(value) => onRoomChange((draft) => {
              draft.b = value;
              syncDistanceOverride(draft, "b");
            })} />
            <NumberField label="c Разстояние между носещи CD/UA профили (mm)" value={room.c} manual={room.overrides.c} onChange={(value) => onRoomChange((draft) => {
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
        <NumberField label="Начално отстояние (cm)" value={room.offset} manual={room.overrides.offset} onChange={(value) => onRoomChange((draft) => {
          draft.offset = value;
          syncDistanceOverride(draft, "offset");
        })} />
        <NumberField label="UD дюбели (mm)" value={room.udAnchorSpacing} manual={room.overrides.udAnchorSpacing} onChange={(value) => onRoomChange((draft) => {
          draft.udAnchorSpacing = value;
          syncDistanceOverride(draft, "udAnchorSpacing");
        })} />
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
      <p className="hint">{construction.materialHint}</p>
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

function ConstantsEditor({ constants, onChange }: { constants: CalculatorConstants; onChange: (patch: Partial<CalculatorConstants>) => void }) {
  return (
    <section className="panel compact">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Глобални константи</p>
          <h2>Настройки</h2>
        </div>
      </div>
      <div className="field-grid">
        <NumberInput label="CD профил (m)" value={constants.cdLength} step={0.1} onChange={(value) => onChange({ cdLength: value })} />
        <NumberInput label="UD профил (m)" value={constants.udLength} step={0.1} onChange={(value) => onChange({ udLength: value })} />
        <NumberInput label="Винтове/връзка" value={constants.metalScrewsPerCrossConnector} onChange={(value) => onChange({ metalScrewsPerCrossConnector: value })} />
        <NumberInput label="Винтове/окачвач" value={constants.metalScrewsPerDirectHanger} onChange={(value) => onChange({ metalScrewsPerDirectHanger: value })} />
        <NumberInput label="Винтове/m2" value={constants.drywallScrewsPerM2} onChange={(value) => onChange({ drywallScrewsPerM2: value })} />
      </div>
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

function RoomsTable({ rooms, constants, activeRoomId, onSelect, onDelete }: {
  rooms: Room[];
  constants: CalculatorConstants;
  activeRoomId: string;
  onSelect: (roomId: string) => void;
  onDelete: (roomId: string) => void;
}) {
  return (
    <section className="panel table-panel">
      <div className="panel-title">
        <div>
          <p className="eyebrow">Стаи</p>
          <h2>Количества по стаи</h2>
        </div>
      </div>
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
    </section>
  );
}

function MaterialsPanel({ rooms, constants, onReserveChange }: {
  rooms: Room[];
  constants: CalculatorConstants;
  onReserveChange: (wastePercent: number) => void;
}) {
  const totals = rooms.reduce((acc, room) => {
    const result = calc(cloneRoom(room), constants);
    acc.cdTotalProfiles += result.cdTotalProfiles;
    acc.cdTotalLength += result.cdTotalLength;
    acc.udProfiles += result.udProfiles;
    acc.udTotalLength += result.udTotalLength;
    acc.crossConnectors += result.crossConnectors;
    acc.hangersTotal += result.hangersTotal;
    acc.hangersPerBearing = Math.max(acc.hangersPerBearing, result.hangersPerBearing);
    acc.anchorsUd += result.anchorsUd;
    acc.anchorsHangers += result.anchorsHangers;
    acc.anchorsTotal += result.anchorsTotal;
    acc.metalScrews += result.metalScrews;
    acc.drywallScrews += result.drywallScrews;
    acc.extensionsTotal += result.extensionsTotal;
    return acc;
  }, {
    cdTotalProfiles: 0,
    cdTotalLength: 0,
    udProfiles: 0,
    udTotalLength: 0,
    crossConnectors: 0,
    hangersTotal: 0,
    hangersPerBearing: 0,
    anchorsUd: 0,
    anchorsHangers: 0,
    anchorsTotal: 0,
    metalScrews: 0,
    drywallScrews: 0,
    extensionsTotal: 0,
  });
  const reserveMultiplier = 1 + (Number(constants.wastePercent) || 0) / 100;
  const reserveCount = (value: number) => Math.ceil(value * reserveMultiplier);
  const reserveLength = (value: number) => formatNumber(value * reserveMultiplier);
  const rows = [
    ["Профили общо", `${reserveCount(totals.cdTotalProfiles)} бр.`, `${reserveLength(totals.cdTotalLength)} m след резерв`],
    ["UD 28/27", `${reserveCount(totals.udProfiles)} бр.`, `${reserveLength(totals.udTotalLength)} m след резерв`],
    ["Връзки", `${reserveCount(totals.crossConnectors)} бр.`, "типът зависи от системата"],
    ["Окачвачи", `${reserveCount(totals.hangersTotal)} бр.`, "общо за всички стаи"],
    ["Дюбели общо", `${reserveCount(totals.anchorsTotal)} бр.`, `UD ${reserveCount(totals.anchorsUd)}, окачвачи ${reserveCount(totals.anchorsHangers)}`],
    ["Винтове", `${reserveCount(totals.metalScrews + totals.drywallScrews)} бр.`, `метал ${reserveCount(totals.metalScrews)}, гипсокартон ${reserveCount(totals.drywallScrews)}`],
    ["Удължители", `${reserveCount(totals.extensionsTotal)} бр.`, "само геометрична оценка"],
  ];
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
      <div className="material-list">
        {rows.map(([label, value, note]) => (
          <div key={label} className="material-row">
            <span>{label}</span>
            <strong>{value}</strong>
            <small>{note}</small>
          </div>
        ))}
      </div>
    </section>
  );
}

export default App;
