import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import { DEFAULT_CONSTANTS, MaterialTakeoffItem, createRoom } from "./domain/calculator";

const STORAGE_KEY = "d113-calculator-v2";

type WorkerMessage = { type: string; requestId: string };

class MockWorker {
  static instances: MockWorker[] = [];
  onmessage: ((event: MessageEvent) => void) | null = null;
  onerror: ((event: Event) => void) | null = null;
  messages: WorkerMessage[] = [];
  terminated = false;

  constructor() {
    MockWorker.instances.push(this);
  }

  postMessage(message: WorkerMessage): void {
    this.messages.push(message);
  }

  terminate(): void {
    this.terminated = true;
  }

  emit(data: unknown): void {
    this.onmessage?.({ data } as MessageEvent);
  }
}

function latestWorkerFor(type: string): MockWorker {
  const worker = [...MockWorker.instances].reverse().find((item) => item.messages.some((message) => message.type === type));
  if (!worker) throw new Error(`Missing worker for ${type}`);
  return worker;
}

async function findWorkerFor(type: string): Promise<MockWorker> {
  await waitFor(() => expect(MockWorker.instances.some((item) => item.messages.some((message) => message.type === type))).toBe(true));
  return latestWorkerFor(type);
}

function latestMessage(worker: MockWorker, type: string): WorkerMessage {
  const message = [...worker.messages].reverse().find((item) => item.type === type);
  if (!message) throw new Error(`Missing worker message ${type}`);
  return message;
}

function seedSavedRoom(): ReturnType<typeof createRoom> {
  const room = createRoom("Тест стая");
  const state = {
    rooms: [room],
    draftRoom: room,
    constants: DEFAULT_CONSTANTS,
    materialPrices: {},
    activeRoomId: room.id,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  return room;
}

beforeEach(() => {
  localStorage.clear();
  MockWorker.instances = [];
  vi.stubGlobal("indexedDB", undefined);
  vi.stubGlobal("Worker", MockWorker);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("App recalculation flow", () => {
  it("keeps typed room changes out of the visualization until preview or save", async () => {
    const user = userEvent.setup();
    render(<App />);

    expect(screen.getByText("L 400 cm")).toBeInTheDocument();

    const widthInput = screen.getByLabelText(/Ширина X/);
    await user.clear(widthInput);
    await user.type(widthInput, "500");

    expect(screen.getByText("L 400 cm")).toBeInTheDocument();
    expect(screen.queryByText("L 500 cm")).not.toBeInTheDocument();

    await user.click(screen.getByRole("button", { name: "Прегледай" }));

    expect(screen.getByText("L 500 cm")).toBeInTheDocument();
    expect(screen.getByText(/Няма запазени стаи/)).toBeInTheDocument();

    await user.click(screen.getAllByRole("button", { name: "Запази стая" })[0]);

    expect(screen.getAllByRole("button", { name: "Стая" }).length).toBeGreaterThan(0);
  });

  it("updates optimized table values from the worker and keeps old values visible while recalculating", async () => {
    const user = userEvent.setup();
    const room = seedSavedRoom();
    render(<App />);

    const firstWorker = await findWorkerFor("optimize-room-table-cuts");
    const firstMessage = latestMessage(firstWorker, "optimize-room-table-cuts");
    firstWorker.emit({
      type: "room-cut-optimization-result",
      requestId: firstMessage.requestId,
      rows: { [room.id]: { cd: 7, ud: 3 } },
      global: { cd: 7, ud: 3 },
    });

    expect(await screen.findByText(/\(7 след разкрой\)/)).toBeInTheDocument();
    expect(screen.getByText(/\(3 след разкрой\)/)).toBeInTheDocument();

    const widthInput = screen.getByLabelText(/Ширина X/);
    await user.clear(widthInput);
    await user.type(widthInput, "420");
    const workerCountBeforeSave = MockWorker.instances.length;
    await user.click(screen.getAllByRole("button", { name: "Запази стая" })[0]);

    expect(await screen.findByText("7...")).toBeInTheDocument();
    expect(screen.getByText("3...")).toBeInTheDocument();
    expect(screen.getAllByText(/Прекалкулиране/).length).toBeGreaterThan(0);

    await waitFor(() => expect(MockWorker.instances.length).toBeGreaterThan(workerCountBeforeSave));
    const secondWorker = latestWorkerFor("optimize-room-table-cuts");
    const secondMessage = latestMessage(secondWorker, "optimize-room-table-cuts");
    secondWorker.emit({
      type: "room-cut-optimization-result",
      requestId: secondMessage.requestId,
      rows: { [room.id]: { cd: 8, ud: 4 } },
      global: { cd: 8, ud: 4 },
    });

    expect(await screen.findByText(/\(8 след разкрой\)/)).toBeInTheDocument();
    expect(screen.getByText(/\(4 след разкрой\)/)).toBeInTheDocument();
  });

  it("keeps previous material rows visible while the material worker recalculates", async () => {
    const user = userEvent.setup();
    seedSavedRoom();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Общи материали" }));

    const firstWorker = await findWorkerFor("build-material-takeoff");
    const firstMessage = latestMessage(firstWorker, "build-material-takeoff");
    const row: MaterialTakeoffItem = {
      key: "d112-boards",
      label: "Тест плоскости",
      quantity: 5,
      unit: "бр.",
      note: "тест",
      source: "estimate",
    };
    firstWorker.emit({
      type: "material-takeoff-result",
      requestId: firstMessage.requestId,
      rows: [row],
    });

    expect(await screen.findByText("Тест плоскости")).toBeInTheDocument();

    const reserveInput = screen.getByLabelText(/Резерв/);
    await user.clear(reserveInput);
    await user.type(reserveInput, "12");
    const workerCountBeforeSave = MockWorker.instances.length;
    await user.click(screen.getByRole("button", { name: "Запази" }));

    expect(screen.getByText("Тест плоскости")).toBeInTheDocument();
    expect(screen.getAllByText("Прекалкулиране...").length).toBeGreaterThan(0);

    await waitFor(() => expect(MockWorker.instances.length).toBeGreaterThan(workerCountBeforeSave));
    const secondWorker = latestWorkerFor("build-material-takeoff");
    const secondMessage = latestMessage(secondWorker, "build-material-takeoff");
    secondWorker.emit({
      type: "material-takeoff-result",
      requestId: secondMessage.requestId,
      rows: [{ ...row, label: "Нови плоскости", quantity: 6 }],
    });

    expect(await screen.findByText("Нови плоскости")).toBeInTheDocument();
    expect(screen.queryByText("Тест плоскости")).not.toBeInTheDocument();
  });

  it("does not apply global settings until settings Save is clicked", async () => {
    const user = userEvent.setup();
    seedSavedRoom();
    render(<App />);

    await user.click(screen.getByRole("button", { name: "Настройки" }));
    const cdLengthInput = screen.getByLabelText(/CD профил/);
    await user.clear(cdLengthInput);
    await user.type(cdLengthInput, "3.5");

    const storedBeforeSave = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") as { constants?: { cdLength?: number } };
    expect(storedBeforeSave.constants?.cdLength).toBe(DEFAULT_CONSTANTS.cdLength);

    const settingsPanel = screen.getByRole("heading", { name: "Настройки" }).closest("section");
    expect(settingsPanel).not.toBeNull();
    await user.click(within(settingsPanel as HTMLElement).getByRole("button", { name: "Запази" }));

    await waitFor(() => {
      const storedAfterSave = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") as { constants?: { cdLength?: number } };
      expect(storedAfterSave.constants?.cdLength).toBe(3.5);
    });
  });
});
