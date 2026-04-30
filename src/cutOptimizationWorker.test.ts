import { beforeEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_CONSTANTS, createRoom } from "./domain/calculator";

type WorkerSelf = {
  onmessage: ((event: MessageEvent) => void) | null;
  postMessage: ReturnType<typeof vi.fn>;
};

async function loadWorker() {
  const workerSelf: WorkerSelf = {
    onmessage: null,
    postMessage: vi.fn(),
  };

  vi.stubGlobal("self", workerSelf);
  vi.resetModules();
  await import("./cutOptimizationWorker");

  if (!workerSelf.onmessage) {
    throw new Error("Worker did not register an onmessage handler.");
  }

  return workerSelf;
}

describe("cutOptimizationWorker", () => {
  beforeEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns optimized room table counts for each room and globally", async () => {
    const worker = await loadWorker();
    const room = createRoom("Worker стая");

    worker.onmessage?.({
      data: {
        type: "optimize-room-table-cuts",
        requestId: "rooms-1",
        rooms: [room],
        constants: DEFAULT_CONSTANTS,
      },
    } as MessageEvent);

    expect(worker.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "room-cut-optimization-result",
        requestId: "rooms-1",
        rows: expect.objectContaining({
          [room.id]: expect.objectContaining({
            cd: expect.any(Number),
            ud: expect.any(Number),
          }),
        }),
        global: expect.objectContaining({
          cd: expect.any(Number),
          ud: expect.any(Number),
        }),
      }),
    );
  });

  it("returns cut plans for room and global requests", async () => {
    const worker = await loadWorker();
    const room = createRoom("Разкрой");

    worker.onmessage?.({
      data: {
        type: "optimize-cut-plan",
        requestId: "plan-room",
        mode: "room",
        room,
        constants: DEFAULT_CONSTANTS,
      },
    } as MessageEvent);
    worker.onmessage?.({
      data: {
        type: "optimize-cut-plan",
        requestId: "plan-global",
        mode: "global",
        rooms: [room],
        constants: DEFAULT_CONSTANTS,
      },
    } as MessageEvent);

    expect(worker.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "cut-plan-result",
        requestId: "plan-room",
        plan: expect.objectContaining({ bars: expect.any(Array) }),
      }),
    );
    expect(worker.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "cut-plan-result",
        requestId: "plan-global",
        plan: expect.objectContaining({ bars: expect.any(Array) }),
      }),
    );
  });

  it("builds material takeoff rows without nested optimized quantities", async () => {
    const worker = await loadWorker();
    const room = createRoom("Материали");

    worker.onmessage?.({
      data: {
        type: "build-material-takeoff",
        requestId: "materials-1",
        rooms: [room],
        constants: DEFAULT_CONSTANTS,
      },
    } as MessageEvent);

    expect(worker.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "material-takeoff-result",
        requestId: "materials-1",
        rows: expect.arrayContaining([
          expect.objectContaining({
            label: expect.any(String),
            quantity: expect.any(Number),
          }),
        ]),
      }),
    );
  });

  it("returns typed errors when a worker request fails", async () => {
    const worker = await loadWorker();
    const room = createRoom("Грешка");

    worker.onmessage?.({
      data: {
        type: "optimize-cut-plan",
        requestId: "broken-plan",
        mode: "room",
        room,
        constants: undefined,
      },
    } as unknown as MessageEvent);

    expect(worker.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "cut-plan-error",
        requestId: "broken-plan",
        error: expect.any(String),
      }),
    );
  });
});
