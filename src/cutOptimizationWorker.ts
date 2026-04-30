import {
  CalculatorConstants,
  CutBar,
  MaterialTakeoffItem,
  Room,
  buildCutOptimizationInput,
  buildMaterialTakeoff,
  calc,
  optimizeAllRoomsSuspendedCeilingCuts,
  optimizeSuspendedCeilingCuts,
} from "./domain/calculator";

type OptimizedProfileCounts = { cd: number | null; ud: number | null };

type CutOptimizationWorkerRequest =
  | {
      type: "optimize-room-table-cuts";
      requestId: string;
      rooms: Room[];
      constants: CalculatorConstants;
    }
  | {
      type: "optimize-cut-plan";
      requestId: string;
      mode: "room";
      room: Room;
      constants: CalculatorConstants;
    }
  | {
      type: "optimize-cut-plan";
      requestId: string;
      mode: "global";
      rooms: Room[];
      constants: CalculatorConstants;
    }
  | {
      type: "build-material-takeoff";
      requestId: string;
      rooms: Room[];
      constants: CalculatorConstants;
    };

type CutOptimizationWorkerResponse =
  | {
      type: "room-cut-optimization-result";
      requestId: string;
      rows: Record<string, OptimizedProfileCounts>;
      global: OptimizedProfileCounts;
    }
  | {
      type: "room-cut-optimization-error";
      requestId: string;
      error: string;
    }
  | {
      type: "cut-plan-result";
      requestId: string;
      plan: ReturnType<typeof optimizeSuspendedCeilingCuts>;
    }
  | {
      type: "cut-plan-error";
      requestId: string;
      error: string;
    }
  | {
      type: "material-takeoff-result";
      requestId: string;
      rows: MaterialTakeoffItem[];
    }
  | {
      type: "material-takeoff-error";
      requestId: string;
      error: string;
    };

function cloneRoom(room: Room): Room {
  return { ...room, overrides: { ...room.overrides } };
}

function isCdCutBar(bar: CutBar): boolean {
  return bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "carrier" || piece.type === "mounting");
}

function isUdCutBar(bar: CutBar): boolean {
  return bar.pieces.length > 0 && bar.pieces.every((piece) => piece.type === "ud");
}

function countOptimizedBars(bars: CutBar[]): OptimizedProfileCounts {
  return {
    cd: bars.filter(isCdCutBar).length,
    ud: bars.filter(isUdCutBar).length,
  };
}

function buildCutConfig(constants: CalculatorConstants) {
  return {
    stockLengthCm: Math.max(constants.cdLength, constants.udLength) * 100,
    kerfCm: 0.3,
    minReusableOffcutCm: 20,
    perTypeStockLengthsCm: {
      carrier: constants.cdLength * 100,
      mounting: constants.cdLength * 100,
      ud: constants.udLength * 100,
    },
  };
}

self.onmessage = (event: MessageEvent<CutOptimizationWorkerRequest>) => {
  const message = event.data;

  try {
    const config = buildCutConfig(message.constants);
    if (message.type === "optimize-cut-plan") {
      const plan = message.mode === "global"
        ? optimizeAllRoomsSuspendedCeilingCuts(message.rooms, message.constants, config)
        : (() => {
          const room = cloneRoom(message.room);
          const result = calc(room, message.constants);
          return optimizeSuspendedCeilingCuts(buildCutOptimizationInput(room, result, message.constants), config);
        })();
      const response: CutOptimizationWorkerResponse = {
        type: "cut-plan-result",
        requestId: message.requestId,
        plan,
      };
      self.postMessage(response);
      return;
    }

    if (message.type === "build-material-takeoff") {
      const response: CutOptimizationWorkerResponse = {
        type: "material-takeoff-result",
        requestId: message.requestId,
        rows: buildMaterialTakeoff(message.rooms, message.constants, { includeOptimizedQuantities: false }),
      };
      self.postMessage(response);
      return;
    }

    const rows: Record<string, OptimizedProfileCounts> = {};

    for (const sourceRoom of message.rooms) {
      try {
        const room = cloneRoom(sourceRoom);
        const result = calc(room, message.constants);
        const plan = optimizeSuspendedCeilingCuts(buildCutOptimizationInput(room, result, message.constants), config);
        rows[room.id] = countOptimizedBars(plan.bars);
      } catch {
        rows[sourceRoom.id] = { cd: null, ud: null };
      }
    }

    const globalPlan = optimizeAllRoomsSuspendedCeilingCuts(message.rooms, message.constants, config);
    const response: CutOptimizationWorkerResponse = {
      type: "room-cut-optimization-result",
      requestId: message.requestId,
      rows,
      global: countOptimizedBars(globalPlan.bars),
    };
    self.postMessage(response);
  } catch (error) {
    if (message.type === "optimize-cut-plan") {
      const response: CutOptimizationWorkerResponse = {
        type: "cut-plan-error",
        requestId: message.requestId,
        error: error instanceof Error ? error.message : "Неуспешен фонов разкрой.",
      };
      self.postMessage(response);
      return;
    }

    if (message.type === "build-material-takeoff") {
      const response: CutOptimizationWorkerResponse = {
        type: "material-takeoff-error",
        requestId: message.requestId,
        error: error instanceof Error ? error.message : "Неуспешен фонов материален списък.",
      };
      self.postMessage(response);
      return;
    }

    const response: CutOptimizationWorkerResponse = {
      type: "room-cut-optimization-error",
      requestId: message.requestId,
      error: error instanceof Error ? error.message : "Неуспешен фонов разкрой.",
    };
    self.postMessage(response);
  }
};
