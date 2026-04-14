# Suspended Ceiling Calculator

## Project Overview

This repository is a React + TypeScript calculator for Knauf D11 suspended ceiling systems. The UI is in Bulgarian and supports multiple rooms, construction type selection, live SVG visualization, JSON import/export, and a simple Excel XML export.

Main files:
- `index.html` is the Vite HTML entry point.
- `src/main.tsx` mounts the React app.
- `src/App.tsx` contains the current UI, import/export handlers, room editing flow, result table, material panel, and SVG visualization.
- `src/styles.css` contains the app styling.
- `src/domain/calculator.ts` contains construction tables, typed domain models, validation, position generation, and calculation logic.
- `src/domain/calculator.test.ts` covers the current domain behavior with Vitest.
- `task.md` tracks current known follow-up work and technical TODOs.
- `BRO D11 BG 2024 03.pdf` is the Knauf reference document used for table/formula checks.

## Runtime And Tests

Install dependencies with:

```powershell
npm install
```

Run the local development server with:

```powershell
npm run dev
```

Run the current tests with:

```powershell
npm run test
```

Run the production build with:

```powershell
npm run build
```

Keep domain logic browser-independent where practical so it can stay covered by Vitest.

## Development Notes

- Default shell for this workspace is PowerShell 7 on Windows.
- Prefer PowerShell syntax in commands.
- Keep Bulgarian user-facing labels and copy consistent with the existing UI.
- Calculation behavior lives in `src/domain/calculator.ts`; avoid duplicating formulas in React components.
- If adding tests, extend Vitest coverage near the domain code.
- `localStorage` key is currently `d113-calculator-v2`; changing it can affect existing saved browser data.

## Current Functional Shape

Supported construction types are configured in `src/domain/calculator.ts` under `CONSTRUCTION_TYPES`:
- `D111`
- `D112`
- `D113`
- `D116`

Important exported functions from `src/domain/calculator.ts` include:
- `calc`
- `validateCombination`
- `syncSpacingFromKnaufTable`
- `getTableValue`
- `getValidCValues`
- `countBySpacing`

These exports are used by Vitest tests; keep them stable or update tests intentionally.

## Known Follow-Up Areas

See `task.md` before starting feature work. Current TODO themes include:
- Rename the generic `CD бр.` result label.
- Split material takeoff by construction type instead of using one CD/UD model for all systems.
- Improve `Дюбели UD` spacing rules.
- Show auto/manual state for `a`, `b`, and `c`, with a reset-to-Knauf action.
- Add more Knauf table tests, especially D113 `c = 600` values.
- Improve visualization positioning for installation-grade drawings.
