# CLAUDE.md — Workbook Manager Excel Add-in

> **Read this at the start of every session.** This document is the single source of truth for architecture, design decisions, and conventions. It is updated as the project evolves — always trust this file over stale comments in code.

---

## 1. Project Overview

**Workbook Manager** is an Excel Office add-in that replaces and augments Excel's built-in ribbon functionality with a modern, friendly task pane UI. It is not a general-purpose spreadsheet tool — it is purpose-built for **financial modelers and report builders** whose workbooks are shared, maintained over time, and require power features like bulk named range management, complex formula composition, shape/layout tools, and live cell watching.

### Target users
Power users who need both depth and polish. They know Excel well and will reject anything that feels clunky or slow. The UI must earn their trust by being faster and more capable than the native ribbon, not just prettier.

### Platform strategy
- **Primary platform**: Excel Desktop (Windows). All features should target this first.
- **Excel Web**: Show a persistent banner — *"Some features require Excel Desktop. [Open in Desktop ↗]"* — but do not block all functionality. Degrade gracefully where the Office JS API supports it (e.g., WatchView already does this for the dependents API).
- **Detection**: `Office.context.platform` at load time. If web, display banner and flag limited features inline.

---

## 2. Tech Stack

| Layer | Library / Tool | Version | Notes |
|---|---|---|---|
| UI framework | React | 18.2 | Functional components + hooks only |
| TypeScript | typescript | 5.4 | Strict mode |
| Component library | **Fluent UI v9** (`@fluentui/react-components`) | ^9.55.1 | See §4 — not locked in |
| Icons | `@fluentui/react-icons` | ^2.0.264 | Fluent icon set |
| CSS-in-JS | Griffel (`makeStyles`) | via Fluent UI v9 | See §9 for usage rules |
| Code editor | Monaco (`@monaco-editor/react`) | ^4.7.0 | Used in Formulas tab |
| Bundler | Webpack | 5.95 | Config in `webpack.config.js` |
| State management | *(none yet — local React state)* | — | Zustand planned — see §7 |
| Office bridge | `@types/office-js`, `@types/office-runtime` | latest | Office JS API + OfficeRuntime.storage |

### Notable absences (intentional)
- No Zustand yet — adding it is a **planned architecture change** (see §7)
- No React Router — navigation is a simple string enum via `NavigationTarget`
- No CSS modules, no Tailwind — Griffel `makeStyles` is the only styling system

---

## 3. Source Layout

```
src/
├── commands/
│   ├── commands.html          # Ribbon command handler entry point
│   └── commands.ts            # Ribbon button actions; writes signals to trigger task pane
│
├── hooks/
│   └── useDebounce.ts
│
├── shared/
│   ├── featureFlags.ts        # isSandboxDebugEnabled() — query-param/localhost gating
│   └── signals.ts             # ⚠ Canonical signal key registry (see §8)
│
├── taskpane/
│   ├── components/
│   │   ├── App.tsx            # Root: mode switcher (legacy/modern), signal polling loop
│   │   ├── LegacyApp.tsx      # Legacy UI — kept alive while modern port is incomplete
│   │   ├── modern/
│   │   │   ├── designTokens.ts        # MODERN_TOKENS + useModernSharedStyles
│   │   │   ├── ModernShell.tsx        # Layout: header + collapsible nav rail + content area
│   │   │   ├── FormulaMonacoView.tsx  # Formulas tab (Monaco editor + formula tools)
│   │   │   ├── NamesView.tsx          # Names tab (1,689 lines — CRUD, bulk edit, LAMBDA)
│   │   │   ├── FormatView.tsx         # Format tab (3,641 lines — monolith, needs split)
│   │   │   ├── WatchView.tsx          # Watch tab (well-written, reference impl)
│   │   │   ├── TablesView.tsx
│   │   │   ├── PivotsView.tsx         # Stub
│   │   │   ├── QueriesView.tsx
│   │   │   ├── ModelBuilderView.tsx
│   │   │   ├── SettingsView.tsx       # Stub
│   │   │   ├── HelpView.tsx
│   │   │   └── SandboxDebugView.tsx   # Dev-only, gated by featureFlags
│   │   └── FormattingTab.tsx, Header.tsx, HeroList.tsx, TextInsertion.tsx  (legacy)
│   │
│   ├── navigation.ts          # NavigationTarget type + NAVIGATION_TARGETS set
│   ├── taskpane.ts            # ⚠ Monolithic Office JS API file — split planned (see §6)
│   ├── taskpane.html
│   ├── index.tsx              # React root mount
│   └── utils/
│       └── nameTransforms.ts
│
└── utils/
    └── names.utils.ts         # Pure utilities: case transforms, LAMBDA formula helpers
```

### Planned additions (not yet created)
```
src/
├── api/                       # Domain-split from taskpane.ts (see §6)
│   ├── names.ts
│   ├── formulas.ts
│   ├── format.ts
│   ├── tables.ts
│   ├── queries.ts
│   ├── watch.ts
│   └── workbook.ts
└── store/
    └── workbookStore.ts       # Zustand store (see §7)
```

---

## 4. Visual Design System

### Design direction
**Notion-inspired**: clean, intuitive, content-forward. The goal is to bring Excel into a contemporary UX without feeling like a Microsoft product from 2019. The add-in should have its own visual identity.

### Color palette (from `designTokens.ts`)

| Token | Value | Usage |
|---|---|---|
| `colorBg` | `#F9FAFB` | Page/pane background (light warm gray) |
| `colorSurface` | `#FFFFFF` | Cards, panels, inputs |
| `colorText` | `#111827` | Primary text |
| `colorTextMuted` | `#6B7280` | Labels, hints, secondary text |
| `colorBorder` | `#E5E7EB` | Dividers, card outlines |
| `colorBrand` | `#2E7D32` | Primary brand (forest green, close to Excel green `#217346`) |
| `colorBrandStrong` | `#256729` | Hover/active brand states |
| `colorAccent` | `#F59E0B` | Amber — warnings, highlights |
| `colorDanger` | `#B91C1C` | Error states |
| `shadowCard` | subtle 2-layer box-shadow | Card resting state |
| `shadowCardHover` | deeper 2-layer box-shadow | Card hover state |

**Active nav item background**: `#E8F4EA` (light green tint). This is intentionally not a Fluent token — it comes from `MODERN_TOKENS`.

### Why Excel green, not Microsoft blue
Fluent UI v9's default theme is blue-dominant. This add-in avoids that. The green anchor (`#2E7D32`) ties the UI to Excel's identity while creating separation from Microsoft's corporate blue palette. Do not introduce Fluent's blue semantic tokens (e.g., `colorBrandBackground`) into the main UI. Use `MODERN_TOKENS` instead.

### Typography
**Not yet finalized.** Current state: Fluent UI v9 defaults (Segoe UI / system-ui). Recommended direction: `Inter` (Google Fonts or self-hosted) — it pairs well with Notion-style UX, is highly legible at small sizes in task panes, and avoids the Segoe/Office look. This should be evaluated and locked in; update this document when decided.

Type scale in use (from `designTokens.ts`):
- Section title: `20px / 700`
- Card title: `16px / 600`
- Body: `~13–14px / 400`
- Muted/label: `12px / 400`
- Section subtitle: `13px / 400`

### Component library status
Fluent UI v9 is the current library but is **not locked in**. If a future session recommends switching (e.g., to Radix UI + Tailwind, or shadcn/ui), justify the recommendation with: bundle size impact, Office Web compatibility, Griffel replacement strategy, and migration cost for existing `makeStyles` calls. Do not make the switch without that analysis documented here.

### Layout
- **Shell**: CSS Grid, `gridTemplateRows: "56px 1fr"` — fixed header, scrollable content
- **Nav rail**: Left-side icon rail (`gridTemplateColumns: "auto 1fr"`), collapsible, sticky
- **Content area**: `overflowY: auto`, scrollable, padding varies by tab (Formula tab: `12px 16px`, Format tab: `0 24px 24px`, others: `24px`)
- **NOT fixed-height**: The pane is scrollable and content-adaptive per tab. Avoid `height: 100vh` in content children.

---

## 5. Navigation & Shell Architecture

### Shell (`ModernShell.tsx`)
The shell owns the layout and nav rail. It receives `activeTarget` and `onTargetChange` as props from `App.tsx` (which owns the state). This keeps navigation state at the App level, accessible to signal handlers.

### Nav items (current)
The nav rail renders these items in order. All are always visible:

| ID | Label | Status |
|---|---|---|
| `names` | Names | ✅ Active (needs parity work) |
| `tables` | Tables | ✅ Active |
| `formulas` | Formulas | ✅ Active (major work planned) |
| `queries` | Queries | ✅ Active |
| `model-builder` | Model Builder | ✅ Active |
| `pivots` | Pivots | ⚠ Stub |
| `format` | Format | ✅ Active (refactor needed) |
| `watch` | Watch Window | ✅ Well-written, near-final |
| `sandbox-debug` | Sandbox | 🔧 Dev-only (featureFlag gated) |

Settings and Help are accessible from the `...` menu in the header, not the nav rail.

### NavigationTarget type (`navigation.ts`)
```typescript
export type NavigationTarget =
  | "names" | "names-create" | "tables" | "formulas"
  | "queries" | "model-builder" | "pivots" | "format"
  | "sandbox-debug" | "watch" | "settings" | "help";
```
Always add new nav destinations here first. Use `isNavigationTarget()` to validate signal values before acting on them.

### Planned: tab structure per spec
Per product spec, the user-visible tabs should be: **Formulas, Names, Format, Pivots, Queries, Watch**. The current nav rail has a superset of this (Tables, Model Builder, Sandbox). Tab reconciliation — deciding which tabs appear in the primary rail vs. overflow/settings — is a **pending design decision**.

---

## 6. Monolithic API File — Split Plan

### Current state
`src/taskpane/taskpane.ts` is enormous (~164KB). Every Office JS call — named ranges, formulas, format tools, watch, queries — is exported from this single file. Every view imports from `../../taskpane`.

### Planned split
Break into domain modules under `src/api/`. Each module should:
- Import and use `Excel.run()` directly
- Be independently testable (or at least independently readable)
- Export only types and functions relevant to its domain

```
src/api/names.ts       → getNamedRanges, addNamedRange, updateNamedRange, deleteNamedRange…
src/api/formulas.ts    → evaluateFormula, insertFormula, beautifyFormula…
src/api/format.ts      → captureSheetFormat, applySheetFormat, freezePanes, shapeBuilder…
src/api/tables.ts      → listTables, getTableData…
src/api/queries.ts     → listPowerQueries, refreshQuery…
src/api/watch.ts       → getWatchWindowData, selectCellAddress…
src/api/workbook.ts    → getWorkbookContext (sheets, named ranges list, active cell)
```

**Do not do this split in one large refactor.** Extract one domain at a time, update imports in the relevant view, and verify the view still works before moving on.

---

## 7. State Management

### Current state
All state is local React component state (`useState`). Each view fetches its own data independently via Office JS when it mounts.

### Planned: Zustand global store
Add Zustand for **shared workbook context** — data that multiple tabs need:

```typescript
// src/store/workbookStore.ts
interface WorkbookState {
  sheetNames: string[];
  namedRanges: NamedRangeRecord[];
  tableNames: string[];
  activeCellAddress: string | null;
  lastRefreshedAt: number;

  refreshWorkbookContext: () => Promise<void>;
}
```

**Why Zustand over Context**: Zustand avoids prop drilling and React Context re-render storms. It also has zero-boilerplate selectors. Keep the store small — it should hold *shared workbook metadata*, not per-tab UI state.

**Per-tab UI state** (accordion open/close, filter text, sort column) stays local to each view component. Use `OfficeRuntime.storage` only for state that should survive tab reloads.

### Signal/IPC stays as-is
The `OfficeRuntime.storage` polling loop in `App.tsx` (900ms interval) is the bridge from Office JS commands (ribbon buttons in `commands.ts`) to the React app. This pattern is intentional — Office JS commands run in a separate JavaScript context and cannot call React directly. Do not replace this with a different IPC mechanism without understanding why it exists.

---

## 8. Signal / IPC Architecture

### How it works
1. User clicks a ribbon button → `commands.ts` writes a signal value to `OfficeRuntime.storage` (and falls back to `Office.context.document.settings` for Excel Web).
2. `App.tsx` polls storage every 900ms in a `useEffect` cleanup loop.
3. On detecting a signal, App clears it and dispatches the appropriate action (navigate, open formula editor, inject selection, etc.).

### Signal key registry (`src/shared/signals.ts`)
```typescript
export const Signals = {
  NAVIGATION_REQUEST: "wbm.navigation.request",
  FORMULA_EDITOR_REQUEST: "wbm.openFormulaEditor.request",
} as const;
```

**⚠ Known gap**: `navigation.ts` also defines `NAVIGATION_SIGNAL_KEY = "wbm.navigation.request"` as an inline string. This is a duplication. The canonical definition should live only in `shared/signals.ts`. Fix by importing from `shared/signals.ts` in `navigation.ts`.

**⚠ Known gap**: `WatchView.tsx` defines `WATCH_ADD_SIGNAL_KEY = "wbm.watch.addAddress"` and `WATCH_ADD_SHEET_KEY = "wbm.watch.addSheet"` as inline string constants. These should be added to `shared/signals.ts`.

### Rule: never inline signal key strings
All signal key strings must be defined in `src/shared/signals.ts` and imported from there. No exceptions. This makes refactoring signal keys safe and makes it possible to audit all IPC paths in one place.

### Formula editor signals (current values)
These are written by `commands.ts` to `OPEN_FORMULA_EDITOR_SIGNAL_KEY`:

| Signal value | Effect |
|---|---|
| `"open-only"` | Navigate to Formulas tab, open editor |
| `"open-and-pull"` | Navigate + pull active cell formula into editor |
| `"formula-pull"` | Pull formula (tab already open) |
| `"formula-apply"` | Apply editor formula back to cell |
| `"formula-beautify"` | Beautify/format formula in editor |
| `"formula-insert-selection"` | Insert selected cell reference at cursor |

These string literals are currently defined inline in `App.tsx` and should be moved to `shared/signals.ts`.

---

## 9. Styling Conventions (Griffel / makeStyles)

### The rule
`makeStyles()` **defines a hook** (it returns a hook function). The convention is:

```typescript
// ✅ CORRECT — define at module scope, call inside component
const useStyles = makeStyles({ ... });

const MyComponent = () => {
  const styles = useStyles();  // called inside component
  return <div className={styles.root} />;
};
```

```typescript
// ❌ WRONG — calling makeStyles result at module scope
const styles = makeStyles({ ... })();  // never do this
```

```typescript
// ✅ CORRECT — shared styles from designTokens.ts
const MyComponent = () => {
  const shared = useModernSharedStyles();
  const own = useStyles();
  return <div className={`${shared.card} ${own.specialCard}`} />;
};
```

### Naming convention
- Local: `const useStyles = makeStyles({...})` — used in a single component file
- Shared: `const useXxxStyles = makeStyles({...})` exported from `designTokens.ts` — used across multiple components
- Always prefix hook names with `use`

### Always use MODERN_TOKENS for colors
Do not hardcode color hex values inline in `makeStyles` calls. Reference `MODERN_TOKENS.colorBrand`, `MODERN_TOKENS.colorBorder`, etc. Exception: one-off values like nav active background `#E8F4EA` are acceptable when they derive conceptually from a token but don't map to an existing one — add a comment explaining the intent.

---

## 10. Feature Roadmap by Tab

### Priority order
1. **Formulas** — highest gap, most architectural work
2. **Names** — significant regression from legacy, must reach parity
3. **Format** — functional but needs decomposition
4. **Pivots / Settings** — stubs, build after 1 and 2 are solid
5. **Watch** — near-complete, minor cleanup only
6. **Queries / Tables / Model Builder** — active, ongoing

---

### 10.1 Formulas Tab (`FormulaMonacoView.tsx`)

**Vision**: The best formula editing experience available in Excel, period.

#### Two modes
- **Advanced mode**: Monaco editor, full-screen feel, live evaluator. Mirrors the native formula bar but more powerful.
- **Wizard mode**: Step-by-step guided builder for complex/nested formulas. Breaks nested IF, VLOOKUP, LAMBDA, etc. into structured argument slots. Forces the user to express their logic one layer at a time. Produces a correct nested formula at the end.

#### Advanced mode — requirements
- Monaco editor with Excel formula syntax highlighting
- **Live evaluator**: evaluates the current formula automatically as the user types (debounced, not on button click). Shows result or error inline below the editor.
- **Reference injection**: when the user clicks a cell or selects a range in the Excel grid while the formula editor is open, inject that reference at the cursor position — matching native Excel formula bar behavior. This uses the `formula-insert-selection` signal from `commands.ts`.
- Table/range reference picker: allow hovering/clicking named ranges and tables to inject them.

#### Wizard mode — requirements
- Parse the formula structure into a tree of argument slots
- Render each argument slot as a labeled input with its own mini-evaluator
- Allow nested "sub-wizards" for nested functions
- **LAMBDA/builder**: the current experience needs a full rethink — it is clunky. The Wizard mode should make LAMBDA authoring feel like filling out a form.
- At any point, the user should be able to switch to Advanced mode and see the combined formula.

#### Popout support
The formula editor supports a `?popout=formula` URL mode where it renders standalone without the shell nav. This already works — preserve it when refactoring.

---

### 10.2 Names Tab (`NamesView.tsx`)

**Current state**: 1,689 lines. Has CRUD, scope management, LAMBDA formula support, named list support, case transforms, debounced search.

**The regression problem**: The legacy app (`LegacyApp.tsx`) had a working **bulk editor** — rename, repath, delete, scope change on multiple named ranges at once. The modern port (`NamesView.tsx`) does not have this fully functional. This is the primary gap.

#### Requirements before legacy can be retired
- [ ] Full CRUD: create Range / Function (LAMBDA) / List entries
- [ ] Inline edit: name, address, scope (worksheet vs workbook)
- [ ] **Bulk select + bulk edit**: rename pattern (find/replace in name), repath (change address prefix), delete, scope change — all applied to a selection of multiple records
- [ ] Sort by any column (name, address, sheet, scope, type)
- [ ] Filter/search with debounce
- [ ] `names-create` signal support: open the create panel when triggered from ribbon
- [ ] Graceful empty state when no named ranges exist

#### Do not remove the legacy app until parity is confirmed
Keep `LegacyApp.tsx` and the "Open Legacy UI" escape hatch until a side-by-side parity review is done against `docs/manual-host-parity-checklist.md`.

---

### 10.3 Format Tab (`FormatView.tsx`)

**Current state**: 3,641-line monolith. Fully functional but unmaintainable as-is.

#### Sub-panels (decomposition plan)
Break into focused sub-components:

| Sub-panel | Responsibility |
|---|---|
| `ShapeBuilderPanel` | Create, style, position, z-order shapes |
| `SheetFormatPanel` | Capture and apply sheet format presets |
| `NavigationPanel` | Navigation destinations (hyperlink-style nav within workbook) |
| `FreezePanel` | Freeze rows/columns, unfreeze |
| `UtilitiesPanel` | Gridlines toggle, column/row sizing, alignment |
| `CellStylePanel` | Apply named cell style presets |
| `TableStylePanel` | Apply Excel table styles |

Each sub-panel should be a separate file in `components/modern/format/`. The parent `FormatView.tsx` becomes a thin orchestrator that renders accordions around each sub-panel.

#### Accordion pattern
All Format sub-panels use the accordion pattern. Multiple sections can be open simultaneously. Open/closed state persists via `OfficeRuntime.storage` (key: `wbm.ui.accordionState.format`).

#### Popout support
Format tab also supports `?popout=format`. Preserve this.

---

### 10.4 Watch Tab (`WatchView.tsx`)

**Reference implementation** — this file is the best-written component in the codebase. Study it before writing new components.

Features already working:
- Persists watch list via `OfficeRuntime.storage` (key: `wbm.watchWindow.items`)
- Live polling every 3,000ms
- Shows value, formula, dependent/precedent counts
- Inline label editing (double-click)
- Graceful degradation on Excel Web (dependents API not available → shows "–")
- Accepts address from context menu signal (`wbm.watch.addAddress`)

Remaining work: minor cleanup only. No major feature work needed.

---

### 10.5 Pivots, Queries, Settings (stubs / partial)

Build these after Formulas and Names are solid. Do not start stub-to-feature work on these until the top priorities are resolved.

---

## 11. UX Patterns

### Accordion sections
- Multiple sections **can** be open simultaneously within a tab (not mutually exclusive)
- Open/closed state persists per-tab via `OfficeRuntime.storage`
- Storage key pattern: `wbm.ui.accordionState.{tabId}`
- On first load (no stored state), default to first section open

### Empty states
When a tab has no content (no named ranges, no power queries, etc.):
- Centered icon + short label + primary action button (`+` to add first item)
- One-line explanation of what the tab does
- Never show an empty list or blank space — always guide the user to the action

```
Example:
  [icon: BookNumber]
  No named ranges in this workbook
  [+ Create your first named range]
```

### Error handling
- **Inline errors**: within the relevant panel section. Not full-page takeovers.
- **Subtle error banner style**: icon + short message + optional retry button. Use `colorDanger` from tokens.
- **Toast notifications**: for async operations (save succeeded, rename failed). Use Fluent UI Toast component.
- **No modal dialogs for errors**: keep the user in context. Dialogs are acceptable only for destructive confirmation (delete all named ranges, etc.) and even then, prefer inline confirmation patterns.

### Web fallback banner
On load, if `Office.context.platform === Office.PlatformType.OfficeOnline`:
```
⚠ Some features require Excel Desktop. [Open in Desktop ↗]
```
- Displayed as a persistent top-of-pane banner below the header
- Does not block the UI — degrade gracefully where possible
- Inline feature limitations (e.g., WatchView's "–" for dependents) are preferred over disabling entire tabs

---

## 12. Architecture Conventions

### Do ✅

- **One source of truth for signal keys**: `src/shared/signals.ts`
- **One source of truth for types**: `src/shared/types.ts` (to be created — consolidate shared types here)
- **One source of truth for navigation targets**: `src/taskpane/navigation.ts`
- **`makeStyles` always returns a hook, called inside the component**
- **Use `MODERN_TOKENS` for all colors in `makeStyles` calls**
- **Domain API modules**: when splitting `taskpane.ts`, put each domain's functions in `src/api/{domain}.ts`
- **Comment signal key usages** with the corresponding commands.ts trigger, so IPC paths are traceable
- **WatchView.tsx is the style reference** — follow its documentation and structure patterns for new views

### Do not ❌

```
❌ makeStyles at module scope (calling the result, not defining the hook)
❌ Inline signal key strings — always import from shared/signals.ts
❌ Duplicate type definitions across view files — put shared types in shared/types.ts
❌ Dynamic imports for modules already statically imported
❌ Building features that assume Excel Web parity without flagging the limitation inline
❌ Removing LegacyApp.tsx before the parity checklist is signed off
❌ Using Fluent UI's blue semantic color tokens in the main UI
❌ Hardcoded hex colors inside makeStyles — use MODERN_TOKENS
```

---

## 13. Open Design Decisions (TBD)

These are intentionally unresolved — do not invent answers; flag them and discuss.

| Decision | Options | Status |
|---|---|---|
| **Typography / font** | Keep system-ui (Segoe), adopt Inter, adopt DM Sans | **Unresolved** — recommend Inter, needs evaluation |
| **Nav rail vs. top tab bar** | Current: left icon rail (collapsible). Spec noted "top tab bar TBD." | **Unresolved** — current left rail feels right for a tall pane; revisit if content area feels cramped |
| **Tab set reconciliation** | Current nav has: Names, Tables, Formulas, Queries, Model Builder, Pivots, Format, Watch. Product spec says: Formulas, Names, Format, Pivots, Queries, Watch. | **Unresolved** — decide what happens to Tables and Model Builder |
| **Component library switch** | Keep Fluent UI v9 or migrate | **Unresolved** — only switch with full justification (bundle size, Office Web compat, migration cost) |
| **Accordion component** | Use Fluent UI v9 `Accordion` vs. custom | Fluent v9 Accordion is available — evaluate it vs. current custom implementations |
| **Zustand introduction** | When and how to add, what goes in the store | **Planned** — do not add until Names and Formulas tabs are stable |

---

## 14. Key Files Quick Reference

| File | Purpose | Health |
|---|---|---|
| `src/shared/signals.ts` | Signal key registry | ⚠ Incomplete — some keys still inline in views |
| `src/shared/featureFlags.ts` | Feature gates (sandbox debug) | ✅ Small, clean |
| `src/taskpane/navigation.ts` | NavigationTarget type + validator | ⚠ Duplicates NAVIGATION_SIGNAL_KEY from signals.ts |
| `src/taskpane/taskpane.ts` | Office JS API monolith | ⚠ ~164KB — split planned |
| `src/taskpane/components/App.tsx` | Root, mode switch, signal polling | ✅ Clean, readable |
| `src/taskpane/components/modern/designTokens.ts` | Color/shadow tokens + shared styles | ✅ Clean — always extend here, not inline |
| `src/taskpane/components/modern/ModernShell.tsx` | Shell layout + nav rail | ✅ Clean, readable |
| `src/taskpane/components/modern/WatchView.tsx` | Watch Window | ✅ **Reference implementation** |
| `src/taskpane/components/modern/NamesView.tsx` | Names CRUD + bulk edit | ⚠ 1,689 lines — bulk edit regression vs. legacy |
| `src/taskpane/components/modern/FormatView.tsx` | Format tools | ⚠ 3,641 lines — decomposition needed |
| `src/taskpane/components/modern/FormulaMonacoView.tsx` | Formula editor | ⚠ Major feature work planned |
| `src/taskpane/components/LegacyApp.tsx` | Legacy fallback UI | 🔒 Keep alive until parity confirmed |
| `docs/manual-host-parity-checklist.md` | Legacy parity checklist | 📋 Gate for retiring LegacyApp |

---

## 15. Development Workflow

```bash
npm start          # Start local dev server + sideload in Excel Desktop
npm run build:dev  # Development build (source maps, unminified)
npm run build      # Production build
npm run watch      # Watch mode build
npm run lint       # ESLint check
npm run lint:fix   # ESLint autofix
```

The dev server runs on port 3000 (set in `package.json` config). Webpack config is in `webpack.config.js`. Monaco editor files are copied to `dist/monaco/` by copy-webpack-plugin.

Manifest: `manifest.xml` — controls ribbon buttons, add-in metadata. Signals written by ribbon buttons in `commands.ts` are consumed by `App.tsx`'s polling loop.

---

*Last updated: 2026-04-06. Update this file whenever architectural decisions change.*
