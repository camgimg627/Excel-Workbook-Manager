# Legacy-to-Modern Functionality Parity Review (Current Workspace Baseline)

## Review Metadata
- Review date: 2026-03-10
- Baseline branch/worktree: `parity/legacy-modern-uncertain` (started from current dirty workspace state)
- Parity definition: outcome-equivalent behavior, not strict control-by-control UI matching
- Evidence standard: static code audit + manual Excel host checklist
- Scope: all legacy user-facing workflows that produce workbook outcomes

## Status Legend
- `Full parity`: Modern has an equivalent trigger and workbook outcome.
- `Outcome-equivalent parity`: Different UX/API path, same workbook outcome.
- `Partial parity`: Modern covers some but not all legacy outcomes.
- `Missing`: No modern path to achieve legacy outcome.
- `Uncertain`: Code path exists, but host/runtime behavior must be validated.

---

## 1) Canonical Legacy Capability Catalog

| Capability ID | Legacy trigger path | User intent | Expected workbook outcome | Preconditions |
| --- | --- | --- | --- | --- |
| NAM-001 | `LegacyApp.tsx:3507` (Names > Ranges grid) | Inspect workbook names/shapes with filter/sort/select | Named item inventory is loaded and browsable | Workbook open |
| NAM-002 | `LegacyApp.tsx:3615` (`goToRange`) | Jump to named range/cell by link | Target sheet activates and target range is selected | Linked record is range-like |
| NAM-003 | `LegacyApp.tsx:3615` + `LegacyApp.tsx:3143` | Open shape editing from shape row in Names | Format > Shapes editing context opens for that shape | Shape record exists |
| NAM-004 | `LegacyApp.tsx:3591`, modal at `5405+` | Edit name/address | Name and/or definition updates in workbook names collection | Selected named item |
| NAM-005 | `LegacyApp.tsx:3599`, modal at `5471+` | Move named range | Named range points to new address | Selected range + address |
| NAM-006 | `LegacyApp.tsx:3607`, modal at `5508+` | Delete name and/or values | Name removed and/or target cells cleared | Selected named item |
| NAM-007 | `LegacyApp.tsx:3521`, modal at `5570+` | Bulk transform/delete names | Batch rename/delete succeeds with preview rules | Selected rows |
| TBL-001 | `LegacyApp.tsx:3638` (Names > Tables grid) | Inspect workbook tables with filter/sort/select | Table inventory loaded and browsable | Workbook open |
| TBL-002 | `LegacyApp.tsx:3719` (inline edit/save) | Rename one table | Target table renamed | Table selected |
| TBL-003 | `LegacyApp.tsx:3657`, modal at `5873+` | Batch rename tables | Bulk rename succeeds with conflict avoidance | Selected tables |
| FML-001 | `LegacyApp.tsx:4844`, `4847`, `4872`, `5046` | Pull/apply/beautify/insert selection | Active-cell formula sync and editor transforms | Active worksheet/cell |
| FML-002 | `LegacyApp.tsx:5237` | Evaluate current formula/function call | Grid output rendered from evaluation | Test formula present |
| FML-003 | `LegacyApp.tsx:5009` and `5051` | Author LAMBDA-based function text | LAMBDA formula body/invocation prepared | Function mode enabled |
| FML-004 | `LegacyApp.tsx:5139` | Use inline suggestions and context hints | Formula text composed with function/name/table hints | Editor focused |
| FML-005 | `LegacyApp.tsx:5199` context commands | Use go-to-definition/advanced context actions | Definition jump and editor actions execute | Token under cursor |
| FML-006 | `LegacyApp.tsx:4886`, popout flow at `5270+` | Continue editing in popout | Formula editor opens in dialog and remains operational | Office dialog allowed |
| FML-007 | `LegacyApp.tsx:4902` settings panel | Tune formula editor behavior | Line numbers, capture, palette settings applied | Formula panel open |
| FMT-001 | `LegacyApp.tsx:3768` | Apply cell style presets | Selection receives preset style | Selection exists |
| FMT-002 | `LegacyApp.tsx:3847` (insert shape) | Insert shape/text box on active cell | New shape appears anchored to selection | Selection exists |
| FMT-003 | `LegacyApp.tsx:4125` (`submitShapeRename`) | Rename selected shape | Shape name changes | Active shape |
| FMT-004 | `LegacyApp.tsx:3882`, `3972` | Edit shape style/text/size attributes | Shape formatting and text update | Active shape |
| FMT-005 | `LegacyApp.tsx:4455`, `4473`, `4493` | Arrange shape (position/nudge/align/z-order) | Shape order/position changes | Active shape + optional selection |
| FMT-006 | `LegacyApp.tsx:4536` | Change geometric shape type | Shape geometry changes | Non-text shape |
| FMT-007 | `LegacyApp.tsx:4658`, `4661`, `3956`, service `getNamedRangeValueText` | Bind shape text to named range values | Shape text reflects linked named value | Active shape + named range |
| FMT-008 | `LegacyApp.tsx:4664`, service `applyFormulaToShapeAnchorCell` | Drive shape text from anchor-cell formula | Anchor cell formula applies and shape text updates | Active shape + formula |
| FMT-009 | `LegacyApp.tsx:4128`, `4025` | Move shape to selected cell and carry value | Shape anchor moves and anchor value transfer happens | Active shape + target selection |
| FMT-010 | `LegacyApp.tsx` open format tools in-pane | Perform format utility actions (grid/pivot/table etc.) | Workbook utility action applied | Valid worksheet context |
| UTL-001 | `LegacyApp.tsx:5324` command bar | Trigger quick utilities (text/pivots/grid/style/accent) | Utility action completes on workbook | Active workbook |
| SIG-001 | `commands.ts` + `App.tsx:96` | Ribbon navigation routing | Task pane opens at requested section | Signal storage available |
| SIG-002 | `commands.ts` + `App.tsx:136` | Ribbon formula action routing | Formula action executes against modern formula view | Formula view mounted |

---

## 2) Modern Mapping Against Legacy Capabilities

| Capability ID | Modern status | Modern trigger + evidence | Service/API evidence | Gap summary | Severity |
| --- | --- | --- | --- | --- | --- |
| NAM-001 | Full parity | `modern/NamesView.tsx:946`, grid/actions `1040+` | `getNamedRanges` (`taskpane.ts:2844`) | Core inventory/filter/sort/select present | - |
| NAM-002 | Uncertain (host/runtime dependent) | Link handler `NamesView.tsx:907` | `selectNamedRangeAddress` (`taskpane.ts:3134`) | Code path exists; reported runtime no-op must be host-validated/fixed | P0 |
| NAM-003 | Missing | Shape rows are non-link text at `NamesView.tsx:1124-1139` | n/a | Legacy shape-row jump-to-editor workflow not implemented | P1 |
| NAM-004 | Full parity | Edit modal + submit `NamesView.tsx:777`, `1310+` | `updateNamedRange` (`taskpane.ts:2970`) | Supports range/list/function-formula updates | - |
| NAM-005 | Full parity | Move modal + submit `NamesView.tsx:821`, `1383+` | `moveNamedRange` (`taskpane.ts:3091`) | Equivalent outcome | - |
| NAM-006 | Full parity | Delete modal + submit `NamesView.tsx:839`, `1414+` | `deleteNamedRangeWithOptions` (`taskpane.ts:3101`) | Equivalent name/value delete options | - |
| NAM-007 | Partial parity | Bulk modal `NamesView.tsx:868`, `1455+` | `updateNamedRange` / `deleteNamedRangeWithOptions` | Missing legacy second-pass transform and delimiter controls | P1 |
| TBL-001 | Full parity | `modern/TablesView.tsx:214`, table grid `240+` | `getTables` (`taskpane.ts:3145`) | Equivalent listing/sort/filter/select | - |
| TBL-002 | Full parity | Inline edit/save `TablesView.tsx:286-300` | `updateTableName` (`taskpane.ts:3377`) | Equivalent outcome | - |
| TBL-003 | Partial parity | Bulk edit modal `TablesView.tsx:315+` | `updateTableName` | Legacy collision-safe rename fallback is not replicated | P1 |
| FML-001 | Full parity | Modern action bar + editor actions (`ModernShell.tsx:719-728`, `FormulaMonacoView.tsx:923-1001`, `1799-1804`, `1881-1887`) | `getActiveCellFormulaState`, `applyFormulaToActiveCell`, `getCurrentSelectionAddress` | Equivalent pull/apply/beautify/insert flows | - |
| FML-002 | Outcome-equivalent parity | Evaluate/Test section (`ModernShell.tsx:833+`) and live output (`FormulaMonacoView.tsx:1976+`) | `evaluateFormula` (`taskpane.ts:3452`) | Same outcome, split across shell + editor sub-tab | - |
| FML-003 | Outcome-equivalent parity | Function metadata/wizard/save (`FormulaMonacoView.tsx:1709+`, `1476+`) | `saveNamedFunction` (`taskpane.ts:3187`) | Modern uses explicit save API vs legacy in-editor lambda build | - |
| FML-004 | Outcome-equivalent parity | Monaco suggestion pipeline (`FormulaMonacoView.tsx:1003+`, completion provider `1306+`) | `getNamedRanges`, `getTables` | Different UX, same suggestion intent | - |
| FML-005 | Missing | No modern go-to-definition/context menu equivalent | n/a | Legacy editor context actions not surfaced in modern UI | P2 |
| FML-006 | Full parity | Popout open and dialog RPC (`FormulaMonacoView.tsx:1066+`, `1506+`) + signal consume (`App.tsx:136+`) | `openFormulaEditorPopout` (`taskpane.ts:3910`) | Ribbon and popout routing preserved | - |
| FML-007 | Partial parity | Modern supports auto-capture toggle, but not legacy full settings panel (`FormulaMonacoView.tsx:1659+`) | n/a | Missing palette customization and several legacy tuning toggles | P2 |
| FMT-001 | Missing | No cell-style preset UI in `modern/FormatView.tsx` | `applyCellStylePreset` is legacy-only consumer | Legacy style preset workflow removed | P1 |
| FMT-002 | Partial parity | Shape creation in modern builder (`FormatView.tsx:1009`, `1467`) | `createShapeBuilderShape` (`taskpane.ts:2569`) | No active-cell anchored insertion equivalent to `addShapeOnActiveCell` | P1 |
| FMT-003 | Missing | No rename control in modern editor tabs (`FormatView.tsx:1575+`) | `renameShape` unused by modern | Legacy rename workflow absent | P1 |
| FMT-004 | Partial parity | Modern shape editor tabs (`FormatView.tsx:1575+`, `1772+`, `2017+`) | `updateShapeBuilderShape`, `formatShapeBuilderTextSelection` | Broad coverage, but lacks some legacy fields (e.g., rotation/lock-aspect controls) | P1 |
| FMT-005 | Partial parity | Position fields exist (`FormatView.tsx:2027-2030`) | `updateShapeBuilderShape` | Missing explicit nudge, align-to-selection, z-order, and move-to-selection parity | P0 |
| FMT-006 | Outcome-equivalent parity | Shape type selector `FormatView.tsx:1577-1580` | `updateShapeBuilderShape` (`shapeType`) | Different API path, same geometry-change outcome | - |
| FMT-007 | Missing | No named-range-to-text binding workflow in modern | `getNamedRangeValueText` unused by modern | Binding parity not implemented | P0 |
| FMT-008 | Missing | No formula-to-anchor binding action in modern | `applyFormulaToShapeAnchorCell` unused by modern | Anchor formula + text sync workflow missing | P0 |
| FMT-009 | Missing | No move-to-selection transfer workflow in modern | `moveShapeToSelection` unused by modern | Legacy anchor/value transfer behavior absent | P0 |
| FMT-010 | Full parity | Format popout button (`FormatView.tsx:1401`) | `openFormatEditorPopout` (`taskpane.ts:3880`) | Equivalent popout access | - |
| UTL-001 | Full parity | Utilities in Format + Sandbox (`FormatView.tsx:1222+`, `SandboxDebugView.tsx:53+`) | shared utility APIs (`toggleGridlines`, `refreshPivotTables`, `applyTableStyle`, `addRectangleShape`, `insertText`) | Command outcomes preserved across relocated UI | - |
| SIG-001 | Full parity | Command set in `commands.ts`; navigation polling in `App.tsx:96+` | `wbm.navigation.request` contract | Routing parity retained | - |
| SIG-002 | Full parity | Formula signal dispatch in `commands.ts`; consume/invoke in `App.tsx:136+` | `wbm.openFormulaEditor.request` contract | Ribbon formula actions still execute | - |

---

## 3) Service-Layer Parity Diff (`taskpane.ts`)

### Legacy-used APIs with no modern consumer

| API | Classification | Notes | Risk |
| --- | --- | --- | --- |
| `applyCellStylePreset` | True missing wiring in modern | Legacy style cards call it; modern has no equivalent control | Medium |
| `addShapeOnActiveCell` | Intentional replacement with gap | Replaced by `createShapeBuilderShape`, but active-cell anchoring parity is incomplete | High |
| `renameShape` | True missing wiring in modern | No modern rename action | Medium |
| `updateShapeFormatting` + `getShapeEditorRecord` | Intentional replacement | Replaced by shape-builder list/update APIs | Medium |
| `updateShapePosition` | Partial replacement | Modern can set left/top via `updateShapeBuilderShape`, but no dedicated workflow parity | Medium |
| `nudgeShape` | True missing wiring | No modern nudge controls | High |
| `alignShapeToSelection` | True missing wiring | No modern align-to-selection controls | High |
| `setShapeZOrder` | True missing wiring | No modern z-order controls | High |
| `moveShapeToSelection` | True missing wiring | Move-and-transfer behavior not implemented | High |
| `applyFormulaToShapeAnchorCell` | True missing wiring | Formula anchor binding missing | High |
| `getNamedRangeValueText` | True missing wiring | Named-value text binding missing | High |

### Modern-only APIs vs legacy

| API group | Classification | Notes |
| --- | --- | --- |
| `listShapeBuilderShapes` / `createShapeBuilderShape` / `updateShapeBuilderShape` / `formatShapeBuilderTextSelection` / `duplicateShapeBuilderShape` / `deleteShapeBuilderShape` | Intentional replacement/expansion | New shape-builder model replaces legacy shape-editor service surface |
| `listNavigationDestinations` / `activateNavigationDestination` / `applyNavigationTemplate` | Intentional expansion | Adds modern link destination and nav-template workflows |
| `listSheetFormats` / `captureSheetFormat` / `recaptureSheetFormatFromSelection` / `applySheetFormat*` / `deleteSheetFormat` | Intentional expansion | New format-template feature (no legacy equivalent) |
| `freezeTopRow` / `freezeFirstColumn` / `unfreezePanes` / `setSelectionColumnWidth` / `setSelectionRowHeight` / `openNewWorkbookWindow` / `arrangeWorkbookWindows` / `hideActiveSheet*` | Intentional expansion | Modern layout utility breadth exceeds legacy |
| `listWorkbookQueries` / `refreshWorkbookQueries` | Intentional expansion | New query-specific data workspace |
| `saveNamedFunction` | Intentional replacement | Explicit named-function save flow in modern formula UX |

---

## 4) Prioritized Remediation Backlog

| Ticket ID | Priority | Capability IDs | Behavior to add/fix | Target modern surface | Acceptance criteria | Dependencies / risk notes |
| --- | --- | --- | --- | --- | --- | --- |
| PAR-001 | P0 | NAM-002 | Fix named-range/cell hyperlink navigation reliability in host runtime | `modern/NamesView.tsx`, `taskpane.ts` nav helpers | Clicking a range row always activates target sheet+cell on supported hosts; failures surface actionable status | Host address parsing edge cases and scoped-name resolution |
| PAR-002 | P0 | FMT-007, FMT-008 | Reintroduce shape binding workflows (named-range text and formula-anchor sync) | `modern/FormatView.tsx` + shape services | User can bind selected shape text to named value and apply formula to anchor cell with reflected shape text | Requires either legacy APIs or new shape-builder binding APIs |
| PAR-003 | P0 | FMT-005, FMT-009 | Add arrange and move-to-selection controls (nudge/align/z-order/move) | `modern/FormatView.tsx` | Modern editor exposes these actions and workbook outcomes match legacy | Needs service wiring to legacy APIs or equivalent replacements |
| PAR-004 | P1 | NAM-003 | Add Names-grid shape-row navigation into modern shape editor context | `modern/NamesView.tsx` + `ModernShell.tsx` routing | Shape row is actionable and opens corresponding shape in Format editor | Depends on shape selection targeting contract |
| PAR-005 | P1 | NAM-007 | Restore legacy-grade bulk naming transforms (second-pass + delimiter options) | `modern/NamesView.tsx` | Bulk preview and results match legacy transformation capability set | UI complexity; ensure deterministic collision handling |
| PAR-006 | P1 | TBL-003 | Add collision-safe bulk rename behavior | `modern/TablesView.tsx` | Bulk rename succeeds without hard-fail on collisions; deterministic suffix fallback applied | Requires preflight uniqueness map |
| PAR-007 | P1 | FMT-001 | Reintroduce cell style preset workflow | `modern/FormatView.tsx` | Modern UI can apply Input/Parameter/Header/Subheader presets to selected range | Reuse `applyCellStylePreset` |
| PAR-008 | P1 | FMT-002 | Add explicit insert-on-active-cell option for shape creation | `modern/FormatView.tsx` | User can create shape anchored to selected cell with expected defaults | Could call `addShapeOnActiveCell` or implement equivalent in shape-builder API |
| PAR-009 | P1 | FMT-003 | Add shape rename action in modern editor | `modern/FormatView.tsx` | Shape rename updates workbook and refreshes list without losing selection | Service reuse: `renameShape` or `updateShapeBuilderShape` extension |
| PAR-010 | P2 | FML-007 | Add advanced editor settings parity (line-number toggle, optional palette controls) | `modern/FormulaMonacoView.tsx` | Settings persist and affect editor behavior as expected | Avoid over-coupling to legacy custom colorization model |
| PAR-011 | P2 | FML-005 | Add modern equivalents for high-value context actions (go-to-definition first) | `modern/FormulaMonacoView.tsx` | User can jump to named-range/table reference definitions from editor tokens | Requires token parse + navigation integration |

---

## 5) Uncertain/Host-Ambiguous Items

Branch: `parity/legacy-modern-uncertain`

### Strategy
- Umbrella branch created: `parity/legacy-modern-uncertain`
- Initial branch scope: `Uncertain` + `Partial` items with host-only ambiguity first (NAM-002 and shape host-event/navigation edge cases)
- Traceability rule: every branch PR change must reference one or more capability IDs from this document

### Checklist
- [ ] U-001 (`NAM-002`): Named range hyperlink navigation consistently selects target range across hosts (Desktop/Web, channel variants).
- [ ] U-002 (`NAM-002`): Worksheet-scoped name navigation resolves correctly when workbook and local names collide.
- [ ] U-003 (`FMT-005`): Shape activation event handlers in modern format view remain stable after shape create/duplicate/delete cycles.
- [ ] U-004 (`FMT-005`, `FMT-009`): Shape destination activation does not break when sheet names contain spaces/quotes.
- [ ] U-005 (`NAM-007`, `TBL-003`): Bulk operations produce deterministic outcomes under rename collisions and protected-sheet scenarios.
- [ ] U-006 (`FML-006`): Formula popout dialog RPC remains reliable after repeated open/close cycles.

### Validation Notes

| Item | Build/Host | Result | Evidence (screenshot/log) | Follow-up PR |
| --- | --- | --- | --- | --- |
| U-001 |  |  |  |  |
| U-002 |  |  |  |  |
| U-003 |  |  |  |  |
| U-004 |  |  |  |  |
| U-005 |  |  |  |  |
| U-006 |  |  |  |  |

---

## 6) Manual Host Parity Test Execution

Use this section to verify Legacy vs Modern outcomes hands-on. Compare using the same workbook where possible. Record `Pass` only when workbook state matches expected outcome. Capture host/channel/build in notes when behavior differs.

### Run Metadata
- Tester:
- Date:
- Excel host/version:
- Workbook used:

### Names

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| N-01 | NAM-001 | Workbook has workbook + worksheet scoped names | Open Modern > Structure > Named Ranges, click Refresh, apply sort/filter | Name inventory matches workbook Name Manager entries |  |  |
| N-02 | NAM-002 | At least one range name points to visible sheet/cell | Click linked range name in grid | Target sheet activates and target cell/range is selected |  |  |
| N-03 | Modern watchlist | Select a valid range and open create dialog | Click `+ Create Name`, use selection, create as `Range`, then refresh | New workbook name appears in Name Manager and maps to selected address |  |  |
| N-04 | NAM-004 | Existing named range record | Edit name/address and save | Updated name/definition is visible in Name Manager and grid after refresh |  |  |
| N-05 | NAM-005 | Existing range name | Open Move, choose new address, submit | Named range points to new address |  |  |
| N-06 | NAM-006 | Existing range name with values | Delete with `Delete name` only, then with `Delete values` | Name and value deletion behavior matches selected options |  |  |
| N-07 | NAM-007 | 3+ selected names | Run Bulk Prefix/Suffix/Replace and verify preview + result | Names update according to preview |  |  |
| N-08 | NAM-003 | Names grid includes shape rows | Try navigating from shape row | Shape editor opens with selected shape context (legacy parity target) |  |  |

### Tables

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| T-01 | TBL-001 | Workbook has at least 2 tables | Open Modern > Structure > Tables, refresh/sort/filter | Table list and metadata are accurate |  |  |
| T-02 | TBL-002 | Table exists | Inline edit table name and save | Table renamed in workbook and grid reload reflects new name |  |  |
| T-03 | TBL-003 | Multiple selected tables with potential rename collisions | Run bulk rename | All selected tables renamed with deterministic collision handling |  |  |

### Formulas

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| F-01 | FML-001, SIG-002 | Active cell contains formula | Trigger ribbon `Open and Pull Formula Editor` | Formula editor opens and loads active formula |  |  |
| F-02 | FML-001 | Active-cell formula available | Use Pull, Beautify, Insert Selection, Apply actions | Formula is transformed/inserted/applied correctly to active cell |  |  |
| F-03 | FML-002 | Formula text present | Open Evaluate and Test section, run test | Output grid renders expected calculation result |  |  |
| F-04 | FML-002 | Editor formula changes over time | Use Live Output sub-tab and edit formula | Live output updates and reports errors when applicable |  |  |
| F-05 | FML-003 | Valid LAMBDA formula and function name | Save Named Function and verify in Name Manager | Named function exists and can be invoked in worksheet |  |  |
| F-06 | FML-006 | Task pane mode modern | Open popout from formula editor and continue editing | Popout loads current state and remains interactive |  |  |
| F-07 | FML-007 | Formula editor open | Check line numbers, capture behavior, settings controls | Required settings parity behavior is available and functional |  |  |
| F-08 | FML-005 | Formula references named range/table | Attempt go-to-definition equivalent action | Definition jump capability works (if implemented) |  |  |

### Format / Layout

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| L-01 | FMT-010, UTL-001 | Normal worksheet | Use Gridlines, Freeze Top Row, Freeze First Column, Unfreeze | Worksheet view/freeze states update correctly |  |  |
| L-02 | FMT-010 | Multiple windows allowed | Open new window and apply each arrange mode | Workbook windows are arranged per selected layout |  |  |
| L-03 | FMT-002 | Active cell selected | Add shape and inspect initial location | Shape insertion behavior matches expected anchor rule |  |  |
| L-04 | FMT-004 | Existing shape selected | Edit text, fill, outline, font, size, position and save/blur | Shape reflects edits immediately and persists after refresh |  |  |
| L-05 | FMT-006 | Geometric shape selected | Change shape type (e.g., rectangle -> chevron) | Shape geometry updates without losing key metadata |  |  |
| L-06 | FMT-005 | Existing shape + target selection | Run arrange actions (nudge/align/z-order/move) | Shape position/order behavior matches legacy parity intent |  |  |
| L-07 | FMT-007 | Named range with visible value exists | Bind shape text to named range value | Shape text reflects named range value updates |  |  |
| L-08 | FMT-008 | Shape has anchor cell | Apply formula to shape anchor from modern UI | Anchor cell formula updates and shape text syncs |  |  |
| L-09 | FMT-001 | Select a range | Apply each cell style preset | Selection receives corresponding preset style |  |  |
| L-10 | FMT-010 | Format popout enabled | Open format popout and execute a shape or utility action | Action succeeds from popout context |  |  |

### Data

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| D-01 | (modern expansion) | Host supports query APIs | Open Queries and load list | Query metadata list loads with expected columns |  |  |
| D-02 | (modern expansion) | Workbook has refreshable queries | Trigger Refresh Queries | Refresh status/method/warnings are surfaced accurately |  |  |
| D-03 | UTL-001 | Workbook has pivot tables | Run Refresh Pivot Tables (Pivots view + utility buttons) | Pivot tables refresh without error |  |  |

### Tools / Settings / Help

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| O-01 | UTL-001 | Modern Sandbox available | Run insert text, accent shape, table style, gridlines, pivots | Each action updates workbook as indicated by status |  |  |
| O-02 | SIG-001 | UI preference previously saved | Click Reset Saved UI Preference in Settings | UI mode preference resets and default mode behavior returns |  |  |
| O-03 | SIG-001 | Help view open | Use Help jump links to move to sections | Navigation targets open correctly |  |  |

### Defect Log

| Defect ID | Scenario ID | Capability IDs | Severity | Actual behavior | Repro notes |
| --- | --- | --- | --- | --- | --- |
|  |  |  |  |  |  |

---

## 7) Additional Notes
- Legacy tabs `Format > Cells/View` and `Sandbox > Playground/Actions/Debug` are mostly placeholder surfaces in `LegacyApp.tsx`; parity decisions were based on workflows that produce workbook outcomes.
- Modern includes net-new domains (`Queries`, `Model Builder`, navigation templates, sheet format templates). These were treated as expansions, not parity blockers.
