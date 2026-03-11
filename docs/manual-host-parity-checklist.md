# Manual Excel Host Parity Checklist (Legacy vs Modern)

## How To Use
- Baseline: current workspace branch `parity/legacy-modern-uncertain`
- Compare Legacy and Modern outcomes for each scenario using the same workbook where possible.
- Record `Pass` only when workbook state matches expected outcome.
- Capture host/channel/build in notes when behavior differs (Desktop vs Web, Insider vs Current Channel, etc.).

## Run Metadata
- Tester:
- Date:
- Excel host/version:
- Workbook used:

## Names

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

## Tables

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| T-01 | TBL-001 | Workbook has at least 2 tables | Open Modern > Structure > Tables, refresh/sort/filter | Table list and metadata are accurate |  |  |
| T-02 | TBL-002 | Table exists | Inline edit table name and save | Table renamed in workbook and grid reload reflects new name |  |  |
| T-03 | TBL-003 | Multiple selected tables with potential rename collisions | Run bulk rename | All selected tables renamed with deterministic collision handling |  |  |

## Formulas

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

## Format / Layout

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

## Data

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| D-01 | (modern expansion) | Host supports query APIs | Open Queries and load list | Query metadata list loads with expected columns |  |  |
| D-02 | (modern expansion) | Workbook has refreshable queries | Trigger Refresh Queries | Refresh status/method/warnings are surfaced accurately |  |  |
| D-03 | UTL-001 | Workbook has pivot tables | Run Refresh Pivot Tables (Pivots view + utility buttons) | Pivot tables refresh without error |  |  |

## Tools / Settings / Help

| Scenario ID | Capability IDs | Preconditions | Steps | Expected workbook outcome | Result (Pass/Fail) | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| O-01 | UTL-001 | Modern Sandbox available | Run insert text, accent shape, table style, gridlines, pivots | Each action updates workbook as indicated by status |  |  |
| O-02 | SIG-001 | UI preference previously saved | Click Reset Saved UI Preference in Settings | UI mode preference resets and default mode behavior returns |  |  |
| O-03 | SIG-001 | Help view open | Use Help jump links to move to sections | Navigation targets open correctly |  |  |

## Defect Log (fill during run)

| Defect ID | Scenario ID | Capability IDs | Severity | Actual behavior | Repro notes |
| --- | --- | --- | --- | --- | --- |
|  |  |  |  |  |  |
