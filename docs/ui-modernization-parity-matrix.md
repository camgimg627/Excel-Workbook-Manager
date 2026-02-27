# Workbook Manager UI Modernization Parity Matrix

## Rollout Controls
- Default mode: `modern`
- Fallback mode: `legacy` (emergency)
- Query override: `?ui=modern` or `?ui=legacy`
- Formula popout: `?popout=formula`
- Sticky user mode key: `localStorage["wbm.ui.mode"]`

## Feature Coverage
| Area | Legacy Mode | Modern Mode | Notes |
| --- | --- | --- | --- |
| Names list/filter/sort/select | ✅ | ✅ | Modern uses new card + grid shell |
| Create name | ✅ | ✅ | Modern routed via `names-create` navigation signal |
| Edit/move/delete named range | ✅ | ✅ | Same `taskpane.ts` service surface |
| Bulk named range operations | ✅ | ✅ | Modern supports prefix/suffix/replace/delete bulk |
| Tables list/filter/sort/select | ✅ | ✅ | Modern uses new card + grid shell |
| Inline table rename | ✅ | ✅ | Same `updateTableName` service |
| Bulk table rename | ✅ | ✅ | Modern bulk modal |
| Formula editor | ✅ | ✅ | Modern uses Monaco |
| Formula pull/apply/beautify/insert selection | ✅ | ✅ | Command-signal compatible |
| Formula test runner | ✅ | ✅ | Uses `evaluateFormula` |
| Formula popout | ✅ | ✅ | Modern popout when mode is `modern` |
| Format styles/grid/freeze/table/pivot quick actions | ✅ | ✅ | Native modern Format view |
| Advanced shape editor workflows | ✅ | ✅ (Legacy Bridge) | Modern Format includes embedded legacy advanced panel |
| Pivots section | ⚠️ (via other sections) | ✅ | Native modern Pivots view with refresh action |
| Settings section | ⚠️ | ✅ | Native modern Settings view |
| Help section | ⚠️ | ✅ | Native modern Help view with guidance + section links |
| Sandbox/debug quick actions | ✅ | ✅ | Modern has dedicated Sandbox Debug card |
| Ribbon formulas shortcuts | ✅ | ✅ | Existing shortcut IDs unchanged |

## Ribbon Command Mapping (Current)
- Group layout:
  - `Names`: `TaskpaneButton`, `CreateNameButton`
  - `Formulas`: formula command buttons
  - `Formatting`: `OpenFormattingButton`
  - `Settings & Help`: `OpenSettingsButton`, `OpenHelpButton`
- `TaskpaneButton` -> `openNamesCommand`
- `CreateNameButton` -> `createNameCommand`
- `OpenFormulaEditorButton` -> `openFormulaEditorCommand`
- `OpenAndPullFormulaEditorButton` -> `openAndPullFormulaEditorCommand`
- `PullFormulaButton` -> `pullFormulaCommand`
- `ApplyFormulaButton` -> `applyFormulaCommand`
- `BeautifyFormulaButton` -> `beautifyFormulaCommand`
- `InsertSelectionButton` -> `insertSelectionFormulaCommand`
- `OpenFormattingButton` -> `openFormattingCommand`
- `OpenSettingsButton` -> `openSettingsCommand`
- `OpenHelpButton` -> `openHelpCommand`

## Keyboard Shortcuts
- `Ctrl+Shift+Alt+E` -> `openAndPullFormulaEditorCommand`
- `Ctrl+Shift+Alt+P` -> `pullFormulaCommand`
- `Ctrl+Shift+Alt+A` -> `applyFormulaCommand`

## Navigation Signal Contract
- Key: `wbm.navigation.request`
- Targets:
  - `names`
  - `names-create`
  - `tables`
  - `formulas`
  - `pivots`
  - `format`
  - `sandbox-debug`
  - `settings`
  - `help`
