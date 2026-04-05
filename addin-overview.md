# Workbook Manager Add-in Overview

## What this add-in is
Workbook Manager is an **Excel task pane add-in** that centralizes common workbook-authoring workflows (names, formulas, formatting, tables, query refresh, and helper utilities) into a single side panel and ribbon command set. It is designed to reduce repetitive Excel setup steps and make workbook structure easier to manage consistently.

At runtime, the add-in supports two UI modes:
- **Modern UI** (default)
- **Legacy UI** (fallback / emergency compatibility path)

The mode can be controlled by querystring (`?ui=modern` / `?ui=legacy`) and is persisted in local storage (`wbm.ui.mode`).

## How users enter the add-in
Users interact with Workbook Manager through a custom Excel ribbon tab called **Workbook Manager** with grouped commands:
- **Names**: open task pane, start name creation
- **Formulas**: open/pull/apply/beautify/insert-selection formula actions
- **Formatting**: open formatting tools
- **Settings & Help**: open those sections directly

Command handlers do not directly manipulate UI. Instead, they set signal keys (for navigation or formula actions), then show the task pane. The React app polls and consumes those signals to route to the right section.

## App architecture at a glance
The app has three layers:
1. **Ribbon command layer** (`src/commands/commands.ts`) for command-to-signal mapping.
2. **Task pane shell/router** (`src/taskpane/components/App.tsx`, `ModernShell.tsx`) for UI mode selection and navigation.
3. **Excel service layer** (`src/taskpane/taskpane.ts`) containing the Office.js operations (named ranges, queries, pivots, formatting, shapes, etc.).

This separation means most UI screens call stable service functions, while command handling stays focused on intent routing.

## Main sections (Modern UI)

| Section | Purpose | Key Workflows |
|---------|---------|---------------|
| **Names** | Treat workbook names as first-class assets | Inventory, create/edit/move/delete, bulk transform |
| **Tables** | Keep table naming clean and consistent | List, inline rename, bulk rename |
| **Formulas** | Mini IDE for Excel formula authoring/debugging | Monaco editor, pull/apply/beautify, evaluate/test, named functions, popout |
| **Queries** | Surface Power Query status and trigger refreshes | Metadata listing, refresh with method/warnings reporting |
| **Model Builder** | Scaffold parameterized worksheet inputs | Anchor selection, parameter rows, named cells, validation lists |
| **Pivots** | One-click pivot refresh | Refresh pivot tables on active worksheet |
| **Format** | Centralize formatting and layout operations | Style/grid/freeze/shape/table quick actions, popout editor |
| **Sandbox Debug** | Dev/troubleshooting utility actions | Pivot refresh, gridline toggle, sample shapes, debug text |
| **Settings** | Migration behavior and signal transparency | UI mode reset, signal key documentation |
| **Help** | Built-in onboarding | Quick-start guidance, formula shortcuts, jump links |

For detailed capability inventories, parity status, and gap analysis, see `ui-modernization-parity-matrix.md`.

## Navigation and signal model
The add-in uses a **signal contract** rather than direct cross-context calls:
- `wbm.navigation.request` for section navigation targets (names, tables, formulas, etc.)
- `wbm.openFormulaEditor.request` for formula action intents (open/pull/apply/beautify/insert)

Signals are written by ribbon commands to runtime/document storage; the task pane polls, consumes, and clears them. This approach keeps command execution robust even when task pane timing/lifecycle varies.

## Current product intent
Workbook Manager is evolving from a sample add-in into a practical workbook operations console:
- **Authoring speed:** faster setup/editing of names, formulas, and table structure.
- **Operational reliability:** explicit refresh/status surfaces for queries and pivots.
- **Consistency at scale:** bulk operations and standardized model-builder patterns.
- **Migration safety:** modern-by-default with legacy fallback during transition.
