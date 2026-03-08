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

## Intention of each main section (Modern UI)

### Names
**Purpose:** Treat workbook names as first-class assets, not hidden metadata.

**What it does:**
- Lists named ranges/objects with filtering, sorting, and selection.
- Supports create/edit/move/delete operations.
- Supports bulk operations (prefix/suffix/replace/delete/case transform).
- Can route from ribbon directly into “create name” flow.

**Why it exists:** Named ranges are core to scalable spreadsheet models; this section gives governance and batch-maintenance tools that Excel’s default UX makes cumbersome.

### Tables
**Purpose:** Keep table naming clean and consistent across a workbook.

**What it does:**
- Lists workbook tables.
- Supports inline rename and bulk rename workflows.
- Includes filtering/sorting patterns similar to Names.

**Why it exists:** Table names drive formulas, Power Query, and model readability. This area enforces naming hygiene at scale.

### Formulas
**Purpose:** Provide a richer authoring and debugging environment for formulas than the formula bar.

**What it does:**
- Monaco-based formula editor.
- Pull active-cell formula into editor.
- Apply editor formula back to active cell.
- Beautify and selection-insert helpers.
- Live test output/evaluation and save-as-named-function support.
- Supports direct ribbon and keyboard-shortcut driven actions.

**Why it exists:** Formula development is iterative; this section is intended as a focused “mini IDE” for Excel formulas.

### Queries
**Purpose:** Surface Power Query status/metadata and trigger refreshes from one place.

**What it does:**
- Lists query metadata (load target, model loading, refresh timestamps, row counts, errors).
- Refreshes workbook queries and reports method/warnings.
- Detects Excel API capability limitations and communicates support status.

**Why it exists:** Query operations are frequently operational (refresh + verify), so this section is designed for reliability/visibility.

### Model Builder
**Purpose:** Quickly scaffold parameterized worksheet inputs for model-driven workbooks.

**What it does:**
- Lets user pick an anchor and define parameter rows.
- Creates named cells and optional validation lists.
- Supports manual/default values or formula-based parameter values.
- Returns an apply summary of created artifacts.

**Why it exists:** Building parameter blocks repeatedly is error-prone. This section standardizes the pattern.

### Pivots
**Purpose:** Keep pivot outputs synchronized after source changes.

**What it does:**
- Dedicated action to refresh pivot tables on the active worksheet.
- Shows success/failure feedback.

**Why it exists:** Pivot refresh is a common post-update step and is exposed as a one-click maintenance action.

### Format
**Purpose:** Centralize formatting and layout operations used in workbook UI building.

**What it does:**
- Includes style/grid/freeze/table/pivot quick actions.
- Serves as the modern home for formatting workflows.
- Keeps a bridge to advanced legacy formatting/editor workflows where needed.

**Why it exists:** Workbook presentation and usability rely on consistent formatting; this section consolidates those tools.

### Sandbox Debug
**Purpose:** Provide safe utility/testing actions during development and troubleshooting.

**What it does:**
- Quick actions (pivot refresh, gridline toggle, sample shape insertion/accent fill).
- Debug text insertion helper.
- Table style utility.

**Why it exists:** Speeds up manual verification and diagnostics without navigating multiple Excel surfaces.

### Settings
**Purpose:** Operational controls for migration behavior and command signaling transparency.

**What it does:**
- Explains modern-vs-legacy rollout behavior.
- Offers “reset saved UI preference”.
- Documents signal keys used by commands (`wbm.navigation.request`, `wbm.openFormulaEditor.request`).

**Why it exists:** Helps support teams and users understand routing/mode behavior and recover from mismatched preferences.

### Help
**Purpose:** Built-in onboarding for first-time or occasional users.

**What it does:**
- Provides quick-start guidance by feature area.
- Documents formula shortcuts.
- Includes jump links to key sections.

**Why it exists:** Reduces training burden and makes intent of each section discoverable inside the add-in itself.

## Navigation and signal model (important concept)
The add-in uses a **signal contract** rather than direct cross-context calls:
- `wbm.navigation.request` for section navigation targets (names, tables, formulas, etc.)
- `wbm.openFormulaEditor.request` for formula action intents (open/pull/apply/beautify/insert)

Signals are written by ribbon commands to runtime/document storage; the task pane polls, consumes, and clears them. This approach keeps command execution robust even when task pane timing/lifecycle varies.

## Current product intent (big picture)
Workbook Manager is evolving from a sample add-in into a practical workbook operations console:
- **Authoring speed:** faster setup/editing of names, formulas, and table structure.
- **Operational reliability:** explicit refresh/status surfaces for queries and pivots.
- **Consistency at scale:** bulk operations and standardized model-builder patterns.
- **Migration safety:** modern-by-default with legacy fallback during transition.
