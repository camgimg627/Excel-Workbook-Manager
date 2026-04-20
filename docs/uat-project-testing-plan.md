# Workbook Manager UAT Project Plan & Test Strategy

## 1) Objective and context
This plan is designed to move the add-in from bug-by-bug triage to **structured, area-by-area UAT execution** with clear entry/exit criteria, ownership, and defect governance.

Primary goals:
- Validate each major platform function end-to-end in a reproducible order.
- Reduce regression risk by introducing phase gates and smoke checks between areas.
- Create a common language for Product, QA, and business UAT stakeholders.

## 2) Platform functional inventory (what must be tested)
The modern shell exposes the primary user-facing areas and navigation structure for UAT execution order:
- Names
- Tables
- Formulas
- Queries
- Model Builder
- Pivots
- Format
- Sandbox Debug
- Settings
- Help

These map to explicit navigation targets/signals and command routing contracts, which means they can be validated both from the ribbon entry points and from in-app navigation.

## 3) Area-by-area UAT scope

### A. Entry points, navigation, and mode behavior
**Purpose:** Verify users can reliably open the intended section from ribbon commands and in-pane navigation.

**In scope:**
- Ribbon command-to-target routing for names, names-create, formulas actions, format, settings, help.
- Signal persistence/consumption behavior for `wbm.navigation.request` and `wbm.openFormulaEditor.request`.
- UI mode behavior: default modern mode, forced querystring mode, localStorage mode preference.
- Menu pathways to Settings/Help/Legacy UI.

**High-risk scenarios:**
- Task pane closed when command executes.
- Stale signal values causing misrouting.
- User stuck in wrong mode because saved preference is not reset.

---

### B. Names management
**Purpose:** Validate governance workflows for named ranges/shapes.

**In scope:**
- List names, filtering/sorting/selection behavior.
- Create named range (workbook/worksheet scope).
- Edit/update named range metadata.
- Move and delete workflows (including option variants).
- Bulk operations: prefix/suffix/replace/delete/case conversion.
- Address selection/jump from result row.

**High-risk scenarios:**
- Name conflicts and invalid names.
- Sheet-scoped vs workbook-scoped behavior.
- Names referring to formulas/functions vs ranges.

---

### C. Tables management
**Purpose:** Ensure table naming and metadata workflows are robust.

**In scope:**
- Table listing and navigation to table addresses.
- Inline rename and bulk rename.
- Column listing for selected tables.
- Create named ranges from table columns (conflict modes).

**High-risk scenarios:**
- Duplicate table name collisions.
- Broken references after rename.
- Mixed case and special-character normalization.

---

### D. Formula authoring and execution
**Purpose:** Validate formula editor behavior as a mini IDE workflow.

**In scope:**
- Open editor, pull active-cell formula, apply formula to active cell.
- Beautify and insert-selection shortcuts.
- Formula evaluation/test runner.
- Save named function flow.
- Formula popout behavior.
- Ribbon shortcuts (`open-and-pull`, `pull`, `apply`) end-to-end.

**High-risk scenarios:**
- Incorrect active cell context.
- Formula parse/localization edge cases.
- Signal race when running command sequences quickly.

---

### E. Queries and refresh operations
**Purpose:** Validate operational reliability for Power Query monitoring and refresh.

**In scope:**
- Query inventory metadata: load target, data model flag, refresh timestamps, row counts, errors.
- Refresh workbook queries and method reporting (`DataConnections.refreshAll` vs unsupported warning path).
- User messaging for unsupported API capabilities.

**High-risk scenarios:**
- Partial refresh with stale status.
- False-success responses.
- Workbooks with mixed query load destinations.

---

### F. Model Builder
**Purpose:** Validate parameter block scaffolding and naming output quality.

**In scope:**
- Anchor selection behavior.
- Parameter creation with manual values and formula values.
- Optional list validation artifacts and generated names.
- Apply summary accuracy (what was created and where).

**High-risk scenarios:**
- Invalid desired names and conflict resolution.
- Large parameter sets causing partial completion.
- Formula-based defaults producing unexpected errors.

---

### G. Pivots
**Purpose:** Validate pivot refresh as a quick maintenance workflow.

**In scope:**
- Refresh action on active worksheet.
- Success/failure feedback and resilience.

**High-risk scenarios:**
- Worksheet with no pivots.
- Refresh failures due to upstream query/table issues.

---

### H. Format and workbook UI tooling
**Purpose:** Validate formatting and layout operations used for workbook UX.

**In scope:**
- Gridline toggle, freeze/unfreeze operations.
- Selection formatting (bold/fill/presets/row-height/column-width).
- Table style application.
- Sheet format lifecycle: capture/list/apply/update/recapture/delete.
- Navigation template destination list, apply template, destination activation.
- Shape insertion, formatting, positioning, alignment, z-order, geometric-type updates.
- Format editor popout behavior where applicable.

**High-risk scenarios:**
- Actions applied to unintended selection.
- Named destination drift after sheet/table/name changes.
- Shape metadata corruption across duplicate/update flows.

---

### I. Sandbox Debug utilities
**Purpose:** Confirm internal utility actions remain safe and predictable.

**In scope:**
- Text insertion helper.
- Quick utility actions (pivot refresh, gridline toggle, sample shape/accent operations).

**High-risk scenarios:**
- Debug actions modifying user workbook unintentionally.
- Unexpected dependencies on dev-only assumptions.

---

### J. Settings and Help
**Purpose:** Validate adoption and supportability workflows.

**In scope:**
- Reset UI preference.
- Modern/legacy explanation and behavior alignment.
- Help navigation links and quick-start guidance correctness.

**High-risk scenarios:**
- Settings claims not matching actual behavior.
- Broken navigation links in help.

## 4) UAT execution waves (recommended sequence)

### Wave 0 — Stabilization smoke (2-3 days)
- Goal: confirm add-in launches, routing works, and no blocker-level crashes.
- Coverage: entry points, navigation, formulas open/apply happy path, names list, table list.
- Exit criteria: no Sev-1 blockers; smoke pass rate >= 90%.

### Wave 1 — Core authoring workflows (4-6 days)
- Coverage: Names, Tables, Formulas, Model Builder.
- Exit criteria: all P0/P1 defects fixed or approved workaround; retest complete.

### Wave 2 — Data operations (3-4 days)
- Coverage: Queries + Pivots with representative production-like workbooks.
- Exit criteria: refresh reliability confirmed in all required workbook archetypes.

### Wave 3 — Presentation/tooling (4-5 days)
- Coverage: Format, shape/navigation tooling, popouts, Sandbox safety checks.
- Exit criteria: no destructive formatting defects; deterministic results across runs.

### Wave 4 — Readiness and supportability (2-3 days)
- Coverage: Settings, Help, legacy fallback checks, full regression + command mapping verification.
- Exit criteria: UAT sign-off package complete.

## 5) Test design approach
- **Risk-based prioritization:** target high-defect and high-business-impact features first.
- **Workbook archetype matrix:** run each area against at least 3 workbook types:
  1. Clean/simple workbook
  2. Medium complexity (names + tables + pivots + queries)
  3. High complexity (heavy formulas, many sheets/shapes, model-builder artifacts)
- **Dual entry validation:** every major workflow tested from both ribbon command and in-app navigation where applicable.
- **Negative-path coverage:** invalid names, conflicts, unsupported APIs, missing selections, empty states.

## 6) Defect governance model for UAT
- **Severity definitions (suggested):**
  - **Sev-1:** Crash/data loss/workbook corruption/blocker to continue testing.
  - **Sev-2:** Core workflow broken with no practical workaround.
  - **Sev-3:** Workflow impaired but workaround exists.
  - **Sev-4:** Cosmetic, copy, non-blocking UX issues.
- **Triage cadence:** daily 30-minute triage during active waves.
- **Exit quality gates:**
  - 0 open Sev-1
  - Sev-2 count below agreed threshold (preferably 0)
  - 100% pass on critical-path scenarios
  - Documented acceptance of deferred defects

## 7) UAT deliverables and templates

### A. Master test suite structure
1. Navigation & commands
2. Names
3. Tables
4. Formulas
5. Queries
6. Model Builder
7. Pivots
8. Format + Shapes + Navigation Templates
9. Sandbox Debug
10. Settings & Help
11. Regression smoke

### B. Suggested test case schema
- Test ID
- Area
- Preconditions
- Steps
- Expected result
- Actual result
- Pass/Fail
- Defect ID (if failed)
- Build/version
- Tester/date

### C. Daily reporting snapshot
- Planned vs executed tests
- Pass/fail/block percentages
- New defects by severity
- Defects closed today
- Top 3 risk themes + next-day focus

## 8) Ownership model (RACI starter)
- **Product Owner:** approve scope, prioritize defect fixes, sign-off.
- **QA Lead:** maintain suite, run triage, quality gate recommendation.
- **UAT Business Testers:** execute business-critical scenarios and validate usability.
- **Engineering Lead:** defect assignment/escalation and fix readiness.
- **Release Manager:** wave scheduling and go/no-go package.

## 9) Practical kickoff checklist
- Confirm UAT workbook pack (simple/medium/complex) is finalized.
- Freeze a UAT build branch and version naming format.
- Align on severity definitions and gate thresholds.
- Baseline current known defects and map each to functional area.
- Start Wave 0 smoke before deep area testing.

## 10) Fast-start test backlog (first 15 scenarios)
1. Open add-in from ribbon to Names; verify route and state.
2. Trigger create-name command from ribbon and complete create flow.
3. Rename, move, and delete a named range; verify references remain valid.
4. Bulk prefix/suffix update names and verify output.
5. Open Tables, inline rename one table, verify formula references.
6. Bulk rename multiple tables with conflict handling.
7. Open formula editor via shortcut, pull active formula, apply edited formula.
8. Beautify formula and insert selection token into expression.
9. Evaluate a formula and verify returned value/error shape.
10. Save named function and consume it in worksheet formula.
11. List workbook queries and validate metadata displays non-empty/accurate fields.
12. Execute query refresh and validate result method + warnings.
13. Run model builder for 5+ parameters with mixed manual/formula values.
14. Apply formatting quick actions (freeze/gridline/style) and validate persistence.
15. Refresh pivots after query/table updates and verify output consistency.

---

## Appendix: Source-to-area mapping (for traceability)
- Command routing and signal writing: `src/commands/commands.ts`
- Navigation targets contract: `src/taskpane/navigation.ts`
- Modern shell/section surfaces: `src/taskpane/components/modern/ModernShell.tsx`
- Functional service operations (names, tables, formulas, queries, model builder, pivots, format, shapes, popouts): `src/taskpane/taskpane.ts`
- Product intent and section rationale: `docs/addin-overview.md`
- Legacy/modern parity baseline: `docs/ui-modernization-parity-matrix.md`
