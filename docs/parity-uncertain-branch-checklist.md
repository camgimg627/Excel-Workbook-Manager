# Uncertain/Host-Ambiguous Parity Checklist

Branch: `parity/legacy-modern-uncertain`

Use this checklist for changes that are `Uncertain` or `Partial` due host/runtime ambiguity. Each item must be tied to a capability ID from `docs/ui-modernization-parity-matrix.md`.

## Checklist
- [ ] U-001 (`NAM-002`): Named range hyperlink navigation consistently selects target range across hosts (Desktop/Web, channel variants).
- [ ] U-002 (`NAM-002`): Worksheet-scoped name navigation resolves correctly when workbook and local names collide.
- [ ] U-003 (`FMT-005`): Shape activation event handlers in modern format view remain stable after shape create/duplicate/delete cycles.
- [ ] U-004 (`FMT-005`, `FMT-009`): Shape destination activation does not break when sheet names contain spaces/quotes.
- [ ] U-005 (`NAM-007`, `TBL-003`): Bulk operations produce deterministic outcomes under rename collisions and protected-sheet scenarios.
- [ ] U-006 (`FML-006`): Formula popout dialog RPC remains reliable after repeated open/close cycles.

## Notes Template
| Item | Build/Host | Result | Evidence (screenshot/log) | Follow-up PR |
| --- | --- | --- | --- | --- |
| U-001 |  |  |  |  |
| U-002 |  |  |  |  |
| U-003 |  |  |  |  |
| U-004 |  |  |  |  |
| U-005 |  |  |  |  |
| U-006 |  |  |  |  |
