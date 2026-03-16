import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import {
  TableColumnRecord,
  TableRecord,
  createNamedRangesFromTableColumnsV2,
  getTableColumns,
  getTables,
  listWorkbookNames,
  selectTableAddress,
  updateTableName,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";
import {
  CaseStyle,
  ColumnConfig,
  NamePreviewResult,
  buildColumnCreateSpecs,
  derivePreview,
  isValidExcelName,
} from "../../utils/nameTransforms";
import {
  TABLE_CONTEXT_SIGNAL_KEY,
  TABLE_CONTEXT_TABLE_NAME_KEY,
  TABLE_CONTEXT_SHEET_NAME_KEY,
} from "../../../commands/commands";

// ─── Local types ──────────────────────────────────────────────────────────────

type SortColumn = "name" | "address" | "sheet" | "scope";
type SortDirection = "asc" | "desc";
type BulkCaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";

interface TablesViewProps {
  onOpenLegacy: () => void;
}

interface BulkState {
  open: boolean;
  mode: "Prefix" | "Suffix" | "Replace";
  value: string;
  replaceWith: string;
  caseTransform: BulkCaseTransform;
}

/** State for the enhanced Create Named Ranges modal. */
interface RangesModalState {
  open: boolean;
  table: TableRecord | null;
  /** Full column list loaded from Excel. */
  columns: ColumnConfig[];
  /** Global prefix applied to all non-overridden column names. */
  globalPrefix: string;
  /** Global suffix applied to all non-overridden column names. */
  globalSuffix: string;
  /** Case style applied to all non-overridden column names. */
  caseStyle: CaseStyle;
  scopeType: "Workbook" | "Worksheet";
  /**
   * Existing named range names in the target scope (upper-cased).
   * Pre-loaded when the modal opens so the preview can flag duplicates.
   */
  existingNames: Set<string>;
}

// ─── Utility: bulk table name case transform ──────────────────────────────────

const applyBulkCase = (value: string, mode: BulkCaseTransform): string => {
  if (mode === "none") return value;
  const parts = value
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
  if (!parts.length) return value;
  if (mode === "camelCase") {
    return parts
      .map((part, i) =>
        i === 0
          ? `${part.charAt(0).toLowerCase()}${part.slice(1).toLowerCase()}`
          : `${part.charAt(0).toUpperCase()}${part.slice(1).toLowerCase()}`
      )
      .join("");
  }
  if (mode === "snake_case") return parts.map((p) => p.toLowerCase()).join("_");
  return parts.map((p) => p.toUpperCase()).join("_");
};

// ─── Styles ───────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: { display: "grid", gap: "24px" },
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  spacer: { flexGrow: 1 },
  tableWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    backgroundColor: "#fff",
    overflow: "auto",
  },
  table: { width: "100%", borderCollapse: "separate", borderSpacing: 0, minWidth: "840px", fontSize: "12px" },
  headCell: {
    position: "sticky",
    top: 0,
    zIndex: 1,
    backgroundColor: "#F3F4F6",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px 8px",
    whiteSpace: "nowrap",
  },
  sortBtn: { border: "none", background: "transparent", fontWeight: 600, cursor: "pointer", padding: 0 },
  row: {
    selectors: {
      "&:nth-child(even)": { backgroundColor: "#FCFCFD" },
      "&:hover": { backgroundColor: "#F3F8F4" },
      "&:hover .row-actions": { opacity: 1, pointerEvents: "auto" },
    },
  },
  selectedRow: { backgroundColor: "#EAF2FF" },
  cell: { padding: "10px 8px", borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, whiteSpace: "nowrap" },
  clickableCellBtn: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
    padding: 0,
    textDecorationLine: "underline",
    fontSize: "12px",
  },
  rowActions: { display: "flex", gap: "4px", opacity: 0, pointerEvents: "none", transition: "opacity 150ms ease" },

  // ── Modals ──
  modalBackdrop: {
    position: "fixed",
    inset: 0,
    backgroundColor: "rgba(15, 23, 42, 0.24)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 60,
    padding: "16px",
  },
  modal: {
    width: "min(760px, 100%)",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "12px",
    maxHeight: "90vh",
    overflow: "auto",
  },
  modalGrid: { display: "grid", gap: "12px", gridTemplateColumns: "repeat(2, minmax(0, 1fr))" },
  modalGrid3: { display: "grid", gap: "12px", gridTemplateColumns: "repeat(3, minmax(0, 1fr))" },
  full: { gridColumn: "1 / -1" },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: "8px", flexWrap: "wrap", alignItems: "center" },
  modalDivider: { borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`, marginTop: "4px" },

  // ── Column picker table inside the modal ──
  colTable: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: "12px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "6px",
    overflow: "hidden",
  },
  colTableHead: {
    backgroundColor: "#F3F4F6",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
  },
  colTh: {
    padding: "8px",
    textAlign: "left",
    fontWeight: 600,
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    whiteSpace: "nowrap",
  },
  colTd: {
    padding: "6px 8px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    verticalAlign: "middle",
  },
  colTableWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "6px",
    overflow: "auto",
    maxHeight: "260px",
  },
  previewValid: {
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "11px",
    color: "#166534",
    backgroundColor: "#dcfce7",
    borderRadius: "4px",
    padding: "2px 6px",
    display: "inline-block",
  },
  previewInvalid: {
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "11px",
    color: "#991b1b",
    backgroundColor: "#fee2e2",
    borderRadius: "4px",
    padding: "2px 6px",
    display: "inline-block",
  },
  previewDuplicate: {
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "11px",
    color: "#92400e",
    backgroundColor: "#fef3c7",
    borderRadius: "4px",
    padding: "2px 6px",
    display: "inline-block",
  },
  previewExcluded: {
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    fontStyle: "italic",
  },
  mutedCode: {
    fontFamily: "Consolas, 'Courier New', monospace",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "11px",
  },
  selectAllRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    paddingBottom: "4px",
  },
  countBadge: {
    marginRight: "auto",
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
  },
});

// ─── Component ────────────────────────────────────────────────────────────────

const TablesView: React.FC<TablesViewProps> = ({ onOpenLegacy }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();

  // Main table list
  const [rows, setRows] = useState<TableRecord[]>([]);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [loading, setLoading] = useState<boolean>(false);
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [search, setSearch] = useState<string>("");
  const [sortColumn, setSortColumn] = useState<SortColumn>("name");
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [inlineEdit, setInlineEdit] = useState<{ id: string; value: string } | null>(null);

  // Bulk rename modal
  const [bulkState, setBulkState] = useState<BulkState>({
    open: false,
    mode: "Prefix",
    value: "",
    replaceWith: "",
    caseTransform: "none",
  });

  // Create Named Ranges modal
  const emptyRangesModal: RangesModalState = {
    open: false,
    table: null,
    columns: [],
    globalPrefix: "",
    globalSuffix: "",
    caseStyle: "none",
    scopeType: "Worksheet",
    existingNames: new Set(),
  };
  const [rangesModal, setRangesModal] = useState<RangesModalState>(emptyRangesModal);

  // ── Load table list ────────────────────────────────────────────────────────

  const load = async () => {
    setLoading(true);
    try {
      const loaded = await getTables();
      setRows(loaded);
      setSelectedIds(new Set());
      setInlineEdit(null);
      setStatusType("success");
      setStatus(`Loaded ${loaded.length} table(s).`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Load failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);

  // ── Context menu signal polling ────────────────────────────────────────────
  //
  // When the user right-clicks inside a table and selects one of our context
  // menu items, commands.ts writes signals and calls showAsTaskpane(). The
  // task pane may already be open; polling here catches the signal on mount
  // and on each poll interval, then clears it so it fires only once.

  useEffect(() => {
    let cancelled = false;

    const consumeSignal = async () => {
      try {
        let action: string | null = null;
        let tableName: string | null = null;
        let sheetName: string | null = null;

        // Prefer OfficeRuntime.storage (shared runtime, persists across pane hide/show).
        if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
          action = await OfficeRuntime.storage.getItem(TABLE_CONTEXT_SIGNAL_KEY).catch(() => null);
          tableName = await OfficeRuntime.storage.getItem(TABLE_CONTEXT_TABLE_NAME_KEY).catch(() => null);
          sheetName = await OfficeRuntime.storage.getItem(TABLE_CONTEXT_SHEET_NAME_KEY).catch(() => null);

          if (action) {
            // Clear the signal immediately so it doesn't re-fire.
            await OfficeRuntime.storage.removeItem(TABLE_CONTEXT_SIGNAL_KEY).catch(() => undefined);
          }
        } else if (Office.context?.document?.settings) {
          action = Office.context.document.settings.get(TABLE_CONTEXT_SIGNAL_KEY) as string | null;
          tableName = Office.context.document.settings.get(TABLE_CONTEXT_TABLE_NAME_KEY) as string | null;
          sheetName = Office.context.document.settings.get(TABLE_CONTEXT_SHEET_NAME_KEY) as string | null;

          if (action) {
            Office.context.document.settings.remove(TABLE_CONTEXT_SIGNAL_KEY);
            await new Promise<void>((res) => Office.context.document.settings.saveAsync(() => res()));
          }
        }

        if (!action || !tableName || !sheetName || cancelled) return;

        const table = rows.find((r) => r.name === tableName && r.sheet === sheetName);

        if (action === "editTableName") {
          // Jump directly to inline edit for the matching table row.
          if (table) {
            setInlineEdit({ id: table.id, value: table.name });
            setSelectedIds(new Set([table.id]));
          }
        } else if (action === "createNamedRanges") {
          // Open the Create Named Ranges modal for this table.
          const targetTable = table ?? { id: `${sheetName}::${tableName}`, name: tableName, sheet: sheetName, address: "", scope: sheetName };
          await openRangesModalForTable(targetTable);
        }
      } catch {
        // Polling errors are non-fatal.
      }
    };

    // Check once on mount (rows may not be loaded yet — check again after load).
    void consumeSignal();

    // Poll every 600ms to catch signals that arrive while the pane is visible.
    const interval = setInterval(() => void consumeSignal(), 600);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [rows]); // Re-run when rows change so the table lookup is up to date.

  // ── Computed values ────────────────────────────────────────────────────────

  const visibleRows = useMemo(() => {
    const token = search.trim().toLowerCase();
    return rows
      .filter((row) =>
        token ? `${row.name} ${row.address} ${row.sheet} ${row.scope}`.toLowerCase().includes(token) : true
      )
      .sort((a, b) => {
        const compare = a[sortColumn].localeCompare(b[sortColumn], undefined, { sensitivity: "base" });
        return sortDirection === "asc" ? compare : -compare;
      });
  }, [rows, search, sortColumn, sortDirection]);

  const allSelected = visibleRows.length > 0 && visibleRows.every((row) => selectedIds.has(row.id));

  const bulkPreview = useMemo(() => {
    const selected = rows.filter((row) => selectedIds.has(row.id));
    return selected.map((row) => {
      let nextName = row.name;
      if (bulkState.mode === "Prefix") nextName = `${bulkState.value}${row.name}`;
      if (bulkState.mode === "Suffix") nextName = `${row.name}${bulkState.value}`;
      if (bulkState.mode === "Replace") nextName = row.name.replace(bulkState.value, bulkState.replaceWith);
      nextName = applyBulkCase(nextName, bulkState.caseTransform);
      return { row, oldName: row.name, newName: nextName };
    });
  }, [rows, selectedIds, bulkState]);

  // ── Named Ranges modal — preview computation ───────────────────────────────
  //
  // For each included column, derive a preview result (finalName, isValid,
  // isDuplicate).  This is pure derived state — computed in render, never
  // stored in state — so it updates instantly as the user types.

  const columnPreviews = useMemo((): NamePreviewResult[] => {
    if (!rangesModal.open) return [];

    const seenInBatch = new Set<string>();
    return rangesModal.columns.map((col) => {
      const result = derivePreview(
        col,
        rangesModal.globalPrefix,
        rangesModal.globalSuffix,
        rangesModal.caseStyle,
        rangesModal.existingNames,
        seenInBatch
      );
      if (col.included && result.finalName && result.isValid && !result.isDuplicate) {
        seenInBatch.add(result.finalName.toUpperCase());
      }
      return result;
    });
  }, [rangesModal]);

  const includedCount = rangesModal.columns.filter((c) => c.included).length;
  const hasErrors = columnPreviews.some(
    (p, i) => rangesModal.columns[i].included && (!p.isValid || p.isDuplicate)
  );

  // ── Inline rename ──────────────────────────────────────────────────────────

  const saveInline = async (row: TableRecord) => {
    if (!inlineEdit) return;
    setSubmitting(true);
    try {
      await updateTableName(row.sheet, row.name, inlineEdit.value.trim());
      setInlineEdit(null);
      await load();
      setStatusType("success");
      setStatus("Table renamed.");
    } catch (error) {
      setStatusType("error");
      setStatus(`Rename failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  // ── Open Named Ranges modal ────────────────────────────────────────────────

  /**
   * Load columns + existing names, then open the modal.
   * Used by both the toolbar button and the context menu signal handler.
   */
  const openRangesModalForTable = async (table: TableRecord) => {
    setSubmitting(true);
    try {
      const [rawColumns, allNames] = await Promise.all([
        getTableColumns(table.sheet, table.name),
        listWorkbookNames(),
      ]);

      const existingNames = new Set(allNames.map((n: { name: string }) => n.name.toUpperCase()));

      const columns: ColumnConfig[] = rawColumns.map((col: TableColumnRecord) => ({
        id: col.id,
        originalName: col.name,
        address: col.address,
        included: true,
        perColumnOverride: "",
      }));

      setRangesModal({
        open: true,
        table,
        columns,
        globalPrefix: "",
        globalSuffix: "",
        caseStyle: "none",
        scopeType: "Worksheet",
        existingNames,
      });
    } catch (error) {
      setStatusType("error");
      setStatus(`Unable to load table columns: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const openCreateRangesModal = async () => {
    const table = rows.find((row) => selectedIds.has(row.id));
    if (!table) {
      setStatusType("error");
      setStatus("Select one table first.");
      return;
    }
    await openRangesModalForTable(table);
  };

  // ── Submit Named Ranges creation ───────────────────────────────────────────

  const applyCreateRangesFromTable = async () => {
    if (!rangesModal.table) return;

    const specs = buildColumnCreateSpecs(
      rangesModal.columns,
      rangesModal.globalPrefix,
      rangesModal.globalSuffix,
      rangesModal.caseStyle,
      rangesModal.existingNames
    );

    if (specs.length === 0) {
      setStatusType("error");
      setStatus("No valid columns selected. Resolve any errors in the preview first.");
      return;
    }

    setSubmitting(true);
    try {
      const result = await createNamedRangesFromTableColumnsV2({
        sheetName: rangesModal.table.sheet,
        tableName: rangesModal.table.name,
        columns: specs,
        scopeType: rangesModal.scopeType,
      });
      setRangesModal(emptyRangesModal);
      setStatusType("success");
      setStatus(
        `Created ${result.created.length} named range(s)` +
          (result.skipped.length ? `, skipped ${result.skipped.length}.` : ".")
      );
    } catch (error) {
      setStatusType("error");
      setStatus(`Create from table failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  // ── Bulk rename ────────────────────────────────────────────────────────────

  const applyBulk = async () => {
    if (bulkPreview.length === 0) return;
    setSubmitting(true);
    try {
      for (const item of bulkPreview) {
        if (item.oldName === item.newName) continue;
        await updateTableName(item.row.sheet, item.oldName, item.newName);
      }
      setBulkState((prev) => ({ ...prev, open: false }));
      setSelectedIds(new Set());
      await load();
      setStatusType("success");
      setStatus(`Bulk update completed for ${bulkPreview.length} table(s).`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Bulk update failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  // ── Render ─────────────────────────────────────────────────────────────────

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
      {/* Header */}
      <div>
        <Text className={shared.sectionTitle}>Tables</Text>
        <Text className={shared.sectionSubtitle}>Selection-driven table actions and quick range generation.</Text>
      </div>

      {/* Toolbar */}
      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button onClick={() => void openCreateRangesModal()} disabled={selectedIds.size !== 1 || submitting}>
            Create Named Ranges from Table
          </Button>
          <Button onClick={() => setBulkState((prev) => ({ ...prev, open: true }))} disabled={selectedIds.size === 0}>
            Bulk Edit ({selectedIds.size})
          </Button>
          <Button onClick={() => void load()} disabled={loading}>Refresh</Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          <div className={styles.spacer} />
          <Input placeholder="Search tables..." value={search} onChange={(_, data) => setSearch(data.value)} />
        </div>
      </div>

      {/* Table list */}
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.headCell}>
                <input
                  type="checkbox"
                  checked={allSelected}
                  onChange={() =>
                    setSelectedIds((prev) => {
                      const next = new Set(prev);
                      if (allSelected) visibleRows.forEach((row) => next.delete(row.id));
                      else visibleRows.forEach((row) => next.add(row.id));
                      return next;
                    })
                  }
                />
              </th>
              <th className={styles.headCell}>Actions</th>
              {(["name", "address", "sheet", "scope"] as const).map((column) => (
                <th key={column} className={styles.headCell}>
                  <button
                    type="button"
                    className={styles.sortBtn}
                    onClick={() => {
                      if (sortColumn === column) setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
                      else {
                        setSortColumn(column);
                        setSortDirection("asc");
                      }
                    }}
                  >
                    {column[0].toUpperCase() + column.slice(1)}
                    {sortColumn === column ? (sortDirection === "asc" ? " ▲" : " ▼") : ""}
                  </button>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={row.id} className={`${styles.row} ${selectedIds.has(row.id) ? styles.selectedRow : ""}`}>
                <td className={styles.cell}>
                  <input
                    type="checkbox"
                    checked={selectedIds.has(row.id)}
                    onChange={() =>
                      setSelectedIds((prev) => {
                        const next = new Set(prev);
                        if (next.has(row.id)) next.delete(row.id);
                        else next.add(row.id);
                        return next;
                      })
                    }
                  />
                </td>
                <td className={styles.cell}>
                  <div className={`${styles.rowActions} row-actions`}>
                    {inlineEdit?.id === row.id ? (
                      <>
                        <Button size="small" disabled={submitting} onClick={() => void saveInline(row)}>Save</Button>
                        <Button size="small" onClick={() => setInlineEdit(null)}>Cancel</Button>
                      </>
                    ) : (
                      <Button size="small" onClick={() => setInlineEdit({ id: row.id, value: row.name })}>Edit</Button>
                    )}
                  </div>
                </td>
                <td className={styles.cell}>
                  {inlineEdit?.id === row.id ? (
                    <Input
                      value={inlineEdit.value}
                      onChange={(_, data) => setInlineEdit({ id: row.id, value: data.value })}
                    />
                  ) : (
                    <button
                      type="button"
                      className={styles.clickableCellBtn}
                      onClick={() => void selectTableAddress(row.address, row.sheet)}
                    >
                      {row.name}
                    </button>
                  )}
                </td>
                <td className={styles.cell}>
                  <span className={styles.mutedCode}>{row.address}</span>
                </td>
                <td className={styles.cell}>{row.sheet}</td>
                <td className={styles.cell}>{row.scope}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <Text className={shared.mutedText}>{`Showing ${visibleRows.length} of ${rows.length}`}</Text>
      {status ? <Text className={statusClass}>{status}</Text> : null}

      {/* ================================================================== */}
      {/* CREATE NAMED RANGES MODAL                                           */}
      {/* ================================================================== */}
      {rangesModal.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create Named Ranges from Table</Text>
            <Text className={shared.mutedText}>
              Table: <strong>{rangesModal.table?.name}</strong> — Sheet: {rangesModal.table?.sheet}
            </Text>

            {/* Global settings row */}
            <div className={styles.modalGrid3}>
              <div>
                <Text className={shared.mutedText}>Global Prefix</Text>
                <Input
                  placeholder="e.g. tbl_"
                  value={rangesModal.globalPrefix}
                  onChange={(_, data) => setRangesModal((prev) => ({ ...prev, globalPrefix: data.value }))}
                />
              </div>
              <div>
                <Text className={shared.mutedText}>Global Suffix</Text>
                <Input
                  placeholder="e.g. _range"
                  value={rangesModal.globalSuffix}
                  onChange={(_, data) => setRangesModal((prev) => ({ ...prev, globalSuffix: data.value }))}
                />
              </div>
              <div>
                <Text className={shared.mutedText}>Case Style</Text>
                <Select
                  value={rangesModal.caseStyle}
                  onChange={(_, data) =>
                    setRangesModal((prev) => ({ ...prev, caseStyle: data.value as CaseStyle }))
                  }
                >
                  <option value="none">None (keep as-is)</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAM_SNAKE_CASE">SCREAM_SNAKE_CASE</option>
                </Select>
              </div>
            </div>

            <div>
              <Text className={shared.mutedText}>Scope</Text>
              <Select
                value={rangesModal.scopeType}
                onChange={(_, data) =>
                  setRangesModal((prev) => ({ ...prev, scopeType: data.value as "Workbook" | "Worksheet" }))
                }
              >
                <option value="Worksheet">Worksheet</option>
                <option value="Workbook">Workbook</option>
              </Select>
            </div>

            <div className={styles.modalDivider} />

            {/* Column picker table */}
            <div>
              {/* Select all / none controls */}
              <div className={styles.selectAllRow}>
                <Button
                  size="small"
                  onClick={() =>
                    setRangesModal((prev) => ({
                      ...prev,
                      columns: prev.columns.map((c) => ({ ...c, included: true })),
                    }))
                  }
                >
                  Select All
                </Button>
                <Button
                  size="small"
                  onClick={() =>
                    setRangesModal((prev) => ({
                      ...prev,
                      columns: prev.columns.map((c) => ({ ...c, included: false })),
                    }))
                  }
                >
                  Deselect All
                </Button>
                <Text className={styles.countBadge}>
                  {includedCount} of {rangesModal.columns.length} columns selected
                </Text>
              </div>

              <div className={styles.colTableWrap}>
                <table className={styles.colTable}>
                  <thead>
                    <tr className={styles.colTableHead}>
                      <th className={styles.colTh} style={{ width: "32px" }}></th>
                      <th className={styles.colTh}>Column</th>
                      <th className={styles.colTh}>Address</th>
                      <th className={styles.colTh} style={{ width: "180px" }}>
                        Custom Name Override
                        <span style={{ fontWeight: 400, marginLeft: "4px" }}>(optional)</span>
                      </th>
                      <th className={styles.colTh}>Preview</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rangesModal.columns.map((col, idx) => {
                      const preview = columnPreviews[idx];
                      let previewClass = styles.previewExcluded;
                      let previewLabel = "—";
                      if (col.included && preview) {
                        if (!preview.isValid) {
                          previewClass = styles.previewInvalid;
                          previewLabel = `${preview.finalName} ✗ invalid`;
                        } else if (preview.isDuplicate) {
                          previewClass = styles.previewDuplicate;
                          previewLabel = `${preview.finalName} ⚠ duplicate`;
                        } else {
                          previewClass = styles.previewValid;
                          previewLabel = preview.finalName;
                        }
                      }

                      return (
                        <tr key={col.id}>
                          <td className={styles.colTd}>
                            <input
                              type="checkbox"
                              checked={col.included}
                              onChange={() =>
                                setRangesModal((prev) => ({
                                  ...prev,
                                  columns: prev.columns.map((c) =>
                                    c.id === col.id ? { ...c, included: !c.included } : c
                                  ),
                                }))
                              }
                            />
                          </td>
                          <td className={styles.colTd}>{col.originalName}</td>
                          <td className={styles.colTd}>
                            <span className={styles.mutedCode}>{col.address}</span>
                          </td>
                          <td className={styles.colTd}>
                            <Input
                              size="small"
                              disabled={!col.included}
                              placeholder="Optional override…"
                              value={col.perColumnOverride}
                              onChange={(_, data) =>
                                setRangesModal((prev) => ({
                                  ...prev,
                                  columns: prev.columns.map((c) =>
                                    c.id === col.id ? { ...c, perColumnOverride: data.value } : c
                                  ),
                                }))
                              }
                            />
                          </td>
                          <td className={styles.colTd}>
                            <span className={previewClass}>{previewLabel}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Legend */}
            <Text className={shared.mutedText} style={{ fontSize: "11px" }}>
              🟢 Valid &nbsp; 🟡 Duplicate — already exists in target scope &nbsp; 🔴 Invalid — fails Excel naming rules
            </Text>

            {/* Actions */}
            <div className={styles.modalActions}>
              {hasErrors && (
                <Text className={shared.errorText} style={{ fontSize: "11px", marginRight: "auto" }}>
                  Resolve highlighted errors before creating.
                </Text>
              )}
              <Button onClick={() => setRangesModal(emptyRangesModal)}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={submitting || includedCount === 0 || hasErrors}
                onClick={() => void applyCreateRangesFromTable()}
              >
                Create {includedCount > 0 ? `${includedCount} Range(s)` : ""}
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {/* ================================================================== */}
      {/* BULK EDIT MODAL                                                     */}
      {/* ================================================================== */}
      {bulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Bulk Edit Tables</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Mode</Text>
                <Select
                  value={bulkState.mode}
                  onChange={(_, data) => setBulkState((p) => ({ ...p, mode: data.value as BulkState["mode"] }))}
                >
                  <option value="Prefix">Prefix</option>
                  <option value="Suffix">Suffix</option>
                  <option value="Replace">Replace</option>
                </Select>
              </div>
              <div>
                <Text className={shared.mutedText}>Case Transform</Text>
                <Select
                  value={bulkState.caseTransform}
                  onChange={(_, data) =>
                    setBulkState((p) => ({ ...p, caseTransform: data.value as BulkCaseTransform }))
                  }
                >
                  <option value="none">None</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={shared.mutedText}>{bulkState.mode === "Replace" ? "Find Text" : "Text"}</Text>
                <Input
                  value={bulkState.value}
                  onChange={(_, data) => setBulkState((p) => ({ ...p, value: data.value }))}
                />
              </div>
              {bulkState.mode === "Replace" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>Replace With</Text>
                  <Input
                    value={bulkState.replaceWith}
                    onChange={(_, data) => setBulkState((p) => ({ ...p, replaceWith: data.value }))}
                  />
                </div>
              ) : null}
            </div>
            <Text className={shared.mutedText}>{`Preview rows: ${bulkPreview.length}`}</Text>
            <div className={styles.modalActions}>
              <Button onClick={() => setBulkState((p) => ({ ...p, open: false }))}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={submitting || bulkPreview.length === 0}
                onClick={() => void applyBulk()}
              >
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
};

export default TablesView;
