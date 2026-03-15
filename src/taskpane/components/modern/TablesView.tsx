import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import { Checkmark20Regular, Dismiss20Regular, Edit20Regular } from "@fluentui/react-icons";
import {
  CreateNamedRangesFromTableRequest,
  TableRecord,
  createNamedRangesFromTableColumns,
  getTableColumns,
  getTables,
  selectTableAddress,
  updateTableName,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

type SortColumn = "name" | "address" | "sheet" | "scope";
type SortDirection = "asc" | "desc";
type CaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";

interface TablesViewProps {
  onOpenLegacy: () => void;
  embedded?: boolean;
}

interface BulkState {
  open: boolean;
  mode: "Prefix" | "Suffix" | "Replace";
  value: string;
  replaceWith: string;
  caseTransform: CaseTransform;
}

interface TableRangeModalState {
  open: boolean;
  table: TableRecord | null;
  columns: Array<{ id: string; name: string; address: string }>;
  selectedColumns: Set<string>;
  scopeType: "Workbook" | "Worksheet";
  conflictMode: CreateNamedRangesFromTableRequest["conflictMode"];
  conflictValue: string;
}

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
  table: {
    width: "100%",
    borderCollapse: "separate",
    borderSpacing: 0,
    minWidth: "840px",
    fontSize: "12px",
  },
  headCell: {
    position: "sticky",
    top: 0,
    zIndex: 1,
    backgroundColor: "#F3F4F6",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px 8px",
    whiteSpace: "nowrap",
  },
  sortBtn: {
    border: "none",
    background: "transparent",
    fontWeight: 600,
    cursor: "pointer",
    padding: 0,
  },
  row: {
    selectors: {
      "&:nth-child(even)": { backgroundColor: "#FCFCFD" },
      "&:hover": { backgroundColor: "#F3F8F4" },
    },
  },
  selectedRow: { backgroundColor: "#EAF2FF" },
  cell: {
    padding: "10px 8px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    whiteSpace: "nowrap",
  },
  clickableCellBtn: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
    padding: 0,
    textDecorationLine: "underline",
    fontSize: "12px",
  },
  rowActions: {
    display: "flex",
    gap: "4px",
    opacity: 1,
    pointerEvents: "auto",
    flexWrap: "wrap",
  },
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
    width: "min(720px, 100%)",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "12px",
    maxHeight: "88vh",
    overflow: "auto",
  },
  modalGrid: { display: "grid", gap: "12px", gridTemplateColumns: "repeat(2, minmax(0, 1fr))" },
  full: { gridColumn: "1 / -1" },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: "8px", flexWrap: "wrap" },
  columnList: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "8px",
    display: "grid",
    gap: "6px",
    maxHeight: "220px",
    overflow: "auto",
  },
  columnRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: "12px",
  },
  mutedCode: {
    fontFamily: "Consolas, 'Courier New', monospace",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "11px",
  },
});

const applyCase = (value: string, mode: CaseTransform): string => {
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

const TablesView: React.FC<TablesViewProps> = ({ onOpenLegacy, embedded = false }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
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
  const [rangesFromTable, setRangesFromTable] = useState<TableRangeModalState>({
    open: false,
    table: null,
    columns: [],
    selectedColumns: new Set<string>(),
    scopeType: "Worksheet",
    conflictMode: "prefix",
    conflictValue: "nr_",
  });
  const [bulkState, setBulkState] = useState<BulkState>({
    open: false,
    mode: "Prefix",
    value: "",
    replaceWith: "",
    caseTransform: "none",
  });

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

  const visibleRows = useMemo(() => {
    const token = search.trim().toLowerCase();
    return rows
      .filter((row) =>
        token
          ? `${row.name} ${row.address} ${row.sheet} ${row.scope}`.toLowerCase().includes(token)
          : true
      )
      .sort((a, b) => {
        const compare = a[sortColumn].localeCompare(b[sortColumn], undefined, {
          sensitivity: "base",
        });
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
      if (bulkState.mode === "Replace")
        nextName = row.name.replace(bulkState.value, bulkState.replaceWith);
      nextName = applyCase(nextName, bulkState.caseTransform);
      return { row, oldName: row.name, newName: nextName };
    });
  }, [rows, selectedIds, bulkState]);

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

  const openCreateRangesModal = async () => {
    const table = rows.find((row) => selectedIds.has(row.id));
    if (!table) {
      setStatusType("error");
      setStatus("Select one table first.");
      return;
    }
    setSubmitting(true);
    try {
      const columns = await getTableColumns(table.sheet, table.name);
      setRangesFromTable({
        open: true,
        table,
        columns,
        selectedColumns: new Set(columns.map((item) => item.name)),
        scopeType: "Worksheet",
        conflictMode: "prefix",
        conflictValue: "nr_",
      });
    } catch (error) {
      setStatusType("error");
      setStatus(
        `Unable to load table columns: ${error instanceof Error ? error.message : String(error)}`
      );
    } finally {
      setSubmitting(false);
    }
  };

  const applyCreateRangesFromTable = async () => {
    if (!rangesFromTable.table) return;
    const selectedColumns = Array.from(rangesFromTable.selectedColumns);
    if (selectedColumns.length === 0) {
      setStatusType("error");
      setStatus("Select at least one table column.");
      return;
    }

    setSubmitting(true);
    try {
      const result = await createNamedRangesFromTableColumns({
        sheetName: rangesFromTable.table.sheet,
        tableName: rangesFromTable.table.name,
        columns: selectedColumns,
        scopeType: rangesFromTable.scopeType,
        conflictMode: rangesFromTable.conflictMode,
        conflictValue: rangesFromTable.conflictValue,
      });
      setRangesFromTable((prev) => ({ ...prev, open: false }));
      setStatusType("success");
      setStatus(
        `Created ${result.created.length} named range(s)` +
          (result.skipped.length ? `, skipped ${result.skipped.length}.` : ".")
      );
    } catch (error) {
      setStatusType("error");
      setStatus(
        `Create from table failed: ${error instanceof Error ? error.message : String(error)}`
      );
    } finally {
      setSubmitting(false);
    }
  };

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

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Tables</Text>
          <Text className={shared.sectionSubtitle}>
            Selection-driven table actions and quick range generation.
          </Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button onClick={openCreateRangesModal} disabled={selectedIds.size !== 1 || submitting}>
            Create New Ranges from Table
          </Button>
          <Button
            onClick={() => setBulkState((prev) => ({ ...prev, open: true }))}
            disabled={selectedIds.size === 0}
          >
            Bulk Edit ({selectedIds.size})
          </Button>
          <Button onClick={() => void load()} disabled={loading}>
            Refresh
          </Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          <div className={styles.spacer} />
          <Input
            placeholder="Search tables..."
            value={search}
            onChange={(_, data) => setSearch(data.value)}
          />
        </div>
      </div>

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
                      if (sortColumn === column)
                        setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
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
              <tr
                key={row.id}
                className={`${styles.row} ${selectedIds.has(row.id) ? styles.selectedRow : ""}`}
              >
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
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<Checkmark20Regular />}
                          aria-label={`Save ${row.name}`}
                          title={`Save ${row.name}`}
                          disabled={submitting}
                          onClick={() => void saveInline(row)}
                        />
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<Dismiss20Regular />}
                          aria-label={`Cancel editing ${row.name}`}
                          title={`Cancel editing ${row.name}`}
                          onClick={() => setInlineEdit(null)}
                        />
                      </>
                    ) : (
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Edit20Regular />}
                        aria-label={`Edit ${row.name}`}
                        title={`Edit ${row.name}`}
                        onClick={() => setInlineEdit({ id: row.id, value: row.name })}
                      />
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

      {rangesFromTable.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create New Ranges from Table</Text>
            <Text className={shared.mutedText}>
              Table: {rangesFromTable.table?.name} ({rangesFromTable.table?.sheet})
            </Text>

            <div>
              <Text className={shared.mutedText}>Columns</Text>
              <div className={styles.columnList}>
                {rangesFromTable.columns.map((column) => (
                  <label key={column.id} className={styles.columnRow}>
                    <input
                      type="checkbox"
                      checked={rangesFromTable.selectedColumns.has(column.name)}
                      onChange={() =>
                        setRangesFromTable((prev) => {
                          const next = new Set(prev.selectedColumns);
                          if (next.has(column.name)) next.delete(column.name);
                          else next.add(column.name);
                          return { ...prev, selectedColumns: next };
                        })
                      }
                    />
                    <span>{column.name}</span>
                    <span className={styles.mutedCode}>{column.address}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Scope</Text>
                <Select
                  value={rangesFromTable.scopeType}
                  onChange={(_, data) =>
                    setRangesFromTable((prev) => ({
                      ...prev,
                      scopeType: data.value as "Workbook" | "Worksheet",
                    }))
                  }
                >
                  <option value="Worksheet">Worksheet</option>
                  <option value="Workbook">Workbook</option>
                </Select>
              </div>
              <div>
                <Text className={shared.mutedText}>Conflict Handling</Text>
                <Select
                  value={rangesFromTable.conflictMode}
                  onChange={(_, data) =>
                    setRangesFromTable((prev) => ({
                      ...prev,
                      conflictMode: data.value as CreateNamedRangesFromTableRequest["conflictMode"],
                    }))
                  }
                >
                  <option value="prefix">Add prefix</option>
                  <option value="suffix">Add suffix</option>
                  <option value="rename">Replace with explicit name</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={shared.mutedText}>
                  {rangesFromTable.conflictMode === "prefix"
                    ? "Prefix"
                    : rangesFromTable.conflictMode === "suffix"
                      ? "Suffix"
                      : "Replacement Name"}
                </Text>
                <Input
                  value={rangesFromTable.conflictValue}
                  onChange={(_, data) =>
                    setRangesFromTable((prev) => ({ ...prev, conflictValue: data.value }))
                  }
                />
              </div>
            </div>

            <div className={styles.modalActions}>
              <Button onClick={() => setRangesFromTable((prev) => ({ ...prev, open: false }))}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => void applyCreateRangesFromTable()}
                disabled={submitting}
              >
                Create
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {bulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Bulk Edit Tables</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Mode</Text>
                <Select
                  value={bulkState.mode}
                  onChange={(_, data) =>
                    setBulkState((p) => ({ ...p, mode: data.value as BulkState["mode"] }))
                  }
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
                    setBulkState((p) => ({ ...p, caseTransform: data.value as CaseTransform }))
                  }
                >
                  <option value="none">None</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={shared.mutedText}>
                  {bulkState.mode === "Replace" ? "Find Text" : "Text"}
                </Text>
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
