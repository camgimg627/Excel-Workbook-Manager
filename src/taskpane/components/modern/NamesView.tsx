import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import {
  NamedRangeRecord,
  addNamedRange,
  deleteNamedRangeWithOptions,
  getCurrentSelectionAddress,
  getNamedRanges,
  moveNamedRange,
  selectNamedRangeAddress,
  updateNamedRange,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

type SortColumn = "name" | "address" | "sheet" | "scope" | "type";
type SortDirection = "asc" | "desc";
type ScopeType = "Workbook" | "Worksheet";
type CaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";

interface NamesViewProps {
  createRequestId: number;
  onOpenLegacy: () => void;
}

interface EditState {
  open: boolean;
  record: NamedRangeRecord | null;
  name: string;
  address: string;
  fallbackSheet: string;
  caseTransform: CaseTransform;
}

interface CreateState {
  open: boolean;
  scopeType: ScopeType;
  scope: string;
  name: string;
  address: string;
  fallbackSheet: string;
}

interface MoveState {
  open: boolean;
  record: NamedRangeRecord | null;
  address: string;
  fallbackSheet: string;
}

interface DeleteState {
  open: boolean;
  record: NamedRangeRecord | null;
  deleteName: boolean;
  deleteValues: boolean;
}

interface BulkState {
  open: boolean;
  mode: "Prefix" | "Suffix" | "Replace" | "Delete";
  value: string;
  replaceWith: string;
  caseTransform: CaseTransform;
  deleteName: boolean;
  deleteValues: boolean;
}

const RANGE_COLUMNS: SortColumn[] = ["name", "address", "sheet", "scope", "type"];

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
  table: { width: "100%", borderCollapse: "separate", borderSpacing: 0, minWidth: "980px", fontSize: "12px" },
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
  cell: { padding: "10px 8px", borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, whiteSpace: "nowrap" },
  rowActions: {
    display: "flex",
    gap: "4px",
    opacity: 0,
    pointerEvents: "none",
    transition: "opacity 150ms ease",
  },
  linkBtn: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
    textDecorationLine: "underline",
    padding: 0,
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
    maxHeight: "90vh",
    overflow: "auto",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "12px",
  },
  modalGrid: { display: "grid", gap: "12px", gridTemplateColumns: "repeat(2, minmax(0, 1fr))" },
  full: { gridColumn: "1 / -1" },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: "8px", flexWrap: "wrap" },
});

const applyCase = (value: string, mode: CaseTransform): string => {
  if (mode === "none") return value;
  const tokens = value
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
  if (tokens.length === 0) return value;
  if (mode === "camelCase") {
    return tokens
      .map((part, idx) =>
        idx === 0
          ? `${part.charAt(0).toLowerCase()}${part.slice(1).toLowerCase()}`
          : `${part.charAt(0).toUpperCase()}${part.slice(1).toLowerCase()}`
      )
      .join("");
  }
  if (mode === "snake_case") return tokens.map((t) => t.toLowerCase()).join("_");
  return tokens.map((t) => t.toUpperCase()).join("_");
};

const err = (error: unknown): string => (error instanceof Error ? error.message : String(error));

const NamesView: React.FC<NamesViewProps> = ({ createRequestId, onOpenLegacy }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [rows, setRows] = useState<NamedRangeRecord[]>([]);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [loading, setLoading] = useState<boolean>(false);
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [search, setSearch] = useState<string>("");
  const [scopeFilter, setScopeFilter] = useState<"all" | ScopeType>("all");
  const [sortColumn, setSortColumn] = useState<SortColumn>("name");
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());

  const [editState, setEditState] = useState<EditState>({
    open: false,
    record: null,
    name: "",
    address: "",
    fallbackSheet: "",
    caseTransform: "none",
  });
  const [createState, setCreateState] = useState<CreateState>({
    open: false,
    scopeType: "Workbook",
    scope: "",
    name: "",
    address: "",
    fallbackSheet: "",
  });
  const [moveState, setMoveState] = useState<MoveState>({
    open: false,
    record: null,
    address: "",
    fallbackSheet: "",
  });
  const [deleteState, setDeleteState] = useState<DeleteState>({
    open: false,
    record: null,
    deleteName: true,
    deleteValues: false,
  });
  const [bulkState, setBulkState] = useState<BulkState>({
    open: false,
    mode: "Prefix",
    value: "",
    replaceWith: "",
    caseTransform: "none",
    deleteName: true,
    deleteValues: false,
  });

  const load = async () => {
    setLoading(true);
    try {
      const loaded = await getNamedRanges();
      setRows(loaded);
      setSelectedIds(new Set());
      setStatusType("success");
      setStatus(`Loaded ${loaded.length} name/shape record(s).`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Load failed: ${err(error)}`);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);

  useEffect(() => {
    if (createRequestId > 0) {
      setCreateState((prev) => ({ ...prev, open: true }));
    }
  }, [createRequestId]);

  const visibleRows = useMemo(() => {
    const token = search.trim().toLowerCase();
    return rows
      .filter((row) => (scopeFilter === "all" ? true : row.scopeType === scopeFilter))
      .filter((row) =>
        token
          ? `${row.name} ${row.address} ${row.sheet} ${row.scope} ${row.type}`
              .toLowerCase()
              .includes(token)
          : true
      )
      .sort((a, b) => {
        const compare = a[sortColumn].localeCompare(b[sortColumn], undefined, { sensitivity: "base" });
        return sortDirection === "asc" ? compare : -compare;
      });
  }, [rows, scopeFilter, search, sortColumn, sortDirection]);

  const selectableRows = useMemo(() => visibleRows.filter((row) => row.kind === "NamedRange"), [visibleRows]);
  const allSelected =
    selectableRows.length > 0 && selectableRows.every((row) => selectedIds.has(row.id));

  const bulkPreview = useMemo(() => {
    const selected = rows.filter((row) => row.kind === "NamedRange" && selectedIds.has(row.id));
    return selected.map((row) => {
      if (bulkState.mode === "Delete") {
        return { id: row.id, oldName: row.name, newName: row.name, row };
      }
      let newName = row.name;
      if (bulkState.mode === "Prefix") newName = `${bulkState.value}${row.name}`;
      if (bulkState.mode === "Suffix") newName = `${row.name}${bulkState.value}`;
      if (bulkState.mode === "Replace") newName = row.name.replace(bulkState.value, bulkState.replaceWith);
      newName = applyCase(newName, bulkState.caseTransform);
      return { id: row.id, oldName: row.name, newName, row };
    });
  }, [rows, selectedIds, bulkState]);

  const toggleSort = (column: SortColumn) => {
    if (sortColumn === column) {
      setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
      return;
    }
    setSortColumn(column);
    setSortDirection("asc");
  };

  const captureSelection = async (
    setter: (selection: { address: string; sheet: string }) => void
  ) => {
    try {
      const selection = await getCurrentSelectionAddress();
      setter(selection);
      setStatusType("success");
      setStatus("Current selection captured.");
    } catch (error) {
      setStatusType("error");
      setStatus(`Capture selection failed: ${err(error)}`);
    }
  };

  const runSubmit = async (action: () => Promise<void>, success: string, failure: string) => {
    setSubmitting(true);
    try {
      await action();
      setStatusType("success");
      setStatus(success);
      await load();
    } catch (error) {
      setStatusType("error");
      setStatus(`${failure}: ${err(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const submitCreate = async () => {
    const name = createState.name.trim();
    const address = createState.address.trim();
    if (!name || !address) {
      setStatusType("error");
      setStatus("Name and address are required.");
      return;
    }
    await runSubmit(
      async () => {
        const scope =
          createState.scopeType === "Worksheet"
            ? createState.scope.trim() || createState.fallbackSheet
            : createState.scope.trim();
        await addNamedRange(createState.scopeType, scope, name, address, createState.fallbackSheet);
        setCreateState((prev) => ({ ...prev, open: false, name: "", address: "" }));
      },
      "Named range created.",
      "Create failed"
    );
  };

  const submitEdit = async () => {
    const record = editState.record;
    if (!record) return;
    await runSubmit(
      async () => {
        await updateNamedRange(
          record.scopeType,
          record.scope,
          record.name,
          applyCase(editState.name.trim(), editState.caseTransform),
          editState.address.trim(),
          editState.fallbackSheet
        );
        setEditState({
          open: false,
          record: null,
          name: "",
          address: "",
          fallbackSheet: "",
          caseTransform: "none",
        });
      },
      "Named range updated.",
      "Update failed"
    );
  };

  const submitMove = async () => {
    const record = moveState.record;
    if (!record) return;
    await runSubmit(
      async () => {
        await moveNamedRange(record.scopeType, record.scope, record.name, moveState.address, moveState.fallbackSheet);
        setMoveState({
          open: false,
          record: null,
          address: "",
          fallbackSheet: "",
        });
      },
      "Named range moved.",
      "Move failed"
    );
  };

  const submitDelete = async () => {
    const record = deleteState.record;
    if (!record) return;
    if (!deleteState.deleteName && !deleteState.deleteValues) {
      setStatusType("error");
      setStatus("Select delete name and/or delete values.");
      return;
    }
    await runSubmit(
      async () => {
        await deleteNamedRangeWithOptions(
          record.scopeType,
          record.scope,
          record.name,
          deleteState.deleteName,
          deleteState.deleteValues
        );
        setDeleteState({
          open: false,
          record: null,
          deleteName: true,
          deleteValues: false,
        });
      },
      "Delete operation completed.",
      "Delete failed"
    );
  };

  const submitBulk = async () => {
    if (bulkPreview.length === 0) return;
    if (bulkState.mode === "Delete" && !bulkState.deleteName && !bulkState.deleteValues) {
      setStatusType("error");
      setStatus("For delete mode, pick delete name and/or delete values.");
      return;
    }

    await runSubmit(
      async () => {
        for (const row of bulkPreview) {
          if (bulkState.mode === "Delete") {
            await deleteNamedRangeWithOptions(
              row.row.scopeType,
              row.row.scope,
              row.row.name,
              bulkState.deleteName,
              bulkState.deleteValues
            );
            continue;
          }
          if (row.newName === row.oldName) continue;
          await updateNamedRange(
            row.row.scopeType,
            row.row.scope,
            row.oldName,
            row.newName,
            row.row.address,
            row.row.sheet
          );
        }
        setBulkState((prev) => ({ ...prev, open: false }));
        setSelectedIds(new Set());
      },
      `Bulk operation completed for ${bulkPreview.length} range(s).`,
      "Bulk operation failed"
    );
  };

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
      <div>
        <Text className={shared.sectionTitle}>Names</Text>
        <Text className={shared.sectionSubtitle}>Manage named ranges and workbook objects.</Text>
      </div>

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button appearance="primary" onClick={() => setCreateState((prev) => ({ ...prev, open: true }))}>
            + Create Name
          </Button>
          <Button onClick={() => setBulkState((prev) => ({ ...prev, open: true }))} disabled={selectedIds.size === 0}>
            Bulk Update ({selectedIds.size})
          </Button>
          <Button onClick={() => void load()} disabled={loading}>
            Refresh
          </Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          <div className={styles.spacer} />
          <Select value={scopeFilter} onChange={(_, data) => setScopeFilter(data.value as "all" | ScopeType)}>
            <option value="all">All Scopes</option>
            <option value="Workbook">Workbook</option>
            <option value="Worksheet">Worksheet</option>
          </Select>
          <Input
            placeholder="Search names, address, scope..."
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
                      if (allSelected) {
                        selectableRows.forEach((row) => next.delete(row.id));
                      } else {
                        selectableRows.forEach((row) => next.add(row.id));
                      }
                      return next;
                    })
                  }
                />
              </th>
              <th className={styles.headCell}>Actions</th>
              {RANGE_COLUMNS.map((column) => (
                <th key={column} className={styles.headCell}>
                  <button type="button" className={styles.sortBtn} onClick={() => toggleSort(column)}>
                    {column[0].toUpperCase() + column.slice(1)}
                    {sortColumn === column ? (sortDirection === "asc" ? " ▲" : " ▼") : ""}
                  </button>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={row.id} className={styles.row}>
                <td className={styles.cell}>
                  <input
                    type="checkbox"
                    checked={selectedIds.has(row.id)}
                    disabled={row.kind !== "NamedRange"}
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
                    <Button
                      size="small"
                      disabled={row.kind !== "NamedRange" || !row.isRange}
                      onClick={() =>
                        setEditState({
                          open: true,
                          record: row,
                          name: row.name,
                          address: row.address,
                          fallbackSheet: row.sheet,
                          caseTransform: "none",
                        })
                      }
                    >
                      Edit
                    </Button>
                    <Button
                      size="small"
                      disabled={!row.isRange}
                      onClick={() =>
                        setMoveState({
                          open: true,
                          record: row,
                          address: row.address,
                          fallbackSheet: row.sheet,
                        })
                      }
                    >
                      Move
                    </Button>
                    <Button
                      size="small"
                      disabled={row.kind !== "NamedRange"}
                      onClick={() =>
                        setDeleteState({
                          open: true,
                          record: row,
                          deleteName: true,
                          deleteValues: false,
                        })
                      }
                    >
                      Delete
                    </Button>
                  </div>
                </td>
                <td className={styles.cell}>
                  {row.isRange || row.kind === "Shape" ? (
                    <button
                      type="button"
                      className={styles.linkBtn}
                      onClick={() => void selectNamedRangeAddress(row.address, row.sheet)}
                    >
                      {row.name}
                    </button>
                  ) : (
                    row.name
                  )}
                </td>
                <td className={styles.cell}>{row.address}</td>
                <td className={styles.cell}>{row.sheet}</td>
                <td className={styles.cell}>{row.scope}</td>
                <td className={styles.cell}>{row.type}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <Text className={shared.mutedText}>{`Showing ${visibleRows.length} of ${rows.length}`}</Text>
      {status ? <Text className={statusClass}>{status}</Text> : null}

      {createState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create Named Range</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Name</Text>
                <Input value={createState.name} onChange={(_, data) => setCreateState((p) => ({ ...p, name: data.value }))} />
              </div>
              <div>
                <Text className={shared.mutedText}>Scope Type</Text>
                <Select
                  value={createState.scopeType}
                  onChange={(_, data) => setCreateState((p) => ({ ...p, scopeType: data.value as ScopeType }))}
                >
                  <option value="Workbook">Workbook</option>
                  <option value="Worksheet">Worksheet</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={shared.mutedText}>Address</Text>
                <Input
                  value={createState.address}
                  onChange={(_, data) => setCreateState((p) => ({ ...p, address: data.value }))}
                />
              </div>
              {createState.scopeType === "Worksheet" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>Worksheet Scope</Text>
                  <Input
                    value={createState.scope}
                    placeholder={createState.fallbackSheet || "Sheet1"}
                    onChange={(_, data) => setCreateState((p) => ({ ...p, scope: data.value }))}
                  />
                </div>
              ) : null}
            </div>
            <div className={styles.modalActions}>
              <Button
                onClick={() =>
                  void captureSelection((selection) =>
                    setCreateState((p) => ({
                      ...p,
                      address: selection.address,
                      fallbackSheet: selection.sheet,
                      scope: selection.sheet,
                    }))
                  )
                }
              >
                Use Selection
              </Button>
              <Button onClick={() => setCreateState((p) => ({ ...p, open: false }))}>Cancel</Button>
              <Button appearance="primary" disabled={submitting} onClick={() => void submitCreate()}>
                Create
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {editState.open && editState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Edit Named Range</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Name</Text>
                <Input value={editState.name} onChange={(_, data) => setEditState((p) => ({ ...p, name: data.value }))} />
              </div>
              <div>
                <Text className={shared.mutedText}>Case Transform</Text>
                <Select
                  value={editState.caseTransform}
                  onChange={(_, data) =>
                    setEditState((p) => ({ ...p, caseTransform: data.value as CaseTransform }))
                  }
                >
                  <option value="none">None</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={shared.mutedText}>Address</Text>
                <Input
                  value={editState.address}
                  onChange={(_, data) => setEditState((p) => ({ ...p, address: data.value }))}
                />
              </div>
            </div>
            <div className={styles.modalActions}>
              <Button
                onClick={() =>
                  void captureSelection((selection) =>
                    setEditState((p) => ({
                      ...p,
                      address: selection.address,
                      fallbackSheet: selection.sheet,
                    }))
                  )
                }
              >
                Use Selection
              </Button>
              <Button
                onClick={() =>
                  setEditState({
                    open: false,
                    record: null,
                    name: "",
                    address: "",
                    fallbackSheet: "",
                    caseTransform: "none",
                  })
                }
              >
                Cancel
              </Button>
              <Button appearance="primary" disabled={submitting} onClick={() => void submitEdit()}>
                Save
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {moveState.open && moveState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Move Named Range</Text>
            <Text className={shared.mutedText}>Address</Text>
            <Input value={moveState.address} onChange={(_, data) => setMoveState((p) => ({ ...p, address: data.value }))} />
            <div className={styles.modalActions}>
              <Button
                onClick={() =>
                  void captureSelection((selection) =>
                    setMoveState((p) => ({
                      ...p,
                      address: selection.address,
                      fallbackSheet: selection.sheet,
                    }))
                  )
                }
              >
                Use Selection
              </Button>
              <Button onClick={() => setMoveState({ open: false, record: null, address: "", fallbackSheet: "" })}>
                Cancel
              </Button>
              <Button appearance="primary" disabled={submitting} onClick={() => void submitMove()}>
                Move
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {deleteState.open && deleteState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Delete Named Range</Text>
            <label>
              <input
                type="checkbox"
                checked={deleteState.deleteName}
                onChange={(event) => setDeleteState((p) => ({ ...p, deleteName: event.target.checked }))}
              />{" "}
              Delete name
            </label>
            <label>
              <input
                type="checkbox"
                checked={deleteState.deleteValues}
                onChange={(event) => setDeleteState((p) => ({ ...p, deleteValues: event.target.checked }))}
              />{" "}
              Delete range values
            </label>
            <div className={styles.modalActions}>
              <Button
                onClick={() =>
                  setDeleteState({
                    open: false,
                    record: null,
                    deleteName: true,
                    deleteValues: false,
                  })
                }
              >
                Cancel
              </Button>
              <Button appearance="primary" disabled={submitting} onClick={() => void submitDelete()}>
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {bulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Bulk Update Named Ranges</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Mode</Text>
                <Select value={bulkState.mode} onChange={(_, data) => setBulkState((p) => ({ ...p, mode: data.value as BulkState["mode"] }))}>
                  <option value="Prefix">Prefix</option>
                  <option value="Suffix">Suffix</option>
                  <option value="Replace">Replace</option>
                  <option value="Delete">Delete</option>
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
              {bulkState.mode !== "Delete" ? (
                <>
                  <div className={styles.full}>
                    <Text className={shared.mutedText}>{bulkState.mode === "Replace" ? "Find Text" : "Text"}</Text>
                    <Input value={bulkState.value} onChange={(_, data) => setBulkState((p) => ({ ...p, value: data.value }))} />
                  </div>
                  {bulkState.mode === "Replace" ? (
                    <div className={styles.full}>
                      <Text className={shared.mutedText}>Replace With</Text>
                      <Input value={bulkState.replaceWith} onChange={(_, data) => setBulkState((p) => ({ ...p, replaceWith: data.value }))} />
                    </div>
                  ) : null}
                </>
              ) : (
                <div className={styles.full}>
                  <label>
                    <input
                      type="checkbox"
                      checked={bulkState.deleteName}
                      onChange={(event) => setBulkState((p) => ({ ...p, deleteName: event.target.checked }))}
                    />{" "}
                    Delete names
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={bulkState.deleteValues}
                      onChange={(event) => setBulkState((p) => ({ ...p, deleteValues: event.target.checked }))}
                    />{" "}
                    Delete range values
                  </label>
                </div>
              )}
            </div>
            <Text className={shared.mutedText}>{`Preview rows: ${bulkPreview.length}`}</Text>
            <div className={styles.modalActions}>
              <Button onClick={() => setBulkState((p) => ({ ...p, open: false }))}>Cancel</Button>
              <Button appearance="primary" disabled={submitting || bulkPreview.length === 0} onClick={() => void submitBulk()}>
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
};

export default NamesView;
