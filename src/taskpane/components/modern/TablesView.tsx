import * as React from "react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import {
<<<<<<< ours
<<<<<<< ours
<<<<<<< ours
  addNamedRange,
  getNamedRanges,
  getTableColumns,
  getTableFromCurrentSelection,
  getTables,
  selectTableRange,
  TableColumnRecord,
  TableRecord,
=======
=======
>>>>>>> theirs
=======
>>>>>>> theirs
  CreateNamedRangesFromTableRequest,
  TableRecord,
  createNamedRangesFromTableColumns,
  getTableColumns,
  getTables,
  selectTableAddress,
<<<<<<< ours
<<<<<<< ours
>>>>>>> theirs
=======
>>>>>>> theirs
=======
>>>>>>> theirs
  updateTableName,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

type SortColumn = "name" | "address" | "sheet" | "scope";
type SortDirection = "asc" | "desc";
type CaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";
type ScopeType = "Workbook" | "Worksheet";
type ConflictMode = "prefix" | "suffix" | "rename";

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

<<<<<<< ours
<<<<<<< ours
<<<<<<< ours
interface ExistingNameRecord {
  name: string;
  scopeType: ScopeType;
  scope: string;
}

interface CreateRangesState {
  open: boolean;
  table: TableRecord | null;
  columns: TableColumnRecord[];
  selectedColumnIds: Set<string>;
  existingNames: ExistingNameRecord[];
  scopeType: ScopeType;
  scope: string;
  conflictMode: ConflictMode;
  conflictValue: string;
  renameMap: Record<string, string>;
  loading: boolean;
}

const buildCreateRangesState = (): CreateRangesState => ({
  open: false,
  table: null,
  columns: [],
  selectedColumnIds: new Set<string>(),
  existingNames: [],
  scopeType: "Worksheet",
  scope: "",
  conflictMode: "prefix",
  conflictValue: "",
  renameMap: {},
  loading: false,
});

=======
=======
>>>>>>> theirs
=======
>>>>>>> theirs
interface TableRangeModalState {
  open: boolean;
  table: TableRecord | null;
  columns: Array<{ id: string; name: string; address: string }>;
  selectedColumns: Set<string>;
  scopeType: "Workbook" | "Worksheet";
  conflictMode: CreateNamedRangesFromTableRequest["conflictMode"];
  conflictValue: string;
}

<<<<<<< ours
<<<<<<< ours
>>>>>>> theirs
=======
>>>>>>> theirs
=======
>>>>>>> theirs
const useStyles = makeStyles({
  root: { display: "grid", gap: "18px" },
  toolbar: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    justifyContent: "space-between",
    flexWrap: "wrap",
  },
  toolbarGroup: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
  },
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
    minWidth: "880px",
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
    cursor: "pointer",
    selectors: {
      "&:nth-child(even)": { backgroundColor: "#FCFCFD" },
      "&:hover": { backgroundColor: "#F3F8F4" },
      "&:hover .row-actions": { opacity: 1, pointerEvents: "auto" },
    },
  },
<<<<<<< ours
<<<<<<< ours
<<<<<<< ours
  selectedRow: {
    backgroundColor: "#EAF3FF",
  },
  cell: {
    padding: "10px 8px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    whiteSpace: "nowrap",
    verticalAlign: "top",
  },
  rowActions: {
    display: "flex",
    gap: "4px",
    opacity: 0,
    pointerEvents: "none",
    transition: "opacity 150ms ease",
  },
  rowActionsOpen: {
    opacity: 1,
    pointerEvents: "auto",
  },
  linkBtn: {
=======
  selectedRow: { backgroundColor: "#EAF2FF" },
  cell: { padding: "10px 8px", borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, whiteSpace: "nowrap" },
  clickableCellBtn: {
>>>>>>> theirs
=======
  selectedRow: { backgroundColor: "#EAF2FF" },
  cell: { padding: "10px 8px", borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, whiteSpace: "nowrap" },
  clickableCellBtn: {
>>>>>>> theirs
=======
  selectedRow: { backgroundColor: "#EAF2FF" },
  cell: { padding: "10px 8px", borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, whiteSpace: "nowrap" },
  clickableCellBtn: {
>>>>>>> theirs
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
<<<<<<< ours
<<<<<<< ours
<<<<<<< ours
    textDecorationLine: "underline",
    padding: 0,
    fontSize: "12px",
  },
  cellEllipsis: {
    display: "block",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    minWidth: 0,
  },
  detailCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    display: "grid",
    gap: "12px",
  },
  detailHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "10px",
    flexWrap: "wrap",
  },
  detailMetaGrid: {
    display: "grid",
    gap: "10px",
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  detailLabel: {
    fontSize: "11px",
    lineHeight: "15px",
    fontWeight: 700,
    letterSpacing: "0.03em",
    textTransform: "uppercase",
    color: MODERN_TOKENS.colorTextMuted,
  },
  detailValue: {
    fontSize: "13px",
    lineHeight: "18px",
    color: MODERN_TOKENS.colorText,
  },
  detailActions: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
=======
=======
>>>>>>> theirs
=======
>>>>>>> theirs
    padding: 0,
    textDecorationLine: "underline",
    fontSize: "12px",
  },
  rowActions: { display: "flex", gap: "4px", opacity: 0, pointerEvents: "none", transition: "opacity 150ms ease" },
>>>>>>> theirs
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
<<<<<<< ours
<<<<<<< ours
<<<<<<< ours
    width: "min(760px, 100%)",
    maxHeight: "90vh",
    overflow: "auto",
=======
    width: "min(720px, 100%)",
>>>>>>> theirs
=======
    width: "min(720px, 100%)",
>>>>>>> theirs
=======
    width: "min(720px, 100%)",
>>>>>>> theirs
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "12px",
    maxHeight: "88vh",
    overflow: "auto",
<<<<<<< ours
<<<<<<< ours
  },
  modalGrid: {
    display: "grid",
    gap: "12px",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
=======
>>>>>>> theirs
=======
>>>>>>> theirs
  },
  full: { gridColumn: "1 / -1" },
<<<<<<< ours
  modalActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    flexWrap: "wrap",
  },
  columnList: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    backgroundColor: "#F9FAFB",
    display: "grid",
    gap: "1px",
    overflow: "hidden",
  },
  columnRow: {
    display: "grid",
    gridTemplateColumns: "auto minmax(0, 1fr) minmax(0, 1fr)",
    gap: "10px",
    alignItems: "center",
    padding: "10px 12px",
    backgroundColor: "#fff",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "auto minmax(0, 1fr)",
    },
  },
  columnAddress: {
    fontSize: "11px",
    lineHeight: "15px",
    color: MODERN_TOKENS.colorTextMuted,
    textAlign: "right",
    "@media (max-width: 720px)": {
      gridColumn: "2 / -1",
      textAlign: "left",
    },
  },
  conflictCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFF8EB",
    padding: "12px",
    display: "grid",
    gap: "10px",
  },
  renameGrid: {
    display: "grid",
    gap: "8px",
  },
  renameRow: {
    display: "grid",
    gap: "6px",
    gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)",
    "@media (max-width: 720px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  previewList: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    overflow: "hidden",
    backgroundColor: "#fff",
  },
  previewRow: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr) auto",
    gap: "10px",
    alignItems: "center",
    padding: "10px 12px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    "@media (max-width: 720px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
    selectors: {
      "&:last-child": {
        borderBottom: "none",
      },
    },
  },
  previewStatus: {
    fontSize: "11px",
    lineHeight: "15px",
    fontWeight: 700,
    color: MODERN_TOKENS.colorTextMuted,
  },
  previewStatusError: {
    color: MODERN_TOKENS.colorDanger,
=======
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
<<<<<<< ours
<<<<<<< ours
>>>>>>> theirs
=======
>>>>>>> theirs
=======
>>>>>>> theirs
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

const normalizeNameKey = (value: string) => value.trim().toUpperCase();

const buildUniqueTableName = (
  baseName: string,
  occupied: Set<string>,
  fallbackName: string
): string => {
  const trimmedBase = baseName.trim() || fallbackName;
  let candidate = trimmedBase;
  let suffix = 2;
  while (occupied.has(normalizeNameKey(candidate))) {
    candidate = `${trimmedBase}_${suffix.toString()}`;
    suffix += 1;
  }
  return candidate;
};

const toExcelNameCandidate = (value: string): string => {
  let candidate = value
    .trim()
    .replace(/[^A-Za-z0-9_.\\]+/g, "_")
    .replace(/_{2,}/g, "_")
    .replace(/^_+/, "");

  if (!candidate) {
    candidate = "Range";
  }
  if (!/^[A-Za-z_\\]/.test(candidate)) {
    candidate = `n_${candidate}`;
  }
  if (/^[A-Za-z]{1,3}\d+$/i.test(candidate)) {
    candidate = `n_${candidate}`;
  }
  return candidate;
};

const buildScopeKey = (scopeType: ScopeType, scope: string) =>
  scopeType === "Workbook" ? "Workbook" : `Worksheet::${scope.trim().toUpperCase()}`;

const err = (error: unknown): string => (error instanceof Error ? error.message : String(error));

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
  const [selectedTableId, setSelectedTableId] = useState<string | null>(null);
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
  const [createRangesState, setCreateRangesState] = useState<CreateRangesState>(
    buildCreateRangesState()
  );

  const load = useCallback(
    async (showStatus = true) => {
      setLoading(true);
      try {
        const loaded = await getTables();
        setRows(loaded);
        setSelectedIds(new Set());
        setInlineEdit(null);
        setSelectedTableId((prev) => {
          if (prev && loaded.some((row) => row.id === prev)) {
            return prev;
          }
          return loaded[0]?.id ?? null;
        });
        if (showStatus) {
          setStatusType("success");
          setStatus(`Loaded ${loaded.length} table(s).`);
        }
      } catch (error) {
        setStatusType("error");
        setStatus(`Load failed: ${err(error)}`);
      } finally {
        setLoading(false);
      }
    },
    []
  );

  useEffect(() => {
    void load();
  }, [load]);

  useEffect(() => {
    const handleFocus = () => {
      if (inlineEdit || bulkState.open || createRangesState.open) {
        return;
      }
      void load(false);
    };
    window.addEventListener("focus", handleFocus);
    return () => {
      window.removeEventListener("focus", handleFocus);
    };
  }, [bulkState.open, createRangesState.open, inlineEdit, load]);

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

  const selectedTable = useMemo(
    () => rows.find((row) => row.id === selectedTableId) ?? null,
    [rows, selectedTableId]
  );

  const allSelected = visibleRows.length > 0 && visibleRows.every((row) => selectedIds.has(row.id));

  const bulkPreview = useMemo(() => {
    const selected = rows.filter((row) => selectedIds.has(row.id));
    const occupied = new Set(
      rows.filter((row) => !selectedIds.has(row.id)).map((row) => normalizeNameKey(row.name))
    );

    return selected.map((row) => {
      let requestedName = row.name;
      if (bulkState.mode === "Prefix") requestedName = `${bulkState.value}${row.name}`;
      if (bulkState.mode === "Suffix") requestedName = `${row.name}${bulkState.value}`;
      if (bulkState.mode === "Replace")
        requestedName = row.name.replace(bulkState.value, bulkState.replaceWith);
      requestedName = applyCase(requestedName, bulkState.caseTransform).trim() || row.name;
      const newName = buildUniqueTableName(requestedName, occupied, row.name);
      occupied.add(normalizeNameKey(newName));
      return { row, oldName: row.name, requestedName, newName };
    });
  }, [rows, selectedIds, bulkState]);

  const createScopeValue =
    createRangesState.scopeType === "Worksheet"
      ? createRangesState.scope.trim() || createRangesState.table?.sheet || ""
      : "";

  const createPreview = useMemo(() => {
    const selectedColumns = createRangesState.columns.filter((column) =>
      createRangesState.selectedColumnIds.has(column.id)
    );
    const scopeKey = buildScopeKey(createRangesState.scopeType, createScopeValue);
    const existingNamesInScope = new Set(
      createRangesState.existingNames
        .filter((item) => buildScopeKey(item.scopeType, item.scope) === scopeKey)
        .map((item) => normalizeNameKey(item.name))
    );
    const occupied = new Set(existingNamesInScope);

    return selectedColumns.map((column) => {
      const baseName = toExcelNameCandidate(column.name);
      const existingConflict = existingNamesInScope.has(normalizeNameKey(baseName));
      let requestedName = baseName;
      let requiresUserResolution = false;

      if (existingConflict) {
        if (createRangesState.conflictMode === "rename") {
          requestedName = toExcelNameCandidate(createRangesState.renameMap[column.id] ?? "");
          requiresUserResolution = !(createRangesState.renameMap[column.id] ?? "").trim();
        } else if (createRangesState.conflictMode === "prefix") {
          const prefix = createRangesState.conflictValue.trim();
          requestedName = prefix ? toExcelNameCandidate(`${prefix}_${column.name}`) : baseName;
          requiresUserResolution = !prefix;
        } else {
          const suffix = createRangesState.conflictValue.trim();
          requestedName = suffix ? toExcelNameCandidate(`${column.name}_${suffix}`) : baseName;
          requiresUserResolution = !suffix;
        }
      }

      if (!requestedName.trim()) {
        requestedName = baseName;
        requiresUserResolution = true;
      }

      const finalName = requiresUserResolution
        ? requestedName
        : buildUniqueTableName(requestedName, occupied, baseName);

      if (!requiresUserResolution) {
        occupied.add(normalizeNameKey(finalName));
      }

      return {
        column,
        baseName,
        finalName,
        existingConflict,
        requiresUserResolution,
      };
    });
  }, [createRangesState, createScopeValue]);

  const conflictingPreviewRows = useMemo(
    () => createPreview.filter((item) => item.existingConflict),
    [createPreview]
  );

  const saveInline = async (row: TableRecord) => {
    if (!inlineEdit) return;
    setSubmitting(true);
    try {
      await updateTableName(row.sheet, row.name, inlineEdit.value.trim());
      setInlineEdit(null);
      await load(false);
      setSelectedTableId(row.id);
      setStatusType("success");
      setStatus("Table renamed.");
    } catch (error) {
      setStatusType("error");
      setStatus(`Rename failed: ${err(error)}`);
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
      setStatus(`Unable to load table columns: ${error instanceof Error ? error.message : String(error)}`);
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
      setStatus(`Create from table failed: ${error instanceof Error ? error.message : String(error)}`);
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
      setStatus(`Unable to load table columns: ${error instanceof Error ? error.message : String(error)}`);
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
      setStatus(`Create from table failed: ${error instanceof Error ? error.message : String(error)}`);
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
      setStatus(`Unable to load table columns: ${error instanceof Error ? error.message : String(error)}`);
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
      setStatus(`Create from table failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const applyBulk = async () => {
    if (bulkPreview.length === 0) return;
    setSubmitting(true);
    try {
      const changed = bulkPreview.filter(
        (item) => normalizeNameKey(item.oldName) !== normalizeNameKey(item.newName)
      );
      const occupied = new Set(rows.map((row) => normalizeNameKey(row.name)));
      const temporaryNames = changed.map((item, index) => {
        let tempName = `__WBM_TMP_TABLE_${index.toString()}__`;
        while (occupied.has(normalizeNameKey(tempName))) {
          tempName = `__WBM_TMP_TABLE_${index.toString()}_${occupied.size.toString()}__`;
        }
        occupied.add(normalizeNameKey(tempName));
        return { ...item, tempName };
      });

      for (const item of temporaryNames) {
        await updateTableName(item.row.sheet, item.oldName, item.tempName);
      }
      for (const item of temporaryNames) {
        await updateTableName(item.row.sheet, item.tempName, item.newName);
      }

      setBulkState((prev) => ({ ...prev, open: false }));
      setSelectedIds(new Set());
      await load(false);
      setStatusType("success");
      const adjustedCount = bulkPreview.filter(
        (item) => normalizeNameKey(item.requestedName) !== normalizeNameKey(item.newName)
      ).length;
      setStatus(
        adjustedCount > 0
          ? `Bulk update completed for ${bulkPreview.length} table(s). ${adjustedCount.toString()} rename(s) were adjusted to avoid collisions.`
          : `Bulk update completed for ${bulkPreview.length} table(s).`
      );
    } catch (error) {
      setStatusType("error");
      setStatus(`Bulk update failed: ${err(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const openCreateRanges = async (table: TableRecord) => {
    setSelectedTableId(table.id);
    setCreateRangesState({
      ...buildCreateRangesState(),
      open: true,
      loading: true,
      table,
      scopeType: "Worksheet",
      scope: table.sheet,
    });

    try {
      const [columns, names] = await Promise.all([
        getTableColumns(table.sheet, table.name),
        getNamedRanges(),
      ]);
      setCreateRangesState({
        open: true,
        table,
        columns,
        selectedColumnIds: new Set(columns.map((column) => column.id)),
        existingNames: names
          .filter((item) => item.kind === "NamedRange")
          .map((item) => ({
            name: item.name,
            scopeType: item.scopeType,
            scope: item.scope,
          })),
        scopeType: "Worksheet",
        scope: table.sheet,
        conflictMode: "prefix",
        conflictValue: "",
        renameMap: {},
        loading: false,
      });
    } catch (error) {
      setCreateRangesState(buildCreateRangesState());
      setStatusType("error");
      setStatus(`Unable to load table columns: ${err(error)}`);
    }
  };

  const applyCreateRanges = async () => {
    if (!createRangesState.table) {
      return;
    }
    if (createPreview.length === 0) {
      setStatusType("error");
      setStatus("Select at least one table column.");
      return;
    }
    if (createRangesState.scopeType === "Worksheet" && !createScopeValue.trim()) {
      setStatusType("error");
      setStatus("Worksheet scope is required.");
      return;
    }

    const unresolved = createPreview.filter((item) => item.requiresUserResolution);
    if (unresolved.length > 0) {
      setStatusType("error");
      setStatus("Resolve the naming conflict options before creating ranges.");
      return;
    }

    setSubmitting(true);
    try {
      for (const item of createPreview) {
        await addNamedRange(
          createRangesState.scopeType,
          createRangesState.scopeType === "Worksheet" ? createScopeValue : "",
          item.finalName,
          item.column.address,
          item.column.sheet,
          "Reference"
        );
      }
      setCreateRangesState(buildCreateRangesState());
      setStatusType("success");
      setStatus(
        `Created ${createPreview.length.toString()} named range(s) from ${createRangesState.table.name}.`
      );
    } catch (error) {
      setStatusType("error");
      setStatus(`Create ranges failed: ${err(error)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const handleSelectOnSheet = async (row: TableRecord) => {
    try {
      setSelectedTableId(row.id);
      await selectTableRange(row.sheet, row.name);
      setStatusType("success");
      setStatus(`Selected table "${row.name}" on ${row.sheet}.`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Unable to select table "${row.name}": ${err(error)}`);
    }
  };

  const useCurrentSelection = async () => {
    try {
      const table = await getTableFromCurrentSelection();
      if (!table) {
        throw new Error("Current selection is not inside a table.");
      }
      setSelectedTableId(table.id);
      setStatusType("success");
      setStatus(`Current selection mapped to table "${table.name}".`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Use current selection failed: ${err(error)}`);
    }
  };

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
<<<<<<< ours
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Tables</Text>
          <Text className={shared.sectionSubtitle}>
            Select a table, act on it, and create related named ranges without leaving the pane.
          </Text>
        </div>
      ) : null}
      <div className={shared.card}>
        <div className={styles.toolbar}>
          <div className={styles.toolbarGroup}>
            <Button onClick={() => void useCurrentSelection()}>Use Current Selection</Button>
            <Button
              appearance="primary"
              onClick={() => (selectedTable ? void openCreateRanges(selectedTable) : undefined)}
              disabled={!selectedTable}
            >
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
          </div>
          <div className={styles.toolbarGroup}>
            <Input
              placeholder="Search tables..."
              aria-label="Search tables"
              value={search}
              onChange={(_, data) => setSearch(data.value)}
            />
          </div>
=======
      <div>
        <Text className={shared.sectionTitle}>Tables</Text>
        <Text className={shared.sectionSubtitle}>Selection-driven table actions and quick range generation.</Text>
      </div>

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button onClick={openCreateRangesModal} disabled={selectedIds.size !== 1 || submitting}>
            Create New Ranges from Table
          </Button>
          <Button onClick={() => setBulkState((prev) => ({ ...prev, open: true }))} disabled={selectedIds.size === 0}>
            Bulk Edit ({selectedIds.size})
          </Button>
          <Button onClick={() => void load()} disabled={loading}>Refresh</Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          <div className={styles.spacer} />
          <Input placeholder="Search tables..." value={search} onChange={(_, data) => setSearch(data.value)} />
>>>>>>> theirs
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
                  aria-label={
                    allSelected
                      ? "Clear selection for all visible tables"
                      : "Select all visible tables"
                  }
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
                      if (sortColumn === column) {
                        setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
                      } else {
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
<<<<<<< ours
            {visibleRows.map((row) => {
              const selected = row.id === selectedTableId;
              return (
                <tr
                  key={row.id}
                  className={`${styles.row} ${selected ? styles.selectedRow : ""}`}
                  onClick={() => setSelectedTableId(row.id)}
                >
                  <td className={styles.cell} onClick={(event) => event.stopPropagation()}>
                    <input
                      type="checkbox"
                      checked={selectedIds.has(row.id)}
                      aria-label={`Select table ${row.name}`}
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
                  <td className={styles.cell} onClick={(event) => event.stopPropagation()}>
                    <div
                      className={`${styles.rowActions} row-actions ${
                        selected ? styles.rowActionsOpen : ""
                      }`}
                    >
                      {inlineEdit?.id === row.id ? (
                        <>
                          <Button
                            size="small"
                            disabled={submitting}
                            onClick={() => void saveInline(row)}
                          >
                            Save
                          </Button>
                          <Button size="small" onClick={() => setInlineEdit(null)}>
                            Cancel
                          </Button>
                        </>
                      ) : (
                        <>
                          <Button
                            size="small"
                            onClick={() => {
                              setSelectedTableId(row.id);
                              setInlineEdit({ id: row.id, value: row.name });
                            }}
                          >
                            Edit
                          </Button>
                          <Button size="small" onClick={() => void handleSelectOnSheet(row)}>
                            Select
                          </Button>
                          <Button size="small" onClick={() => void openCreateRanges(row)}>
                            Ranges
                          </Button>
                        </>
                      )}
                    </div>
                  </td>
                  <td className={styles.cell}>
=======
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
>>>>>>> theirs
                    {inlineEdit?.id === row.id ? (
                      <Input
                        aria-label={`Edit table name for ${row.name}`}
                        value={inlineEdit.value}
                        onChange={(_, data) => setInlineEdit({ id: row.id, value: data.value })}
                        onClick={(event) => event.stopPropagation()}
                      />
                    ) : (
                      <button
                        type="button"
                        className={styles.linkBtn}
                        onClick={(event) => {
                          event.stopPropagation();
                          void handleSelectOnSheet(row);
                        }}
                      >
                        {row.name}
                      </button>
                    )}
<<<<<<< ours
                  </td>
                  <td className={styles.cell}>
                    <span className={styles.cellEllipsis} title={row.address}>
                      {row.address}
                    </span>
                  </td>
                  <td className={styles.cell}>{row.sheet}</td>
                  <td className={styles.cell}>{row.scope}</td>
                </tr>
              );
            })}
=======
                  </div>
                </td>
                <td className={styles.cell}>
                  {inlineEdit?.id === row.id ? (
                    <Input value={inlineEdit.value} onChange={(_, data) => setInlineEdit({ id: row.id, value: data.value })} />
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
>>>>>>> theirs
          </tbody>
        </table>
      </div>

      <div className={styles.detailCard}>
        <div className={styles.detailHeader}>
          <Text className={shared.cardTitle}>Selected Table</Text>
          <div className={styles.detailActions}>
            <Button size="small" onClick={() => void useCurrentSelection()}>
              Use Current Selection
            </Button>
            <Button
              size="small"
              appearance="primary"
              onClick={() => (selectedTable ? void openCreateRanges(selectedTable) : undefined)}
              disabled={!selectedTable}
            >
              Create New Ranges from Table
            </Button>
          </div>
        </div>
        {selectedTable ? (
          <>
            <div className={styles.detailMetaGrid}>
              <div>
                <Text className={styles.detailLabel}>Name</Text>
                <Text className={styles.detailValue}>{selectedTable.name}</Text>
              </div>
              <div>
                <Text className={styles.detailLabel}>Sheet</Text>
                <Text className={styles.detailValue}>{selectedTable.sheet}</Text>
              </div>
              <div>
                <Text className={styles.detailLabel}>Address</Text>
                <Text className={styles.detailValue}>{selectedTable.address}</Text>
              </div>
            </div>
            <Text className={shared.mutedText}>
              Select the table, act on it, and keep the result visible here. The range creation
              flow reuses this selection and prompts when a name already exists.
            </Text>
            <div className={styles.detailActions}>
              <Button onClick={() => void handleSelectOnSheet(selectedTable)}>
                Select on Sheet
              </Button>
              <Button onClick={() => setInlineEdit({ id: selectedTable.id, value: selectedTable.name })}>
                Rename Table
              </Button>
            </div>
          </>
        ) : (
          <Text className={shared.mutedText}>
            Click a table in the list or use the current Excel selection to load its actions here.
          </Text>
        )}
      </div>

      <Text className={shared.mutedText}>{`Showing ${visibleRows.length} of ${rows.length}`}</Text>
      {status ? <Text className={statusClass}>{status}</Text> : null}
<<<<<<< ours
=======

      {rangesFromTable.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create New Ranges from Table</Text>
            <Text className={shared.mutedText}>
              Table: {rangesFromTable.table?.name} ({rangesFromTable.table?.sheet})
            </Text>

<<<<<<< ours
<<<<<<< ours
=======
      {rangesFromTable.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create New Ranges from Table</Text>
            <Text className={shared.mutedText}>
              Table: {rangesFromTable.table?.name} ({rangesFromTable.table?.sheet})
            </Text>

>>>>>>> theirs
=======
>>>>>>> theirs
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
                    setRangesFromTable((prev) => ({ ...prev, scopeType: data.value as "Workbook" | "Worksheet" }))
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
                  onChange={(_, data) => setRangesFromTable((prev) => ({ ...prev, conflictValue: data.value }))}
                />
              </div>
            </div>

            <div className={styles.modalActions}>
              <Button onClick={() => setRangesFromTable((prev) => ({ ...prev, open: false }))}>Cancel</Button>
              <Button appearance="primary" onClick={() => void applyCreateRangesFromTable()} disabled={submitting}>
                Create
              </Button>
            </div>
          </div>
        </div>
      ) : null}

<<<<<<< ours
<<<<<<< ours
>>>>>>> theirs
=======
>>>>>>> theirs
=======
>>>>>>> theirs
      {bulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Bulk Edit Tables</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Mode</Text>
                <Select
                  aria-label="Bulk edit mode"
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
                  aria-label="Bulk edit case transform"
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
                  aria-label={
                    bulkState.mode === "Replace" ? "Bulk edit find text" : "Bulk edit text"
                  }
                  value={bulkState.value}
                  onChange={(_, data) => setBulkState((p) => ({ ...p, value: data.value }))}
                />
              </div>
              {bulkState.mode === "Replace" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>Replace With</Text>
                  <Input
                    aria-label="Bulk edit replacement text"
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

      {createRangesState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create New Ranges from Table</Text>
            {createRangesState.table ? (
              <Text className={shared.mutedText}>
                {createRangesState.table.name} on {createRangesState.table.sheet}
              </Text>
            ) : null}
            {createRangesState.loading ? (
              <Text className={shared.mutedText}>Loading table columns...</Text>
            ) : (
              <>
                <div className={styles.modalGrid}>
                  <div>
                    <Text className={shared.mutedText}>Scope</Text>
                    <Select
                      value={createRangesState.scopeType}
                      onChange={(_, data) =>
                        setCreateRangesState((prev) => ({
                          ...prev,
                          scopeType: data.value as ScopeType,
                        }))
                      }
                    >
                      <option value="Worksheet">Worksheet</option>
                      <option value="Workbook">Workbook</option>
                    </Select>
                  </div>
                  <div>
                    <Text className={shared.mutedText}>Worksheet Scope</Text>
                    <Input
                      disabled={createRangesState.scopeType !== "Worksheet"}
                      value={createRangesState.scopeType === "Worksheet" ? createScopeValue : ""}
                      onChange={(_, data) =>
                        setCreateRangesState((prev) => ({ ...prev, scope: data.value }))
                      }
                    />
                  </div>
                </div>
                <div className={styles.full}>
                  <Text className={shared.mutedText}>
                    Select the table columns to convert into named ranges.
                  </Text>
                  <div className={styles.columnList}>
                    {createRangesState.columns.map((column) => (
                      <label key={column.id} className={styles.columnRow}>
                        <input
                          type="checkbox"
                          checked={createRangesState.selectedColumnIds.has(column.id)}
                          onChange={() =>
                            setCreateRangesState((prev) => {
                              const next = new Set(prev.selectedColumnIds);
                              if (next.has(column.id)) next.delete(column.id);
                              else next.add(column.id);
                              return { ...prev, selectedColumnIds: next };
                            })
                          }
                        />
                        <span>{column.name}</span>
                        <span className={styles.columnAddress}>{column.address}</span>
                      </label>
                    ))}
                  </div>
                </div>
                {conflictingPreviewRows.length > 0 ? (
                  <div className={styles.conflictCard}>
                    <Text className={shared.cardTitle}>Conflict Handling</Text>
                    <Text className={shared.mutedText}>
                      One or more column names already exist in the selected scope. Choose how to
                      resolve them before creating the ranges.
                    </Text>
                    <div className={styles.modalGrid}>
                      <div>
                        <Text className={shared.mutedText}>Resolution</Text>
                        <Select
                          value={createRangesState.conflictMode}
                          onChange={(_, data) =>
                            setCreateRangesState((prev) => ({
                              ...prev,
                              conflictMode: data.value as ConflictMode,
                            }))
                          }
                        >
                          <option value="prefix">Add Prefix</option>
                          <option value="suffix">Add Suffix</option>
                          <option value="rename">Rename Conflicts</option>
                        </Select>
                      </div>
                      {createRangesState.conflictMode !== "rename" ? (
                        <div>
                          <Text className={shared.mutedText}>
                            {createRangesState.conflictMode === "prefix" ? "Prefix" : "Suffix"}
                          </Text>
                          <Input
                            value={createRangesState.conflictValue}
                            onChange={(_, data) =>
                              setCreateRangesState((prev) => ({
                                ...prev,
                                conflictValue: data.value,
                              }))
                            }
                          />
                        </div>
                      ) : null}
                    </div>
                    {createRangesState.conflictMode === "rename" ? (
                      <div className={styles.renameGrid}>
                        {conflictingPreviewRows.map((item) => (
                          <div key={item.column.id} className={styles.renameRow}>
                            <Text className={shared.mutedText}>{item.column.name}</Text>
                            <Input
                              placeholder={`New name for ${item.column.name}`}
                              value={createRangesState.renameMap[item.column.id] ?? ""}
                              onChange={(_, data) =>
                                setCreateRangesState((prev) => ({
                                  ...prev,
                                  renameMap: {
                                    ...prev.renameMap,
                                    [item.column.id]: data.value,
                                  },
                                }))
                              }
                            />
                          </div>
                        ))}
                      </div>
                    ) : null}
                  </div>
                ) : null}
                <div className={styles.previewList}>
                  {createPreview.map((item) => (
                    <div key={item.column.id} className={styles.previewRow}>
                      <div>
                        <Text>{item.column.name}</Text>
                        <Text className={shared.mutedText}>{item.column.address}</Text>
                      </div>
                      <div>
                        <Text>{item.finalName}</Text>
                        <Text className={shared.mutedText}>{item.baseName}</Text>
                      </div>
                      <div
                        className={`${styles.previewStatus} ${
                          item.requiresUserResolution ? styles.previewStatusError : ""
                        }`}
                      >
                        {item.requiresUserResolution
                          ? "Needs input"
                          : item.existingConflict
                            ? "Conflict resolved"
                            : "Ready"}
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
            <div className={styles.modalActions}>
              <Button onClick={() => setCreateRangesState(buildCreateRangesState())}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={submitting || createRangesState.loading}
                onClick={() => void applyCreateRanges()}
              >
                Create
              </Button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
};

export default TablesView;
