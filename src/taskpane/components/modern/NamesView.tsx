import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import { ArrowMove20Regular, Delete20Regular, Edit20Regular } from "@fluentui/react-icons";
import {
  NamedRangeRecord,
  addNamedRange,
  deleteNamedRangeWithOptions,
  getCurrentSelectionAddress,
  getNamedRanges,
  moveNamedRange,
  openFormulaEditorPopout,
  selectNamedRangeAddress,
  updateNamedRange,
} from "../../taskpane";
import { useDebounce } from "../../../hooks/useDebounce";
import {
  applyCase,
  buildLambdaFormula,
  buildNamedListFormula,
  buildUniqueScopedName,
  extractLambdaArgsFromFormula,
  normalizeNameKey,
  parseNamedListFormula,
} from "../../../utils/names.utils";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

type SortColumn = "name" | "address" | "sheet" | "scope" | "type";
type SortDirection = "asc" | "desc";
type ScopeType = "Workbook" | "Worksheet";
type CaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";
type CreateEntryType = "Range" | "Function" | "List";

interface NamesViewProps {
  createRequestId: number;
  onOpenLegacy: () => void;
  embedded?: boolean;
  onOpenShape?: (request: { sheetName: string; shapeName: string }) => void;
}

interface EditState {
  record: NamedRangeRecord | null;
  editType: "Range" | "List" | "Formula";
  name: string;
  address: string;
  listValues: string;
  fallbackSheet: string;
  caseTransform: CaseTransform;
}

interface CreateState {
  createType: CreateEntryType;
  scopeType: ScopeType;
  scope: string;
  name: string;
  address: string;
  lambdaArgs: string;
  functionBody: string;
  listValues: string;
  fallbackSheet: string;
}

interface MoveState {
  record: NamedRangeRecord | null;
  address: string;
  fallbackSheet: string;
}

interface DeleteState {
  record: NamedRangeRecord | null;
  deleteName: boolean;
  deleteValues: boolean;
}

interface BulkState {
  mode: "Prefix" | "Suffix" | "Replace" | "Delete";
  value: string;
  replaceWith: string;
  caseTransform: CaseTransform;
  deleteName: boolean;
  deleteValues: boolean;
}

type ActiveModal = "none" | "create" | "edit" | "move" | "delete" | "bulk";

const isInternalWorkbookSheetName = (sheetName: string): boolean => {
  const normalized = sheetName.trim();
  if (!normalized) {
    return false;
  }
  if (
    normalized === "__WBM_SHEET_FORMATS" ||
    normalized === "__WBM_META" ||
    normalized === "__WBM_FUNCTION_EVAL"
  ) {
    return true;
  }
  return normalized.startsWith("__WBM_FMT_");
};

const buildDefaultCreateState = (): CreateState => ({
  createType: "Range",
  scopeType: "Workbook",
  scope: "",
  name: "",
  address: "",
  lambdaArgs: "",
  functionBody: "",
  listValues: "",
  fallbackSheet: "",
});

const RANGE_COLUMNS: SortColumn[] = ["name", "address", "sheet", "scope", "type"];
const ADDRESS_COLUMN_MIN_WIDTH = 220;
const ADDRESS_COLUMN_MAX_WIDTH = 960;

const useStyles = makeStyles({
  root: { display: "grid", gap: "24px" },
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  toolbarPrimary: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  toolbarFilters: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
    marginLeft: "auto",
  },
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
    minWidth: "1120px",
    tableLayout: "fixed",
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
  cell: {
    padding: "10px 8px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    whiteSpace: "nowrap",
  },
  rowActions: {
    display: "flex",
    gap: "4px",
    opacity: 1,
    pointerEvents: "auto",
    flexWrap: "wrap",
  },
  linkBtn: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
    textDecorationLine: "underline",
    padding: 0,
  },
  cellEllipsis: {
    display: "block",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    minWidth: 0,
  },
  typeFilterWrap: { position: "relative" },
  typeFilterSummary: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "6px",
    height: "32px",
    display: "inline-flex",
    alignItems: "center",
    padding: "0 10px",
    cursor: "pointer",
    userSelect: "none",
    backgroundColor: "#fff",
    color: MODERN_TOKENS.colorText,
    fontSize: "12px",
    fontWeight: 400,
    textAlign: "left",
    selectors: {
      "&:focus-visible": {
        outline: `2px solid ${MODERN_TOKENS.colorBrand}`,
        outlineOffset: "1px",
      },
    },
  },
  typeFilterButtonOpen: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    boxShadow: MODERN_TOKENS.shadowCard,
  },
  typeFilterMenu: {
    position: "absolute",
    top: "calc(100% + 4px)",
    right: 0,
    zIndex: 6,
    minWidth: "220px",
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "8px",
    display: "grid",
    gap: "6px",
  },
  typeFilterRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: MODERN_TOKENS.colorText,
    fontSize: "12px",
    whiteSpace: "nowrap",
  },
  typeFilterDivider: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    margin: "2px 0",
  },
  addressHeadCell: {
    position: "sticky",
    top: 0,
    paddingRight: "14px",
  },
  columnResizeHandle: {
    position: "absolute",
    top: 0,
    right: 0,
    width: "9px",
    height: "100%",
    cursor: "col-resize",
    userSelect: "none",
    touchAction: "none",
    backgroundColor: "transparent",
    selectors: {
      "&:hover": { backgroundColor: "#D1D5DB" },
    },
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
  multilineInput: {
    width: "100%",
    minHeight: "120px",
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "12px",
    lineHeight: "1.4",
    boxSizing: "border-box",
    resize: "vertical",
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: "6px" },
  chip: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "999px",
    padding: "2px 8px",
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    backgroundColor: "#F9FAFB",
  },
});

const buildScopeKey = (row: Pick<NamedRangeRecord, "scopeType" | "scope">) =>
  row.scopeType === "Workbook" ? "Workbook" : `Worksheet::${row.scope.toUpperCase()}`;

const err = (error: unknown): string => (error instanceof Error ? error.message : String(error));

const parseFunctionArgs = (value: string): string[] =>
  value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);

const NamesView: React.FC<NamesViewProps> = ({
  createRequestId,
  onOpenLegacy,
  embedded = false,
  onOpenShape,
}) => {
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
  const [activeModal, setActiveModal] = useState<ActiveModal>("none");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [typeFilters, setTypeFilters] = useState<Set<string>>(new Set());
  const [typeFilterOpen, setTypeFilterOpen] = useState<boolean>(false);
  const [addressColumnWidth, setAddressColumnWidth] = useState<number>(320);
  const [resizingAddressColumn, setResizingAddressColumn] = useState<boolean>(false);
  const addressResizeRef = useRef<{ startX: number; startWidth: number } | null>(null);
  const typeFilterRef = useRef<HTMLDivElement | null>(null);
  const anyModalOpenRef = useRef(false);

  const [editState, setEditState] = useState<EditState>({
    record: null,
    editType: "Range",
    name: "",
    address: "",
    listValues: "",
    fallbackSheet: "",
    caseTransform: "none",
  });
  const [createState, setCreateState] = useState<CreateState>({
    ...buildDefaultCreateState(),
  });
  const [moveState, setMoveState] = useState<MoveState>({
    record: null,
    address: "",
    fallbackSheet: "",
  });
  const [deleteState, setDeleteState] = useState<DeleteState>({
    record: null,
    deleteName: true,
    deleteValues: false,
  });
  const [bulkState, setBulkState] = useState<BulkState>({
    mode: "Prefix",
    value: "",
    replaceWith: "",
    caseTransform: "none",
    deleteName: true,
    deleteValues: false,
  });
  const debouncedBulkValue = useDebounce(bulkState.value, 150);
  const debouncedBulkReplaceWith = useDebounce(bulkState.replaceWith, 150);
  anyModalOpenRef.current = activeModal !== "none";

  const openCreateDialog = async () => {
    let selectionSheet = "";
    try {
      const selection = await getCurrentSelectionAddress();
      if (!isInternalWorkbookSheetName(selection.sheet)) {
        selectionSheet = selection.sheet;
      }
    } catch {
      // best effort only
    }
    setCreateState({
      ...buildDefaultCreateState(),
      fallbackSheet: selectionSheet,
      scope: selectionSheet,
    });
    setActiveModal("create");
  };

  const closeCreateDialog = () => {
    setCreateState(buildDefaultCreateState());
    setActiveModal("none");
  };

  const load = useCallback(
    async (showStatus = true) => {
      setLoading(true);
      try {
        const loaded = await getNamedRanges();
        setRows(loaded);
        setSelectedIds(new Set());
        if (showStatus) {
          setStatusType("success");
          setStatus(`Loaded ${loaded.length} name/shape record(s).`);
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
      if (anyModalOpenRef.current) {
        return;
      }
      void load(false);
    };
    window.addEventListener("focus", handleFocus);
    return () => {
      window.removeEventListener("focus", handleFocus);
    };
  }, [load]);

  useEffect(() => {
    if (createRequestId > 0) {
      void openCreateDialog();
    }
  }, [createRequestId]);

  useEffect(() => {
    if (!resizingAddressColumn) {
      return undefined;
    }
    const handleMouseMove = (event: MouseEvent) => {
      const resizeState = addressResizeRef.current;
      if (!resizeState) {
        return;
      }
      const delta = event.clientX - resizeState.startX;
      setAddressColumnWidth(
        Math.max(
          ADDRESS_COLUMN_MIN_WIDTH,
          Math.min(ADDRESS_COLUMN_MAX_WIDTH, resizeState.startWidth + delta)
        )
      );
    };
    const handleMouseUp = () => {
      setResizingAddressColumn(false);
      addressResizeRef.current = null;
    };
    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [resizingAddressColumn]);

  useEffect(() => {
    if (!typeFilterOpen) {
      return undefined;
    }

    const handlePointerDown = (event: MouseEvent) => {
      if (typeFilterRef.current?.contains(event.target as Node)) {
        return;
      }
      setTypeFilterOpen(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setTypeFilterOpen(false);
      }
    };

    window.addEventListener("mousedown", handlePointerDown);
    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("mousedown", handlePointerDown);
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [typeFilterOpen]);

  const visibleRows = useMemo(() => {
    const token = search.trim().toLowerCase();
    return rows
      .filter((row) => (scopeFilter === "all" ? true : row.scopeType === scopeFilter))
      .filter((row) => (typeFilters.size === 0 ? true : typeFilters.has(row.type)))
      .filter((row) =>
        token
          ? `${row.name} ${row.address} ${row.sheet} ${row.scope} ${row.type}`
              .toLowerCase()
              .includes(token)
          : true
      )
      .sort((a, b) => {
        const compare = a[sortColumn].localeCompare(b[sortColumn], undefined, {
          sensitivity: "base",
        });
        return sortDirection === "asc" ? compare : -compare;
      });
  }, [rows, scopeFilter, typeFilters, search, sortColumn, sortDirection]);

  const availableTypes = useMemo(
    () => Array.from(new Set(rows.map((row) => row.type))).sort((a, b) => a.localeCompare(b)),
    [rows]
  );
  const typeFilterLabel =
    typeFilters.size === 0 ? "All Types" : `Types (${typeFilters.size.toString()})`;

  const selectableRows = useMemo(
    () => visibleRows.filter((row) => row.kind === "NamedRange" && row.isRange),
    [visibleRows]
  );
  const allSelected =
    selectableRows.length > 0 && selectableRows.every((row) => selectedIds.has(row.id));
  const createFunctionArgsPreview = useMemo(
    () => parseFunctionArgs(createState.lambdaArgs),
    [createState.lambdaArgs]
  );

  const bulkPreview = useMemo(() => {
    const selected = rows.filter((row) => row.kind === "NamedRange" && selectedIds.has(row.id));
    const occupiedByScope = new Map<string, Set<string>>();
    rows
      .filter((row) => row.kind === "NamedRange" && !selectedIds.has(row.id))
      .forEach((row) => {
        const scopeKey = buildScopeKey(row);
        const occupied = occupiedByScope.get(scopeKey) ?? new Set<string>();
        occupied.add(normalizeNameKey(row.name));
        occupiedByScope.set(scopeKey, occupied);
      });

    return selected.map((row) => {
      const scopeKey = buildScopeKey(row);
      const occupied = occupiedByScope.get(scopeKey) ?? new Set<string>();
      if (bulkState.mode === "Delete") {
        return {
          id: row.id,
          oldName: row.name,
          requestedName: row.name,
          newName: row.name,
          row,
          scopeKey,
        };
      }
      let requestedName = row.name;
      if (bulkState.mode === "Prefix") requestedName = `${debouncedBulkValue}${row.name}`;
      if (bulkState.mode === "Suffix") requestedName = `${row.name}${debouncedBulkValue}`;
      if (bulkState.mode === "Replace")
        requestedName = row.name.replace(debouncedBulkValue, debouncedBulkReplaceWith);
      requestedName = applyCase(requestedName, bulkState.caseTransform);
      const newName = buildUniqueScopedName(requestedName, occupied, row.name);
      occupied.add(normalizeNameKey(newName));
      occupiedByScope.set(scopeKey, occupied);
      return { id: row.id, oldName: row.name, requestedName, newName, row, scopeKey };
    });
  }, [
    bulkState.caseTransform,
    bulkState.mode,
    debouncedBulkReplaceWith,
    debouncedBulkValue,
    rows,
    selectedIds,
  ]);

  const toggleSort = (column: SortColumn) => {
    if (sortColumn === column) {
      setSortDirection((prev) => (prev === "asc" ? "desc" : "asc"));
      return;
    }
    setSortColumn(column);
    setSortDirection("asc");
  };

  const toggleTypeFilter = (type: string) => {
    setTypeFilters((prev) => {
      const next = new Set(prev);
      if (next.has(type)) {
        next.delete(type);
      } else {
        next.add(type);
      }
      return next;
    });
  };

  const startAddressResize = (event: React.MouseEvent<HTMLSpanElement>) => {
    event.preventDefault();
    event.stopPropagation();
    addressResizeRef.current = {
      startX: event.clientX,
      startWidth: addressColumnWidth,
    };
    setResizingAddressColumn(true);
  };

  const captureSelection = async (
    setter: (selection: { address: string; sheet: string }) => void
  ) => {
    try {
      const selection = await getCurrentSelectionAddress();
      if (isInternalWorkbookSheetName(selection.sheet)) {
        setStatusType("error");
        setStatus("Select a cell on a workbook sheet before capturing selection.");
        return;
      }
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
    if (!name) {
      setStatusType("error");
      setStatus("Name is required.");
      return;
    }
    if (
      createState.scopeType === "Worksheet" &&
      !createState.scope.trim() &&
      !createState.fallbackSheet.trim()
    ) {
      setStatusType("error");
      setStatus("Worksheet scope is required for worksheet-level names.");
      return;
    }

    let definition = createState.address.trim();
    let referenceType: "Reference" | "Formula" = "Reference";
    let effectiveFallbackSheet = createState.fallbackSheet.trim();

    if (createState.createType === "Range") {
      if (!definition) {
        setStatusType("error");
        setStatus("Address is required for range names.");
        return;
      }
      if (!definition.includes("!") && !effectiveFallbackSheet) {
        try {
          const selection = await getCurrentSelectionAddress();
          effectiveFallbackSheet = selection.sheet;
          setCreateState((prev) => ({
            ...prev,
            fallbackSheet: prev.fallbackSheet || selection.sheet,
          }));
        } catch {
          // continue with explicit validation below
        }
      }
      if (!definition.includes("!") && !effectiveFallbackSheet) {
        setStatusType("error");
        setStatus(
          "Use a qualified reference (for example, Sheet1!A1) or capture a selection first."
        );
        return;
      }
    }

    if (createState.createType === "Function") {
      try {
        definition = buildLambdaFormula(createState.functionBody, createState.lambdaArgs);
        referenceType = "Formula";
      } catch (error) {
        setStatusType("error");
        setStatus(err(error));
        return;
      }
    }

    if (createState.createType === "List") {
      try {
        definition = buildNamedListFormula(createState.listValues);
        referenceType = "Formula";
      } catch (error) {
        setStatusType("error");
        setStatus(err(error));
        return;
      }
    }

    await runSubmit(
      async () => {
        const scope =
          createState.scopeType === "Worksheet"
            ? createState.scope.trim() || effectiveFallbackSheet
            : createState.scope.trim();
        if (createState.scopeType === "Worksheet" && isInternalWorkbookSheetName(scope)) {
          throw new Error(
            "Choose a workbook worksheet scope (internal __WBM sheets are not allowed)."
          );
        }
        if (
          createState.createType === "Range" &&
          !definition.includes("!") &&
          isInternalWorkbookSheetName(effectiveFallbackSheet)
        ) {
          throw new Error(
            "Use a workbook sheet selection or a qualified reference like Sheet1!A1."
          );
        }
        await addNamedRange(
          createState.scopeType,
          scope,
          name,
          definition,
          effectiveFallbackSheet,
          referenceType
        );
        setCreateState((prev) => ({
          ...prev,
          ...buildDefaultCreateState(),
        }));
        setActiveModal("none");
      },
      `${createState.createType} name created.`,
      "Create failed"
    );
  };

  const submitEdit = async () => {
    const record = editState.record;
    if (!record) return;
    let editAddress = editState.address.trim();
    let referenceType: "Reference" | "Formula" = record.isRange ? "Reference" : "Formula";

    if (editState.editType === "List") {
      try {
        editAddress = buildNamedListFormula(editState.listValues);
        referenceType = "Formula";
      } catch (error) {
        setStatusType("error");
        setStatus(err(error));
        return;
      }
    }

    await runSubmit(
      async () => {
        await updateNamedRange(
          record.scopeType,
          record.scope,
          record.name,
          applyCase(editState.name.trim(), editState.caseTransform),
          editAddress,
          editState.fallbackSheet,
          referenceType
        );
        setEditState({
          record: null,
          editType: "Range",
          name: "",
          address: "",
          listValues: "",
          fallbackSheet: "",
          caseTransform: "none",
        });
        setActiveModal("none");
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
        await moveNamedRange(
          record.scopeType,
          record.scope,
          record.name,
          moveState.address,
          moveState.fallbackSheet
        );
        setMoveState({
          record: null,
          address: "",
          fallbackSheet: "",
        });
        setActiveModal("none");
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
          record: null,
          deleteName: true,
          deleteValues: false,
        });
        setActiveModal("none");
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
        }

        if (bulkState.mode !== "Delete") {
          const changedRows = bulkPreview.filter(
            (row) => normalizeNameKey(row.newName) !== normalizeNameKey(row.oldName)
          );
          const occupiedByScope = new Map<string, Set<string>>();
          rows
            .filter((row) => row.kind === "NamedRange")
            .forEach((row) => {
              const scopeKey = buildScopeKey(row);
              const occupied = occupiedByScope.get(scopeKey) ?? new Set<string>();
              occupied.add(normalizeNameKey(row.name));
              occupiedByScope.set(scopeKey, occupied);
            });

          const temporaryRows = changedRows.map((row, index) => {
            const occupied = occupiedByScope.get(row.scopeKey) ?? new Set<string>();
            let tempName = `__WBM_TMP_NAME_${index.toString()}__`;
            while (occupied.has(normalizeNameKey(tempName))) {
              tempName = `__WBM_TMP_NAME_${index.toString()}_${occupied.size.toString()}__`;
            }
            occupied.add(normalizeNameKey(tempName));
            occupiedByScope.set(row.scopeKey, occupied);
            return { ...row, tempName };
          });

          for (const row of temporaryRows) {
            await updateNamedRange(
              row.row.scopeType,
              row.row.scope,
              row.oldName,
              row.tempName,
              row.row.address,
              row.row.sheet
            );
          }
          for (const row of temporaryRows) {
            await updateNamedRange(
              row.row.scopeType,
              row.row.scope,
              row.tempName,
              row.newName,
              row.row.address,
              row.row.sheet
            );
          }
        }
        setActiveModal("none");
        setSelectedIds(new Set());
      },
      bulkState.mode === "Delete"
        ? `Bulk operation completed for ${bulkPreview.length} range(s).`
        : (() => {
            const adjustedCount = bulkPreview.filter(
              (row) => normalizeNameKey(row.requestedName) !== normalizeNameKey(row.newName)
            ).length;
            return adjustedCount > 0
              ? `Bulk operation completed for ${bulkPreview.length} range(s). ${adjustedCount.toString()} rename(s) were adjusted to avoid collisions.`
              : `Bulk operation completed for ${bulkPreview.length} range(s).`;
          })(),
      "Bulk operation failed"
    );
  };

  const openNameLink = async (row: NamedRangeRecord) => {
    if (row.kind === "Shape") {
      if (!onOpenShape) {
        setStatusType("error");
        setStatus(`Shape routing is unavailable for "${row.name}".`);
        return;
      }
      onOpenShape({ sheetName: row.sheet, shapeName: row.name });
      setStatusType("success");
      setStatus(`Opened shape editor for "${row.name}".`);
      return;
    }

    if (row.kind !== "NamedRange") {
      return;
    }
    if (row.isRange) {
      try {
        await selectNamedRangeAddress(row.address, row.sheet);
      } catch (error) {
        setStatusType("error");
        setStatus(`Unable to navigate to "${row.name}": ${err(error)}`);
      }
      return;
    }
    const baseFormula = (row.formula || row.address).trim();
    const normalizedFormula = baseFormula.startsWith("=") ? baseFormula : `=${baseFormula}`;
    try {
      const functionArgs =
        row.type === "Function"
          ? extractLambdaArgsFromFormula(normalizedFormula).join(",")
          : undefined;
      await openFormulaEditorPopout({
        formula: normalizedFormula,
        name: row.name,
        entryType: row.type === "Function" ? "Function" : row.type === "List" ? "List" : "Formula",
        creationMode: row.type === "Function" ? "Function" : "Formula",
        authoringMode: "Editor",
        functionArgs,
      });
      setStatusType("success");
      setStatus(`Opened formula editor pop-out for "${row.name}".`);
    } catch (error) {
      setStatusType("error");
      setStatus(`Unable to open formula editor for "${row.name}": ${err(error)}`);
    }
  };

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Names</Text>
          <Text className={shared.sectionSubtitle}>
            Manage named ranges, lists, functions, and workbook objects.
          </Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <div className={styles.toolbarPrimary}>
            <Button appearance="primary" onClick={() => void openCreateDialog()}>
              + Create Name
            </Button>
            <Button
              onClick={() => {
                setActiveModal("bulk");
              }}
              disabled={selectedIds.size === 0}
            >
              Bulk Update ({selectedIds.size})
            </Button>
            <Button onClick={() => void load()} disabled={loading}>
              Refresh
            </Button>
            <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          </div>
          <div className={styles.toolbarFilters}>
            <Select
              value={scopeFilter}
              onChange={(_, data) => setScopeFilter(data.value as "all" | ScopeType)}
            >
              <option value="all">All Scopes</option>
              <option value="Workbook">Workbook</option>
              <option value="Worksheet">Worksheet</option>
            </Select>
            <div className={styles.typeFilterWrap} ref={typeFilterRef}>
              <button
                type="button"
                className={`${styles.typeFilterSummary} ${typeFilterOpen ? styles.typeFilterButtonOpen : ""}`}
                aria-expanded={typeFilterOpen ? "true" : "false"}
                aria-controls="names-type-filter-menu"
                onClick={() => setTypeFilterOpen((prev) => !prev)}
              >
                {typeFilterLabel}
              </button>
              {typeFilterOpen ? (
                <div id="names-type-filter-menu" className={styles.typeFilterMenu}>
                  <label className={styles.typeFilterRow}>
                    <input
                      type="checkbox"
                      checked={typeFilters.size === 0}
                      onChange={() => setTypeFilters(new Set())}
                    />
                    All Types
                  </label>
                  <div className={styles.typeFilterDivider} />
                  {availableTypes.map((type) => (
                    <label key={type} className={styles.typeFilterRow}>
                      <input
                        type="checkbox"
                        checked={typeFilters.has(type)}
                        onChange={() => toggleTypeFilter(type)}
                      />
                      {type}
                    </label>
                  ))}
                </div>
              ) : null}
            </div>
            <Input
              placeholder="Search names, address, scope..."
              value={search}
              onChange={(_, data) => setSearch(data.value)}
            />
          </div>
        </div>
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <colgroup>
            <col width={44} />
            <col width={180} />
            <col width={200} />
            <col width={addressColumnWidth} />
            <col width={140} />
            <col width={140} />
            <col width={180} />
          </colgroup>
          <thead>
            <tr>
              <th className={styles.headCell}>
                <input
                  type="checkbox"
                  checked={allSelected}
                  aria-label={
                    allSelected
                      ? "Clear selection for all visible named ranges"
                      : "Select all visible named ranges"
                  }
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
                <th
                  key={column}
                  className={`${styles.headCell} ${column === "address" ? styles.addressHeadCell : ""}`}
                >
                  <button
                    type="button"
                    className={styles.sortBtn}
                    onClick={() => toggleSort(column)}
                  >
                    {column[0].toUpperCase() + column.slice(1)}
                    {sortColumn === column ? (sortDirection === "asc" ? " ▲" : " ▼") : ""}
                  </button>
                  {column === "address" ? (
                    <span
                      className={styles.columnResizeHandle}
                      onMouseDown={startAddressResize}
                      title="Drag to resize Address column"
                    />
                  ) : null}
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
                    disabled={row.kind !== "NamedRange" || !row.isRange}
                    aria-label={`Select named range ${row.name}`}
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
                      appearance="subtle"
                      size="small"
                      icon={<Edit20Regular />}
                      aria-label={`Edit ${row.name}`}
                      title={`Edit ${row.name}`}
                      disabled={row.kind !== "NamedRange"}
                      onClick={() => {
                        setEditState({
                          record: row,
                          editType:
                            row.type === "List" ? "List" : row.isRange ? "Range" : "Formula",
                          name: row.name,
                          address: row.isRange ? row.address : row.formula || row.address,
                          listValues:
                            row.type === "List"
                              ? parseNamedListFormula(row.formula || row.address).join("\n")
                              : "",
                          fallbackSheet: row.sheet,
                          caseTransform: "none",
                        });
                        setActiveModal("edit");
                      }}
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<ArrowMove20Regular />}
                      aria-label={`Move ${row.name}`}
                      title={`Move ${row.name}`}
                      disabled={!row.isRange}
                      onClick={() => {
                        setMoveState({
                          record: row,
                          address: row.address,
                          fallbackSheet: row.sheet,
                        });
                        setActiveModal("move");
                      }}
                    />
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<Delete20Regular />}
                      aria-label={`Delete ${row.name}`}
                      title={`Delete ${row.name}`}
                      disabled={row.kind !== "NamedRange"}
                      onClick={() => {
                        setDeleteState({
                          record: row,
                          deleteName: true,
                          deleteValues: false,
                        });
                        setActiveModal("delete");
                      }}
                    />
                  </div>
                </td>
                <td className={styles.cell}>
                  {row.kind === "NamedRange" || row.kind === "Shape" ? (
                    <button
                      type="button"
                      className={styles.linkBtn}
                      onClick={() => void openNameLink(row)}
                      title={
                        row.kind === "Shape"
                          ? `Open shape editor for ${row.name}`
                          : row.isRange
                            ? `Go to ${row.address}`
                            : `Formula-based name (${row.type})`
                      }
                    >
                      {row.name}
                    </button>
                  ) : (
                    row.name
                  )}
                </td>
                <td className={styles.cell}>
                  <span className={styles.cellEllipsis} title={row.address}>
                    {row.address}
                  </span>
                </td>
                <td className={styles.cell}>
                  <span className={styles.cellEllipsis} title={row.sheet}>
                    {row.sheet}
                  </span>
                </td>
                <td className={styles.cell}>
                  <span className={styles.cellEllipsis} title={row.scope}>
                    {row.scope}
                  </span>
                </td>
                <td className={styles.cell}>
                  <span className={styles.cellEllipsis} title={row.type}>
                    {row.type}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <Text className={shared.mutedText}>{`Showing ${visibleRows.length} of ${rows.length}`}</Text>
      {status ? <Text className={statusClass}>{status}</Text> : null}

      {activeModal === "create" ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Create Name</Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Name</Text>
                <Input
                  value={createState.name}
                  onChange={(_, data) => setCreateState((p) => ({ ...p, name: data.value }))}
                />
              </div>
              {createState.createType === "Range" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>Address / Selected Range</Text>
                  <Input
                    value={createState.address}
                    onChange={(_, data) => setCreateState((p) => ({ ...p, address: data.value }))}
                  />
                </div>
              ) : null}
              <div>
                <Text className={shared.mutedText}>Scope</Text>
                <Select
                  value={createState.scopeType}
                  onChange={(_, data) =>
                    setCreateState((p) => ({ ...p, scopeType: data.value as ScopeType }))
                  }
                >
                  <option value="Workbook">Workbook</option>
                  <option value="Worksheet">Worksheet</option>
                </Select>
              </div>
              <div>
                <Text className={shared.mutedText}>Type</Text>
                <Select
                  value={createState.createType}
                  onChange={(_, data) =>
                    setCreateState((p) => ({ ...p, createType: data.value as CreateEntryType }))
                  }
                >
                  <option value="Range">Range</option>
                  <option value="Function">Function (LAMBDA)</option>
                  <option value="List">List (Named Array)</option>
                </Select>
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
              {createState.createType === "Function" ? (
                <>
                  <div className={styles.full}>
                    <Text className={shared.mutedText}>Arguments (comma-separated)</Text>
                    <Input
                      value={createState.lambdaArgs}
                      placeholder="table, lookupValue"
                      onChange={(_, data) =>
                        setCreateState((p) => ({ ...p, lambdaArgs: data.value }))
                      }
                    />
                  </div>
                  <div className={styles.full}>
                    <Text className={shared.mutedText}>Arguments Preview</Text>
                    {createFunctionArgsPreview.length > 0 ? (
                      <div className={styles.chipRow}>
                        {createFunctionArgsPreview.map((arg, index) => (
                          <span key={`${arg}-${index.toString()}`} className={styles.chip}>
                            {arg}
                          </span>
                        ))}
                      </div>
                    ) : (
                      <Text className={shared.mutedText}>
                        No arguments set. The function will save as `LAMBDA(calculation)`.
                      </Text>
                    )}
                  </div>
                  <div className={styles.full}>
                    <Text className={shared.mutedText}>Function Definition</Text>
                    <textarea
                      className={styles.multilineInput}
                      value={createState.functionBody}
                      onChange={(event) =>
                        setCreateState((p) => ({ ...p, functionBody: event.target.value }))
                      }
                      placeholder="Enter a formula body, or paste a full LAMBDA formula."
                    />
                  </div>
                </>
              ) : null}
              {createState.createType === "List" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>List Values</Text>
                  <textarea
                    className={styles.multilineInput}
                    value={createState.listValues}
                    onChange={(event) =>
                      setCreateState((p) => ({ ...p, listValues: event.target.value }))
                    }
                    placeholder={
                      "Enter values separated by commas or new lines.\nExample:\nOpen\nIn Progress\nClosed"
                    }
                  />
                  <Text className={shared.mutedText}>
                    Saved as a named array formula so it can be used in data validation lists.
                  </Text>
                </div>
              ) : null}
            </div>
            <div className={styles.modalActions}>
              {createState.createType === "Range" ? (
                <Button
                  onClick={() =>
                    void captureSelection((selection) =>
                      setCreateState((p) => ({
                        ...p,
                        address: selection.address,
                        fallbackSheet: selection.sheet,
                        scope: p.scopeType === "Worksheet" ? selection.sheet : p.scope,
                      }))
                    )
                  }
                >
                  Use Selection
                </Button>
              ) : null}
              <Button onClick={closeCreateDialog}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={submitting}
                onClick={() => void submitCreate()}
              >
                Create
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {activeModal === "edit" && editState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>
              {editState.editType === "List" ? "Edit Named List" : "Edit Named Range"}
            </Text>
            <div className={styles.modalGrid}>
              <div>
                <Text className={shared.mutedText}>Name</Text>
                <Input
                  value={editState.name}
                  onChange={(_, data) => setEditState((p) => ({ ...p, name: data.value }))}
                />
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
              {editState.editType === "List" ? (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>List Values</Text>
                  <textarea
                    className={styles.multilineInput}
                    value={editState.listValues}
                    onChange={(event) =>
                      setEditState((p) => ({ ...p, listValues: event.target.value }))
                    }
                    placeholder={"Open\nIn Progress\nClosed"}
                  />
                </div>
              ) : (
                <div className={styles.full}>
                  <Text className={shared.mutedText}>
                    {editState.editType === "Formula" ? "Definition" : "Address"}
                  </Text>
                  <Input
                    value={editState.address}
                    onChange={(_, data) => setEditState((p) => ({ ...p, address: data.value }))}
                  />
                </div>
              )}
            </div>
            <div className={styles.modalActions}>
              {editState.editType === "Range" ? (
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
              ) : null}
              <Button
                onClick={() => {
                  setEditState({
                    record: null,
                    editType: "Range",
                    name: "",
                    address: "",
                    listValues: "",
                    fallbackSheet: "",
                    caseTransform: "none",
                  });
                  setActiveModal("none");
                }}
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

      {activeModal === "move" && moveState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Move Named Range</Text>
            <Text className={shared.mutedText}>Address</Text>
            <Input
              value={moveState.address}
              onChange={(_, data) => setMoveState((p) => ({ ...p, address: data.value }))}
            />
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
              <Button
                onClick={() => {
                  setMoveState({ record: null, address: "", fallbackSheet: "" });
                  setActiveModal("none");
                }}
              >
                Cancel
              </Button>
              <Button appearance="primary" disabled={submitting} onClick={() => void submitMove()}>
                Move
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {activeModal === "delete" && deleteState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Delete Named Range</Text>
            <label>
              <input
                type="checkbox"
                checked={deleteState.deleteName}
                onChange={(event) =>
                  setDeleteState((p) => ({ ...p, deleteName: event.target.checked }))
                }
              />{" "}
              Delete name
            </label>
            <label>
              <input
                type="checkbox"
                checked={deleteState.deleteValues}
                onChange={(event) =>
                  setDeleteState((p) => ({ ...p, deleteValues: event.target.checked }))
                }
              />{" "}
              Delete range values
            </label>
            <div className={styles.modalActions}>
              <Button
                onClick={() => {
                  setDeleteState({
                    record: null,
                    deleteName: true,
                    deleteValues: false,
                  });
                  setActiveModal("none");
                }}
              >
                Cancel
              </Button>
              <Button
                appearance="primary"
                disabled={submitting}
                onClick={() => void submitDelete()}
              >
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {activeModal === "bulk" ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={shared.cardTitle}>Bulk Update Named Ranges</Text>
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
                        onChange={(_, data) =>
                          setBulkState((p) => ({ ...p, replaceWith: data.value }))
                        }
                      />
                    </div>
                  ) : null}
                </>
              ) : (
                <div className={styles.full}>
                  <label>
                    <input
                      type="checkbox"
                      checked={bulkState.deleteName}
                      onChange={(event) =>
                        setBulkState((p) => ({ ...p, deleteName: event.target.checked }))
                      }
                    />{" "}
                    Delete names
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={bulkState.deleteValues}
                      onChange={(event) =>
                        setBulkState((p) => ({ ...p, deleteValues: event.target.checked }))
                      }
                    />{" "}
                    Delete range values
                  </label>
                </div>
              )}
            </div>
            <Text className={shared.mutedText}>{`Preview rows: ${bulkPreview.length}`}</Text>
            <div className={styles.modalActions}>
              <Button onClick={() => setActiveModal("none")}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={submitting || bulkPreview.length === 0}
                onClick={() => void submitBulk()}
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

export default NamesView;
