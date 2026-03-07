import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import Editor, { OnMount, loader } from "@monaco-editor/react";
import type * as Monaco from "monaco-editor";
import {
  FormulaEvaluationResult,
  TableRecord,
  applyFormulaToActiveCell,
  evaluateFormula,
  getActiveCellFormulaState,
  getCurrentSelectionAddress,
  getNamedRanges,
  getTables,
  openFormulaEditorPopout,
  saveNamedFunction,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

/* global Office */

const MONACO_LANGUAGE_ID = "wbm-formula";
const DIALOG_RPC_CHANNEL = "wbm-formula-dialog-rpc";

interface DialogEvalRequestMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "eval-request";
  requestId: string;
  formula: string;
}

interface DialogEvalResponseMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "eval-response";
  requestId: string;
  ok: boolean;
  result?: FormulaEvaluationResult;
  error?: string;
}

interface DialogReadyMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "ready";
}

interface DialogOpenFormulaMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "open-formula";
  formula: string;
  name?: string;
  entryType?: "Formula" | "Function" | "List";
  functionArgs?: string;
  description?: string;
  creationMode?: "Formula" | "Function";
  authoringMode?: "Editor" | "Wizard";
  wizardTemplate?: "LAMBDA" | "LET";
  wizardArgs?: string;
  wizardReturnExpression?: string;
  wizardVariables?: Array<{ name: string; expression: string }>;
}

let monacoLoaderConfigured = false;
const configureMonacoLoader = () => {
  if (monacoLoaderConfigured || typeof window === "undefined") {
    return;
  }
  // Keep Monaco local to the add-in origin so Office host CSP/AppDomain rules do not block editor boot.
  loader.config({
    paths: {
      vs: `${window.location.origin}/monaco/vs`,
    },
  });
  monacoLoaderConfigured = true;
};
configureMonacoLoader();

const FUNCTION_SUGGESTIONS = [
  "ABS",
  "AVERAGE",
  "CHOOSE",
  "COUNT",
  "FILTER",
  "IF",
  "IFS",
  "INDEX",
  "INDIRECT",
  "LAMBDA",
  "LET",
  "LOOKUP",
  "MATCH",
  "MAX",
  "MIN",
  "OFFSET",
  "SEQUENCE",
  "SORT",
  "SORTBY",
  "SUBTOTAL",
  "SUM",
  "SUMIF",
  "SUMIFS",
  "SUMPRODUCT",
  "SWITCH",
  "TEXT",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTJOIN",
  "TEXTSPLIT",
  "UNIQUE",
  "VLOOKUP",
  "XLOOKUP",
  "XMATCH",
] as const;

interface FormulaMonacoViewProps {
  isPopout: boolean;
  onOpenLegacy: () => void;
}

export interface FormulaViewHandle {
  openOnly: () => Promise<void>;
  openAndPull: () => Promise<void>;
  pull: () => Promise<void>;
  apply: () => Promise<void>;
  beautify: () => Promise<void>;
  insertSelection: () => Promise<void>;
}

interface ActiveCellState {
  sheet: string;
  address: string;
  formula: string;
  hasFormula: boolean;
}

type FormulaCreationMode = "Formula" | "Function";
type FormulaAuthoringMode = "Editor" | "Wizard";
type WizardTemplate = "LAMBDA" | "LET";
type FormulaSubTab = "metadata" | "lambda-test" | "editor" | "wizard" | "live-output";

interface WizardLetVariable {
  id: string;
  name: string;
  expression: string;
}

const useStyles = makeStyles({
  root: { display: "grid", gap: "16px" },
  card: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    display: "grid",
    gap: "12px",
  },
  editorShell: {
    position: "relative",
  },
  overlayTop: {
    position: "absolute",
    top: "8px",
    right: "8px",
    zIndex: 6,
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  overlayBottom: {
    position: "absolute",
    left: "8px",
    right: "8px",
    bottom: "8px",
    zIndex: 6,
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  overlayCluster: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "rgba(255, 255, 255, 0.95)",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "8px",
    display: "grid",
    gap: "6px",
  },
  overlayTitle: {
    fontSize: "11px",
    fontWeight: 700,
    color: MODERN_TOKENS.colorTextMuted,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  actionRow: {
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
    alignItems: "center",
  },
  slicer: {
    display: "inline-flex",
    borderRadius: "999px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "2px",
    gap: "2px",
  },
  slicerBtn: {
    borderRadius: "999px",
    border: "none",
    backgroundColor: "transparent",
    color: MODERN_TOKENS.colorTextMuted,
    padding: "4px 10px",
    fontSize: "12px",
    fontWeight: 600,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  slicerBtnActive: {
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    boxShadow: "0 1px 2px rgba(17,24,39,0.08)",
  },
  editorWrap: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    overflow: "hidden",
    minHeight: "220px",
    backgroundColor: "#fff",
  },
  editorLoading: {
    minHeight: "220px",
    display: "grid",
    alignItems: "center",
    justifyItems: "center",
    backgroundColor: "#fff",
  },
  fallbackEditor: {
    width: "100%",
    minHeight: "220px",
    border: "none",
    outline: "none",
    resize: "vertical",
    padding: "42px 12px 128px",
    boxSizing: "border-box",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "13px",
    lineHeight: "1.4",
    color: MODERN_TOKENS.colorText,
    backgroundColor: "#fff",
  },
  liveTestHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  outputWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "6px",
    overflow: "auto",
    backgroundColor: "#fff",
    minHeight: "84px",
  },
  outputTable: { width: "100%", borderCollapse: "collapse", fontSize: "12px" },
  outputCell: { borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, padding: "8px 10px" },
  autoCaptureRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "12px",
    flexWrap: "wrap",
  },
  checkboxLabel: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
  },
  modeRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    flexWrap: "wrap",
  },
  subTabBar: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  subTabBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
    fontWeight: 600,
    padding: "6px 10px",
    cursor: "pointer",
    textTransform: "none",
  },
  subTabBtnActive: {
    backgroundColor: "#E8F4EA",
    border: "1px solid #B9D9BD",
    color: MODERN_TOKENS.colorBrandStrong,
  },
  subTabPanel: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "10px 12px",
    display: "grid",
    gap: "10px",
  },
  actionDock: {
    position: "sticky",
    bottom: 0,
    zIndex: 7,
    marginTop: "8px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "rgba(255,255,255,0.96)",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  modeSelect: { minWidth: "160px" },
  functionMetaGrid: {
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  fullRow: { gridColumn: "1 / -1" },
  wizardCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "12px",
    display: "grid",
    gap: "10px",
  },
  wizardVarRow: {
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1.2fr) auto",
    alignItems: "center",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  wizardPreview: {
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "12px",
    lineHeight: "1.35",
    padding: "8px",
    margin: 0,
    maxHeight: "180px",
    overflow: "auto",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  lambdaTestGrid: {
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    "@media (max-width: 900px)": {
      gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    },
    "@media (max-width: 620px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
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
    width: "min(520px, 100%)",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "10px",
  },
  modalField: { display: "grid", gap: "6px" },
  modalLabel: {
    fontSize: "12px",
    fontWeight: 600,
    color: MODERN_TOKENS.colorTextMuted,
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: "6px" },
  chip: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "999px",
    padding: "2px 10px",
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    backgroundColor: "#F9FAFB",
  },
  chipBtn: {
    border: "none",
    backgroundColor: "transparent",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "11px",
    cursor: "pointer",
    marginLeft: "4px",
    padding: 0,
  },
  argsInput: {
    minWidth: "220px",
    flex: 1,
  },
  invocationText: {
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    fontFamily: "Consolas, 'Courier New', monospace",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  formulaPreview: {
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "12px",
    lineHeight: "1.4",
    padding: "8px",
    margin: 0,
    maxHeight: "120px",
    overflow: "auto",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: "8px", flexWrap: "wrap" },
});

const normalizeError = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);
const LAMBDA_IDENTIFIER_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;
const EXCEL_NAME_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;
const CELL_REFERENCE_LIKE_PATTERN = /^[A-Za-z]{1,3}\d+$/i;

const parseLambdaArgs = (input: string): string[] =>
  input
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);

const mergeArgsWithDraft = (argsInput: string, draftInput: string): string[] => {
  const args = parseLambdaArgs(argsInput);
  const draft = draftInput.trim();
  if (!draft) {
    return args;
  }
  if (!args.some((item) => item.toUpperCase() === draft.toUpperCase())) {
    args.push(draft);
  }
  return args;
};

const stringifyArgs = (args: string[]): string => args.join(",");

const getFirstErrorCell = (result: FormulaEvaluationResult): string => {
  for (let rowIndex = 0; rowIndex < result.values.length; rowIndex += 1) {
    const row = result.values[rowIndex];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      if (result.valueTypes[rowIndex]?.[colIndex] === "Error") {
        const value = row[colIndex];
        return value === null || value === undefined || value === "" ? "#ERROR!" : String(value);
      }
    }
  }
  return "#ERROR!";
};

const buildNamedFunctionFormula = (formulaInput: string, argsInput: string): string => {
  const normalized = formulaInput.trim();
  if (!normalized) {
    throw new Error("Formula is required.");
  }

  const formulaWithoutEquals = normalized.replace(/^=/, "").trim();
  if (/^LAMBDA\s*\(/i.test(formulaWithoutEquals)) {
    return `=${formulaWithoutEquals}`;
  }

  const args = parseLambdaArgs(argsInput);
  const seen = new Set<string>();
  args.forEach((arg) => {
    if (!LAMBDA_IDENTIFIER_PATTERN.test(arg)) {
      throw new Error(`Invalid function argument "${arg}".`);
    }
    const key = arg.toUpperCase();
    if (seen.has(key)) {
      throw new Error(`Duplicate function argument "${arg}".`);
    }
    seen.add(key);
  });

  return args.length > 0
    ? `=LAMBDA(${args.join(",")},${formulaWithoutEquals})`
    : `=LAMBDA(${formulaWithoutEquals})`;
};

const splitTopLevelComma = (input: string): string[] => {
  const parts: string[] = [];
  let depth = 0;
  let inString = false;
  let current = "";
  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];
    if (ch === '"') {
      inString = !inString;
      current += ch;
      continue;
    }
    if (!inString) {
      if (ch === "(") {
        depth += 1;
      } else if (ch === ")") {
        depth = Math.max(0, depth - 1);
      } else if (ch === "," && depth === 0) {
        parts.push(current.trim());
        current = "";
        continue;
      }
    }
    current += ch;
  }
  if (current.trim()) {
    parts.push(current.trim());
  }
  return parts;
};

const extractLambdaArgsFromFormula = (formulaInput: string): string[] => {
  const normalized = formulaInput.trim().replace(/^=/, "").trim();
  if (!/^LAMBDA\s*\(/i.test(normalized)) {
    return [];
  }
  const start = normalized.indexOf("(");
  const end = normalized.lastIndexOf(")");
  if (start < 0 || end <= start + 1) {
    return [];
  }
  const inner = normalized.slice(start + 1, end);
  const segments = splitTopLevelComma(inner);
  if (segments.length <= 1) {
    return [];
  }
  return segments.slice(0, -1).filter((segment) => LAMBDA_IDENTIFIER_PATTERN.test(segment));
};

const buildLambdaInvokeFormula = (lambdaFormula: string, rawInputs: string[]): string => {
  const normalized = lambdaFormula.trim();
  if (!/^=\s*LAMBDA\s*\(/i.test(normalized)) {
    throw new Error("A valid LAMBDA formula is required for input testing.");
  }
  const lambdaBody = normalized.replace(/^=/, "").trim();
  const callArgs = rawInputs.map((item) => (item.trim().length > 0 ? item.trim() : "0"));
  return `=LET(__wbmFn,${lambdaBody},__wbmFn(${callArgs.join(",")}))`;
};

const prettyFormatFormula = (input: string): string => {
  const trimmed = input.trim();
  if (!trimmed) return "";
  const formula = trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
  let depth = 0;
  let inString = false;
  let out = "";
  for (let i = 0; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"') {
      inString = !inString;
      out += ch;
      continue;
    }
    if (inString) {
      out += ch;
      continue;
    }
    if (ch === "(") {
      depth += 1;
      out += "(\n" + "  ".repeat(depth);
      continue;
    }
    if (ch === ")") {
      depth = Math.max(0, depth - 1);
      out += "\n" + "  ".repeat(depth) + ")";
      continue;
    }
    if (ch === ",") {
      out += ",\n" + "  ".repeat(depth);
      continue;
    }
    out += ch;
  }
  return out;
};

const getTokenBeforeCursor = (text: string, cursorOffset: number): string => {
  const before = text.slice(0, cursorOffset);
  const tokenMatch = before.match(/([A-Za-z_\\][A-Za-z0-9_.\\]*)$/);
  return tokenMatch ? tokenMatch[1] : "";
};

const collectLetLambdaLocals = (text: string, cursorOffset: number): string[] => {
  interface ParseFrame {
    name: string;
    args: string[];
    current: string;
  }

  const isIdentifier = (value: string) => /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(value.trim());
  const beforeCursor = text.slice(0, cursorOffset);
  const stack: ParseFrame[] = [];
  let token = "";
  let inString = false;

  const appendToCurrent = (ch: string) => {
    if (stack.length > 0) {
      stack[stack.length - 1].current += ch;
    }
  };

  for (let i = 0; i < beforeCursor.length; i += 1) {
    const ch = beforeCursor[i];
    if (ch === '"') {
      inString = !inString;
      appendToCurrent(ch);
      token = "";
      continue;
    }
    if (inString) {
      appendToCurrent(ch);
      continue;
    }
    if (/[A-Za-z0-9_.\\]/.test(ch)) {
      if (stack.length === 0) {
        token += ch;
      } else {
        appendToCurrent(ch);
      }
      continue;
    }
    if (ch === "(") {
      stack.push({ name: token.toUpperCase(), args: [], current: "" });
      token = "";
      continue;
    }
    if (ch === ",") {
      if (stack.length > 0) {
        const top = stack[stack.length - 1];
        top.args.push(top.current.trim());
        top.current = "";
      }
      token = "";
      continue;
    }
    if (ch === ")") {
      if (stack.length > 0) {
        const top = stack[stack.length - 1];
        top.args.push(top.current.trim());
        stack.pop();
      }
      token = "";
      continue;
    }
    appendToCurrent(ch);
    token = "";
  }

  const locals: string[] = [];
  const addLocal = (candidate: string) => {
    const cleaned = candidate.trim();
    if (!isIdentifier(cleaned) || cleaned.includes(".")) {
      return;
    }
    if (!locals.some((item) => item.toUpperCase() === cleaned.toUpperCase())) {
      locals.push(cleaned);
    }
  };

  stack.forEach((frame) => {
    const currentIndex = frame.args.length;
    if (frame.name === "LET") {
      for (let i = 0; i + 1 < currentIndex; i += 2) {
        addLocal(frame.args[i]);
      }
    }
    if (frame.name === "LAMBDA") {
      frame.args.forEach((arg) => addLocal(arg));
    }
  });

  return locals;
};

const FormulaMonacoView = React.forwardRef<FormulaViewHandle, FormulaMonacoViewProps>(
  ({ isPopout, onOpenLegacy }, ref) => {
    const shared = useModernSharedStyles();
    const styles = useStyles();
    const [formulaText, setFormulaText] = useState<string>("");
    const [status, setStatus] = useState<string>("");
    const [statusType, setStatusType] = useState<"success" | "error">("success");
    const [activeCell, setActiveCell] = useState<ActiveCellState | null>(null);
    const [formulaOutput, setFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
    const [autoCapture, setAutoCapture] = useState<boolean>(true);
    const [loadingMetadata, setLoadingMetadata] = useState<boolean>(false);
    const [editorReady, setEditorReady] = useState<boolean>(false);
    const [editorLoadError, setEditorLoadError] = useState<string>("");
    const [showSaveFunctionModal, setShowSaveFunctionModal] = useState<boolean>(false);
    const [namedFunctionName, setNamedFunctionName] = useState<string>("");
    const [namedFunctionArgs, setNamedFunctionArgs] = useState<string>("");
    const [namedFunctionArgDraft, setNamedFunctionArgDraft] = useState<string>("");
    const [namedFunctionDescription, setNamedFunctionDescription] = useState<string>("");
    const [creationMode, setCreationMode] = useState<FormulaCreationMode>("Formula");
    const [authoringMode, setAuthoringMode] = useState<FormulaAuthoringMode>("Editor");
    const [wizardTemplate, setWizardTemplate] = useState<WizardTemplate>("LAMBDA");
    const [wizardArgsInput, setWizardArgsInput] = useState<string>("");
    const [wizardReturnExpression, setWizardReturnExpression] = useState<string>("");
    const [wizardVariables, setWizardVariables] = useState<WizardLetVariable[]>([
      { id: "var-1", name: "", expression: "" },
    ]);
    const [lambdaTestInputs, setLambdaTestInputs] = useState<string[]>([]);
    const [liveTestBusy, setLiveTestBusy] = useState<boolean>(false);
    const [liveTestError, setLiveTestError] = useState<string>("");
    const [lastTestedAt, setLastTestedAt] = useState<string>("");
    const [lastTestInvocation, setLastTestInvocation] = useState<string>("");
    const [isEditorExpanded, setIsEditorExpanded] = useState<boolean>(false);
    const [editorHasFocus, setEditorHasFocus] = useState<boolean>(false);
    const [activeSubTab, setActiveSubTab] = useState<FormulaSubTab | null>("editor");

    const editorRef = useRef<Monaco.editor.IStandaloneCodeEditor | null>(null);
    const monacoRef = useRef<typeof Monaco | null>(null);
    const providerRef = useRef<Monaco.IDisposable | null>(null);
    const namesRef = useRef<string[]>([]);
    const tablesRef = useRef<string[]>([]);
    const pendingEvalRequestsRef = useRef<
      Map<string, { resolve: (result: FormulaEvaluationResult) => void; reject: (error: Error) => void; timeoutId: number }>
    >(new Map());
    const lastSyncedFormulaRef = useRef<string>("");
    const suppressAutoCaptureUntilRef = useRef<number>(0);

    const statusClass = statusType === "success" ? shared.successText : shared.errorText;
    const effectiveNamedFunctionArgs = useMemo(
      () => mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft),
      [namedFunctionArgs, namedFunctionArgDraft]
    );
    const namedFunctionArgPreview = useMemo(
      () => effectiveNamedFunctionArgs,
      [effectiveNamedFunctionArgs]
    );
    const namedFunctionFormulaPreview = useMemo(() => {
      const current = (editorRef.current?.getValue() ?? formulaText).trim();
      if (!current) {
        return { formula: "", error: "Enter a formula in the editor first." };
      }
      try {
        return { formula: buildNamedFunctionFormula(current, stringifyArgs(effectiveNamedFunctionArgs)), error: "" };
      } catch (error) {
        return { formula: "", error: normalizeError(error) };
      }
    }, [effectiveNamedFunctionArgs, formulaText]);
    const wizardPreview = useMemo(() => {
      try {
        const returnExpression = wizardReturnExpression.trim() || "0";
        const populatedVars = wizardVariables
          .map((item) => ({ name: item.name.trim(), expression: item.expression.trim() }))
          .filter((item) => item.name && item.expression);
        const seen = new Set<string>();
        populatedVars.forEach((item) => {
          if (!LAMBDA_IDENTIFIER_PATTERN.test(item.name)) {
            throw new Error(`Invalid variable name "${item.name}".`);
          }
          const key = item.name.toUpperCase();
          if (seen.has(key)) {
            throw new Error(`Duplicate variable name "${item.name}".`);
          }
          seen.add(key);
        });

        const letExpression =
          populatedVars.length > 0
            ? `LET(${populatedVars.map((item) => `${item.name},${item.expression}`).join(",")},${returnExpression})`
            : returnExpression;

        if (wizardTemplate === "LET") {
          if (populatedVars.length === 0) {
            throw new Error("LET builder requires at least one variable.");
          }
          return { formula: `=${letExpression}`, error: "" };
        }

        const lambdaArgs = parseLambdaArgs(wizardArgsInput);
        lambdaArgs.forEach((arg) => {
          if (!LAMBDA_IDENTIFIER_PATTERN.test(arg)) {
            throw new Error(`Invalid argument "${arg}".`);
          }
        });
        const lambdaFormula =
          lambdaArgs.length > 0
            ? `=LAMBDA(${lambdaArgs.join(",")},${letExpression})`
            : `=LAMBDA(${letExpression})`;
        return { formula: lambdaFormula, error: "" };
      } catch (error) {
        return { formula: "", error: normalizeError(error) };
      }
    }, [wizardArgsInput, wizardReturnExpression, wizardTemplate, wizardVariables]);
    const activeLambdaArgs = useMemo(() => {
      if (authoringMode === "Wizard" && wizardTemplate === "LAMBDA") {
        return parseLambdaArgs(wizardArgsInput);
      }
      return effectiveNamedFunctionArgs;
    }, [authoringMode, wizardArgsInput, wizardTemplate, effectiveNamedFunctionArgs]);
    const availableSubTabs = useMemo<Array<{ key: FormulaSubTab; label: string }>>(() => {
      const tabs: Array<{ key: FormulaSubTab; label: string }> = [];
      if (creationMode === "Function") {
        tabs.push({ key: "metadata", label: "Function Metadata" });
      }
      if (creationMode === "Function" && !(authoringMode === "Wizard" && wizardTemplate === "LET")) {
        tabs.push({ key: "lambda-test", label: "Lambda Test" });
      }
      tabs.push(authoringMode === "Editor" ? { key: "editor", label: "Editor" } : { key: "wizard", label: "Wizard" });
      tabs.push({ key: "live-output", label: "Live Output" });
      return tabs;
    }, [authoringMode, creationMode, wizardTemplate]);

    useEffect(() => {
      if (!activeSubTab) {
        return;
      }
      if (availableSubTabs.some((tab) => tab.key === activeSubTab)) {
        return;
      }
      setActiveSubTab(availableSubTabs[0]?.key ?? null);
    }, [activeSubTab, availableSubTabs]);

    useEffect(() => {
      let disposed = false;
      void loader
        .init()
        .then(() => {
          if (!disposed) {
            setEditorReady(true);
          }
        })
        .catch((error) => {
          if (disposed) {
            return;
          }
          const message = normalizeError(error);
          setEditorLoadError(message);
          setStatusType("error");
          setStatus(`Formula editor failed to initialize. Fallback mode enabled: ${message}`);
        });

      return () => {
        disposed = true;
      };
    }, []);

    const runAction = async (label: string, action: () => Promise<void>) => {
      setStatus("");
      try {
        await action();
        setStatusType("success");
        setStatus(`${label} completed.`);
      } catch (error) {
        setStatusType("error");
        setStatus(`${label} failed: ${normalizeError(error)}`);
      }
    };

    const commitNamedFunctionArgDraft = useCallback(() => {
      const draft = namedFunctionArgDraft.trim();
      if (!draft) {
        return;
      }
      const next = parseLambdaArgs(namedFunctionArgs);
      if (!next.some((item) => item.toUpperCase() === draft.toUpperCase())) {
        next.push(draft);
        setNamedFunctionArgs(next.join(","));
      }
      setNamedFunctionArgDraft("");
    }, [namedFunctionArgDraft, namedFunctionArgs]);

    const removeNamedFunctionArgAt = useCallback((index: number) => {
      setNamedFunctionArgs((prev) => {
        const next = parseLambdaArgs(prev);
        if (index >= 0 && index < next.length) {
          next.splice(index, 1);
        }
        return next.join(",");
      });
    }, []);

    const pullFromActiveCell = useCallback(
      async (requireFormula: boolean) => {
        const state = await getActiveCellFormulaState();
        setActiveCell(state);
        if (!state.hasFormula) {
          if (requireFormula) {
            throw new Error(`Cell ${state.sheet}!${state.address} does not contain a formula.`);
          }
          return false;
        }
        setFormulaText(state.formula);
        lastSyncedFormulaRef.current = state.formula;
        return true;
      },
      []
    );

    const applyToActiveCell = useCallback(async () => {
      const normalized = (editorRef.current?.getValue() ?? formulaText).trim();
      if (!normalized) {
        throw new Error("Formula editor is empty.");
      }
      const argsForApply = stringifyArgs(mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft));
      const formulaToApply =
        creationMode === "Function" ? buildNamedFunctionFormula(normalized, argsForApply) : normalized;
      await applyFormulaToActiveCell(formulaToApply);
      const refreshed = await getActiveCellFormulaState();
      setActiveCell(refreshed);
      lastSyncedFormulaRef.current = refreshed.formula;
      setFormulaText(refreshed.formula);
    }, [creationMode, formulaText, namedFunctionArgDraft, namedFunctionArgs]);

    const autoPullOnEditorBlur = useCallback(async () => {
      const currentEditorValue = (editorRef.current?.getValue() ?? formulaText).trim();
      if (currentEditorValue !== lastSyncedFormulaRef.current.trim()) {
        setStatusType("success");
        setStatus("Auto-pull skipped to preserve unsaved editor changes.");
        return;
      }
      const state = await getActiveCellFormulaState();
      setActiveCell(state);
      if (!state.hasFormula) {
        return;
      }
      if (state.formula.trim() === currentEditorValue) {
        return;
      }
      setFormulaText(state.formula);
      lastSyncedFormulaRef.current = state.formula;
      setStatusType("success");
      setStatus(`Auto-pulled ${state.sheet}!${state.address} after editor blur.`);
    }, [formulaText]);

    const insertSelection = useCallback(async () => {
      const selection = await getCurrentSelectionAddress();
      const token = selection.address;
      const current = editorRef.current?.getValue() ?? formulaText;
      if (!editorRef.current) {
        setFormulaText(current ? `${current}${token}` : token);
        return;
      }
      const selectionRange = editorRef.current.getSelection();
      if (!selectionRange) {
        setFormulaText(current ? `${current}${token}` : token);
        return;
      }
      editorRef.current.executeEdits("wbm-insert-selection", [
        {
          range: selectionRange,
          text: token,
          forceMoveMarkers: true,
        },
      ]);
      setFormulaText(editorRef.current.getValue());
    }, [formulaText]);

    const beautify = useCallback(async () => {
      setFormulaText((prev) => prettyFormatFormula(prev));
    }, []);

    const loadSuggestions = useCallback(async () => {
      setLoadingMetadata(true);
      try {
        const [namedRanges, tables] = await Promise.all([getNamedRanges(), getTables()]);
        namesRef.current = namedRanges
          .filter((item) => item.kind === "NamedRange")
          .map((item) => item.name);
        tablesRef.current = tables.map((item: TableRecord) => item.name);
      } finally {
        setLoadingMetadata(false);
      }
    }, []);

    useEffect(() => {
      void loadSuggestions();
    }, [loadSuggestions]);

    useEffect(() => {
      if (!autoCapture) return undefined;
      let disposed = false;
      let busy = false;
      const timerId = window.setInterval(() => {
        if (disposed || busy) return;
        busy = true;
        void (async () => {
          try {
            const state = await getActiveCellFormulaState();
            if (disposed) return;
            setActiveCell((prev) => {
              if (
                prev &&
                prev.sheet === state.sheet &&
                prev.address === state.address &&
                prev.formula === state.formula &&
                prev.hasFormula === state.hasFormula
              ) {
                return prev;
              }
              return state;
            });
            if (!state.hasFormula) return;
            if (Date.now() < suppressAutoCaptureUntilRef.current) {
              return;
            }
            const editorFocused = editorRef.current?.hasTextFocus() ?? false;
            if (!editorFocused && state.formula !== formulaText) {
              setFormulaText(state.formula);
              lastSyncedFormulaRef.current = state.formula;
            }
          } catch {
            // ignore transient excel editing states
          } finally {
            busy = false;
          }
        })();
      }, 1100);

      return () => {
        disposed = true;
        window.clearInterval(timerId);
      };
    }, [autoCapture, formulaText]);

    useEffect(() => {
      if (!isPopout || typeof Office === "undefined") {
        return undefined;
      }

      try {
        const readyMessage: DialogReadyMessage = {
          channel: DIALOG_RPC_CHANNEL,
          type: "ready",
        };
        Office.context.ui?.messageParent(JSON.stringify(readyMessage));
      } catch {
        // ignore if bridge not ready yet
      }

      const handler = (args: Office.DialogParentMessageReceivedEventArgs) => {
        let payload: unknown;
        try {
          payload = JSON.parse(args.message);
        } catch {
          return;
        }

        if (!payload || typeof payload !== "object") {
          return;
        }

        const data = payload as Partial<DialogEvalResponseMessage | DialogOpenFormulaMessage>;
        if (data.channel !== DIALOG_RPC_CHANNEL || typeof data.type !== "string") {
          return;
        }
        if (data.type === "open-formula" && typeof data.formula === "string") {
          const incomingFormula = data.formula.trim();
          if (!incomingFormula) {
            return;
          }
          const normalizedFormula = incomingFormula.startsWith("=") ? incomingFormula : `=${incomingFormula}`;
          setFormulaText(normalizedFormula);
          lastSyncedFormulaRef.current = normalizedFormula;
          suppressAutoCaptureUntilRef.current = Date.now() + 1800;
          if (typeof data.name === "string" && data.name.trim()) {
            setNamedFunctionName(data.name.trim());
          }
          if (typeof data.functionArgs === "string") {
            setNamedFunctionArgs(data.functionArgs);
            setNamedFunctionArgDraft("");
          }
          if (typeof data.description === "string") {
            setNamedFunctionDescription(data.description);
          }
          if (data.creationMode === "Formula" || data.creationMode === "Function") {
            setCreationMode(data.creationMode);
          }
          if (data.authoringMode === "Editor" || data.authoringMode === "Wizard") {
            setAuthoringMode(data.authoringMode);
          }
          if (data.wizardTemplate === "LAMBDA" || data.wizardTemplate === "LET") {
            setWizardTemplate(data.wizardTemplate);
          }
          if (typeof data.wizardArgs === "string") {
            setWizardArgsInput(data.wizardArgs);
          }
          if (typeof data.wizardReturnExpression === "string") {
            setWizardReturnExpression(data.wizardReturnExpression);
          }
          if (Array.isArray(data.wizardVariables)) {
            const mapped = data.wizardVariables
              .map((item, index) => ({
                id: `var-preload-${index.toString()}`,
                name: typeof item?.name === "string" ? item.name : "",
                expression: typeof item?.expression === "string" ? item.expression : "",
              }))
              .filter((item) => item.name || item.expression);
            if (mapped.length > 0) {
              setWizardVariables(mapped);
            }
          }
          if (data.entryType === "Function") {
            setCreationMode("Function");
            const extractedArgs = extractLambdaArgsFromFormula(normalizedFormula);
            if (!data.functionArgs && extractedArgs.length > 0) {
              setNamedFunctionArgs(extractedArgs.join(","));
              setNamedFunctionArgDraft("");
            }
          } else if (!data.creationMode) {
            setCreationMode("Formula");
          }
          if (!data.authoringMode) {
            setAuthoringMode("Editor");
          }
          setStatusType("success");
          setStatus(`Loaded "${data.name || "selected name"}" into the formula editor.`);
          return;
        }
        if (data.type !== "eval-response" || typeof data.requestId !== "string") {
          return;
        }

        const pending = pendingEvalRequestsRef.current.get(data.requestId);
        if (!pending) {
          return;
        }

        pendingEvalRequestsRef.current.delete(data.requestId);
        window.clearTimeout(pending.timeoutId);

        if (data.ok && data.result) {
          pending.resolve(data.result);
          return;
        }

        pending.reject(new Error(typeof data.error === "string" ? data.error : "Live test failed."));
      };

      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, handler);

      return () => {
        pendingEvalRequestsRef.current.forEach((pending) => {
          window.clearTimeout(pending.timeoutId);
          pending.reject(new Error("Dialog listener was reset."));
        });
        pendingEvalRequestsRef.current.clear();
      };
    }, [isPopout]);

    const evaluateFormulaForLive = useCallback(
      async (normalizedFormula: string): Promise<FormulaEvaluationResult> => {
        if (!isPopout) {
          return evaluateFormula(normalizedFormula);
        }
        if (typeof Office === "undefined" || !Office.context.ui?.messageParent) {
          throw new Error("Dialog bridge is unavailable for live testing.");
        }

        return await new Promise<FormulaEvaluationResult>((resolve, reject) => {
          const requestId = `eval-${Date.now()}-${Math.random().toString(16).slice(2)}`;
          const timeoutId = window.setTimeout(() => {
            pendingEvalRequestsRef.current.delete(requestId);
            reject(new Error("Live test timed out waiting for task pane response."));
          }, 10000);

          pendingEvalRequestsRef.current.set(requestId, { resolve, reject, timeoutId });

          const request: DialogEvalRequestMessage = {
            channel: DIALOG_RPC_CHANNEL,
            type: "eval-request",
            requestId,
            formula: normalizedFormula,
          };

          try {
            Office.context.ui.messageParent(JSON.stringify(request));
          } catch (error) {
            pendingEvalRequestsRef.current.delete(requestId);
            window.clearTimeout(timeoutId);
            reject(new Error(normalizeError(error)));
          }
        });
      },
      [isPopout]
    );

    useEffect(() => {
      const normalized = formulaText.trim();
      if (!normalized) {
        setFormulaOutput(null);
        setLiveTestError("");
        setLiveTestBusy(false);
        setLastTestedAt("");
        return undefined;
      }

      let disposed = false;
      const timerId = window.setTimeout(() => {
        setLiveTestBusy(true);
        setLiveTestError("");
        void evaluateFormulaForLive(normalized)
          .then((result) => {
            if (disposed) {
              return;
            }
            setFormulaOutput(result);
            if (result.hasError) {
              const errorText = getFirstErrorCell(result);
              setLiveTestError(`Evaluation returned ${errorText}.`);
            } else {
              setLiveTestError("");
            }
            setLastTestedAt(new Date().toLocaleTimeString());
          })
          .catch((error) => {
            if (disposed) {
              return;
            }
            setFormulaOutput(null);
            setLiveTestError(normalizeError(error));
            setLastTestedAt(new Date().toLocaleTimeString());
          })
          .finally(() => {
            if (!disposed) {
              setLiveTestBusy(false);
            }
          });
      }, 320);

      return () => {
        disposed = true;
        window.clearTimeout(timerId);
      };
    }, [evaluateFormulaForLive, formulaText]);

    const registerCompletionProvider = useCallback(() => {
      if (!monacoRef.current) return;
      if (providerRef.current) {
        providerRef.current.dispose();
      }

      const monaco = monacoRef.current;
      try {
        providerRef.current = monaco.languages.registerCompletionItemProvider(MONACO_LANGUAGE_ID, {
          triggerCharacters: ["=", "(", ",", "_", "[", "."],
          provideCompletionItems: (model, position) => {
            const cursorOffset = model.getOffsetAt(position);
            const token = getTokenBeforeCursor(model.getValue(), cursorOffset);
            const tokenUpper = token.toUpperCase();
            const word = model.getWordUntilPosition(position);
            const range = {
              startLineNumber: position.lineNumber,
              endLineNumber: position.lineNumber,
              startColumn: word.startColumn,
              endColumn: word.endColumn,
            };

            const suggestions = new Map<string, Monaco.languages.CompletionItem>();
            const matchesToken = (candidate: string): boolean =>
              !tokenUpper || candidate.toUpperCase().includes(tokenUpper);

            const addSuggestion = (
              label: string,
              kind: Monaco.languages.CompletionItemKind,
              insertText: string,
              detail: string,
              order: string,
              asSnippet = false
            ) => {
              if (!matchesToken(label)) {
                return;
              }
              const key = label.toUpperCase();
              if (suggestions.has(key)) {
                return;
              }
              suggestions.set(key, {
                label,
                kind,
                insertText,
                insertTextRules: asSnippet
                  ? monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet
                  : undefined,
                detail,
                range,
                sortText: `${order}_${label.toUpperCase()}`,
              });
            };

            collectLetLambdaLocals(model.getValue(), cursorOffset).forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Variable,
                name,
                "LET/LAMBDA variable",
                "0"
              );
            });

            namesRef.current.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Variable,
                name,
                "Named range/function",
                "1"
              );
            });

            tablesRef.current.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Class,
                `${name}[`,
                "Table",
                "2"
              );
            });

            FUNCTION_SUGGESTIONS.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Function,
                `${name}($1)`,
                "Function",
                "3",
                true
              );
            });

            return { suggestions: Array.from(suggestions.values()).slice(0, 200) };
          },
        });
      } catch (error) {
        const message = normalizeError(error);
        setEditorLoadError(message);
        setStatusType("error");
        setStatus(`Autocomplete registration failed. Fallback mode enabled: ${message}`);
      }
    }, []);

    const beforeMount = useCallback(
      (monaco: typeof Monaco) => {
        monacoRef.current = monaco;
        if (!monaco.languages.getLanguages().some((lang) => lang.id === MONACO_LANGUAGE_ID)) {
          monaco.languages.register({ id: MONACO_LANGUAGE_ID });
          monaco.languages.setLanguageConfiguration(MONACO_LANGUAGE_ID, {
            brackets: [
              ["(", ")"],
              ["[", "]"],
            ],
            autoClosingPairs: [
              { open: "(", close: ")" },
              { open: "[", close: "]" },
              { open: '"', close: '"' },
            ],
            surroundingPairs: [
              { open: "(", close: ")" },
              { open: "[", close: "]" },
              { open: '"', close: '"' },
            ],
            wordPattern: /[A-Za-z_\\][A-Za-z0-9_.\\]*/g,
          });
        }
        registerCompletionProvider();
      },
      [registerCompletionProvider]
    );

    const onMount: OnMount = useCallback(
      (editor) => {
        editorRef.current = editor;
        editor.onDidChangeModelContent(() => {
          setFormulaText(editor.getValue());
        });
        editor.onDidFocusEditorText(() => {
          setEditorHasFocus(true);
        });
        editor.onDidBlurEditorText(() => {
          setEditorHasFocus(false);
          void autoPullOnEditorBlur();
        });
      },
      [autoPullOnEditorBlur]
    );

    useEffect(() => {
      registerCompletionProvider();
      return () => {
        providerRef.current?.dispose();
      };
    }, [registerCompletionProvider]);

    useEffect(() => {
      setLambdaTestInputs((prev) => {
        const next = activeLambdaArgs.map((_, index) => prev[index] ?? "");
        return next;
      });
    }, [activeLambdaArgs]);

    useEffect(() => {
      if (!editorRef.current) return;
      if (editorRef.current.getValue() !== formulaText) {
        editorRef.current.setValue(formulaText);
      }
    }, [formulaText]);

    const closeSaveFunctionModal = useCallback(() => {
      setShowSaveFunctionModal(false);
      setNamedFunctionName("");
      setNamedFunctionArgs("");
      setNamedFunctionArgDraft("");
      setNamedFunctionDescription("");
    }, []);

    const buildSaveFailureMessage = (
      error: unknown,
      functionName: string
    ): string => {
      const source = normalizeError(error);
      const tips: string[] = [];
      if (!EXCEL_NAME_PATTERN.test(functionName) || CELL_REFERENCE_LIKE_PATTERN.test(functionName)) {
        tips.push("Use a valid Excel name (starts with letter/_ and cannot look like A1).");
      }
      if (namedFunctionArgPreview.length > 0 && namedFunctionArgPreview.some((arg) => !LAMBDA_IDENTIFIER_PATTERN.test(arg))) {
        tips.push("Each argument must be a valid identifier (letters/numbers/_/./\\).");
      }
      if (/argument is invalid|incorrect format/i.test(source)) {
        tips.push("Verify LET pairs (name,value,...) plus a final return expression.");
      }
      const guidance = tips.length > 0 ? ` ${tips.join(" ")}` : "";
      return `Function was not saved. ${source}.${guidance}`;
    };

    const saveCurrentAsNamedFunction = async () => {
      setStatus("");
      try {
        const current = (editorRef.current?.getValue() ?? formulaText).trim();
        const functionName = namedFunctionName.trim();
        if (!functionName) {
          throw new Error("Function name is required.");
        }
        if (!EXCEL_NAME_PATTERN.test(functionName) || CELL_REFERENCE_LIKE_PATTERN.test(functionName)) {
          throw new Error("Function name must be a valid Excel defined name and cannot look like a cell reference.");
        }
        const argsForSave = stringifyArgs(mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft));
        setNamedFunctionArgs(argsForSave);
        setNamedFunctionArgDraft("");
        const lambdaFormula = buildNamedFunctionFormula(current, argsForSave);
        await saveNamedFunction(functionName, lambdaFormula);
        await loadSuggestions();
        const saved = namesRef.current.some((item) => item.toUpperCase() === functionName.toUpperCase());
        if (!saved) {
          throw new Error("Save request completed, but the function is not visible yet. Refresh and try again.");
        }
        setStatusType("success");
        setStatus(`Function "${functionName}" saved.`);
        closeSaveFunctionModal();
      } catch (error) {
        setStatusType("error");
        setStatus(buildSaveFailureMessage(error, namedFunctionName.trim()));
      }
    };

    const openPopoutWithCurrentState = async () => {
      const currentFormula = (editorRef.current?.getValue() ?? formulaText).trim();
      const argsForPopout = stringifyArgs(mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft));
      setNamedFunctionArgs(argsForPopout);
      setNamedFunctionArgDraft("");
      await openFormulaEditorPopout({
        formula: currentFormula || "=",
        name: namedFunctionName.trim() || undefined,
        entryType: creationMode === "Function" ? "Function" : "Formula",
        functionArgs: argsForPopout,
        description: namedFunctionDescription,
        creationMode,
        authoringMode,
        wizardTemplate,
        wizardArgs: wizardArgsInput,
        wizardReturnExpression,
        wizardVariables: wizardVariables.map((item) => ({
          name: item.name,
          expression: item.expression,
        })),
      });
    };

    const addWizardVariable = () => {
      setWizardVariables((prev) => [
        ...prev,
        { id: `var-${Date.now().toString()}-${Math.random().toString(16).slice(2)}`, name: "", expression: "" },
      ]);
    };

    const removeWizardVariable = (id: string) => {
      setWizardVariables((prev) => {
        const next = prev.filter((item) => item.id !== id);
        return next.length > 0 ? next : [{ id: "var-1", name: "", expression: "" }];
      });
    };

    const updateWizardVariable = (id: string, field: "name" | "expression", value: string) => {
      setWizardVariables((prev) =>
        prev.map((item) => (item.id === id ? { ...item, [field]: value } : item))
      );
    };

    const applyWizardToEditor = async () => {
      await runAction("Build formula from wizard", async () => {
        if (wizardPreview.error) {
          throw new Error(wizardPreview.error);
        }
        if (!wizardPreview.formula) {
          throw new Error("Wizard formula is empty.");
        }
        setFormulaText(wizardPreview.formula);
        if (wizardTemplate === "LAMBDA") {
          setCreationMode("Function");
          setNamedFunctionArgs(parseLambdaArgs(wizardArgsInput).join(","));
          setNamedFunctionArgDraft("");
        }
        setAuthoringMode("Editor");
      });
    };

    const runLambdaInputTest = async () => {
      await runAction("Run lambda input test", async () => {
        if (authoringMode === "Wizard" && wizardTemplate === "LET") {
          throw new Error("Switch wizard type to LAMBDA before running lambda input tests.");
        }
        const lambdaFormula =
          authoringMode === "Wizard" && wizardTemplate === "LAMBDA"
            ? wizardPreview.formula
            : buildNamedFunctionFormula(
                (editorRef.current?.getValue() ?? formulaText).trim(),
                stringifyArgs(mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft))
              );
        const invokeFormula = buildLambdaInvokeFormula(lambdaFormula, lambdaTestInputs);
        setLastTestInvocation(invokeFormula);
        const result = await evaluateFormulaForLive(invokeFormula);
        setFormulaOutput(result);
        setLastTestedAt(new Date().toLocaleTimeString());
        if (result.hasError) {
          const errorText = getFirstErrorCell(result);
          setLiveTestError(`Lambda test returned ${errorText}.`);
          throw new Error(`Lambda invocation returned ${errorText}. Check arguments and LET/LAMBDA structure.`);
        }
        setLiveTestError("");
      });
    };

    const handleNamedArgInputKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
      if (event.key === " " || event.key === "Enter" || event.key === ",") {
        event.preventDefault();
        commitNamedFunctionArgDraft();
        return;
      }
      if (event.key === "Backspace" && !namedFunctionArgDraft.trim() && namedFunctionArgPreview.length > 0) {
        event.preventDefault();
        removeNamedFunctionArgAt(namedFunctionArgPreview.length - 1);
      }
    };

    React.useImperativeHandle(ref, () => ({
      openOnly: async () => {
        await runAction("Load active cell state", async () => {
          const state = await getActiveCellFormulaState();
          setActiveCell(state);
        });
      },
      openAndPull: async () => {
        await runAction("Open active cell formula in editor", async () => {
          await pullFromActiveCell(true);
        });
      },
      pull: async () => {
        await runAction("Pull active cell formula", async () => {
          await pullFromActiveCell(true);
        });
      },
      apply: async () => {
        await runAction("Apply editor formula to active cell", applyToActiveCell);
      },
      beautify: async () => {
        await runAction("Beautify formula", beautify);
      },
      insertSelection: async () => {
        await runAction("Insert grid selection", insertSelection);
      },
    }));

    const editorLineHeight = 20;
    const defaultVisibleLines = 10;
    const editorVerticalChrome = 176;
    const compactEditorHeight =
      editorLineHeight * defaultVisibleLines + editorVerticalChrome + (isPopout ? 80 : 0);
    const expandedEditorHeight = compactEditorHeight + (isPopout ? 220 : 180);
    const editorHeight = `${isEditorExpanded ? expandedEditorHeight : compactEditorHeight}px`;

    return (
      <div className={styles.root}>
        {!isPopout ? (
          <div>
            <Text className={shared.sectionTitle}>Formulas</Text>
            <Text className={shared.sectionSubtitle}>Manage and edit workbook formulas with live test output.</Text>
          </div>
        ) : null}

        <div className={styles.card}>
          <div className={styles.autoCaptureRow}>
            <Text className={shared.mutedText}>
              Active Cell: {activeCell ? `${activeCell.sheet}!${activeCell.address}` : "not captured"}
              {" • "}
              <span style={{ color: MODERN_TOKENS.colorBrandStrong, fontWeight: 700 }}>Connected</span>
              {loadingMetadata ? " • loading suggestions..." : ""}
            </Text>
            <label className={styles.checkboxLabel}>
              <input type="checkbox" checked={autoCapture} onChange={(event) => setAutoCapture(event.target.checked)} />{" "}
              Auto-capture active formula
            </label>
          </div>

          <div className={styles.modeRow}>
            <Text className={shared.mutedText}>Create:</Text>
            <div className={styles.slicer}>
              {(["Formula", "Function"] as FormulaCreationMode[]).map((mode) => (
                <button
                  key={mode}
                  type="button"
                  className={`${styles.slicerBtn} ${creationMode === mode ? styles.slicerBtnActive : ""}`}
                  onClick={() => setCreationMode(mode)}
                >
                  {mode === "Function" ? "Function (LAMBDA)" : "Formula"}
                </button>
              ))}
            </div>
            <Text className={shared.mutedText}>Authoring:</Text>
            <div className={styles.slicer}>
              {(["Editor", "Wizard"] as FormulaAuthoringMode[]).map((mode) => (
                <button
                  key={mode}
                  type="button"
                  className={`${styles.slicerBtn} ${authoringMode === mode ? styles.slicerBtnActive : ""}`}
                  onClick={() => setAuthoringMode(mode)}
                >
                  {mode}
                </button>
              ))}
            </div>
          </div>

          <div className={styles.subTabBar}>
            {availableSubTabs.map((tab) => (
              <button
                key={tab.key}
                type="button"
                className={`${styles.subTabBtn} ${activeSubTab === tab.key ? styles.subTabBtnActive : ""}`}
                onClick={() => setActiveSubTab((prev) => (prev === tab.key ? null : tab.key))}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {activeSubTab ? (
            <div className={styles.subTabPanel}>
              {activeSubTab === "metadata" && creationMode === "Function" ? (
                <div className={styles.functionMetaGrid}>
                  <div>
                    <Text className={styles.modalLabel}>Function name</Text>
                    <Input
                      placeholder="Function name (example: CalcMargin)"
                      value={namedFunctionName}
                      onChange={(_, data) => setNamedFunctionName(data.value)}
                    />
                  </div>
                  <div>
                    <Text className={styles.modalLabel}>Arguments (space creates a pill)</Text>
                    <Input
                      className={styles.argsInput}
                      placeholder="Type argument and press Space"
                      value={namedFunctionArgDraft}
                      onChange={(_, data) => setNamedFunctionArgDraft(data.value)}
                      onKeyDown={handleNamedArgInputKeyDown}
                      onBlur={commitNamedFunctionArgDraft}
                    />
                    {namedFunctionArgPreview.length > 0 ? (
                      <div className={styles.chipRow}>
                        {namedFunctionArgPreview.map((arg, index) => (
                          <span key={`${arg}-${index.toString()}`} className={styles.chip}>
                            {arg}
                            <button
                              type="button"
                              className={styles.chipBtn}
                              onClick={() => removeNamedFunctionArgAt(index)}
                              aria-label={`Remove ${arg}`}
                            >
                              x
                            </button>
                          </span>
                        ))}
                      </div>
                    ) : (
                      <Text className={shared.mutedText}>No arguments yet.</Text>
                    )}
                  </div>
                  <div className={styles.fullRow}>
                    <Text className={styles.modalLabel}>Description (optional)</Text>
                    <Input
                      placeholder="Description for your team docs"
                      value={namedFunctionDescription}
                      onChange={(_, data) => setNamedFunctionDescription(data.value)}
                    />
                  </div>
                </div>
              ) : null}

              {activeSubTab === "lambda-test" &&
              creationMode === "Function" &&
              !(authoringMode === "Wizard" && wizardTemplate === "LET") ? (
                <>
                  {activeLambdaArgs.length > 0 ? (
                    <div className={styles.lambdaTestGrid}>
                      {activeLambdaArgs.map((arg, index) => (
                        <div key={`${arg}-${index.toString()}`}>
                          <Text className={styles.modalLabel}>{arg}</Text>
                          <Input
                            placeholder="Input expression"
                            value={lambdaTestInputs[index] ?? ""}
                            onChange={(_, data) =>
                              setLambdaTestInputs((prev) => {
                                const next = [...prev];
                                next[index] = data.value;
                                return next;
                              })
                            }
                          />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <Text className={shared.mutedText}>
                      No explicit arguments defined. Test will invoke the lambda with no parameters.
                    </Text>
                  )}
                  <div className={styles.actionRow}>
                    <Button appearance="primary" size="small" onClick={() => void runLambdaInputTest()}>
                      Run Lambda Test
                    </Button>
                  </div>
                </>
              ) : null}

              {activeSubTab === "editor" && authoringMode === "Editor" ? (
                <div className={styles.editorShell}>
                  <div className={styles.overlayTop}>
                    <Button size="small" onClick={() => void runAction("Beautify formula", beautify)}>
                      Beautify
                    </Button>
                    <Button size="small" onClick={() => void runAction("Insert selection", insertSelection)}>
                      Insert Selection
                    </Button>
                    <Button size="small" onClick={() => setIsEditorExpanded((prev) => !prev)}>
                      {isEditorExpanded ? "Use Smaller Editor" : "Expand Editor"}
                    </Button>
                    {!isPopout ? (
                      <Button
                        size="small"
                        onClick={() => void runAction("Open formula editor popout", openPopoutWithCurrentState)}
                      >
                        Open Pop-out
                      </Button>
                    ) : null}
                    {!isPopout ? (
                      <Button size="small" onClick={onOpenLegacy}>
                        Legacy View
                      </Button>
                    ) : null}
                  </div>

                  <div className={styles.editorWrap}>
                    {editorLoadError ? (
                      <textarea
                        aria-label="Formula editor fallback"
                        className={styles.fallbackEditor}
                        spellCheck={false}
                        rows={10}
                        style={{ height: editorHeight }}
                        value={formulaText}
                        onChange={(event) => setFormulaText(event.target.value)}
                        onFocus={() => setEditorHasFocus(true)}
                        onBlur={() => {
                          setEditorHasFocus(false);
                          void autoPullOnEditorBlur();
                        }}
                      />
                    ) : editorReady ? (
                      <Editor
                        height={editorHeight}
                        language={MONACO_LANGUAGE_ID}
                        value={formulaText}
                        beforeMount={beforeMount}
                        onMount={onMount}
                        theme="vs"
                        options={{
                          minimap: { enabled: false },
                          scrollBeyondLastLine: false,
                          fontSize: 13,
                          lineHeight: editorLineHeight,
                          lineNumbers: "on",
                          wordWrap: "off",
                          automaticLayout: true,
                          suggestOnTriggerCharacters: true,
                          quickSuggestions: { other: true, comments: false, strings: false },
                          wordBasedSuggestions: "off",
                          tabCompletion: "on",
                          padding: {
                            top: 38,
                            bottom: 128,
                          },
                          suggest: {
                            showWords: false,
                            showSnippets: true,
                            preview: true,
                          },
                        }}
                      />
                    ) : (
                      <div className={styles.editorLoading} style={{ height: editorHeight }}>
                        <Text className={shared.mutedText}>Loading formula editor...</Text>
                      </div>
                    )}
                  </div>

                  {!editorHasFocus ? (
                    <div className={styles.actionDock}>
                      <Text className={styles.overlayTitle}>Apply / Save</Text>
                      <div className={styles.actionRow}>
                        <Button
                          appearance="primary"
                          size="small"
                          onClick={() => void runAction("Apply formula", applyToActiveCell)}
                        >
                          Apply Formula
                        </Button>
                        {creationMode === "Function" ? (
                          <Button size="small" onClick={() => void saveCurrentAsNamedFunction()}>
                            Save Function
                          </Button>
                        ) : (
                          <Button size="small" onClick={() => setShowSaveFunctionModal(true)}>
                            Save Named Function
                          </Button>
                        )}
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : null}

              {activeSubTab === "wizard" && authoringMode === "Wizard" ? (
                <>
                  <div className={styles.modeRow}>
                    <Text className={shared.mutedText}>Wizard type:</Text>
                    <Select
                      className={styles.modeSelect}
                      value={wizardTemplate}
                      onChange={(_, data) => setWizardTemplate(data.value as WizardTemplate)}
                    >
                      <option value="LAMBDA">LAMBDA Builder</option>
                      <option value="LET">LET Builder</option>
                    </Select>
                  </div>
                  {wizardTemplate === "LAMBDA" ? (
                    <div>
                      <Text className={styles.modalLabel}>Arguments (comma-separated)</Text>
                      <Input
                        placeholder="table, lookupValue"
                        value={wizardArgsInput}
                        onChange={(_, data) => setWizardArgsInput(data.value)}
                      />
                    </div>
                  ) : null}
                  <div className={styles.modeRow}>
                    <Text className={styles.modalLabel}>Variables</Text>
                    <Button size="small" onClick={addWizardVariable}>
                      + Add Variable
                    </Button>
                  </div>
                  {wizardVariables.map((item) => (
                    <div key={item.id} className={styles.wizardVarRow}>
                      <Input
                        placeholder="name"
                        value={item.name}
                        onChange={(_, data) => updateWizardVariable(item.id, "name", data.value)}
                      />
                      <Input
                        placeholder="expression/value"
                        value={item.expression}
                        onChange={(_, data) => updateWizardVariable(item.id, "expression", data.value)}
                      />
                      <Button size="small" onClick={() => removeWizardVariable(item.id)}>
                        Remove
                      </Button>
                    </div>
                  ))}
                  <div>
                    <Text className={styles.modalLabel}>
                      {wizardTemplate === "LAMBDA" ? "Return expression" : "LET final expression"}
                    </Text>
                    <Input
                      placeholder="Result expression"
                      value={wizardReturnExpression}
                      onChange={(_, data) => setWizardReturnExpression(data.value)}
                    />
                  </div>
                  <div>
                    <Text className={styles.modalLabel}>Wizard preview</Text>
                    {wizardPreview.error ? (
                      <Text className={shared.errorText}>{wizardPreview.error}</Text>
                    ) : (
                      <pre className={styles.wizardPreview}>{wizardPreview.formula}</pre>
                    )}
                  </div>
                  <div className={styles.actionRow}>
                    <Button appearance="primary" onClick={() => void applyWizardToEditor()}>
                      Build Into Editor
                    </Button>
                    <Button onClick={() => setAuthoringMode("Editor")}>Switch To Editor</Button>
                  </div>
                </>
              ) : null}

              {activeSubTab === "live-output" ? (
                <>
                  <div className={styles.liveTestHeader}>
                    <Text className={shared.cardTitle}>Live Test Output</Text>
                    <Text className={shared.mutedText}>
                      {liveTestBusy
                        ? "Testing..."
                        : liveTestError
                          ? `Last test failed${lastTestedAt ? ` at ${lastTestedAt}` : ""}`
                          : lastTestedAt
                            ? `Auto-tested at ${lastTestedAt}`
                            : "Waiting for formula input"}
                    </Text>
                  </div>

                  {lastTestInvocation ? (
                    <Text className={styles.invocationText}>Invocation: {lastTestInvocation}</Text>
                  ) : null}
                  {liveTestError ? <Text className={shared.errorText}>{liveTestError}</Text> : null}

                  <div className={styles.outputWrap}>
                    {formulaOutput ? (
                      <table className={styles.outputTable}>
                        <tbody>
                          {formulaOutput.values.map((row, rowIndex) => (
                            <tr key={rowIndex}>
                              {row.map((cell, colIndex) => (
                                <td key={`${rowIndex}-${colIndex}`} className={styles.outputCell}>
                                  {cell === null ? "" : String(cell)}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <div className={styles.outputCell}>
                        <Text className={shared.mutedText}>Live test will render results here as you edit.</Text>
                      </div>
                    )}
                  </div>
                </>
              ) : null}
            </div>
          ) : null}

          {status ? <Text className={statusClass}>{status}</Text> : null}
        </div>

        {showSaveFunctionModal ? (
          <div className={styles.modalBackdrop}>
            <div className={styles.modal}>
              <Text className={shared.cardTitle}>Save As Named Function</Text>
              <Text className={shared.cardSubtitle}>
                Add function metadata and arguments. If the editor formula is not already a LAMBDA expression, it will
                be wrapped automatically.
              </Text>
              <div className={styles.modalField}>
                <Text className={styles.modalLabel}>Function name</Text>
                <Input
                  placeholder="Function name (example: CalcMargin)"
                  value={namedFunctionName}
                  onChange={(_, data) => setNamedFunctionName(data.value)}
                />
              </div>
              <div className={styles.modalField}>
                <Text className={styles.modalLabel}>Description (optional)</Text>
                <Input
                  placeholder="Description for your team docs"
                  value={namedFunctionDescription}
                  onChange={(_, data) => setNamedFunctionDescription(data.value)}
                />
                <Text className={shared.mutedText}>Descriptions are not persisted by Excel named formulas.</Text>
              </div>
              <div className={styles.modalField}>
                <Text className={styles.modalLabel}>Arguments (comma-separated)</Text>
                <Input
                  placeholder="table, lookupValue"
                  value={namedFunctionArgs}
                  onChange={(_, data) => setNamedFunctionArgs(data.value)}
                />
                {namedFunctionArgPreview.length > 0 ? (
                  <div className={styles.chipRow}>
                    {namedFunctionArgPreview.map((arg, index) => (
                      <span key={`${arg}-${index.toString()}`} className={styles.chip}>
                        {arg}
                      </span>
                    ))}
                  </div>
                ) : (
                  <Text className={shared.mutedText}>No arguments set. The editor formula will be used as LAMBDA body.</Text>
                )}
              </div>
              <div className={styles.modalField}>
                <Text className={styles.modalLabel}>Final formula preview</Text>
                {namedFunctionFormulaPreview.error ? (
                  <Text className={shared.errorText}>{namedFunctionFormulaPreview.error}</Text>
                ) : (
                  <pre className={styles.formulaPreview}>{namedFunctionFormulaPreview.formula}</pre>
                )}
              </div>
              <div className={styles.modalActions}>
                <Button onClick={closeSaveFunctionModal}>Cancel</Button>
                <Button appearance="primary" onClick={() => void saveCurrentAsNamedFunction()}>
                  Save Function
                </Button>
              </div>
            </div>
          </div>
        ) : null}
      </div>
    );
  }
);

FormulaMonacoView.displayName = "FormulaMonacoView";

export default FormulaMonacoView;
