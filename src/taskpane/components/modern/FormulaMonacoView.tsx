import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Checkbox, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import Editor, { OnMount, loader } from "@monaco-editor/react";
import { Delete20Regular } from "@fluentui/react-icons";
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
  "ABS","ACCRINT","ACCRINTM","ACOS","ACOSH","ACOT","ACOTH","ADDRESS","AGGREGATE","AMORDEGRC",
  "AMORLINC","AND","ARABIC","AREAS","ASC","ASIN","ASINH","ATAN","ATAN2","ATANH","AVEDEV",
  "AVERAGE","AVERAGEA","AVERAGEIF","AVERAGEIFS","BAHTTEXT","BASE","BESSELI","BESSELJ","BESSELK",
  "BESSELY","BETA.DIST","BETA.INV","BETADIST","BETAINV","BIN2DEC","BIN2HEX","BIN2OCT",
  "BINOM.DIST","BINOM.DIST.RANGE","BINOM.INV","BINOMDIST","BITAND","BITLSHIFT","BITOR",
  "BITRSHIFT","BITXOR","CEILING","CEILING.MATH","CEILING.PRECISE","CELL","CHAR","CHISQ.DIST",
  "CHISQ.DIST.RT","CHISQ.INV","CHISQ.INV.RT","CHISQ.TEST","CHISQDIST","CHISQINV","CHITEST",
  "CHOOSE","CLEAN","CODE","COLUMN","COLUMNS","COMBIN","COMBINA","COMPLEX","CONCAT","CONCATENATE",
  "CONFIDENCE","CONFIDENCE.NORM","CONFIDENCE.T","CONVERT","CORREL","COS","COSH","COT","COTH",
  "COUNT","COUNTA","COUNTBLANK","COUNTIF","COUNTIFS","COUPDAYBS","COUPDAYS","COUPDAYSNC",
  "COUPNCD","COUPNUM","COUPPCD","COVARIANCE.P","COVARIANCE.S","COVAR","CRITBINOM","CSC","CSCH",
  "CUBEKPIMEMBER","CUBEMEMBER","CUBEMEMBERPROPERTY","CUBERANKEDMEMBER","CUBESET","CUBESETCOUNT",
  "CUBEVALUE","CUMIPMT","CUMPRINC","DATE","DATEDIF","DATEVALUE","DAVERAGE","DAY","DAYS","DAYS360",
  "DB","DCOUNT","DCOUNTA","DDB","DEC2BIN","DEC2HEX","DEC2OCT","DECIMAL","DEGREES","DELTA","DEVSQ",
  "DGET","DISC","DMAX","DMIN","DOLLAR","DOLLARDE","DOLLARFR","DPRODUCT","DSTDEV","DSTDEVP","DSUM",
  "DURATION","DVAR","DVARP","ECMA.CEILING","EDATE","EFFECT","EOMONTH","ERF","ERF.PRECISE","ERFC",
  "ERFC.PRECISE","ERROR.TYPE","EUROCONVERT","EVEN","EXACT","EXP","EXPON.DIST","EXPONDIST","FACT",
  "FACTDOUBLE","FALSE","FDIST","FILTER","FIND","FINDB","FISHER","FISHERINV","FIXED","FLOOR",
  "FLOOR.MATH","FLOOR.PRECISE","FORECAST","FORECAST.ETS","FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY","FORECAST.ETS.STAT","FORECAST.LINEAR","FORMULATEXT","FREQUENCY",
  "FTEST","FV","FVSCHEDULE","GAMMA","GAMMA.DIST","GAMMA.INV","GAMMADIST","GAMMAINV","GAMMALN",
  "GAMMALN.PRECISE","GAUSS","GCD","GEOMEAN","GESTEP","GETPIVOTDATA","GROWTH","HARMEAN","HEX2BIN",
  "HEX2DEC","HEX2OCT","HLOOKUP","HOUR","HYPERLINK","HYPGEOM.DIST","HYPGEOMDIST","IF","IFERROR",
  "IFNA","IFS","IMABS","IMAGINARY","IMARGUMENT","IMCONJUGATE","IMCOS","IMCOSH","IMCOT","IMCSC",
  "IMCSCH","IMDIV","IMEXP","IMLN","IMLOG10","IMLOG2","IMPOWER","IMPRODUCT","IMREAL","IMSEC",
  "IMSECH","IMSIN","IMSINH","IMSQRT","IMSUB","IMSUM","INDEX","INDIRECT","INFO","INT","INTERCEPT",
  "INTRATE","IPMT","IRR","ISBLANK","ISERR","ISERROR","ISEVEN","ISFORMULA","ISLOGICAL","ISNA",
  "ISNONTEXT","ISNUMBER","ISODD","ISPMT","ISREF","ISTEXT","KURT","LARGE","LAMBDA","LCM","LEFT",
  "LEFTB","LEN","LENB","LET","LINEST","LN","LOG","LOG10","LOGEST","LOGINV","LOGNORM.DIST",
  "LOGNORM.INV","LOGNORMDIST","LOOKUP","LOWER","MATCH","MAX","MAXA","MAXIFS","MDETERM","MDURATION",
  "MEDIAN","MID","MIDB","MIN","MINA","MINIFS","MINUTE","MINVERSE","MIRR","MMULT","MOD","MODE",
  "MODE.MULT","MODE.SNGL","MONTH","MROUND","MULTINOMIAL","N","NA","NEGBINOM.DIST","NEGBINOMDIST",
  "NETWORKDAYS","NETWORKDAYS.INTL","NOMINAL","NORM.DIST","NORM.INV","NORM.S.DIST","NORM.S.INV",
  "NORMDIST","NORMINV","NORMSDIST","NORMSINV","NOT","NOW","NPER","NPV","NUMBERVALUE","OCT2BIN",
  "OCT2DEC","OCT2HEX","ODD","ODDFPRICE","ODDFYIELD","ODDLPRICE","ODDLYIELD","OFFSET","OR",
  "PDURATION","PEARSON","PERCENTILE","PERCENTILE.EXC","PERCENTILE.INC","PERCENTRANK",
  "PERCENTRANK.EXC","PERCENTRANK.INC","PERMUT","PERMUTATIONA","PHI","PI","PMT","POISSON",
  "POISSON.DIST","POWER","PPMT","PRICE","PRICEDISC","PRICEMAT","PROB","PRODUCT","PROPER","PV",
  "QUARTILE","QUARTILE.EXC","QUARTILE.INC","QUOTIENT","RADIANS","RAND","RANDARRAY","RANDBETWEEN",
  "RANK","RANK.AVG","RANK.EQ","RATE","RECEIVED","REDUCE","REPLACE","REPLACEB","REPT","RIGHT",
  "RIGHTB","ROMAN","ROUND","ROUNDDOWN","ROUNDUP","ROW","ROWS","RRI","RSQ","RTD","SCAN","SEARCH",
  "SEARCHB","SEC","SECH","SECOND","SEQUENCE","SERIESSUM","SHEET","SHEETS","SIGN","SIN","SINH",
  "SKEW","SKEW.P","SLN","SLOPE","SMALL","SORT","SORTBY","SQRT","SQRTPI","STANDARDIZE","STDEV",
  "STDEV.P","STDEV.S","STDEVA","STDEVP","STDEVPA","STEYX","SUBSTITUTE","SUBTOTAL","SUM","SUMIF",
  "SUMIFS","SUMPRODUCT","SUMSQ","SUMX2MY2","SUMX2PY2","SUMXMY2","SWITCH","SYD","T","T.DIST",
  "T.DIST.2T","T.DIST.RT","T.INV","T.INV.2T","T.TEST","TAN","TANH","TBILLEQ","TBILLPRICE",
  "TBILLYIELD","TEXT","TEXTAFTER","TEXTBEFORE","TEXTJOIN","TEXTSPLIT","TIME","TIMEVALUE","TOCOL",
  "TODAY","TOROW","TRANSPOSE","TREND","TRIM","TRIMMEAN","TRUE","TRUNC","TYPE","UNICHAR","UNICODE",
  "UNIQUE","UPPER","VALUE","VAR","VAR.P","VAR.S","VARA","VARP","VARPA","VDB","VLOOKUP","WEBSERVICE",
  "WEEKDAY","WEEKNUM","WEIBULL","WEIBULL.DIST","WORKDAY","WORKDAY.INTL","XIRR","XLOOKUP","XMATCH",
  "XNPV","XOR","YEAR","YEARFRAC","YIELD","YIELDDISC","YIELDMAT","Z.TEST",
] as const;

interface FormulaMonacoViewProps {
  isPopout: boolean;
  onOpenLegacy: () => void;
  embedded?: boolean;
}

export interface FormulaViewHandle {
  openOnly: () => Promise<void>;
  openAndPull: () => Promise<void>;
  pull: () => Promise<void>;
  apply: () => Promise<void>;
  beautify: () => Promise<void>;
  insertSelection: () => Promise<void>;
  getCurrentFormula: () => string;
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
type FormulaSubTab = "metadata" | "lambda-test" | "editor" | "wizard";

interface WizardLetVariable {
  id: string;
  name: string;
  expression: string;
}

interface PreparedWizardLetVariable extends WizardLetVariable {
  name: string;
  expression: string;
}

interface WizardOutputState {
  output: FormulaEvaluationResult | null;
  error: string;
}

const WIZARD_FINAL_OUTPUT_TARGET = "__wbm-let-final-output";
const DEFAULT_EDITOR_HEIGHT = 268;
const DEFAULT_POPOUT_EDITOR_HEIGHT = 360;
const MIN_EDITOR_HEIGHT = 220;
const MAX_EDITOR_HEIGHT = 760;

const buildLetFormula = (
  variables: Array<Pick<PreparedWizardLetVariable, "name" | "expression">>,
  outputExpression: string
): string => {
  const normalizedOutput = outputExpression.trim() || "0";
  if (variables.length === 0) {
    return `=${normalizedOutput}`;
  }
  return `=LET(${variables.map((item) => `${item.name},${item.expression}`).join(",")},${normalizedOutput})`;
};

const useStyles = makeStyles({
  root: { display: "grid", gap: "16px" },
  rootEmbedded: { gap: 0 },
  card: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    display: "grid",
    gap: "12px",
  },
  cardEmbedded: {
    borderRadius: 0,
    border: "none",
    boxShadow: "none",
    padding: "12px 0 0 0",
  },
  editorShell: {
    display: "grid",
    gap: "10px",
  },
  editorStage: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    overflow: "hidden",
    display: "grid",
    gridTemplateRows: "minmax(0, 1fr) auto",
    minHeight: 0,
  },
  editorToolbar: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  editorToolbarButtons: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    alignItems: "center",
  },
  editorToolbarMeta: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  sectionInfoRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "12px",
    flexWrap: "wrap",
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
    borderRadius: 0,
    border: "none",
    overflow: "hidden",
    minHeight: 0,
    backgroundColor: "#fff",
  },
  editorWrapResizable: {
    height: `${DEFAULT_EDITOR_HEIGHT}px`,
    resize: "vertical",
    overflow: "auto",
    minHeight: `${MIN_EDITOR_HEIGHT}px`,
    maxHeight: `min(${MAX_EDITOR_HEIGHT}px, 72vh)`,
  },
  editorWrapResizablePopout: {
    height: `${DEFAULT_POPOUT_EDITOR_HEIGHT}px`,
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
    resize: "none",
    padding: "12px",
    boxSizing: "border-box",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "14px",
    lineHeight: "1.6",
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
  editorFooter: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "10px",
    display: "grid",
    gap: "8px",
  },
  editorFooterPinned: {
    borderRadius: 0,
    border: "none",
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    boxShadow: "0 -6px 18px rgba(17,24,39,0.08)",
  },
  editorResizeHint: {
    fontSize: "11px",
    lineHeight: "15px",
    color: MODERN_TOKENS.colorTextMuted,
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
  wizardList: {
    display: "grid",
    gap: "10px",
  },
  wizardCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "12px",
    display: "grid",
    gap: "10px",
  },
  wizardCardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: "10px",
    flexWrap: "wrap",
  },
  wizardCardHeading: {
    display: "grid",
    gap: "4px",
    minWidth: 0,
  },
  wizardCardTitle: {
    fontSize: "13px",
    lineHeight: "18px",
    fontWeight: 700,
    color: MODERN_TOKENS.colorText,
  },
  wizardCardHint: {
    fontSize: "11px",
    lineHeight: "15px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  wizardCardActions: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  wizardCardFields: {
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  wizardAddRow: {
    display: "flex",
    justifyContent: "flex-start",
  },
  wizardOutputStack: {
    display: "grid",
    gap: "6px",
  },
  wizardOutputLabel: {
    fontSize: "11px",
    lineHeight: "15px",
    fontWeight: 600,
    color: MODERN_TOKENS.colorTextMuted,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
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
  fullHeight: { height: "100%" },
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

  // Special case: LET formula — format as paired name, value lines
  const letMatch = formula.match(/^=\s*LET\s*\(([\s\S]*)\)\s*$/i);
  if (letMatch) {
    const args = splitTopLevelComma(letMatch[1]);
    if (args.length >= 3) {
      const indent = "  ";
      const lines: string[] = ["=LET("];
      const calculationArg = args[args.length - 1];
      const pairArgs = args.slice(0, -1);
      for (let i = 0; i < pairArgs.length; i += 2) {
        const nameArg = pairArgs[i] ?? "";
        const valueArg = pairArgs[i + 1] ?? "";
        if (valueArg) {
          lines.push(`${indent}${nameArg}, ${valueArg},`);
        } else if (nameArg) {
          lines.push(`${indent}${nameArg},`);
        }
      }
      lines.push(`${indent}${calculationArg}`);
      lines.push(")");
      return lines.join("\n");
    }
  }

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

// ── F4 reference-lock cycling ─────────────────────────────────────────────────

type LockState = "none" | "abs" | "rowAbs" | "colAbs";

const getLockState = (cellRef: string): LockState | null => {
  const match = cellRef.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  if (!match) return null;
  const hasColLock = match[1] === "$";
  const hasRowLock = match[3] === "$";
  if (hasColLock && hasRowLock) return "abs";
  if (!hasColLock && hasRowLock) return "rowAbs";
  if (hasColLock && !hasRowLock) return "colAbs";
  return "none";
};

const applyLockState = (cellRef: string, state: LockState): string => {
  const match = cellRef.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
  if (!match) return cellRef;
  const col = match[2];
  const row = match[4];
  if (state === "abs") return `$${col}$${row}`;
  if (state === "rowAbs") return `${col}$${row}`;
  if (state === "colAbs") return `$${col}${row}`;
  return `${col}${row}`;
};

const nextLockState = (state: LockState): LockState => {
  if (state === "none") return "abs";
  if (state === "abs") return "rowAbs";
  if (state === "rowAbs") return "colAbs";
  return "none";
};

const cycleSingleCellLock = (cellRef: string): string => {
  const state = getLockState(cellRef);
  if (!state) return cellRef;
  return applyLockState(cellRef, nextLockState(state));
};

const cycleRangeLock = (rangeRef: string): string => {
  const rangeMatch = rangeRef.match(/^(\$?[A-Za-z]{1,3}\$?\d+)\s*:\s*(\$?[A-Za-z]{1,3}\$?\d+)$/);
  if (!rangeMatch) return rangeRef;
  const leftState = getLockState(rangeMatch[1]);
  const rightState = getLockState(rangeMatch[2]);
  const baseState = leftState ?? rightState;
  if (!baseState) return rangeRef;
  const nextState = nextLockState(baseState);
  return `${applyLockState(rangeMatch[1], nextState)}:${applyLockState(rangeMatch[2], nextState)}`;
};

const cycleCellOrRangeLockAtCursor = (
  text: string,
  cursor: number,
  selectionStart: number,
  selectionEnd: number
): { updated: string; nextStart: number; nextEnd: number } => {
  if (selectionEnd > selectionStart) {
    const selectedText = text.slice(selectionStart, selectionEnd);
    const isCell = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(selectedText);
    const isRange = /^\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+$/.test(selectedText);
    if (isCell || isRange) {
      const replacement = isRange ? cycleRangeLock(selectedText) : cycleSingleCellLock(selectedText);
      const updated = `${text.slice(0, selectionStart)}${replacement}${text.slice(selectionEnd)}`;
      return { updated, nextStart: selectionStart, nextEnd: selectionStart + replacement.length };
    }
  }

  const rangeRegex = /\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+/g;
  let rangeMatch: RegExpExecArray | null;
  while ((rangeMatch = rangeRegex.exec(text)) !== null) {
    const start = rangeMatch.index;
    const end = start + rangeMatch[0].length;
    if (cursor >= start && cursor <= end) {
      const replacement = cycleRangeLock(rangeMatch[0]);
      const updated = `${text.slice(0, start)}${replacement}${text.slice(end)}`;
      const nextPos = start + replacement.length;
      return { updated, nextStart: nextPos, nextEnd: nextPos };
    }
  }

  const cellRegex = /\$?[A-Za-z]{1,3}\$?\d+/g;
  let chosen: { start: number; end: number; value: string } | null = null;
  let cellMatch: RegExpExecArray | null;
  while ((cellMatch = cellRegex.exec(text)) !== null) {
    const start = cellMatch.index;
    const end = start + cellMatch[0].length;
    if (cursor >= start && cursor <= end) {
      chosen = { start, end, value: cellMatch[0] };
      break;
    }
    if (end < cursor) {
      chosen = { start, end, value: cellMatch[0] };
    }
  }

  if (!chosen) return { updated: text, nextStart: cursor, nextEnd: cursor };
  const replacement = cycleSingleCellLock(chosen.value);
  const updated = `${text.slice(0, chosen.start)}${replacement}${text.slice(chosen.end)}`;
  const nextPos = chosen.start + replacement.length;
  return { updated, nextStart: nextPos, nextEnd: nextPos };
};

// ─────────────────────────────────────────────────────────────────────────────

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
  ({ isPopout, onOpenLegacy, embedded = false }, ref) => {
    const shared = useModernSharedStyles();
    const styles = useStyles();
    const [formulaText, setFormulaText] = useState<string>("");
    const [status, setStatus] = useState<string>("");
    const [statusType, setStatusType] = useState<"success" | "error">("success");
    const [activeCell, setActiveCell] = useState<ActiveCellState | null>(null);
    const [formulaOutput, setFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
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
    const [wizardOutputTarget, setWizardOutputTarget] = useState<string>(WIZARD_FINAL_OUTPUT_TARGET);
    const [wizardVariableOutputs, setWizardVariableOutputs] = useState<Record<string, WizardOutputState>>({});
    const [wizardResultOutput, setWizardResultOutput] = useState<WizardOutputState>({
      output: null,
      error: "",
    });
    const [wizardFormulaOutput, setWizardFormulaOutput] = useState<WizardOutputState>({
      output: null,
      error: "",
    });
    const [wizardOutputBusy, setWizardOutputBusy] = useState<boolean>(false);
    const [lastWizardOutputAt, setLastWizardOutputAt] = useState<string>("");
    const [lambdaTestInputs, setLambdaTestInputs] = useState<string[]>([]);
    const [liveTestBusy, setLiveTestBusy] = useState<boolean>(false);
    const [liveTestError, setLiveTestError] = useState<string>("");
    const [lastTestedAt, setLastTestedAt] = useState<string>("");
    const [lambdaTestOutput, setLambdaTestOutput] = useState<FormulaEvaluationResult | null>(null);
    const [lambdaTestBusy, setLambdaTestBusy] = useState<boolean>(false);
    const [lambdaTestError, setLambdaTestError] = useState<string>("");
    const [lastLambdaTestedAt, setLastLambdaTestedAt] = useState<string>("");
    const [lastTestInvocation, setLastTestInvocation] = useState<string>("");
    const [activeSubTab, setActiveSubTab] = useState<FormulaSubTab | null>("editor");
    const [autoCaptureFormula, setAutoCaptureFormula] = useState<boolean>(false);

    const editorRef = useRef<Monaco.editor.IStandaloneCodeEditor | null>(null);
    const editorHostRef = useRef<HTMLDivElement | null>(null);
    const monacoRef = useRef<typeof Monaco | null>(null);
    const providerRef = useRef<Monaco.IDisposable | null>(null);
    const layoutFrameRef = useRef<number | null>(null);
    const namesRef = useRef<string[]>([]);
    const tablesRef = useRef<string[]>([]);
    const pendingEvalRequestsRef = useRef<
      Map<
        string,
        {
          resolve: (result: FormulaEvaluationResult) => void;
          reject: (error: Error) => void;
          timeoutId: number;
        }
      >
    >(new Map());
    const lastSyncedFormulaRef = useRef<string>("");
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
        return {
          formula: buildNamedFunctionFormula(current, stringifyArgs(effectiveNamedFunctionArgs)),
          error: "",
        };
      } catch (error) {
        return { formula: "", error: normalizeError(error) };
      }
    }, [effectiveNamedFunctionArgs, formulaText]);
    const wizardPreview = useMemo(() => {
      const returnExpression = wizardReturnExpression.trim() || "0";

      if (wizardTemplate === "LET") {
        const issuesById: Record<string, string> = {};
        const preparedVars: PreparedWizardLetVariable[] = [];
        const seen = new Set<string>();
        let firstError = "";

        wizardVariables.forEach((item) => {
          const name = item.name.trim();
          const expression = item.expression.trim();
          if (!name && !expression) {
            return;
          }
          if (!name || !expression) {
            const message = "Enter both a variable name and expression.";
            issuesById[item.id] = message;
            if (!firstError) {
              firstError = message;
            }
            return;
          }
          if (!LAMBDA_IDENTIFIER_PATTERN.test(name)) {
            const message = `Invalid variable name "${name}".`;
            issuesById[item.id] = message;
            if (!firstError) {
              firstError = message;
            }
            return;
          }
          const key = name.toUpperCase();
          if (seen.has(key)) {
            const message = `Duplicate variable name "${name}".`;
            issuesById[item.id] = message;
            if (!firstError) {
              firstError = message;
            }
            return;
          }
          seen.add(key);
          preparedVars.push({ id: item.id, name, expression });
        });

        if (!firstError && preparedVars.length === 0) {
          firstError = "LET builder requires at least one variable.";
        }

        const formula = firstError ? "" : buildLetFormula(preparedVars, returnExpression);
        let testingFormula = formula;
        let testingError = firstError;
        let outputLabel = "LET final expression";

        if (!firstError && wizardOutputTarget !== WIZARD_FINAL_OUTPUT_TARGET) {
          const selectedVariable = preparedVars.find((item) => item.id === wizardOutputTarget);
          if (selectedVariable) {
            testingFormula = buildLetFormula(preparedVars, selectedVariable.name);
            testingError = "";
            outputLabel = `Variable ${selectedVariable.name}`;
          } else {
            const selectedDraft = wizardVariables.find((item) => item.id === wizardOutputTarget);
            testingFormula = "";
            testingError = "Select a complete variable to test formula output.";
            outputLabel = selectedDraft?.name.trim()
              ? `Variable ${selectedDraft.name.trim()}`
              : "Selected variable";
          }
        }

        return {
          formula,
          error: firstError,
          testingFormula,
          testingError,
          outputLabel,
          preparedVars,
          issuesById,
        };
      }

      try {
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
        return {
          formula: lambdaFormula,
          error: "",
          testingFormula: lambdaFormula,
          testingError: "",
          outputLabel: "Formula output",
          preparedVars: [],
          issuesById: {},
        };
      } catch (error) {
        return {
          formula: "",
          error: normalizeError(error),
          testingFormula: "",
          testingError: normalizeError(error),
          outputLabel: "Formula output",
          preparedVars: [],
          issuesById: {},
        };
      }
    }, [
      wizardArgsInput,
      wizardOutputTarget,
      wizardReturnExpression,
      wizardTemplate,
      wizardVariables,
    ]);
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
      if (
        creationMode === "Function" &&
        !(authoringMode === "Wizard" && wizardTemplate === "LET")
      ) {
        tabs.push({ key: "lambda-test", label: "Lambda Test" });
      }
      tabs.push(
        authoringMode === "Editor"
          ? { key: "editor", label: "Editor" }
          : { key: "wizard", label: "Wizard" }
      );
      return tabs;
    }, [authoringMode, creationMode, wizardTemplate]);

    useEffect(() => {
      let disposed = false;
      void getActiveCellFormulaState()
        .then((state) => {
          if (!disposed) {
            setActiveCell(state);
          }
        })
        .catch(() => {
          // Ignore transient host states until the user pulls explicitly.
        });
      return () => {
        disposed = true;
      };
    }, []);

    useEffect(() => {
      if (!autoCaptureFormula) return undefined;
      let disposed = false;
      let busy = false;
      const timerId = window.setInterval(() => {
        if (disposed || busy) return;
        busy = true;
        void getActiveCellFormulaState()
          .then((state) => {
            if (disposed) return;
            setActiveCell(state);
            if (!state.hasFormula) return;
            if (editorRef.current?.hasTextFocus()) return;
            if (state.formula !== lastSyncedFormulaRef.current) {
              setFormulaText(state.formula);
              lastSyncedFormulaRef.current = state.formula;
            }
          })
          .catch(() => { /* ignore transient Excel states */ })
          .finally(() => { busy = false; });
      }, 1100);
      return () => {
        disposed = true;
        window.clearInterval(timerId);
      };
    }, [autoCaptureFormula]);

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
      if (wizardTemplate !== "LET") {
        setWizardOutputTarget(WIZARD_FINAL_OUTPUT_TARGET);
        return;
      }
      if (
        wizardOutputTarget !== WIZARD_FINAL_OUTPUT_TARGET &&
        !wizardVariables.some((item) => item.id === wizardOutputTarget)
      ) {
        setWizardOutputTarget(WIZARD_FINAL_OUTPUT_TARGET);
      }
    }, [wizardOutputTarget, wizardTemplate, wizardVariables]);

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

    const scheduleEditorLayout = useCallback(() => {
      if (typeof window === "undefined") {
        return;
      }
      if (layoutFrameRef.current !== null) {
        window.cancelAnimationFrame(layoutFrameRef.current);
      }
      layoutFrameRef.current = window.requestAnimationFrame(() => {
        layoutFrameRef.current = null;
        editorRef.current?.layout();
      });
    }, []);

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

    const pullFromActiveCell = useCallback(async (requireFormula: boolean) => {
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
    }, []);

    const applyToActiveCell = useCallback(async () => {
      const normalized = (editorRef.current?.getValue() ?? formulaText).trim();
      if (!normalized) {
        throw new Error("Formula editor is empty.");
      }
      const argsForApply = stringifyArgs(
        mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft)
      );
      const formulaToApply =
        creationMode === "Function"
          ? buildNamedFunctionFormula(normalized, argsForApply)
          : normalized;
      await applyFormulaToActiveCell(formulaToApply);
      const refreshed = await getActiveCellFormulaState();
      setActiveCell(refreshed);
      lastSyncedFormulaRef.current = refreshed.formula;
      setFormulaText(refreshed.formula);
    }, [creationMode, formulaText, namedFunctionArgDraft, namedFunctionArgs]);

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
      if (!isPopout || typeof Office === "undefined") {
        return undefined;
      }

      // Build readyMessage here — sent AFTER addHandlerAsync (see Bug 2 fix below).
      const readyMessage: DialogReadyMessage = {
        channel: DIALOG_RPC_CHANNEL,
        type: "ready",
      };

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
          const normalizedFormula = incomingFormula.startsWith("=")
            ? incomingFormula
            : `=${incomingFormula}`;
          setFormulaText(normalizedFormula);
          lastSyncedFormulaRef.current = normalizedFormula;
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

        pending.reject(
          new Error(typeof data.error === "string" ? data.error : "Live test failed.")
        );
      };

      // Bug 2 fix: send "ready" INSIDE the addHandlerAsync callback, not before it.
      // Previously, messageParent("ready") fired before the handler was registered, so the
      // host's immediate "open-formula" response arrived while we weren't listening — formula dropped.
      Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        handler,
        () => {
          try {
            Office.context.ui?.messageParent(JSON.stringify(readyMessage));
          } catch {
            // ignore if bridge not ready
          }
        }
      );

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
      if (authoringMode !== "Wizard" || wizardTemplate !== "LET") {
        setWizardVariableOutputs({});
        setWizardResultOutput({ output: null, error: "" });
        setWizardFormulaOutput({ output: null, error: "" });
        setWizardOutputBusy(false);
        setLastWizardOutputAt("");
        return undefined;
      }

      let disposed = false;
      const timerId = window.setTimeout(() => {
        setWizardOutputBusy(true);
        void (async () => {
          const nextVariableOutputs: Record<string, WizardOutputState> = {};
          const preparedPrefix: PreparedWizardLetVariable[] = [];
          const preparedMap = new Map(
            wizardPreview.preparedVars.map((item) => [item.id, item] as const)
          );

          for (const item of wizardVariables) {
            const name = item.name.trim();
            const expression = item.expression.trim();
            if (!name && !expression) {
              nextVariableOutputs[item.id] = {
                output: null,
                error: "Enter a variable name and expression to preview this step.",
              };
              continue;
            }

            const issue = wizardPreview.issuesById[item.id];
            if (issue) {
              nextVariableOutputs[item.id] = { output: null, error: issue };
              continue;
            }

            const preparedItem = preparedMap.get(item.id);
            if (!preparedItem) {
              nextVariableOutputs[item.id] = {
                output: null,
                error: "Complete this variable before previewing it.",
              };
              continue;
            }

            try {
              const result = await evaluateFormulaForLive(
                buildLetFormula([...preparedPrefix, preparedItem], preparedItem.name)
              );
              if (disposed) {
                return;
              }
              nextVariableOutputs[item.id] = {
                output: result,
                error: result.hasError
                  ? `Variable preview returned ${getFirstErrorCell(result)}.`
                  : "",
              };
            } catch (error) {
              if (disposed) {
                return;
              }
              nextVariableOutputs[item.id] = {
                output: null,
                error: normalizeError(error),
              };
            }

            preparedPrefix.push(preparedItem);
          }

          let nextResultOutput: WizardOutputState =
            wizardPreview.error || !wizardPreview.formula
              ? { output: null, error: wizardPreview.error }
              : { output: null, error: "" };

          if (!nextResultOutput.error && wizardPreview.formula) {
            try {
              const result = await evaluateFormulaForLive(wizardPreview.formula);
              if (disposed) {
                return;
              }
              nextResultOutput = {
                output: result,
                error: result.hasError ? `Result preview returned ${getFirstErrorCell(result)}.` : "",
              };
            } catch (error) {
              if (disposed) {
                return;
              }
              nextResultOutput = { output: null, error: normalizeError(error) };
            }
          }

          let nextFormulaOutput: WizardOutputState;
          if (wizardPreview.testingFormula && wizardPreview.testingFormula === wizardPreview.formula) {
            nextFormulaOutput = nextResultOutput;
          } else if (wizardPreview.testingError || !wizardPreview.testingFormula) {
            nextFormulaOutput = { output: null, error: wizardPreview.testingError };
          } else {
            try {
              const result = await evaluateFormulaForLive(wizardPreview.testingFormula);
              if (disposed) {
                return;
              }
              nextFormulaOutput = {
                output: result,
                error: result.hasError ? `Formula output returned ${getFirstErrorCell(result)}.` : "",
              };
            } catch (error) {
              if (disposed) {
                return;
              }
              nextFormulaOutput = { output: null, error: normalizeError(error) };
            }
          }

          if (disposed) {
            return;
          }

          setWizardVariableOutputs(nextVariableOutputs);
          setWizardResultOutput(nextResultOutput);
          setWizardFormulaOutput(nextFormulaOutput);
          setLastWizardOutputAt(new Date().toLocaleTimeString());
        })().finally(() => {
          if (!disposed) {
            setWizardOutputBusy(false);
          }
        });
      }, 320);

      return () => {
        disposed = true;
        window.clearTimeout(timerId);
      };
    }, [
      authoringMode,
      evaluateFormulaForLive,
      wizardPreview,
      wizardTemplate,
      wizardVariables,
    ]);

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
        scheduleEditorLayout();
        // formulaText state is now synced via the <Editor onChange> prop below.
        // Removed manual onDidChangeModelContent here to prevent double-firing.

        // F4 — cycle cell/range reference lock ($A$1 → A$1 → $A1 → A1 → …)
        const monaco = monacoRef.current;
        if (monaco) {
          editor.addCommand(monaco.KeyCode.F4, () => {
            const model = editor.getModel();
            if (!model) return;
            const currentText = editor.getValue();
            const position = editor.getPosition();
            const cursorOffset = position ? model.getOffsetAt(position) : 0;
            const sel = editor.getSelection();
            const selStart = sel
              ? model.getOffsetAt({ lineNumber: sel.startLineNumber, column: sel.startColumn })
              : cursorOffset;
            const selEnd = sel
              ? model.getOffsetAt({ lineNumber: sel.endLineNumber, column: sel.endColumn })
              : cursorOffset;
            const { updated, nextStart, nextEnd } = cycleCellOrRangeLockAtCursor(
              currentText,
              cursorOffset,
              selStart,
              selEnd
            );
            if (updated === currentText) return;
            editor.pushUndoStop();
            editor.executeEdits("wbm-f4-lock", [{ range: model.getFullModelRange(), text: updated }]);
            editor.pushUndoStop();
            const nextPos = model.getPositionAt(nextStart);
            const nextEndPos = model.getPositionAt(nextEnd);
            editor.setSelection({
              startLineNumber: nextPos.lineNumber,
              startColumn: nextPos.column,
              endLineNumber: nextEndPos.lineNumber,
              endColumn: nextEndPos.column,
            });
          });
        }
      },
      [scheduleEditorLayout]
    );

    useEffect(() => {
      registerCompletionProvider();
      return () => {
        providerRef.current?.dispose();
      };
    }, [registerCompletionProvider]);

    useEffect(() => {
      if (typeof window === "undefined") {
        return undefined;
      }

      const host = editorHostRef.current;
      const handleWindowResize = () => {
        scheduleEditorLayout();
      };

      window.addEventListener("resize", handleWindowResize);

      let observer: ResizeObserver | null = null;
      if (host && typeof ResizeObserver !== "undefined") {
        observer = new ResizeObserver(() => {
          scheduleEditorLayout();
        });
        observer.observe(host);
      }

      scheduleEditorLayout();

      return () => {
        observer?.disconnect();
        window.removeEventListener("resize", handleWindowResize);
        if (layoutFrameRef.current !== null) {
          window.cancelAnimationFrame(layoutFrameRef.current);
          layoutFrameRef.current = null;
        }
      };
    }, [scheduleEditorLayout]);

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

    const buildSaveFailureMessage = (error: unknown, functionName: string): string => {
      const source = normalizeError(error);
      const tips: string[] = [];
      if (
        !EXCEL_NAME_PATTERN.test(functionName) ||
        CELL_REFERENCE_LIKE_PATTERN.test(functionName)
      ) {
        tips.push("Use a valid Excel name (starts with letter/_ and cannot look like A1).");
      }
      if (
        namedFunctionArgPreview.length > 0 &&
        namedFunctionArgPreview.some((arg) => !LAMBDA_IDENTIFIER_PATTERN.test(arg))
      ) {
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
        if (
          !EXCEL_NAME_PATTERN.test(functionName) ||
          CELL_REFERENCE_LIKE_PATTERN.test(functionName)
        ) {
          throw new Error(
            "Function name must be a valid Excel defined name and cannot look like a cell reference."
          );
        }
        const argsForSave = stringifyArgs(
          mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft)
        );
        setNamedFunctionArgs(argsForSave);
        setNamedFunctionArgDraft("");
        const lambdaFormula = buildNamedFunctionFormula(current, argsForSave);
        await saveNamedFunction(functionName, lambdaFormula);
        await loadSuggestions();
        const saved = namesRef.current.some(
          (item) => item.toUpperCase() === functionName.toUpperCase()
        );
        if (!saved) {
          throw new Error(
            "Save request completed, but the function is not visible yet. Refresh and try again."
          );
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
      const argsForPopout = stringifyArgs(
        mergeArgsWithDraft(namedFunctionArgs, namedFunctionArgDraft)
      );
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
        {
          id: `var-${Date.now().toString()}-${Math.random().toString(16).slice(2)}`,
          name: "",
          expression: "",
        },
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
      setLambdaTestBusy(true);
      setLambdaTestError("");
      try {
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
          setLambdaTestOutput(result);
          setLastLambdaTestedAt(new Date().toLocaleTimeString());
          if (result.hasError) {
            const errorText = getFirstErrorCell(result);
            setLambdaTestError(`Lambda test returned ${errorText}.`);
            throw new Error(
              `Lambda invocation returned ${errorText}. Check arguments and LET/LAMBDA structure.`
            );
          }
          setLambdaTestError("");
        });
      } finally {
        setLambdaTestBusy(false);
      }
    };

    const handleNamedArgInputKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
      if (event.key === " " || event.key === "Enter" || event.key === ",") {
        event.preventDefault();
        commitNamedFunctionArgDraft();
        return;
      }
      if (
        event.key === "Backspace" &&
        !namedFunctionArgDraft.trim() &&
        namedFunctionArgPreview.length > 0
      ) {
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
      getCurrentFormula: () => (editorRef.current?.getValue() ?? formulaText).trim(),
    }));

    const renderOutputPanel = ({
      title,
      idleMessage,
      busyMessage,
      successPrefix,
      output,
      busy,
      error,
      testedAt,
      showInvocation = false,
      className = "",
    }: {
      title: string;
      idleMessage: string;
      busyMessage: string;
      successPrefix: string;
      output: FormulaEvaluationResult | null;
      busy: boolean;
      error: string;
      testedAt: string;
      showInvocation?: boolean;
      className?: string;
    }) => {
      const statusText = busy
        ? busyMessage
        : error
          ? `Last calculation failed${testedAt ? ` at ${testedAt}` : ""}`
          : testedAt
            ? `${successPrefix}${testedAt ? ` at ${testedAt}` : ""}`
            : idleMessage;

      return (
        <div className={className ? `${styles.editorFooter} ${className}` : styles.editorFooter}>
          <div className={styles.liveTestHeader}>
            <Text className={shared.cardTitle}>{title}</Text>
            <Text className={shared.mutedText}>{statusText}</Text>
          </div>

          {showInvocation && lastTestInvocation ? (
            <Text className={styles.invocationText}>Invocation: {lastTestInvocation}</Text>
          ) : null}
          {error ? <Text className={shared.errorText}>{error}</Text> : null}

          <div className={styles.outputWrap}>
            {output ? (
              <table className={styles.outputTable}>
                <tbody>
                  {output.values.map((row, rowIndex) => (
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
                <Text className={shared.mutedText}>{idleMessage}</Text>
              </div>
            )}
          </div>
        </div>
      );
    };

    const renderCompactOutput = (
      state: WizardOutputState | undefined,
      idleMessage: string,
      busyMessage = "Updating preview..."
    ) => {
      const message = wizardOutputBusy ? busyMessage : idleMessage;

      return (
        <div className={styles.wizardOutputStack}>
          {state?.error ? <Text className={shared.errorText}>{state.error}</Text> : null}
          <div className={styles.outputWrap}>
            {state?.output ? (
              <table className={styles.outputTable}>
                <tbody>
                  {state.output.values.map((row, rowIndex) => (
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
                <Text className={shared.mutedText}>
                  {state?.error ? idleMessage : message}
                </Text>
              </div>
            )}
          </div>
        </div>
      );
    };

    const editorLineHeight = 22;

    return (
      <div className={`${styles.root} ${embedded ? styles.rootEmbedded : ""}`}>
        {!isPopout && !embedded ? (
          <div>
            <Text className={shared.sectionTitle}>Formulas</Text>
            <Text className={shared.sectionSubtitle}>
              Manage and edit workbook formulas with live test output.
            </Text>
          </div>
        ) : null}

        <div className={`${styles.card} ${embedded ? styles.cardEmbedded : ""}`}>
          <div className={styles.sectionInfoRow}>
            <Text className={shared.mutedText}>
              Active Cell: {activeCell ? `${activeCell.sheet}!${activeCell.address}` : "not linked"}
              {loadingMetadata ? " • loading suggestions..." : ""}
            </Text>
            <Text className={shared.mutedText}>
              {autoCaptureFormula
                ? "Auto-capture active — formula updates when you select a cell with a formula."
                : "Manual sync. Use Pull Active Formula or enable Auto-capture."}
            </Text>
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
                    <Button
                      appearance="primary"
                      size="small"
                      onClick={() => void runLambdaInputTest()}
                      disabled={lambdaTestBusy}
                    >
                      Run Lambda Test
                    </Button>
                  </div>
                  {renderOutputPanel({
                    title: "Lambda Test Output",
                    idleMessage: "Run a lambda test to render results here.",
                    busyMessage: "Running lambda test...",
                    successPrefix: "Lambda test updated",
                    output: lambdaTestOutput,
                    busy: lambdaTestBusy,
                    error: lambdaTestError,
                    testedAt: lastLambdaTestedAt,
                    showInvocation: true,
                  })}
                </>
              ) : null}

              {activeSubTab === "editor" && authoringMode === "Editor" ? (
                <div className={styles.editorShell}>
                  <div className={styles.editorToolbar}>
                    <div className={styles.editorToolbarButtons}>
                      <Button
                        size="small"
                        onClick={() =>
                          void runAction("Pull active formula", async () => {
                            await pullFromActiveCell(true);
                          })
                        }
                      >
                        Pull Active Formula
                      </Button>
                      <Checkbox
                        label="Auto-capture"
                        title="Automatically pull the active cell formula whenever you select a cell that contains one"
                        checked={autoCaptureFormula}
                        onChange={(_, data) => setAutoCaptureFormula(Boolean(data.checked))}
                      />
                      <Button
                        size="small"
                        onClick={() => void runAction("Beautify formula", beautify)}
                      >
                        Beautify
                      </Button>
                      <Button
                        size="small"
                        onClick={() => void runAction("Insert selection", insertSelection)}
                      >
                        Insert Selection
                      </Button>
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
                      {!isPopout ? (
                        <Button
                          size="small"
                          onClick={() =>
                            void runAction(
                              "Open formula editor popout",
                              openPopoutWithCurrentState
                            )
                          }
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
                    <div className={styles.editorToolbarMeta}>
                      <Text className={shared.mutedText}>
                        {activeCell
                          ? `Editing against ${activeCell.sheet}!${activeCell.address}`
                          : "Pull a worksheet formula to load it into the editor."}
                      </Text>
                      <Text className={styles.editorResizeHint}>
                        Drag the editor's bottom edge to resize it. Calculate output stays pinned below.
                      </Text>
                    </div>
                  </div>

                  <div className={styles.editorStage}>
                    <div
                      ref={editorHostRef}
                      className={`${styles.editorWrap} ${styles.editorWrapResizable} ${
                        isPopout ? styles.editorWrapResizablePopout : ""
                      }`}
                    >
                      {editorLoadError ? (
                        <textarea
                          aria-label="Formula editor fallback"
                          className={`${styles.fallbackEditor} ${styles.fullHeight}`}
                          spellCheck={false}
                          rows={10}
                          value={formulaText}
                          onChange={(event) => setFormulaText(event.target.value)}
                        />
                      ) : editorReady ? (
                        <Editor
                          height="100%"
                          language={MONACO_LANGUAGE_ID}
                          value={formulaText}
                          onChange={(value) => setFormulaText(value ?? "")}
                          beforeMount={beforeMount}
                          onMount={onMount}
                          theme="vs"
                          options={{
                            minimap: { enabled: false },
                            scrollBeyondLastLine: false,
                            fontSize: 14,
                            lineHeight: editorLineHeight,
                            lineNumbers: "on",
                            wordWrap: "on",
                            suggestOnTriggerCharacters: true,
                            quickSuggestions: { other: true, comments: false, strings: false },
                            wordBasedSuggestions: "off",
                            tabCompletion: "on",
                            padding: {
                              top: 12,
                              bottom: 12,
                            },
                            suggest: {
                              showWords: false,
                              showSnippets: true,
                              preview: true,
                            },
                          }}
                        />
                      ) : (
                        <div className={`${styles.editorLoading} ${styles.fullHeight}`}>
                          <Text className={shared.mutedText}>Loading formula editor...</Text>
                        </div>
                      )}
                    </div>

                    {renderOutputPanel({
                      title: "Calculate Output",
                      idleMessage: "Results update here as you edit the formula.",
                      busyMessage: "Calculating...",
                      successPrefix: "Updated",
                      output: formulaOutput,
                      busy: liveTestBusy,
                      error: liveTestError,
                      testedAt: lastTestedAt,
                      className: styles.editorFooterPinned,
                    })}
                  </div>
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
                  <Text className={styles.modalLabel}>Variables</Text>
                  <div className={styles.wizardList}>
                    {wizardVariables.map((item, index) => (
                      <div key={item.id} className={styles.wizardCard}>
                        <div className={styles.wizardCardHeader}>
                          <div className={styles.wizardCardHeading}>
                            <Text className={styles.wizardCardTitle}>Variable {index + 1}</Text>
                            <Text className={styles.wizardCardHint}>
                              {wizardTemplate === "LET"
                                ? "Preview each LET step and optionally route formula output through this variable."
                                : "Optional LET variable for the lambda body."}
                            </Text>
                          </div>
                          <div className={styles.wizardCardActions}>
                            {wizardTemplate === "LET" ? (
                              <Checkbox
                                checked={wizardOutputTarget === item.id}
                                label="Use as formula output"
                                onChange={(_, data) =>
                                  setWizardOutputTarget(
                                    data.checked ? item.id : WIZARD_FINAL_OUTPUT_TARGET
                                  )
                                }
                              />
                            ) : null}
                            <Button
                              appearance="subtle"
                              size="small"
                              icon={<Delete20Regular />}
                              title="Delete variable"
                              aria-label={`Delete variable ${index + 1}`}
                              onClick={() => removeWizardVariable(item.id)}
                            />
                          </div>
                        </div>
                        <div className={styles.wizardCardFields}>
                          <div>
                            <Text className={styles.modalLabel}>Variable name</Text>
                            <Input
                              placeholder="greeting"
                              value={item.name}
                              onChange={(_, data) =>
                                updateWizardVariable(item.id, "name", data.value)
                              }
                            />
                          </div>
                          <div>
                            <Text className={styles.modalLabel}>Expression</Text>
                            <Input
                              placeholder={'"Hello"'}
                              value={item.expression}
                              onChange={(_, data) =>
                                updateWizardVariable(item.id, "expression", data.value)
                              }
                            />
                          </div>
                        </div>
                        {wizardTemplate === "LET" ? (
                          <div className={styles.wizardOutputStack}>
                            <Text className={styles.wizardOutputLabel}>Variable Output</Text>
                            {renderCompactOutput(
                              wizardVariableOutputs[item.id],
                              "This variable preview appears once the name and expression are complete."
                            )}
                          </div>
                        ) : null}
                      </div>
                    ))}
                  </div>
                  <div className={styles.wizardAddRow}>
                    <Button size="small" onClick={addWizardVariable}>
                      + Add Variable
                    </Button>
                  </div>
                  {wizardTemplate === "LET" ? (
                    <>
                      <div className={styles.wizardCard}>
                        <div className={styles.wizardCardHeader}>
                          <div className={styles.wizardCardHeading}>
                            <Text className={styles.wizardCardTitle}>Final Expression</Text>
                            <Text className={styles.wizardCardHint}>
                              This is the real LET return expression that will be built into the
                              editor.
                            </Text>
                          </div>
                          <div className={styles.wizardCardActions}>
                            <Checkbox
                              checked={wizardOutputTarget === WIZARD_FINAL_OUTPUT_TARGET}
                              label="Use as formula output"
                              onChange={(_, data) =>
                                setWizardOutputTarget(
                                  data.checked
                                    ? WIZARD_FINAL_OUTPUT_TARGET
                                    : WIZARD_FINAL_OUTPUT_TARGET
                                )
                              }
                            />
                          </div>
                        </div>
                        <div>
                          <Text className={styles.modalLabel}>Expression</Text>
                          <Input
                            placeholder="greeting&audience"
                            value={wizardReturnExpression}
                            onChange={(_, data) => setWizardReturnExpression(data.value)}
                          />
                        </div>
                        <div className={styles.wizardOutputStack}>
                          <Text className={styles.wizardOutputLabel}>Final Expression Output</Text>
                          {renderCompactOutput(
                            wizardResultOutput,
                            "The LET final expression result will preview here."
                          )}
                        </div>
                      </div>
                      {renderOutputPanel({
                        title: `Formula Output: ${wizardPreview.outputLabel}`,
                        idleMessage: "Formula output updates as you switch output targets.",
                        busyMessage: "Updating formula output...",
                        successPrefix: "Updated",
                        output: wizardFormulaOutput.output,
                        busy: wizardOutputBusy,
                        error: wizardFormulaOutput.error,
                        testedAt: lastWizardOutputAt,
                      })}
                    </>
                  ) : (
                    <div>
                      <Text className={styles.modalLabel}>Return expression</Text>
                      <Input
                        placeholder="Result expression"
                        value={wizardReturnExpression}
                        onChange={(_, data) => setWizardReturnExpression(data.value)}
                      />
                    </div>
                  )}
                  <div>
                    <Text className={styles.modalLabel}>Wizard preview</Text>
                    {wizardTemplate === "LET" ? (
                      <Text className={shared.mutedText}>
                        Testing output target: {wizardPreview.outputLabel}
                      </Text>
                    ) : null}
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

            </div>
          ) : null}

          {status ? <Text className={statusClass}>{status}</Text> : null}
        </div>

        {showSaveFunctionModal ? (
          <div className={styles.modalBackdrop}>
            <div className={styles.modal}>
              <Text className={shared.cardTitle}>Save As Named Function</Text>
              <Text className={shared.cardSubtitle}>
                Add function metadata and arguments. If the editor formula is not already a LAMBDA
                expression, it will be wrapped automatically.
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
                <Text className={shared.mutedText}>
                  Descriptions are not persisted by Excel named formulas.
                </Text>
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
                  <Text className={shared.mutedText}>
                    No arguments set. The editor formula will be used as LAMBDA body.
                  </Text>
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
