import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, Select, Text, makeStyles, tokens } from "@fluentui/react-components";
import {
  Add20Regular,
  Checkmark20Regular,
  CloudArrowDown20Regular,
  Dismiss20Regular,
  Edit20Regular,
  Filter20Regular,
  Grid20Regular,
  Save20Regular,
  TextAlignJustify20Regular,
} from "@fluentui/react-icons";
import {
  CellStylePreset,
  FormulaEvaluationResult,
  InsertableShapeType,
  InsertedShapeRecord,
  NamedRangeRecord,
  TableRecord,
  addShapeOnActiveCell,
  addRectangleShape,
  applyAccentFill,
  alignShapeToSelection,
  applyCellStylePreset,
  applyFormulaToActiveCell,
  applyFormulaToShapeAnchorCell,
  applyTableStyle,
  deleteNamedRangeWithOptions,
  getActiveCellFormulaState,
  getCurrentSelectionAddress,
  getNamedRangeValueText,
  getNamedRanges,
  getShapeEditorRecord,
  getTables,
  insertText,
  openFormulaEditorPopout,
  moveNamedRange,
  moveShapeToSelection,
  nudgeShape,
  renameShape,
  refreshPivotTables,
  evaluateFormula,
  setShapeGeometricType,
  setShapeZOrder,
  selectNamedRangeAddress,
  toggleGridlines,
  ShapeFormatOptions,
  ShapeEditorRecord,
  updateShapePosition,
  updateShapeFormatting,
  updateTableName,
  updateNamedRange,
} from "../taskpane";
import { NavigationTarget } from "../navigation";

type StatusType = "success" | "error";
type PrimaryTab = "Names" | "Format" | "Sandbox";
type SortColumn = "name" | "address" | "sheet" | "scope" | "type";
type TableSortColumn = "name" | "address" | "sheet" | "scope";
type SortDirection = "asc" | "desc";
type CaseTransform = "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE";
type FormulaColorScheme = "Advanced" | "Classic" | "Monochrome";

interface LegacyAppProps {
  initialTarget?: NavigationTarget;
  bridgeMode?: boolean;
  onExitBridge?: () => void;
}

interface SecondaryTab {
  id: string;
  label: string;
}

interface EditState {
  open: boolean;
  record: NamedRangeRecord | null;
  name: string;
  address: string;
  fallbackSheet: string;
  caseTransform: CaseTransform;
}

interface BulkState {
  open: boolean;
  operation: "Add" | "Remove" | "Replace" | "Delete";
  position: "Prefix" | "Suffix";
  delimiter: "none" | "underscore" | "dot";
  textValue: string;
  replaceWith: string;
  enableSecondEdit: boolean;
  secondOperation: "Add" | "Remove" | "Replace";
  secondPosition: "Prefix" | "Suffix";
  secondDelimiter: "none" | "underscore" | "dot";
  secondTextValue: string;
  secondReplaceWith: string;
  caseTransform: CaseTransform;
  secondCaseTransform: CaseTransform;
  deleteName: boolean;
  deleteValues: boolean;
}

interface BulkPreviewRow {
  id: string;
  oldName: string;
  newName: string;
  scopeType: "Workbook" | "Worksheet";
  scope: string;
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

interface TableBulkState {
  open: boolean;
  operation: "Add" | "Remove" | "Replace";
  position: "Prefix" | "Suffix";
  delimiter: "none" | "underscore" | "dot";
  textValue: string;
  replaceWith: string;
  caseTransform: CaseTransform;
}

interface FunctionDefinition {
  name: string;
  description: string;
  params: string[];
}

interface ActiveFormulaPrompt {
  sheet: string;
  address: string;
  formula: string;
  hasFormula: boolean;
}

interface FormulaSuggestion {
  name: string;
  kind: "Function" | "Name" | "Table" | "Local";
  signature?: string;
}

interface FormulaReferenceHint {
  kind: "Function" | "Name" | "Table" | "Local";
  name: string;
  detail: string;
}

type FormulaTokenKind =
  | "function"
  | "name"
  | "table"
  | "local"
  | "string"
  | "number"
  | "operator"
  | "default";

interface FormulaColorPalette {
  function: string;
  name: string;
  table: string;
  local: string;
  string: string;
  number: string;
  operator: string;
  default: string;
}

interface CustomFormulaScheme {
  id: string;
  name: string;
  palette: FormulaColorPalette;
}

interface ShapeEditorState {
  shapeName: string;
  shapeType: InsertableShapeType;
  bindingMode: "StaticText" | "NamedRange" | "Formula";
  geometricShapeType: string;
  left: number;
  top: number;
  zOrderPosition: number;
  fillColor: string;
  outlineColor: string;
  fontColor: string;
  text: string;
  valueBinding: string;
  formula: string;
  lineWeight: number;
  fillTransparency: number;
  width: number;
  height: number;
  rotation: number;
  textHorizontalAlignment: "Left" | "Center" | "Right";
  textVerticalAlignment: "Top" | "Middle" | "Bottom";
  fontSize: number;
  bold: boolean;
  italic: boolean;
  lockAspectRatio: boolean;
  selectedThemeId: string;
}

interface ShapeTheme {
  id: string;
  label: string;
  themeColors: string[];
  standardColors: string[];
}

const DEFAULT_TABLE_STYLE = "TableStyleMedium2";
const FORMULA_TOKEN_KEYS: Array<keyof FormulaColorPalette> = [
  "function",
  "name",
  "table",
  "local",
  "string",
  "number",
  "operator",
  "default",
];
const FORMULA_TOKEN_LABELS: Record<keyof FormulaColorPalette, string> = {
  function: "Function",
  name: "Name",
  table: "Table",
  local: "Local variable",
  string: "String",
  number: "Number",
  operator: "Operator",
  default: "Default",
};
const FORMULA_COLOR_SCHEMES: Record<FormulaColorScheme, FormulaColorPalette> = {
  Advanced: {
    function: "#143a8a",
    name: "#295dbf",
    table: "#2f6b3f",
    local: "#3d7ae6",
    string: "#8b2f1f",
    number: "#6a3cc7",
    operator: "#2a2a2a",
    default: "#111111",
  },
  Classic: {
    function: "#1f4e79",
    name: "#4472c4",
    table: "#548235",
    local: "#5b9bd5",
    string: "#9e480e",
    number: "#7030a0",
    operator: "#3f3f3f",
    default: "#111111",
  },
  Monochrome: {
    function: "#1f1f1f",
    name: "#303030",
    table: "#3a3a3a",
    local: "#4a4a4a",
    string: "#5a5a5a",
    number: "#444444",
    operator: "#2a2a2a",
    default: "#121212",
  },
};

const WORKBOOK_THEME_SWATCHES = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
  "#264478",
  "#9E480E",
  "#636363",
  "#997300",
];
const NO_FILL_COLOR_TOKEN = "__NO_FILL__";
const OPEN_FORMULA_EDITOR_SIGNAL_KEY = "wbm.openFormulaEditor.request";
const OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL = "open-and-pull";
const FORMULA_PULL_SIGNAL = "formula-pull";
const FORMULA_APPLY_SIGNAL = "formula-apply";
const FORMULA_BEAUTIFY_SIGNAL = "formula-beautify";
const FORMULA_INSERT_SELECTION_SIGNAL = "formula-insert-selection";

const SHAPE_THEMES: ShapeTheme[] = [
  {
    id: "theme1",
    label: "Theme1",
    themeColors: ["#000000", "#FFFFFF", "#1F4E79", "#4F81BD", "#9BBB59", "#8064A2", "#4BACC6", "#F79646", "#C0504D", "#7F7F7F"],
    standardColors: ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"],
  },
  {
    id: "office",
    label: "Office",
    themeColors: ["#000000", "#FFFFFF", "#1F497D", "#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646", "#7F7F7F"],
    standardColors: ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"],
  },
  {
    id: "gallery",
    label: "Gallery",
    themeColors: ["#2F2F2F", "#FFFFFF", "#6D4A2C", "#A37446", "#9C7A47", "#7D8D4E", "#5E8D93", "#6B6B8A", "#C0672D", "#8E8E8E"],
    standardColors: ["#9E480E", "#C55A11", "#F4B183", "#FFD966", "#A9D18E", "#70AD47", "#5B9BD5", "#2E75B6", "#1F3864", "#5F497A"],
  },
];

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));

const mixColor = (hex: string, target: string, factor: number) => {
  const clean = hex.replace("#", "");
  const tgt = target.replace("#", "");
  const r1 = parseInt(clean.slice(0, 2), 16);
  const g1 = parseInt(clean.slice(2, 4), 16);
  const b1 = parseInt(clean.slice(4, 6), 16);
  const r2 = parseInt(tgt.slice(0, 2), 16);
  const g2 = parseInt(tgt.slice(2, 4), 16);
  const b2 = parseInt(tgt.slice(4, 6), 16);
  const r = Math.round(r1 + (r2 - r1) * factor);
  const g = Math.round(g1 + (g2 - g1) * factor);
  const b = Math.round(b1 + (b2 - b1) * factor);
  return `#${[r, g, b]
    .map((v) => clamp(v, 0, 255).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
};

const buildThemeToneGrid = (baseColors: string[]) => {
  const tintTargets = ["#FFFFFF", "#FFFFFF", "#FFFFFF", "#000000", "#000000"];
  const tintFactors = [0.82, 0.58, 0.32, 0.22, 0.42];
  return tintTargets.map((target, idx) =>
    baseColors.map((color) => mixColor(color, target, tintFactors[idx]))
  );
};

const parseThemeColorList = (value: string): string[] =>
  value
    .split(",")
    .map((item) => item.trim().toUpperCase())
    .filter((item) => /^#[0-9A-F]{6}$/.test(item))
    .slice(0, 10);

const shapeColorInputValue = (value: string, fallback: string): string =>
  value === NO_FILL_COLOR_TOKEN ? fallback : value;

const NAMES_TABS: SecondaryTab[] = [
  { id: "ranges", label: "Ranges" },
  { id: "tables", label: "Tables" },
  { id: "functions", label: "Formulas" },
  { id: "pivot-tables", label: "Pivot Tables" },
];

const FORMAT_TABS: SecondaryTab[] = [
  { id: "cells", label: "Cells" },
  { id: "styles", label: "Styles" },
  { id: "shapes", label: "Shapes/Text Boxes" },
  { id: "view", label: "View" },
];

const SANDBOX_TABS: SecondaryTab[] = [
  { id: "playground", label: "Playground" },
  { id: "actions", label: "Actions" },
  { id: "debug", label: "Debug" },
];

const cloneFormulaPalette = (palette: FormulaColorPalette): FormulaColorPalette => ({ ...palette });

const RANGE_COLUMNS: SortColumn[] = ["name", "address", "sheet", "scope", "type"];
const TABLE_COLUMNS: TableSortColumn[] = ["name", "address", "sheet", "scope"];
const FUNCTION_DEFINITIONS: FunctionDefinition[] = [
  { name: "SUM", description: "Add values.", params: ["number1", "[number2]"] },
  { name: "AVERAGE", description: "Average values.", params: ["number1", "[number2]"] },
  { name: "IF", description: "Conditional branch.", params: ["logical_test", "value_if_true", "value_if_false"] },
  { name: "XLOOKUP", description: "Lookup with exact/default matching.", params: ["lookup_value", "lookup_array", "return_array", "[if_not_found]"] },
  { name: "FILTER", description: "Filter array by include condition.", params: ["array", "include", "[if_empty]"] },
  { name: "INDEX", description: "Return value by row/column.", params: ["array", "row_num", "[column_num]"] },
  { name: "MATCH", description: "Return relative position.", params: ["lookup_value", "lookup_array", "[match_type]"] },
  { name: "TEXTJOIN", description: "Join values with delimiter.", params: ["delimiter", "ignore_empty", "text1", "[text2]"] },
  { name: "LET", description: "Assign names inside formula.", params: ["name1", "value1", "calculation"] },
  { name: "LAMBDA", description: "Define custom function.", params: ["parameter1", "[parameter2]", "calculation"] },
];

const EXCEL_FUNCTION_NAMES = [
  "ABS","ACCRINT","ACCRINTM","ACOS","ACOSH","ACOT","ACOTH","ADDRESS","AGGREGATE","AMORDEGRC","AMORLINC","AND","ARABIC","AREAS","ASC","ASIN","ASINH","ATAN","ATAN2","ATANH","AVEDEV","AVERAGE","AVERAGEA","AVERAGEIF","AVERAGEIFS","BAHTTEXT","BASE","BESSELI","BESSELJ","BESSELK","BESSELY","BETA.DIST","BETA.INV","BETADIST","BETAINV","BIN2DEC","BIN2HEX","BIN2OCT","BINOM.DIST","BINOM.DIST.RANGE","BINOM.INV","BINOMDIST","BITAND","BITLSHIFT","BITOR","BITRSHIFT","BITXOR","CEILING","CEILING.MATH","CEILING.PRECISE","CELL","CHAR","CHISQ.DIST","CHISQ.DIST.RT","CHISQ.INV","CHISQ.INV.RT","CHISQ.TEST","CHISQDIST","CHISQINV","CHITEST","CHOOSE","CLEAN","CODE","COLUMN","COLUMNS","COMBIN","COMBINA","COMPLEX","CONCAT","CONCATENATE","CONFIDENCE","CONFIDENCE.NORM","CONFIDENCE.T","CONVERT","CORREL","COS","COSH","COT","COTH","COUNT","COUNTA","COUNTBLANK","COUNTIF","COUNTIFS","COUPDAYBS","COUPDAYS","COUPDAYSNC","COUPNCD","COUPNUM","COUPPCD","COVARIANCE.P","COVARIANCE.S","COVAR","CRITBINOM","CSC","CSCH","CUBEKPIMEMBER","CUBEMEMBER","CUBEMEMBERPROPERTY","CUBERANKEDMEMBER","CUBESET","CUBESETCOUNT","CUBEVALUE","CUMIPMT","CUMPRINC","DATE","DATEDIF","DATEVALUE","DAVERAGE","DAY","DAYS","DAYS360","DB","DCOUNT","DCOUNTA","DDB","DEC2BIN","DEC2HEX","DEC2OCT","DECIMAL","DEGREES","DELTA","DEVSQ","DGET","DISC","DMAX","DMIN","DOLLAR","DOLLARDE","DOLLARFR","DPRODUCT","DSTDEV","DSTDEVP","DSUM","DURATION","DVAR","DVARP","ECMA.CEILING","EDATE","EFFECT","EOMONTH","ERF","ERF.PRECISE","ERFC","ERFC.PRECISE","ERROR.TYPE","EUROCONVERT","EVEN","EXACT","EXP","EXPON.DIST","EXPONDIST","FACT","FACTDOUBLE","FALSE","FDIST","FILTER","FIND","FINDB","FISHER","FISHERINV","FIXED","FLOOR","FLOOR.MATH","FLOOR.PRECISE","FORECAST","FORECAST.ETS","FORECAST.ETS.CONFINT","FORECAST.ETS.SEASONALITY","FORECAST.ETS.STAT","FORECAST.LINEAR","FORMULATEXT","FREQUENCY","FTEST","FV","FVSCHEDULE","GAMMA","GAMMA.DIST","GAMMA.INV","GAMMADIST","GAMMAINV","GAMMALN","GAMMALN.PRECISE","GAUSS","GCD","GEOMEAN","GESTEP","GETPIVOTDATA","GROWTH","HARMEAN","HEX2BIN","HEX2DEC","HEX2OCT","HLOOKUP","HOUR","HYPERLINK","HYPGEOM.DIST","HYPGEOMDIST","IF","IFERROR","IFNA","IFS","IMABS","IMAGINARY","IMARGUMENT","IMCONJUGATE","IMCOS","IMCOSH","IMCOT","IMCSC","IMCSCH","IMDIV","IMEXP","IMLN","IMLOG10","IMLOG2","IMPOWER","IMPRODUCT","IMREAL","IMSEC","IMSECH","IMSIN","IMSINH","IMSQRT","IMSUB","IMSUM","INDEX","INDIRECT","INFO","INT","INTERCEPT","INTRATE","IPMT","IRR","ISBLANK","ISERR","ISERROR","ISEVEN","ISFORMULA","ISLOGICAL","ISNA","ISNONTEXT","ISNUMBER","ISODD","ISPMT","ISREF","ISTEXT","KURT","LARGE","LAMBDA","LCM","LEFT","LEFTB","LEN","LENB","LET","LINEST","LN","LOG","LOG10","LOGEST","LOGINV","LOGNORM.DIST","LOGNORM.INV","LOGNORMDIST","LOOKUP","LOWER","MATCH","MAX","MAXA","MAXIFS","MDETERM","MDURATION","MEDIAN","MID","MIDB","MIN","MINA","MINIFS","MINUTE","MINVERSE","MIRR","MMULT","MOD","MODE","MODE.MULT","MODE.SNGL","MONTH","MROUND","MULTINOMIAL","N","NA","NEGBINOM.DIST","NEGBINOMDIST","NETWORKDAYS","NETWORKDAYS.INTL","NOMINAL","NORM.DIST","NORM.INV","NORM.S.DIST","NORM.S.INV","NORMDIST","NORMINV","NORMSDIST","NORMSINV","NOT","NOW","NPER","NPV","NUMBERVALUE","OCT2BIN","OCT2DEC","OCT2HEX","ODD","ODDFPRICE","ODDFYIELD","ODDLPRICE","ODDLYIELD","OFFSET","OR","PDURATION","PEARSON","PERCENTILE","PERCENTILE.EXC","PERCENTILE.INC","PERCENTRANK","PERCENTRANK.EXC","PERCENTRANK.INC","PERMUT","PERMUTATIONA","PHI","PI","PMT","POISSON","POISSON.DIST","POWER","PPMT","PRICE","PRICEDISC","PRICEMAT","PROB","PRODUCT","PROPER","PV","QUARTILE","QUARTILE.EXC","QUARTILE.INC","QUOTIENT","RADIANS","RAND","RANDARRAY","RANDBETWEEN","RANK","RANK.AVG","RANK.EQ","RATE","RECEIVED","REDUCE","REPLACE","REPLACEB","REPT","RIGHT","RIGHTB","ROMAN","ROUND","ROUNDDOWN","ROUNDUP","ROW","ROWS","RRI","RSQ","RTD","SCAN","SEARCH","SEARCHB","SEC","SECH","SECOND","SEQUENCE","SERIESSUM","SHEET","SHEETS","SIGN","SIN","SINH","SKEW","SKEW.P","SLN","SLOPE","SMALL","SORT","SORTBY","SQRT","SQRTPI","STANDARDIZE","STDEV","STDEV.P","STDEV.S","STDEVA","STDEVP","STDEVPA","STEYX","SUBSTITUTE","SUBTOTAL","SUM","SUMIF","SUMIFS","SUMPRODUCT","SUMSQ","SUMX2MY2","SUMX2PY2","SUMXMY2","SWITCH","SYD","T","T.DIST","T.DIST.2T","T.DIST.RT","T.INV","T.INV.2T","T.TEST","TAN","TANH","TBILLEQ","TBILLPRICE","TBILLYIELD","TEXT","TEXTAFTER","TEXTBEFORE","TEXTJOIN","TEXTSPLIT","TIME","TIMEVALUE","TOCOL","TODAY","TOROW","TRANSPOSE","TREND","TRIM","TRIMMEAN","TRUE","TRUNC","TYPE","UNICHAR","UNICODE","UNIQUE","UPPER","VALUE","VAR","VAR.P","VAR.S","VARA","VARP","VARPA","VDB","VLOOKUP","WEBSERVICE","WEEKDAY","WEEKNUM","WEIBULL","WEIBULL.DIST","WORKDAY","WORKDAY.INTL","XIRR","XLOOKUP","XMATCH","XNPV","XOR","YEAR","YEARFRAC","YIELD","YIELDDISC","YIELDMAT","Z.TEST"
] as const;

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: "#efefef",
    color: "#1f1f1f",
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
  },
  frame: {
    display: "flex",
    flexDirection: "column",
    minHeight: "100vh",
  },
  primaryTabsWrap: {
    borderTop: "1px solid #d0d0d0",
  },
  primaryTabs: {
    minHeight: "58px",
    display: "flex",
    alignItems: "flex-end",
    gap: "10px",
    padding: "0 10px",
    backgroundColor: "#317e43",
  },
  primaryTabBtn: {
    border: "none",
    background: "transparent",
    color: "#f8f8f8",
    fontSize: "11px",
    lineHeight: 1.1,
    padding: "0 10px 8px",
    cursor: "pointer",
    borderBottom: "3px solid transparent",
    minHeight: "100%",
  },
  primaryTabActive: {
    borderBottomColor: "#ffffff",
    fontWeight: 700,
    backgroundColor: "#2d6f3d",
  },
  menuCluster: {
    marginLeft: "auto",
    paddingBottom: "7px",
  },
  commandBar: {
    minHeight: "44px",
    backgroundColor: "#1d1d1d",
    borderBottom: "1px solid #161616",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "0 12px",
  },
  commandGroup: {
    display: "flex",
    alignItems: "stretch",
    gap: "8px",
  },
  iconBtn: {
    color: "#f5f5f5",
    backgroundColor: "transparent",
    border: "none",
    minWidth: "24px",
    height: "44px",
    borderRadius: "0",
  },
  iconBtnActive: {
    backgroundColor: "#4b4b4b",
  },
  secondaryTabs: {
    minHeight: "42px",
    display: "flex",
    alignItems: "flex-end",
    gap: "14px",
    padding: "0 12px",
    backgroundColor: "#f5f5f5",
    boxShadow: "0 1px 4px rgba(0, 0, 0, 0.12)",
  },
  secondaryTabBtn: {
    border: "none",
    background: "transparent",
    color: "#1f1f1f",
    fontSize: "11px",
    padding: "0 2px 7px",
    cursor: "pointer",
    borderBottom: "3px solid transparent",
  },
  secondaryTabActive: {
    borderBottomColor: "#317e43",
    fontWeight: 700,
  },
  content: {
    padding: "12px",
    display: "grid",
    gap: "8px",
  },
  topActions: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  miniBtn: {
    minHeight: "24px",
    fontSize: "10px",
  },
  tableWrap: {
    overflowX: "auto",
    border: "1px solid #d6d6d6",
    borderRadius: "6px",
    backgroundColor: "#ffffff",
  },
  table: {
    width: "100%",
    minWidth: "920px",
    borderCollapse: "collapse",
    fontSize: "10px",
  },
  th: {
    borderBottom: "1px solid #d6d6d6",
    borderRight: "1px solid #ececec",
    padding: "6px",
    textAlign: "left",
    backgroundColor: "#f4f4f4",
    whiteSpace: "nowrap",
    verticalAlign: "top",
  },
  thHeader: {
    display: "grid",
    gap: "3px",
    minWidth: "90px",
  },
  thLabelRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "6px",
  },
  thFilterRow: {
    display: "grid",
    gridTemplateColumns: "12px 1fr",
    gap: "4px",
    alignItems: "center",
  },
  thFilterIcon: {
    color: "#6e6e6e",
    fontSize: "10px",
  },
  td: {
    borderBottom: "1px solid #ececec",
    borderRight: "1px solid #f0f0f0",
    padding: "6px",
    whiteSpace: "nowrap",
  },
  sortBtn: {
    border: "none",
    padding: "0",
    margin: "0",
    background: "transparent",
    fontSize: "10px",
    fontWeight: 700,
    color: "#1f1f1f",
    cursor: "pointer",
  },
  filterInput: {
    width: "100%",
  },
  filterNative: {
    width: "100%",
    height: "20px",
    fontSize: "10px",
    border: "1px solid #b7b7b7",
    borderRadius: "4px",
    padding: "0 4px",
    boxSizing: "border-box",
  },
  rowCheckCol: {
    width: "30px",
  },
  rowActionCol: {
    width: "120px",
  },
  typeCol: {
    minWidth: "135px",
  },
  rowActions: {
    display: "flex",
    gap: "2px",
  },
  rowActionBtn: {
    minWidth: "18px",
    height: "18px",
    color: "#317e43",
    border: "none",
    backgroundColor: "transparent",
  },
  rowActionTextBtn: {
    minWidth: "0",
    height: "18px",
    border: "1px solid #c8c8c8",
    backgroundColor: "#ffffff",
    color: "#1f1f1f",
    fontSize: "9px",
    padding: "0 5px",
    borderRadius: "4px",
  },
  nameLink: {
    border: "none",
    background: "transparent",
    padding: "0",
    margin: "0",
    color: "#185abd",
    textDecoration: "underline",
    cursor: "pointer",
    fontSize: "10px",
    fontWeight: 600,
  },
  placeholder: {
    fontSize: "10px",
    fontStyle: "italic",
    color: "#425466",
    backgroundColor: "#f8f8f8",
    border: "1px solid #d8d8d8",
    borderRadius: "6px",
    padding: "12px",
  },
  status: {
    marginTop: "2px",
    padding: "5px 6px",
    borderRadius: tokens.borderRadiusMedium,
    border: "1px solid #d8d8d8",
    backgroundColor: "#ffffff",
    width: "fit-content",
    fontSize: "10px",
  },
  statusSuccess: {
    color: "#0f6f3f",
  },
  statusError: {
    color: "#a4262c",
  },
  modalBackdrop: {
    position: "fixed",
    inset: "0",
    backgroundColor: "rgba(0, 0, 0, 0.35)",
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "center",
    overflowY: "auto",
    padding: "10px 0",
    zIndex: 1000,
  },
  modal: {
    width: "min(540px, 94vw)",
    backgroundColor: "#ffffff",
    borderRadius: "8px",
    border: "1px solid #d6d6d6",
    padding: "12px",
    display: "grid",
    gap: "8px",
    maxHeight: "calc(100vh - 20px)",
    overflowY: "auto",
  },
  modalTitle: {
    fontSize: "12px",
    fontWeight: 700,
  },
  modalRow: {
    display: "grid",
    gap: "6px",
  },
  modalLabel: {
    fontSize: "10px",
    color: "#424242",
  },
  modalActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    marginTop: "4px",
  },
  modalInlineRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "8px",
  },
  checkboxRow: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: "10px",
  },
  smallSelect: {
    width: "100%",
    "> select": {
      fontSize: "10px",
      minHeight: "28px",
    },
  },
  previewBox: {
    border: "1px solid #d8d8d8",
    borderRadius: "6px",
    maxHeight: "150px",
    overflow: "auto",
    backgroundColor: "#fafafa",
  },
  previewTable: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: "10px",
  },
  previewTh: {
    textAlign: "left",
    padding: "6px",
    borderBottom: "1px solid #e7e7e7",
    backgroundColor: "#f1f1f1",
    position: "sticky",
    top: "0",
    zIndex: 1,
  },
  previewTd: {
    padding: "6px",
    borderBottom: "1px solid #ededed",
  },
  previewNoChange: {
    color: "#666666",
  },
  sectionCard: {
    border: "1px solid #d8d8d8",
    backgroundColor: "#ffffff",
    borderRadius: "6px",
    padding: "10px",
    display: "grid",
    gap: "8px",
  },
  sectionTitle: {
    fontSize: "11px",
    fontWeight: 700,
  },
  stylePresetGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: "8px",
  },
  stylePresetCard: {
    border: "1px solid #d8d8d8",
    borderRadius: "6px",
    padding: "8px",
    display: "grid",
    gap: "6px",
    backgroundColor: "#fafafa",
  },
  stylePresetTitle: {
    fontSize: "10px",
    fontWeight: 700,
  },
  stylePresetDesc: {
    fontSize: "9px",
    color: "#495463",
  },
  shapeToolbar: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  plusBtn: {
    minWidth: "30px",
    height: "28px",
    borderRadius: "50%",
    border: "1px solid #c8c8c8",
    backgroundColor: "#ffffff",
    fontSize: "16px",
    lineHeight: 1,
  },
  shapeHint: {
    fontSize: "10px",
    color: "#4e5a69",
  },
  swatchRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
  },
  swatchBtn: {
    width: "16px",
    height: "16px",
    border: "1px solid #c8c8c8",
    borderRadius: "3px",
    padding: "0",
    minWidth: "16px",
  },
  swatchBtnSelected: {
    border: "2px solid #1f4e78",
  },
  noFillBtn: {
    minWidth: "60px",
    height: "16px",
    border: "1px solid #c8c8c8",
    borderRadius: "3px",
    fontSize: "9px",
    lineHeight: "14px",
    backgroundColor: "#ffffff",
    padding: "0 4px",
  },
  noFillBtnSelected: {
    border: "2px solid #1f4e78",
    color: "#1f4e78",
    fontWeight: 700,
  },
  swatchInput: {
    width: "42px",
    height: "24px",
    border: "1px solid #c8c8c8",
    backgroundColor: "transparent",
    padding: 0,
  },
  numberInput: {
    width: "100%",
    height: "24px",
    fontSize: "10px",
    border: "1px solid #b7b7b7",
    borderRadius: "4px",
    padding: "0 6px",
    boxSizing: "border-box",
  },
  collapsibleSection: {
    border: "1px solid #e0e0e0",
    borderRadius: "6px",
    backgroundColor: "#fafafa",
    overflow: "hidden",
  },
  collapsibleSummary: {
    cursor: "pointer",
    listStyle: "none",
    padding: "7px 9px",
    fontSize: "10px",
    fontWeight: 700,
    backgroundColor: "#f3f3f3",
    borderBottom: "1px solid #e0e0e0",
  },
  collapsibleBody: {
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  nestedSection: {
    border: "1px solid #e2e2e2",
    borderRadius: "5px",
    backgroundColor: "#ffffff",
    overflow: "hidden",
  },
  nestedSummary: {
    cursor: "pointer",
    listStyle: "none",
    padding: "6px 8px",
    fontSize: "10px",
    fontWeight: 600,
    backgroundColor: "#f7f7f7",
    borderBottom: "1px solid #ececec",
  },
  nestedBody: {
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  themePaletteHeader: {
    display: "grid",
    gridTemplateColumns: "repeat(10, minmax(0, 1fr))",
    gap: "4px",
  },
  themePaletteRow: {
    display: "grid",
    gridTemplateColumns: "repeat(10, minmax(0, 1fr))",
    gap: "4px",
  },
  paletteSwatch: {
    width: "100%",
    minWidth: "16px",
    height: "16px",
    border: "1px solid #cfcfcf",
    borderRadius: "2px",
    padding: 0,
  },
  formulaMenuBar: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  formulaMenuSpacer: {
    flex: 1,
  },
  settingsPanel: {
    border: "1px solid #d8d8d8",
    borderRadius: "6px",
    padding: "8px",
    display: "grid",
    gap: "8px",
    backgroundColor: "#fafafa",
  },
  formulaEditorArea: {
    position: "absolute",
    inset: "0",
    width: "100%",
    height: "100%",
    fontSize: "11px",
    fontFamily: "Consolas, 'Courier New', monospace",
    border: "none",
    borderRadius: "4px",
    padding: "6px",
    boxSizing: "border-box",
    resize: "vertical",
    backgroundColor: "transparent",
    color: "transparent",
    caretColor: "#111111",
    WebkitTextFillColor: "transparent",
    lineHeight: "1.45",
    overflowX: "auto",
    overflowY: "auto",
    whiteSpace: "pre",
    "&::placeholder": {
      color: "#6e6e6e",
      WebkitTextFillColor: "#6e6e6e",
    },
  },
  formulaEditorWrap: {
    position: "relative",
    border: "1px solid #b7b7b7",
    borderRadius: "4px",
    backgroundColor: "#ffffff",
    minHeight: "82px",
    overflow: "hidden",
    resize: "vertical",
  },
  formulaLineNumbers: {
    position: "absolute",
    left: "0",
    top: "0",
    bottom: "0",
    width: "28px",
    borderRight: "1px solid #e2e2e2",
    backgroundColor: "#f7f7f7",
    overflow: "hidden",
    pointerEvents: "none",
  },
  formulaLineNumberText: {
    margin: "0",
    padding: "6px 4px 6px 0",
    fontSize: "11px",
    lineHeight: "1.45",
    fontFamily: "Consolas, 'Courier New', monospace",
    color: "#8b8b8b",
    textAlign: "right",
    whiteSpace: "pre",
  },
  formulaEditorHighlight: {
    margin: "0",
    padding: "6px",
    fontSize: "11px",
    fontFamily: "Consolas, 'Courier New', monospace",
    lineHeight: "1.45",
    whiteSpace: "pre",
    overflowX: "auto",
    overflowY: "auto",
    minHeight: "82px",
    boxSizing: "border-box",
    pointerEvents: "none",
    color: "#1d1d1d",
  },
  formulaTokenFunction: {
    color: "#143a8a",
  },
  formulaTokenName: {
    color: "#295dbf",
  },
  formulaTokenTable: {
    color: "#2f6b3f",
  },
  formulaTokenLocal: {
    color: "#3d7ae6",
    fontStyle: "italic",
  },
  formulaTokenString: {
    color: "#8b2f1f",
  },
  formulaTokenNumber: {
    color: "#6a3cc7",
  },
  formulaTokenOperator: {
    color: "#2a2a2a",
  },
  formulaTokenDefault: {
    color: "#111111",
  },
  suggestionsBox: {
    border: "1px solid #d8d8d8",
    borderRadius: "4px",
    backgroundColor: "#fff",
    maxHeight: "100px",
    overflowY: "auto",
  },
  suggestionBtn: {
    width: "100%",
    border: "none",
    backgroundColor: "#fff",
    textAlign: "left",
    fontSize: "10px",
    padding: "6px",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  suggestionBtnActive: {
    backgroundColor: "#2f63b7",
    color: "#ffffff",
  },
  suggestionKind: {
    color: "#5c7cb8",
    fontWeight: 600,
    minWidth: "18px",
  },
  suggestionMatch: {
    color: "#2f63b7",
    fontWeight: 700,
  },
  suggestionAccentOnActive: {
    color: "#dbe7ff",
  },
  formulaHint: {
    fontSize: "10px",
    color: "#2f63b7",
  },
  formulaPromptCard: {
    border: "1px solid #bfd2f5",
    borderRadius: "4px",
    backgroundColor: "#f2f7ff",
    padding: "8px",
    display: "grid",
    gap: "6px",
    marginBottom: "8px",
  },
  formulaPromptTitle: {
    fontSize: "11px",
    fontWeight: 700,
    color: "#194a96",
  },
  formulaPromptText: {
    fontSize: "10px",
    color: "#2f3b4a",
  },
  formulaPromptMeta: {
    fontSize: "10px",
    color: "#4f5d6d",
  },
  formulaPaletteEditor: {
    display: "grid",
    gap: "6px",
  },
  formulaPaletteGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "6px 10px",
  },
  formulaPaletteRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    fontSize: "10px",
    color: "#2f3b4a",
  },
  colorInput: {
    width: "34px",
    height: "22px",
    border: "1px solid #c8c8c8",
    backgroundColor: "transparent",
    padding: "0",
  },
  formulaContextMenu: {
    position: "fixed",
    zIndex: 1200,
    minWidth: "220px",
    border: "1px solid #cfcfcf",
    borderRadius: "4px",
    backgroundColor: "#f7f7f7",
    boxShadow: "0 6px 16px rgba(0, 0, 0, 0.2)",
    padding: "4px 0",
    display: "grid",
  },
  formulaContextItem: {
    border: "none",
    background: "transparent",
    width: "100%",
    textAlign: "left",
    padding: "6px 10px",
    fontSize: "10px",
    color: "#2f3b4a",
    display: "flex",
    justifyContent: "space-between",
    cursor: "pointer",
  },
  formulaContextItemActive: {
    backgroundColor: "#2f63b7",
    color: "#ffffff",
  },
  formulaContextSep: {
    border: "none",
    borderTop: "1px solid #dcdcdc",
    margin: "3px 0",
  },
  gridOutputWrap: {
    border: "1px solid #d8d8d8",
    borderRadius: "4px",
    overflowX: "auto",
    backgroundColor: "#fff",
  },
  gridOutputTable: {
    borderCollapse: "collapse",
    width: "100%",
    fontSize: "10px",
  },
  gridOutputCell: {
    border: "1px solid #ececec",
    padding: "4px 6px",
    whiteSpace: "nowrap",
  },
  bridgeBanner: {
    backgroundColor: "#fff7ed",
    borderBottom: "1px solid #fed7aa",
    padding: "8px 12px",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  bridgeBannerText: {
    fontSize: "12px",
    color: "#9a3412",
  },
});

const mapTargetToLegacy = (
  target: NavigationTarget | undefined
): { primaryTab: PrimaryTab; secondaryTabId: string } => {
  switch (target) {
    case "formulas":
      return { primaryTab: "Names", secondaryTabId: "functions" };
    case "tables":
      return { primaryTab: "Names", secondaryTabId: "tables" };
    case "format":
      return { primaryTab: "Format", secondaryTabId: "styles" };
    case "sandbox-debug":
      return { primaryTab: "Sandbox", secondaryTabId: "debug" };
    case "names":
    case "names-create":
    default:
      return { primaryTab: "Names", secondaryTabId: "ranges" };
  }
};

const LegacyApp: React.FC<LegacyAppProps> = ({
  initialTarget = "names",
  bridgeMode = false,
  onExitBridge,
}) => {
  const styles = useStyles();
  const mappedTarget = mapTargetToLegacy(initialTarget);
  const isFormulaPopout =
    typeof window !== "undefined" &&
    new URLSearchParams(window.location.search).get("popout") === "formula";
  const [primaryTab, setPrimaryTab] = useState<PrimaryTab>(mappedTarget.primaryTab);
  const [secondaryTabId, setSecondaryTabId] = useState<string>(mappedTarget.secondaryTabId);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<StatusType>("success");
  const [namedRanges, setNamedRanges] = useState<NamedRangeRecord[]>([]);
  const [tables, setTables] = useState<TableRecord[]>([]);
  const [isLoadingRanges, setIsLoadingRanges] = useState<boolean>(false);
  const [isLoadingTables, setIsLoadingTables] = useState<boolean>(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [selectedTableIds, setSelectedTableIds] = useState<Set<string>>(new Set());
  const [inlineTableEdit, setInlineTableEdit] = useState<{ id: string; value: string } | null>(null);
  const [tableBulkState, setTableBulkState] = useState<TableBulkState>({
    open: false,
    operation: "Add",
    position: "Prefix",
    delimiter: "underscore",
    textValue: "",
    replaceWith: "",
    caseTransform: "none",
  });
  const [sortColumn, setSortColumn] = useState<SortColumn>("name");
  const [tableSortColumn, setTableSortColumn] = useState<TableSortColumn>("name");
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");
  const [tableSortDirection, setTableSortDirection] = useState<SortDirection>("asc");
  const [columnFilters, setColumnFilters] = useState<Record<SortColumn, string>>({
    name: "",
    address: "",
    sheet: "",
    scope: "",
    type: "",
  });
  const [tableFilters, setTableFilters] = useState<Record<TableSortColumn, string>>({
    name: "",
    address: "",
    sheet: "",
    scope: "",
  });
  const [editState, setEditState] = useState<EditState>({
    open: false,
    record: null,
    name: "",
    address: "",
    fallbackSheet: "",
    caseTransform: "none",
  });
  const [bulkState, setBulkState] = useState<BulkState>({
    open: false,
    operation: "Add",
    position: "Prefix",
    delimiter: "underscore",
    textValue: "",
    replaceWith: "",
    enableSecondEdit: false,
    secondOperation: "Remove",
    secondPosition: "Suffix",
    secondDelimiter: "underscore",
    secondTextValue: "",
    secondReplaceWith: "",
    caseTransform: "none",
    secondCaseTransform: "none",
    deleteName: true,
    deleteValues: false,
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
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [formulaText, setFormulaText] = useState<string>("");
  const [formulaMode, setFormulaMode] = useState<"Formula" | "Function">("Formula");
  const [formulaSettingsOpen, setFormulaSettingsOpen] = useState<boolean>(false);
  const [formulaColorSchemeId, setFormulaColorSchemeId] = useState<string>("preset:Advanced");
  const [customFormulaSchemes, setCustomFormulaSchemes] = useState<CustomFormulaScheme[]>([]);
  const [formulaPaletteDraft, setFormulaPaletteDraft] = useState<FormulaColorPalette>(
    cloneFormulaPalette(FORMULA_COLOR_SCHEMES.Advanced)
  );
  const [customSchemeName, setCustomSchemeName] = useState<string>("My Scheme");
  const [showFormulaLineNumbers, setShowFormulaLineNumbers] = useState<boolean>(true);
  const [autoCaptureExcelFormula, setAutoCaptureExcelFormula] = useState<boolean>(true);
  const [autoOpenFormulaTab, setAutoOpenFormulaTab] = useState<boolean>(true);
  const [customFunctionName, setCustomFunctionName] = useState<string>("MyFunction");
  const [customFunctionBody, setCustomFunctionBody] = useState<string>("param1+param2");
  const [customFunctionParams, setCustomFunctionParams] = useState<string[]>(["param1", "param2"]);
  const [testFormulaCall, setTestFormulaCall] = useState<string>("");
  const [formulaOutput, setFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
  const [activeFormulaPrompt, setActiveFormulaPrompt] = useState<ActiveFormulaPrompt | null>(null);
  const [formulaCursor, setFormulaCursor] = useState<number>(formulaText.length);
  const [activeSuggestionIndex, setActiveSuggestionIndex] = useState<number>(0);
  const [formulaContextMenu, setFormulaContextMenu] = useState<{ x: number; y: number } | null>(null);
  const [shapeAddOpen, setShapeAddOpen] = useState<boolean>(false);
  const [activeShape, setActiveShape] = useState<InsertedShapeRecord | null>(null);
  const [shapeState, setShapeState] = useState<ShapeEditorState>({
    shapeName: "",
    shapeType: "Rectangle",
    bindingMode: "StaticText",
    geometricShapeType: "Rectangle",
    left: 0,
    top: 0,
    zOrderPosition: 0,
    fillColor: "#E3F2FD",
    outlineColor: "#1F4E78",
    fontColor: "#1F1F1F",
    text: "",
    valueBinding: "",
    formula: "",
    lineWeight: 1,
    fillTransparency: 0,
    width: 110,
    height: 34,
    rotation: 0,
    textHorizontalAlignment: "Center",
    textVerticalAlignment: "Middle",
    fontSize: 11,
    bold: false,
    italic: false,
    lockAspectRatio: false,
    selectedThemeId: "theme1",
  });
  const [shapeSuggestionIndex, setShapeSuggestionIndex] = useState<number>(0);
  const [shapeThemes, setShapeThemes] = useState<ShapeTheme[]>(SHAPE_THEMES);
  const [newShapeThemeName, setNewShapeThemeName] = useState<string>("MyTheme");
  const [customThemeColorsText, setCustomThemeColorsText] = useState<string>(
    SHAPE_THEMES[0].themeColors.join(",")
  );
  const [customStandardColorsText, setCustomStandardColorsText] = useState<string>(
    SHAPE_THEMES[0].standardColors.join(",")
  );
  const formulaEditorRef = useRef<HTMLTextAreaElement | null>(null);
  const formulaHighlightRef = useRef<HTMLPreElement | null>(null);
  const formulaLineNumbersRef = useRef<HTMLDivElement | null>(null);
  const shapeActivationHandlersRef = useRef<Array<{ remove: () => Promise<void> | void }>>([]);

  const primaryTabs: PrimaryTab[] = useMemo(() => ["Names", "Format", "Sandbox"], []);

  const secondaryTabs = useMemo(() => {
    if (primaryTab === "Names") {
      return NAMES_TABS;
    }
    if (primaryTab === "Format") {
      return FORMAT_TABS;
    }
    return SANDBOX_TABS;
  }, [primaryTab]);

  useEffect(() => {
    if (!isFormulaPopout) {
      return;
    }
    setPrimaryTab("Names");
    setSecondaryTabId("functions");
  }, [isFormulaPopout]);

  useEffect(() => {
    if (isFormulaPopout) {
      return;
    }
    const mapped = mapTargetToLegacy(initialTarget);
    setPrimaryTab(mapped.primaryTab);
    setSecondaryTabId(mapped.secondaryTabId);
  }, [initialTarget, isFormulaPopout]);

  useEffect(() => {
    const currentExists = secondaryTabs.some((tab) => tab.id === secondaryTabId);
    if (!currentExists) {
      setSecondaryTabId(secondaryTabs[0].id);
    }
  }, [secondaryTabId, secondaryTabs]);

  const runAction = async (label: string, action: () => Promise<void>) => {
    setStatus("");
    try {
      await action();
      setStatusType("success");
      setStatus(`${label} completed.`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`${label} failed: ${message}`);
    }
  };

  const loadRanges = async (resetFilters = false) => {
    setIsLoadingRanges(true);
    try {
      const ranges = await getNamedRanges();
      setNamedRanges(ranges);
      setSelectedIds(new Set());
      if (resetFilters) {
        setColumnFilters({
          name: "",
          address: "",
          sheet: "",
          scope: "",
          type: "",
        });
      }
      setStatusType("success");
      setStatus(`Loaded ${ranges.length} name/shape record(s).`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Load named ranges failed: ${message}`);
    } finally {
      setIsLoadingRanges(false);
    }
  };

  const loadTables = async (resetFilters = false) => {
    setIsLoadingTables(true);
    try {
      const loaded = await getTables();
      setTables(loaded);
      setSelectedTableIds(new Set());
      setInlineTableEdit(null);
      if (resetFilters) {
        setTableFilters({
          name: "",
          address: "",
          sheet: "",
          scope: "",
        });
      }
      setStatusType("success");
      setStatus(`Loaded ${loaded.length} table(s).`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Load tables failed: ${message}`);
    } finally {
      setIsLoadingTables(false);
    }
  };

  useEffect(() => {
    if (primaryTab === "Names" && secondaryTabId === "ranges") {
      void loadRanges();
    }
    if (primaryTab === "Names" && secondaryTabId === "tables") {
      void loadTables();
    }
    if (primaryTab === "Names" && secondaryTabId === "functions") {
      if (namedRanges.length === 0) {
        void loadRanges();
      }
      if (tables.length === 0) {
        void loadTables();
      }
    }
    if (primaryTab === "Format" && secondaryTabId === "shapes" && namedRanges.length === 0) {
      void loadRanges();
    }
  }, [primaryTab, secondaryTabId, namedRanges.length, tables.length]);

  const filteredAndSortedRanges = useMemo(() => {
    const filtered = namedRanges.filter((item) =>
      RANGE_COLUMNS.every((col) =>
        item[col].toLowerCase().includes(columnFilters[col].trim().toLowerCase())
      )
    );

    return filtered.sort((a, b) => {
      const compare = a[sortColumn].localeCompare(b[sortColumn], undefined, {
        sensitivity: "base",
      });
      return sortDirection === "asc" ? compare : -compare;
    });
  }, [namedRanges, columnFilters, sortColumn, sortDirection]);

  const filteredAndSortedTables = useMemo(() => {
    const filtered = tables.filter((item) =>
      TABLE_COLUMNS.every((col) =>
        item[col].toLowerCase().includes(tableFilters[col].trim().toLowerCase())
      )
    );

    return filtered.sort((a, b) => {
      const compare = a[tableSortColumn].localeCompare(b[tableSortColumn], undefined, {
        sensitivity: "base",
      });
      return tableSortDirection === "asc" ? compare : -compare;
    });
  }, [tables, tableFilters, tableSortColumn, tableSortDirection]);

  const selectableVisibleRanges = useMemo(
    () => filteredAndSortedRanges.filter((item) => item.kind === "NamedRange"),
    [filteredAndSortedRanges]
  );
  const allVisibleSelected =
    selectableVisibleRanges.length > 0 &&
    selectableVisibleRanges.every((item) => selectedIds.has(item.id));
  const hasActiveFilters = RANGE_COLUMNS.some((col) => columnFilters[col].trim().length > 0);

  const clearFilters = () => {
    setColumnFilters({
      name: "",
      address: "",
      sheet: "",
      scope: "",
      type: "",
    });
  };

  const clearTableFilters = () => {
    setTableFilters({
      name: "",
      address: "",
      sheet: "",
      scope: "",
    });
  };

  const resetBulkState = () => {
    setBulkState({
      open: false,
      operation: "Add",
      position: "Prefix",
      delimiter: "underscore",
      textValue: "",
      replaceWith: "",
      enableSecondEdit: false,
      secondOperation: "Remove",
      secondPosition: "Suffix",
      secondDelimiter: "underscore",
      secondTextValue: "",
      secondReplaceWith: "",
      caseTransform: "none",
      secondCaseTransform: "none",
      deleteName: true,
      deleteValues: false,
    });
  };

  const getDelimiterText = (delimiter: BulkState["delimiter"]) => {
    switch (delimiter) {
      case "underscore":
        return "_";
      case "dot":
        return ".";
      default:
        return "";
    }
  };

  const splitNameTokens = (value: string): string[] => {
    const spaced = value
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/[_\-.]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!spaced) {
      return [];
    }
    return spaced.split(" ");
  };

  const normalizeAddressForCompare = (address: string) =>
    address.replace(/^=/, "").replace(/\$/g, "").replace(/'/g, "").trim().toUpperCase();

  type LockState = "none" | "abs" | "rowAbs" | "colAbs";

  const getLockState = (cellRef: string): LockState | null => {
    const match = cellRef.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
    if (!match) {
      return null;
    }
    const hasColLock = match[1] === "$";
    const hasRowLock = match[3] === "$";
    if (hasColLock && hasRowLock) {
      return "abs";
    }
    if (!hasColLock && hasRowLock) {
      return "rowAbs";
    }
    if (hasColLock && !hasRowLock) {
      return "colAbs";
    }
    return "none";
  };

  const applyLockState = (cellRef: string, state: LockState) => {
    const match = cellRef.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/);
    if (!match) {
      return cellRef;
    }
    const col = match[2];
    const row = match[4];
    if (state === "abs") {
      return `$${col}$${row}`;
    }
    if (state === "rowAbs") {
      return `${col}$${row}`;
    }
    if (state === "colAbs") {
      return `$${col}${row}`;
    }
    return `${col}${row}`;
  };

  const nextLockState = (state: LockState): LockState => {
    if (state === "none") {
      return "abs";
    }
    if (state === "abs") {
      return "rowAbs";
    }
    if (state === "rowAbs") {
      return "colAbs";
    }
    return "none";
  };

  const cycleSingleCellLock = (cellRef: string) => {
    const state = getLockState(cellRef);
    if (!state) {
      return cellRef;
    }
    return applyLockState(cellRef, nextLockState(state));
  };

  const cycleRangeLock = (rangeRef: string) => {
    const rangeMatch = rangeRef.match(
      /^(\$?[A-Za-z]{1,3}\$?\d+)\s*:\s*(\$?[A-Za-z]{1,3}\$?\d+)$/
    );
    if (!rangeMatch) {
      return rangeRef;
    }
    const left = rangeMatch[1];
    const right = rangeMatch[2];
    const leftState = getLockState(left);
    const rightState = getLockState(right);
    const baseState = leftState ?? rightState;
    if (!baseState) {
      return rangeRef;
    }
    const nextStateValue = nextLockState(baseState);
    return `${applyLockState(left, nextStateValue)}:${applyLockState(right, nextStateValue)}`;
  };

  const cycleCellOrRangeLockAtCursor = (
    text: string,
    cursor: number,
    selectionStart: number,
    selectionEnd: number
  ) => {
    if (selectionEnd > selectionStart) {
      const selectedText = text.slice(selectionStart, selectionEnd);
      const isCell = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(selectedText);
      const isRange = /^\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+$/.test(selectedText);
      if (isCell || isRange) {
        const replacement = isRange ? cycleRangeLock(selectedText) : cycleSingleCellLock(selectedText);
        const updated = `${text.slice(0, selectionStart)}${replacement}${text.slice(selectionEnd)}`;
        const end = selectionStart + replacement.length;
        return { updated, nextStart: selectionStart, nextEnd: end };
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

    const regex = /\$?[A-Za-z]{1,3}\$?\d+/g;
    let chosen: { start: number; end: number; value: string } | null = null;
    let match: RegExpExecArray | null;

    while ((match = regex.exec(text)) !== null) {
      const start = match.index;
      const end = start + match[0].length;
      if (cursor >= start && cursor <= end) {
        chosen = { start, end, value: match[0] };
        break;
      }
      if (end < cursor) {
        chosen = { start, end, value: match[0] };
      }
    }

    if (!chosen) {
      return { updated: text, nextStart: cursor, nextEnd: cursor };
    }

    const replacement = cycleSingleCellLock(chosen.value);
    const updated = `${text.slice(0, chosen.start)}${replacement}${text.slice(chosen.end)}`;
    const nextPos = chosen.start + replacement.length;
    return { updated, nextStart: nextPos, nextEnd: nextPos };
  };

  const applyCaseTransform = (name: string, transform: CaseTransform) => {
    if (transform === "none") {
      return name;
    }

    const words = splitNameTokens(name).map((word) => word.toLowerCase());
    if (words.length === 0) {
      return name;
    }

    if (transform === "camelCase") {
      return words
        .map((word, index) =>
          index === 0 ? word : `${word.charAt(0).toUpperCase()}${word.slice(1)}`
        )
        .join("");
    }

    if (transform === "snake_case") {
      return words.join("_");
    }

    return words.join("_").toUpperCase();
  };

  const prettyFormatFormula = (input: string) => {
    const source = input.trim();
    if (!source) {
      return "";
    }

    let formula = source;
    if (!formula.startsWith("=")) {
      formula = `=${formula}`;
    }

    const splitTopLevelArgs = (value: string): string[] => {
      const args: string[] = [];
      let current = "";
      let depth = 0;
      let inString = false;
      for (let i = 0; i < value.length; i += 1) {
        const ch = value[i];
        if (ch === '"') {
          inString = !inString;
          current += ch;
          continue;
        }
        if (!inString) {
          if (ch === "(") {
            depth += 1;
            current += ch;
            continue;
          }
          if (ch === ")") {
            depth = Math.max(0, depth - 1);
            current += ch;
            continue;
          }
          if (ch === "," && depth === 0) {
            args.push(current.trim());
            current = "";
            continue;
          }
        }
        current += ch;
      }
      if (current.trim().length > 0) {
        args.push(current.trim());
      }
      return args;
    };

    const tryPairwiseLetFormatting = () => {
      const letMatch = formula.match(/^=\s*LET\s*\(([\s\S]*)\)\s*$/i);
      if (!letMatch) {
        return null as string | null;
      }
      const args = splitTopLevelArgs(letMatch[1]);
      if (args.length < 3) {
        return null;
      }
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
    };

    const pairwiseLet = tryPairwiseLetFormatting();
    if (pairwiseLet) {
      return pairwiseLet;
    }

    const indentUnit = "  ";
    const lines: string[] = [];
    let current = "";
    let depth = 0;
    let inString = false;
    let previousNonWs = "";

    const pushCurrent = () => {
      const trimmed = current.trimEnd();
      if (trimmed.length > 0) {
        lines.push(trimmed);
      }
      current = "";
    };

    const startLine = (indentDepth: number) => {
      current = `${indentUnit.repeat(Math.max(0, indentDepth))}`;
    };

    startLine(0);

    for (let i = 0; i < formula.length; i += 1) {
      const ch = formula[i];

      if (ch === '"') {
        current += ch;
        inString = !inString;
        previousNonWs = ch;
        continue;
      }

      if (inString) {
        current += ch;
        continue;
      }

      if (ch === "\r" || ch === "\n") {
        continue;
      }

      if (ch === "(") {
        current += ch;
        pushCurrent();
        depth += 1;
        startLine(depth);
        previousNonWs = ch;
        continue;
      }

      if (ch === ",") {
        current += ch;
        pushCurrent();
        startLine(depth);
        previousNonWs = ch;
        continue;
      }

      if (ch === ")") {
        const trimmedCurrent = current.trim();
        if (trimmedCurrent.length > 0) {
          pushCurrent();
        }
        depth = Math.max(0, depth - 1);
        startLine(depth);
        current += ")";
        previousNonWs = ch;
        continue;
      }

      if (ch === " " || ch === "\t") {
        if (previousNonWs && previousNonWs !== "(" && previousNonWs !== "," && previousNonWs !== " ") {
          current += " ";
          previousNonWs = " ";
        }
        continue;
      }

      current += ch;
      previousNonWs = ch;
    }

    if (current.trim().length > 0) {
      pushCurrent();
    }

    return lines.join("\n");
  };

  const getCurrentFunctionContext = useMemo(() => {
    const text = formulaText.slice(0, formulaCursor);
    const stack: Array<{ name: string; argIndex: number }> = [];
    let token = "";
    let inString = false;

    for (let i = 0; i < text.length; i += 1) {
      const ch = text[i];
      if (ch === '"') {
        inString = !inString;
        token = "";
        continue;
      }
      if (inString) {
        continue;
      }
      if (/[A-Za-z0-9_.]/.test(ch)) {
        token += ch;
        continue;
      }
      if (ch === "(") {
        if (token) {
          stack.push({ name: token.toUpperCase(), argIndex: 0 });
        }
        token = "";
        continue;
      }
      if (ch === ",") {
        if (stack.length > 0) {
          stack[stack.length - 1].argIndex += 1;
        }
        token = "";
        continue;
      }
      if (ch === ")") {
        stack.pop();
        token = "";
        continue;
      }
      token = "";
    }

    return stack.length > 0 ? stack[stack.length - 1] : null;
  }, [formulaText, formulaCursor]);

  const formulaTokenBeforeCursorRaw = useMemo(() => {
    const beforeCursor = formulaText.slice(0, formulaCursor);
    const tokenMatch = beforeCursor.match(/([A-Za-z_][A-Za-z0-9_.]*)$/);
    return tokenMatch ? tokenMatch[1] : "";
  }, [formulaText, formulaCursor]);

  const formulaTokenBeforeCursor = useMemo(() => {
    return formulaTokenBeforeCursorRaw.toUpperCase();
  }, [formulaTokenBeforeCursorRaw]);

  const shouldShowIntellisense = useMemo(() => {
    const beforeCursor = formulaText.slice(0, formulaCursor);
    if (!beforeCursor.startsWith("=")) {
      return false;
    }
    return formulaTokenBeforeCursorRaw.trim().length > 0;
  }, [formulaText, formulaCursor, formulaTokenBeforeCursorRaw]);

  const localVariableSuggestions = useMemo<FormulaSuggestion[]>(() => {
    interface ParseFrame {
      name: string;
      args: string[];
      current: string;
    }

    const toIdentifier = (value: string) => value.trim();
    const isIdentifier = (value: string) => /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(value);
    const beforeCursor = formulaText.slice(0, formulaCursor);
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

    const collected: string[] = [];
    const addLocal = (candidate: string) => {
      const cleaned = toIdentifier(candidate);
      if (!isIdentifier(cleaned)) {
        return;
      }
      if (cleaned.includes(".")) {
        return;
      }
      if (!collected.some((existing) => existing.toUpperCase() === cleaned.toUpperCase())) {
        collected.push(cleaned);
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

    return collected.map((name) => ({
      name,
      kind: "Local",
      signature: "LET/LAMBDA variable",
    }));
  }, [formulaText, formulaCursor]);

  const colorizedFormulaTokens = useMemo(() => {
    const namedSet = new Set(
      namedRanges.filter((item) => item.kind === "NamedRange").map((item) => item.name.toUpperCase())
    );
    const tableSet = new Set(tables.map((item) => item.name.toUpperCase()));
    const localSet = new Set(localVariableSuggestions.map((item) => item.name.toUpperCase()));
    const functionSet = new Set<string>(EXCEL_FUNCTION_NAMES as readonly string[]);
    const tokens: Array<{ text: string; kind: FormulaTokenKind }> = [];
    const text = formulaText;
    let i = 0;

    const isIdentStart = (value: string) => /[A-Za-z_\\]/.test(value);
    const isIdentPart = (value: string) => /[A-Za-z0-9_.\\]/.test(value);
    const isDigit = (value: string) => /[0-9]/.test(value);

    while (i < text.length) {
      const ch = text[i];

      if (ch === '"') {
        let j = i + 1;
        while (j < text.length) {
          if (text[j] === '"' && text[j + 1] === '"') {
            j += 2;
            continue;
          }
          if (text[j] === '"') {
            j += 1;
            break;
          }
          j += 1;
        }
        tokens.push({ text: text.slice(i, j), kind: "string" });
        i = j;
        continue;
      }

      if (isIdentStart(ch)) {
        let j = i + 1;
        while (j < text.length && isIdentPart(text[j])) {
          j += 1;
        }
        const value = text.slice(i, j);
        const upper = value.toUpperCase();

        let k = j;
        while (k < text.length && /\s/.test(text[k])) {
          k += 1;
        }
        const hasCallParen = text[k] === "(";

        let kind: FormulaTokenKind = "default";
        if (localSet.has(upper)) {
          kind = "local";
        } else if (namedSet.has(upper)) {
          kind = "name";
        } else if (tableSet.has(upper)) {
          kind = "table";
        } else if (functionSet.has(upper) && hasCallParen) {
          kind = "function";
        }
        tokens.push({ text: value, kind });
        i = j;
        continue;
      }

      if (isDigit(ch)) {
        let j = i + 1;
        while (j < text.length && /[0-9.]/.test(text[j])) {
          j += 1;
        }
        tokens.push({ text: text.slice(i, j), kind: "number" });
        i = j;
        continue;
      }

      if ("=+-*/^&<>:,()".includes(ch)) {
        tokens.push({ text: ch, kind: "operator" });
        i += 1;
        continue;
      }

      tokens.push({ text: ch, kind: "default" });
      i += 1;
    }

    return tokens;
  }, [formulaText, namedRanges, tables, localVariableSuggestions]);

  const selectedCustomFormulaScheme = useMemo(() => {
    return customFormulaSchemes.find((item) => item.id === formulaColorSchemeId) ?? null;
  }, [customFormulaSchemes, formulaColorSchemeId]);

  const isSelectedPresetScheme = formulaColorSchemeId.startsWith("preset:");
  const activeFormulaSchemeLabel = isSelectedPresetScheme
    ? formulaColorSchemeId.replace("preset:", "")
    : selectedCustomFormulaScheme?.name ?? "Custom";

  useEffect(() => {
    if (formulaColorSchemeId.startsWith("preset:")) {
      const presetName = formulaColorSchemeId.replace("preset:", "") as FormulaColorScheme;
      const preset = FORMULA_COLOR_SCHEMES[presetName] ?? FORMULA_COLOR_SCHEMES.Advanced;
      setFormulaPaletteDraft(cloneFormulaPalette(preset));
      setCustomSchemeName(presetName);
      return;
    }
    if (selectedCustomFormulaScheme) {
      setFormulaPaletteDraft(cloneFormulaPalette(selectedCustomFormulaScheme.palette));
      setCustomSchemeName(selectedCustomFormulaScheme.name);
    }
  }, [formulaColorSchemeId, selectedCustomFormulaScheme]);

  const selectedFormulaPalette = formulaPaletteDraft;

  const allSuggestions = useMemo(() => {
    const functionDefsByName = new Map(FUNCTION_DEFINITIONS.map((fn) => [fn.name, fn]));
    const functions: FormulaSuggestion[] = EXCEL_FUNCTION_NAMES.map((name) => {
      const detail = functionDefsByName.get(name);
      return {
        name,
        kind: "Function",
        signature: detail ? `${name}(${detail.params.join(", ")})` : undefined,
      };
    });

    const names: FormulaSuggestion[] = namedRanges
      .filter((item) => item.kind === "NamedRange")
      .map((item) => ({
        name: item.name,
        kind: "Name",
      }));
    const tableItems: FormulaSuggestion[] = tables.map((item) => ({
      name: item.name,
      kind: "Table",
    }));

    const byKey = new Map<string, FormulaSuggestion>();
    [...localVariableSuggestions, ...functions, ...names, ...tableItems].forEach((item) => {
      const key = item.name.toUpperCase();
      if (!byKey.has(key)) {
        byKey.set(key, item);
      }
    });
    return Array.from(byKey.values());
  }, [localVariableSuggestions, namedRanges, tables]);

  const createCustomFormulaScheme = () => {
    const name = customSchemeName.trim() || `Custom ${customFormulaSchemes.length + 1}`;
    const id = `custom:${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    const next: CustomFormulaScheme = {
      id,
      name,
      palette: cloneFormulaPalette(formulaPaletteDraft),
    };
    setCustomFormulaSchemes((prev) => [...prev, next]);
    setFormulaColorSchemeId(id);
    setStatusType("success");
    setStatus(`Created color scheme "${name}".`);
  };

  const updateCustomFormulaScheme = () => {
    if (!selectedCustomFormulaScheme) {
      return;
    }
    const nextName = customSchemeName.trim() || selectedCustomFormulaScheme.name;
    setCustomFormulaSchemes((prev) =>
      prev.map((item) =>
        item.id === selectedCustomFormulaScheme.id
          ? { ...item, name: nextName, palette: cloneFormulaPalette(formulaPaletteDraft) }
          : item
      )
    );
    setStatusType("success");
    setStatus(`Updated color scheme "${nextName}".`);
  };

  const deleteCustomFormulaScheme = () => {
    if (!selectedCustomFormulaScheme) {
      return;
    }
    const deletedName = selectedCustomFormulaScheme.name;
    setCustomFormulaSchemes((prev) => prev.filter((item) => item.id !== selectedCustomFormulaScheme.id));
    setFormulaColorSchemeId("preset:Advanced");
    setStatusType("success");
    setStatus(`Deleted color scheme "${deletedName}".`);
  };

  const functionSuggestions = useMemo(() => {
    const token = formulaTokenBeforeCursor.trim();
    if (!token) {
      return [];
    }
    const isTokenSearch = true;
    const prioritized = allSuggestions
      .filter((item) => {
        if (!isTokenSearch) {
          return true;
        }
        const upper = item.name.toUpperCase();
        return upper.startsWith(token) || upper.includes(token);
      })
      .sort((a, b) => {
        const aStarts = isTokenSearch ? a.name.toUpperCase().startsWith(token) : false;
        const bStarts = isTokenSearch ? b.name.toUpperCase().startsWith(token) : false;
        if (aStarts !== bStarts) {
          return aStarts ? -1 : 1;
        }
        const order = isTokenSearch
          ? ({ Local: 0, Name: 1, Table: 2, Function: 3 } as const)
          : ({ Local: 0, Function: 1, Name: 2, Table: 3 } as const);
        if (a.kind !== b.kind) {
          return order[a.kind] - order[b.kind];
        }
        return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
      });
    return prioritized.slice(0, isTokenSearch ? 120 : 40);
  }, [allSuggestions, formulaTokenBeforeCursor]);

  useEffect(() => {
    setActiveSuggestionIndex(0);
  }, [formulaTokenBeforeCursor, shouldShowIntellisense, functionSuggestions.length]);

  const formulaReferenceHint = useMemo<FormulaReferenceHint | null>(() => {
    const token = formulaTokenBeforeCursor;
    if (!token) {
      return null;
    }

    const localMatch = localVariableSuggestions.find((item) => item.name.toUpperCase() === token);
    if (localMatch) {
      return {
        kind: "Local",
        name: localMatch.name,
        detail: "LET/LAMBDA variable",
      };
    }

    const namedMatch = namedRanges.find(
      (item) => item.kind === "NamedRange" && item.name.toUpperCase() === token
    );
    if (namedMatch) {
      return {
        kind: "Name",
        name: namedMatch.name,
        detail: namedMatch.address,
      };
    }

    const tableMatch = tables.find((item) => item.name.toUpperCase() === token);
    if (tableMatch) {
      return {
        kind: "Table",
        name: tableMatch.name,
        detail: tableMatch.address,
      };
    }

    const functionMatch = EXCEL_FUNCTION_NAMES.find((item) => item === token);
    if (functionMatch) {
      const definition = FUNCTION_DEFINITIONS.find((fn) => fn.name === functionMatch);
      return {
        kind: "Function",
        name: functionMatch,
        detail: definition ? `${functionMatch}(${definition.params.join(", ")})` : `${functionMatch}(...)`,
      };
    }

    return null;
  }, [formulaTokenBeforeCursor, localVariableSuggestions, namedRanges, tables]);

  const contextFunctionDefinition = useMemo(() => {
    if (!getCurrentFunctionContext) {
      return null;
    }
    return (
      FUNCTION_DEFINITIONS.find((fn) => fn.name === getCurrentFunctionContext.name) ?? {
        name: getCurrentFunctionContext.name,
        description: "Function signature unavailable.",
        params: [],
      }
    );
  }, [getCurrentFunctionContext]);

  const insertAtCursor = (insertText: string) => {
    const start = formulaEditorRef.current?.selectionStart ?? formulaCursor;
    const end = formulaEditorRef.current?.selectionEnd ?? formulaCursor;
    const updated = `${formulaText.slice(0, start)}${insertText}${formulaText.slice(end)}`;
    setFormulaText(updated);
    const nextPos = start + insertText.length;
    setTimeout(() => {
      formulaEditorRef.current?.focus();
      formulaEditorRef.current?.setSelectionRange(nextPos, nextPos);
      setFormulaCursor(nextPos);
    }, 0);
  };

  const insertFunctionTemplate = (item: FormulaSuggestion) => {
    const template = item.kind === "Function" ? `${item.name}()` : item.name;
    const beforeCursor = formulaText.slice(0, formulaCursor);
    const tokenMatch = beforeCursor.match(/([A-Za-z_][A-Za-z0-9_.]*)$/);
    if (!tokenMatch) {
      insertAtCursor(template);
      return;
    }
    const tokenStart = (tokenMatch.index ?? beforeCursor.length) + 0;
    const updated = `${formulaText.slice(0, tokenStart)}${template}${formulaText.slice(formulaCursor)}`;
    setFormulaText(updated);
    const cursorPos =
      item.kind === "Function" ? tokenStart + template.length - 1 : tokenStart + template.length;
    setTimeout(() => {
      formulaEditorRef.current?.focus();
      formulaEditorRef.current?.setSelectionRange(cursorPos, cursorPos);
      setFormulaCursor(cursorPos);
    }, 0);
  };

  const applyActiveSuggestion = () => {
    if (functionSuggestions.length === 0) {
      return;
    }
    const safeIndex = Math.min(Math.max(activeSuggestionIndex, 0), functionSuggestions.length - 1);
    insertFunctionTemplate(functionSuggestions[safeIndex]);
  };

  const getSuggestionTextParts = (name: string, tokenUpper: string) => {
    if (!tokenUpper) {
      return { pre: name, match: "", post: "" };
    }
    const upperName = name.toUpperCase();
    const idx = upperName.indexOf(tokenUpper);
    if (idx < 0) {
      return { pre: name, match: "", post: "" };
    }
    const pre = name.slice(0, idx);
    const match = name.slice(idx, idx + tokenUpper.length);
    const post = name.slice(idx + tokenUpper.length);
    return { pre, match, post };
  };

  const addSelectionToFormula = async () => {
    try {
      const selection = await getCurrentSelectionAddress();
      const token = selection.address.includes("!")
        ? selection.address
        : `${selection.sheet}!${selection.address}`;
      const rangeCandidates = (
        namedRanges.length > 0 ? namedRanges : await getNamedRanges()
      ).filter((item) => item.kind === "NamedRange");
      const match = rangeCandidates.find(
        (item) =>
          item.isRange &&
          normalizeAddressForCompare(item.address) === normalizeAddressForCompare(token)
      );
      insertAtCursor(match ? match.name : token);
      setStatusType("success");
      setStatus("Grid selection inserted into formula.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Insert selection failed: ${message}`);
    }
  };

  const generatedLambda = useMemo(() => {
    const cleanName = customFunctionName.trim();
    const params = customFunctionParams
      .map((param) => param.trim())
      .filter((param) => param.length > 0);
    const paramList = params.join(",");
    const body = customFunctionBody.trim() || "0";
    return {
      name: cleanName,
      formula: `=LAMBDA(${paramList}${paramList ? "," : ""}${body})`,
      invokeExample: `=${cleanName}(${params.map((_, idx) => idx + 1).join(",")})`,
    };
  }, [customFunctionName, customFunctionBody, customFunctionParams]);

  const runFormulaTest = async () => {
    const formulaToRun = formulaText.trim();
    if (!formulaToRun) {
      setFormulaOutput(null);
      return;
    }
    try {
      const result = await evaluateFormula(formulaToRun);
      setFormulaOutput(result);
      setStatusType("success");
      setStatus(`Formula test completed (${result.address}).`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Formula test failed: ${message}`);
      setFormulaOutput(null);
    }
  };

  const syncFormulaFromActiveCell = async (requireFormula: boolean): Promise<boolean> => {
    const activeCell = await getActiveCellFormulaState();
    setActiveFormulaPrompt(activeCell);
    if (!activeCell.hasFormula) {
      if (requireFormula) {
        throw new Error(`Cell ${activeCell.sheet}!${activeCell.address} does not contain a formula.`);
      }
      return false;
    }
    setFormulaMode("Formula");
    setFormulaText(activeCell.formula);
    setTestFormulaCall(activeCell.formula);
    return true;
  };

  const pullActiveCellFormula = async () => {
    await runAction("Open active cell formula in editor", async () => {
      await syncFormulaFromActiveCell(true);
    });
  };

  useEffect(() => {
    if (!autoCaptureExcelFormula) {
      return undefined;
    }
    let disposed = false;
    let busy = false;
    const timerId = window.setInterval(() => {
      if (disposed || busy) {
        return;
      }
      busy = true;
      void (async () => {
        try {
          const activeCell = await getActiveCellFormulaState();
          if (disposed) {
            return;
          }
          setActiveFormulaPrompt((prev) => {
            if (
              prev &&
              prev.sheet === activeCell.sheet &&
              prev.address === activeCell.address &&
              prev.formula === activeCell.formula &&
              prev.hasFormula === activeCell.hasFormula
            ) {
              return prev;
            }
            return activeCell;
          });
          if (!activeCell.hasFormula) {
            return;
          }
          const formulaTabOpen = primaryTab === "Names" && secondaryTabId === "functions";
          if (autoOpenFormulaTab && !formulaTabOpen) {
            setPrimaryTab("Names");
            setSecondaryTabId("functions");
          }
          const editorHasFocus = document.activeElement === formulaEditorRef.current;
          if (!editorHasFocus && activeCell.formula !== formulaText) {
            setFormulaMode("Formula");
            setFormulaText(activeCell.formula);
            setTestFormulaCall(activeCell.formula);
          }
        } catch {
          // Ignore sync misses from transient Excel editing states.
        } finally {
          busy = false;
        }
      })();
    }, 1100);

    return () => {
      disposed = true;
      window.clearInterval(timerId);
    };
  }, [autoCaptureExcelFormula, autoOpenFormulaTab, formulaText, primaryTab, secondaryTabId]);

  useEffect(() => {
    let disposed = false;
    const runtime = (window as unknown as { OfficeRuntime?: typeof OfficeRuntime }).OfficeRuntime;

    const readDocumentSignal = (): string | null => {
      try {
        return (Office.context.document.settings.get(OPEN_FORMULA_EDITOR_SIGNAL_KEY) as string) || null;
      } catch {
        return null;
      }
    };

    const clearDocumentSignal = async (): Promise<void> => {
      try {
        Office.context.document.settings.remove(OPEN_FORMULA_EDITOR_SIGNAL_KEY);
        await new Promise<void>((resolve) => {
          Office.context.document.settings.saveAsync(() => resolve());
        });
      } catch {
        // Ignore clear failures.
      }
    };

    const consumeSignal = async () => {
      try {
        const runtimeSignal = runtime?.storage
          ? await runtime.storage.getItem(OPEN_FORMULA_EDITOR_SIGNAL_KEY)
          : null;
        const documentSignal = readDocumentSignal();
        const signal = runtimeSignal || documentSignal;
        if (!signal || disposed) {
          return;
        }
        if (runtime?.storage) {
          await runtime.storage.removeItem(OPEN_FORMULA_EDITOR_SIGNAL_KEY);
        }
        await clearDocumentSignal();
        if (disposed) {
          return;
        }

        setPrimaryTab("Names");
        setSecondaryTabId("functions");
        if (signal === OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL || signal === FORMULA_PULL_SIGNAL) {
          await pullActiveCellFormula();
          return;
        }
        if (signal === FORMULA_APPLY_SIGNAL) {
          await pushFormulaToActiveCell();
          return;
        }
        if (signal === FORMULA_BEAUTIFY_SIGNAL) {
          await executeFormulaEditorCommand("beautifyFormula");
          return;
        }
        if (signal === FORMULA_INSERT_SELECTION_SIGNAL) {
          await addSelectionToFormula();
          return;
        }
      } catch {
        // Ignore signal read errors to keep editor responsive.
      }
    };

    void consumeSignal();
    const timerId = window.setInterval(() => {
      void consumeSignal();
    }, 1000);

    return () => {
      disposed = true;
      window.clearInterval(timerId);
    };
  }, []);

  const pushFormulaToActiveCell = async () => {
    await runAction("Apply editor formula to active cell", async () => {
      const normalized = formulaText.trim();
      if (!normalized) {
        throw new Error("Formula editor is empty.");
      }
      await applyFormulaToActiveCell(normalized);
      const refreshed = await getActiveCellFormulaState();
      setActiveFormulaPrompt(refreshed);
    });
  };

  useEffect(() => {
    setTestFormulaCall(formulaText);
    if (!(primaryTab === "Names" && secondaryTabId === "functions")) {
      return undefined;
    }
    const formulaToRun = formulaText.trim();
    if (!formulaToRun) {
      setFormulaOutput(null);
      return undefined;
    }

    const timer = setTimeout(async () => {
      try {
        const result = await evaluateFormula(formulaToRun);
        setFormulaOutput(result);
      } catch {
        setFormulaOutput(null);
      }
    }, 450);

    return () => clearTimeout(timer);
  }, [formulaText, primaryTab, secondaryTabId]);

  useEffect(() => {
    if (formulaMode === "Function") {
      setFormulaText(generatedLambda.formula);
      setTestFormulaCall(generatedLambda.invokeExample);
    }
  }, [formulaMode, generatedLambda.formula, generatedLambda.invokeExample]);

  useEffect(() => {
    if (!formulaContextMenu) {
      return undefined;
    }
    const closeMenu = () => setFormulaContextMenu(null);
    window.addEventListener("click", closeMenu);
    window.addEventListener("scroll", closeMenu, true);
    return () => {
      window.removeEventListener("click", closeMenu);
      window.removeEventListener("scroll", closeMenu, true);
    };
  }, [formulaContextMenu]);

  const cycleFormulaReferenceLock = () => {
    const editor = formulaEditorRef.current;
    if (!editor) {
      return;
    }
    const selectionStart = editor.selectionStart ?? formulaCursor;
    const selectionEnd = editor.selectionEnd ?? selectionStart;
    const cycled = cycleCellOrRangeLockAtCursor(
      formulaText,
      selectionStart,
      selectionStart,
      selectionEnd
    );
    setFormulaText(cycled.updated);
    setTimeout(() => {
      formulaEditorRef.current?.focus();
      formulaEditorRef.current?.setSelectionRange(cycled.nextStart, cycled.nextEnd);
      setFormulaCursor(cycled.nextStart);
    }, 0);
  };

  const executeFormulaEditorCommand = async (command: string) => {
    try {
      const editor = formulaEditorRef.current;
      if (!editor) {
        return;
      }
      const start = editor.selectionStart ?? formulaCursor;
      const end = editor.selectionEnd ?? start;
      const selectedText = formulaText.slice(start, end);

      if (command === "goToDefinition") {
        if (
          !formulaReferenceHint ||
          formulaReferenceHint.kind === "Function" ||
          formulaReferenceHint.kind === "Local"
        ) {
          setStatusType("error");
          setStatus("Go to definition is available for names and tables.");
          return;
        }
        const nameMatch = namedRanges.find(
          (item) => item.kind === "NamedRange" && item.name === formulaReferenceHint.name
        );
        if (nameMatch && nameMatch.isRange) {
          await selectNamedRangeAddress(nameMatch.address, nameMatch.sheet);
          setStatusType("success");
          setStatus(`Navigated to ${nameMatch.name}.`);
          return;
        }
        const tableMatch = tables.find((item) => item.name === formulaReferenceHint.name);
        if (tableMatch) {
          await selectNamedRangeAddress(tableMatch.address, tableMatch.sheet);
          setStatusType("success");
          setStatus(`Navigated to ${tableMatch.name}.`);
          return;
        }
        setStatusType("error");
        setStatus("Definition target is not a selectable range.");
        return;
      }

      if (command === "format" || command === "beautifyFormula") {
        const normalized = prettyFormatFormula(formulaText);
        setFormulaText(normalized);
        setStatusType("success");
        setStatus("Formula beautified.");
        return;
      }

      if (command === "toggleSelectionRange") {
        cycleFormulaReferenceLock();
        return;
      }

      if (command === "cut") {
        if (start === end) {
          return;
        }
        await navigator.clipboard.writeText(selectedText);
        const nextText = `${formulaText.slice(0, start)}${formulaText.slice(end)}`;
        setFormulaText(nextText);
        setTimeout(() => {
          formulaEditorRef.current?.focus();
          formulaEditorRef.current?.setSelectionRange(start, start);
          setFormulaCursor(start);
        }, 0);
        return;
      }

      if (command === "copy") {
        if (start === end) {
          return;
        }
        await navigator.clipboard.writeText(selectedText);
        return;
      }

      if (command === "paste") {
        const text = await navigator.clipboard.readText();
        const nextText = `${formulaText.slice(0, start)}${text}${formulaText.slice(end)}`;
        const nextPos = start + text.length;
        setFormulaText(nextText);
        setTimeout(() => {
          formulaEditorRef.current?.focus();
          formulaEditorRef.current?.setSelectionRange(nextPos, nextPos);
          setFormulaCursor(nextPos);
        }, 0);
        return;
      }

      setStatusType("success");
      setStatus(`${command} is available in this menu but not yet fully implemented.`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Formula editor command failed: ${message}`);
    }
  };

  const handleFormulaEditorKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const hasSuggestions = shouldShowIntellisense && functionSuggestions.length > 0;
    if (event.key === "Escape") {
      setFormulaContextMenu(null);
    }
    if (hasSuggestions && (event.key === "Tab" || (event.key === "Enter" && !event.shiftKey))) {
      event.preventDefault();
      applyActiveSuggestion();
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      const editor = event.currentTarget;
      const start = editor.selectionStart ?? formulaCursor;
      const end = editor.selectionEnd ?? start;
      const tabText = "  ";
      const nextText = `${formulaText.slice(0, start)}${tabText}${formulaText.slice(end)}`;
      const nextPos = start + tabText.length;
      setFormulaText(nextText);
      setTimeout(() => {
        formulaEditorRef.current?.focus();
        formulaEditorRef.current?.setSelectionRange(nextPos, nextPos);
        setFormulaCursor(nextPos);
      }, 0);
      return;
    }
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void runFormulaTest();
      return;
    }
    if (event.key === "F4") {
      event.preventDefault();
      cycleFormulaReferenceLock();
      return;
    }
  };

  const applySingleOperation = (
    input: string,
    operation: "Add" | "Remove" | "Replace",
    position: "Prefix" | "Suffix",
    delimiter: BulkState["delimiter"],
    textValue: string,
    replaceWith: string
  ) => {
    const normalizedText = textValue.trim();
    const separator = getDelimiterText(delimiter);
    let nextName = input;
    const lower = (value: string) => value.toLowerCase();

    if (operation === "Add" && normalizedText) {
      if (position === "Prefix") {
        nextName = `${normalizedText}${separator}${nextName}`;
      } else {
        nextName = `${nextName}${separator}${normalizedText}`;
      }
    }

    if (operation === "Remove" && normalizedText) {
      if (position === "Prefix") {
        const prefixCandidates = [
          `${normalizedText}${separator}`,
          normalizedText,
          `${normalizedText}_`,
          `${normalizedText}.`,
        ];
        for (const candidate of prefixCandidates) {
          if (lower(nextName).startsWith(lower(candidate))) {
            const stripped = nextName.slice(candidate.length);
            nextName = stripped.length > 0 ? stripped : input;
            break;
          }
        }
      } else {
        const suffixCandidates = [
          `${separator}${normalizedText}`,
          normalizedText,
          `_${normalizedText}`,
          `.${normalizedText}`,
        ];
        for (const candidate of suffixCandidates) {
          if (lower(nextName).endsWith(lower(candidate))) {
            const stripped = nextName.slice(0, -candidate.length);
            nextName = stripped.length > 0 ? stripped : input;
            break;
          }
        }
      }
    }

    if (operation === "Replace" && normalizedText) {
      nextName = nextName.split(normalizedText).join(replaceWith.trim());
    }

    return nextName;
  };

  const applyBulkRule = (name: string) => {
    let nextName = applySingleOperation(
      name,
      bulkState.operation === "Delete" ? "Add" : bulkState.operation,
      bulkState.position,
      bulkState.delimiter,
      bulkState.textValue,
      bulkState.replaceWith
    );
    nextName = applyCaseTransform(nextName, bulkState.caseTransform);

    if (bulkState.enableSecondEdit && bulkState.operation !== "Delete") {
      nextName = applySingleOperation(
        nextName,
        bulkState.secondOperation,
        bulkState.secondPosition,
        bulkState.secondDelimiter,
        bulkState.secondTextValue,
        bulkState.secondReplaceWith
      );
      nextName = applyCaseTransform(nextName, bulkState.secondCaseTransform);
    }

    return nextName;
  };

  const bulkPreviewRows: BulkPreviewRow[] = useMemo(() => {
    const selected = namedRanges.filter(
      (item) => item.kind === "NamedRange" && selectedIds.has(item.id)
    );
    const existingByScope = new Map<string, Set<string>>();

    const getScopeKey = (scopeType: "Workbook" | "Worksheet", scope: string) =>
      `${scopeType}::${scope}`;

    namedRanges.forEach((item) => {
      if (item.kind !== "NamedRange") {
        return;
      }
      const key = getScopeKey(item.scopeType, item.scope);
      if (!existingByScope.has(key)) {
        existingByScope.set(key, new Set());
      }
      existingByScope.get(key)?.add(item.name.toLowerCase());
    });

    const preview: BulkPreviewRow[] = [];

    selected.forEach((item) => {
      if (bulkState.operation === "Delete") {
        preview.push({
          id: item.id,
          oldName: item.name,
          newName: bulkState.deleteName ? "(name deleted)" : item.name,
          scopeType: item.scopeType,
          scope: item.scope,
        });
        return;
      }

      const transformed = applyBulkRule(item.name);

      const scopeKey = getScopeKey(item.scopeType, item.scope);
      const usedNames = existingByScope.get(scopeKey) ?? new Set<string>();
      usedNames.delete(item.name.toLowerCase());

      let candidate = transformed;
      let index = 1;
      while (usedNames.has(candidate.toLowerCase())) {
        candidate = `${transformed}_${index}`;
        index += 1;
      }
      usedNames.add(candidate.toLowerCase());
      existingByScope.set(scopeKey, usedNames);

      preview.push({
        id: item.id,
        oldName: item.name,
        newName: candidate,
        scopeType: item.scopeType,
        scope: item.scope,
      });
    });

    return preview;
  }, [
    namedRanges,
    selectedIds,
    bulkState.operation,
    bulkState.deleteName,
    bulkState.position,
    bulkState.delimiter,
    bulkState.textValue,
    bulkState.replaceWith,
    bulkState.enableSecondEdit,
    bulkState.secondOperation,
    bulkState.secondPosition,
    bulkState.secondDelimiter,
    bulkState.secondTextValue,
    bulkState.secondReplaceWith,
    bulkState.caseTransform,
    bulkState.secondCaseTransform,
  ]);

  const toggleSort = (column: SortColumn) => {
    if (sortColumn === column) {
      setSortDirection(sortDirection === "asc" ? "desc" : "asc");
      return;
    }
    setSortColumn(column);
    setSortDirection("asc");
  };

  const toggleTableSort = (column: TableSortColumn) => {
    if (tableSortColumn === column) {
      setTableSortDirection(tableSortDirection === "asc" ? "desc" : "asc");
      return;
    }
    setTableSortColumn(column);
    setTableSortDirection("asc");
  };

  const allVisibleTablesSelected =
    filteredAndSortedTables.length > 0 &&
    filteredAndSortedTables.every((item) => selectedTableIds.has(item.id));

  const toggleTableSelection = (id: string) => {
    setSelectedTableIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const toggleVisibleTableSelection = () => {
    setSelectedTableIds((prev) => {
      const next = new Set(prev);
      if (allVisibleTablesSelected) {
        filteredAndSortedTables.forEach((item) => next.delete(item.id));
      } else {
        filteredAndSortedTables.forEach((item) => next.add(item.id));
      }
      return next;
    });
  };

  const applyTableNameRule = (name: string) => {
    const first = applySingleOperation(
      name,
      tableBulkState.operation,
      tableBulkState.position,
      tableBulkState.delimiter,
      tableBulkState.textValue,
      tableBulkState.replaceWith
    );
    return applyCaseTransform(first, tableBulkState.caseTransform);
  };

  const tableBulkPreview = useMemo(() => {
    return tables
      .filter((item) => selectedTableIds.has(item.id))
      .map((item) => ({
        id: item.id,
        oldName: item.name,
        newName: applyTableNameRule(item.name),
        sheet: item.sheet,
      }));
  }, [
    tables,
    selectedTableIds,
    tableBulkState.operation,
    tableBulkState.position,
    tableBulkState.delimiter,
    tableBulkState.textValue,
    tableBulkState.replaceWith,
    tableBulkState.caseTransform,
  ]);

  const toggleRowSelection = (id: string) => {
    const record = namedRanges.find((item) => item.id === id);
    if (!record || record.kind !== "NamedRange") {
      return;
    }
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const toggleVisibleSelection = () => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (allVisibleSelected) {
        selectableVisibleRanges.forEach((item) => next.delete(item.id));
      } else {
        selectableVisibleRanges.forEach((item) => next.add(item.id));
      }
      return next;
    });
  };

  const openEdit = (record: NamedRangeRecord) => {
    if (record.kind === "Shape") {
      void openShapeEditorFromNames(record);
      return;
    }
    setEditState({
      open: true,
      record,
      name: record.name,
      address: record.address,
      fallbackSheet: record.sheet,
      caseTransform: "none",
    });
  };

  const openMove = (record: NamedRangeRecord) => {
    setMoveState({
      open: true,
      record,
      address: record.address,
      fallbackSheet: record.sheet,
    });
  };

  const openDelete = (record: NamedRangeRecord) => {
    if (record.kind !== "NamedRange") {
      return;
    }
    setDeleteState({
      open: true,
      record,
      deleteName: true,
      deleteValues: false,
    });
  };

  const goToRange = async (record: NamedRangeRecord) => {
    if (record.kind === "Shape") {
      await runAction("Open shape editor", async () => {
        await openShapeEditorFromNames(record);
      });
      return;
    }
    await runAction("Navigate to range", async () => {
      await selectNamedRangeAddress(record.address, record.sheet);
    });
  };

  const hydrateShapeEditor = (details: ShapeEditorRecord) => {
    const effectiveAnchor = details.anchorAddress || "A1";
    setPrimaryTab("Format");
    setSecondaryTabId("shapes");
    setActiveShape({
      name: details.name,
      shapeType: details.shapeType,
      sheet: details.sheet,
      anchorAddress: effectiveAnchor,
      width: details.width,
      height: details.height,
    });
    setShapeState((prev) => ({
      ...prev,
      shapeName: details.name,
      shapeType: details.shapeType,
      geometricShapeType: details.geometricShapeType,
      left: details.left,
      top: details.top,
      zOrderPosition: details.zOrderPosition,
      fillColor: details.fillColor,
      outlineColor: details.outlineColor,
      fontColor: details.fontColor,
      text: details.text,
      lineWeight: details.lineWeight,
      fillTransparency: details.fillTransparency,
      width: details.width,
      height: details.height,
      rotation: details.rotation,
      textHorizontalAlignment: details.textHorizontalAlignment,
      textVerticalAlignment: details.textVerticalAlignment,
      fontSize: details.fontSize,
      bold: details.bold,
      italic: details.italic,
      lockAspectRatio: details.lockAspectRatio,
    }));
  };

  const openShapeEditorFromNames = async (record: NamedRangeRecord) => {
    const details: ShapeEditorRecord = await getShapeEditorRecord(record.sheet, record.name);
    hydrateShapeEditor(details);
  };

  useEffect(() => {
    let cancelled = false;
    const register = async () => {
      try {
        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          await context.sync();

          for (const sheet of worksheets.items) {
            const shapes = sheet.shapes;
            shapes.load("items/name");
            await context.sync();
            for (const shape of shapes.items) {
              const sheetName = sheet.name;
              const shapeName = shape.name;
              const handler = await shape.onActivated.add(async () => {
                if (cancelled) {
                  return;
                }
                try {
                  const details = await getShapeEditorRecord(sheetName, shapeName);
                  hydrateShapeEditor(details);
                } catch {
                  // ignore host or shape-state sync failures
                }
              });
              shapeActivationHandlersRef.current.push(handler as { remove: () => Promise<void> | void });
            }
          }
        });
      } catch {
        // host may not support shape activation events
      }
    };

    void register();
    return () => {
      cancelled = true;
      const handlers = [...shapeActivationHandlersRef.current];
      shapeActivationHandlersRef.current = [];
      handlers.forEach((handler) => {
        try {
          const removed = handler.remove();
          if (removed instanceof Promise) {
            void removed;
          }
        } catch {
          // ignore cleanup errors
        }
      });
    };
  }, [namedRanges.length]);

  const submitMove = async () => {
    if (!moveState.record) {
      return;
    }

    setIsSubmitting(true);
    try {
      await moveNamedRange(
        moveState.record.scopeType,
        moveState.record.scope,
        moveState.record.name,
        moveState.address.trim(),
        moveState.fallbackSheet
      );
      setMoveState({ open: false, record: null, address: "", fallbackSheet: "" });
      await loadRanges();
      setStatusType("success");
      setStatus("Named range moved.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Move failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const submitDelete = async () => {
    if (!deleteState.record) {
      return;
    }
    if (!deleteState.deleteName && !deleteState.deleteValues) {
      setStatusType("error");
      setStatus("Select at least one delete option.");
      return;
    }

    setIsSubmitting(true);
    try {
      await deleteNamedRangeWithOptions(
        deleteState.record.scopeType,
        deleteState.record.scope,
        deleteState.record.name,
        deleteState.deleteName,
        deleteState.deleteValues
      );
      setDeleteState({
        open: false,
        record: null,
        deleteName: true,
        deleteValues: false,
      });
      await loadRanges();
      setStatusType("success");
      setStatus("Delete operation completed.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Delete failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const submitEdit = async () => {
    if (!editState.record) {
      return;
    }

    setIsSubmitting(true);
    try {
      const transformedName = applyCaseTransform(editState.name.trim(), editState.caseTransform);
      await updateNamedRange(
        editState.record.scopeType,
        editState.record.scope,
        editState.record.name,
        transformedName,
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
      await loadRanges();
      setStatusType("success");
      setStatus("Named range updated.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Update failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const useGridSelection = async () => {
    try {
      const selection = await getCurrentSelectionAddress();
      setEditState((prev) => ({
        ...prev,
        address: selection.address,
        fallbackSheet: selection.sheet,
      }));
      setStatusType("success");
      setStatus("Current grid selection captured.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Capture selection failed: ${message}`);
    }
  };

  const useGridSelectionForMove = async () => {
    try {
      const selection = await getCurrentSelectionAddress();
      setMoveState((prev) => ({
        ...prev,
        address: selection.address,
        fallbackSheet: selection.sheet,
      }));
      setStatusType("success");
      setStatus("Current grid selection captured.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Capture selection failed: ${message}`);
    }
  };

  const submitBulkUpdate = async () => {
    const selectedRecords = namedRanges.filter((item) => selectedIds.has(item.id));
    if (selectedRecords.length === 0) {
      return;
    }
    if (bulkState.operation === "Delete" && !bulkState.deleteName && !bulkState.deleteValues) {
      setStatusType("error");
      setStatus("For delete, select delete name and/or delete values.");
      return;
    }

    setIsSubmitting(true);
    try {
      for (const record of selectedRecords) {
        if (bulkState.operation === "Delete") {
          await deleteNamedRangeWithOptions(
            record.scopeType,
            record.scope,
            record.name,
            bulkState.deleteName,
            bulkState.deleteValues
          );
          continue;
        }

        const preview = bulkPreviewRows.find((item) => item.id === record.id);
        if (!preview) {
          continue;
        }

        if (preview.newName !== record.name) {
          await updateNamedRange(
            record.scopeType,
            record.scope,
            record.name,
            preview.newName,
            record.address,
            record.sheet
          );
        }
      }

      resetBulkState();
      await loadRanges();
      setStatusType("success");
      setStatus(`Bulk operation completed for ${selectedRecords.length} range(s).`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Bulk update failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const resetTableBulkState = () => {
    setTableBulkState({
      open: false,
      operation: "Add",
      position: "Prefix",
      delimiter: "underscore",
      textValue: "",
      replaceWith: "",
      caseTransform: "none",
    });
  };

  const startInlineTableEdit = (table: TableRecord) => {
    setInlineTableEdit({ id: table.id, value: table.name });
  };

  const saveInlineTableEdit = async (table: TableRecord) => {
    if (!inlineTableEdit) {
      return;
    }
    const newName = inlineTableEdit.value.trim();
    if (!newName || newName === table.name) {
      setInlineTableEdit(null);
      return;
    }

    await runAction("Rename table", async () => {
      await updateTableName(table.sheet, table.name, newName);
      await loadTables();
    });
    setInlineTableEdit(null);
  };

  const submitTableBulkUpdate = async () => {
    if (tableBulkPreview.length === 0) {
      return;
    }

    setIsSubmitting(true);
    try {
      const usedBySheet = new Map<string, Set<string>>();
      tables.forEach((t) => {
        if (!usedBySheet.has(t.sheet)) {
          usedBySheet.set(t.sheet, new Set());
        }
        usedBySheet.get(t.sheet)?.add(t.name.toLowerCase());
      });

      for (const preview of tableBulkPreview) {
        const current = tables.find((t) => t.id === preview.id);
        if (!current) {
          continue;
        }

        const used = usedBySheet.get(current.sheet) ?? new Set<string>();
        used.delete(current.name.toLowerCase());
        let candidate = preview.newName;
        let index = 1;
        while (used.has(candidate.toLowerCase())) {
          candidate = `${preview.newName}_${index}`;
          index += 1;
        }
        used.add(candidate.toLowerCase());
        usedBySheet.set(current.sheet, used);

        if (candidate !== current.name) {
          await updateTableName(current.sheet, current.name, candidate);
        }
      }

      resetTableBulkState();
      await loadTables();
      setStatusType("success");
      setStatus("Bulk table update completed.");
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Bulk table update failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const renderPlaceholder = () => (
    <div className={styles.placeholder}>
      {primaryTab} tools will appear here. Select commands from the black toolbar to run workbook
      actions.
    </div>
  );

  const formulaLineCount = useMemo(() => {
    if (!formulaText) {
      return 1;
    }
    return formulaText.split(/\r?\n/).length;
  }, [formulaText]);

  const formatStylePresets: Array<{ preset: CellStylePreset; description: string }> = [
    {
      preset: "Input Cell",
      description: "Light input shading for user-entered values.",
    },
    {
      preset: "Parameter Cell",
      description: "Emphasized parameter values with stronger font treatment.",
    },
    {
      preset: "Header",
      description: "Primary table/report header styling.",
    },
    {
      preset: "Subheader",
      description: "Secondary section heading styling.",
    },
  ];

  const renderRangesTable = () => (
    <>
      <div className={styles.topActions}>
        <Button
          className={styles.miniBtn}
          onClick={() => void loadRanges(true)}
          disabled={isLoadingRanges}
        >
          Refresh
        </Button>
        <Button className={styles.miniBtn} onClick={clearFilters} disabled={!hasActiveFilters}>
          Clear Filters
        </Button>
        <Button
          className={styles.miniBtn}
          onClick={() => setBulkState((prev) => ({ ...prev, open: true }))}
          disabled={selectedIds.size === 0}
        >
          Bulk Update ({selectedIds.size})
        </Button>
        <Text>{`Showing ${filteredAndSortedRanges.length} of ${namedRanges.length}`}</Text>
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={`${styles.th} ${styles.rowCheckCol}`}>
                <input
                  type="checkbox"
                  checked={allVisibleSelected}
                  onChange={toggleVisibleSelection}
                />
              </th>
              <th className={`${styles.th} ${styles.rowActionCol}`}>Actions</th>
              {RANGE_COLUMNS.map((col) => (
                <th key={col} className={`${styles.th} ${col === "type" ? styles.typeCol : ""}`}>
                  <div className={styles.thHeader}>
                    <div className={styles.thLabelRow}>
                      <button
                        type="button"
                        className={styles.sortBtn}
                        onClick={() => toggleSort(col)}
                      >
                        {col[0].toUpperCase() + col.slice(1)}
                        {sortColumn === col ? (sortDirection === "asc" ? " ▲" : " ▼") : ""}
                      </button>
                    </div>
                    <div className={styles.thFilterRow}>
                      <Filter20Regular className={styles.thFilterIcon} />
                      <input
                        className={styles.filterNative}
                        placeholder=""
                        value={columnFilters[col]}
                        onChange={(event) =>
                          setColumnFilters((prev) => ({
                            ...prev,
                            [col]: event.target.value,
                          }))
                        }
                      />
                    </div>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredAndSortedRanges.map((item) => (
              <tr key={item.id}>
                <td className={`${styles.td} ${styles.rowCheckCol}`}>
                  <input
                    type="checkbox"
                    checked={selectedIds.has(item.id)}
                    disabled={item.kind !== "NamedRange"}
                    onChange={() => toggleRowSelection(item.id)}
                  />
                </td>
                <td className={`${styles.td} ${styles.rowActionCol}`}>
                  <div className={styles.rowActions}>
                    <Button
                      className={styles.rowActionTextBtn}
                      title="Edit name and address"
                      disabled={item.kind === "NamedRange" && !item.isRange}
                      onClick={() => openEdit(item)}
                    >
                      Edit
                    </Button>
                    <Button
                      className={styles.rowActionTextBtn}
                      title="Move named range"
                      disabled={!item.isRange}
                      onClick={() => openMove(item)}
                    >
                      Move
                    </Button>
                    <Button
                      className={styles.rowActionTextBtn}
                      title="Delete named range"
                      disabled={item.kind !== "NamedRange"}
                      onClick={() => openDelete(item)}
                    >
                      Delete
                    </Button>
                  </div>
                </td>
                <td className={styles.td}>
                  {item.isRange || item.kind === "Shape" ? (
                    <button
                      type="button"
                      className={styles.nameLink}
                      onClick={() => void goToRange(item)}
                    >
                      {item.name}
                    </button>
                  ) : (
                    <Text>{item.name}</Text>
                  )}
                </td>
                <td className={styles.td}>{item.address}</td>
                <td className={styles.td}>{item.sheet}</td>
                <td className={styles.td}>{item.scope}</td>
                <td className={`${styles.td} ${styles.typeCol}`}>{item.type}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );

  const renderTablesTable = () => (
    <>
      <div className={styles.topActions}>
        <Button
          className={styles.miniBtn}
          onClick={() => void loadTables(true)}
          disabled={isLoadingTables}
        >
          Refresh
        </Button>
        <Button
          className={styles.miniBtn}
          onClick={clearTableFilters}
          disabled={!TABLE_COLUMNS.some((col) => tableFilters[col].trim().length > 0)}
        >
          Clear Filters
        </Button>
        <Button
          className={styles.miniBtn}
          onClick={() => setTableBulkState((prev) => ({ ...prev, open: true }))}
          disabled={selectedTableIds.size === 0}
        >
          Bulk Edit ({selectedTableIds.size})
        </Button>
        <Text>{`Showing ${filteredAndSortedTables.length} of ${tables.length}`}</Text>
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={`${styles.th} ${styles.rowCheckCol}`}>
                <input
                  type="checkbox"
                  checked={allVisibleTablesSelected}
                  onChange={toggleVisibleTableSelection}
                />
              </th>
              <th className={`${styles.th} ${styles.rowActionCol}`}>Actions</th>
              {TABLE_COLUMNS.map((col) => (
                <th key={col} className={styles.th}>
                  <div className={styles.thHeader}>
                    <div className={styles.thLabelRow}>
                      <button type="button" className={styles.sortBtn} onClick={() => toggleTableSort(col)}>
                        {col[0].toUpperCase() + col.slice(1)}
                        {tableSortColumn === col ? (tableSortDirection === "asc" ? " ▲" : " ▼") : ""}
                      </button>
                    </div>
                    <div className={styles.thFilterRow}>
                      <Filter20Regular className={styles.thFilterIcon} />
                      <input
                        className={styles.filterNative}
                        value={tableFilters[col]}
                        onChange={(event) =>
                          setTableFilters((prev) => ({
                            ...prev,
                            [col]: event.target.value,
                          }))
                        }
                      />
                    </div>
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredAndSortedTables.map((table) => (
              <tr key={table.id}>
                <td className={`${styles.td} ${styles.rowCheckCol}`}>
                  <input
                    type="checkbox"
                    checked={selectedTableIds.has(table.id)}
                    onChange={() => toggleTableSelection(table.id)}
                  />
                </td>
                <td className={`${styles.td} ${styles.rowActionCol}`}>
                  <div className={styles.rowActions}>
                    {inlineTableEdit?.id === table.id ? (
                      <>
                        <Button
                          className={styles.rowActionTextBtn}
                          title="Save"
                          onClick={() => void saveInlineTableEdit(table)}
                        >
                          Save
                        </Button>
                        <Button
                          className={styles.rowActionTextBtn}
                          title="Cancel"
                          onClick={() => setInlineTableEdit(null)}
                        >
                          Cancel
                        </Button>
                      </>
                    ) : (
                      <Button
                        className={styles.rowActionTextBtn}
                        title="Inline edit"
                        onClick={() => startInlineTableEdit(table)}
                      >
                        Edit
                      </Button>
                    )}
                  </div>
                </td>
                <td className={styles.td}>
                  {inlineTableEdit?.id === table.id ? (
                    <input
                      className={styles.filterNative}
                      value={inlineTableEdit.value}
                      onChange={(event) =>
                        setInlineTableEdit({ id: table.id, value: event.target.value })
                      }
                    />
                  ) : (
                    table.name
                  )}
                </td>
                <td className={styles.td}>{table.address}</td>
                <td className={styles.td}>{table.sheet}</td>
                <td className={styles.td}>{table.scope}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );

  const renderFormatStyles = () => (
    <div className={styles.sectionCard}>
      <Text className={styles.sectionTitle}>Default Cell Styles</Text>
      <Text className={styles.modalLabel}>
        Select a range in Excel, then apply a style preset.
      </Text>
      <div className={styles.stylePresetGrid}>
        {formatStylePresets.map((item) => (
          <div key={item.preset} className={styles.stylePresetCard}>
            <Text className={styles.stylePresetTitle}>{item.preset}</Text>
            <Text className={styles.stylePresetDesc}>{item.description}</Text>
            <Button
              className={styles.miniBtn}
              onClick={() =>
                runAction(`Apply ${item.preset}`, () => applyCellStylePreset(item.preset))
              }
            >
              Apply Style
            </Button>
          </div>
        ))}
      </div>
    </div>
  );

  const selectedShapeTheme = useMemo(() => {
    return (
      shapeThemes.find((theme) => theme.id === shapeState.selectedThemeId) ??
      shapeThemes[0] ??
      SHAPE_THEMES[0]
    );
  }, [shapeState.selectedThemeId, shapeThemes]);
  const themeColorSwatches = selectedShapeTheme.themeColors;
  const themeColorRows = useMemo(
    () => buildThemeToneGrid(selectedShapeTheme.themeColors),
    [selectedShapeTheme]
  );
  useEffect(() => {
    setCustomThemeColorsText(selectedShapeTheme.themeColors.join(","));
    setCustomStandardColorsText(selectedShapeTheme.standardColors.join(","));
  }, [selectedShapeTheme.id]);

  const shapeValueSuggestions = useMemo(() => {
    if (shapeState.bindingMode !== "NamedRange") {
      return [];
    }
    const token = shapeState.valueBinding.trim().toLowerCase();
    if (!token) {
      return [];
    }
    return namedRanges
      .filter((item) => item.kind === "NamedRange")
      .map((item) => item.name)
      .filter((name) => name.toLowerCase().includes(token))
      .slice(0, 25);
  }, [namedRanges, shapeState.bindingMode, shapeState.valueBinding]);

  const shapeFormattingOptions = useMemo(
    () =>
      ({
        fillColor: shapeState.fillColor,
        outlineColor: shapeState.outlineColor,
        fontColor: shapeState.fontColor,
        text: shapeState.text,
        lineWeight: shapeState.lineWeight,
        fillTransparency: shapeState.fillTransparency,
        width: shapeState.width,
        height: shapeState.height,
        rotation: shapeState.rotation,
        textHorizontalAlignment: shapeState.textHorizontalAlignment,
        textVerticalAlignment: shapeState.textVerticalAlignment,
        fontSize: shapeState.fontSize,
        bold: shapeState.bold,
        italic: shapeState.italic,
        lockAspectRatio: shapeState.lockAspectRatio,
      }) satisfies ShapeFormatOptions,
    [shapeState]
  );

  const insertShapeFromSelection = async () => {
    setIsSubmitting(true);
    try {
      const inserted = await addShapeOnActiveCell({
        shapeType: shapeState.shapeType,
        fillColor: shapeState.fillColor,
        outlineColor: shapeState.outlineColor,
        fontColor: shapeState.fontColor,
        text: shapeState.bindingMode === "StaticText" ? shapeState.text : "",
      });
      setActiveShape(inserted);
      setShapeState((prev) => ({
        ...prev,
        shapeName: inserted.name,
        width: inserted.width,
        height: inserted.height,
      }));
      try {
        const details = await getShapeEditorRecord(inserted.sheet, inserted.name);
        hydrateShapeEditor(details);
      } catch {
        // shape details enrichment is optional
      }
      setShapeAddOpen(false);
      setStatusType("success");
      setStatus(`Inserted ${inserted.shapeType} on ${inserted.anchorAddress}.`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Insert shape failed: ${message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const applyShapeFormatting = async () => {
    if (!activeShape) {
      return;
    }
    await runAction("Apply shape formatting", async () => {
      await updateShapeFormatting(activeShape.sheet, activeShape.name, shapeFormattingOptions);
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  useEffect(() => {
    if (!activeShape) {
      return undefined;
    }
    const timer = setTimeout(() => {
      void updateShapeFormatting(activeShape.sheet, activeShape.name, shapeFormattingOptions).catch(
        (error) => {
          const message = error instanceof Error ? error.message : String(error);
          setStatusType("error");
          setStatus(`Auto-update shape formatting failed: ${message}`);
        }
      );
    }, 250);
    return () => clearTimeout(timer);
  }, [activeShape, shapeFormattingOptions]);

  useEffect(() => {
    if (!activeShape) {
      return undefined;
    }
    const timer = setTimeout(() => {
      if (shapeState.bindingMode === "NamedRange" && shapeState.valueBinding.trim()) {
        void getNamedRangeValueText(shapeState.valueBinding.trim())
          .then((valueText) => {
            setShapeState((prev) => (prev.text === valueText ? prev : { ...prev, text: valueText }));
          })
          .catch((error) => {
            const message = error instanceof Error ? error.message : String(error);
            setStatusType("error");
            setStatus(`Named range binding failed: ${message}`);
          });
        return;
      }
      if (shapeState.bindingMode === "Formula" && shapeState.formula.trim()) {
        void applyFormulaToShapeAnchorCell(
          activeShape.sheet,
          activeShape.anchorAddress,
          shapeState.formula,
          activeShape.name
        ).catch((error) => {
          const message = error instanceof Error ? error.message : String(error);
          setStatusType("error");
          setStatus(`Formula binding failed: ${message}`);
        });
      }
    }, 300);
    return () => clearTimeout(timer);
  }, [activeShape, shapeState.bindingMode, shapeState.valueBinding, shapeState.formula]);

  const applyShapeAnchorFormula = async () => {
    if (!activeShape || !shapeState.formula.trim()) {
      return;
    }
    await runAction("Apply anchor cell formula", async () => {
      await applyFormulaToShapeAnchorCell(
        activeShape.sheet,
        activeShape.anchorAddress,
        shapeState.formula,
        activeShape.name
      );
    });
  };

  const applyNamedRangeToShapeText = async (selectedName?: string) => {
    if (!activeShape) {
      return;
    }
    const rangeName = (selectedName ?? shapeState.valueBinding).trim();
    if (!rangeName) {
      return;
    }
    await runAction("Bind shape text to named range", async () => {
      const valueText = await getNamedRangeValueText(rangeName);
      setShapeState((prev) => ({
        ...prev,
        bindingMode: "NamedRange",
        valueBinding: rangeName,
        text: valueText,
      }));
      await updateShapeFormatting(activeShape.sheet, activeShape.name, {
        ...shapeFormattingOptions,
        text: valueText,
      });
    });
  };

  const applyThemeColorToShapeField = (
    field: "fillColor" | "outlineColor" | "fontColor",
    color: string
  ) => {
    setShapeState((prev) => ({ ...prev, [field]: color }));
  };

  const createCustomShapeTheme = () => {
    const name = newShapeThemeName.trim();
    if (!name) {
      return;
    }
    const parsedThemeColors = parseThemeColorList(customThemeColorsText);
    const parsedStandardColors = parseThemeColorList(customStandardColorsText);
    const base = selectedShapeTheme;
    const nextId = `custom-${Date.now().toString()}`;
    const nextTheme: ShapeTheme = {
      id: nextId,
      label: name,
      themeColors:
        parsedThemeColors.length > 0 ? parsedThemeColors : [...base.themeColors],
      standardColors:
        parsedStandardColors.length > 0 ? parsedStandardColors : [...base.standardColors],
    };
    setShapeThemes((prev) => [...prev, nextTheme]);
    setShapeState((prev) => ({ ...prev, selectedThemeId: nextId }));
  };

  const applyShapeNoFill = () => {
    setShapeState((prev) => ({ ...prev, fillColor: NO_FILL_COLOR_TOKEN }));
  };

  const submitShapeRename = async () => {
    if (!activeShape) {
      return;
    }
    const nextName = shapeState.shapeName.trim();
    if (!nextName || nextName === activeShape.name) {
      return;
    }
    await runAction("Rename shape", async () => {
      await renameShape(activeShape.sheet, activeShape.name, nextName);
    });
    setActiveShape((prev) => (prev ? { ...prev, name: nextName } : prev));
  };

  const moveShapeToGridSelection = async () => {
    if (!activeShape) {
      return;
    }
    await runAction("Move shape to selected cell", async () => {
      const moved = await moveShapeToSelection(activeShape.sheet, activeShape.name, activeShape.anchorAddress);
      setActiveShape((prev) => (prev ? { ...prev, anchorAddress: moved.anchorAddress } : prev));
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  const applyShapePosition = async () => {
    if (!activeShape) {
      return;
    }
    await runAction("Update shape position", async () => {
      await updateShapePosition(activeShape.sheet, activeShape.name, {
        left: shapeState.left,
        top: shapeState.top,
      });
    });
  };

  const nudgeActiveShape = async (dx: number, dy: number) => {
    if (!activeShape) {
      return;
    }
    await runAction("Nudge shape", async () => {
      await nudgeShape(activeShape.sheet, activeShape.name, dx, dy);
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  const applyShapeZOrder = async (
    order: "BringToFront" | "BringForward" | "SendToBack" | "SendBackward"
  ) => {
    if (!activeShape) {
      return;
    }
    await runAction("Update shape arrange order", async () => {
      await setShapeZOrder(activeShape.sheet, activeShape.name, order);
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  const applyGeometricShapeType = async (
    shapeType: Exclude<InsertableShapeType, "TextBox">
  ) => {
    if (!activeShape) {
      return;
    }
    await runAction("Edit shape type", async () => {
      await setShapeGeometricType(activeShape.sheet, activeShape.name, shapeType);
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  const alignActiveShape = async (
    alignment: "Left" | "Center" | "Right" | "Top" | "Middle" | "Bottom"
  ) => {
    if (!activeShape) {
      return;
    }
    await runAction("Align shape", async () => {
      await alignShapeToSelection(activeShape.sheet, activeShape.name, alignment);
      const refreshed = await getShapeEditorRecord(activeShape.sheet, activeShape.name);
      hydrateShapeEditor(refreshed);
    });
  };

  const renderFormatShapes = () => (
    <>
      <div className={styles.sectionCard}>
        <Text className={styles.sectionTitle}>Shapes/Text Boxes</Text>
        <div className={styles.shapeToolbar}>
          <button type="button" className={styles.plusBtn} title="Add shape or text box" onClick={() => setShapeAddOpen(true)}>
            +
          </button>
          <Text className={styles.shapeHint}>Formatting is now applied automatically as you edit values.</Text>
        </div>
        {activeShape ? (
          <>
            <details className={styles.collapsibleSection} open>
              <summary className={styles.collapsibleSummary}>General</summary>
              <div className={styles.collapsibleBody}>
                <div className={styles.modalInlineRow}>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Name</Text>
                    <Input value={shapeState.shapeName || activeShape.name} onChange={(_, data) => setShapeState((prev) => ({ ...prev, shapeName: data.value }))} />
                  </div>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Anchor Cell</Text>
                    <Input value={`${activeShape.sheet}!${activeShape.anchorAddress}`} readOnly />
                  </div>
                </div>
                <div className={styles.topActions}>
                  <Button className={styles.miniBtn} onClick={() => void submitShapeRename()}>
                    Rename Shape
                  </Button>
                  <Button className={styles.miniBtn} onClick={() => void moveShapeToGridSelection()}>
                    Move to Grid Selection
                  </Button>
                </div>
              </div>
            </details>
            <details className={styles.collapsibleSection} open>
              <summary className={styles.collapsibleSummary}>Appearance</summary>
              <div className={styles.collapsibleBody}>
            <details className={styles.nestedSection} open>
              <summary className={styles.nestedSummary}>Shape Styles</summary>
              <div className={styles.nestedBody}>
                <div className={styles.modalRow}>
                  <Text className={styles.modalLabel}>Themes</Text>
                  <Select
                    className={styles.smallSelect}
                    value={shapeState.selectedThemeId}
                    onChange={(_, data) =>
                      setShapeState((prev) => ({ ...prev, selectedThemeId: data.value }))
                    }
                  >
                    {shapeThemes.map((theme) => (
                      <option key={theme.id} value={theme.id}>
                        {theme.label}
                      </option>
                    ))}
                  </Select>
                </div>
                <Text className={styles.modalLabel}>Theme Colors</Text>
                <div className={styles.themePaletteHeader}>
                  {selectedShapeTheme.themeColors.map((color) => (
                    <button
                      key={`theme-header-${color}`}
                      type="button"
                      className={styles.paletteSwatch}
                      style={{ backgroundColor: color }}
                      onClick={() => applyThemeColorToShapeField("fillColor", color)}
                    />
                  ))}
                </div>
                {themeColorRows.map((row, rowIndex) => (
                  <div key={`row-${rowIndex.toString()}`} className={styles.themePaletteRow}>
                    {row.map((color) => (
                      <button
                        key={`tone-${rowIndex.toString()}-${color}`}
                        type="button"
                        className={styles.paletteSwatch}
                        style={{ backgroundColor: color }}
                        onClick={() => applyThemeColorToShapeField("fillColor", color)}
                      />
                    ))}
                  </div>
                ))}
                <Text className={styles.modalLabel}>Standard Colors</Text>
                <div className={styles.themePaletteRow}>
                  {selectedShapeTheme.standardColors.map((color) => (
                    <button
                      key={`standard-${color}`}
                      type="button"
                      className={styles.paletteSwatch}
                      style={{ backgroundColor: color }}
                      onClick={() => applyThemeColorToShapeField("fillColor", color)}
                    />
                  ))}
                </div>
                <div className={styles.modalInlineRow}>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Create Custom Theme</Text>
                    <Input
                      value={newShapeThemeName}
                      onChange={(_, data) => setNewShapeThemeName(data.value)}
                    />
                  </div>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Actions</Text>
                    <Button className={styles.miniBtn} onClick={createCustomShapeTheme}>
                      Save Current as Theme
                    </Button>
                  </div>
                </div>
                <div className={styles.modalRow}>
                  <Text className={styles.modalLabel}>Theme Colors (10 hex, comma-separated)</Text>
                  <Input
                    value={customThemeColorsText}
                    onChange={(_, data) => setCustomThemeColorsText(data.value)}
                  />
                </div>
                <div className={styles.modalRow}>
                  <Text className={styles.modalLabel}>Standard Colors (10 hex, comma-separated)</Text>
                  <Input
                    value={customStandardColorsText}
                    onChange={(_, data) => setCustomStandardColorsText(data.value)}
                  />
                </div>
              </div>
            </details>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Fill</Text>
                <div className={styles.swatchRow}>
                  <Button
                    className={`${styles.noFillBtn} ${
                      shapeState.fillColor === NO_FILL_COLOR_TOKEN ? styles.noFillBtnSelected : ""
                    }`}
                    onClick={applyShapeNoFill}
                  >
                    No Fill
                  </Button>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`fill-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.fillColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      title={`Theme color ${swatch}`}
                      onClick={() => applyThemeColorToShapeField("fillColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.fillColor, "#ffffff")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, fillColor: event.target.value }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Outline</Text>
                <div className={styles.swatchRow}>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`outline-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.outlineColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      title={`Theme color ${swatch}`}
                      onClick={() => applyThemeColorToShapeField("outlineColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.outlineColor, "#000000")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, outlineColor: event.target.value }))
                  }
                />
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Font Color</Text>
                <div className={styles.swatchRow}>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`font-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.fontColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      title={`Theme color ${swatch}`}
                      onClick={() => applyThemeColorToShapeField("fontColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.fontColor, "#1f1f1f")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, fontColor: event.target.value }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Shape Text</Text>
                <Input
                  value={shapeState.text}
                  onChange={(_, data) => setShapeState((prev) => ({ ...prev, text: data.value }))}
                />
              </div>
            </div>
            <details className={styles.nestedSection} open>
              <summary className={styles.nestedSummary}>Size & Shape Effects</summary>
              <div className={styles.nestedBody}>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Width</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={20}
                  value={shapeState.width}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      width: Number(event.target.value) || prev.width,
                    }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Height</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={20}
                  value={shapeState.height}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      height: Number(event.target.value) || prev.height,
                    }))
                  }
                />
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Rotation</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={-360}
                  max={360}
                  value={shapeState.rotation}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      rotation: Number(event.target.value) || 0,
                    }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Outline Weight</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={0.25}
                  step={0.25}
                  value={shapeState.lineWeight}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      lineWeight: Number(event.target.value) || prev.lineWeight,
                    }))
                  }
                />
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Fill Transparency (%)</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={0}
                  max={100}
                  value={shapeState.fillTransparency}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      fillTransparency: Number(event.target.value) || 0,
                    }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Font Size</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={6}
                  max={72}
                  value={shapeState.fontSize}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      fontSize: Number(event.target.value) || prev.fontSize,
                    }))
                  }
                />
              </div>
            </div>
              </div>
            </details>
            <details className={styles.nestedSection} open>
              <summary className={styles.nestedSummary}>Arrange & Alignment</summary>
              <div className={styles.nestedBody}>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Position Left</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={0}
                  value={shapeState.left}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      left: Number(event.target.value) || 0,
                    }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Position Top</Text>
                <input
                  className={styles.numberInput}
                  type="number"
                  min={0}
                  value={shapeState.top}
                  onChange={(event) =>
                    setShapeState((prev) => ({
                      ...prev,
                      top: Number(event.target.value) || 0,
                    }))
                  }
                />
              </div>
            </div>
            <div className={styles.topActions}>
              <Button className={styles.miniBtn} onClick={() => void applyShapePosition()}>
                Apply Position
              </Button>
              <Button className={styles.miniBtn} onClick={() => void nudgeActiveShape(-4, 0)}>
                Nudge Left
              </Button>
              <Button className={styles.miniBtn} onClick={() => void nudgeActiveShape(4, 0)}>
                Nudge Right
              </Button>
              <Button className={styles.miniBtn} onClick={() => void nudgeActiveShape(0, -4)}>
                Nudge Up
              </Button>
              <Button className={styles.miniBtn} onClick={() => void nudgeActiveShape(0, 4)}>
                Nudge Down
              </Button>
            </div>
            <div className={styles.topActions}>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Left")}>
                Align Left
              </Button>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Center")}>
                Align Center
              </Button>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Right")}>
                Align Right
              </Button>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Top")}>
                Align Top
              </Button>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Middle")}>
                Align Middle
              </Button>
              <Button className={styles.miniBtn} onClick={() => void alignActiveShape("Bottom")}>
                Align Bottom
              </Button>
            </div>
            <div className={styles.topActions}>
              <Button className={styles.miniBtn} onClick={() => void applyShapeZOrder("BringToFront")}>
                Bring To Front
              </Button>
              <Button className={styles.miniBtn} onClick={() => void applyShapeZOrder("BringForward")}>
                Bring Forward
              </Button>
              <Button className={styles.miniBtn} onClick={() => void applyShapeZOrder("SendBackward")}>
                Send Backward
              </Button>
              <Button className={styles.miniBtn} onClick={() => void applyShapeZOrder("SendToBack")}>
                Send To Back
              </Button>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Edit Shape</Text>
                <Select
                  className={styles.smallSelect}
                  value={shapeState.shapeType}
                  onChange={(_, data) =>
                    setShapeState((prev) => ({
                      ...prev,
                      shapeType: data.value as InsertableShapeType,
                    }))
                  }
                >
                  <option value="Rectangle">Rectangle</option>
                  <option value="RoundedRectangle">Rounded Rectangle</option>
                  <option value="Chevron">Chevron</option>
                  <option value="Hexagon">Hexagon</option>
                  <option value="Diamond">Diamond</option>
                  <option value="Oval">Oval</option>
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Current Z-Order Position</Text>
                <Input value={shapeState.zOrderPosition.toString()} readOnly />
                <Text className={styles.modalLabel}>Current Shape Type</Text>
                <Input value={shapeState.geometricShapeType || shapeState.shapeType} readOnly />
              </div>
            </div>
            <div className={styles.topActions}>
              <Button
                className={styles.miniBtn}
                onClick={() =>
                  void applyGeometricShapeType(shapeState.shapeType as Exclude<InsertableShapeType, "TextBox">)
                }
                disabled={shapeState.shapeType === "TextBox"}
              >
                Apply Edit Shape
              </Button>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Text Horizontal Alignment</Text>
                <Select
                  className={styles.smallSelect}
                  value={shapeState.textHorizontalAlignment}
                  onChange={(_, data) =>
                    setShapeState((prev) => ({
                      ...prev,
                      textHorizontalAlignment: data.value as ShapeEditorState["textHorizontalAlignment"],
                    }))
                  }
                >
                  <option value="Left">Left</option>
                  <option value="Center">Center</option>
                  <option value="Right">Right</option>
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Text Vertical Alignment</Text>
                <Select
                  className={styles.smallSelect}
                  value={shapeState.textVerticalAlignment}
                  onChange={(_, data) =>
                    setShapeState((prev) => ({
                      ...prev,
                      textVerticalAlignment: data.value as ShapeEditorState["textVerticalAlignment"],
                    }))
                  }
                >
                  <option value="Top">Top</option>
                  <option value="Middle">Middle</option>
                  <option value="Bottom">Bottom</option>
                </Select>
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <label className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={shapeState.bold}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, bold: event.target.checked }))
                  }
                />
                Bold
              </label>
              <label className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={shapeState.italic}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, italic: event.target.checked }))
                  }
                />
                Italic
              </label>
              <label className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={shapeState.lockAspectRatio}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, lockAspectRatio: event.target.checked }))
                  }
                />
                Lock Aspect Ratio
              </label>
            </div>
              </div>
            </details>
              </div>
            </details>
            <details className={styles.collapsibleSection}>
              <summary className={styles.collapsibleSummary}>Binding</summary>
              <div className={styles.collapsibleBody}>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Shape Value (Named Range Intellisense)</Text>
              <Input
                value={shapeState.valueBinding}
                onChange={(_, data) => {
                  setShapeSuggestionIndex(0);
                  setShapeState((prev) => ({ ...prev, valueBinding: data.value }));
                }}
              />
              {shapeValueSuggestions.length > 0 ? (
                <div className={styles.suggestionsBox}>
                  {shapeValueSuggestions.map((name, index) => (
                    <button
                      key={name}
                      type="button"
                      className={`${styles.suggestionBtn} ${shapeSuggestionIndex === index ? styles.suggestionBtnActive : ""}`}
                      onMouseEnter={() => setShapeSuggestionIndex(index)}
                      onClick={() => void applyNamedRangeToShapeText(name)}
                    >
                      nm {name}
                    </button>
                  ))}
                </div>
              ) : null}
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>
                Formula Editor (writes into anchor cell; font matches fill to appear hidden)
              </Text>
              <Input
                value={shapeState.formula}
                onChange={(_, data) => setShapeState((prev) => ({ ...prev, formula: data.value }))}
                placeholder='=TEXT(TODAY(),"yyyy-mm-dd")'
              />
            </div>
              </div>
            </details>
            <div className={styles.topActions}>
              <Button className={styles.miniBtn} onClick={() => void applyShapeFormatting()}>
                Update Shape Formatting
              </Button>
              <Button className={styles.miniBtn} onClick={() => void applyNamedRangeToShapeText()}>
                Apply Named Range Value
              </Button>
              <Button className={styles.miniBtn} onClick={() => void applyShapeAnchorFormula()}>
                Apply Formula to Anchor Cell
              </Button>
            </div>
          </>
        ) : (
          <Text className={styles.placeholder}>No active shape in this session. Use the + button to insert one on the selected cell.</Text>
        )}
      </div>
      {shapeAddOpen ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Insert Shape/Text Box</Text>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Type</Text>
              <Select
                className={styles.smallSelect}
                value={shapeState.shapeType}
                onChange={(_, data) =>
                  setShapeState((prev) => ({
                    ...prev,
                    shapeType: data.value as InsertableShapeType,
                  }))
                }
              >
                <option value="Rectangle">Rectangle</option>
                <option value="RoundedRectangle">Rounded Rectangle</option>
                <option value="Chevron">Chevron</option>
                <option value="Hexagon">Hexagon</option>
                <option value="Diamond">Diamond</option>
                <option value="Oval">Oval</option>
                <option value="TextBox">Text Box</option>
              </Select>
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Theme</Text>
              <Select
                className={styles.smallSelect}
                value={shapeState.selectedThemeId}
                onChange={(_, data) =>
                  setShapeState((prev) => ({ ...prev, selectedThemeId: data.value }))
                }
              >
                {shapeThemes.map((theme) => (
                  <option key={`insert-theme-${theme.id}`} value={theme.id}>
                    {theme.label}
                  </option>
                ))}
              </Select>
            </div>
            <Text className={styles.modalLabel}>Theme Colors</Text>
            <div className={styles.themePaletteHeader}>
              {selectedShapeTheme.themeColors.map((color) => (
                <button
                  key={`insert-theme-color-${color}`}
                  type="button"
                  className={styles.paletteSwatch}
                  style={{ backgroundColor: color }}
                  onClick={() => applyThemeColorToShapeField("fillColor", color)}
                />
              ))}
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Fill</Text>
                <div className={styles.swatchRow}>
                  <Button
                    className={`${styles.noFillBtn} ${
                      shapeState.fillColor === NO_FILL_COLOR_TOKEN ? styles.noFillBtnSelected : ""
                    }`}
                    onClick={applyShapeNoFill}
                  >
                    No Fill
                  </Button>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`insert-fill-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.fillColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      onClick={() => applyThemeColorToShapeField("fillColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.fillColor, "#ffffff")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, fillColor: event.target.value }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Outline</Text>
                <div className={styles.swatchRow}>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`insert-outline-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.outlineColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      onClick={() => applyThemeColorToShapeField("outlineColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.outlineColor, "#000000")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, outlineColor: event.target.value }))
                  }
                />
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Font Color</Text>
                <div className={styles.swatchRow}>
                  {themeColorSwatches.map((swatch) => (
                    <Button
                      key={`insert-font-${swatch}`}
                      className={`${styles.swatchBtn} ${
                        shapeState.fontColor === swatch ? styles.swatchBtnSelected : ""
                      }`}
                      style={{ backgroundColor: swatch }}
                      onClick={() => applyThemeColorToShapeField("fontColor", swatch)}
                    />
                  ))}
                </div>
                <input
                  className={styles.swatchInput}
                  type="color"
                  value={shapeColorInputValue(shapeState.fontColor, "#1f1f1f")}
                  onChange={(event) =>
                    setShapeState((prev) => ({ ...prev, fontColor: event.target.value }))
                  }
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Text</Text>
                <Input
                  value={shapeState.text}
                  onChange={(_, data) => setShapeState((prev) => ({ ...prev, text: data.value }))}
                />
              </div>
            </div>
            <div className={styles.modalActions}>
              <Button className={styles.miniBtn} onClick={() => setShapeAddOpen(false)}>Cancel</Button>
              <Button className={styles.miniBtn} disabled={isSubmitting} onClick={() => void insertShapeFromSelection()}>Insert on Active Cell</Button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );

  const renderFunctionsEditor = () => (
    <>
      <div className={styles.sectionCard}>
        <Text className={styles.sectionTitle}>{isFormulaPopout ? "Formula Editor" : "Formula Bar"}</Text>
        {!isFormulaPopout ? (
          <div className={styles.formulaPromptCard}>
            <Text className={styles.formulaPromptTitle}>Use Workbook Manager Formula Editor</Text>
            <Text className={styles.formulaPromptText}>
              If you are editing in Excel&apos;s formula bar, pull that formula into this editor for autocomplete,
              formatting, and testing.
            </Text>
            <Text className={styles.formulaPromptText}>
              Excel editor bridge: {autoCaptureExcelFormula ? "On" : "Off"} (captures committed formula edits and selection changes).
            </Text>
            <Text className={styles.formulaPromptMeta}>
              {activeFormulaPrompt
                ? `Active cell: ${activeFormulaPrompt.sheet}!${activeFormulaPrompt.address}`
                : "Active cell: not captured yet"}
            </Text>
            <div className={styles.topActions}>
              <Button className={styles.miniBtn} onClick={() => void pullActiveCellFormula()}>
                Open Active Cell Formula
              </Button>
              <Button className={styles.miniBtn} onClick={() => void pushFormulaToActiveCell()}>
                Apply Formula Back to Active Cell
              </Button>
              <Button className={styles.miniBtn} onClick={() => void syncFormulaFromActiveCell(true)}>
                Pull Latest From Excel Editor
              </Button>
            </div>
          </div>
        ) : null}
        <div className={styles.formulaMenuBar}>
          {!isFormulaPopout ? (
            <Select
              className={styles.smallSelect}
              value={formulaMode}
              onChange={(_, data) => setFormulaMode(data.value as "Formula" | "Function")}
            >
              <option value="Formula">Formula</option>
              <option value="Function">Function (LAMBDA)</option>
            </Select>
          ) : null}
          {!isFormulaPopout ? (
            <Button className={styles.miniBtn} onClick={() => setFormulaSettingsOpen((prev) => !prev)}>
              {formulaSettingsOpen ? "Hide Settings" : "Settings"}
            </Button>
          ) : null}
          <Button className={styles.miniBtn} onClick={() => void executeFormulaEditorCommand("beautifyFormula")}>
            Beautify Formula
          </Button>
          {isFormulaPopout ? (
            <>
              <Button className={styles.miniBtn} onClick={() => void syncFormulaFromActiveCell(true)}>
                Pull Latest
              </Button>
              <Button className={styles.miniBtn} onClick={() => void pushFormulaToActiveCell()}>
                Apply to Cell
              </Button>
            </>
          ) : null}
          {!isFormulaPopout ? (
            <Button
              className={styles.miniBtn}
              onClick={() => runAction("Open formula editor popout", openFormulaEditorPopout)}
            >
              Pop Out Formula Editor
            </Button>
          ) : null}
          <div className={styles.formulaMenuSpacer} />
          <Text className={styles.modalLabel}>
            {contextFunctionDefinition
              ? `${contextFunctionDefinition.name} arg ${(
                  (getCurrentFunctionContext?.argIndex ?? 0) + 1
                ).toString()}`
              : "Context: None"}
          </Text>
        </div>
        {!isFormulaPopout && formulaSettingsOpen ? (
          <div className={styles.settingsPanel}>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Editor Color Scheme</Text>
                <Select
                  className={styles.smallSelect}
                  value={formulaColorSchemeId}
                  onChange={(_, data) => setFormulaColorSchemeId(data.value)}
                >
                  <option value="preset:Advanced">Advanced (Preset)</option>
                  <option value="preset:Classic">Classic (Preset)</option>
                  <option value="preset:Monochrome">Monochrome (Preset)</option>
                  {customFormulaSchemes.map((item) => (
                    <option key={item.id} value={item.id}>
                      {`${item.name} (Custom)`}
                    </option>
                  ))}
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Preview</Text>
                <Input value={activeFormulaSchemeLabel} readOnly />
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Show Line Numbers</Text>
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={showFormulaLineNumbers}
                    onChange={(event) => setShowFormulaLineNumbers(event.target.checked)}
                  />
                  Show line numbers
                </label>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Excel Editor Bridge</Text>
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={autoCaptureExcelFormula}
                    onChange={(event) => setAutoCaptureExcelFormula(event.target.checked)}
                  />
                  Auto-capture active formula
                </label>
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={autoOpenFormulaTab}
                    onChange={(event) => setAutoOpenFormulaTab(event.target.checked)}
                    disabled={!autoCaptureExcelFormula}
                  />
                  Auto-open Formula tab
                </label>
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Custom Scheme Name</Text>
                <Input
                  value={customSchemeName}
                  onChange={(_, data) => setCustomSchemeName(data.value)}
                  placeholder="My Scheme"
                />
              </div>
            </div>
            <div className={styles.formulaPaletteEditor}>
              <div className={styles.formulaPaletteGrid}>
                {FORMULA_TOKEN_KEYS.map((tokenKey) => (
                  <div key={tokenKey} className={styles.formulaPaletteRow}>
                    <span>{FORMULA_TOKEN_LABELS[tokenKey]}</span>
                    <input
                      className={styles.colorInput}
                      type="color"
                      value={formulaPaletteDraft[tokenKey]}
                      onChange={(event) =>
                        setFormulaPaletteDraft((prev) => ({ ...prev, [tokenKey]: event.target.value }))
                      }
                    />
                  </div>
                ))}
              </div>
              <div className={styles.topActions}>
                <Button className={styles.miniBtn} onClick={createCustomFormulaScheme}>
                  Save As Custom
                </Button>
                <Button
                  className={styles.miniBtn}
                  onClick={updateCustomFormulaScheme}
                  disabled={!selectedCustomFormulaScheme}
                >
                  Update Custom
                </Button>
                <Button
                  className={styles.miniBtn}
                  onClick={deleteCustomFormulaScheme}
                  disabled={!selectedCustomFormulaScheme}
                >
                  Delete Custom
                </Button>
              </div>
            </div>
          </div>
        ) : null}

        {!isFormulaPopout && formulaMode === "Function" ? (
          <>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Function Name</Text>
                <Input
                  value={customFunctionName}
                  onChange={(_, data) => setCustomFunctionName(data.value)}
                />
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Parameters (comma-separated)</Text>
                <Input
                  value={customFunctionParams.join(",")}
                  onChange={(_, data) =>
                    setCustomFunctionParams(
                      data.value
                        .split(",")
                        .map((value) => value.trim())
                        .filter((value) => value.length > 0)
                    )
                  }
                />
              </div>
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>LAMBDA Body</Text>
              <Input
                value={customFunctionBody}
                onChange={(_, data) => setCustomFunctionBody(data.value)}
              />
            </div>
          </>
        ) : null}

        {!isFormulaPopout ? (
          <div className={styles.topActions}>
            <Button className={styles.miniBtn} onClick={() => void addSelectionToFormula()}>
              Insert Grid Selection
            </Button>
            <Button
              className={styles.miniBtn}
              onClick={() => {
                if (formulaMode === "Function") {
                  setFormulaText(generatedLambda.formula);
                  setTestFormulaCall(generatedLambda.invokeExample);
                }
              }}
              disabled={formulaMode !== "Function"}
            >
              Build LAMBDA
            </Button>
          </div>
        ) : null}
        <div className={styles.formulaEditorWrap}>
          {showFormulaLineNumbers ? (
            <div ref={formulaLineNumbersRef} className={styles.formulaLineNumbers} aria-hidden="true">
              <pre className={styles.formulaLineNumberText}>
                {Array.from({ length: formulaLineCount }, (_, i) => (i + 1).toString()).join("\n")}
              </pre>
            </div>
          ) : null}
          <pre
            ref={formulaHighlightRef}
            className={styles.formulaEditorHighlight}
            style={{ paddingLeft: showFormulaLineNumbers ? "34px" : "6px" }}
            aria-hidden="true"
          >
            {colorizedFormulaTokens.map((token, index) => {
              const className =
                token.kind === "function"
                  ? styles.formulaTokenFunction
                  : token.kind === "name"
                    ? styles.formulaTokenName
                    : token.kind === "table"
                      ? styles.formulaTokenTable
                      : token.kind === "local"
                        ? styles.formulaTokenLocal
                        : token.kind === "string"
                          ? styles.formulaTokenString
                          : token.kind === "number"
                            ? styles.formulaTokenNumber
                            : token.kind === "operator"
                              ? styles.formulaTokenOperator
                              : styles.formulaTokenDefault;
              return (
                <span
                  key={`${index.toString()}-${token.kind}`}
                  className={className}
                  style={{
                    color: selectedFormulaPalette[token.kind],
                    fontStyle: token.kind === "local" ? "italic" : "normal",
                  }}
                >
                  {token.text}
                </span>
              );
            })}
            {formulaText.length === 0 ? " " : ""}
          </pre>
          <textarea
            ref={formulaEditorRef}
            className={styles.formulaEditorArea}
            style={{ paddingLeft: showFormulaLineNumbers ? "34px" : "6px" }}
            value={formulaText}
            wrap="off"
            spellCheck={false}
            onChange={(event) => setFormulaText(event.target.value)}
            onScroll={(event) => {
              if (!formulaHighlightRef.current) {
                return;
              }
              formulaHighlightRef.current.scrollTop = event.currentTarget.scrollTop;
              formulaHighlightRef.current.scrollLeft = event.currentTarget.scrollLeft;
              if (formulaLineNumbersRef.current) {
                formulaLineNumbersRef.current.scrollTop = event.currentTarget.scrollTop;
              }
            }}
            onClick={(event) => setFormulaCursor(event.currentTarget.selectionStart ?? 0)}
            onKeyUp={(event) => setFormulaCursor(event.currentTarget.selectionStart ?? 0)}
            onSelect={(event) => setFormulaCursor(event.currentTarget.selectionStart ?? 0)}
            onKeyDown={handleFormulaEditorKeyDown}
            onContextMenu={(event) => {
              event.preventDefault();
              setFormulaContextMenu({ x: event.clientX, y: event.clientY });
            }}
            placeholder="Type a formula, e.g. =LET(name1, value1, calculation)"
          />
        </div>
        {shouldShowIntellisense ? (
          <div className={styles.suggestionsBox}>
            {functionSuggestions.map((item, index) => {
              const parts = getSuggestionTextParts(item.name, formulaTokenBeforeCursor);
              const kindLabel =
                item.kind === "Function" ? "fx" : item.kind === "Name" ? "nm" : item.kind === "Table" ? "tb" : "lv";
              return (
                <button
                  key={`${item.kind}:${item.name}`}
                  type="button"
                  className={`${styles.suggestionBtn} ${index === activeSuggestionIndex ? styles.suggestionBtnActive : ""}`}
                  onClick={() => {
                    setActiveSuggestionIndex(index);
                    insertFunctionTemplate(item);
                  }}
                  onMouseEnter={() => setActiveSuggestionIndex(index)}
                  title={item.signature ?? `${item.kind}: ${item.name}`}
                >
                  <span
                    className={`${styles.suggestionKind} ${index === activeSuggestionIndex ? styles.suggestionAccentOnActive : ""}`}
                  >
                    {kindLabel}
                  </span>
                  <span>
                    {parts.pre}
                    {parts.match ? (
                      <span
                        className={`${styles.suggestionMatch} ${index === activeSuggestionIndex ? styles.suggestionAccentOnActive : ""}`}
                      >
                        {parts.match}
                      </span>
                    ) : null}
                    {parts.post}
                  </span>
                </button>
              );
            })}
          </div>
        ) : null}
        {formulaReferenceHint ? (
          <Text className={styles.formulaHint}>
            {formulaReferenceHint.kind}: {formulaReferenceHint.name} {"->"} {formulaReferenceHint.detail}
          </Text>
        ) : null}
        {contextFunctionDefinition ? (
          <Text className={styles.modalLabel}>
            {contextFunctionDefinition.name}
            {contextFunctionDefinition.params.length > 0
              ? `(${contextFunctionDefinition.params
                  .map((param, idx) =>
                    idx === (getCurrentFunctionContext?.argIndex ?? -1) ? `[${param}]` : param
                  )
                  .join(", ")})`
              : "()"}
          </Text>
        ) : null}
        {formulaContextMenu ? (
          <div
            className={styles.formulaContextMenu}
            style={{ left: formulaContextMenu.x, top: formulaContextMenu.y }}
          >
            {[
              { key: "goToDefinition", label: "Go to Definition", shortcut: "Ctrl+F12" },
              { key: "goToSymbol", label: "Go to Symbol...", shortcut: "Ctrl+Shift+O" },
              { key: "peek", label: "Peek", shortcut: "" },
              { key: "sep-a", label: "-", shortcut: "" },
              { key: "beautifyFormula", label: "Beautify Formula", shortcut: "Ctrl+Shift+F" },
              { key: "renameSymbol", label: "Rename Symbol", shortcut: "F2" },
              { key: "changeAll", label: "Change All Occurrences", shortcut: "Ctrl+F2" },
              { key: "toggleSelectionRange", label: "Toggle selection to insert range", shortcut: "F4" },
              { key: "sep-b", label: "-", shortcut: "" },
              { key: "cut", label: "Cut", shortcut: "Ctrl+X" },
              { key: "copy", label: "Copy", shortcut: "Ctrl+C" },
              { key: "paste", label: "Paste", shortcut: "Ctrl+V" },
              { key: "sep-c", label: "-", shortcut: "" },
              { key: "commandPalette", label: "Command Palette", shortcut: "F1" },
            ].map((item) =>
              item.key.startsWith("sep-") ? (
                <hr key={item.key} className={styles.formulaContextSep} />
              ) : (
                <button
                  key={item.key}
                  type="button"
                  className={`${styles.formulaContextItem} ${item.key === "toggleSelectionRange" ? styles.formulaContextItemActive : ""}`}
                  onClick={() => {
                    void executeFormulaEditorCommand(item.key);
                    setFormulaContextMenu(null);
                  }}
                >
                  <span>{item.label}</span>
                  <span>{item.shortcut}</span>
                </button>
              )
            )}
          </div>
        ) : null}
      </div>

      {!isFormulaPopout ? (
        <div className={styles.sectionCard}>
          <Text className={styles.sectionTitle}>Function Test Runner</Text>
          <div className={styles.modalRow}>
            <Text className={styles.modalLabel}>Current Formula / Function Call</Text>
            <Input value={testFormulaCall} readOnly />
          </div>
          <div className={styles.topActions}>
            <Button className={styles.miniBtn} onClick={() => void runFormulaTest()}>
              Run Test
            </Button>
          </div>
          {formulaOutput ? (
            <div className={styles.gridOutputWrap}>
              <table className={styles.gridOutputTable}>
                <tbody>
                  {formulaOutput.values.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, colIndex) => (
                        <td key={`${rowIndex}-${colIndex}`} className={styles.gridOutputCell}>
                          {cell === null ? "" : String(cell)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
        </div>
      ) : null}
    </>
  );

  if (isFormulaPopout) {
    return (
      <div className={styles.root}>
        <div className={styles.frame}>
          <div className={styles.content}>
            {renderFunctionsEditor()}
            {status ? (
              <div className={styles.status}>
                <Text className={statusType === "success" ? styles.statusSuccess : styles.statusError}>
                  {status}
                </Text>
              </div>
            ) : null}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.frame}>
        {bridgeMode && onExitBridge ? (
          <div className={styles.bridgeBanner}>
            <Text className={styles.bridgeBannerText}>
              Running in legacy bridge mode for phased migration.
            </Text>
            <Button size="small" onClick={onExitBridge}>
              Back To Modern
            </Button>
          </div>
        ) : null}
        <div className={styles.primaryTabsWrap}>
          <div className={styles.primaryTabs}>
            {primaryTabs.map((tab) => (
              <button
                key={tab}
                type="button"
                className={`${styles.primaryTabBtn} ${primaryTab === tab ? styles.primaryTabActive : ""}`}
                onClick={() => setPrimaryTab(tab)}
              >
                {tab}
              </button>
            ))}
            <div className={styles.menuCluster}>
              <Button
                className={styles.iconBtn}
                icon={<TextAlignJustify20Regular />}
                title="Menu"
              />
            </div>
          </div>
        </div>

        <div className={styles.commandBar}>
          <div className={styles.commandGroup}>
            <Button
              className={styles.iconBtn}
              icon={<Save20Regular />}
              title="Insert text in A1"
              onClick={() => runAction("Insert text", () => insertText("Workbook Manager"))}
            />
            <Button
              className={styles.iconBtn}
              icon={<CloudArrowDown20Regular />}
              title="Refresh pivot tables"
              onClick={() => runAction("Refresh pivots", refreshPivotTables)}
            />
          </div>
          <div className={styles.commandGroup}>
            <Button
              className={`${styles.iconBtn} ${styles.iconBtnActive}`}
              icon={<Filter20Regular />}
              title="Toggle gridlines"
              onClick={() => runAction("Toggle gridlines", toggleGridlines)}
            />
            <Button
              className={styles.iconBtn}
              icon={<Grid20Regular />}
              title="Apply default table style"
              onClick={() =>
                runAction("Apply table style", () => applyTableStyle(DEFAULT_TABLE_STYLE))
              }
            />
            <Button
              className={styles.iconBtn}
              icon={<Add20Regular />}
              title="Insert rectangle and accent fill"
              onClick={() =>
                runAction("Insert rectangle", async () => {
                  await addRectangleShape();
                  await applyAccentFill();
                })
              }
            />
          </div>
        </div>

        <div className={styles.secondaryTabs}>
          {secondaryTabs.map((tab) => (
            <button
              key={tab.id}
              type="button"
              className={`${styles.secondaryTabBtn} ${secondaryTabId === tab.id ? styles.secondaryTabActive : ""}`}
              onClick={() => setSecondaryTabId(tab.id)}
            >
              {tab.label}
            </button>
          ))}
        </div>

        <div className={styles.content}>
          {primaryTab === "Names" && secondaryTabId === "ranges"
            ? renderRangesTable()
            : primaryTab === "Names" && secondaryTabId === "functions"
              ? renderFunctionsEditor()
            : primaryTab === "Names" && secondaryTabId === "tables"
              ? renderTablesTable()
            : primaryTab === "Format" && secondaryTabId === "styles"
              ? renderFormatStyles()
            : primaryTab === "Format" && secondaryTabId === "shapes"
              ? renderFormatShapes()
            : renderPlaceholder()}
          {status ? (
            <div className={styles.status}>
              <Text
                className={statusType === "success" ? styles.statusSuccess : styles.statusError}
              >
                {status}
              </Text>
            </div>
          ) : null}
        </div>
      </div>

      {editState.open && editState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Edit Named Range</Text>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Name</Text>
              <Input
                value={editState.name}
                onChange={(_, data) => setEditState((prev) => ({ ...prev, name: data.value }))}
              />
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Address</Text>
              <Input
                value={editState.address}
                onChange={(_, data) => setEditState((prev) => ({ ...prev, address: data.value }))}
              />
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Case Transform</Text>
              <Select
                className={styles.smallSelect}
                value={editState.caseTransform}
                onChange={(_, data) =>
                  setEditState((prev) => ({
                    ...prev,
                    caseTransform: data.value as CaseTransform,
                  }))
                }
              >
                <option value="none">None</option>
                <option value="camelCase">camelCase</option>
                <option value="snake_case">snake_case</option>
                <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
              </Select>
            </div>
            <div className={styles.modalActions}>
              <Button onClick={() => void useGridSelection()}>Use Selection</Button>
              <Button
                icon={<Dismiss20Regular />}
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
              <Button
                appearance="primary"
                onClick={() => void submitEdit()}
                disabled={isSubmitting}
              >
                Save
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {moveState.open && moveState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Move Named Range</Text>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Name</Text>
              <Text>{moveState.record.name}</Text>
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>New Address</Text>
              <Input
                value={moveState.address}
                onChange={(_, data) => setMoveState((prev) => ({ ...prev, address: data.value }))}
              />
            </div>
            <div className={styles.modalActions}>
              <Button onClick={() => void useGridSelectionForMove()}>Use Selection</Button>
              <Button
                icon={<Dismiss20Regular />}
                onClick={() =>
                  setMoveState({ open: false, record: null, address: "", fallbackSheet: "" })
                }
              >
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => void submitMove()}
                disabled={isSubmitting}
              >
                Move
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {deleteState.open && deleteState.record ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Delete Named Range</Text>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Name</Text>
              <Text>{deleteState.record.name}</Text>
            </div>
            <div className={styles.modalInlineRow}>
              <label className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={deleteState.deleteName}
                  onChange={(e) =>
                    setDeleteState((prev) => ({
                      ...prev,
                      deleteName: e.target.checked,
                    }))
                  }
                />
                Delete name definition
              </label>
              <label className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={deleteState.deleteValues}
                  onChange={(e) =>
                    setDeleteState((prev) => ({
                      ...prev,
                      deleteValues: e.target.checked,
                    }))
                  }
                />
                Delete values in range
              </label>
            </div>
            <div className={styles.modalActions}>
              <Button
                icon={<Dismiss20Regular />}
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
              <Button
                appearance="primary"
                onClick={() => void submitDelete()}
                disabled={isSubmitting}
              >
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {bulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Bulk Update Named Ranges</Text>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Operation</Text>
                <Select
                  className={styles.smallSelect}
                  value={bulkState.operation}
                  onChange={(_, data) =>
                    setBulkState((prev) => ({
                      ...prev,
                      operation: data.value as BulkState["operation"],
                    }))
                  }
                >
                  <option value="Add">Add</option>
                  <option value="Remove">Remove</option>
                  <option value="Replace">Replace</option>
                  <option value="Delete">Delete</option>
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Case Transform (Edit 1)</Text>
                <Select
                  className={styles.smallSelect}
                  value={bulkState.caseTransform}
                  onChange={(_, data) =>
                    setBulkState((prev) => ({
                      ...prev,
                      caseTransform: data.value as BulkState["caseTransform"],
                    }))
                  }
                >
                  <option value="none">None</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                </Select>
              </div>
            </div>

            {bulkState.operation !== "Delete" ? (
              <>
                <div className={styles.modalInlineRow}>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Position (Edit 1)</Text>
                    <Select
                      className={styles.smallSelect}
                      value={bulkState.position}
                      onChange={(_, data) =>
                        setBulkState((prev) => ({
                          ...prev,
                          position: data.value as BulkState["position"],
                        }))
                      }
                    >
                      <option value="Prefix">Prefix</option>
                      <option value="Suffix">Suffix</option>
                    </Select>
                  </div>
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Delimiter (Edit 1)</Text>
                    <Select
                      className={styles.smallSelect}
                      value={bulkState.delimiter}
                      onChange={(_, data) =>
                        setBulkState((prev) => ({
                          ...prev,
                          delimiter: data.value as BulkState["delimiter"],
                        }))
                      }
                    >
                      <option value="none">None</option>
                      <option value="underscore">Underscore (_)</option>
                      <option value="dot">Dot (.)</option>
                    </Select>
                  </div>
                </div>
                <div className={styles.modalRow}>
                  <Text className={styles.modalLabel}>
                    {bulkState.operation === "Replace" ? "Find Text" : "Text"}
                  </Text>
                  <Input
                    value={bulkState.textValue}
                    onChange={(_, data) =>
                      setBulkState((prev) => ({
                        ...prev,
                        textValue: data.value,
                      }))
                    }
                  />
                </div>
                {bulkState.operation === "Replace" ? (
                  <div className={styles.modalRow}>
                    <Text className={styles.modalLabel}>Replace With</Text>
                    <Input
                      value={bulkState.replaceWith}
                      onChange={(_, data) =>
                        setBulkState((prev) => ({
                          ...prev,
                          replaceWith: data.value,
                        }))
                      }
                    />
                  </div>
                ) : null}
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={bulkState.enableSecondEdit}
                    onChange={(e) =>
                      setBulkState((prev) => ({
                        ...prev,
                        enableSecondEdit: e.target.checked,
                      }))
                    }
                  />
                  Enable second edit step
                </label>
                {bulkState.enableSecondEdit ? (
                  <>
                    <div className={styles.modalInlineRow}>
                      <div className={styles.modalRow}>
                        <Text className={styles.modalLabel}>Operation (Edit 2)</Text>
                        <Select
                          className={styles.smallSelect}
                          value={bulkState.secondOperation}
                          onChange={(_, data) =>
                            setBulkState((prev) => ({
                              ...prev,
                              secondOperation: data.value as BulkState["secondOperation"],
                            }))
                          }
                        >
                          <option value="Add">Add</option>
                          <option value="Remove">Remove</option>
                          <option value="Replace">Replace</option>
                        </Select>
                      </div>
                      <div className={styles.modalRow}>
                        <Text className={styles.modalLabel}>Position (Edit 2)</Text>
                        <Select
                          className={styles.smallSelect}
                          value={bulkState.secondPosition}
                          onChange={(_, data) =>
                            setBulkState((prev) => ({
                              ...prev,
                              secondPosition: data.value as BulkState["secondPosition"],
                            }))
                          }
                        >
                          <option value="Prefix">Prefix</option>
                          <option value="Suffix">Suffix</option>
                        </Select>
                      </div>
                    </div>
                    <div className={styles.modalInlineRow}>
                      <div className={styles.modalRow}>
                        <Text className={styles.modalLabel}>Delimiter (Edit 2)</Text>
                        <Select
                          className={styles.smallSelect}
                          value={bulkState.secondDelimiter}
                          onChange={(_, data) =>
                            setBulkState((prev) => ({
                              ...prev,
                              secondDelimiter: data.value as BulkState["secondDelimiter"],
                            }))
                          }
                        >
                          <option value="none">None</option>
                          <option value="underscore">Underscore (_)</option>
                          <option value="dot">Dot (.)</option>
                        </Select>
                      </div>
                      <div className={styles.modalRow}>
                        <Text className={styles.modalLabel}>
                          {bulkState.secondOperation === "Replace"
                            ? "Find Text (Edit 2)"
                            : "Text (Edit 2)"}
                        </Text>
                        <Input
                          value={bulkState.secondTextValue}
                          onChange={(_, data) =>
                            setBulkState((prev) => ({
                              ...prev,
                              secondTextValue: data.value,
                            }))
                          }
                        />
                      </div>
                    </div>
                    {bulkState.secondOperation === "Replace" ? (
                      <div className={styles.modalRow}>
                        <Text className={styles.modalLabel}>Replace With (Edit 2)</Text>
                        <Input
                          value={bulkState.secondReplaceWith}
                          onChange={(_, data) =>
                            setBulkState((prev) => ({
                              ...prev,
                              secondReplaceWith: data.value,
                            }))
                          }
                        />
                      </div>
                    ) : null}
                    <div className={styles.modalRow}>
                      <Text className={styles.modalLabel}>Case Transform (Edit 2)</Text>
                      <Select
                        className={styles.smallSelect}
                        value={bulkState.secondCaseTransform}
                        onChange={(_, data) =>
                          setBulkState((prev) => ({
                            ...prev,
                            secondCaseTransform: data.value as BulkState["secondCaseTransform"],
                          }))
                        }
                      >
                        <option value="none">None</option>
                        <option value="camelCase">camelCase</option>
                        <option value="snake_case">snake_case</option>
                        <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                      </Select>
                    </div>
                  </>
                ) : null}
              </>
            ) : (
              <div className={styles.modalInlineRow}>
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={bulkState.deleteName}
                    onChange={(e) =>
                      setBulkState((prev) => ({
                        ...prev,
                        deleteName: e.target.checked,
                      }))
                    }
                  />
                  Delete names
                </label>
                <label className={styles.checkboxRow}>
                  <input
                    type="checkbox"
                    checked={bulkState.deleteValues}
                    onChange={(e) =>
                      setBulkState((prev) => ({
                        ...prev,
                        deleteValues: e.target.checked,
                      }))
                    }
                  />
                  Delete cell values
                </label>
              </div>
            )}

            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Preview ({bulkPreviewRows.length} selected)</Text>
              <div className={styles.previewBox}>
                <table className={styles.previewTable}>
                  <thead>
                    <tr>
                      <th className={styles.previewTh}>Current Name</th>
                      <th className={styles.previewTh}>New Name</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bulkPreviewRows.map((row) => (
                      <tr key={row.id}>
                        <td className={styles.previewTd}>{row.oldName}</td>
                        <td
                          className={`${styles.previewTd} ${row.oldName === row.newName ? styles.previewNoChange : ""}`}
                        >
                          {row.newName}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className={styles.modalActions}>
              <Button
                icon={<Dismiss20Regular />}
                onClick={resetBulkState}
              >
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => void submitBulkUpdate()}
                disabled={isSubmitting || bulkPreviewRows.length === 0}
              >
                Apply
              </Button>
            </div>
          </div>
        </div>
      ) : null}

      {tableBulkState.open ? (
        <div className={styles.modalBackdrop}>
          <div className={styles.modal}>
            <Text className={styles.modalTitle}>Bulk Edit Tables</Text>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Operation</Text>
                <Select
                  className={styles.smallSelect}
                  value={tableBulkState.operation}
                  onChange={(_, data) =>
                    setTableBulkState((prev) => ({
                      ...prev,
                      operation: data.value as TableBulkState["operation"],
                    }))
                  }
                >
                  <option value="Add">Add</option>
                  <option value="Remove">Remove</option>
                  <option value="Replace">Replace</option>
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Case Transform</Text>
                <Select
                  className={styles.smallSelect}
                  value={tableBulkState.caseTransform}
                  onChange={(_, data) =>
                    setTableBulkState((prev) => ({
                      ...prev,
                      caseTransform: data.value as CaseTransform,
                    }))
                  }
                >
                  <option value="none">None</option>
                  <option value="camelCase">camelCase</option>
                  <option value="snake_case">snake_case</option>
                  <option value="SCREAMING_SNAKE_CASE">SCREAMING_SNAKE_CASE</option>
                </Select>
              </div>
            </div>
            <div className={styles.modalInlineRow}>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Position</Text>
                <Select
                  className={styles.smallSelect}
                  value={tableBulkState.position}
                  onChange={(_, data) =>
                    setTableBulkState((prev) => ({
                      ...prev,
                      position: data.value as TableBulkState["position"],
                    }))
                  }
                >
                  <option value="Prefix">Prefix</option>
                  <option value="Suffix">Suffix</option>
                </Select>
              </div>
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Delimiter</Text>
                <Select
                  className={styles.smallSelect}
                  value={tableBulkState.delimiter}
                  onChange={(_, data) =>
                    setTableBulkState((prev) => ({
                      ...prev,
                      delimiter: data.value as TableBulkState["delimiter"],
                    }))
                  }
                >
                  <option value="none">None</option>
                  <option value="underscore">Underscore (_)</option>
                  <option value="dot">Dot (.)</option>
                </Select>
              </div>
            </div>
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>
                {tableBulkState.operation === "Replace" ? "Find Text" : "Text"}
              </Text>
              <Input
                value={tableBulkState.textValue}
                onChange={(_, data) =>
                  setTableBulkState((prev) => ({
                    ...prev,
                    textValue: data.value,
                  }))
                }
              />
            </div>
            {tableBulkState.operation === "Replace" ? (
              <div className={styles.modalRow}>
                <Text className={styles.modalLabel}>Replace With</Text>
                <Input
                  value={tableBulkState.replaceWith}
                  onChange={(_, data) =>
                    setTableBulkState((prev) => ({
                      ...prev,
                      replaceWith: data.value,
                    }))
                  }
                />
              </div>
            ) : null}
            <div className={styles.modalRow}>
              <Text className={styles.modalLabel}>Preview ({tableBulkPreview.length} selected)</Text>
              <div className={styles.previewBox}>
                <table className={styles.previewTable}>
                  <thead>
                    <tr>
                      <th className={styles.previewTh}>Current Name</th>
                      <th className={styles.previewTh}>New Name</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tableBulkPreview.map((row) => (
                      <tr key={row.id}>
                        <td className={styles.previewTd}>{row.oldName}</td>
                        <td className={styles.previewTd}>{row.newName}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className={styles.modalActions}>
              <Button icon={<Dismiss20Regular />} onClick={resetTableBulkState}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => void submitTableBulkUpdate()}
                disabled={isSubmitting || tableBulkPreview.length === 0}
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

export default LegacyApp;
