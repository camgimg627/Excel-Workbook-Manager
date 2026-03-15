/* global Excel Office console */

export interface NamedRangeRecord {
  id: string;
  kind: "NamedRange" | "Shape";
  name: string;
  address: string;
  formula: string;
  sheet: string;
  scope: string;
  scopeType: "Workbook" | "Worksheet";
  isRange: boolean;
  type:
    | "Single Cell"
    | "Single Column Array"
    | "Single Row Array"
    | "Multi-column Array"
    | "Multi-Row Array"
    | "Function"
    | "List"
    | "Formula"
    | "Shape";
}

export interface TableRecord {
  id: string;
  name: string;
  address: string;
  sheet: string;
  scope: string;
}

export interface TableColumnRecord {
  id: string;
  name: string;
  address: string;
}

export interface CreateNamedRangesFromTableRequest {
  sheetName: string;
  tableName: string;
  columns: string[];
  scopeType: "Workbook" | "Worksheet";
  conflictMode: "prefix" | "suffix" | "rename";
  conflictValue: string;
}

export interface FormulaEvaluationResult {
  address: string;
  values: (string | number | boolean | null)[][];
  valueTypes: string[][];
  hasError: boolean;
}

export interface ActiveCellFormulaState {
  sheet: string;
  address: string;
  formula: string;
  hasFormula: boolean;
}

export interface WorkbookQueryRecord {
  name: string;
  loadedTo: string;
  loadedToDataModel: boolean;
  refreshDate: string;
  rowsLoadedCount: number;
  error: string;
}

export interface QueryRefreshResult {
  refreshed: boolean;
  method: "DataConnections.refreshAll" | "Unsupported";
  warnings: string[];
}

export interface ModelBuilderParameterInput {
  label: string;
  desiredName: string;
  valueMode: "manual" | "formula";
  value: string;
  listValues: string;
}

export interface ModelBuilderApplyRequest {
  anchorSheet: string;
  anchorAddress: string;
  parameters: ModelBuilderParameterInput[];
}

export interface ModelBuilderApplyResult {
  anchorSheet: string;
  anchorAddress: string;
  created: Array<{
    label: string;
    finalName: string;
    valueAddress: string;
    validationListName?: string;
  }>;
}

export interface SheetFormatStyle {
  fillColor: string;
  fontColor: string;
  fontName: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  shrinkToFit: boolean;
  indentLevel: number;
  numberFormat: string;
}

export interface SheetFormatRecord {
  id: string;
  name: string;
  sourceSheet: string;
  sourceAddress: string;
  createdAt: string;
  updatedAt: string;
  style: SheetFormatStyle;
}

export interface UpdateSheetFormatRequest {
  name: string;
  style: SheetFormatStyle;
}

export type NavigationDestinationType = "Sheet" | "NamedRange" | "Table" | "Chart";

export interface NavigationDestinationOption {
  id: string;
  type: NavigationDestinationType;
  label: string;
  sheetName: string;
  address: string;
}

export interface NavigationButtonTemplate {
  id: string;
  label: string;
  destinationId: string;
  fillColor: string;
  outlineColor: string;
  fontColor: string;
}

export interface ApplyNavigationTemplateRequest {
  layout: "Vertical" | "Horizontal";
  panelFillColor: string;
  panelOutlineColor: string;
  panelWidth: number;
  panelPadding: number;
  buttonWidth: number;
  buttonHeight: number;
  buttonGap: number;
  originLeft: number;
  originTop: number;
  buttons: NavigationButtonTemplate[];
  applyToAllSheets: boolean;
}

export type ShapeBuilderLinkType = "None" | "Internal" | "External";
export type ShapeBuilderEffect = "None" | "Shadow";
export type ShapeBuilderTextHorizontalAlignment = "Left" | "Center" | "Right";
export type ShapeBuilderTextVerticalAlignment = "Top" | "Middle" | "Bottom";

export interface ShapeBuilderShapeRecord {
  id: string;
  sheetName: string;
  shapeName: string;
  shapeType: InsertableShapeType;
  anchorAddress: string;
  text: string;
  iconKey: string;
  fillColor: string;
  outlineColor: string;
  outlineWidth: number;
  width: number;
  height: number;
  left: number;
  top: number;
  zOrderPosition: number;
  fontColor: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  textHorizontalAlignment: ShapeBuilderTextHorizontalAlignment;
  textVerticalAlignment: ShapeBuilderTextVerticalAlignment;
  linkType: ShapeBuilderLinkType;
  internalDestinationId: string;
  externalUrl: string;
  effect: ShapeBuilderEffect;
}

export interface CreateShapeBuilderShapeRequest {
  shapeType: InsertableShapeType;
  text: string;
  iconKey: string;
  fillColor: string;
  outlineColor: string;
  outlineWidth: number;
  width: number;
  height: number;
  left?: number;
  top?: number;
  anchorAddress?: string;
  fontColor: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  textHorizontalAlignment: ShapeBuilderTextHorizontalAlignment;
  textVerticalAlignment: ShapeBuilderTextVerticalAlignment;
  linkType: ShapeBuilderLinkType;
  internalDestinationId: string;
  externalUrl: string;
  effect: ShapeBuilderEffect;
}

export type UpdateShapeBuilderShapeRequest = Partial<CreateShapeBuilderShapeRequest>;

export interface ShapeBuilderTextSelectionFormatRequest {
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

export type CellStylePreset = "Input Cell" | "Parameter Cell" | "Header" | "Subheader";
export type WorkbookWindowArrangeLayout = "Tiled" | "Horizontal" | "Vertical" | "Cascade";
export type InsertableShapeType =
  | "Rectangle"
  | "RoundedRectangle"
  | "Chevron"
  | "Hexagon"
  | "Diamond"
  | "Oval"
  | "TextBox";

export interface ShapeInsertOptions {
  shapeType: InsertableShapeType;
  fillColor: string;
  outlineColor: string;
  fontColor: string;
  text: string;
}

export interface InsertedShapeRecord {
  name: string;
  shapeType: InsertableShapeType;
  sheet: string;
  anchorAddress: string;
  width: number;
  height: number;
}

export interface ShapeFormatOptions {
  fillColor: string;
  outlineColor: string;
  fontColor: string;
  text: string;
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
}

export interface ShapePositionOptions {
  left?: number;
  top?: number;
}

export interface ShapeEditorRecord {
  name: string;
  sheet: string;
  shapeType: InsertableShapeType;
  geometricShapeType: string;
  anchorAddress: string;
  left: number;
  top: number;
  zOrderPosition: number;
  fillColor: string;
  outlineColor: string;
  fontColor: string;
  text: string;
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
}

export const NO_FILL_COLOR_TOKEN = "__NO_FILL__";
const SHAPE_ANCHOR_PREFIX = "WBM_ANCHOR=";
const SHEET_FORMAT_STORE_SHEET = "__WBM_SHEET_FORMATS";
const SHEET_FORMAT_TEMPLATE_PREFIX = "__WBM_FMT_";
const NAV_SHAPE_PREFIX = "WBM_NAV_SHAPE_";
const NAV_SHADOW_PREFIX = "WBM_NAV_SHADOW_";
const NAV_PANEL_PREFIX = "WBM_NAV_PANEL_";
const NAV_TARGET_PREFIX = "WBM_NAV_TARGET=";
const SHAPE_BUILDER_META_PREFIX = "WBM_SHAPE_BUILDER_META=";
const WORKBOOK_THEME_SWATCH_FALLBACK = [
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
const WORKBOOK_THEME_STYLE_SEQUENCE = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Accent1_20",
  "Accent2_20",
  "Accent3_20",
  "Accent4_20",
  "Accent5_20",
  "Accent6_20",
  "Accent1_40",
  "Accent2_40",
  "Accent3_40",
  "Accent4_40",
  "Accent5_40",
  "Accent6_40",
  "Accent1_60",
  "Accent2_60",
  "Accent3_60",
  "Accent4_60",
  "Accent5_60",
  "Accent6_60",
];
const FORMULA_DIALOG_RPC_CHANNEL = "wbm-formula-dialog-rpc";
let formulaEditorDialog: Office.Dialog | null = null;
let formatEditorDialog: Office.Dialog | null = null;

interface FormulaDialogEvalRequestMessage {
  channel: typeof FORMULA_DIALOG_RPC_CHANNEL;
  type: "eval-request";
  requestId: string;
  formula: string;
}

interface FormulaDialogEvalResponseMessage {
  channel: typeof FORMULA_DIALOG_RPC_CHANNEL;
  type: "eval-response";
  requestId: string;
  ok: boolean;
  result?: FormulaEvaluationResult;
  error?: string;
}

interface FormulaDialogReadyMessage {
  channel: typeof FORMULA_DIALOG_RPC_CHANNEL;
  type: "ready";
}

interface FormulaDialogOpenFormulaMessage {
  channel: typeof FORMULA_DIALOG_RPC_CHANNEL;
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

async function runFormattingCommand(command: (context: Excel.RequestContext) => Promise<void>) {
  await Excel.run(async (context) => {
    await command(context);
    await context.sync();
  });
}

function splitQualifiedAddress(formula: string): { sheet: string; address: string } {
  const cleaned = normalizeDisplayAddress(formula);
  const bang = cleaned.indexOf("!");
  if (bang < 0) {
    return { sheet: "", address: cleaned };
  }

  const sheet = cleaned.slice(0, bang).replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
  const address = cleaned.slice(bang + 1);
  return { sheet, address };
}

function quoteSheetName(name: string): string {
  const escaped = name.replace(/'/g, "''");
  return /[\s!']/g.test(name) ? `'${escaped}'` : escaped;
}

function normalizeReference(addressInput: string, fallbackSheet: string): string {
  const cleaned = normalizeDisplayAddress(addressInput);
  if (cleaned.includes("!")) {
    return cleaned;
  }
  const normalizedFallbackSheet = fallbackSheet.trim();
  if (!normalizedFallbackSheet) {
    return cleaned;
  }
  return `${quoteSheetName(normalizedFallbackSheet)}!${cleaned}`;
}

function classifyRangeShape(rowCount: number, columnCount: number): NamedRangeRecord["type"] {
  if (rowCount === 1 && columnCount === 1) {
    return "Single Cell";
  }
  if (rowCount > 1 && columnCount === 1) {
    return "Single Column Array";
  }
  if (rowCount === 1 && columnCount > 1) {
    return "Single Row Array";
  }
  return rowCount >= columnCount ? "Multi-Row Array" : "Multi-column Array";
}

function isLambdaFormula(formula: string): boolean {
  return /^LAMBDA\s*\(/i.test(formula.trim());
}

function isArrayConstantFormula(formula: string): boolean {
  const trimmed = formula.trim();
  return trimmed.startsWith("{") && trimmed.endsWith("}");
}

function classifyNamedItemType(formula: string, isRange: boolean): NamedRangeRecord["type"] {
  if (isRange) {
    return "Single Cell";
  }
  if (isLambdaFormula(formula)) {
    return "Function";
  }
  if (isArrayConstantFormula(formula)) {
    return "List";
  }
  return "Formula";
}

function normalizeDisplayAddress(formula: string): string {
  const trimmed = formula.replace(/^=/, "").trim();
  if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replace(/""/g, '"').trim();
  }
  return trimmed;
}

function normalizeFormulaExpression(formulaInput: string): string {
  const trimmed = formulaInput.trim();
  if (!trimmed) {
    return "";
  }
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

function canUseExcelApi(minVersion: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", minVersion);
  } catch {
    return false;
  }
}

function buildArrayFormulaSeparatorCandidates(formulaInput: string): string[] {
  const normalized = normalizeFormulaExpression(formulaInput);
  const body = normalized.replace(/^=/, "").trim();
  if (!/^\{[\s\S]*\}$/.test(body)) {
    return [normalized];
  }

  const swapSeparators = (input: string, from: ";" | ",", to: ";" | ","): string => {
    let output = "";
    let inString = false;
    for (let index = 0; index < input.length; index += 1) {
      const ch = input[index];
      if (ch === '"') {
        output += ch;
        if (inString && input[index + 1] === '"') {
          output += '"';
          index += 1;
          continue;
        }
        inString = !inString;
        continue;
      }
      if (!inString && ch === from) {
        output += to;
      } else {
        output += ch;
      }
    }
    return output;
  };

  const candidates = new Set<string>([normalized]);
  if (body.includes(";")) {
    candidates.add(`=${swapSeparators(body, ";", ",")}`);
  }
  if (body.includes(",")) {
    candidates.add(`=${swapSeparators(body, ",", ";")}`);
  }
  return Array.from(candidates);
}

function parseInlineListValues(input: string): string[] {
  return input
    .split(/[\r\n,;]+/)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
}

function quoteArrayLiteralValue(value: string): string {
  if (/^-?\d+(\.\d+)?$/.test(value)) {
    return value;
  }
  if (/^(TRUE|FALSE)$/i.test(value)) {
    return value.toUpperCase();
  }
  return `"${value.replace(/"/g, '""')}"`;
}

function buildVerticalArrayFormula(values: string[], separator: ";" | ","): string {
  return `={${values.map((value) => quoteArrayLiteralValue(value)).join(separator)}}`;
}

function toModelSafeName(value: string): string {
  const normalized = value
    .replace(/[^A-Za-z0-9_.\\]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+/, "")
    .replace(/_+$/, "");
  let candidate = normalized || "Parameter";
  if (!/^[A-Za-z_\\]/.test(candidate)) {
    candidate = `_${candidate}`;
  }
  if (/^[A-Za-z]{1,3}\d+$/i.test(candidate)) {
    candidate = `_${candidate}`;
  }
  return candidate;
}

function toUniqueName(candidate: string, taken: Set<string>): string {
  const base = candidate.trim() || "Parameter";
  let next = base;
  let index = 2;
  while (taken.has(next.toUpperCase())) {
    next = `${base}_${index.toString()}`;
    index += 1;
  }
  taken.add(next.toUpperCase());
  return next;
}

function columnIndexToName(index: number): string {
  let remaining = index + 1;
  let output = "";
  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    output = String.fromCharCode(65 + modulo) + output;
    remaining = Math.floor((remaining - modulo) / 26);
  }
  return output;
}

function toAbsoluteCellAddress(rowIndex: number, columnIndex: number): string {
  return `$${columnIndexToName(columnIndex)}$${(rowIndex + 1).toString()}`;
}

function toQualifiedAddress(sheet: string, address: string): string {
  if (!sheet || !address) {
    return address;
  }
  return `${quoteSheetName(sheet)}!${address}`;
}

function isCellAddressExpression(address: string): boolean {
  const normalized = address.replace(/\$/g, "").trim();
  return (
    /^[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$/i.test(normalized) ||
    /^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/i.test(normalized) ||
    /^\d+:\d+$/.test(normalized)
  );
}

function getSheetAndAddress(
  reference: string,
  fallbackSheet: string
): { sheet: string; address: string } {
  const normalized = normalizeReference(reference, fallbackSheet);
  const parsed = splitQualifiedAddress(normalized);
  return {
    sheet: parsed.sheet || fallbackSheet,
    address: parsed.address,
  };
}

function normalizeMaybeAddress(input: string): string {
  const trimmed = input.replace(/^=/, "").trim();
  if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replace(/""/g, '"').trim();
  }
  return trimmed;
}

function parseAddressLikeText(
  text: string,
  fallbackSheet: string
): { sheet: string; address: string } | null {
  const normalized = normalizeMaybeAddress(text);
  const isAddressOnly =
    /^[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$/i.test(normalized) ||
    /^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/i.test(normalized) ||
    /^\d+:\d+$/.test(normalized);
  const isQualified = normalized.includes("!");
  if (!isAddressOnly && !isQualified) {
    return null;
  }
  try {
    const parsed = getSheetAndAddress(normalized, fallbackSheet);
    return parsed.sheet && parsed.address ? parsed : null;
  } catch {
    return null;
  }
}

function extractShapeAnchorAddress(altTextDescription: string | null | undefined): string {
  const text = (altTextDescription ?? "").trim();
  if (!text.startsWith(SHAPE_ANCHOR_PREFIX)) {
    return "";
  }
  return text.slice(SHAPE_ANCHOR_PREFIX.length).trim();
}

function inferInsertableShapeType(
  geometricType?: string | null,
  textAsGeometry?: string | null
): InsertableShapeType {
  const normalized = (geometricType ?? textAsGeometry ?? "").toLowerCase();
  if (normalized.includes("textbox") || normalized.includes("text box")) {
    return "TextBox";
  }
  if (normalized.includes("round")) {
    return "RoundedRectangle";
  }
  if (normalized.includes("chevron")) {
    return "Chevron";
  }
  if (normalized.includes("hexagon")) {
    return "Hexagon";
  }
  if (normalized.includes("diamond")) {
    return "Diamond";
  }
  if (normalized.includes("oval") || normalized.includes("ellipse")) {
    return "Oval";
  }
  if (normalized.length === 0) {
    return "TextBox";
  }
  return "Rectangle";
}

function toGeometricShapeType(
  shapeType: Exclude<InsertableShapeType, "TextBox">
): Excel.GeometricShapeType {
  const geometricMap: Record<Exclude<InsertableShapeType, "TextBox">, Excel.GeometricShapeType> = {
    Rectangle: Excel.GeometricShapeType.rectangle,
    RoundedRectangle: "Round2SameRectangle" as unknown as Excel.GeometricShapeType,
    Chevron: "Chevron" as unknown as Excel.GeometricShapeType,
    Hexagon: "Hexagon" as unknown as Excel.GeometricShapeType,
    Diamond: "Diamond" as unknown as Excel.GeometricShapeType,
    Oval: "Oval" as unknown as Excel.GeometricShapeType,
  };
  return geometricMap[shapeType] ?? Excel.GeometricShapeType.rectangle;
}

export async function insertText(text: string) {
  try {
    await runFormattingCommand(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
    });
  } catch (error) {
    console.log("Error: " + error);
    throw error;
  }
}

export async function toggleGridlines() {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("showGridlines");
    await context.sync();
    sheet.showGridlines = !sheet.showGridlines;
  });
}

function ensureWorkbookWindowApiSupport(): void {
  if (typeof Office === "undefined") {
    throw new Error("Office.js is not available.");
  }
  const supported = Office.context?.requirements?.isSetSupported("ExcelApiDesktop", "1.1");
  if (!supported) {
    throw new Error("Workbook window controls are available in Excel desktop only.");
  }
}

function normalizeWindowGroupKey(windowName: string): string {
  const trimmed = asString(windowName).trim();
  const separatorIndex = trimmed.lastIndexOf(":");
  if (separatorIndex <= 0) {
    return trimmed.toLowerCase();
  }
  return trimmed.slice(0, separatorIndex).toLowerCase();
}

export async function openNewWorkbookWindow(): Promise<void> {
  ensureWorkbookWindowApiSupport();

  await Excel.run(async (context) => {
    context.workbook.application.activeWindow.newWindow();
    await context.sync();
  });
}

export async function arrangeWorkbookWindows(layout: WorkbookWindowArrangeLayout): Promise<void> {
  ensureWorkbookWindowApiSupport();

  await Excel.run(async (context) => {
    const application = context.workbook.application;
    const windows = application.windows;
    const activeWindow = application.activeWindow;

    windows.load("items/name,items/type,items/windowState");
    activeWindow.load("name,left,top,usableWidth,usableHeight");
    await context.sync();

    const activeGroupKey = normalizeWindowGroupKey(activeWindow.name);
    const workbookWindows = windows.items.filter(
      (windowItem) => asString(windowItem.type) === Excel.WindowType.workbook
    );
    const sameWorkbookWindows = workbookWindows.filter(
      (windowItem) => normalizeWindowGroupKey(windowItem.name) === activeGroupKey
    );
    const targetWindows = sameWorkbookWindows.length >= 2 ? sameWorkbookWindows : workbookWindows;

    if (targetWindows.length < 2) {
      throw new Error("Open at least two workbook windows before arranging.");
    }

    const originLeft = Math.max(0, asNumber(activeWindow.left, 0));
    const originTop = Math.max(0, asNumber(activeWindow.top, 0));
    const availableWidth = Math.max(460, asNumber(activeWindow.usableWidth, 1280));
    const availableHeight = Math.max(320, asNumber(activeWindow.usableHeight, 760));

    targetWindows.forEach((windowItem) => {
      windowItem.windowState = Excel.WindowState.normal;
    });

    if (layout === "Vertical") {
      const columnWidth = Math.max(320, Math.floor(availableWidth / targetWindows.length));
      targetWindows.forEach((windowItem, index) => {
        windowItem.left = originLeft + index * columnWidth;
        windowItem.top = originTop;
        windowItem.width = columnWidth;
        windowItem.height = availableHeight;
      });
      await context.sync();
      return;
    }

    if (layout === "Horizontal") {
      const rowHeight = Math.max(240, Math.floor(availableHeight / targetWindows.length));
      targetWindows.forEach((windowItem, index) => {
        windowItem.left = originLeft;
        windowItem.top = originTop + index * rowHeight;
        windowItem.width = availableWidth;
        windowItem.height = rowHeight;
      });
      await context.sync();
      return;
    }

    if (layout === "Cascade") {
      const offset = Math.max(
        24,
        Math.min(48, Math.floor(Math.min(availableWidth, availableHeight) / 16))
      );
      const cascadeWidth = Math.max(
        440,
        availableWidth - offset * Math.max(targetWindows.length - 1, 0)
      );
      const cascadeHeight = Math.max(
        300,
        availableHeight - offset * Math.max(targetWindows.length - 1, 0)
      );
      targetWindows.forEach((windowItem, index) => {
        windowItem.left = originLeft + index * offset;
        windowItem.top = originTop + index * offset;
        windowItem.width = cascadeWidth;
        windowItem.height = cascadeHeight;
      });
      await context.sync();
      return;
    }

    // Tiled
    const count = targetWindows.length;
    const columns = Math.ceil(Math.sqrt(count));
    const rows = Math.ceil(count / columns);
    const cellWidth = Math.max(320, Math.floor(availableWidth / columns));
    const cellHeight = Math.max(220, Math.floor(availableHeight / rows));
    targetWindows.forEach((windowItem, index) => {
      const row = Math.floor(index / columns);
      const column = index % columns;
      windowItem.left = originLeft + column * cellWidth;
      windowItem.top = originTop + row * cellHeight;
      windowItem.width = cellWidth;
      windowItem.height = cellHeight;
    });
    await context.sync();
  });
}

export async function freezeTopRow() {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.freezePanes.freezeRows(1);
  });
}

export async function freezeFirstColumn() {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.freezePanes.freezeColumns(1);
  });
}

export async function unfreezePanes() {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.freezePanes.unfreeze();
  });
}

async function setActiveSheetVisibility(visibility: Excel.SheetVisibility): Promise<void> {
  await runFormattingCommand(async (context) => {
    const sheets = context.workbook.worksheets;
    const activeSheet = sheets.getActiveWorksheet();
    sheets.load("items/name,items/visibility");
    activeSheet.load("name");
    await context.sync();

    const visibleSheets = sheets.items.filter(
      (sheet) =>
        asString(sheet.visibility).toLowerCase() === Excel.SheetVisibility.visible.toLowerCase()
    );
    const isActiveVisible = visibleSheets.some((sheet) => sheet.name === activeSheet.name);
    if (isActiveVisible && visibleSheets.length <= 1) {
      throw new Error("Cannot hide the only visible worksheet.");
    }

    activeSheet.visibility = visibility;
  });
}

export async function hideActiveSheet() {
  await setActiveSheetVisibility(Excel.SheetVisibility.hidden);
}

export async function hideActiveSheetVeryHidden() {
  await setActiveSheetVisibility(Excel.SheetVisibility.veryHidden);
}

async function applyDimensionToSelectionOrSheet(
  apply: (target: Excel.Range, value: number) => void,
  value: number
): Promise<void> {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selection = context.workbook.getSelectedRange();
    selection.load("rowCount,columnCount");
    await context.sync();

    const target =
      selection.rowCount <= 1 && selection.columnCount <= 1 ? sheet.getRange() : selection;
    apply(target, value);
  });
}

export async function setSelectionColumnWidth(width: number) {
  const next = Math.max(0, asNumber(width, 0));
  await applyDimensionToSelectionOrSheet((target, value) => {
    target.format.columnWidth = value;
  }, next);
}

export async function setSelectionRowHeight(height: number) {
  const next = Math.max(0, asNumber(height, 0));
  await applyDimensionToSelectionOrSheet((target, value) => {
    target.format.rowHeight = value;
  }, next);
}

export async function toggleBold() {
  await runFormattingCommand(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.font.load("bold");
    await context.sync();
    range.format.font.bold = !range.format.font.bold;
  });
}

export async function applyAccentFill() {
  await runFormattingCommand(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = "#DDEBF7";
  });
}

export async function applyCellStylePreset(preset: CellStylePreset) {
  await runFormattingCommand(async (context) => {
    const range = context.workbook.getSelectedRange();

    // Clear baseline style first for consistent results.
    range.format.font.bold = false;
    range.format.font.italic = false;
    range.format.font.color = "#1f1f1f";
    range.format.fill.color = "#ffffff";
    range.format.horizontalAlignment = "Left";
    range.format.verticalAlignment = "Center";
    range.format.wrapText = false;

    if (preset === "Input Cell") {
      range.format.fill.color = "#FFF2CC";
      range.format.font.color = "#1f1f1f";
      range.format.horizontalAlignment = "Left";
    } else if (preset === "Parameter Cell") {
      range.format.fill.color = "#E2F0D9";
      range.format.font.bold = true;
      range.format.font.color = "#1F4E78";
      range.format.horizontalAlignment = "Left";
    } else if (preset === "Header") {
      range.format.fill.color = "#1F4E78";
      range.format.font.bold = true;
      range.format.font.color = "#FFFFFF";
      range.format.horizontalAlignment = "Center";
      range.format.wrapText = true;
    } else if (preset === "Subheader") {
      range.format.fill.color = "#D9E1F2";
      range.format.font.bold = true;
      range.format.font.color = "#1F4E78";
      range.format.horizontalAlignment = "Left";
    }

    range.format.autofitColumns();
  });
}

export async function addRectangleShape() {
  await runFormattingCommand(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    shape.name = `FormattingShape_${Date.now()}`;
    shape.left = 40;
    shape.top = 40;
    shape.height = 70;
    shape.width = 180;
    shape.fill.setSolidColor("#E3F2FD");
    shape.textFrame.textRange.text = "Formatting tab";
  });
}

export async function refreshPivotTables() {
  await runFormattingCommand(async (context) => {
    const pivotTables = context.workbook.worksheets.getActiveWorksheet().pivotTables;
    pivotTables.load("items/name");
    await context.sync();

    if (pivotTables.items.length === 0) {
      throw new Error("No pivot tables found on the active worksheet.");
    }

    pivotTables.items.forEach((pivotTable) => pivotTable.refresh());
  });
}

export async function applyTableStyle(styleName: string) {
  await runFormattingCommand(async (context) => {
    const tables = context.workbook.worksheets.getActiveWorksheet().tables;
    tables.load("items/name");
    await context.sync();

    if (tables.items.length === 0) {
      throw new Error("No tables found on the active worksheet.");
    }

    tables.items[0].style = styleName;
  });
}

const DEFAULT_SHEET_FORMAT_STYLE: SheetFormatStyle = {
  fillColor: "#FFFFFF",
  fontColor: "#1F1F1F",
  fontName: "Calibri",
  fontSize: 11,
  bold: false,
  italic: false,
  underline: false,
  horizontalAlignment: "General",
  verticalAlignment: "Center",
  wrapText: false,
  shrinkToFit: false,
  indentLevel: 0,
  numberFormat: "General",
};

function asString(value: unknown, fallback = ""): string {
  if (typeof value === "string") {
    return value;
  }
  if (value === null || value === undefined) {
    return fallback;
  }
  return String(value);
}

function asNumber(value: unknown, fallback: number): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function asBoolean(value: unknown, fallback = false): boolean {
  if (typeof value === "boolean") {
    return value;
  }
  if (value === null || value === undefined) {
    return fallback;
  }
  const normalized = String(value).trim().toLowerCase();
  if (normalized === "true" || normalized === "1") {
    return true;
  }
  if (normalized === "false" || normalized === "0") {
    return false;
  }
  return fallback;
}

function loadSheetFormatTemplateCell(templateCell: Excel.Range): void {
  templateCell.load("numberFormat");
  templateCell.format.load(
    "fill/color,horizontalAlignment,verticalAlignment,wrapText,shrinkToFit,indentLevel"
  );
  templateCell.format.font.load("name,size,bold,italic,underline,color");
}

function buildSheetFormatStyle(templateCell: Excel.Range): SheetFormatStyle {
  const underlineRaw = asString(templateCell.format.font.underline, "None").toLowerCase();
  const isUnderlined =
    underlineRaw.length > 0 &&
    underlineRaw !== "none" &&
    underlineRaw !== "false" &&
    underlineRaw !== "0";
  const formatValue = templateCell.numberFormat?.[0]?.[0];

  return {
    fillColor: asString(templateCell.format.fill.color, DEFAULT_SHEET_FORMAT_STYLE.fillColor),
    fontColor: asString(templateCell.format.font.color, DEFAULT_SHEET_FORMAT_STYLE.fontColor),
    fontName: asString(templateCell.format.font.name, DEFAULT_SHEET_FORMAT_STYLE.fontName),
    fontSize: asNumber(templateCell.format.font.size, DEFAULT_SHEET_FORMAT_STYLE.fontSize),
    bold: asBoolean(templateCell.format.font.bold, DEFAULT_SHEET_FORMAT_STYLE.bold),
    italic: asBoolean(templateCell.format.font.italic, DEFAULT_SHEET_FORMAT_STYLE.italic),
    underline: isUnderlined,
    horizontalAlignment: asString(
      templateCell.format.horizontalAlignment,
      DEFAULT_SHEET_FORMAT_STYLE.horizontalAlignment
    ),
    verticalAlignment: asString(
      templateCell.format.verticalAlignment,
      DEFAULT_SHEET_FORMAT_STYLE.verticalAlignment
    ),
    wrapText: asBoolean(templateCell.format.wrapText, DEFAULT_SHEET_FORMAT_STYLE.wrapText),
    shrinkToFit: asBoolean(templateCell.format.shrinkToFit, DEFAULT_SHEET_FORMAT_STYLE.shrinkToFit),
    indentLevel: asNumber(templateCell.format.indentLevel, DEFAULT_SHEET_FORMAT_STYLE.indentLevel),
    numberFormat: asString(formatValue, DEFAULT_SHEET_FORMAT_STYLE.numberFormat),
  };
}

function applySheetFormatStyle(templateCell: Excel.Range, style: SheetFormatStyle): void {
  templateCell.format.fill.color = style.fillColor || DEFAULT_SHEET_FORMAT_STYLE.fillColor;
  templateCell.format.font.color = style.fontColor || DEFAULT_SHEET_FORMAT_STYLE.fontColor;
  templateCell.format.font.name = style.fontName || DEFAULT_SHEET_FORMAT_STYLE.fontName;
  templateCell.format.font.size = Math.max(
    6,
    asNumber(style.fontSize, DEFAULT_SHEET_FORMAT_STYLE.fontSize)
  );
  templateCell.format.font.bold = asBoolean(style.bold, false);
  templateCell.format.font.italic = asBoolean(style.italic, false);
  templateCell.format.font.underline = (style.underline
    ? "Single"
    : "None") as unknown as Excel.RangeUnderlineStyle;
  templateCell.format.horizontalAlignment = (style.horizontalAlignment ||
    DEFAULT_SHEET_FORMAT_STYLE.horizontalAlignment) as unknown as Excel.HorizontalAlignment;
  templateCell.format.verticalAlignment = (style.verticalAlignment ||
    DEFAULT_SHEET_FORMAT_STYLE.verticalAlignment) as unknown as Excel.VerticalAlignment;
  templateCell.format.wrapText = asBoolean(style.wrapText, false);
  templateCell.format.shrinkToFit = asBoolean(style.shrinkToFit, false);
  templateCell.format.indentLevel = Math.max(0, Math.round(asNumber(style.indentLevel, 0)));
  templateCell.numberFormat = [[style.numberFormat || DEFAULT_SHEET_FORMAT_STYLE.numberFormat]];
}

async function ensureSheetFormatStore(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const existing = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
  existing.load("name");
  await context.sync();

  const store = existing.isNullObject
    ? context.workbook.worksheets.add(SHEET_FORMAT_STORE_SHEET)
    : (existing as Excel.Worksheet);
  store.visibility = Excel.SheetVisibility.hidden;
  store.getRange("A1:H1").values = [
    [
      "Id",
      "Name",
      "SourceSheet",
      "SourceAddress",
      "CreatedAt",
      "UpdatedAt",
      "Template",
      "TemplateSheet",
    ],
  ];
  return store;
}

async function findSheetFormatRowIndex(
  context: Excel.RequestContext,
  store: Excel.Worksheet,
  id: string
): Promise<number> {
  const used = store.getUsedRangeOrNullObject(true);
  used.load("rowCount");
  await context.sync();

  if (used.isNullObject || used.rowCount <= 1) {
    return -1;
  }

  const idsRange = store.getRangeByIndexes(1, 0, used.rowCount - 1, 1);
  idsRange.load("values");
  await context.sync();

  const targetId = id.trim();
  for (let i = 0; i < idsRange.values.length; i += 1) {
    if (asString(idsRange.values[i]?.[0]).trim() === targetId) {
      return i + 1;
    }
  }

  return -1;
}

async function readSheetFormatsFromStore(
  context: Excel.RequestContext,
  store: Excel.Worksheet
): Promise<SheetFormatRecord[]> {
  const used = store.getUsedRangeOrNullObject(true);
  used.load("rowCount");
  await context.sync();

  if (used.isNullObject || used.rowCount <= 1) {
    return [];
  }

  const rowCount = used.rowCount - 1;
  const metaRange = store.getRangeByIndexes(1, 0, rowCount, 8);
  metaRange.load("values");
  await context.sync();

  const pendingRows: Array<{
    rowIndex: number;
    id: string;
    name: string;
    sourceSheet: string;
    sourceAddress: string;
    createdAt: string;
    updatedAt: string;
    templateSheetName: string;
    templateCell: Excel.Range;
  }> = [];

  for (let i = 0; i < metaRange.values.length; i += 1) {
    const row = metaRange.values[i] as unknown[];
    const id = asString(row?.[0]).trim();
    const name = asString(row?.[1]).trim();
    if (!id || !name) {
      continue;
    }
    const rowIndex = i + 1;
    const templateCell = store.getRangeByIndexes(rowIndex, 6, 1, 1);
    loadSheetFormatTemplateCell(templateCell);
    pendingRows.push({
      rowIndex,
      id,
      name,
      sourceSheet: asString(row?.[2]),
      sourceAddress: asString(row?.[3]),
      createdAt: asString(row?.[4]),
      updatedAt: asString(row?.[5]),
      templateSheetName: asString(row?.[7]),
      templateCell,
    });
  }

  if (pendingRows.length === 0) {
    return [];
  }

  await context.sync();

  const records = pendingRows.map((entry) => ({
    id: entry.id,
    name: entry.name,
    sourceSheet: entry.sourceSheet,
    sourceAddress: entry.sourceAddress,
    createdAt: entry.createdAt,
    updatedAt: entry.updatedAt,
    style: buildSheetFormatStyle(entry.templateCell),
  }));

  records.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
  return records;
}

export async function listSheetFormats(): Promise<SheetFormatRecord[]> {
  return Excel.run(async (context) => {
    const store = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
    store.load("name");
    await context.sync();
    if (store.isNullObject) {
      return [];
    }
    return readSheetFormatsFromStore(context, store as Excel.Worksheet);
  });
}

export async function captureSheetFormat(name: string): Promise<SheetFormatRecord> {
  const trimmedName = name.trim();
  if (!trimmedName) {
    throw new Error("Format name is required.");
  }

  return Excel.run(async (context) => {
    const store = await ensureSheetFormatStore(context);
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    const sourceCell = context.workbook.getSelectedRange().getCell(0, 0);
    sourceCell.load("address");
    const usedRange = activeSheet.getUsedRangeOrNullObject(true);
    usedRange.load("address");

    const used = store.getUsedRangeOrNullObject(true);
    used.load("rowCount");
    await context.sync();

    const rowIndex = used.isNullObject ? 1 : Math.max(1, used.rowCount);
    const id = `fmt-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
    const nowIso = new Date().toISOString();
    const parsedAddress = !usedRange.isNullObject
      ? splitQualifiedAddress(usedRange.address)
      : splitQualifiedAddress(sourceCell.address);
    const sourceAddress =
      parsedAddress.address || (!usedRange.isNullObject ? usedRange.address : sourceCell.address);

    const templateSheet = activeSheet.copy(Excel.WorksheetPositionType.end);
    const templateSheetName =
      `${SHEET_FORMAT_TEMPLATE_PREFIX}${Date.now().toString(36)}${Math.random()
        .toString(36)
        .slice(2, 6)}`.slice(0, 31);
    templateSheet.name = templateSheetName;
    templateSheet.visibility = Excel.SheetVisibility.hidden;

    const metaRange = store.getRangeByIndexes(rowIndex, 0, 1, 8);
    metaRange.values = [
      [id, trimmedName, activeSheet.name, sourceAddress, nowIso, nowIso, "", templateSheetName],
    ];

    const templateCell = store.getRangeByIndexes(rowIndex, 6, 1, 1);
    templateCell.clear(Excel.ClearApplyTo.all);
    templateCell.copyFrom(sourceCell, Excel.RangeCopyType.formats);
    templateCell.values = [[""]];
    loadSheetFormatTemplateCell(templateCell);

    await context.sync();

    return {
      id,
      name: trimmedName,
      sourceSheet: activeSheet.name,
      sourceAddress,
      createdAt: nowIso,
      updatedAt: nowIso,
      style: buildSheetFormatStyle(templateCell),
    };
  });
}

async function applySheetFormatInternal(id: string, applyToAllSheets: boolean): Promise<void> {
  const trimmedId = id.trim();
  if (!trimmedId) {
    throw new Error("A sheet format must be selected.");
  }

  await Excel.run(async (context) => {
    const store = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
    store.load("name");
    await context.sync();
    if (store.isNullObject) {
      throw new Error("No sheet formats are saved yet.");
    }

    const rowIndex = await findSheetFormatRowIndex(context, store as Excel.Worksheet, trimmedId);
    if (rowIndex < 0) {
      throw new Error("Selected sheet format was not found.");
    }

    const templateSheetCell = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 7, 1, 1);
    templateSheetCell.load("values");
    await context.sync();
    const templateSheetName = asString(templateSheetCell.values?.[0]?.[0]).trim();

    if (templateSheetName) {
      const templateSheet = context.workbook.worksheets.getItemOrNullObject(templateSheetName);
      templateSheet.load("name");
      await context.sync();

      if (!templateSheet.isNullObject) {
        const templateUsed = (templateSheet as Excel.Worksheet).getUsedRangeOrNullObject(true);
        templateUsed.load("address");
        await context.sync();

        if (!templateUsed.isNullObject) {
          const parsed = splitQualifiedAddress(templateUsed.address);
          const targetAddress = parsed.address || templateUsed.address;

          if (applyToAllSheets) {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();
            worksheets.items.forEach((sheet) => {
              if (
                sheet.name === SHEET_FORMAT_STORE_SHEET ||
                sheet.name === templateSheetName ||
                sheet.name.startsWith(SHEET_FORMAT_TEMPLATE_PREFIX)
              ) {
                return;
              }
              const targetRange = sheet.getRange(targetAddress);
              targetRange.copyFrom(templateUsed, Excel.RangeCopyType.formats);
            });
          } else {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            const targetRange = activeSheet.getRange(targetAddress);
            targetRange.copyFrom(templateUsed, Excel.RangeCopyType.formats);
          }

          await context.sync();
          return;
        }
      }
    }

    const templateCell = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 6, 1, 1);
    if (applyToAllSheets) {
      throw new Error(
        "This format was saved before full-sheet capture support. Recapture it first, then apply to all sheets."
      );
    }
    const selection = context.workbook.getSelectedRange();
    selection.copyFrom(templateCell, Excel.RangeCopyType.formats);
    await context.sync();
  });
}

export async function applySheetFormat(id: string): Promise<void> {
  await applySheetFormatInternal(id, false);
}

export async function applySheetFormatToAllSheets(id: string): Promise<void> {
  await applySheetFormatInternal(id, true);
}

export async function updateSheetFormat(
  id: string,
  request: UpdateSheetFormatRequest
): Promise<SheetFormatRecord> {
  const trimmedId = id.trim();
  const trimmedName = request.name.trim();
  if (!trimmedId) {
    throw new Error("A sheet format must be selected.");
  }
  if (!trimmedName) {
    throw new Error("Format name is required.");
  }

  return Excel.run(async (context) => {
    const store = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
    store.load("name");
    await context.sync();
    if (store.isNullObject) {
      throw new Error("No sheet format store is available.");
    }

    const rowIndex = await findSheetFormatRowIndex(context, store as Excel.Worksheet, trimmedId);
    if (rowIndex < 0) {
      throw new Error("Selected sheet format was not found.");
    }

    const metaRange = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 0, 1, 8);
    metaRange.load("values");
    await context.sync();
    const values = (metaRange.values[0] as unknown[]) ?? [];
    const sourceSheet = asString(values[2]);
    const sourceAddress = asString(values[3]);
    const createdAt = asString(values[4], new Date().toISOString());
    const updatedAt = new Date().toISOString();
    const templateSheetName = asString(values[7]);

    metaRange.values = [
      [
        trimmedId,
        trimmedName,
        sourceSheet,
        sourceAddress,
        createdAt,
        updatedAt,
        "",
        templateSheetName,
      ],
    ];

    const templateCell = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 6, 1, 1);
    applySheetFormatStyle(templateCell, request.style);
    loadSheetFormatTemplateCell(templateCell);

    await context.sync();

    return {
      id: trimmedId,
      name: trimmedName,
      sourceSheet,
      sourceAddress,
      createdAt,
      updatedAt,
      style: buildSheetFormatStyle(templateCell),
    };
  });
}

export async function recaptureSheetFormatFromSelection(id: string): Promise<SheetFormatRecord> {
  const trimmedId = id.trim();
  if (!trimmedId) {
    throw new Error("A sheet format must be selected.");
  }

  return Excel.run(async (context) => {
    const store = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
    store.load("name");
    await context.sync();
    if (store.isNullObject) {
      throw new Error("No sheet format store is available.");
    }

    const rowIndex = await findSheetFormatRowIndex(context, store as Excel.Worksheet, trimmedId);
    if (rowIndex < 0) {
      throw new Error("Selected sheet format was not found.");
    }

    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    const sourceCell = context.workbook.getSelectedRange().getCell(0, 0);
    sourceCell.load("address");
    const sourceUsedRange = activeSheet.getUsedRangeOrNullObject(true);
    sourceUsedRange.load("address");

    const metaRange = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 0, 1, 8);
    metaRange.load("values");
    await context.sync();

    const values = (metaRange.values[0] as unknown[]) ?? [];
    const name = asString(values[1], "Sheet Format");
    const createdAt = asString(values[4], new Date().toISOString());
    const existingTemplateSheetName = asString(values[7]).trim();
    const parsedAddress = !sourceUsedRange.isNullObject
      ? splitQualifiedAddress(sourceUsedRange.address)
      : splitQualifiedAddress(sourceCell.address);
    const sourceAddress =
      parsedAddress.address ||
      (!sourceUsedRange.isNullObject ? sourceUsedRange.address : sourceCell.address);
    const updatedAt = new Date().toISOString();
    const templateSheetName =
      `${SHEET_FORMAT_TEMPLATE_PREFIX}${Date.now().toString(36)}${Math.random()
        .toString(36)
        .slice(2, 6)}`.slice(0, 31);

    if (existingTemplateSheetName) {
      const existingTemplateSheet =
        context.workbook.worksheets.getItemOrNullObject(existingTemplateSheetName);
      existingTemplateSheet.load("name");
      await context.sync();
      if (!existingTemplateSheet.isNullObject) {
        existingTemplateSheet.delete();
      }
    }

    const copiedTemplateSheet = activeSheet.copy(Excel.WorksheetPositionType.end);
    copiedTemplateSheet.name = templateSheetName;
    copiedTemplateSheet.visibility = Excel.SheetVisibility.hidden;

    metaRange.values = [
      [
        trimmedId,
        name,
        activeSheet.name,
        sourceAddress,
        createdAt,
        updatedAt,
        "",
        templateSheetName,
      ],
    ];

    const templateCell = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 6, 1, 1);
    templateCell.clear(Excel.ClearApplyTo.all);
    templateCell.copyFrom(sourceCell, Excel.RangeCopyType.formats);
    templateCell.values = [[""]];
    loadSheetFormatTemplateCell(templateCell);

    await context.sync();

    return {
      id: trimmedId,
      name,
      sourceSheet: activeSheet.name,
      sourceAddress,
      createdAt,
      updatedAt,
      style: buildSheetFormatStyle(templateCell),
    };
  });
}

export async function deleteSheetFormat(id: string): Promise<void> {
  const trimmedId = id.trim();
  if (!trimmedId) {
    throw new Error("A sheet format must be selected.");
  }

  await Excel.run(async (context) => {
    const store = context.workbook.worksheets.getItemOrNullObject(SHEET_FORMAT_STORE_SHEET);
    store.load("name");
    await context.sync();
    if (store.isNullObject) {
      throw new Error("No sheet format store is available.");
    }

    const rowIndex = await findSheetFormatRowIndex(context, store as Excel.Worksheet, trimmedId);
    if (rowIndex < 0) {
      throw new Error("Selected sheet format was not found.");
    }

    const templateSheetCell = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 7, 1, 1);
    templateSheetCell.load("values");
    await context.sync();
    const templateSheetName = asString(templateSheetCell.values?.[0]?.[0]).trim();
    if (templateSheetName) {
      const templateSheet = context.workbook.worksheets.getItemOrNullObject(templateSheetName);
      templateSheet.load("name");
      await context.sync();
      if (!templateSheet.isNullObject) {
        templateSheet.delete();
      }
    }

    const rowRange = (store as Excel.Worksheet).getRangeByIndexes(rowIndex, 0, 1, 8);
    rowRange.delete(Excel.DeleteShiftDirection.up);
    await context.sync();
  });
}

function isInternalWorkbookSheet(sheetName: string): boolean {
  if (!sheetName) {
    return true;
  }
  if (
    sheetName === SHEET_FORMAT_STORE_SHEET ||
    sheetName === "__WBM_META" ||
    sheetName === "__WBM_FUNCTION_EVAL"
  ) {
    return true;
  }
  return sheetName.startsWith(SHEET_FORMAT_TEMPLATE_PREFIX);
}

function buildDestinationId(
  type: NavigationDestinationType,
  sheetName: string,
  objectName: string
): string {
  if (type === "Sheet") {
    return `sheet::${sheetName}`;
  }
  if (type === "NamedRange") {
    const normalizedSheetName = sheetName.trim();
    if (normalizedSheetName && normalizedSheetName.toLowerCase() !== "workbook") {
      return `named::${normalizedSheetName}::${objectName}`;
    }
    return `named::${objectName}`;
  }
  if (type === "Table") {
    return `table::${sheetName}::${objectName}`;
  }
  return `chart::${sheetName}::${objectName}`;
}

interface ParsedNavigationDestination {
  type: NavigationDestinationType | "ExternalUrl" | "SheetCell";
  sheetName: string;
  objectName: string;
}

function parseLegacyDestinationId(destination: string): ParsedNavigationDestination {
  const normalized = normalizeMaybeAddress(destination);
  const parsed = splitQualifiedAddress(normalized);
  if (parsed.sheet && parsed.address) {
    if (isCellAddressExpression(parsed.address)) {
      return { type: "SheetCell", sheetName: parsed.sheet, objectName: parsed.address };
    }
    return { type: "NamedRange", sheetName: parsed.sheet, objectName: parsed.address };
  }
  if (isCellAddressExpression(normalized)) {
    return { type: "SheetCell", sheetName: "", objectName: normalized };
  }
  return { type: "NamedRange", sheetName: "", objectName: normalized };
}

function parseDestinationId(destinationId: string): ParsedNavigationDestination {
  const normalized = destinationId.trim();
  if (!normalized) {
    throw new Error("Destination is required.");
  }
  const parts = normalized.split("::");
  const prefix = parts[0]?.toLowerCase();
  if (prefix === "url" && parts.length >= 2) {
    return { type: "ExternalUrl", sheetName: "", objectName: parts.slice(1).join("::") };
  }
  if (prefix === "cell" && parts.length >= 3) {
    return { type: "SheetCell", sheetName: parts[1], objectName: parts.slice(2).join("::") };
  }
  if (prefix === "sheet" && parts.length >= 2) {
    return { type: "Sheet", sheetName: parts.slice(1).join("::"), objectName: "" };
  }
  if (prefix === "named" && parts.length >= 3) {
    return { type: "NamedRange", sheetName: parts[1], objectName: parts.slice(2).join("::") };
  }
  if (prefix === "named" && parts.length >= 2) {
    return { type: "NamedRange", sheetName: "", objectName: parts.slice(1).join("::") };
  }
  if (prefix === "table" && parts.length >= 3) {
    return { type: "Table", sheetName: parts[1], objectName: parts.slice(2).join("::") };
  }
  if (prefix === "chart" && parts.length >= 3) {
    return { type: "Chart", sheetName: parts[1], objectName: parts.slice(2).join("::") };
  }
  if (!normalized.includes("::")) {
    return parseLegacyDestinationId(normalized);
  }
  throw new Error("Invalid navigation destination.");
}

export async function listNavigationDestinations(): Promise<NavigationDestinationOption[]> {
  return Excel.run(async (context) => {
    const options: NavigationDestinationOption[] = [];
    const workbookNames = context.workbook.names;
    workbookNames.load("items/name,items/formula");

    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const tableMaps = worksheets.items.map((sheet) => {
      const tables = sheet.tables;
      tables.load("items/name");
      const charts = sheet.charts;
      charts.load("items/name");
      const names = sheet.names;
      names.load("items/name,items/formula");
      return { sheetName: sheet.name, tables, charts, names };
    });
    await context.sync();

    worksheets.items.forEach((sheet) => {
      if (isInternalWorkbookSheet(sheet.name)) {
        return;
      }
      options.push({
        id: buildDestinationId("Sheet", sheet.name, ""),
        type: "Sheet",
        label: `Sheet: ${sheet.name}`,
        sheetName: sheet.name,
        address: "A1",
      });
    });

    workbookNames.items.forEach((namedItem) => {
      const name = namedItem.name;
      const formula = normalizeDisplayAddress(asString(namedItem.formula));
      if (name.toLowerCase().startsWith("_xl") || name.includes("(") || formula.includes("(")) {
        return;
      }
      const parsed = splitQualifiedAddress(asString(namedItem.formula));
      options.push({
        id: buildDestinationId("NamedRange", "", name),
        type: "NamedRange",
        label: `Name: ${name}`,
        sheetName: parsed.sheet || "Workbook",
        address: parsed.address || formula,
      });
    });

    tableMaps.forEach((entry) => {
      if (isInternalWorkbookSheet(entry.sheetName)) {
        return;
      }
      entry.names.items.forEach((namedItem) => {
        const name = namedItem.name;
        const formula = normalizeDisplayAddress(asString(namedItem.formula));
        if (name.toLowerCase().startsWith("_xl") || name.includes("(") || formula.includes("(")) {
          return;
        }
        const parsed = splitQualifiedAddress(asString(namedItem.formula));
        options.push({
          id: buildDestinationId("NamedRange", entry.sheetName, name),
          type: "NamedRange",
          label: `Name: ${name} (${entry.sheetName})`,
          sheetName: parsed.sheet || entry.sheetName,
          address: parsed.address || formula,
        });
      });
      entry.tables.items.forEach((table) => {
        options.push({
          id: buildDestinationId("Table", entry.sheetName, table.name),
          type: "Table",
          label: `Table: ${table.name} (${entry.sheetName})`,
          sheetName: entry.sheetName,
          address: table.name,
        });
      });
      entry.charts.items.forEach((chart) => {
        options.push({
          id: buildDestinationId("Chart", entry.sheetName, chart.name),
          type: "Chart",
          label: `Chart: ${chart.name} (${entry.sheetName})`,
          sheetName: entry.sheetName,
          address: chart.name,
        });
      });
    });

    options.sort((left, right) =>
      left.label.localeCompare(right.label, undefined, { sensitivity: "base" })
    );
    return options;
  });
}

export async function listWorkbookThemeColors(): Promise<string[]> {
  try {
    return await Excel.run(async (context) => {
      const metaSheetOrNull = context.workbook.worksheets.getItemOrNullObject("__WBM_META");
      metaSheetOrNull.load("name");
      await context.sync();

      const metaSheet = metaSheetOrNull.isNullObject
        ? context.workbook.worksheets.add("__WBM_META")
        : (metaSheetOrNull as Excel.Worksheet);
      metaSheet.visibility = Excel.SheetVisibility.hidden;

      const cells = WORKBOOK_THEME_STYLE_SEQUENCE.map((styleName, index) => {
        const cell = metaSheet.getRangeByIndexes(index, 2, 1, 1);
        cell.style = styleName;
        cell.load("format/fill/color");
        return cell;
      });
      await context.sync();

      const seen = new Set<string>();
      const colors = cells
        .map((cell) => sanitizeHexColor(asString(cell.format.fill.color), ""))
        .filter((color) => {
          if (!color || seen.has(color)) {
            return false;
          }
          seen.add(color);
          return true;
        });

      return colors.length ? colors : [...WORKBOOK_THEME_SWATCH_FALLBACK];
    });
  } catch {
    return [...WORKBOOK_THEME_SWATCH_FALLBACK];
  }
}

export async function activateNavigationDestination(destinationId: string): Promise<void> {
  const parsed = parseDestinationId(destinationId);

  if (parsed.type === "ExternalUrl") {
    if (typeof window !== "undefined") {
      window.open(parsed.objectName, "_blank");
    }
    return;
  }

  await Excel.run(async (context) => {
    if (parsed.type === "SheetCell") {
      const targetSheet = parsed.sheetName
        ? context.workbook.worksheets.getItem(parsed.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      targetSheet.activate();
      const targetRange = targetSheet.getRange(parsed.objectName);
      targetRange.select();
      await context.sync();
      return;
    }

    if (parsed.type === "Sheet") {
      const targetSheet = context.workbook.worksheets.getItem(parsed.sheetName);
      targetSheet.activate();
      targetSheet.getRange("A1").select();
      await context.sync();
      return;
    }

    if (parsed.type === "NamedRange") {
      const namedRangeName = parsed.objectName.trim();
      if (!namedRangeName) {
        throw new Error("Named range destination is missing a name.");
      }

      if (parsed.sheetName) {
        const sheet = context.workbook.worksheets.getItem(parsed.sheetName);
        const localNamed = sheet.names.getItemOrNullObject(namedRangeName);
        localNamed.load("name");
        await context.sync();
        if (!localNamed.isNullObject) {
          sheet.activate();
          localNamed.getRange().select();
          await context.sync();
          return;
        }
      }

      const workbookNamed = context.workbook.names.getItemOrNullObject(namedRangeName);
      workbookNamed.load("name");
      await context.sync();
      if (!workbookNamed.isNullObject) {
        workbookNamed.getRange().select();
        await context.sync();
        return;
      }
      throw new Error(`Named range "${namedRangeName}" was not found.`);
    }

    if (parsed.type === "Table") {
      const sheet = context.workbook.worksheets.getItem(parsed.sheetName);
      sheet.activate();
      const table = sheet.tables.getItem(parsed.objectName);
      table.getRange().select();
      await context.sync();
      return;
    }

    const sheet = context.workbook.worksheets.getItem(parsed.sheetName);
    sheet.activate();
    const chart = sheet.charts.getItem(parsed.objectName);
    chart.activate();
    await context.sync();
  });
}

export async function applyNavigationTemplate(
  request: ApplyNavigationTemplateRequest
): Promise<void> {
  if (!request.buttons.length) {
    throw new Error("Add at least one navigation button before applying.");
  }

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();

    const targets = request.applyToAllSheets
      ? worksheets.items.filter((sheet) => !isInternalWorkbookSheet(sheet.name))
      : worksheets.items.filter((sheet) => sheet.name === activeSheet.name);
    if (!targets.length) {
      throw new Error("No target worksheet was available.");
    }

    for (const sheet of targets) {
      const shapes = sheet.shapes;
      shapes.load("items/name,items/altTextDescription");
      await context.sync();

      shapes.items.forEach((shape) => {
        const alt = asString(shape.altTextDescription);
        if (
          shape.name.startsWith(NAV_SHAPE_PREFIX) ||
          shape.name.startsWith(NAV_SHADOW_PREFIX) ||
          shape.name.startsWith(NAV_PANEL_PREFIX) ||
          alt.startsWith(NAV_TARGET_PREFIX)
        ) {
          shape.delete();
        }
      });
      await context.sync();

      const navId = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
      const panelShape = sheet.shapes.addGeometricShape(
        "Round2SameRectangle" as unknown as Excel.GeometricShapeType
      );
      panelShape.name = `${NAV_PANEL_PREFIX}${navId}`;
      panelShape.left = Math.max(0, request.originLeft);
      panelShape.top = Math.max(0, request.originTop);
      panelShape.fill.setSolidColor(request.panelFillColor);
      panelShape.lineFormat.color = request.panelOutlineColor;
      panelShape.lineFormat.weight = 1.5;
      panelShape.textFrame.textRange.text = "";

      const buttonCount = request.buttons.length;
      if (request.layout === "Vertical") {
        panelShape.width = request.panelWidth;
        panelShape.height =
          request.panelPadding * 2 +
          buttonCount * request.buttonHeight +
          Math.max(0, buttonCount - 1) * request.buttonGap;
      } else {
        panelShape.width =
          request.panelPadding * 2 +
          buttonCount * request.buttonWidth +
          Math.max(0, buttonCount - 1) * request.buttonGap;
        panelShape.height = request.panelPadding * 2 + request.buttonHeight;
      }

      request.buttons.forEach((button, index) => {
        const buttonShape = sheet.shapes.addGeometricShape(
          "Round2SameRectangle" as unknown as Excel.GeometricShapeType
        );
        buttonShape.name = `${NAV_SHAPE_PREFIX}${navId}_${index + 1}`;
        buttonShape.fill.setSolidColor(button.fillColor);
        buttonShape.lineFormat.color = button.outlineColor;
        buttonShape.lineFormat.weight = 1;
        buttonShape.textFrame.textRange.text = button.label;
        buttonShape.textFrame.textRange.font.color = button.fontColor;
        buttonShape.textFrame.textRange.font.bold = true;
        buttonShape.textFrame.horizontalAlignment =
          "Left" as unknown as Excel.ShapeTextHorizontalAlignment;
        buttonShape.textFrame.verticalAlignment =
          "Middle" as unknown as Excel.ShapeTextVerticalAlignment;
        buttonShape.altTextDescription = `${NAV_TARGET_PREFIX}${button.destinationId}`;
        buttonShape.width = request.buttonWidth;
        buttonShape.height = request.buttonHeight;

        if (request.layout === "Vertical") {
          buttonShape.left = request.originLeft + request.panelPadding;
          buttonShape.top =
            request.originTop +
            request.panelPadding +
            index * (request.buttonHeight + request.buttonGap);
        } else {
          buttonShape.left =
            request.originLeft +
            request.panelPadding +
            index * (request.buttonWidth + request.buttonGap);
          buttonShape.top = request.originTop + request.panelPadding;
        }
      });

      await context.sync();
    }
  });
}

export interface ShapeBuilderIconOption {
  key: string;
  label: string;
  symbol: string;
  keywords: string[];
}

const RAW_SHAPE_BUILDER_ICON_OPTIONS: Array<Omit<ShapeBuilderIconOption, "keywords">> = [
  { key: "none", label: "None", symbol: "" },
  { key: "home", label: "Home", symbol: "⌂" },
  { key: "home_alt", label: "Home Alt", symbol: "🏠" },
  { key: "dashboard", label: "Dashboard", symbol: "⌗" },
  { key: "data", label: "Data", symbol: "▦" },
  { key: "controls", label: "Controls", symbol: "⌘" },
  { key: "reports", label: "Reports", symbol: "▤" },
  { key: "table", label: "Table", symbol: "▥" },
  { key: "database", label: "Database", symbol: "🗄" },
  { key: "document", label: "Document", symbol: "🗎" },
  { key: "folder", label: "Folder", symbol: "📁" },
  { key: "file", label: "File", symbol: "📄" },
  { key: "template", label: "Template", symbol: "🗒" },
  { key: "search", label: "Search", symbol: "⌕" },
  { key: "settings", label: "Settings", symbol: "⚙" },
  { key: "gear_alt", label: "Gear Alt", symbol: "⚒" },
  { key: "tools", label: "Tools", symbol: "🛠" },
  { key: "wrench", label: "Wrench", symbol: "🔧" },
  { key: "palette", label: "Palette", symbol: "🎨" },
  { key: "help", label: "Help", symbol: "?" },
  { key: "info", label: "Info", symbol: "ℹ" },
  { key: "warning", label: "Warning", symbol: "⚠" },
  { key: "check", label: "Check", symbol: "✓" },
  { key: "close", label: "Close", symbol: "✕" },
  { key: "plus", label: "Plus", symbol: "+" },
  { key: "minus", label: "Minus", symbol: "−" },
  { key: "edit", label: "Edit", symbol: "✎" },
  { key: "save", label: "Save", symbol: "💾" },
  { key: "refresh", label: "Refresh", symbol: "↻" },
  { key: "sync", label: "Sync", symbol: "⟳" },
  { key: "filter", label: "Filter", symbol: "⛃" },
  { key: "sort", label: "Sort", symbol: "⇅" },
  { key: "link", label: "Link", symbol: "🔗" },
  { key: "external", label: "External", symbol: "↗" },
  { key: "download", label: "Download", symbol: "⬇" },
  { key: "upload", label: "Upload", symbol: "⬆" },
  { key: "share", label: "Share", symbol: "⤴" },
  { key: "mail", label: "Mail", symbol: "✉" },
  { key: "calendar", label: "Calendar", symbol: "📅" },
  { key: "clock", label: "Clock", symbol: "🕒" },
  { key: "phone", label: "Phone", symbol: "☎" },
  { key: "chat", label: "Chat", symbol: "💬" },
  { key: "user", label: "User", symbol: "👤" },
  { key: "users", label: "Users", symbol: "👥" },
  { key: "group", label: "Group", symbol: "👪" },
  { key: "lock", label: "Lock", symbol: "🔒" },
  { key: "unlock", label: "Unlock", symbol: "🔓" },
  { key: "key", label: "Key", symbol: "🔑" },
  { key: "shield", label: "Shield", symbol: "🛡" },
  { key: "bell", label: "Bell", symbol: "🔔" },
  { key: "bookmark", label: "Bookmark", symbol: "🔖" },
  { key: "star", label: "Star", symbol: "★" },
  { key: "flag", label: "Flag", symbol: "⚑" },
  { key: "pin", label: "Pin", symbol: "📌" },
  { key: "chart_bar", label: "Chart Bar", symbol: "📊" },
  { key: "chart_line", label: "Chart Line", symbol: "📈" },
  { key: "chart_down", label: "Chart Down", symbol: "📉" },
  { key: "chart_pie", label: "Chart Pie", symbol: "◔" },
  { key: "chart_area", label: "Chart Area", symbol: "▰" },
  { key: "target", label: "Target", symbol: "◎" },
  { key: "lightning", label: "Lightning", symbol: "⚡" },
  { key: "rocket", label: "Rocket", symbol: "🚀" },
  { key: "globe", label: "Globe", symbol: "🌐" },
  { key: "map", label: "Map", symbol: "🗺" },
  { key: "location", label: "Location", symbol: "📍" },
  { key: "compass", label: "Compass", symbol: "🧭" },
  { key: "camera", label: "Camera", symbol: "📷" },
  { key: "image", label: "Image", symbol: "🖼" },
  { key: "video", label: "Video", symbol: "🎬" },
  { key: "play", label: "Play", symbol: "▶" },
  { key: "pause", label: "Pause", symbol: "⏸" },
  { key: "stop", label: "Stop", symbol: "⏹" },
  { key: "arrow_left", label: "Arrow Left", symbol: "←" },
  { key: "arrow_right", label: "Arrow Right", symbol: "→" },
  { key: "arrow_up", label: "Arrow Up", symbol: "↑" },
  { key: "arrow_down", label: "Arrow Down", symbol: "↓" },
  { key: "chevron_left", label: "Chevron Left", symbol: "‹" },
  { key: "chevron_right", label: "Chevron Right", symbol: "›" },
  { key: "chevron_up", label: "Chevron Up", symbol: "˄" },
  { key: "chevron_down", label: "Chevron Down", symbol: "˅" },
  { key: "back", label: "Back", symbol: "↩" },
  { key: "forward", label: "Forward", symbol: "↪" },
  { key: "undo", label: "Undo", symbol: "↶" },
  { key: "redo", label: "Redo", symbol: "↷" },
  { key: "print", label: "Print", symbol: "🖨" },
  { key: "clipboard", label: "Clipboard", symbol: "📋" },
  { key: "note", label: "Note", symbol: "📝" },
  { key: "code", label: "Code", symbol: "⌨" },
  { key: "formula", label: "Formula", symbol: "∑" },
  { key: "function", label: "Function", symbol: "ƒ" },
  { key: "calculator", label: "Calculator", symbol: "🧮" },
  { key: "cube", label: "Cube", symbol: "⬢" },
  { key: "diamond", label: "Diamond", symbol: "◆" },
  { key: "circle", label: "Circle", symbol: "●" },
  { key: "square", label: "Square", symbol: "■" },
  { key: "triangle", label: "Triangle", symbol: "▲" },
  { key: "hexagon", label: "Hexagon", symbol: "⬡" },
  { key: "eye", label: "Eye", symbol: "👁" },
  { key: "hide", label: "Hide", symbol: "🙈" },
  { key: "package", label: "Package", symbol: "📦" },
  { key: "truck", label: "Truck", symbol: "🚚" },
  { key: "money", label: "Money", symbol: "💲" },
  { key: "briefcase", label: "Briefcase", symbol: "💼" },
  { key: "building", label: "Building", symbol: "🏢" },
  { key: "spark", label: "Spark", symbol: "✦" },
  { key: "magic", label: "Magic", symbol: "✨" },
];

const EMOJI_SURROGATE_PATTERN = /[\uD800-\uDBFF][\uDC00-\uDFFF]/;

const tokenizeIconSearchTerms = (value: string): string[] =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .split(" ")
    .map((part) => part.trim())
    .filter((part) => part.length > 1);

const ICON_KEYWORD_OVERRIDES: Record<string, string[]> = {
  none: ["clear", "remove", "empty"],
  home: ["house", "start"],
  home_alt: ["house", "start"],
  data: ["analytics", "insights", "grid", "sheet"],
  controls: ["button", "action", "command"],
  reports: ["chart", "graph", "analytics"],
  table: ["grid", "cells"],
  database: ["store", "storage", "records"],
  settings: ["preferences", "options", "config"],
  tools: ["build", "configure"],
  help: ["support", "question"],
  info: ["details", "about"],
  warning: ["alert", "caution", "error"],
  check: ["success", "done", "confirm"],
  close: ["cancel", "dismiss"],
  edit: ["write", "compose", "text"],
  save: ["store", "persist"],
  refresh: ["reload", "update"],
  sync: ["refresh", "reload", "update"],
  filter: ["funnel", "criteria"],
  sort: ["order", "rank"],
  link: ["url", "hyperlink", "connect"],
  external: ["url", "open", "launch"],
  mail: ["email", "message"],
  users: ["people", "team"],
  group: ["people", "team"],
  chart_bar: ["report", "analytics", "graph"],
  chart_line: ["report", "analytics", "graph"],
  chart_down: ["report", "analytics", "graph"],
  chart_pie: ["report", "analytics", "graph"],
  chart_area: ["report", "analytics", "graph"],
  formula: ["math", "function", "fx"],
  function: ["formula", "math", "fx"],
  calculator: ["math", "formula"],
  eye: ["view", "preview", "show"],
  hide: ["view", "preview", "show"],
};

function fallbackIconSymbolFromKey(key: string): string {
  const normalized = key.toLowerCase();
  if (normalized.includes("none")) return "";
  if (normalized.includes("arrow_left")) return "←";
  if (normalized.includes("arrow_right")) return "→";
  if (normalized.includes("arrow_up")) return "↑";
  if (normalized.includes("arrow_down")) return "↓";
  if (normalized.includes("chevron_left")) return "‹";
  if (normalized.includes("chevron_right")) return "›";
  if (normalized.includes("chevron_up")) return "˄";
  if (normalized.includes("chevron_down")) return "˅";
  if (normalized.includes("home")) return "⌂";
  if (
    normalized.includes("dashboard") ||
    normalized.includes("data") ||
    normalized.includes("table") ||
    normalized.includes("database")
  )
    return "▦";
  if (normalized.includes("report") || normalized.includes("chart")) return "▤";
  if (
    normalized.includes("document") ||
    normalized.includes("file") ||
    normalized.includes("folder") ||
    normalized.includes("template")
  )
    return "▣";
  if (normalized.includes("search")) return "⌕";
  if (
    normalized.includes("setting") ||
    normalized.includes("gear") ||
    normalized.includes("tool") ||
    normalized.includes("wrench")
  )
    return "⚙";
  if (normalized.includes("help")) return "?";
  if (normalized.includes("info")) return "ℹ";
  if (normalized.includes("warning")) return "⚠";
  if (normalized.includes("check")) return "✓";
  if (normalized.includes("close")) return "✕";
  if (normalized.includes("plus")) return "+";
  if (normalized.includes("minus")) return "−";
  if (normalized.includes("edit")) return "✎";
  if (normalized.includes("save")) return "⎙";
  if (normalized.includes("refresh") || normalized.includes("sync")) return "↻";
  if (normalized.includes("filter")) return "≣";
  if (normalized.includes("sort")) return "⇅";
  if (
    normalized.includes("link") ||
    normalized.includes("external") ||
    normalized.includes("share")
  )
    return "↗";
  if (normalized.includes("download")) return "⬇";
  if (normalized.includes("upload")) return "⬆";
  if (normalized.includes("mail")) return "✉";
  if (normalized.includes("calendar") || normalized.includes("clock")) return "◷";
  if (normalized.includes("phone")) return "☎";
  if (normalized.includes("chat")) return "☰";
  if (normalized.includes("user") || normalized.includes("group")) return "◎";
  if (normalized.includes("lock")) return normalized.includes("unlock") ? "○" : "●";
  if (normalized.includes("key")) return "✦";
  if (normalized.includes("shield")) return "⬟";
  if (normalized.includes("bell")) return "◉";
  if (normalized.includes("bookmark")) return "▮";
  if (normalized.includes("star")) return "★";
  if (normalized.includes("flag")) return "⚑";
  if (normalized.includes("pin")) return "•";
  if (normalized.includes("target")) return "◎";
  if (normalized.includes("lightning")) return "⚡";
  if (normalized.includes("rocket")) return "▲";
  if (normalized.includes("globe")) return "◌";
  if (normalized.includes("map")) return "▧";
  if (normalized.includes("location") || normalized.includes("compass")) return "⌖";
  if (normalized.includes("camera")) return "◫";
  if (normalized.includes("image")) return "▣";
  if (normalized.includes("video") || normalized.includes("play")) return "▶";
  if (normalized.includes("pause")) return "⏸";
  if (normalized.includes("stop")) return "⏹";
  if (normalized.includes("back")) return "↩";
  if (normalized.includes("forward")) return "↪";
  if (normalized.includes("undo")) return "↶";
  if (normalized.includes("redo")) return "↷";
  if (normalized.includes("print")) return "⎙";
  if (normalized.includes("clipboard")) return "▣";
  if (normalized.includes("note")) return "✎";
  if (normalized.includes("code")) return "⌨";
  if (normalized.includes("formula")) return "∑";
  if (normalized.includes("function")) return "ƒ";
  if (normalized.includes("calculator")) return "∑";
  if (normalized.includes("cube")) return "⬢";
  if (normalized.includes("diamond")) return "◆";
  if (normalized.includes("circle")) return "●";
  if (normalized.includes("square")) return "■";
  if (normalized.includes("triangle")) return "▲";
  if (normalized.includes("hexagon")) return "⬡";
  if (normalized.includes("eye")) return "◉";
  if (normalized.includes("hide")) return "◌";
  if (
    normalized.includes("package") ||
    normalized.includes("truck") ||
    normalized.includes("briefcase") ||
    normalized.includes("building")
  )
    return "▣";
  if (normalized.includes("money")) return "$";
  if (normalized.includes("spark") || normalized.includes("magic")) return "✦";
  return "•";
}

function toExcelSafeIconSymbol(key: string, symbol: string): string {
  const trimmed = symbol.trim();
  if (!trimmed) {
    return "";
  }
  if (!EMOJI_SURROGATE_PATTERN.test(trimmed)) {
    return trimmed;
  }
  return fallbackIconSymbolFromKey(key);
}

function buildShapeBuilderIconKeywords(option: Omit<ShapeBuilderIconOption, "keywords">): string[] {
  const set = new Set<string>([
    ...tokenizeIconSearchTerms(option.key),
    ...tokenizeIconSearchTerms(option.label),
    ...(ICON_KEYWORD_OVERRIDES[option.key] ?? []),
  ]);
  return Array.from(set).sort((left, right) => left.localeCompare(right));
}

function normalizeFluentIconName(rawName: string): string {
  return rawName
    .replace(/\d+(Filled|Regular)$/i, "")
    .replace(/(Filled|Regular)$/i, "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .replace(/_/g, " ")
    .toLowerCase()
    .trim();
}

function loadFluentIconNameCatalog(): string[] {
  try {
    const regular =
      require("../../node_modules/@fluentui/react-icons/lib-cjs/utils/fonts/FluentSystemIcons-Regular.json") as Record<
        string,
        number
      >;
    const resizable =
      require("../../node_modules/@fluentui/react-icons/lib-cjs/utils/fonts/FluentSystemIcons-Resizable.json") as Record<
        string,
        number
      >;
    const unique = new Set<string>();
    [...Object.keys(regular), ...Object.keys(resizable)].forEach((name) => {
      const normalized = normalizeFluentIconName(name);
      if (normalized) {
        unique.add(normalized);
      }
    });
    return Array.from(unique).sort((left, right) => left.localeCompare(right));
  } catch {
    return [];
  }
}

export const SHAPE_BUILDER_ICON_OPTIONS: ShapeBuilderIconOption[] =
  RAW_SHAPE_BUILDER_ICON_OPTIONS.map((option) => ({
    ...option,
    symbol: toExcelSafeIconSymbol(option.key, option.symbol),
    keywords: buildShapeBuilderIconKeywords(option),
  }));

export const FLUENT_ICON_NAME_CATALOG: string[] = loadFluentIconNameCatalog();

const SHAPE_BUILDER_ICON_SYMBOLS: Record<string, string> = SHAPE_BUILDER_ICON_OPTIONS.reduce(
  (accumulator, option) => {
    accumulator[option.key] = option.symbol;
    return accumulator;
  },
  {} as Record<string, string>
);

interface ShapeBuilderMetadata {
  iconKey: string;
  linkType: ShapeBuilderLinkType;
  internalDestinationId: string;
  externalUrl: string;
  effect: ShapeBuilderEffect;
  anchorAddress: string;
}

const DEFAULT_SHAPE_BUILDER_METADATA: ShapeBuilderMetadata = {
  iconKey: "none",
  linkType: "None",
  internalDestinationId: "",
  externalUrl: "",
  effect: "None",
  anchorAddress: "",
};

function sanitizeHexColor(value: string, fallback: string): string {
  return /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : fallback;
}

function isNoFillColor(value: string): boolean {
  return value.trim().toUpperCase() === NO_FILL_COLOR_TOKEN;
}

function normalizeShapeBuilderFillColor(value: string, fallback: string): string {
  if (isNoFillColor(value)) {
    return NO_FILL_COLOR_TOKEN;
  }
  const normalizedFallback = isNoFillColor(fallback)
    ? "#A8D5AD"
    : sanitizeHexColor(fallback, "#A8D5AD");
  return sanitizeHexColor(value, normalizedFallback);
}

function applyShapeBuilderFill(shape: Excel.Shape, fillColor: string, fallback: string): string {
  const normalized = normalizeShapeBuilderFillColor(fillColor, fallback);
  if (normalized === NO_FILL_COLOR_TOKEN) {
    shape.fill.clear();
  } else {
    shape.fill.setSolidColor(normalized);
  }
  return normalized;
}

function getShapeBuilderShadowName(shapeName: string): string {
  return `${NAV_SHADOW_PREFIX}${shapeName}`;
}

async function syncShapeBuilderShadow(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  record: {
    shapeName: string;
    shapeType: InsertableShapeType;
    left: number;
    top: number;
    width: number;
    height: number;
    fillColor: string;
    effect: ShapeBuilderEffect;
  }
): Promise<void> {
  const shadowName = getShapeBuilderShadowName(record.shapeName);
  const shadowOrNull = sheet.shapes.getItemOrNullObject(shadowName);
  shadowOrNull.load("name");
  await context.sync();

  // Excel JavaScript API doesn't expose shape shadow preset APIs.
  // Remove any legacy synthetic "shadow clone" so we don't create duplicate shapes.
  if (!shadowOrNull.isNullObject) {
    shadowOrNull.delete();
  }
}

function normalizeShapeBuilderLinkType(value: string): ShapeBuilderLinkType {
  if (value === "Internal" || value === "External") {
    return value;
  }
  return "None";
}

function normalizeShapeBuilderEffect(value: string): ShapeBuilderEffect {
  return value === "Shadow" ? "Shadow" : "None";
}

function normalizeShapeBuilderTextHorizontalAlignment(
  value: string
): ShapeBuilderTextHorizontalAlignment {
  const normalized = value.trim().toLowerCase();
  if (normalized === "center") {
    return "Center";
  }
  if (normalized === "right") {
    return "Right";
  }
  return "Left";
}

function normalizeShapeBuilderTextVerticalAlignment(
  value: string
): ShapeBuilderTextVerticalAlignment {
  const normalized = value.trim().toLowerCase();
  if (normalized === "top") {
    return "Top";
  }
  if (normalized === "bottom") {
    return "Bottom";
  }
  return "Middle";
}

function getShapeBuilderIconSymbol(iconKey: string): string {
  const key = iconKey.trim().toLowerCase();
  if (key.startsWith("custom:")) {
    const encoded = iconKey.trim().slice("custom:".length);
    try {
      return decodeURIComponent(encoded);
    } catch {
      return encoded;
    }
  }
  return SHAPE_BUILDER_ICON_SYMBOLS[key] ?? "";
}

function stripShapeBuilderIcon(displayText: string, iconKey: string): string {
  const symbol = getShapeBuilderIconSymbol(iconKey);
  if (!symbol) {
    return displayText;
  }
  const prefix = `${symbol} `;
  if (displayText.startsWith(prefix)) {
    return displayText.slice(prefix.length);
  }
  if (displayText === symbol) {
    return "";
  }
  if (displayText.startsWith(symbol)) {
    return displayText.slice(symbol.length);
  }
  return displayText;
}

function buildShapeBuilderDisplayText(iconKey: string, text: string): string {
  const symbol = getShapeBuilderIconSymbol(iconKey);
  const label = text ?? "";
  if (!symbol) {
    return label;
  }
  if (!label) {
    return symbol;
  }
  return /^\s/.test(label) ? `${symbol}${label}` : `${symbol} ${label}`;
}

function getShapeBuilderDisplayTextOffset(displayText: string, iconKey: string): number {
  const symbol = getShapeBuilderIconSymbol(iconKey);
  if (!symbol) {
    return 0;
  }
  const prefixedWithSpace = `${symbol} `;
  if (displayText.startsWith(prefixedWithSpace)) {
    return prefixedWithSpace.length;
  }
  if (displayText.startsWith(symbol)) {
    return symbol.length;
  }
  return 0;
}

function parseShapeBuilderMetadata(
  titleText: string,
  descriptionText: string
): ShapeBuilderMetadata {
  const defaultFromDescription = (() => {
    const anchorAddress = descriptionText.startsWith(SHAPE_ANCHOR_PREFIX)
      ? descriptionText.slice(SHAPE_ANCHOR_PREFIX.length).trim()
      : "";
    const target = descriptionText.startsWith(NAV_TARGET_PREFIX)
      ? descriptionText.slice(NAV_TARGET_PREFIX.length).trim()
      : "";
    if (target.startsWith("url::")) {
      return {
        ...DEFAULT_SHAPE_BUILDER_METADATA,
        linkType: "External" as ShapeBuilderLinkType,
        externalUrl: target.slice("url::".length).trim(),
        anchorAddress,
      };
    }
    if (target) {
      return {
        ...DEFAULT_SHAPE_BUILDER_METADATA,
        linkType: "Internal" as ShapeBuilderLinkType,
        internalDestinationId: target,
        anchorAddress,
      };
    }
    return { ...DEFAULT_SHAPE_BUILDER_METADATA, anchorAddress };
  })();

  if (!titleText.startsWith(SHAPE_BUILDER_META_PREFIX)) {
    return defaultFromDescription;
  }

  try {
    const raw = JSON.parse(titleText.slice(SHAPE_BUILDER_META_PREFIX.length));
    const asObject = raw as Partial<ShapeBuilderMetadata>;
    const linkType = normalizeShapeBuilderLinkType(
      asString(asObject.linkType, defaultFromDescription.linkType)
    );
    const metadata: ShapeBuilderMetadata = {
      iconKey: asString(asObject.iconKey, defaultFromDescription.iconKey).trim() || "none",
      linkType,
      internalDestinationId: asString(
        asObject.internalDestinationId,
        defaultFromDescription.internalDestinationId
      ).trim(),
      externalUrl: asString(asObject.externalUrl, defaultFromDescription.externalUrl).trim(),
      effect: normalizeShapeBuilderEffect(asString(asObject.effect, defaultFromDescription.effect)),
      anchorAddress: asString(asObject.anchorAddress, defaultFromDescription.anchorAddress).trim(),
    };
    if (metadata.linkType !== "Internal") {
      metadata.internalDestinationId = "";
    }
    if (metadata.linkType !== "External") {
      metadata.externalUrl = "";
    }
    return metadata;
  } catch {
    return defaultFromDescription;
  }
}

function applyShapeBuilderMetadata(shape: Excel.Shape, metadataInput: ShapeBuilderMetadata): void {
  const metadata: ShapeBuilderMetadata = {
    ...metadataInput,
    iconKey: metadataInput.iconKey.trim() || "none",
    linkType: normalizeShapeBuilderLinkType(metadataInput.linkType),
    internalDestinationId: metadataInput.internalDestinationId.trim(),
    externalUrl: metadataInput.externalUrl.trim(),
    effect: normalizeShapeBuilderEffect(metadataInput.effect),
    anchorAddress: metadataInput.anchorAddress.trim(),
  };

  if (metadata.linkType !== "Internal") {
    metadata.internalDestinationId = "";
  }
  if (metadata.linkType !== "External") {
    metadata.externalUrl = "";
  }

  shape.altTextTitle = `${SHAPE_BUILDER_META_PREFIX}${JSON.stringify(metadata)}`;
  if (metadata.linkType === "Internal" && metadata.internalDestinationId) {
    shape.altTextDescription = `${NAV_TARGET_PREFIX}${metadata.internalDestinationId}`;
    return;
  }
  if (metadata.linkType === "External" && metadata.externalUrl) {
    shape.altTextDescription = `${NAV_TARGET_PREFIX}url::${metadata.externalUrl}`;
    return;
  }
  shape.altTextDescription = "";
}

function loadShapeBuilderProperties(shape: Excel.Shape): void {
  shape.load(
    "id,name,type,geometricShapeType,left,top,zOrderPosition,width,height,altTextDescription,altTextTitle,fill/foregroundColor,fill/type,lineFormat/color,lineFormat/weight,textFrame/horizontalAlignment,textFrame/verticalAlignment,textFrame/textRange/text,textFrame/textRange/font/color,textFrame/textRange/font/size,textFrame/textRange/font/bold,textFrame/textRange/font/italic"
  );
}

function isShapeBuilderEditableShapeType(shapeType: string): boolean {
  return shapeType === "GeometricShape" || shapeType === "TextBox";
}

function buildShapeBuilderRecord(sheetName: string, shape: Excel.Shape): ShapeBuilderShapeRecord {
  const metadata = parseShapeBuilderMetadata(
    asString(shape.altTextTitle),
    asString(shape.altTextDescription)
  );
  const displayText = asString(shape.textFrame.textRange.text);
  const text = stripShapeBuilderIcon(displayText, metadata.iconKey);
  const fillColor =
    asString(shape.fill.type) === "NoFill"
      ? NO_FILL_COLOR_TOKEN
      : sanitizeHexColor(asString(shape.fill.foregroundColor), "#A8D5AD");

  return {
    id: `${sheetName}::${shape.name}`,
    sheetName,
    shapeName: shape.name,
    shapeType: inferInsertableShapeType(asString(shape.geometricShapeType), asString(shape.type)),
    anchorAddress: metadata.anchorAddress,
    text,
    iconKey: metadata.iconKey,
    fillColor,
    outlineColor: sanitizeHexColor(asString(shape.lineFormat.color), "#8CBF95"),
    outlineWidth: asNumber(shape.lineFormat.weight, 1),
    width: asNumber(shape.width, 160),
    height: asNumber(shape.height, 60),
    left: asNumber(shape.left, 100),
    top: asNumber(shape.top, 175),
    zOrderPosition: asNumber(shape.zOrderPosition, 0),
    fontColor: sanitizeHexColor(asString(shape.textFrame.textRange.font.color), "#1F2937"),
    fontSize: asNumber(shape.textFrame.textRange.font.size, 16),
    bold: asBoolean(shape.textFrame.textRange.font.bold, true),
    italic: asBoolean(shape.textFrame.textRange.font.italic, false),
    textHorizontalAlignment: normalizeShapeBuilderTextHorizontalAlignment(
      asString(shape.textFrame.horizontalAlignment)
    ),
    textVerticalAlignment: normalizeShapeBuilderTextVerticalAlignment(
      asString(shape.textFrame.verticalAlignment)
    ),
    linkType: metadata.linkType,
    internalDestinationId: metadata.internalDestinationId,
    externalUrl: metadata.externalUrl,
    effect: metadata.effect,
  };
}

function compareShapeBuilderRecordsBySelectionPanePosition(
  left: ShapeBuilderShapeRecord,
  right: ShapeBuilderShapeRecord
): number {
  return (
    right.zOrderPosition - left.zOrderPosition ||
    left.top - right.top ||
    left.left - right.left ||
    left.shapeName.localeCompare(right.shapeName)
  );
}

export async function listShapeBuilderShapes(): Promise<ShapeBuilderShapeRecord[]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const shapes = sheet.shapes;
    shapes.load("items/name,items/type,items/altTextTitle");
    await context.sync();

    const candidateShapes = shapes.items.filter(
      (shape) =>
        isShapeBuilderEditableShapeType(asString(shape.type)) &&
        !shape.name.startsWith(NAV_SHADOW_PREFIX) &&
        !shape.name.startsWith(NAV_PANEL_PREFIX)
    );

    candidateShapes.forEach((shape) => {
      loadShapeBuilderProperties(shape);
    });
    await context.sync();

    const records = candidateShapes.map((shape) => buildShapeBuilderRecord(sheet.name, shape));
    records.sort(compareShapeBuilderRecordsBySelectionPanePosition);
    return records;
  });
}

export async function createShapeBuilderShape(
  request: CreateShapeBuilderShapeRequest
): Promise<ShapeBuilderShapeRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const shapes = sheet.shapes;
    shapes.load("items/name,items/top,items/height,items/altTextTitle");
    await context.sync();

    let left = typeof request.left === "number" ? request.left : 40;
    let top = typeof request.top === "number" ? request.top : 140;
    if (typeof request.top !== "number") {
      const navShapes = shapes.items.filter((shape) => {
        const title = asString(shape.altTextTitle);
        return (
          title.startsWith(SHAPE_BUILDER_META_PREFIX) || shape.name.startsWith(NAV_SHAPE_PREFIX)
        );
      });
      if (navShapes.length > 0) {
        const maxBottom = Math.max(
          ...navShapes.map((shape) => asNumber(shape.top, 0) + asNumber(shape.height, 60))
        );
        top = maxBottom + 16;
      }
    }

    const shapeName = `${NAV_SHAPE_PREFIX}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    let shape: Excel.Shape;
    if (request.shapeType === "TextBox") {
      shape = sheet.shapes.addTextBox("");
    } else {
      shape = sheet.shapes.addGeometricShape(toGeometricShapeType(request.shapeType));
    }

    shape.name = shapeName;
    shape.left = Math.max(0, left);
    shape.top = Math.max(0, top);
    shape.width = Math.max(80, asNumber(request.width, 160));
    shape.height = Math.max(24, asNumber(request.height, 60));
    const normalizedFillColor = applyShapeBuilderFill(
      shape,
      asString(request.fillColor, "#A8D5AD"),
      "#A8D5AD"
    );
    shape.lineFormat.color = sanitizeHexColor(request.outlineColor, "#8CBF95");
    shape.lineFormat.weight = Math.max(0, asNumber(request.outlineWidth, 2));
    shape.textFrame.textRange.font.color = sanitizeHexColor(request.fontColor, "#1F2937");
    shape.textFrame.textRange.font.size = Math.max(8, asNumber(request.fontSize, 16));
    shape.textFrame.textRange.font.bold = asBoolean(request.bold, true);
    shape.textFrame.textRange.font.italic = asBoolean(request.italic, false);
    shape.textFrame.horizontalAlignment = normalizeShapeBuilderTextHorizontalAlignment(
      asString(request.textHorizontalAlignment, "Left")
    ) as unknown as Excel.ShapeTextHorizontalAlignment;
    shape.textFrame.verticalAlignment = normalizeShapeBuilderTextVerticalAlignment(
      asString(request.textVerticalAlignment, "Middle")
    ) as unknown as Excel.ShapeTextVerticalAlignment;
    shape.textFrame.textRange.text = buildShapeBuilderDisplayText(request.iconKey, request.text);
    shape.placement = Excel.Placement.oneCell;

    applyShapeBuilderMetadata(shape, {
      iconKey: request.iconKey,
      linkType: request.linkType,
      internalDestinationId: request.internalDestinationId,
      externalUrl: request.externalUrl,
      effect: request.effect,
      anchorAddress: asString(request.anchorAddress, ""),
    });
    await syncShapeBuilderShadow(context, sheet, {
      shapeName,
      shapeType: request.shapeType,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      fillColor: normalizedFillColor,
      effect: normalizeShapeBuilderEffect(request.effect),
    });

    loadShapeBuilderProperties(shape);
    await context.sync();

    return buildShapeBuilderRecord(sheet.name, shape);
  });
}

export async function updateShapeBuilderShape(
  sheetName: string,
  shapeName: string,
  updates: UpdateShapeBuilderShapeRequest
): Promise<ShapeBuilderShapeRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    loadShapeBuilderProperties(shape);
    await context.sync();

    const existing = buildShapeBuilderRecord(sheetName, shape);
    const nextShapeType = updates.shapeType ?? existing.shapeType;
    if (nextShapeType !== existing.shapeType) {
      if (nextShapeType === "TextBox" || existing.shapeType === "TextBox") {
        throw new Error("Switching between Text Box and geometric shape is not supported.");
      }
      (shape as unknown as { geometricShapeType: Excel.GeometricShapeType }).geometricShapeType =
        toGeometricShapeType(nextShapeType);
    }

    shape.left = Math.max(0, asNumber(updates.left, existing.left));
    shape.top = Math.max(0, asNumber(updates.top, existing.top));
    shape.width = Math.max(80, asNumber(updates.width, existing.width));
    shape.height = Math.max(24, asNumber(updates.height, existing.height));
    const nextFillColor = applyShapeBuilderFill(
      shape,
      asString(updates.fillColor, existing.fillColor),
      existing.fillColor
    );
    shape.lineFormat.color = sanitizeHexColor(
      asString(updates.outlineColor, existing.outlineColor),
      existing.outlineColor
    );
    shape.lineFormat.weight = Math.max(0, asNumber(updates.outlineWidth, existing.outlineWidth));
    shape.textFrame.textRange.font.color = sanitizeHexColor(
      asString(updates.fontColor, existing.fontColor),
      existing.fontColor
    );
    shape.textFrame.textRange.font.size = Math.max(
      8,
      asNumber(updates.fontSize, existing.fontSize)
    );
    shape.textFrame.textRange.font.bold = asBoolean(updates.bold, existing.bold);
    shape.textFrame.textRange.font.italic = asBoolean(updates.italic, existing.italic);
    shape.textFrame.horizontalAlignment = normalizeShapeBuilderTextHorizontalAlignment(
      asString(updates.textHorizontalAlignment, existing.textHorizontalAlignment)
    ) as unknown as Excel.ShapeTextHorizontalAlignment;
    shape.textFrame.verticalAlignment = normalizeShapeBuilderTextVerticalAlignment(
      asString(updates.textVerticalAlignment, existing.textVerticalAlignment)
    ) as unknown as Excel.ShapeTextVerticalAlignment;

    const nextMetadata: ShapeBuilderMetadata = {
      iconKey: asString(updates.iconKey, existing.iconKey).trim() || "none",
      linkType: normalizeShapeBuilderLinkType(asString(updates.linkType, existing.linkType)),
      internalDestinationId: asString(
        updates.internalDestinationId,
        existing.internalDestinationId
      ).trim(),
      externalUrl: asString(updates.externalUrl, existing.externalUrl).trim(),
      effect: normalizeShapeBuilderEffect(asString(updates.effect, existing.effect)),
      anchorAddress: asString(updates.anchorAddress, existing.anchorAddress).trim(),
    };
    if (nextMetadata.linkType !== "Internal") {
      nextMetadata.internalDestinationId = "";
    }
    if (nextMetadata.linkType !== "External") {
      nextMetadata.externalUrl = "";
    }

    const nextText = asString(updates.text, existing.text);
    shape.textFrame.textRange.text = buildShapeBuilderDisplayText(nextMetadata.iconKey, nextText);
    applyShapeBuilderMetadata(shape, nextMetadata);
    await syncShapeBuilderShadow(context, sheet, {
      shapeName,
      shapeType: nextShapeType,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      fillColor: nextFillColor,
      effect: nextMetadata.effect,
    });

    loadShapeBuilderProperties(shape);
    await context.sync();
    return buildShapeBuilderRecord(sheetName, shape);
  });
}

export async function formatShapeBuilderTextSelection(
  sheetName: string,
  shapeName: string,
  selectionStart: number,
  selectionLength: number,
  updates: ShapeBuilderTextSelectionFormatRequest
): Promise<ShapeBuilderShapeRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    loadShapeBuilderProperties(shape);
    await context.sync();

    const existing = buildShapeBuilderRecord(sheetName, shape);
    const textLength = existing.text.length;
    const safeStart = Math.min(Math.max(Math.floor(asNumber(selectionStart, 0)), 0), textLength);
    const requestedLength = Math.max(Math.floor(asNumber(selectionLength, 0)), 0);
    const safeLength = Math.min(requestedLength, Math.max(textLength - safeStart, 0));
    if (safeLength <= 0) {
      return existing;
    }

    const displayText = asString(shape.textFrame.textRange.text);
    const displayOffset = getShapeBuilderDisplayTextOffset(displayText, existing.iconKey);
    const displayStart = displayOffset + safeStart;
    const selection = shape.textFrame.textRange.getSubstring(displayStart, safeLength);

    if (typeof updates.bold === "boolean") {
      selection.font.bold = updates.bold;
    }
    if (typeof updates.italic === "boolean") {
      selection.font.italic = updates.italic;
    }
    if (typeof updates.color === "string" && updates.color.trim()) {
      selection.font.color = sanitizeHexColor(updates.color, existing.fontColor);
    }

    await context.sync();

    loadShapeBuilderProperties(shape);
    await context.sync();
    return buildShapeBuilderRecord(sheetName, shape);
  });
}

export async function deleteShapeBuilderShape(sheetName: string, shapeName: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    const shadowOrNull = sheet.shapes.getItemOrNullObject(getShapeBuilderShadowName(shapeName));
    shadowOrNull.load("name");
    await context.sync();
    if (!shadowOrNull.isNullObject) {
      shadowOrNull.delete();
    }
    shape.delete();
    await context.sync();
  });
}

export async function duplicateShapeBuilderShape(
  sheetName: string,
  shapeName: string
): Promise<ShapeBuilderShapeRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    loadShapeBuilderProperties(shape);
    await context.sync();

    const source = buildShapeBuilderRecord(sheetName, shape);
    let duplicated: Excel.Shape;
    if (source.shapeType === "TextBox") {
      duplicated = sheet.shapes.addTextBox("");
    } else {
      duplicated = sheet.shapes.addGeometricShape(toGeometricShapeType(source.shapeType));
    }

    duplicated.name = `${NAV_SHAPE_PREFIX}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    duplicated.left = source.left + 16;
    duplicated.top = source.top + 16;
    duplicated.width = source.width;
    duplicated.height = source.height;
    applyShapeBuilderFill(duplicated, source.fillColor, "#A8D5AD");
    duplicated.lineFormat.color = source.outlineColor;
    duplicated.lineFormat.weight = source.outlineWidth;
    duplicated.textFrame.textRange.font.color = source.fontColor;
    duplicated.textFrame.textRange.font.size = source.fontSize;
    duplicated.textFrame.textRange.font.bold = source.bold;
    duplicated.textFrame.textRange.font.italic = source.italic;
    duplicated.textFrame.horizontalAlignment =
      source.textHorizontalAlignment as unknown as Excel.ShapeTextHorizontalAlignment;
    duplicated.textFrame.verticalAlignment =
      source.textVerticalAlignment as unknown as Excel.ShapeTextVerticalAlignment;
    duplicated.textFrame.textRange.text = buildShapeBuilderDisplayText(source.iconKey, source.text);
    duplicated.placement = Excel.Placement.oneCell;
    applyShapeBuilderMetadata(duplicated, {
      iconKey: source.iconKey,
      linkType: source.linkType,
      internalDestinationId: source.internalDestinationId,
      externalUrl: source.externalUrl,
      effect: source.effect,
      anchorAddress: source.anchorAddress,
    });
    await syncShapeBuilderShadow(context, sheet, {
      shapeName: duplicated.name,
      shapeType: source.shapeType,
      left: duplicated.left,
      top: duplicated.top,
      width: duplicated.width,
      height: duplicated.height,
      fillColor: source.fillColor,
      effect: source.effect,
    });

    loadShapeBuilderProperties(duplicated);
    await context.sync();
    return buildShapeBuilderRecord(sheetName, duplicated);
  });
}

export async function getNamedRanges(): Promise<NamedRangeRecord[]> {
  return Excel.run(async (context) => {
    const records: NamedRangeRecord[] = [];
    const pendingShapeLoads: Array<{ recordIndex: number; range: Excel.Range }> = [];

    const workbookNames = context.workbook.names;
    workbookNames.load("items/name,items/type,items/formula");

    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    for (const namedItem of workbookNames.items) {
      if (namedItem.name.toLowerCase().startsWith("_xl")) {
        continue;
      }
      const normalizedFormula = normalizeDisplayAddress(namedItem.formula);
      const rawFormula = normalizeFormulaExpression(asString(namedItem.formula));
      const parsed = splitQualifiedAddress(namedItem.formula);
      const isRange =
        parsed.sheet.length > 0 &&
        parsed.address.length > 0 &&
        isCellAddressExpression(parsed.address);
      const resolvedSheet = parsed.sheet || (isRange ? "N/A" : "Workbook");
      records.push({
        id: `workbook::${namedItem.name}`,
        kind: "NamedRange",
        name: namedItem.name,
        address:
          isRange && parsed.sheet && parsed.address
            ? toQualifiedAddress(parsed.sheet, parsed.address)
            : normalizedFormula,
        formula: rawFormula,
        sheet: resolvedSheet,
        scope: "Workbook",
        scopeType: "Workbook",
        isRange,
        type: classifyNamedItemType(normalizedFormula, isRange),
      });

      if (isRange && parsed.sheet && parsed.address) {
        const range = context.workbook.worksheets.getItem(parsed.sheet).getRange(parsed.address);
        range.load("rowCount,columnCount");
        pendingShapeLoads.push({ recordIndex: records.length - 1, range });
      }
    }

    const worksheetArtifacts = worksheets.items.map((sheet) => {
      const collection = sheet.names;
      const shapes = sheet.shapes;
      collection.load("items/name,items/type,items/formula");
      shapes.load("items/name");
      return { sheetName: sheet.name, collection, shapes };
    });
    await context.sync();

    for (const map of worksheetArtifacts) {
      for (const namedItem of map.collection.items) {
        if (namedItem.name.toLowerCase().startsWith("_xl")) {
          continue;
        }
        const normalizedFormula = normalizeDisplayAddress(namedItem.formula);
        const rawFormula = normalizeFormulaExpression(asString(namedItem.formula));
        const parsed = splitQualifiedAddress(namedItem.formula);
        const isRange = parsed.address.length > 0 && isCellAddressExpression(parsed.address);
        const resolvedSheet = parsed.sheet || map.sheetName;
        records.push({
          id: `worksheet::${map.sheetName}::${namedItem.name}`,
          kind: "NamedRange",
          name: namedItem.name,
          address:
            isRange && resolvedSheet && parsed.address
              ? toQualifiedAddress(resolvedSheet, parsed.address)
              : normalizedFormula,
          formula: rawFormula,
          sheet: resolvedSheet,
          scope: map.sheetName,
          scopeType: "Worksheet",
          isRange,
          type: classifyNamedItemType(normalizedFormula, isRange),
        });

        if (isRange && parsed.address) {
          const range = context.workbook.worksheets
            .getItem(parsed.sheet || map.sheetName)
            .getRange(parsed.address);
          range.load("rowCount,columnCount");
          pendingShapeLoads.push({ recordIndex: records.length - 1, range });
        }
      }

      for (const shape of map.shapes.items) {
        records.push({
          id: `shape::${map.sheetName}::${shape.name}`,
          kind: "Shape",
          name: shape.name,
          address: toQualifiedAddress(map.sheetName, "[Shape]"),
          formula: "",
          sheet: map.sheetName,
          scope: map.sheetName,
          scopeType: "Worksheet",
          isRange: false,
          type: "Shape",
        });
      }
    }

    if (pendingShapeLoads.length > 0) {
      await context.sync();
      pendingShapeLoads.forEach(({ recordIndex, range }) => {
        records[recordIndex].type = classifyRangeShape(range.rowCount, range.columnCount);
      });
    }

    return records;
  });
}

export async function getCurrentSelectionAddress(): Promise<{ address: string; sheet: string }> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,worksheet/name");
    await context.sync();

    return {
      address: range.address,
      sheet: range.worksheet.name,
    };
  });
}

export async function activateWorksheet(sheetName: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRange("A1").select();
    await context.sync();
  });
}

export async function updateNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  oldName: string,
  newName: string,
  address: string,
  fallbackSheet: string,
  referenceType: "Reference" | "Formula" = "Reference"
) {
  const normalizedScope = scope.trim();
  const normalizedFallbackSheet = fallbackSheet.trim();
  if (scopeType === "Worksheet" && isInternalWorkbookSheet(normalizedScope)) {
    throw new Error("Worksheet-scoped names cannot target internal Workbook Manager sheets.");
  }
  if (
    referenceType === "Reference" &&
    !address.includes("!") &&
    normalizedFallbackSheet &&
    isInternalWorkbookSheet(normalizedFallbackSheet)
  ) {
    throw new Error("Select a workbook sheet before creating or updating a range reference.");
  }

  const candidates =
    referenceType === "Formula"
      ? buildArrayFormulaSeparatorCandidates(address)
      : [normalizeReference(address, normalizedFallbackSheet)];

  let lastError: unknown = null;
  for (const candidate of candidates) {
    try {
      await Excel.run(async (context) => {
        const collection =
          scopeType === "Workbook"
            ? context.workbook.names
            : context.workbook.worksheets.getItem(normalizedScope).names;

        if (oldName.toUpperCase() === newName.toUpperCase()) {
          const item = collection.getItem(oldName);
          item.formula = candidate;
          await context.sync();
          return;
        }

        collection.add(newName, candidate);
        collection.getItem(oldName).delete();
        await context.sync();
      });
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError instanceof Error ? lastError : new Error("Unable to update named range.");
}

export async function addNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  name: string,
  address: string,
  fallbackSheet: string,
  referenceType: "Reference" | "Formula" = "Reference"
) {
  const normalizedScope = scope.trim();
  const normalizedFallbackSheet = fallbackSheet.trim();
  if (scopeType === "Worksheet" && isInternalWorkbookSheet(normalizedScope)) {
    throw new Error("Worksheet-scoped names cannot target internal Workbook Manager sheets.");
  }
  if (
    referenceType === "Reference" &&
    !address.includes("!") &&
    normalizedFallbackSheet &&
    isInternalWorkbookSheet(normalizedFallbackSheet)
  ) {
    throw new Error("Select a workbook sheet before creating a range reference.");
  }

  const candidates =
    referenceType === "Formula"
      ? buildArrayFormulaSeparatorCandidates(address)
      : [normalizeReference(address, normalizedFallbackSheet)];

  let lastError: unknown = null;
  for (const candidate of candidates) {
    try {
      await Excel.run(async (context) => {
        const collection =
          scopeType === "Workbook"
            ? context.workbook.names
            : context.workbook.worksheets.getItem(normalizedScope).names;

        collection.add(name, candidate);
        await context.sync();
      });
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError instanceof Error ? lastError : new Error("Unable to add named range.");
}

export async function deleteNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  name: string
) {
  await Excel.run(async (context) => {
    const collection =
      scopeType === "Workbook"
        ? context.workbook.names
        : context.workbook.worksheets.getItem(scope).names;

    collection.getItem(name).delete();
    await context.sync();
  });
}

export async function moveNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  name: string,
  newAddress: string,
  fallbackSheet: string
) {
  await updateNamedRange(scopeType, scope, name, name, newAddress, fallbackSheet);
}

export async function deleteNamedRangeWithOptions(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  name: string,
  deleteName: boolean,
  deleteValues: boolean
) {
  await Excel.run(async (context) => {
    const collection =
      scopeType === "Workbook"
        ? context.workbook.names
        : context.workbook.worksheets.getItem(scope).names;

    const namedItem = collection.getItem(name);
    namedItem.load("formula");
    await context.sync();

    const parsed = splitQualifiedAddress(namedItem.formula);
    const targetSheet = parsed.sheet || (scopeType === "Worksheet" ? scope : "");

    if (deleteValues && targetSheet && parsed.address) {
      const range = context.workbook.worksheets.getItem(targetSheet).getRange(parsed.address);
      range.clear(Excel.ClearApplyTo.contents);
    }

    if (deleteName) {
      namedItem.delete();
    }

    await context.sync();
  });
}

export async function selectNamedRangeAddress(address: string, fallbackSheet: string) {
  await Excel.run(async (context) => {
    const parsed = getSheetAndAddress(address, fallbackSheet);
    const targetSheet = context.workbook.worksheets.getItem(parsed.sheet);
    targetSheet.activate();
    const range = targetSheet.getRange(parsed.address);
    range.select();
    await context.sync();
  });
}

export async function getTables(): Promise<TableRecord[]> {
  return Excel.run(async (context) => {
    const records: TableRecord[] = [];
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const tableMaps = worksheets.items.map((sheet) => {
      const tables = sheet.tables;
      tables.load("items/name");
      return { sheetName: sheet.name, tables };
    });
    await context.sync();

    const tableRanges = tableMaps.flatMap((map) =>
      map.tables.items.map((table) => {
        const range = table.getRange();
        range.load("address");
        return { sheetName: map.sheetName, table, range };
      })
    );
    await context.sync();

    tableRanges.forEach((item) => {
      const normalizedAddress = item.range.address.includes("!")
        ? item.range.address.split("!").slice(1).join("!")
        : item.range.address;
      records.push({
        id: `${item.sheetName}::${item.table.name}`,
        name: item.table.name,
        address: `${quoteSheetName(item.sheetName)}!${normalizedAddress}`,
        sheet: item.sheetName,
        scope: "Worksheet",
      });
    });

    await syncWorkbookTablesNamedRange(context, records);

    return records;
  });
}

export async function saveNamedFunction(name: string, lambdaFormula: string) {
  const trimmedName = name.trim();
  if (!trimmedName) {
    throw new Error("Function name is required.");
  }

  const normalized = lambdaFormula.trim();
  if (!normalized) {
    throw new Error("Formula is required.");
  }

  const formula = normalized.startsWith("=") ? normalized : `=${normalized}`;
  if (!/^=\s*LAMBDA\s*\(/i.test(formula)) {
    throw new Error("Named Function must use a LAMBDA formula.");
  }

  await Excel.run(async (context) => {
    const existing = context.workbook.names.getItemOrNullObject(trimmedName);
    existing.load("name");
    await context.sync();

    if (existing.isNullObject) {
      context.workbook.names.add(trimmedName, formula);
    } else {
      existing.formula = formula;
    }
    await context.sync();
  });
}

export async function listWorkbookQueries(): Promise<WorkbookQueryRecord[]> {
  if (!canUseExcelApi("1.14")) {
    return [];
  }

  return Excel.run(async (context) => {
    const queries = context.workbook.queries;
    queries.load(
      "items/name,items/loadedTo,items/loadedToDataModel,items/refreshDate,items/rowsLoadedCount,items/error"
    );
    await context.sync();

    return queries.items
      .map((query) => {
        const refreshDate =
          query.refreshDate instanceof Date && !Number.isNaN(query.refreshDate.valueOf())
            ? query.refreshDate.toLocaleString()
            : "";
        const error =
          query.error && query.error !== Excel.QueryError.none ? String(query.error) : "";
        return {
          name: query.name,
          loadedTo: String(query.loadedTo || ""),
          loadedToDataModel: Boolean(query.loadedToDataModel),
          refreshDate,
          rowsLoadedCount: query.rowsLoadedCount,
          error,
        } as WorkbookQueryRecord;
      })
      .sort((left, right) =>
        left.name.localeCompare(right.name, undefined, { sensitivity: "base" })
      );
  });
}

export async function refreshWorkbookQueries(): Promise<QueryRefreshResult> {
  if (!canUseExcelApi("1.7")) {
    return {
      refreshed: false,
      method: "Unsupported",
      warnings: ["ExcelApi 1.7 is unavailable. Cannot trigger workbook data connection refresh."],
    };
  }

  await Excel.run(async (context) => {
    context.workbook.dataConnections.refreshAll();
    await context.sync();
  });

  const warnings = [
    "Used DataConnections.refreshAll(). Excel JS does not expose a dedicated direct Power Query refresh API in this surface.",
  ];
  if (!canUseExcelApi("1.14")) {
    warnings.push("ExcelApi 1.14 query metadata is unavailable on this host.");
  }

  return {
    refreshed: true,
    method: "DataConnections.refreshAll",
    warnings,
  };
}

export async function applyModelBuilderParameters(
  request: ModelBuilderApplyRequest
): Promise<ModelBuilderApplyResult> {
  const parameters = request.parameters.filter(
    (item) =>
      item.label.trim() || item.desiredName.trim() || item.value.trim() || item.listValues.trim()
  );
  if (parameters.length === 0) {
    throw new Error("Add at least one parameter before applying.");
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(request.anchorSheet);
    const anchorAddress =
      splitQualifiedAddress(request.anchorAddress).address || request.anchorAddress;
    const anchorCell = sheet.getRange(anchorAddress).getCell(0, 0);
    anchorCell.load("rowIndex,columnIndex,address,worksheet/name");

    const workbookNames = context.workbook.names;
    workbookNames.load("items/name");

    const metaSheetName = "__WBM_META";
    const metaSheetOrNull = context.workbook.worksheets.getItemOrNullObject(metaSheetName);
    metaSheetOrNull.load("name");

    await context.sync();

    const metaSheet = metaSheetOrNull.isNullObject
      ? context.workbook.worksheets.add(metaSheetName)
      : (metaSheetOrNull as Excel.Worksheet);
    metaSheet.visibility = Excel.SheetVisibility.hidden;

    const metaUsedRange = metaSheet.getUsedRangeOrNullObject(true);
    metaUsedRange.load("columnCount");
    await context.sync();

    let listColumnCursor = metaUsedRange.isNullObject
      ? 3
      : Math.max(metaUsedRange.columnCount + 1, 3);
    const takenNames = new Set<string>(workbookNames.items.map((item) => item.name.toUpperCase()));
    const created: ModelBuilderApplyResult["created"] = [];

    parameters.forEach((parameter, index) => {
      const rowIndex = anchorCell.rowIndex + index;
      const labelCell = sheet.getRangeByIndexes(rowIndex, anchorCell.columnIndex, 1, 1);
      const valueCell = sheet.getRangeByIndexes(rowIndex, anchorCell.columnIndex + 1, 1, 1);

      const label = parameter.label.trim() || `Parameter ${index + 1}`;
      labelCell.values = [[label]];

      if (parameter.valueMode === "formula") {
        const normalizedFormula = normalizeFormulaExpression(parameter.value.trim() || "0");
        valueCell.formulas = [[normalizedFormula]];
      } else {
        valueCell.values = [[parameter.value]];
      }

      const preferredName = toModelSafeName(parameter.desiredName.trim() || label);
      const finalName = toUniqueName(preferredName, takenNames);
      const valueAddress = toAbsoluteCellAddress(rowIndex, anchorCell.columnIndex + 1);
      context.workbook.names.add(
        finalName,
        `=${quoteSheetName(anchorCell.worksheet.name)}!${valueAddress}`
      );

      let validationListName: string | undefined;
      const listValues = parseInlineListValues(parameter.listValues);
      if (listValues.length > 0) {
        validationListName = toUniqueName(toModelSafeName(`${finalName}_list`), takenNames);
        const listRange = metaSheet.getRangeByIndexes(0, listColumnCursor, listValues.length, 1);
        listRange.values = listValues.map((value) => [value]);

        const startAddress = toAbsoluteCellAddress(0, listColumnCursor);
        const endAddress = toAbsoluteCellAddress(listValues.length - 1, listColumnCursor);
        context.workbook.names.add(
          validationListName,
          `=${quoteSheetName(metaSheetName)}!${startAddress}:${endAddress}`
        );

        valueCell.dataValidation.clear();
        valueCell.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: `=${validationListName}`,
          },
        };
        listColumnCursor += 1;
      } else {
        valueCell.dataValidation.clear();
      }

      created.push({
        label,
        finalName,
        valueAddress: `${quoteSheetName(anchorCell.worksheet.name)}!${valueAddress}`,
        validationListName,
      });
    });

    await context.sync();

    return {
      anchorSheet: anchorCell.worksheet.name,
      anchorAddress: anchorCell.address,
      created,
    };
  });
}

export async function updateTableName(sheetName: string, oldName: string, newName: string) {
  await Excel.run(async (context) => {
    const table = context.workbook.worksheets.getItem(sheetName).tables.getItem(oldName);
    table.name = newName;

    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const tableMaps = worksheets.items.map((sheet) => {
      const tables = sheet.tables;
      tables.load("items/name");
      return { sheetName: sheet.name, tables };
    });
    await context.sync();

    const records: TableRecord[] = [];
    tableMaps.forEach((tableMap) => {
      tableMap.tables.items.forEach((item) => {
        records.push({
          id: `${tableMap.sheetName}::${item.name}`,
          name: item.name,
          address: "",
          sheet: tableMap.sheetName,
          scope: "Worksheet",
        });
      });
    });
    await syncWorkbookTablesNamedRange(context, records);

    await context.sync();
  });
}

export async function selectTableAddress(address: string, fallbackSheet: string) {
  await Excel.run(async (context) => {
    const parsed = getSheetAndAddress(address, fallbackSheet);
    const range = context.workbook.worksheets.getItem(parsed.sheet).getRange(parsed.address);
    range.select();
    await context.sync();
  });
}

export async function getTableColumns(
  sheetName: string,
  tableName: string
): Promise<TableColumnRecord[]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItem(tableName);
    table.columns.load("items/name");
    await context.sync();

    const ranges = table.columns.items.map((column) => {
      const bodyRange = column.getDataBodyRange();
      bodyRange.load("address");
      return { column, bodyRange };
    });
    await context.sync();

    return ranges.map((item) => ({
      id: `${sheetName}::${tableName}::${item.column.name}`,
      name: item.column.name,
      address: item.bodyRange.address,
    }));
  });
}

export async function createNamedRangesFromTableColumns(
  request: CreateNamedRangesFromTableRequest
): Promise<{ created: string[]; skipped: string[] }> {
  return Excel.run(async (context) => {
    const table = context.workbook.worksheets
      .getItem(request.sheetName)
      .tables.getItem(request.tableName);
    table.columns.load("items/name");

    const workbookNames = context.workbook.names;
    workbookNames.load("items/name");

    const sheetNames = context.workbook.worksheets.getItem(request.sheetName).names;
    sheetNames.load("items/name");
    await context.sync();

    const targetCollection = request.scopeType === "Workbook" ? workbookNames : sheetNames;
    const taken = new Set(targetCollection.items.map((item) => item.name.toUpperCase()));
    const created: string[] = [];
    const skipped: string[] = [];

    for (const columnName of request.columns) {
      const column = table.columns.getItem(columnName);
      const dataBodyRange = column.getDataBodyRange();
      dataBodyRange.load("address");
      await context.sync();

      const normalizedBase = toModelSafeName(columnName);
      const fallbackName = toModelSafeName(`${request.tableName}_${columnName}`);
      let desiredName = normalizedBase || fallbackName;
      const desiredUpper = desiredName.toUpperCase();
      if (taken.has(desiredUpper)) {
        const override = request.conflictValue.trim();
        if (request.conflictMode === "prefix" && override) {
          desiredName = toModelSafeName(`${override}${desiredName}`);
        } else if (request.conflictMode === "suffix" && override) {
          desiredName = toModelSafeName(`${desiredName}${override}`);
        } else if (request.conflictMode === "rename" && override) {
          desiredName = toModelSafeName(override);
        } else {
          skipped.push(columnName);
          continue;
        }
      }

      const finalName = toUniqueName(desiredName || fallbackName, taken);
      targetCollection.add(finalName, `=${dataBodyRange.address}`);
      taken.add(finalName.toUpperCase());
      created.push(finalName);
    }

    await context.sync();
    return { created, skipped };
  });
}

async function syncWorkbookTablesNamedRange(context: Excel.RequestContext, tables: TableRecord[]) {
  const metaSheetName = "__WBM_META";
  const metaSheetOrNull = context.workbook.worksheets.getItemOrNullObject(metaSheetName);
  metaSheetOrNull.load("name");
  await context.sync();

  const metaSheet = metaSheetOrNull.isNullObject
    ? context.workbook.worksheets.add(metaSheetName)
    : (metaSheetOrNull as Excel.Worksheet);
  metaSheet.visibility = Excel.SheetVisibility.hidden;

  const column = metaSheet.getRange("A:A");
  column.clear(Excel.ClearApplyTo.contents);

  const sortedNames = tables
    .map((table) => table.name.trim())
    .filter((name) => name.length > 0)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));

  if (sortedNames.length > 0) {
    const values = sortedNames.map((name) => [name]);
    const target = metaSheet.getRangeByIndexes(0, 0, values.length, 1);
    target.values = values;
  } else {
    metaSheet.getRange("A1").values = [[""]];
  }

  const wbTablesName = context.workbook.names.getItemOrNullObject("wb_Tables");
  wbTablesName.load("name");
  await context.sync();

  const targetAddress = sortedNames.length > 0 ? `=$A$1:$A$${sortedNames.length}` : "=$A$1";
  const formula = `=${quoteSheetName(metaSheetName)}!${targetAddress.slice(1)}`;

  if (wbTablesName.isNullObject) {
    context.workbook.names.add("wb_Tables", formula);
  } else {
    wbTablesName.formula = formula;
  }
}

export async function evaluateFormula(formulaText: string): Promise<FormulaEvaluationResult> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const evalSheetName = "__WBM_FUNCTION_EVAL";
    const evalSheet = workbook.worksheets.getItemOrNullObject(evalSheetName);
    evalSheet.load("name");
    await context.sync();

    let sheet: Excel.Worksheet = evalSheet as Excel.Worksheet;
    if (evalSheet.isNullObject) {
      sheet = workbook.worksheets.add(evalSheetName);
      sheet.visibility = Excel.SheetVisibility.hidden;
    }

    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("address");
    await context.sync();
    if (!usedRange.isNullObject) {
      usedRange.clear(Excel.ClearApplyTo.contents);
    }

    const target = sheet.getRange("A1");
    target.formulas = [[formulaText.startsWith("=") ? formulaText : `=${formulaText}`]];
    await context.sync();

    const output = sheet.getUsedRangeOrNullObject(true);
    output.load("address,rowCount,columnCount,values,valueTypes");
    await context.sync();

    if (output.isNullObject) {
      return { address: "A1", values: [[""]], valueTypes: [["Empty"]], hasError: false };
    }

    const values = output.values as (string | number | boolean | null)[][];
    const valueTypes = (output.valueTypes as Excel.RangeValueType[][]).map((row) =>
      row.map((item) => String(item))
    );
    const hasError = valueTypes.some((row) =>
      row.some((item) => item === Excel.RangeValueType.error || item === "Error")
    );
    const result: FormulaEvaluationResult = {
      address: output.address,
      values,
      valueTypes,
      hasError,
    };

    output.clear(Excel.ClearApplyTo.contents);
    await context.sync();
    return result;
  });
}

export async function getActiveCellFormulaState(): Promise<ActiveCellFormulaState> {
  return Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange().getCell(0, 0);
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    selection.load("address,formulas");
    worksheet.load("name");
    await context.sync();

    const formulaValue = String(selection.formulas?.[0]?.[0] ?? "");
    const parsed = splitQualifiedAddress(selection.address);
    return {
      sheet: worksheet.name,
      address: parsed.address || selection.address,
      formula: formulaValue,
      hasFormula: formulaValue.startsWith("="),
    };
  });
}

export async function applyFormulaToActiveCell(formulaText: string): Promise<void> {
  await Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange().getCell(0, 0);
    const normalized = formulaText.trim();
    selection.formulas = [[normalized.startsWith("=") ? normalized : `=${normalized}`]];
    await context.sync();
  });
}

export async function addShapeOnActiveCell(
  options: ShapeInsertOptions
): Promise<InsertedShapeRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const selection = context.workbook.getSelectedRange();
    selection.load("address,left,top,width,height");
    await context.sync();
    const selectionParsed = splitQualifiedAddress(selection.address);
    const anchorAddress = selectionParsed.address || selection.address;

    const shapeName = `WBM_${options.shapeType}_${Date.now().toString()}`;
    let shape: Excel.Shape;

    if (options.shapeType === "TextBox") {
      shape = sheet.shapes.addTextBox(options.text || "");
    } else {
      const geometricMap: Record<
        Exclude<InsertableShapeType, "TextBox">,
        Excel.GeometricShapeType
      > = {
        Rectangle: Excel.GeometricShapeType.rectangle,
        RoundedRectangle: "Round2SameRectangle" as unknown as Excel.GeometricShapeType,
        Chevron: "Chevron" as unknown as Excel.GeometricShapeType,
        Hexagon: "Hexagon" as unknown as Excel.GeometricShapeType,
        Diamond: "Diamond" as unknown as Excel.GeometricShapeType,
        Oval: "Oval" as unknown as Excel.GeometricShapeType,
      };
      const geometricType = geometricMap[options.shapeType] ?? Excel.GeometricShapeType.rectangle;
      shape = sheet.shapes.addGeometricShape(geometricType);
      shape.textFrame.textRange.text = options.text || "";
    }

    shape.name = shapeName;
    shape.left = selection.left;
    shape.top = selection.top;
    shape.width = Math.max(selection.width, 110);
    shape.height = Math.max(selection.height, 34);
    if (options.fillColor === NO_FILL_COLOR_TOKEN) {
      shape.fill.clear();
    } else {
      shape.fill.setSolidColor(options.fillColor);
    }
    shape.lineFormat.color = options.outlineColor;
    shape.lineFormat.weight = 1;
    shape.textFrame.textRange.font.color = options.fontColor;
    shape.placement = Excel.Placement.oneCell;
    shape.altTextDescription = `${SHAPE_ANCHOR_PREFIX}${anchorAddress}`;

    await context.sync();

    return {
      name: shapeName,
      shapeType: options.shapeType,
      sheet: sheet.name,
      anchorAddress,
      width: shape.width,
      height: shape.height,
    };
  });
}

export async function updateShapeFormatting(
  sheetName: string,
  shapeName: string,
  options: ShapeFormatOptions
) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    if (options.fillColor === NO_FILL_COLOR_TOKEN) {
      shape.fill.clear();
    } else {
      shape.fill.setSolidColor(options.fillColor);
    }
    shape.fill.transparency = Math.min(Math.max(options.fillTransparency, 0), 100);
    shape.lineFormat.color = options.outlineColor;
    shape.lineFormat.weight = Math.max(options.lineWeight, 0.25);
    shape.width = Math.max(options.width, 20);
    shape.height = Math.max(options.height, 20);
    shape.rotation = options.rotation;
    shape.textFrame.horizontalAlignment =
      options.textHorizontalAlignment as Excel.ShapeTextHorizontalAlignment;
    shape.textFrame.verticalAlignment =
      options.textVerticalAlignment as Excel.ShapeTextVerticalAlignment;
    shape.textFrame.textRange.font.color = options.fontColor;
    shape.textFrame.textRange.font.size = Math.max(options.fontSize, 6);
    shape.textFrame.textRange.font.bold = options.bold;
    shape.textFrame.textRange.font.italic = options.italic;
    shape.textFrame.textRange.text = options.text;
    shape.lockAspectRatio = options.lockAspectRatio;
    await context.sync();
  });
}

export async function renameShape(sheetName: string, oldName: string, newName: string) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(oldName);
    shape.name = newName;
    await context.sync();
  });
}

export async function getShapeEditorRecord(
  sheetName: string,
  shapeName: string
): Promise<ShapeEditorRecord> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    shape.load(
      "name,left,top,zOrderPosition,width,height,rotation,lockAspectRatio,geometricShapeType,altTextDescription,fill/foregroundColor,fill/transparency,lineFormat/color,lineFormat/weight,textFrame/horizontalAlignment,textFrame/verticalAlignment,textFrame/textRange/text,textFrame/textRange/font/color,textFrame/textRange/font/size,textFrame/textRange/font/bold,textFrame/textRange/font/italic"
    );
    await context.sync();

    const fillColor = shape.fill.foregroundColor ? shape.fill.foregroundColor : NO_FILL_COLOR_TOKEN;
    const anchorAddress = extractShapeAnchorAddress(shape.altTextDescription);

    return {
      name: shape.name,
      sheet: sheetName,
      shapeType: inferInsertableShapeType(shape.geometricShapeType),
      geometricShapeType: String(shape.geometricShapeType || ""),
      anchorAddress,
      left: shape.left || 0,
      top: shape.top || 0,
      zOrderPosition: shape.zOrderPosition || 0,
      fillColor,
      outlineColor: shape.lineFormat.color || "#000000",
      fontColor: shape.textFrame.textRange.font.color || "#1f1f1f",
      text: shape.textFrame.textRange.text || "",
      lineWeight: shape.lineFormat.weight || 1,
      fillTransparency: shape.fill.transparency || 0,
      width: shape.width || 110,
      height: shape.height || 34,
      rotation: shape.rotation || 0,
      textHorizontalAlignment:
        (shape.textFrame.horizontalAlignment as ShapeEditorRecord["textHorizontalAlignment"]) ||
        "Center",
      textVerticalAlignment:
        (shape.textFrame.verticalAlignment as ShapeEditorRecord["textVerticalAlignment"]) ||
        "Middle",
      fontSize: shape.textFrame.textRange.font.size || 11,
      bold: Boolean(shape.textFrame.textRange.font.bold),
      italic: Boolean(shape.textFrame.textRange.font.italic),
      lockAspectRatio: Boolean(shape.lockAspectRatio),
    };
  });
}

export async function updateShapePosition(
  sheetName: string,
  shapeName: string,
  options: ShapePositionOptions
) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    if (typeof options.left === "number") {
      shape.left = options.left;
    }
    if (typeof options.top === "number") {
      shape.top = options.top;
    }
    await context.sync();
  });
}

export async function nudgeShape(sheetName: string, shapeName: string, dx: number, dy: number) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    shape.load("left,top");
    await context.sync();
    shape.left = shape.left + dx;
    shape.top = shape.top + dy;
    await context.sync();
  });
}

export async function setShapeZOrder(
  sheetName: string,
  shapeName: string,
  order: "BringToFront" | "BringForward" | "SendToBack" | "SendBackward"
) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    const zOrderValue =
      order === "BringToFront"
        ? Excel.ShapeZOrder.bringToFront
        : order === "BringForward"
          ? Excel.ShapeZOrder.bringForward
          : order === "SendToBack"
            ? Excel.ShapeZOrder.sendToBack
            : Excel.ShapeZOrder.sendBackward;
    shape.setZOrder(zOrderValue);
    await context.sync();
  });
}

export async function setShapeGeometricType(
  sheetName: string,
  shapeName: string,
  shapeType: Exclude<InsertableShapeType, "TextBox">
) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    const geometricMap: Record<
      Exclude<InsertableShapeType, "TextBox">,
      Excel.GeometricShapeType
    > = {
      Rectangle: Excel.GeometricShapeType.rectangle,
      RoundedRectangle: "Round2SameRectangle" as unknown as Excel.GeometricShapeType,
      Chevron: "Chevron" as unknown as Excel.GeometricShapeType,
      Hexagon: "Hexagon" as unknown as Excel.GeometricShapeType,
      Diamond: "Diamond" as unknown as Excel.GeometricShapeType,
      Oval: "Oval" as unknown as Excel.GeometricShapeType,
    };
    (shape as unknown as { geometricShapeType: Excel.GeometricShapeType }).geometricShapeType =
      geometricMap[shapeType] ?? Excel.GeometricShapeType.rectangle;
    await context.sync();
  });
}

export async function alignShapeToSelection(
  sheetName: string,
  shapeName: string,
  alignment: "Left" | "Center" | "Right" | "Top" | "Middle" | "Bottom"
) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    const selection = context.workbook.getSelectedRange();
    shape.load("left,top,width,height");
    selection.load("left,top,width,height");
    await context.sync();

    const left = selection.left;
    const top = selection.top;
    const right = selection.left + selection.width;
    const bottom = selection.top + selection.height;

    if (alignment === "Left") {
      shape.left = left;
    } else if (alignment === "Center") {
      shape.left = left + (selection.width - shape.width) / 2;
    } else if (alignment === "Right") {
      shape.left = right - shape.width;
    } else if (alignment === "Top") {
      shape.top = top;
    } else if (alignment === "Middle") {
      shape.top = top + (selection.height - shape.height) / 2;
    } else if (alignment === "Bottom") {
      shape.top = bottom - shape.height;
    }

    await context.sync();
  });
}

export async function moveShapeToSelection(
  sheetName: string,
  shapeName: string,
  oldAnchorAddress: string
): Promise<{ anchorAddress: string }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shape = sheet.shapes.getItem(shapeName);
    const selected = context.workbook.getSelectedRange();
    const selectedCell = selected.getCell(0, 0);
    const oldParsed = splitQualifiedAddress(oldAnchorAddress);
    const oldAnchor = sheet.getRange(oldParsed.address || oldAnchorAddress).getCell(0, 0);

    selected.load("left,top");
    selectedCell.load("address");
    oldAnchor.load("address");
    await context.sync();

    shape.left = selected.left;
    shape.top = selected.top;

    const targetParsed = splitQualifiedAddress(selectedCell.address);
    const targetAddress = targetParsed.address || selectedCell.address;
    if (oldAnchor.address !== selectedCell.address) {
      selectedCell.copyFrom(oldAnchor, Excel.RangeCopyType.all);
      oldAnchor.clear(Excel.ClearApplyTo.contents);
    }
    shape.altTextDescription = `${SHAPE_ANCHOR_PREFIX}${targetAddress}`;

    await context.sync();
    return { anchorAddress: targetAddress };
  });
}

export async function getNamedRangeValueText(rangeName: string): Promise<string> {
  return Excel.run(async (context) => {
    const evalSheet = context.workbook.worksheets.getActiveWorksheet();
    const probe = evalSheet.getRange("XFD1048576");
    evalSheet.load("name");
    probe.formulas = [[`=${rangeName}`]];
    probe.load("values");
    await context.sync();
    const value = probe.values[0]?.[0];
    probe.clear(Excel.ClearApplyTo.contents);
    await context.sync();
    if (typeof value === "string" && value.startsWith("#")) {
      throw new Error(`Named range "${rangeName}" was not found or returned an error.`);
    }
    const asText = value === null || value === undefined ? "" : String(value);
    const maybeRef = parseAddressLikeText(asText, evalSheet.name);
    if (!maybeRef) {
      return asText;
    }
    const derefRange = context.workbook.worksheets
      .getItem(maybeRef.sheet)
      .getRange(maybeRef.address);
    derefRange.load("values");
    await context.sync();
    const derefValue = derefRange.values[0]?.[0];
    return derefValue === null || derefValue === undefined ? "" : String(derefValue);
  });
}

export async function applyFormulaToShapeAnchorCell(
  sheetName: string,
  anchorAddress: string,
  formulaText: string,
  shapeName: string
) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(anchorAddress);
    range.load("format/fill/color,values");
    await context.sync();

    const formula = formulaText.trim();
    range.formulas = [[formula.startsWith("=") ? formula : `=${formula}`]];
    await context.sync();

    range.load("format/fill/color,values");
    await context.sync();
    const fillColor = range.format.fill.color || "#ffffff";
    range.format.font.color = fillColor;

    const computed = range.values[0]?.[0];
    const shape = sheet.shapes.getItem(shapeName);
    shape.textFrame.textRange.text =
      computed === null || computed === undefined ? "" : String(computed);

    await context.sync();
  });
}

export async function openFormatEditorPopout(): Promise<void> {
  const url = `${window.location.origin}/taskpane.html?popout=format`;
  if (formatEditorDialog) {
    try {
      formatEditorDialog.close();
    } catch {
      // Ignore close errors.
    }
    formatEditorDialog = null;
  }

  await new Promise<void>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 85, width: 60, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error);
          return;
        }
        formatEditorDialog = result.value;
        formatEditorDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          formatEditorDialog = null;
        });
        resolve();
      }
    );
  });
}

export async function openFormulaEditorPopout(initialState?: {
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
}): Promise<void> {
  const url = `${window.location.origin}/taskpane.html?popout=formula`;
  if (formulaEditorDialog) {
    try {
      formulaEditorDialog.close();
    } catch {
      // Ignore close errors.
    }
    formulaEditorDialog = null;
  }

  await new Promise<void>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 85, width: 60, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error);
          return;
        }
        const dialog = result.value;
        formulaEditorDialog = dialog;
        let pendingOpenMessage: FormulaDialogOpenFormulaMessage | null =
          initialState && initialState.formula.trim()
            ? {
                channel: FORMULA_DIALOG_RPC_CHANNEL,
                type: "open-formula",
                formula: initialState.formula.trim(),
                name: initialState.name?.trim() || undefined,
                entryType: initialState.entryType,
                functionArgs: initialState.functionArgs,
                description: initialState.description,
                creationMode: initialState.creationMode,
                authoringMode: initialState.authoringMode,
                wizardTemplate: initialState.wizardTemplate,
                wizardArgs: initialState.wizardArgs,
                wizardReturnExpression: initialState.wizardReturnExpression,
                wizardVariables: initialState.wizardVariables,
              }
            : null;
        const dispatchOpenMessage = () => {
          if (!pendingOpenMessage) {
            return;
          }
          try {
            dialog.messageChild(JSON.stringify(pendingOpenMessage));
            pendingOpenMessage = null;
          } catch {
            // Dialog may not be ready to receive yet.
          }
        };
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (formulaEditorDialog === dialog) {
            formulaEditorDialog = null;
          }
        });
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args: { message: string; origin: string } | { error: number }) => {
            if (!("message" in args)) {
              return;
            }
            let payload: unknown;
            try {
              payload = JSON.parse(args.message);
            } catch {
              return;
            }

            if (!payload || typeof payload !== "object") {
              return;
            }

            const data = payload as Partial<
              FormulaDialogEvalRequestMessage | FormulaDialogReadyMessage
            >;
            if (data.channel !== FORMULA_DIALOG_RPC_CHANNEL || typeof data.type !== "string") {
              return;
            }
            if (data.type === "ready") {
              dispatchOpenMessage();
              return;
            }
            if (
              data.type !== "eval-request" ||
              typeof data.requestId !== "string" ||
              typeof data.formula !== "string"
            ) {
              return;
            }

            void (async () => {
              try {
                const evalResult = await evaluateFormula(data.formula);
                const response: FormulaDialogEvalResponseMessage = {
                  channel: FORMULA_DIALOG_RPC_CHANNEL,
                  type: "eval-response",
                  requestId: data.requestId,
                  ok: true,
                  result: evalResult,
                };
                try {
                  dialog.messageChild(JSON.stringify(response));
                } catch {
                  // Dialog may have been closed before response dispatch.
                }
              } catch (error) {
                const response: FormulaDialogEvalResponseMessage = {
                  channel: FORMULA_DIALOG_RPC_CHANNEL,
                  type: "eval-response",
                  requestId: data.requestId,
                  ok: false,
                  error: error instanceof Error ? error.message : String(error),
                };
                try {
                  dialog.messageChild(JSON.stringify(response));
                } catch {
                  // Dialog may have been closed before response dispatch.
                }
              }
            })();
          }
        );
        window.setTimeout(dispatchOpenMessage, 450);
        window.setTimeout(dispatchOpenMessage, 1100);
        resolve();
      }
    );
  });
}
