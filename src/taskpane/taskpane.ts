/* global Excel console */

export interface NamedRangeRecord {
  id: string;
  kind: "NamedRange" | "Shape";
  name: string;
  address: string;
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
    | "Shape";
}

export interface TableRecord {
  id: string;
  name: string;
  address: string;
  sheet: string;
  scope: string;
}

export interface FormulaEvaluationResult {
  address: string;
  values: (string | number | boolean | null)[][];
}

export interface ActiveCellFormulaState {
  sheet: string;
  address: string;
  formula: string;
  hasFormula: boolean;
}

export type CellStylePreset = "Input Cell" | "Parameter Cell" | "Header" | "Subheader";
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

export interface ShapePositionOptions {
  left: number;
  top: number;
}

const NO_FILL_COLOR_TOKEN = "__NO_FILL__";
const SHAPE_ANCHOR_PREFIX = "WBM_ANCHOR=";

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
  return `${quoteSheetName(fallbackSheet)}!${cleaned}`;
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

function normalizeDisplayAddress(formula: string): string {
  const trimmed = formula.replace(/^=/, "").trim();
  if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replace(/""/g, '"').trim();
  }
  return trimmed;
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
      if (
        namedItem.name.includes("(") ||
        normalizeDisplayAddress(namedItem.formula).includes("(")
      ) {
        continue;
      }
      const parsed = splitQualifiedAddress(namedItem.formula);
      const isRange =
        parsed.sheet.length > 0 &&
        parsed.address.length > 0 &&
        isCellAddressExpression(parsed.address);
      const resolvedSheet = parsed.sheet || "N/A";
      records.push({
        id: `workbook::${namedItem.name}`,
        kind: "NamedRange",
        name: namedItem.name,
        address:
          parsed.sheet && parsed.address
            ? toQualifiedAddress(parsed.sheet, parsed.address)
            : parsed.address || normalizeDisplayAddress(namedItem.formula),
        sheet: resolvedSheet,
        scope: "Workbook",
        scopeType: "Workbook",
        isRange,
        type: "Single Cell",
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
        if (
          namedItem.name.includes("(") ||
          normalizeDisplayAddress(namedItem.formula).includes("(")
        ) {
          continue;
        }
        const parsed = splitQualifiedAddress(namedItem.formula);
        const isRange = parsed.address.length > 0 && isCellAddressExpression(parsed.address);
        const resolvedSheet = parsed.sheet || map.sheetName;
        records.push({
          id: `worksheet::${map.sheetName}::${namedItem.name}`,
          kind: "NamedRange",
          name: namedItem.name,
          address:
            resolvedSheet && parsed.address
              ? toQualifiedAddress(resolvedSheet, parsed.address)
              : parsed.address || normalizeDisplayAddress(namedItem.formula),
          sheet: resolvedSheet,
          scope: map.sheetName,
          scopeType: "Worksheet",
          isRange,
          type: "Single Cell",
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

export async function updateNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  oldName: string,
  newName: string,
  address: string,
  fallbackSheet: string
) {
  await Excel.run(async (context) => {
    const collection =
      scopeType === "Workbook"
        ? context.workbook.names
        : context.workbook.worksheets.getItem(scope).names;

    const oldItem = collection.getItem(oldName);
    oldItem.delete();
    collection.add(newName, normalizeReference(address, fallbackSheet));
    await context.sync();
  });
}

export async function addNamedRange(
  scopeType: "Workbook" | "Worksheet",
  scope: string,
  name: string,
  address: string,
  fallbackSheet: string
) {
  await Excel.run(async (context) => {
    const collection =
      scopeType === "Workbook"
        ? context.workbook.names
        : context.workbook.worksheets.getItem(scope).names;

    collection.add(name, normalizeReference(address, fallbackSheet));
    await context.sync();
  });
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
    const range = context.workbook.worksheets.getItem(parsed.sheet).getRange(parsed.address);
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

    return records;
  });
}

export async function updateTableName(sheetName: string, oldName: string, newName: string) {
  await Excel.run(async (context) => {
    const table = context.workbook.worksheets.getItem(sheetName).tables.getItem(oldName);
    table.name = newName;
    await context.sync();
  });
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
    output.load("address,rowCount,columnCount,values");
    await context.sync();

    if (output.isNullObject) {
      return { address: "A1", values: [[""]] };
    }

    const values = output.values as (string | number | boolean | null)[][];
    const result: FormulaEvaluationResult = {
      address: output.address,
      values,
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
    shape.left = Math.max(0, options.left);
    shape.top = Math.max(0, options.top);
    await context.sync();
  });
}

export async function nudgeShape(sheetName: string, shapeName: string, dx: number, dy: number) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    if (dx !== 0) {
      shape.incrementLeft(dx);
    }
    if (dy !== 0) {
      shape.incrementTop(dy);
    }
    await context.sync();
  });
}

export async function setShapeZOrder(
  sheetName: string,
  shapeName: string,
  position: "BringToFront" | "BringForward" | "SendToBack" | "SendBackward"
) {
  await Excel.run(async (context) => {
    const shape = context.workbook.worksheets.getItem(sheetName).shapes.getItem(shapeName);
    shape.setZOrder(position);
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
    const geometricMap: Record<Exclude<InsertableShapeType, "TextBox">, string> = {
      Rectangle: "Rectangle",
      RoundedRectangle: "Round2SameRectangle",
      Chevron: "Chevron",
      Hexagon: "Hexagon",
      Diamond: "Diamond",
      Oval: "Ellipse",
    };
    shape.geometricShapeType = geometricMap[shapeType] as unknown as Excel.GeometricShapeType;
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
    const selected = context.workbook.getSelectedRange();
    shape.load("left,top,width,height");
    selected.load("left,top,width,height");
    await context.sync();

    if (alignment === "Left") {
      shape.left = selected.left;
    } else if (alignment === "Center") {
      shape.left = selected.left + (selected.width - shape.width) / 2;
    } else if (alignment === "Right") {
      shape.left = selected.left + selected.width - shape.width;
    } else if (alignment === "Top") {
      shape.top = selected.top;
    } else if (alignment === "Middle") {
      shape.top = selected.top + (selected.height - shape.height) / 2;
    } else if (alignment === "Bottom") {
      shape.top = selected.top + selected.height - shape.height;
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
