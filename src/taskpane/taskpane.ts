/* global Excel console */

export interface NamedRangeRecord {
  id: string;
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
    | "Multi-Row Array";
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

export type CellStylePreset = "Input Cell" | "Parameter Cell" | "Header" | "Subheader";

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

    const sheetNameMaps = worksheets.items.map((sheet) => {
      const collection = sheet.names;
      collection.load("items/name,items/type,items/formula");
      return { sheetName: sheet.name, collection };
    });
    await context.sync();

    for (const map of sheetNameMaps) {
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
