/* global Excel console */

async function runFormattingCommand(command: (context: Excel.RequestContext) => Promise<void>) {
  await Excel.run(async (context) => {
    await command(context);
    await context.sync();
  });
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
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const pivotTables = sheet.pivotTables;
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
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    if (tables.items.length === 0) {
      throw new Error("No tables found on the active worksheet.");
    }

    tables.items[0].style = styleName;
  });
}
