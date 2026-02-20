/* global Excel */

export interface WorkbookNameRecord {
  name: string;
  value: string;
  type: string;
}

export interface WorkbookSheetRecord {
  name: string;
  position: number;
  visibility: string;
}

export interface NamedFormulaRecord {
  name: string;
  formula: string;
  value: string;
}

export async function getWorkbookNames(): Promise<WorkbookNameRecord[]> {
  return Excel.run(async (context) => {
    const names = context.workbook.names;
    names.load("items/name,type,value");
    await context.sync();

    return names.items.map((name) => ({
      name: name.name,
      type: name.type,
      value: String(name.value ?? ""),
    }));
  });
}

export async function getWorkbookSheets(): Promise<WorkbookSheetRecord[]> {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,position,visibility");
    await context.sync();

    return worksheets.items.map((sheet) => ({
      name: sheet.name,
      position: sheet.position,
      visibility: String(sheet.visibility),
    }));
  });
}

export async function getNamedFormulas(): Promise<NamedFormulaRecord[]> {
  return Excel.run(async (context) => {
    const names = context.workbook.names;
    names.load("items/name,formula,value");
    await context.sync();

    return names.items.map((name) => ({
      name: name.name,
      formula: String(name.formula ?? ""),
      value: String(name.value ?? ""),
    }));
  });
}
