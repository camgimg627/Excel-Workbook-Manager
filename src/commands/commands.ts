/* global Excel, Office, OfficeRuntime */
import { NAVIGATION_SIGNAL_KEY, NavigationTarget } from "../taskpane/navigation";
import { OPEN_FORMULA_EDITOR_SIGNAL_KEY } from "../shared/signals";
const OPEN_FORMULA_EDITOR_ONLY_SIGNAL = "open-only";
const OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL = "open-and-pull";
const FORMULA_PULL_SIGNAL = "formula-pull";
const FORMULA_APPLY_SIGNAL = "formula-apply";
const FORMULA_BEAUTIFY_SIGNAL = "formula-beautify";
const FORMULA_INSERT_SELECTION_SIGNAL = "formula-insert-selection";
export const WATCH_ADD_SIGNAL_KEY = "wbm.watch.addAddress";
export const WATCH_ADD_SHEET_KEY  = "wbm.watch.addSheet";
export const TABLE_CONTEXT_SIGNAL_KEY = "wbm.tables.context.action";
export const TABLE_CONTEXT_TABLE_NAME_KEY = "wbm.tables.context.tableName";
export const TABLE_CONTEXT_SHEET_NAME_KEY = "wbm.tables.context.sheetName";

type TableContextAction = "editTableName" | "createNamedRanges";

async function setSignal(key: string, value: string): Promise<void> {
  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(key, value);
  }
  if (Office.context?.document?.settings) {
    Office.context.document.settings.set(key, value);
    await new Promise<void>((resolve, reject) => {
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        reject(result.error);
      });
    });
  }
}

async function showTaskpane(): Promise<void> {
  if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
    await Office.addin.showAsTaskpane();
  }
}

async function executeNavigationCommand(
  event: Office.AddinCommands.Event,
  target: NavigationTarget
): Promise<void> {
  try {
    await setSignal(NAVIGATION_SIGNAL_KEY, target);
    await showTaskpane();
  } catch {
    // Best-effort command.
  } finally {
    event.completed();
  }
}

async function executeFormulaSignalCommand(
  event: Office.AddinCommands.Event,
  signal: string
): Promise<void> {
  try {
    // `saveAsync` on document settings is not safe to fire in parallel.
    await setSignal(OPEN_FORMULA_EDITOR_SIGNAL_KEY, signal);
    await setSignal(NAVIGATION_SIGNAL_KEY, "formulas");
    await showTaskpane();
  } catch {
    // Best-effort command.
  } finally {
    event.completed();
  }
}

Office.onReady(() => {
  // Office.js is ready.
});

async function openNamesCommand(event: Office.AddinCommands.Event) {
  await executeNavigationCommand(event, "names");
}

async function createNameCommand(event: Office.AddinCommands.Event) {
  await executeNavigationCommand(event, "names-create");
}

async function openFormattingCommand(event: Office.AddinCommands.Event) {
  await executeNavigationCommand(event, "format");
}

async function openSettingsCommand(event: Office.AddinCommands.Event) {
  await executeNavigationCommand(event, "settings");
}

async function openHelpCommand(event: Office.AddinCommands.Event) {
  await executeNavigationCommand(event, "help");
}

async function openFormulaEditorCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, OPEN_FORMULA_EDITOR_ONLY_SIGNAL);
}

async function openAndPullFormulaEditorCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL);
}

async function pullFormulaCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, FORMULA_PULL_SIGNAL);
}

async function applyFormulaCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, FORMULA_APPLY_SIGNAL);
}

async function beautifyFormulaCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, FORMULA_BEAUTIFY_SIGNAL);
}

async function insertSelectionFormulaCommand(event: Office.AddinCommands.Event) {
  await executeFormulaSignalCommand(event, FORMULA_INSERT_SELECTION_SIGNAL);
}

async function executeTableContextCommand(
  event: Office.AddinCommands.Event,
  action: TableContextAction
): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const selectedCell = context.workbook.getSelectedRange().getCell(0, 0);
      const sheet = selectedCell.worksheet;
      const tables = sheet.tables;
      sheet.load("name");
      tables.load("items/name");
      await context.sync();

      const matches = tables.items.map((table) => {
        const tableRange = table.getRange();
        const intersection = tableRange.getIntersectionOrNullObject(selectedCell);
        intersection.load("isNullObject");
        return { table, intersection };
      });
      await context.sync();

      const match = matches.find(({ intersection }) => !intersection.isNullObject);

      await setSignal(NAVIGATION_SIGNAL_KEY, "tables");
      if (!match) {
        await showTaskpane();
        return;
      }

      await setSignal(TABLE_CONTEXT_SIGNAL_KEY, action);
      await setSignal(TABLE_CONTEXT_TABLE_NAME_KEY, match.table.name as string);
      await setSignal(TABLE_CONTEXT_SHEET_NAME_KEY, sheet.name as string);
      await showTaskpane();
    });
  } catch {
    // Best-effort.
  } finally {
    event.completed();
  }
}

async function editTableNameCommand(event: Office.AddinCommands.Event): Promise<void> {
  await executeTableContextCommand(event, "editTableName");
}

async function createNamedRangesFromTableCommand(
  event: Office.AddinCommands.Event
): Promise<void> {
  await executeTableContextCommand(event, "createNamedRanges");
}

async function addToWatchWindowCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("address");
      const sheet = selectedRange.worksheet;
      sheet.load("name");
      await context.sync();

      const rawAddress = selectedRange.address as string;
      const bareAddress = rawAddress.includes("!")
        ? rawAddress.split("!").slice(1).join("!")
        : rawAddress;

      await setSignal(WATCH_ADD_SIGNAL_KEY, bareAddress);
      await setSignal(WATCH_ADD_SHEET_KEY, sheet.name as string);
      await setSignal(NAVIGATION_SIGNAL_KEY, "watch");
      await showTaskpane();
    });
  } catch {
    // Best-effort.
  } finally {
    event.completed();
  }
}

Office.actions.associate("openNamesCommand", openNamesCommand);
Office.actions.associate("createNameCommand", createNameCommand);
Office.actions.associate("openFormattingCommand", openFormattingCommand);
Office.actions.associate("openSettingsCommand", openSettingsCommand);
Office.actions.associate("openHelpCommand", openHelpCommand);
Office.actions.associate("openFormulaEditorCommand", openFormulaEditorCommand);
Office.actions.associate("openAndPullFormulaEditorCommand", openAndPullFormulaEditorCommand);
Office.actions.associate("pullFormulaCommand", pullFormulaCommand);
Office.actions.associate("applyFormulaCommand", applyFormulaCommand);
Office.actions.associate("beautifyFormulaCommand", beautifyFormulaCommand);
Office.actions.associate("insertSelectionFormulaCommand", insertSelectionFormulaCommand);
Office.actions.associate("editTableNameCommand", editTableNameCommand);
Office.actions.associate("createNamedRangesFromTableCommand", createNamedRangesFromTableCommand);
Office.actions.associate("addToWatchWindowCommand", addToWatchWindowCommand);
