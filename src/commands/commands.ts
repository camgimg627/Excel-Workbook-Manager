/* global Office, OfficeRuntime */

const OPEN_FORMULA_EDITOR_SIGNAL_KEY = "wbm.openFormulaEditor.request";
const OPEN_FORMULA_EDITOR_ONLY_SIGNAL = "open-only";
const OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL = "open-and-pull";

async function setFormulaEditorSignal(value: string): Promise<void> {
  if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(OPEN_FORMULA_EDITOR_SIGNAL_KEY, value);
  }
  if (Office.context?.document?.settings) {
    Office.context.document.settings.set(OPEN_FORMULA_EDITOR_SIGNAL_KEY, value);
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

Office.onReady(() => {
  // Office.js is ready.
});

async function openFormulaEditorCommand(event: Office.AddinCommands.Event) {
  try {
    await setFormulaEditorSignal(OPEN_FORMULA_EDITOR_ONLY_SIGNAL);
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      await Office.addin.showAsTaskpane();
    }
  } catch {
    // Best-effort command; the task pane will still handle the signal on next open.
  } finally {
    event.completed();
  }
}

async function openAndPullFormulaEditorCommand(event: Office.AddinCommands.Event) {
  try {
    await setFormulaEditorSignal(OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL);
    if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      await Office.addin.showAsTaskpane();
    }
  } catch {
    // Best-effort command.
  } finally {
    event.completed();
  }
}

Office.actions.associate("openFormulaEditorCommand", openFormulaEditorCommand);
Office.actions.associate("openAndPullFormulaEditorCommand", openAndPullFormulaEditorCommand);
