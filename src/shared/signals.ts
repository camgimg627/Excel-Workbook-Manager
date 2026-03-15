export const Signals = {
  NAVIGATION_REQUEST: "wbm.navigation.request",
  FORMULA_EDITOR_REQUEST: "wbm.openFormulaEditor.request",
} as const;

export const NAVIGATION_SIGNAL_KEY = Signals.NAVIGATION_REQUEST;
export const OPEN_FORMULA_EDITOR_SIGNAL_KEY = Signals.FORMULA_EDITOR_REQUEST;
