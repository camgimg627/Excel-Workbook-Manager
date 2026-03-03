import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import { Button, makeStyles } from "@fluentui/react-components";
import LegacyApp from "./LegacyApp";
import ModernShell from "./modern/ModernShell";
import FormulaMonacoView, { FormulaViewHandle } from "./modern/FormulaMonacoView";
import FormatView from "./modern/FormatView";
import { NAVIGATION_SIGNAL_KEY, NavigationTarget, isNavigationTarget } from "../navigation";

/* global Office, OfficeRuntime */

type UiMode = "legacy" | "modern";

const OPEN_FORMULA_EDITOR_SIGNAL_KEY = "wbm.openFormulaEditor.request";
const OPEN_FORMULA_EDITOR_ONLY_SIGNAL = "open-only";
const OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL = "open-and-pull";
const FORMULA_PULL_SIGNAL = "formula-pull";
const FORMULA_APPLY_SIGNAL = "formula-apply";
const FORMULA_BEAUTIFY_SIGNAL = "formula-beautify";
const FORMULA_INSERT_SELECTION_SIGNAL = "formula-insert-selection";

const UI_MODE_STORAGE_KEY = "wbm.ui.mode";

const useStyles = makeStyles({
  root: {
    position: "relative",
    minHeight: "100vh",
  },
  previewToggle: {
    position: "fixed",
    top: "8px",
    right: "8px",
    zIndex: 9999,
  },
});

const getQueryParam = (key: string): string | null => {
  if (typeof window === "undefined") {
    return null;
  }
  return new URLSearchParams(window.location.search).get(key);
};

const getInitialMode = (): UiMode => {
  const fromQuery = getQueryParam("ui");
  if (fromQuery === "modern" || fromQuery === "legacy") {
    return fromQuery;
  }
  if (typeof window === "undefined") {
    return "modern";
  }
  const fromStorage = window.localStorage.getItem(UI_MODE_STORAGE_KEY);
  return fromStorage === "legacy" ? "legacy" : "modern";
};

const readDocumentSignal = (key: string): string | null => {
  try {
    return (Office.context.document.settings.get(key) as string) || null;
  } catch {
    return null;
  }
};

const clearDocumentSignal = async (key: string): Promise<void> => {
  try {
    Office.context.document.settings.remove(key);
    await new Promise<void>((resolve) => {
      Office.context.document.settings.saveAsync(() => resolve());
    });
  } catch {
    // best-effort
  }
};

const App: React.FC = () => {
  const styles = useStyles();
  const popoutMode = getQueryParam("popout");
  const isFormulaPopout = popoutMode === "formula";
  const isFormatPopout = popoutMode === "format";
  const [uiMode, setUiMode] = useState<UiMode>(getInitialMode());
  const [activeTarget, setActiveTarget] = useState<NavigationTarget>("names");
  const [createRequestId, setCreateRequestId] = useState<number>(0);
  const formulaViewRef = useRef<FormulaViewHandle>(null);

  const resetUiPreference = () => {
    if (typeof window !== "undefined") {
      window.localStorage.removeItem(UI_MODE_STORAGE_KEY);
    }
    setUiMode("modern");
  };

  useEffect(() => {
    if (typeof window !== "undefined") {
      window.localStorage.setItem(UI_MODE_STORAGE_KEY, uiMode);
    }
  }, [uiMode]);

  useEffect(() => {
    let disposed = false;
    const runtime = (window as unknown as { OfficeRuntime?: typeof OfficeRuntime }).OfficeRuntime;

    const consumeNavigationSignal = async () => {
      try {
        const runtimeSignal = runtime?.storage
          ? await runtime.storage.getItem(NAVIGATION_SIGNAL_KEY)
          : null;
        const documentSignal = readDocumentSignal(NAVIGATION_SIGNAL_KEY);
        const signal = runtimeSignal || documentSignal;
        if (!signal || disposed) {
          return;
        }
        if (runtime?.storage) {
          await runtime.storage.removeItem(NAVIGATION_SIGNAL_KEY);
        }
        await clearDocumentSignal(NAVIGATION_SIGNAL_KEY);
        if (!isNavigationTarget(signal)) {
          return;
        }
        if (signal === "names-create") {
          setActiveTarget("names");
          setCreateRequestId((prev) => prev + 1);
          return;
        }
        setActiveTarget(signal);
      } catch {
        // ignore storage/read errors
      }
    };

    void consumeNavigationSignal();
    const timerId = window.setInterval(() => {
      void consumeNavigationSignal();
    }, 900);

    return () => {
      disposed = true;
      window.clearInterval(timerId);
    };
  }, []);

  useEffect(() => {
    if (uiMode !== "modern") {
      return undefined;
    }

    let disposed = false;
    const runtime = (window as unknown as { OfficeRuntime?: typeof OfficeRuntime }).OfficeRuntime;

    const consumeFormulaSignal = async () => {
      try {
        const runtimeSignal = runtime?.storage
          ? await runtime.storage.getItem(OPEN_FORMULA_EDITOR_SIGNAL_KEY)
          : null;
        const documentSignal = readDocumentSignal(OPEN_FORMULA_EDITOR_SIGNAL_KEY);
        const signal = runtimeSignal || documentSignal;
        if (!signal || disposed) {
          return;
        }
        if (runtime?.storage) {
          await runtime.storage.removeItem(OPEN_FORMULA_EDITOR_SIGNAL_KEY);
        }
        await clearDocumentSignal(OPEN_FORMULA_EDITOR_SIGNAL_KEY);
        setActiveTarget("formulas");

        const invoke = async (attempt = 0): Promise<void> => {
          if (disposed) {
            return;
          }
          const view = formulaViewRef.current;
          if (!view) {
            if (attempt < 20) {
              window.setTimeout(() => void invoke(attempt + 1), 80);
            }
            return;
          }
          if (signal === OPEN_FORMULA_EDITOR_ONLY_SIGNAL) {
            await view.openOnly();
            return;
          }
          if (signal === OPEN_FORMULA_EDITOR_AND_PULL_SIGNAL) {
            await view.openAndPull();
            return;
          }
          if (signal === FORMULA_PULL_SIGNAL) {
            await view.pull();
            return;
          }
          if (signal === FORMULA_APPLY_SIGNAL) {
            await view.apply();
            return;
          }
          if (signal === FORMULA_BEAUTIFY_SIGNAL) {
            await view.beautify();
            return;
          }
          if (signal === FORMULA_INSERT_SELECTION_SIGNAL) {
            await view.insertSelection();
          }
        };

        await invoke();
      } catch {
        // ignore signal/read errors
      }
    };

    void consumeFormulaSignal();
    const timerId = window.setInterval(() => {
      void consumeFormulaSignal();
    }, 900);

    return () => {
      disposed = true;
      window.clearInterval(timerId);
    };
  }, [uiMode]);

  const legacyInitialTarget = useMemo<NavigationTarget>(() => {
    if (isFormulaPopout) {
      return "formulas";
    }
    if (isFormatPopout) {
      return "format";
    }
    return activeTarget;
  }, [activeTarget, isFormatPopout, isFormulaPopout]);

  if (isFormulaPopout && uiMode === "modern") {
    return (
      <div className={styles.root}>
        <FormulaMonacoView ref={formulaViewRef} isPopout onOpenLegacy={() => setUiMode("legacy")} />
      </div>
    );
  }

  if (isFormatPopout && uiMode === "modern") {
    return (
      <div className={styles.root}>
        <FormatView isPopout onOpenLegacy={() => setUiMode("legacy")} />
      </div>
    );
  }

  if (uiMode === "legacy") {
    return (
      <div className={styles.root}>
        {!isFormulaPopout ? (
          <div className={styles.previewToggle}>
            <Button size="small" onClick={() => setUiMode("modern")}>
              Return To Modern
            </Button>
          </div>
        ) : null}
        <LegacyApp initialTarget={legacyInitialTarget} />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <ModernShell
        activeTarget={activeTarget}
        createRequestId={createRequestId}
        onTargetChange={setActiveTarget}
        onSwitchToLegacy={() => setUiMode("legacy")}
        onResetUiPreference={resetUiPreference}
        formulaViewRef={formulaViewRef}
      />
    </div>
  );
};

export default App;
