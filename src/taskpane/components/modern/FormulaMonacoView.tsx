import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { Button, Input, Text, makeStyles } from "@fluentui/react-components";
import Editor, { OnMount, loader } from "@monaco-editor/react";
import type * as Monaco from "monaco-editor";
import {
  FormulaEvaluationResult,
  TableRecord,
  applyFormulaToActiveCell,
  evaluateFormula,
  getActiveCellFormulaState,
  getCurrentSelectionAddress,
  getNamedRanges,
  getTables,
  openFormulaEditorPopout,
  saveNamedFunction,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

/* global Office */

const MONACO_LANGUAGE_ID = "wbm-formula";
const DIALOG_RPC_CHANNEL = "wbm-formula-dialog-rpc";

interface DialogEvalRequestMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "eval-request";
  requestId: string;
  formula: string;
}

interface DialogEvalResponseMessage {
  channel: typeof DIALOG_RPC_CHANNEL;
  type: "eval-response";
  requestId: string;
  ok: boolean;
  result?: FormulaEvaluationResult;
  error?: string;
}

let monacoLoaderConfigured = false;
const configureMonacoLoader = () => {
  if (monacoLoaderConfigured || typeof window === "undefined") {
    return;
  }
  // Keep Monaco local to the add-in origin so Office host CSP/AppDomain rules do not block editor boot.
  loader.config({
    paths: {
      vs: `${window.location.origin}/monaco/vs`,
    },
  });
  monacoLoaderConfigured = true;
};
configureMonacoLoader();

const FUNCTION_SUGGESTIONS = [
  "ABS",
  "AVERAGE",
  "CHOOSE",
  "COUNT",
  "FILTER",
  "IF",
  "IFS",
  "INDEX",
  "INDIRECT",
  "LAMBDA",
  "LET",
  "LOOKUP",
  "MATCH",
  "MAX",
  "MIN",
  "OFFSET",
  "SEQUENCE",
  "SORT",
  "SORTBY",
  "SUBTOTAL",
  "SUM",
  "SUMIF",
  "SUMIFS",
  "SUMPRODUCT",
  "SWITCH",
  "TEXT",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTJOIN",
  "TEXTSPLIT",
  "UNIQUE",
  "VLOOKUP",
  "XLOOKUP",
  "XMATCH",
] as const;

interface FormulaMonacoViewProps {
  isPopout: boolean;
  onOpenLegacy: () => void;
}

export interface FormulaViewHandle {
  openOnly: () => Promise<void>;
  openAndPull: () => Promise<void>;
  pull: () => Promise<void>;
  apply: () => Promise<void>;
  beautify: () => Promise<void>;
  insertSelection: () => Promise<void>;
}

interface ActiveCellState {
  sheet: string;
  address: string;
  formula: string;
  hasFormula: boolean;
}

const useStyles = makeStyles({
  root: { display: "grid", gap: "16px" },
  card: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    display: "grid",
    gap: "12px",
  },
  editorShell: {
    position: "relative",
  },
  overlayTop: {
    position: "absolute",
    top: "8px",
    right: "8px",
    zIndex: 6,
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  overlayBottom: {
    position: "absolute",
    left: "8px",
    right: "8px",
    bottom: "8px",
    zIndex: 6,
    display: "grid",
    gap: "8px",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    "@media (max-width: 780px)": {
      gridTemplateColumns: "minmax(0, 1fr)",
    },
  },
  overlayCluster: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "rgba(255, 255, 255, 0.95)",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "8px",
    display: "grid",
    gap: "6px",
  },
  overlayTitle: {
    fontSize: "11px",
    fontWeight: 700,
    color: MODERN_TOKENS.colorTextMuted,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  actionRow: {
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
    alignItems: "center",
  },
  editorWrap: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    overflow: "hidden",
    minHeight: "220px",
    backgroundColor: "#fff",
  },
  editorLoading: {
    minHeight: "220px",
    display: "grid",
    alignItems: "center",
    justifyItems: "center",
    backgroundColor: "#fff",
  },
  fallbackEditor: {
    width: "100%",
    minHeight: "220px",
    border: "none",
    outline: "none",
    resize: "vertical",
    padding: "42px 12px 128px",
    boxSizing: "border-box",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "13px",
    lineHeight: "1.4",
    color: MODERN_TOKENS.colorText,
    backgroundColor: "#fff",
  },
  liveTestHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  outputWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "6px",
    overflow: "auto",
    backgroundColor: "#fff",
    minHeight: "84px",
  },
  outputTable: { width: "100%", borderCollapse: "collapse", fontSize: "12px" },
  outputCell: { borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`, padding: "8px 10px" },
  autoCaptureRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "12px",
    flexWrap: "wrap",
  },
  checkboxLabel: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
  },
  modalBackdrop: {
    position: "fixed",
    inset: 0,
    backgroundColor: "rgba(15, 23, 42, 0.24)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 60,
    padding: "16px",
  },
  modal: {
    width: "min(520px, 100%)",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "20px",
    display: "grid",
    gap: "10px",
  },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: "8px", flexWrap: "wrap" },
});

const normalizeError = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

const prettyFormatFormula = (input: string): string => {
  const trimmed = input.trim();
  if (!trimmed) return "";
  const formula = trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
  let depth = 0;
  let inString = false;
  let out = "";
  for (let i = 0; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"') {
      inString = !inString;
      out += ch;
      continue;
    }
    if (inString) {
      out += ch;
      continue;
    }
    if (ch === "(") {
      depth += 1;
      out += "(\n" + "  ".repeat(depth);
      continue;
    }
    if (ch === ")") {
      depth = Math.max(0, depth - 1);
      out += "\n" + "  ".repeat(depth) + ")";
      continue;
    }
    if (ch === ",") {
      out += ",\n" + "  ".repeat(depth);
      continue;
    }
    out += ch;
  }
  return out;
};

const getTokenBeforeCursor = (text: string, cursorOffset: number): string => {
  const before = text.slice(0, cursorOffset);
  const tokenMatch = before.match(/([A-Za-z_\\][A-Za-z0-9_.\\]*)$/);
  return tokenMatch ? tokenMatch[1] : "";
};

const collectLetLambdaLocals = (text: string, cursorOffset: number): string[] => {
  interface ParseFrame {
    name: string;
    args: string[];
    current: string;
  }

  const isIdentifier = (value: string) => /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(value.trim());
  const beforeCursor = text.slice(0, cursorOffset);
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

  const locals: string[] = [];
  const addLocal = (candidate: string) => {
    const cleaned = candidate.trim();
    if (!isIdentifier(cleaned) || cleaned.includes(".")) {
      return;
    }
    if (!locals.some((item) => item.toUpperCase() === cleaned.toUpperCase())) {
      locals.push(cleaned);
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

  return locals;
};

const FormulaMonacoView = React.forwardRef<FormulaViewHandle, FormulaMonacoViewProps>(
  ({ isPopout, onOpenLegacy }, ref) => {
    const shared = useModernSharedStyles();
    const styles = useStyles();
    const [formulaText, setFormulaText] = useState<string>("");
    const [status, setStatus] = useState<string>("");
    const [statusType, setStatusType] = useState<"success" | "error">("success");
    const [activeCell, setActiveCell] = useState<ActiveCellState | null>(null);
    const [formulaOutput, setFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
    const [autoCapture, setAutoCapture] = useState<boolean>(true);
    const [loadingMetadata, setLoadingMetadata] = useState<boolean>(false);
    const [editorReady, setEditorReady] = useState<boolean>(false);
    const [editorLoadError, setEditorLoadError] = useState<string>("");
    const [showSaveFunctionModal, setShowSaveFunctionModal] = useState<boolean>(false);
    const [namedFunctionName, setNamedFunctionName] = useState<string>("");
    const [liveTestBusy, setLiveTestBusy] = useState<boolean>(false);
    const [liveTestError, setLiveTestError] = useState<string>("");
    const [lastTestedAt, setLastTestedAt] = useState<string>("");
    const [isEditorExpanded, setIsEditorExpanded] = useState<boolean>(false);

    const editorRef = useRef<Monaco.editor.IStandaloneCodeEditor | null>(null);
    const monacoRef = useRef<typeof Monaco | null>(null);
    const providerRef = useRef<Monaco.IDisposable | null>(null);
    const namesRef = useRef<string[]>([]);
    const tablesRef = useRef<string[]>([]);
    const pendingEvalRequestsRef = useRef<
      Map<string, { resolve: (result: FormulaEvaluationResult) => void; reject: (error: Error) => void; timeoutId: number }>
    >(new Map());

    const statusClass = statusType === "success" ? shared.successText : shared.errorText;

    useEffect(() => {
      let disposed = false;
      void loader
        .init()
        .then(() => {
          if (!disposed) {
            setEditorReady(true);
          }
        })
        .catch((error) => {
          if (disposed) {
            return;
          }
          const message = normalizeError(error);
          setEditorLoadError(message);
          setStatusType("error");
          setStatus(`Formula editor failed to initialize. Fallback mode enabled: ${message}`);
        });

      return () => {
        disposed = true;
      };
    }, []);

    const runAction = async (label: string, action: () => Promise<void>) => {
      setStatus("");
      try {
        await action();
        setStatusType("success");
        setStatus(`${label} completed.`);
      } catch (error) {
        setStatusType("error");
        setStatus(`${label} failed: ${normalizeError(error)}`);
      }
    };

    const pullFromActiveCell = useCallback(
      async (requireFormula: boolean) => {
        const state = await getActiveCellFormulaState();
        setActiveCell(state);
        if (!state.hasFormula) {
          if (requireFormula) {
            throw new Error(`Cell ${state.sheet}!${state.address} does not contain a formula.`);
          }
          return false;
        }
        setFormulaText(state.formula);
        return true;
      },
      []
    );

    const applyToActiveCell = useCallback(async () => {
      const normalized = formulaText.trim();
      if (!normalized) {
        throw new Error("Formula editor is empty.");
      }
      await applyFormulaToActiveCell(normalized);
      const refreshed = await getActiveCellFormulaState();
      setActiveCell(refreshed);
    }, [formulaText]);

    const insertSelection = useCallback(async () => {
      const selection = await getCurrentSelectionAddress();
      const token = selection.address;
      const current = editorRef.current?.getValue() ?? formulaText;
      if (!editorRef.current) {
        setFormulaText(current ? `${current}${token}` : token);
        return;
      }
      const selectionRange = editorRef.current.getSelection();
      if (!selectionRange) {
        setFormulaText(current ? `${current}${token}` : token);
        return;
      }
      editorRef.current.executeEdits("wbm-insert-selection", [
        {
          range: selectionRange,
          text: token,
          forceMoveMarkers: true,
        },
      ]);
      setFormulaText(editorRef.current.getValue());
    }, [formulaText]);

    const beautify = useCallback(async () => {
      setFormulaText((prev) => prettyFormatFormula(prev));
    }, []);

    const loadSuggestions = useCallback(async () => {
      setLoadingMetadata(true);
      try {
        const [namedRanges, tables] = await Promise.all([getNamedRanges(), getTables()]);
        namesRef.current = namedRanges
          .filter((item) => item.kind === "NamedRange")
          .map((item) => item.name);
        tablesRef.current = tables.map((item: TableRecord) => item.name);
      } finally {
        setLoadingMetadata(false);
      }
    }, []);

    useEffect(() => {
      void loadSuggestions();
    }, [loadSuggestions]);

    useEffect(() => {
      if (!autoCapture) return undefined;
      let disposed = false;
      let busy = false;
      const timerId = window.setInterval(() => {
        if (disposed || busy) return;
        busy = true;
        void (async () => {
          try {
            const state = await getActiveCellFormulaState();
            if (disposed) return;
            setActiveCell((prev) => {
              if (
                prev &&
                prev.sheet === state.sheet &&
                prev.address === state.address &&
                prev.formula === state.formula &&
                prev.hasFormula === state.hasFormula
              ) {
                return prev;
              }
              return state;
            });
            if (!state.hasFormula) return;
            const editorFocused = editorRef.current?.hasTextFocus() ?? false;
            if (!editorFocused && state.formula !== formulaText) {
              setFormulaText(state.formula);
            }
          } catch {
            // ignore transient excel editing states
          } finally {
            busy = false;
          }
        })();
      }, 1100);

      return () => {
        disposed = true;
        window.clearInterval(timerId);
      };
    }, [autoCapture, formulaText]);

    useEffect(() => {
      if (!isPopout || typeof Office === "undefined") {
        return undefined;
      }

      const handler = (args: Office.DialogParentMessageReceivedEventArgs) => {
        let payload: unknown;
        try {
          payload = JSON.parse(args.message);
        } catch {
          return;
        }

        if (!payload || typeof payload !== "object") {
          return;
        }

        const data = payload as Partial<DialogEvalResponseMessage>;
        if (data.channel !== DIALOG_RPC_CHANNEL || data.type !== "eval-response" || typeof data.requestId !== "string") {
          return;
        }

        const pending = pendingEvalRequestsRef.current.get(data.requestId);
        if (!pending) {
          return;
        }

        pendingEvalRequestsRef.current.delete(data.requestId);
        window.clearTimeout(pending.timeoutId);

        if (data.ok && data.result) {
          pending.resolve(data.result);
          return;
        }

        pending.reject(new Error(typeof data.error === "string" ? data.error : "Live test failed."));
      };

      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, handler);

      return () => {
        pendingEvalRequestsRef.current.forEach((pending) => {
          window.clearTimeout(pending.timeoutId);
          pending.reject(new Error("Dialog listener was reset."));
        });
        pendingEvalRequestsRef.current.clear();
      };
    }, [isPopout]);

    const evaluateFormulaForLive = useCallback(
      async (normalizedFormula: string): Promise<FormulaEvaluationResult> => {
        if (!isPopout) {
          return evaluateFormula(normalizedFormula);
        }
        if (typeof Office === "undefined" || !Office.context.ui?.messageParent) {
          throw new Error("Dialog bridge is unavailable for live testing.");
        }

        return await new Promise<FormulaEvaluationResult>((resolve, reject) => {
          const requestId = `eval-${Date.now()}-${Math.random().toString(16).slice(2)}`;
          const timeoutId = window.setTimeout(() => {
            pendingEvalRequestsRef.current.delete(requestId);
            reject(new Error("Live test timed out waiting for task pane response."));
          }, 10000);

          pendingEvalRequestsRef.current.set(requestId, { resolve, reject, timeoutId });

          const request: DialogEvalRequestMessage = {
            channel: DIALOG_RPC_CHANNEL,
            type: "eval-request",
            requestId,
            formula: normalizedFormula,
          };

          try {
            Office.context.ui.messageParent(JSON.stringify(request));
          } catch (error) {
            pendingEvalRequestsRef.current.delete(requestId);
            window.clearTimeout(timeoutId);
            reject(new Error(normalizeError(error)));
          }
        });
      },
      [isPopout]
    );

    useEffect(() => {
      const normalized = formulaText.trim();
      if (!normalized) {
        setFormulaOutput(null);
        setLiveTestError("");
        setLiveTestBusy(false);
        setLastTestedAt("");
        return undefined;
      }

      let disposed = false;
      const timerId = window.setTimeout(() => {
        setLiveTestBusy(true);
        setLiveTestError("");
        void evaluateFormulaForLive(normalized)
          .then((result) => {
            if (disposed) {
              return;
            }
            setFormulaOutput(result);
            setLastTestedAt(new Date().toLocaleTimeString());
          })
          .catch((error) => {
            if (disposed) {
              return;
            }
            setFormulaOutput(null);
            setLiveTestError(normalizeError(error));
            setLastTestedAt(new Date().toLocaleTimeString());
          })
          .finally(() => {
            if (!disposed) {
              setLiveTestBusy(false);
            }
          });
      }, 320);

      return () => {
        disposed = true;
        window.clearTimeout(timerId);
      };
    }, [evaluateFormulaForLive, formulaText]);

    const registerCompletionProvider = useCallback(() => {
      if (!monacoRef.current) return;
      if (providerRef.current) {
        providerRef.current.dispose();
      }

      const monaco = monacoRef.current;
      try {
        providerRef.current = monaco.languages.registerCompletionItemProvider(MONACO_LANGUAGE_ID, {
          triggerCharacters: ["=", "(", ",", "_", "[", "."],
          provideCompletionItems: (model, position) => {
            const cursorOffset = model.getOffsetAt(position);
            const token = getTokenBeforeCursor(model.getValue(), cursorOffset);
            const tokenUpper = token.toUpperCase();
            const word = model.getWordUntilPosition(position);
            const range = {
              startLineNumber: position.lineNumber,
              endLineNumber: position.lineNumber,
              startColumn: word.startColumn,
              endColumn: word.endColumn,
            };

            const suggestions = new Map<string, Monaco.languages.CompletionItem>();
            const matchesToken = (candidate: string): boolean =>
              !tokenUpper || candidate.toUpperCase().includes(tokenUpper);

            const addSuggestion = (
              label: string,
              kind: Monaco.languages.CompletionItemKind,
              insertText: string,
              detail: string,
              order: string,
              asSnippet = false
            ) => {
              if (!matchesToken(label)) {
                return;
              }
              const key = label.toUpperCase();
              if (suggestions.has(key)) {
                return;
              }
              suggestions.set(key, {
                label,
                kind,
                insertText,
                insertTextRules: asSnippet
                  ? monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet
                  : undefined,
                detail,
                range,
                sortText: `${order}_${label.toUpperCase()}`,
              });
            };

            collectLetLambdaLocals(model.getValue(), cursorOffset).forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Variable,
                name,
                "LET/LAMBDA variable",
                "0"
              );
            });

            namesRef.current.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Variable,
                name,
                "Named range/function",
                "1"
              );
            });

            tablesRef.current.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Class,
                `${name}[`,
                "Table",
                "2"
              );
            });

            FUNCTION_SUGGESTIONS.forEach((name) => {
              addSuggestion(
                name,
                monaco.languages.CompletionItemKind.Function,
                `${name}($1)`,
                "Function",
                "3",
                true
              );
            });

            return { suggestions: Array.from(suggestions.values()).slice(0, 200) };
          },
        });
      } catch (error) {
        const message = normalizeError(error);
        setEditorLoadError(message);
        setStatusType("error");
        setStatus(`Autocomplete registration failed. Fallback mode enabled: ${message}`);
      }
    }, []);

    const beforeMount = useCallback(
      (monaco: typeof Monaco) => {
        monacoRef.current = monaco;
        if (!monaco.languages.getLanguages().some((lang) => lang.id === MONACO_LANGUAGE_ID)) {
          monaco.languages.register({ id: MONACO_LANGUAGE_ID });
          monaco.languages.setLanguageConfiguration(MONACO_LANGUAGE_ID, {
            brackets: [
              ["(", ")"],
              ["[", "]"],
            ],
            autoClosingPairs: [
              { open: "(", close: ")" },
              { open: "[", close: "]" },
              { open: '"', close: '"' },
            ],
            surroundingPairs: [
              { open: "(", close: ")" },
              { open: "[", close: "]" },
              { open: '"', close: '"' },
            ],
            wordPattern: /[A-Za-z_\\][A-Za-z0-9_.\\]*/g,
          });
        }
        registerCompletionProvider();
      },
      [registerCompletionProvider]
    );

    const onMount: OnMount = useCallback((editor) => {
      editorRef.current = editor;
      editor.onDidChangeModelContent(() => {
        setFormulaText(editor.getValue());
      });
    }, []);

    useEffect(() => {
      registerCompletionProvider();
      return () => {
        providerRef.current?.dispose();
      };
    }, [registerCompletionProvider]);

    useEffect(() => {
      if (!editorRef.current) return;
      if (editorRef.current.getValue() !== formulaText) {
        editorRef.current.setValue(formulaText);
      }
    }, [formulaText]);

    const saveCurrentAsNamedFunction = async () => {
      await runAction("Save named function", async () => {
        const current = (editorRef.current?.getValue() ?? formulaText).trim();
        if (!namedFunctionName.trim()) {
          throw new Error("Function name is required.");
        }
        await saveNamedFunction(namedFunctionName, current);
        await loadSuggestions();
        setShowSaveFunctionModal(false);
      });
    };

    React.useImperativeHandle(ref, () => ({
      openOnly: async () => {
        await runAction("Load active cell state", async () => {
          const state = await getActiveCellFormulaState();
          setActiveCell(state);
        });
      },
      openAndPull: async () => {
        await runAction("Open active cell formula in editor", async () => {
          await pullFromActiveCell(true);
        });
      },
      pull: async () => {
        await runAction("Pull active cell formula", async () => {
          await pullFromActiveCell(true);
        });
      },
      apply: async () => {
        await runAction("Apply editor formula to active cell", applyToActiveCell);
      },
      beautify: async () => {
        await runAction("Beautify formula", beautify);
      },
      insertSelection: async () => {
        await runAction("Insert grid selection", insertSelection);
      },
    }));

    const compactEditorHeight = isPopout ? 340 : 280;
    const expandedEditorHeight = isPopout ? 560 : 440;
    const editorHeight = `${isEditorExpanded ? expandedEditorHeight : compactEditorHeight}px`;

    return (
      <div className={styles.root}>
        {!isPopout ? (
          <div>
            <Text className={shared.sectionTitle}>Formulas</Text>
            <Text className={shared.sectionSubtitle}>Manage and edit workbook formulas with live test output.</Text>
          </div>
        ) : null}

        <div className={styles.card}>
          <div className={styles.autoCaptureRow}>
            <Text className={shared.mutedText}>
              Active Cell: {activeCell ? `${activeCell.sheet}!${activeCell.address}` : "not captured"}
              {" • "}
              <span style={{ color: MODERN_TOKENS.colorBrandStrong, fontWeight: 700 }}>Connected</span>
              {loadingMetadata ? " • loading suggestions..." : ""}
            </Text>
            <label className={styles.checkboxLabel}>
              <input type="checkbox" checked={autoCapture} onChange={(event) => setAutoCapture(event.target.checked)} />{" "}
              Auto-capture active formula
            </label>
          </div>

          <div className={styles.editorShell}>
            <div className={styles.overlayTop}>
              <Button size="small" onClick={() => void runAction("Beautify formula", beautify)}>
                Beautify
              </Button>
              <Button size="small" onClick={() => setIsEditorExpanded((prev) => !prev)}>
                {isEditorExpanded ? "Use Smaller Editor" : "Expand Editor"}
              </Button>
              {!isPopout ? (
                <Button size="small" onClick={() => void runAction("Open formula editor popout", openFormulaEditorPopout)}>
                  Open Pop-out
                </Button>
              ) : null}
              {!isPopout ? (
                <Button size="small" onClick={onOpenLegacy}>
                  Legacy View
                </Button>
              ) : null}
            </div>

            <div className={styles.editorWrap}>
              {editorLoadError ? (
                <textarea
                  aria-label="Formula editor fallback"
                  className={styles.fallbackEditor}
                  spellCheck={false}
                  style={{ height: editorHeight }}
                  value={formulaText}
                  onChange={(event) => setFormulaText(event.target.value)}
                />
              ) : editorReady ? (
                <Editor
                  height={editorHeight}
                  language={MONACO_LANGUAGE_ID}
                  value={formulaText}
                  beforeMount={beforeMount}
                  onMount={onMount}
                  theme="vs"
                  options={{
                    minimap: { enabled: false },
                    scrollBeyondLastLine: false,
                    fontSize: 13,
                    lineNumbers: "on",
                    wordWrap: "off",
                    automaticLayout: true,
                    suggestOnTriggerCharacters: true,
                    quickSuggestions: { other: true, comments: false, strings: false },
                    wordBasedSuggestions: "off",
                    tabCompletion: "on",
                    padding: {
                      top: 38,
                      bottom: 128,
                    },
                    suggest: {
                      showWords: false,
                      showSnippets: true,
                      preview: true,
                    },
                  }}
                />
              ) : (
                <div className={styles.editorLoading} style={{ height: editorHeight }}>
                  <Text className={shared.mutedText}>Loading formula editor...</Text>
                </div>
              )}
            </div>

            <div className={styles.overlayBottom}>
              <div className={styles.overlayCluster}>
                <Text className={styles.overlayTitle}>Pull Into Editor</Text>
                <div className={styles.actionRow}>
                  <Button
                    appearance="primary"
                    size="small"
                    onClick={() =>
                      void runAction("Pull formula", async () => {
                        await pullFromActiveCell(true);
                      })
                    }
                  >
                    Pull In
                  </Button>
                  <Button size="small" onClick={() => void runAction("Insert selection", insertSelection)}>
                    Insert Selection
                  </Button>
                </div>
              </div>

              <div className={styles.overlayCluster}>
                <Text className={styles.overlayTitle}>Apply / Save</Text>
                <div className={styles.actionRow}>
                  <Button appearance="primary" size="small" onClick={() => void runAction("Apply formula", applyToActiveCell)}>
                    Apply Formula
                  </Button>
                  <Button size="small" onClick={() => setShowSaveFunctionModal(true)}>
                    Save Named Function
                  </Button>
                </div>
              </div>
            </div>
          </div>

          <div className={styles.liveTestHeader}>
            <Text className={shared.cardTitle}>Live Test Output</Text>
            <Text className={shared.mutedText}>
              {liveTestBusy
                ? "Testing..."
                : liveTestError
                  ? `Last test failed${lastTestedAt ? ` at ${lastTestedAt}` : ""}`
                  : lastTestedAt
                    ? `Auto-tested at ${lastTestedAt}`
                    : "Waiting for formula input"}
            </Text>
          </div>

          {liveTestError ? <Text className={shared.errorText}>{liveTestError}</Text> : null}

          <div className={styles.outputWrap}>
            {formulaOutput ? (
              <table className={styles.outputTable}>
                <tbody>
                  {formulaOutput.values.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, colIndex) => (
                        <td key={`${rowIndex}-${colIndex}`} className={styles.outputCell}>
                          {cell === null ? "" : String(cell)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className={styles.outputCell}>
                <Text className={shared.mutedText}>Live test will render results here as you edit.</Text>
              </div>
            )}
          </div>

          {status ? <Text className={statusClass}>{status}</Text> : null}
        </div>

        {showSaveFunctionModal ? (
          <div className={styles.modalBackdrop}>
            <div className={styles.modal}>
              <Text className={shared.cardTitle}>Save As Named Function</Text>
              <Text className={shared.cardSubtitle}>
                Enter a function name. The current formula must be a valid LAMBDA expression.
              </Text>
              <Input
                placeholder="Function name (example: CalcMargin)"
                value={namedFunctionName}
                onChange={(_, data) => setNamedFunctionName(data.value)}
              />
              <div className={styles.modalActions}>
                <Button onClick={() => setShowSaveFunctionModal(false)}>Cancel</Button>
                <Button appearance="primary" onClick={() => void saveCurrentAsNamedFunction()}>
                  Save Function
                </Button>
              </div>
            </div>
          </div>
        ) : null}
      </div>
    );
  }
);

FormulaMonacoView.displayName = "FormulaMonacoView";

export default FormulaMonacoView;
