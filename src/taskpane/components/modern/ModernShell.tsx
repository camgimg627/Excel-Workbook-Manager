import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Button, Input, Text, makeStyles } from "@fluentui/react-components";
import {
  BookNumber20Regular,
  Box20Regular,
  ChevronDown20Regular,
  ChevronRight20Regular,
  DataArea20Regular,
  DataPie20Regular,
  Grid20Regular,
  QuestionCircle20Regular,
  Settings20Regular,
  Table20Regular,
  TextAlignJustify20Regular,
} from "@fluentui/react-icons";
import { NavigationTarget } from "../../navigation";
import FormulaMonacoView, { FormulaViewHandle } from "./FormulaMonacoView";
import NamesView from "./NamesView";
import TablesView from "./TablesView";
import SandboxDebugView from "./SandboxDebugView";
import FormatView from "./FormatView";
import PivotsView from "./PivotsView";
import QueriesView from "./QueriesView";
import ModelBuilderView from "./ModelBuilderView";
import SettingsView from "./SettingsView";
import HelpView from "./HelpView";
import { MODERN_TOKENS } from "./designTokens";
import { FormulaEvaluationResult, evaluateFormula } from "../../taskpane";

interface ModernShellProps {
  activeTarget: NavigationTarget;
  createRequestId: number;
  onTargetChange: (target: NavigationTarget) => void;
  onSwitchToLegacy: () => void;
  onResetUiPreference: () => void;
  formulaViewRef: React.RefObject<FormulaViewHandle>;
}

type WorkspaceId = "structure" | "formulas" | "data" | "model" | "layout" | "tools";

interface WorkspaceRoute {
  workspace: WorkspaceId;
  sectionId: string;
}

interface WorkspaceNavItem {
  id: WorkspaceId;
  label: string;
  icon: React.ReactNode;
}

const TARGET_ROUTE_MAP: Record<NavigationTarget, WorkspaceRoute> = {
  names: { workspace: "structure", sectionId: "namedRanges" },
  "names-create": { workspace: "structure", sectionId: "namedRanges" },
  tables: { workspace: "structure", sectionId: "tables" },
  formulas: { workspace: "formulas", sectionId: "formulaEditor" },
  queries: { workspace: "data", sectionId: "queryStatus" },
  pivots: { workspace: "data", sectionId: "pivotRefresh" },
  "model-builder": { workspace: "model", sectionId: "parameterBuilder" },
  format: { workspace: "layout", sectionId: "gridFreeze" },
  "sandbox-debug": { workspace: "tools", sectionId: "debugUtilities" },
  settings: { workspace: "tools", sectionId: "settings" },
  help: { workspace: "tools", sectionId: "help" },
};

const DEFAULT_TARGET_BY_WORKSPACE: Record<WorkspaceId, NavigationTarget> = {
  structure: "names",
  formulas: "formulas",
  data: "queries",
  model: "model-builder",
  layout: "format",
  tools: "settings",
};

const SECTION_TARGET_MAP: Record<WorkspaceId, Record<string, NavigationTarget>> = {
  structure: {
    namedRanges: "names",
    tables: "tables",
    bulkRename: "names",
    cleanupGovernance: "names",
  },
  formulas: {
    formulaEditor: "formulas",
    evaluateTest: "formulas",
    insertHelpers: "formulas",
    namedFunctions: "formulas",
  },
  data: {
    queryStatus: "queries",
    refreshQueries: "queries",
    pivotRefresh: "pivots",
    errorsWarnings: "queries",
  },
  model: {
    parameterBuilder: "model-builder",
    validationLists: "model-builder",
    defaultFormulaValues: "model-builder",
    applySummary: "model-builder",
  },
  layout: {
    gridFreeze: "format",
  },
  tools: {
    debugUtilities: "sandbox-debug",
    settings: "settings",
    help: "help",
    legacyCompatibility: "settings",
  },
};

const DEFAULT_EXPANDED_BY_WORKSPACE: Record<WorkspaceId, string[]> = {
  structure: ["namedRanges"],
  formulas: ["formulaEditor", "evaluateTest"],
  data: ["queryStatus"],
  model: ["parameterBuilder"],
  layout: ["gridFreeze"],
  tools: ["settings"],
};

const WORKSPACE_META: Record<WorkspaceId, { title: string; description: string }> = {
  structure: {
    title: "Structure",
    description: "Manage workbook names, tables, and structural assets.",
  },
  formulas: {
    title: "Formulas",
    description: "Write, test, and apply formulas with a focused editing workflow.",
  },
  data: {
    title: "Data",
    description: "Refresh, verify, and monitor workbook data operations.",
  },
  model: {
    title: "Model",
    description: "Build structured workbook inputs and parameter controls.",
  },
  layout: {
    title: "Layout",
    description: "Apply workbook presentation and usability improvements.",
  },
  tools: {
    title: "Tools",
    description: "Access support, diagnostics, compatibility, and advanced utilities.",
  },
};

const useStyles = makeStyles({
  root: {
    height: "100vh",
    backgroundColor: MODERN_TOKENS.colorBg,
    color: MODERN_TOKENS.colorText,
    display: "grid",
    gridTemplateColumns: "132px 1fr",
    overflow: "hidden",
  },
  navRail: {
    borderRight: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "10px 8px 8px",
    display: "grid",
    gridTemplateRows: "auto 1fr auto",
    gap: "10px",
    minHeight: 0,
  },
  identity: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "8px",
    display: "grid",
    gap: "2px",
  },
  identityTitle: {
    fontSize: "12px",
    lineHeight: "16px",
    fontWeight: 700,
  },
  identitySub: {
    fontSize: "10px",
    lineHeight: "14px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  navList: {
    display: "grid",
    alignContent: "start",
    gap: "4px",
    minHeight: 0,
    overflowY: "auto",
  },
  navButton: {
    border: "1px solid transparent",
    backgroundColor: "transparent",
    color: MODERN_TOKENS.colorText,
    borderRadius: "9px",
    padding: "8px 7px",
    display: "grid",
    gap: "3px",
    justifyItems: "center",
    alignItems: "center",
    cursor: "pointer",
    fontSize: "10px",
    fontWeight: 600,
    lineHeight: "12px",
    textAlign: "center",
    minHeight: "54px",
    selectors: {
      "&:hover": {
        backgroundColor: "#F7FAFF",
      },
      "&:focus-visible": {
        outline: `2px solid ${MODERN_TOKENS.colorBrand}`,
        outlineOffset: "1px",
      },
    },
  },
  navButtonActive: {
    backgroundColor: "#E9F1FD",
    border: "1px solid #B8CDED",
    color: "#194279",
    fontWeight: 700,
  },
  navIcon: {
    width: "20px",
    height: "20px",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
  },
  railFooter: {
    borderRadius: "9px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "7px",
    display: "grid",
    gap: "5px",
  },
  railFooterText: {
    fontSize: "10px",
    lineHeight: "14px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  contentViewport: {
    overflowY: "auto",
    minHeight: 0,
    padding: "12px",
  },
  workspacePage: {
    display: "grid",
    gap: "10px",
    minHeight: "100%",
  },
  workspaceHeader: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    boxShadow: "0 1px 2px rgba(17,24,39,0.05)",
    padding: "10px 12px",
    display: "grid",
    gap: "4px",
  },
  workspaceTitleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    flexWrap: "wrap",
  },
  workspaceTitle: {
    fontSize: "20px",
    lineHeight: "22px",
    fontWeight: 700,
  },
  workspaceDescription: {
    fontSize: "12px",
    lineHeight: "16px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  actionBar: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "8px",
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  actionHint: {
    marginLeft: "auto",
    fontSize: "11px",
    lineHeight: "14px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  sectionStack: {
    display: "grid",
    gap: "8px",
    alignContent: "start",
  },
  sectionCard: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    overflow: "hidden",
    boxShadow: "0 1px 2px rgba(17,24,39,0.05)",
  },
  sectionHeaderButton: {
    width: "100%",
    border: "none",
    borderRadius: 0,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    padding: "10px 12px",
    display: "grid",
    gridTemplateColumns: "1fr auto",
    alignItems: "center",
    gap: "8px",
    textAlign: "left",
    cursor: "pointer",
    selectors: {
      "&:hover": {
        backgroundColor: "#F8FAFC",
      },
      "&:focus-visible": {
        outline: `2px solid ${MODERN_TOKENS.colorBrand}`,
        outlineOffset: "-2px",
      },
    },
  },
  sectionHeaderMain: {
    display: "inline-flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
  },
  sectionHeaderTitle: {
    fontSize: "12px",
    lineHeight: "16px",
    fontWeight: 700,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  sectionHeaderMeta: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  sectionCount: {
    borderRadius: "999px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "2px 7px",
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
  },
  sectionChip: {
    borderRadius: "999px",
    border: "1px solid #B8CDED",
    backgroundColor: "#E9F1FD",
    color: "#194279",
    padding: "2px 7px",
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
  },
  sectionBody: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px",
    backgroundColor: "#FFFFFF",
  },
  evaluatePanel: {
    display: "grid",
    gap: "8px",
  },
  evaluateRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  evaluateInput: {
    flex: 1,
    minWidth: "220px",
  },
  evaluateOutputWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    overflow: "auto",
    backgroundColor: "#FFFFFF",
    minHeight: "56px",
  },
  evaluateOutputTable: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: "12px",
  },
  evaluateOutputCell: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px 10px",
  },
  evaluateStatus: {
    fontSize: "12px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  evaluateStatusError: {
    color: MODERN_TOKENS.colorDanger,
  },
  placeholder: {
    borderRadius: "8px",
    border: `1px dashed ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FAFBFC",
    padding: "12px",
    display: "grid",
    gap: "8px",
  },
  placeholderText: {
    fontSize: "12px",
    lineHeight: "16px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  placeholderActions: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
});

const navItems: WorkspaceNavItem[] = [
  { id: "structure", label: "Structure", icon: <BookNumber20Regular /> },
  { id: "formulas", label: "Formulas", icon: <DataArea20Regular /> },
  { id: "data", label: "Data", icon: <DataPie20Regular /> },
  { id: "model", label: "Model", icon: <TextAlignJustify20Regular /> },
  { id: "layout", label: "Layout", icon: <Grid20Regular /> },
  { id: "tools", label: "Tools", icon: <Settings20Regular /> },
];

const ModernShell: React.FC<ModernShellProps> = ({
  activeTarget,
  createRequestId,
  onTargetChange,
  onSwitchToLegacy,
  onResetUiPreference,
  formulaViewRef,
}) => {
  const styles = useStyles();
  const activeRoute = TARGET_ROUTE_MAP[activeTarget];
  const activeWorkspace = activeRoute.workspace;
  const [expandedSections, setExpandedSections] = useState<Record<WorkspaceId, string[]>>(DEFAULT_EXPANDED_BY_WORKSPACE);
  const [testFormulaCall, setTestFormulaCall] = useState<string>("");
  const [testFormulaOutput, setTestFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
  const [testBusy, setTestBusy] = useState<boolean>(false);
  const [testStatus, setTestStatus] = useState<string>("");
  const [testStatusType, setTestStatusType] = useState<"success" | "error">("success");

  const normalizeError = (error: unknown): string => (error instanceof Error ? error.message : String(error));

  const openLegacy = () => {
    onSwitchToLegacy();
  };

  useEffect(() => {
    const route = TARGET_ROUTE_MAP[activeTarget];
    if (!route) {
      return;
    }
    setExpandedSections((prev) => {
      const current = prev[route.workspace];
      if (current.includes(route.sectionId)) {
        return prev;
      }
      const nextSections =
        route.workspace === "formulas" ? [...current, route.sectionId].slice(-2) : [route.sectionId];
      return {
        ...prev,
        [route.workspace]: nextSections,
      };
    });
  }, [activeTarget]);

  const openWorkspace = (workspace: WorkspaceId) => {
    onTargetChange(DEFAULT_TARGET_BY_WORKSPACE[workspace]);
    setExpandedSections((prev) => {
      if (prev[workspace].length > 0) {
        return prev;
      }
      return {
        ...prev,
        [workspace]: [...DEFAULT_EXPANDED_BY_WORKSPACE[workspace]],
      };
    });
  };

  const openSection = (workspace: WorkspaceId, sectionId: string) => {
    const nextTarget = SECTION_TARGET_MAP[workspace][sectionId] ?? DEFAULT_TARGET_BY_WORKSPACE[workspace];
    onTargetChange(nextTarget);
    setExpandedSections((prev) => {
      const current = prev[workspace];
      if (current.includes(sectionId)) {
        return prev;
      }
      const nextSections =
        workspace === "formulas" ? [...current, sectionId].slice(-2) : [sectionId];
      return {
        ...prev,
        [workspace]: nextSections,
      };
    });
  };

  const toggleSection = (workspace: WorkspaceId, sectionId: string) => {
    setExpandedSections((prev) => {
      const current = prev[workspace];
      const isOpen = current.includes(sectionId);
      let nextSections: string[];
      if (workspace === "formulas") {
        nextSections = isOpen ? current.filter((id) => id !== sectionId) : [...current, sectionId].slice(-2);
      } else {
        nextSections = isOpen ? [] : [sectionId];
      }
      return {
        ...prev,
        [workspace]: nextSections,
      };
    });
  };

  const runFormulaAction = (action: keyof FormulaViewHandle) => {
    const view = formulaViewRef.current;
    if (!view) {
      return;
    }
    void view[action]();
  };

  const isSectionExpanded = (workspace: WorkspaceId, sectionId: string): boolean =>
    expandedSections[workspace].includes(sectionId);

  const renderSection = (
    workspace: WorkspaceId,
    sectionId: string,
    title: string,
    icon: React.ReactNode,
    content: React.ReactNode,
    options?: { count?: string; statusChip?: string }
  ) => {
    const expanded = isSectionExpanded(workspace, sectionId);
    return (
      <div className={styles.sectionCard} key={`${workspace}:${sectionId}`}>
        <button
          type="button"
          className={styles.sectionHeaderButton}
          onClick={() => {
            if (!expanded) {
              openSection(workspace, sectionId);
            } else {
              toggleSection(workspace, sectionId);
            }
          }}
          aria-expanded={expanded}
        >
          <span className={styles.sectionHeaderMain}>
            <span className={styles.navIcon}>{icon}</span>
            <span className={styles.sectionHeaderTitle}>{title}</span>
          </span>
          <span className={styles.sectionHeaderMeta}>
            {options?.count ? <span className={styles.sectionCount}>{options.count}</span> : null}
            {options?.statusChip ? <span className={styles.sectionChip}>{options.statusChip}</span> : null}
            {expanded ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
          </span>
        </button>
        {expanded ? <div className={styles.sectionBody}>{content}</div> : null}
      </div>
    );
  };

  const workspaceMeta = WORKSPACE_META[activeWorkspace];

  const syncTestFormulaFromEditor = () => {
    const fromEditor = formulaViewRef.current?.getCurrentFormula?.().trim() ?? "";
    if (!fromEditor) {
      setTestStatusType("error");
      setTestStatus("Formula editor is empty.");
      return "";
    }
    setTestFormulaCall(fromEditor);
    return fromEditor;
  };

  const executeFormulaTest = async (formulaToRun: string, silent: boolean): Promise<void> => {
    const normalized = formulaToRun.trim();
    if (!normalized) {
      setTestFormulaOutput(null);
      if (!silent) {
        setTestStatusType("error");
        setTestStatus("Enter a formula to test.");
      }
      return;
    }
    setTestBusy(true);
    try {
      const result = await evaluateFormula(normalized);
      setTestFormulaOutput(result);
      if (!silent) {
        setTestStatusType("success");
        setTestStatus(`Formula test completed (${result.address}).`);
      }
    } catch (error) {
      setTestFormulaOutput(null);
      setTestStatusType("error");
      setTestStatus(`Formula test failed: ${normalizeError(error)}`);
    } finally {
      setTestBusy(false);
    }
  };

  const runManualFormulaTest = async () => {
    const editorFormula = formulaViewRef.current?.getCurrentFormula?.().trim() ?? "";
    const candidate = testFormulaCall.trim() || editorFormula;
    if (!candidate) {
      setTestStatusType("error");
      setTestStatus("Formula editor is empty.");
      return;
    }
    if (!testFormulaCall.trim()) {
      setTestFormulaCall(candidate);
    }
    await executeFormulaTest(candidate, false);
  };

  useEffect(() => {
    if (activeWorkspace !== "formulas" || !isSectionExpanded("formulas", "evaluateTest")) {
      return;
    }
    if (testFormulaCall.trim()) {
      return;
    }
    const fromEditor = formulaViewRef.current?.getCurrentFormula?.().trim() ?? "";
    if (fromEditor) {
      setTestFormulaCall(fromEditor);
    }
  }, [activeWorkspace, expandedSections, testFormulaCall, formulaViewRef]);

  useEffect(() => {
    if (activeWorkspace !== "formulas" || !isSectionExpanded("formulas", "evaluateTest")) {
      return undefined;
    }
    const formulaToRun = testFormulaCall.trim();
    if (!formulaToRun) {
      setTestFormulaOutput(null);
      return undefined;
    }
    const timerId = window.setTimeout(() => {
      void executeFormulaTest(formulaToRun, true);
    }, 450);
    return () => {
      window.clearTimeout(timerId);
    };
  }, [activeWorkspace, expandedSections, testFormulaCall]);

  const workspacePrimaryAction = useMemo(() => {
    if (activeWorkspace === "structure") {
      return (
        <Button size="small" appearance="primary" onClick={() => openSection("structure", "namedRanges")}>
          Create
        </Button>
      );
    }
    if (activeWorkspace === "formulas") {
      return (
        <Button size="small" appearance="primary" onClick={() => runFormulaAction("pull")}>
          Pull
        </Button>
      );
    }
    if (activeWorkspace === "data") {
      return (
        <Button size="small" appearance="primary" onClick={() => openSection("data", "queryStatus")}>
          Refresh All
        </Button>
      );
    }
    if (activeWorkspace === "model") {
      return (
        <Button size="small" appearance="primary" onClick={() => openSection("model", "parameterBuilder")}>
          Build
        </Button>
      );
    }
    if (activeWorkspace === "layout") {
      return (
        <Button size="small" appearance="primary" onClick={() => openSection("layout", "gridFreeze")}>
          Apply
        </Button>
      );
    }
    return null;
  }, [activeWorkspace]);

  const renderWorkspaceActionBar = () => {
    if (activeWorkspace === "structure") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" appearance={activeRoute.sectionId === "namedRanges" ? "primary" : "secondary"} onClick={() => openSection("structure", "namedRanges")}>
            Named Ranges
          </Button>
          <Button size="small" appearance={activeRoute.sectionId === "tables" ? "primary" : "secondary"} onClick={() => openSection("structure", "tables")}>
            Tables
          </Button>
          <Text className={styles.actionHint}>Use accordion sections to keep structure tasks compact.</Text>
        </div>
      );
    }
    if (activeWorkspace === "formulas") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" onClick={() => runFormulaAction("pull")}>
            Pull
          </Button>
          <Button size="small" onClick={() => runFormulaAction("apply")}>
            Apply
          </Button>
          <Button size="small" onClick={() => runFormulaAction("beautify")}>
            Beautify
          </Button>
          <Button size="small" onClick={() => runFormulaAction("insertSelection")}>
            Insert Selection
          </Button>
          <Text className={styles.actionHint}>Cell context and diagnostics are inside Formula Editor.</Text>
        </div>
      );
    }
    if (activeWorkspace === "data") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" appearance={activeRoute.sectionId === "queryStatus" ? "primary" : "secondary"} onClick={() => openSection("data", "queryStatus")}>
            Query Status
          </Button>
          <Button size="small" appearance={activeRoute.sectionId === "pivotRefresh" ? "primary" : "secondary"} onClick={() => openSection("data", "pivotRefresh")}>
            Pivot Refresh
          </Button>
          <Text className={styles.actionHint}>Refresh controls remain inside each data section.</Text>
        </div>
      );
    }
    if (activeWorkspace === "model") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" appearance={activeRoute.sectionId === "parameterBuilder" ? "primary" : "secondary"} onClick={() => openSection("model", "parameterBuilder")}>
            Parameter Builder
          </Button>
          <Text className={styles.actionHint}>Use one anchor and apply structured parameter rows.</Text>
        </div>
      );
    }
    if (activeWorkspace === "layout") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" appearance="secondary" onClick={() => openSection("layout", "gridFreeze")}>
            Gridlines and Freeze
          </Button>
          <Text className={styles.actionHint}>Presentation helpers and templates are grouped below.</Text>
        </div>
      );
    }
    return (
      <div className={styles.actionBar}>
        <Button size="small" appearance={activeRoute.sectionId === "settings" ? "primary" : "secondary"} onClick={() => openSection("tools", "settings")}>
          Settings
        </Button>
        <Button size="small" appearance={activeRoute.sectionId === "help" ? "primary" : "secondary"} onClick={() => openSection("tools", "help")}>
          Help
        </Button>
        <Text className={styles.actionHint}>Diagnostics and compatibility tools stay in this workspace.</Text>
      </div>
    );
  };

  const renderWorkspaceSections = () => {
    if (activeWorkspace === "structure") {
      return (
        <div className={styles.sectionStack}>
          {renderSection("structure", "namedRanges", "Named Ranges", <BookNumber20Regular />, (
            <NamesView createRequestId={createRequestId} onOpenLegacy={openLegacy} embedded />
          ), { statusChip: "Active" })}
          {renderSection("structure", "tables", "Tables", <Table20Regular />, (
            <TablesView onOpenLegacy={openLegacy} embedded />
          ))}
          {renderSection("structure", "bulkRename", "Bulk Rename", <TextAlignJustify20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Bulk rename is available in the Named Ranges and Tables sections. Open either section to run scoped batch changes.
              </Text>
              <div className={styles.placeholderActions}>
                <Button size="small" onClick={() => openSection("structure", "namedRanges")}>
                  Open Named Ranges
                </Button>
                <Button size="small" onClick={() => openSection("structure", "tables")}>
                  Open Tables
                </Button>
              </div>
            </div>
          ))}
          {renderSection("structure", "cleanupGovernance", "Cleanup and Governance", <Settings20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Use search, sort, type filtering, and bulk tools to maintain naming consistency across workbook assets.
              </Text>
            </div>
          ))}
        </div>
      );
    }

    if (activeWorkspace === "formulas") {
      return (
        <div className={styles.sectionStack}>
          {renderSection("formulas", "formulaEditor", "Formula Editor", <DataArea20Regular />, (
            <FormulaMonacoView ref={formulaViewRef} isPopout={false} onOpenLegacy={openLegacy} embedded />
          ), { statusChip: "Live" })}
          {renderSection("formulas", "evaluateTest", "Evaluate and Test", <DataPie20Regular />, (
            <div className={styles.evaluatePanel}>
              <Text className={styles.placeholderText}>Current Formula / Function Call</Text>
              <div className={styles.evaluateRow}>
                <Input
                  className={styles.evaluateInput}
                  value={testFormulaCall}
                  placeholder="=SUM(A1:A10)"
                  onChange={(_, data) => setTestFormulaCall(data.value)}
                />
                <Button size="small" onClick={syncTestFormulaFromEditor}>
                  Use Editor Formula
                </Button>
                <Button size="small" appearance="primary" onClick={() => void runManualFormulaTest()} disabled={testBusy}>
                  Run Test
                </Button>
              </div>
              <div className={styles.evaluateOutputWrap}>
                {testFormulaOutput ? (
                  <table className={styles.evaluateOutputTable}>
                    <tbody>
                      {testFormulaOutput.values.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {row.map((cell, colIndex) => (
                            <td key={`${rowIndex}-${colIndex}`} className={styles.evaluateOutputCell}>
                              {cell === null ? "" : String(cell)}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className={styles.evaluateOutputCell}>
                    <Text className={styles.placeholderText}>Results will render here after running the test.</Text>
                  </div>
                )}
              </div>
              {testStatus ? (
                <Text className={`${styles.evaluateStatus} ${testStatusType === "error" ? styles.evaluateStatusError : ""}`}>
                  {testStatus}
                </Text>
              ) : null}
            </div>
          ))}
          {renderSection("formulas", "insertHelpers", "Insert Helpers", <Table20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Use Insert Selection and named/table suggestions from the Formula Editor overlay for helper insertion.
              </Text>
            </div>
          ))}
          {renderSection("formulas", "namedFunctions", "Named Functions", <BookNumber20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Named function authoring is included in Function mode within the Formula Editor workflow.
              </Text>
            </div>
          ))}
        </div>
      );
    }

    if (activeWorkspace === "data") {
      return (
        <div className={styles.sectionStack}>
          {renderSection("data", "queryStatus", "Query Status", <DataPie20Regular />, (
            <QueriesView onOpenLegacy={openLegacy} embedded />
          ), { statusChip: "Live" })}
          {renderSection("data", "refreshQueries", "Refresh Queries", <DataArea20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Query refresh actions and load diagnostics are available in the Query Status section controls.
              </Text>
              <div className={styles.placeholderActions}>
                <Button size="small" onClick={() => openSection("data", "queryStatus")}>
                  Open Query Status
                </Button>
              </div>
            </div>
          ))}
          {renderSection("data", "pivotRefresh", "Pivot Refresh", <Table20Regular />, (
            <PivotsView onOpenLegacy={openLegacy} embedded />
          ))}
          {renderSection("data", "errorsWarnings", "Errors and Warnings", <QuestionCircle20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Refresh warnings are surfaced inline in query and pivot operations to avoid full-page alert banners.
              </Text>
            </div>
          ))}
        </div>
      );
    }

    if (activeWorkspace === "model") {
      return (
        <div className={styles.sectionStack}>
          {renderSection("model", "parameterBuilder", "Parameter Builder", <TextAlignJustify20Regular />, (
            <ModelBuilderView onOpenLegacy={openLegacy} embedded />
          ), { statusChip: "Ready" })}
          {renderSection("model", "validationLists", "Validation Lists", <Table20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Validation list configuration is part of each parameter row and appears automatically during apply.
              </Text>
            </div>
          ))}
          {renderSection("model", "defaultFormulaValues", "Default and Formula Values", <DataArea20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Manual and formula-driven defaults are configured per row in the builder grid.
              </Text>
            </div>
          ))}
          {renderSection("model", "applySummary", "Apply Summary", <DataPie20Regular />, (
            <div className={styles.placeholder}>
              <Text className={styles.placeholderText}>
                Apply output includes created names, validations, and warnings for workbook setup verification.
              </Text>
            </div>
          ))}
        </div>
      );
    }

    if (activeWorkspace === "layout") {
      return (
        <div className={styles.sectionStack}>
          {renderSection("layout", "gridFreeze", "Gridlines and Freeze Panes", <Grid20Regular />, (
            <FormatView onOpenLegacy={openLegacy} embedded />
          ), { statusChip: "Active" })}
        </div>
      );
    }

    return (
      <div className={styles.sectionStack}>
        {renderSection("tools", "debugUtilities", "Debug Utilities", <Box20Regular />, (
          <SandboxDebugView onOpenLegacy={openLegacy} embedded />
        ))}
        {renderSection("tools", "settings", "Settings", <Settings20Regular />, (
          <SettingsView onOpenLegacy={openLegacy} onResetUiPreference={onResetUiPreference} embedded />
        ), { statusChip: "Active" })}
        {renderSection("tools", "help", "Help", <QuestionCircle20Regular />, (
          <HelpView onNavigate={onTargetChange} embedded />
        ))}
        {renderSection("tools", "legacyCompatibility", "Legacy and Compatibility", <BookNumber20Regular />, (
          <div className={styles.placeholder}>
            <Text className={styles.placeholderText}>
              Keep modern mode as default. Legacy mode remains available for temporary compatibility fallback.
            </Text>
            <div className={styles.placeholderActions}>
              <Button size="small" onClick={openLegacy}>
                Open Legacy UI
              </Button>
            </div>
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className={styles.root}>
      <aside className={styles.navRail}>
        <div className={styles.identity}>
          <Text className={styles.identityTitle}>Workbook Manager</Text>
          <Text className={styles.identitySub}>Workbook operations</Text>
        </div>

        <nav className={styles.navList} aria-label="Workspace navigation">
          {navItems.map((item) => {
            const selected = activeWorkspace === item.id;
            return (
              <button
                key={item.id}
                type="button"
                className={`${styles.navButton} ${selected ? styles.navButtonActive : ""}`}
                onClick={() => openWorkspace(item.id)}
                aria-current={selected ? "page" : undefined}
              >
                <span className={styles.navIcon}>{item.icon}</span>
                <span>{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className={styles.railFooter}>
          <Text className={styles.railFooterText}>Mode: Modern</Text>
          <Button size="small" onClick={openLegacy}>
            Legacy
          </Button>
        </div>
      </aside>

      <main className={styles.contentViewport}>
        <div className={styles.workspacePage}>
          <div className={styles.workspaceHeader}>
            <div className={styles.workspaceTitleRow}>
              <Text className={styles.workspaceTitle}>{workspaceMeta.title}</Text>
              {workspacePrimaryAction}
            </div>
            <Text className={styles.workspaceDescription}>{workspaceMeta.description}</Text>
          </div>
          {renderWorkspaceActionBar()}
          {renderWorkspaceSections()}
        </div>
      </main>
    </div>
  );
};

export default ModernShell;
