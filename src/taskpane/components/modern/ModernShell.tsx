import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
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
  },
  formulas: {
    formulaEditor: "formulas",
    evaluateTest: "formulas",
  },
  data: {
    queryStatus: "queries",
    pivotRefresh: "pivots",
  },
  model: {
    parameterBuilder: "model-builder",
  },
  layout: {
    gridFreeze: "format",
  },
  tools: {
    debugUtilities: "sandbox-debug",
    settings: "settings",
    help: "help",
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
    "@media (max-width: 720px)": {
      marginLeft: 0,
      flexBasis: "100%",
    },
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
  evaluateFormulaPreview: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    backgroundColor: "#F8FBFF",
    padding: "10px 12px",
    maxHeight: "180px",
    overflow: "auto",
  },
  evaluateFormulaCode: {
    fontFamily: "Consolas, 'Cascadia Mono', 'Courier New', monospace",
    fontSize: "12px",
    lineHeight: "1.65",
    color: MODERN_TOKENS.colorText,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    margin: 0,
  },
  evaluateRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    flexWrap: "wrap",
  },
  evaluateHint: {
    flex: 1,
    minWidth: "180px",
    fontSize: "11px",
    lineHeight: "15px",
    color: MODERN_TOKENS.colorTextMuted,
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
  const [expandedSections, setExpandedSections] = useState<Record<WorkspaceId, string[]>>(
    DEFAULT_EXPANDED_BY_WORKSPACE
  );
  const [editorFormulaText, setEditorFormulaText] = useState<string>("");
  const [testFormulaOutput, setTestFormulaOutput] = useState<FormulaEvaluationResult | null>(null);
  const [testBusy, setTestBusy] = useState<boolean>(false);
  const [testStatus, setTestStatus] = useState<string>("");
  const [testStatusType, setTestStatusType] = useState<"success" | "error">("success");
  const [lastEvaluatedFormula, setLastEvaluatedFormula] = useState<string>("");
  const [formatSelectionRequest, setFormatSelectionRequest] = useState<{
    requestId: number;
    sheetName: string;
    shapeName: string;
  } | null>(null);

  const normalizeError = (error: unknown): string =>
    error instanceof Error ? error.message : String(error);

  const openLegacy = () => {
    onSwitchToLegacy();
  };

  const openShapeEditor = (request: { sheetName: string; shapeName: string }) => {
    setFormatSelectionRequest((prev) => ({
      requestId: (prev?.requestId ?? 0) + 1,
      ...request,
    }));
    onTargetChange("format");
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
        route.workspace === "formulas"
          ? [...current, route.sectionId].slice(-2)
          : [route.sectionId];
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
    const nextTarget =
      SECTION_TARGET_MAP[workspace][sectionId] ?? DEFAULT_TARGET_BY_WORKSPACE[workspace];
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
        nextSections = isOpen
          ? current.filter((id) => id !== sectionId)
          : [...current, sectionId].slice(-2);
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
    const triggerId = `${workspace}-${sectionId}-trigger`;
    const panelId = `${workspace}-${sectionId}-panel`;
    return (
      <div className={styles.sectionCard} key={`${workspace}:${sectionId}`}>
        <button
          id={triggerId}
          type="button"
          className={styles.sectionHeaderButton}
          onClick={() => {
            if (!expanded) {
              openSection(workspace, sectionId);
            } else {
              toggleSection(workspace, sectionId);
            }
          }}
          aria-expanded={expanded ? "true" : "false"}
          aria-controls={panelId}
        >
          <span className={styles.sectionHeaderMain}>
            <span className={styles.navIcon}>{icon}</span>
            <span className={styles.sectionHeaderTitle}>{title}</span>
          </span>
          <span className={styles.sectionHeaderMeta}>
            {options?.count ? <span className={styles.sectionCount}>{options.count}</span> : null}
            {options?.statusChip ? (
              <span className={styles.sectionChip}>{options.statusChip}</span>
            ) : null}
            {expanded ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
          </span>
        </button>
        {expanded ? (
          <div
            id={panelId}
            className={styles.sectionBody}
            role="region"
            aria-labelledby={triggerId}
          >
            {content}
          </div>
        ) : null}
      </div>
    );
  };

  const workspaceMeta = WORKSPACE_META[activeWorkspace];

  const executeFormulaTest = async (formulaToRun: string, silent: boolean): Promise<void> => {
    const normalized = formulaToRun.trim();
    if (!normalized) {
      setTestFormulaOutput(null);
      setLastEvaluatedFormula("");
      if (!silent) {
        setTestStatusType("error");
        setTestStatus("Enter a formula to test.");
      }
      return;
    }
    setLastEvaluatedFormula(normalized);
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
    const candidate =
      editorFormulaText.trim() || formulaViewRef.current?.getCurrentFormula?.().trim() || "";
    if (!candidate) {
      setTestStatusType("error");
      setTestStatus("Formula editor is empty.");
      return;
    }
    await executeFormulaTest(candidate, false);
  };

  useEffect(() => {
    const normalized = editorFormulaText.trim();
    if (!normalized) {
      setTestFormulaOutput(null);
      setLastEvaluatedFormula("");
      setTestStatus("");
      return;
    }
    if (lastEvaluatedFormula && normalized !== lastEvaluatedFormula) {
      setTestFormulaOutput(null);
      setLastEvaluatedFormula("");
      setTestStatusType("success");
      setTestStatus("Formula editor changed. Click Calculate to refresh output.");
    }
  }, [editorFormulaText, lastEvaluatedFormula]);

  const workspacePrimaryAction = useMemo(() => {
    if (activeWorkspace === "structure") {
      return (
        <Button
          size="small"
          appearance="primary"
          onClick={() => openSection("structure", "namedRanges")}
        >
          Open Names
        </Button>
      );
    }
    if (activeWorkspace === "formulas") {
      return null;
    }
    if (activeWorkspace === "data") {
      return (
        <Button
          size="small"
          appearance="primary"
          onClick={() => openSection("data", "queryStatus")}
        >
          Open Queries
        </Button>
      );
    }
    if (activeWorkspace === "model") {
      return (
        <Button
          size="small"
          appearance="primary"
          onClick={() => openSection("model", "parameterBuilder")}
        >
          Open Builder
        </Button>
      );
    }
    if (activeWorkspace === "layout") {
      return (
        <Button
          size="small"
          appearance="primary"
          onClick={() => openSection("layout", "gridFreeze")}
        >
          Open Layout
        </Button>
      );
    }
    return null;
  }, [activeWorkspace]);

  const renderWorkspaceActionBar = () => {
    if (activeWorkspace === "structure") {
      return (
        <div className={styles.actionBar}>
          <Button
            size="small"
            appearance={activeRoute.sectionId === "namedRanges" ? "primary" : "secondary"}
            onClick={() => openSection("structure", "namedRanges")}
          >
            Named Ranges
          </Button>
          <Button
            size="small"
            appearance={activeRoute.sectionId === "tables" ? "primary" : "secondary"}
            onClick={() => openSection("structure", "tables")}
          >
            Tables
          </Button>
          <Text className={styles.actionHint}>
            Bulk rename and cleanup tools stay inside the Names and Tables views.
          </Text>
        </div>
      );
    }
    if (activeWorkspace === "formulas") {
      return (
        <div className={styles.actionBar}>
          <Button size="small" onClick={() => runFormulaAction("pull")}>
            Pull Active Formula
          </Button>
          <Button size="small" onClick={() => runFormulaAction("apply")}>
            Apply To Cell
          </Button>
          <Button size="small" onClick={() => runFormulaAction("beautify")}>
            Beautify
          </Button>
          <Button size="small" onClick={() => runFormulaAction("insertSelection")}>
            Use Current Selection
          </Button>
          <Text className={styles.actionHint}>
            The formula editor is the source of truth. Calculate output below uses the current
            editor text automatically.
          </Text>
        </div>
      );
    }
    if (activeWorkspace === "data") {
      return (
        <div className={styles.actionBar}>
          <Button
            size="small"
            appearance={activeRoute.sectionId === "queryStatus" ? "primary" : "secondary"}
            onClick={() => openSection("data", "queryStatus")}
          >
            Query Status
          </Button>
          <Button
            size="small"
            appearance={activeRoute.sectionId === "pivotRefresh" ? "primary" : "secondary"}
            onClick={() => openSection("data", "pivotRefresh")}
          >
            Pivot Refresh
          </Button>
          <Text className={styles.actionHint}>
            Query refresh, warnings, and diagnostics are inside Query Status.
          </Text>
        </div>
      );
    }
    if (activeWorkspace === "model") {
      return (
        <div className={styles.actionBar}>
          <Button
            size="small"
            appearance={activeRoute.sectionId === "parameterBuilder" ? "primary" : "secondary"}
            onClick={() => openSection("model", "parameterBuilder")}
          >
            Parameter Builder
          </Button>
          <Text className={styles.actionHint}>
            Validation lists, defaults, and the apply summary are part of the builder workflow.
          </Text>
        </div>
      );
    }
    if (activeWorkspace === "layout") {
      return (
        <div className={styles.actionBar}>
          <Button
            size="small"
            appearance="secondary"
            onClick={() => openSection("layout", "gridFreeze")}
          >
            Gridlines and Freeze
          </Button>
          <Text className={styles.actionHint}>
            Presentation helpers and templates are grouped below.
          </Text>
        </div>
      );
    }
    return (
      <div className={styles.actionBar}>
        <Button
          size="small"
          appearance={activeRoute.sectionId === "settings" ? "primary" : "secondary"}
          onClick={() => openSection("tools", "settings")}
        >
          Settings
        </Button>
        <Button
          size="small"
          appearance={activeRoute.sectionId === "help" ? "primary" : "secondary"}
          onClick={() => openSection("tools", "help")}
        >
          Help
        </Button>
        <Text className={styles.actionHint}>
          Diagnostics, onboarding, and migration controls stay in this workspace.
        </Text>
      </div>
    );
  };

  const renderWorkspaceSections = () => {
    if (activeWorkspace === "structure") {
      return (
        <div className={styles.sectionStack}>
          {renderSection(
            "structure",
            "namedRanges",
            "Named Ranges",
            <BookNumber20Regular />,
            <NamesView
              createRequestId={createRequestId}
              onOpenLegacy={openLegacy}
              onOpenShape={openShapeEditor}
              embedded
            />,
            { statusChip: "Active" }
          )}
          {renderSection(
            "structure",
            "tables",
            "Tables",
            <Table20Regular />,
            <TablesView onOpenLegacy={openLegacy} embedded />
          )}
        </div>
      );
    }

    if (activeWorkspace === "formulas") {
      return (
        <div className={styles.sectionStack}>
          {renderSection(
            "formulas",
            "formulaEditor",
            "Formula Editor",
            <DataArea20Regular />,
            <FormulaMonacoView
              ref={formulaViewRef}
              isPopout={false}
              onOpenLegacy={openLegacy}
              embedded
              onFormulaChange={setEditorFormulaText}
            />,
            { statusChip: "Live" }
          )}
          {renderSection(
            "formulas",
            "evaluateTest",
            "Calculate Output",
            <DataPie20Regular />,
            <div className={styles.evaluatePanel}>
              <Text className={styles.placeholderText}>Editor Formula</Text>
              <div className={styles.evaluateFormulaPreview}>
                <pre className={styles.evaluateFormulaCode}>
                  {editorFormulaText.trim() || "Formula editor is empty."}
                </pre>
              </div>
              <div className={styles.evaluateRow}>
                <Text className={styles.evaluateHint}>
                  Calculate always uses the formula currently shown in Formula Editor.
                </Text>
                <Button
                  size="small"
                  appearance="primary"
                  onClick={() => void runManualFormulaTest()}
                  disabled={testBusy}
                >
                  Calculate
                </Button>
              </div>
              <div className={styles.evaluateOutputWrap}>
                {testFormulaOutput ? (
                  <table className={styles.evaluateOutputTable}>
                    <tbody>
                      {testFormulaOutput.values.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {row.map((cell, colIndex) => (
                            <td
                              key={`${rowIndex}-${colIndex}`}
                              className={styles.evaluateOutputCell}
                            >
                              {cell === null ? "" : String(cell)}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className={styles.evaluateOutputCell}>
                    <Text className={styles.placeholderText}>
                      Results will render here after running the test.
                    </Text>
                  </div>
                )}
              </div>
              {testStatus ? (
                <Text
                  className={`${styles.evaluateStatus} ${testStatusType === "error" ? styles.evaluateStatusError : ""}`}
                >
                  {testStatus}
                </Text>
              ) : null}
            </div>
          )}
        </div>
      );
    }

    if (activeWorkspace === "data") {
      return (
        <div className={styles.sectionStack}>
          {renderSection(
            "data",
            "queryStatus",
            "Query Status",
            <DataPie20Regular />,
            <QueriesView onOpenLegacy={openLegacy} embedded />,
            { statusChip: "Live" }
          )}
          {renderSection(
            "data",
            "pivotRefresh",
            "Pivot Refresh",
            <Table20Regular />,
            <PivotsView onOpenLegacy={openLegacy} embedded />
          )}
        </div>
      );
    }

    if (activeWorkspace === "model") {
      return (
        <div className={styles.sectionStack}>
          {renderSection(
            "model",
            "parameterBuilder",
            "Parameter Builder",
            <TextAlignJustify20Regular />,
            <ModelBuilderView onOpenLegacy={openLegacy} embedded />,
            { statusChip: "Ready" }
          )}
        </div>
      );
    }

    if (activeWorkspace === "layout") {
      return (
        <div className={styles.sectionStack}>
          {renderSection(
            "layout",
            "gridFreeze",
            "Gridlines and Freeze Panes",
            <Grid20Regular />,
            <FormatView
              onOpenLegacy={openLegacy}
              selectionRequest={formatSelectionRequest}
              embedded
            />,
            { statusChip: "Active" }
          )}
        </div>
      );
    }

    return (
      <div className={styles.sectionStack}>
        {renderSection(
          "tools",
          "debugUtilities",
          "Debug Utilities",
          <Box20Regular />,
          <SandboxDebugView onOpenLegacy={openLegacy} embedded />
        )}
        {renderSection(
          "tools",
          "settings",
          "Settings",
          <Settings20Regular />,
          <SettingsView
            onOpenLegacy={openLegacy}
            onResetUiPreference={onResetUiPreference}
            embedded
          />,
          { statusChip: "Active" }
        )}
        {renderSection(
          "tools",
          "help",
          "Help",
          <QuestionCircle20Regular />,
          <HelpView onNavigate={onTargetChange} embedded />
        )}
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
