import * as React from "react";
import { useMemo } from "react";
import { Badge, Button, Input, Text, makeStyles } from "@fluentui/react-components";
import {
  BookNumber20Regular,
  Box20Regular,
  DataArea20Regular,
  DataPie20Regular,
  Filter20Regular,
  Grid20Regular,
  Search20Regular,
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

interface ModernShellProps {
  activeTarget: NavigationTarget;
  createRequestId: number;
  onTargetChange: (target: NavigationTarget) => void;
  onSwitchToLegacy: () => void;
  onResetUiPreference: () => void;
  formulaViewRef: React.RefObject<FormulaViewHandle>;
}

type WorkspaceKey = "structure" | "formulas" | "data" | "model" | "layout" | "tools";

interface WorkspaceItem {
  id: WorkspaceKey;
  label: string;
  icon: React.ReactNode;
  defaultTarget: NavigationTarget;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
    backgroundColor: MODERN_TOKENS.colorBg,
    color: MODERN_TOKENS.colorText,
    display: "grid",
    gridTemplateRows: "auto auto auto 1fr auto",
    overflow: "hidden",
  },
  header: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "12px 16px 10px",
    display: "grid",
    gap: "2px",
  },
  productTitle: {
    fontWeight: 700,
    fontSize: "14px",
    lineHeight: "20px",
  },
  productSubtitle: {
    fontSize: "12px",
    color: MODERN_TOKENS.colorTextMuted,
    lineHeight: "16px",
  },
  workspaceSwitcher: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "8px 12px",
    display: "grid",
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    gap: "8px",
  },
  workspaceButton: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    color: MODERN_TOKENS.colorText,
    fontSize: "12px",
    fontWeight: 600,
    padding: "8px 6px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    gap: "6px",
    cursor: "pointer",
  },
  workspaceButtonActive: {
    border: `1px solid ${MODERN_TOKENS.colorBrand}`,
    backgroundColor: "#E8F4EA",
    color: MODERN_TOKENS.colorBrandStrong,
  },
  contextBar: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "8px 12px",
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
  },
  contextTop: {
    display: "grid",
    gap: "8px",
    minWidth: 0,
  },
  contextTitleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  workspaceTitle: {
    fontSize: "13px",
    fontWeight: 700,
  },
  subnav: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
  },
  chip: {
    borderRadius: "999px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    fontSize: "11px",
    padding: "2px 10px",
    cursor: "pointer",
  },
  chipActive: {
    backgroundColor: "#F3F8F4",
    border: `1px solid ${MODERN_TOKENS.colorBrand}`,
    color: MODERN_TOKENS.colorBrandStrong,
    fontWeight: 700,
  },
  quickActions: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  content: {
    padding: "16px",
    overflowY: "auto",
    overflowX: "hidden",
    minHeight: 0,
  },
  footer: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    padding: "6px 12px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
});

const workspaceForTarget = (target: NavigationTarget): WorkspaceKey => {
  if (target === "names" || target === "names-create" || target === "tables") return "structure";
  if (target === "formulas") return "formulas";
  if (target === "queries" || target === "pivots") return "data";
  if (target === "model-builder") return "model";
  if (target === "format") return "layout";
  return "tools";
};

const ModernShell: React.FC<ModernShellProps> = ({
  activeTarget,
  createRequestId,
  onTargetChange,
  onSwitchToLegacy,
  onResetUiPreference,
  formulaViewRef,
}) => {
  const styles = useStyles();

  const workspaces = useMemo<WorkspaceItem[]>(
    () => [
      { id: "structure", label: "Structure", icon: <BookNumber20Regular />, defaultTarget: "names" },
      { id: "formulas", label: "Formulas", icon: <DataArea20Regular />, defaultTarget: "formulas" },
      { id: "data", label: "Data", icon: <Filter20Regular />, defaultTarget: "queries" },
      { id: "model", label: "Model", icon: <TextAlignJustify20Regular />, defaultTarget: "model-builder" },
      { id: "layout", label: "Layout", icon: <Grid20Regular />, defaultTarget: "format" },
      { id: "tools", label: "Tools", icon: <Box20Regular />, defaultTarget: "sandbox-debug" },
    ],
    []
  );

  const activeWorkspace = workspaceForTarget(activeTarget);
  const openLegacy = () => onSwitchToLegacy();

  const isFormulaTarget = activeTarget === "formulas";
  const isFormatTarget = activeTarget === "format";

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.productTitle}>Workbook Manager</Text>
        <Text className={styles.productSubtitle}>Manage workbook structure, formulas, and data operations</Text>
      </div>

      <div className={styles.workspaceSwitcher}>
        {workspaces.map((workspace) => (
          <button
            key={workspace.id}
            type="button"
            className={`${styles.workspaceButton} ${workspace.id === activeWorkspace ? styles.workspaceButtonActive : ""}`}
            onClick={() => onTargetChange(workspace.defaultTarget)}
          >
            {workspace.icon}
            {workspace.label}
          </button>
        ))}
      </div>

      <div className={styles.contextBar}>
        <div className={styles.contextTop}>
          <div className={styles.contextTitleRow}>
            <Text className={styles.workspaceTitle}>{workspaces.find((w) => w.id === activeWorkspace)?.label}</Text>
            {activeWorkspace === "structure" ? <Badge appearance="outline">Assets</Badge> : null}
            {activeWorkspace === "data" ? <Badge appearance="outline">Refresh Ops</Badge> : null}
          </div>
          <div className={styles.subnav}>
            {activeWorkspace === "structure" ? (
              <>
                <button
                  type="button"
                  className={`${styles.chip} ${(activeTarget === "names" || activeTarget === "names-create") ? styles.chipActive : ""}`}
                  onClick={() => onTargetChange("names")}
                >
                  Names
                </button>
                <button
                  type="button"
                  className={`${styles.chip} ${activeTarget === "tables" ? styles.chipActive : ""}`}
                  onClick={() => onTargetChange("tables")}
                >
                  Tables
                </button>
              </>
            ) : null}
            {activeWorkspace === "data" ? (
              <>
                <button type="button" className={`${styles.chip} ${activeTarget === "queries" ? styles.chipActive : ""}`} onClick={() => onTargetChange("queries")}>
                  Queries
                </button>
                <button type="button" className={`${styles.chip} ${activeTarget === "pivots" ? styles.chipActive : ""}`} onClick={() => onTargetChange("pivots")}>
                  <DataPie20Regular /> Pivots
                </button>
              </>
            ) : null}
            {activeWorkspace === "tools" ? (
              <>
                <button type="button" className={`${styles.chip} ${activeTarget === "sandbox-debug" ? styles.chipActive : ""}`} onClick={() => onTargetChange("sandbox-debug")}>
                  Utilities
                </button>
                <button type="button" className={`${styles.chip} ${activeTarget === "settings" ? styles.chipActive : ""}`} onClick={() => onTargetChange("settings")}>
                  Settings
                </button>
                <button type="button" className={`${styles.chip} ${activeTarget === "help" ? styles.chipActive : ""}`} onClick={() => onTargetChange("help")}>
                  Help
                </button>
              </>
            ) : null}
          </div>
        </div>
        <div className={styles.quickActions}>
          {(activeWorkspace === "structure" || activeWorkspace === "data") ? (
            <Input size="small" contentBefore={<Search20Regular />} placeholder="Search..." />
          ) : null}
          {activeWorkspace === "formulas" ? <Button size="small">Apply</Button> : null}
          {activeWorkspace === "data" ? <Button size="small">Refresh</Button> : null}
          {activeWorkspace === "model" ? <Button size="small">Build</Button> : null}
          {activeWorkspace === "layout" ? <Button size="small">Apply</Button> : null}
          {activeWorkspace === "tools" ? <Button size="small" onClick={openLegacy}>Open Legacy</Button> : null}
        </div>
      </div>

      <div
        className={styles.content}
        style={{ padding: isFormulaTarget ? "12px 16px" : isFormatTarget ? "0 24px 24px" : "16px" }}
      >
        {(activeTarget === "names" || activeTarget === "names-create") ? (
          <NamesView createRequestId={createRequestId} onOpenLegacy={openLegacy} />
        ) : null}
        {activeTarget === "tables" ? <TablesView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "formulas" ? <FormulaMonacoView ref={formulaViewRef} isPopout={false} onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "queries" ? <QueriesView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "model-builder" ? <ModelBuilderView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "format" ? <FormatView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "sandbox-debug" ? <SandboxDebugView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "pivots" ? <PivotsView onOpenLegacy={openLegacy} /> : null}
        {activeTarget === "settings" ? (
          <SettingsView onOpenLegacy={openLegacy} onResetUiPreference={onResetUiPreference} />
        ) : null}
        {activeTarget === "help" ? <HelpView onNavigate={onTargetChange} /> : null}
      </div>

      <div className={styles.footer}>
        <span>Status: Ready</span>
        <span>Modern mode</span>
      </div>
    </div>
  );
};

export default ModernShell;
