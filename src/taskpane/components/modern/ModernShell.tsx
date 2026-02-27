import * as React from "react";
import { useMemo, useState } from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
import {
  BookNumber20Regular,
  Box20Regular,
  DataArea20Regular,
  DataPie20Regular,
  Grid20Regular,
  MoreHorizontal20Regular,
  QuestionCircle20Regular,
  Settings20Regular,
  Table20Regular,
} from "@fluentui/react-icons";
import { NavigationTarget } from "../../navigation";
import FormulaMonacoView, { FormulaViewHandle } from "./FormulaMonacoView";
import NamesView from "./NamesView";
import TablesView from "./TablesView";
import SandboxDebugView from "./SandboxDebugView";
import FormatView from "./FormatView";
import PivotsView from "./PivotsView";
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

interface NavItem {
  id: NavigationTarget;
  label: string;
  icon: React.ReactNode;
  enabled: boolean;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: MODERN_TOKENS.colorBg,
    color: MODERN_TOKENS.colorText,
    display: "grid",
    gridTemplateRows: "56px 1fr",
  },
  header: {
    height: "56px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    display: "flex",
    alignItems: "center",
    padding: "0 16px",
    gap: "12px",
    position: "relative",
  },
  headerTitle: { fontWeight: 700, fontSize: "16px" },
  main: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    minHeight: 0,
  },
  navRail: {
    borderRight: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "12px 8px",
    display: "grid",
    alignContent: "start",
    gap: "8px",
    transition: "width 150ms ease",
  },
  navButton: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorText,
    borderRadius: "8px",
    padding: "8px 10px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
    cursor: "pointer",
    width: "100%",
    fontSize: "12px",
    textAlign: "left",
  },
  navButtonActive: {
    backgroundColor: "#E8F4EA",
    color: MODERN_TOKENS.colorBrandStrong,
    fontWeight: 700,
  },
  navButtonDisabled: {
    opacity: 0.45,
    cursor: "not-allowed",
  },
  navLabelCollapsed: { display: "none" },
  content: {
    padding: "24px",
    overflow: "auto",
    minHeight: 0,
  },
  menuPopover: {
    position: "absolute",
    top: "48px",
    right: "12px",
    width: "190px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCardHover,
    padding: "8px",
    display: "grid",
    gap: "4px",
    zIndex: 10,
  },
  menuBtn: {
    border: "none",
    background: "transparent",
    textAlign: "left",
    borderRadius: "6px",
    padding: "8px",
    cursor: "pointer",
    selectors: {
      "&:hover": { backgroundColor: "#F3F4F6" },
    },
  },
});

const ModernShell: React.FC<ModernShellProps> = ({
  activeTarget,
  createRequestId,
  onTargetChange,
  onSwitchToLegacy,
  onResetUiPreference,
  formulaViewRef,
}) => {
  const styles = useStyles();
  const [expandedRail, setExpandedRail] = useState<boolean>(false);
  const [menuOpen, setMenuOpen] = useState<boolean>(false);
  const isFormulaTarget = activeTarget === "formulas";

  const navItems = useMemo<NavItem[]>(
    () => [
      { id: "names", label: "Names", icon: <BookNumber20Regular />, enabled: true },
      { id: "tables", label: "Tables", icon: <Table20Regular />, enabled: true },
      { id: "formulas", label: "Formulas", icon: <DataArea20Regular />, enabled: true },
      { id: "pivots", label: "Pivots", icon: <DataPie20Regular />, enabled: true },
      { id: "format", label: "Format", icon: <Grid20Regular />, enabled: true },
      { id: "sandbox-debug", label: "Sandbox", icon: <Box20Regular />, enabled: true },
    ],
    []
  );

  const openLegacy = () => {
    onSwitchToLegacy();
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button size="small" onClick={() => setExpandedRail((prev) => !prev)}>
          {expandedRail ? "Collapse" : "Expand"}
        </Button>
        <Text className={styles.headerTitle}>Workbook Manager</Text>
        <div style={{ flexGrow: 1 }} />
        <Button icon={<MoreHorizontal20Regular />} onClick={() => setMenuOpen((prev) => !prev)} />
        {menuOpen ? (
          <div className={styles.menuPopover}>
            <button type="button" className={styles.menuBtn} onClick={() => { onTargetChange("settings"); setMenuOpen(false); }}>
              <Settings20Regular /> Settings
            </button>
            <button type="button" className={styles.menuBtn} onClick={() => { onTargetChange("help"); setMenuOpen(false); }}>
              <QuestionCircle20Regular /> Help
            </button>
            <button type="button" className={styles.menuBtn} onClick={() => { openLegacy(); setMenuOpen(false); }}>
              Open Legacy UI
            </button>
          </div>
        ) : null}
      </div>

      <div className={styles.main}>
        <div
          className={styles.navRail}
          style={{ width: expandedRail ? (isFormulaTarget ? "128px" : "160px") : isFormulaTarget ? "56px" : "72px" }}
        >
          {navItems.map((item) => {
            const selected =
              activeTarget === item.id ||
              (item.id === "names" && activeTarget === "names-create");
            const className = `${styles.navButton} ${selected ? styles.navButtonActive : ""} ${
              !item.enabled ? styles.navButtonDisabled : ""
            }`;
            return (
              <button
                key={item.id}
                type="button"
                className={className}
                disabled={!item.enabled}
                title={expandedRail ? "" : item.label}
                onClick={() => item.enabled && onTargetChange(item.id)}
              >
                {item.icon}
                <span className={expandedRail ? "" : styles.navLabelCollapsed}>{item.label}</span>
              </button>
            );
          })}
        </div>

        <div className={styles.content} style={{ padding: isFormulaTarget ? "12px 16px" : "24px" }}>
          {activeTarget === "names" || activeTarget === "names-create" ? (
            <NamesView createRequestId={createRequestId} onOpenLegacy={openLegacy} />
          ) : null}
          {activeTarget === "tables" ? <TablesView onOpenLegacy={openLegacy} /> : null}
          {activeTarget === "formulas" ? (
            <FormulaMonacoView ref={formulaViewRef} isPopout={false} onOpenLegacy={openLegacy} />
          ) : null}
          {activeTarget === "format" ? <FormatView onOpenLegacy={openLegacy} /> : null}
          {activeTarget === "sandbox-debug" ? <SandboxDebugView onOpenLegacy={openLegacy} /> : null}
          {activeTarget === "pivots" ? <PivotsView onOpenLegacy={openLegacy} /> : null}
          {activeTarget === "settings" ? (
            <SettingsView onOpenLegacy={openLegacy} onResetUiPreference={onResetUiPreference} />
          ) : null}
          {activeTarget === "help" ? <HelpView onNavigate={onTargetChange} /> : null}
        </div>
      </div>
    </div>
  );
};

export default ModernShell;
