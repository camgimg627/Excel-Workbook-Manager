import * as React from "react";
import { useState } from "react";
import { Button, Select, Text, makeStyles } from "@fluentui/react-components";
import {
  CellStylePreset,
  addRectangleShape,
  applyAccentFill,
  applyCellStylePreset,
  applyTableStyle,
  freezeFirstColumn,
  freezeTopRow,
  refreshPivotTables,
  toggleGridlines,
  unfreezePanes,
} from "../../taskpane";
import LegacyApp from "../LegacyApp";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

interface FormatViewProps {
  onOpenLegacy: () => void;
}

const TABLE_STYLES = ["TableStyleMedium2", "TableStyleMedium9", "TableStyleLight11"];

const STYLE_PRESETS: Array<{ preset: CellStylePreset; description: string }> = [
  {
    preset: "Input Cell",
    description: "Light input shading for user-entered values.",
  },
  {
    preset: "Parameter Cell",
    description: "Emphasized parameter values with stronger font treatment.",
  },
  {
    preset: "Header",
    description: "Primary table/report header styling.",
  },
  {
    preset: "Subheader",
    description: "Secondary section heading styling.",
  },
];

const useStyles = makeStyles({
  root: { display: "grid", gap: "24px" },
  cards: {
    display: "grid",
    gap: "20px",
  },
  styleGrid: {
    display: "grid",
    gap: "12px",
    gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))",
  },
  styleCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    padding: "12px",
    display: "grid",
    gap: "8px",
  },
  styleTitle: { fontWeight: 700, fontSize: "13px" },
  actions: { display: "flex", flexWrap: "wrap", gap: "8px" },
  advancedWrap: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    paddingTop: "12px",
    marginTop: "4px",
  },
});

const FormatView: React.FC<FormatViewProps> = ({ onOpenLegacy }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [tableStyle, setTableStyle] = useState<string>(TABLE_STYLES[0]);
  const [showAdvancedLegacy, setShowAdvancedLegacy] = useState<boolean>(false);

  const runAction = async (label: string, action: () => Promise<void>) => {
    setStatus("");
    try {
      await action();
      setStatusType("success");
      setStatus(`${label} completed.`);
    } catch (error) {
      setStatusType("error");
      setStatus(`${label} failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  };

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  return (
    <div className={styles.root}>
      <div>
        <Text className={shared.sectionTitle}>Format</Text>
        <Text className={shared.sectionSubtitle}>
          Native modern formatting tools are available below. Advanced shape tooling remains in bridge mode.
        </Text>
      </div>

      <div className={styles.cards}>
        <div className={shared.card}>
          <Text className={shared.cardTitle}>Default Cell Styles</Text>
          <Text className={shared.cardSubtitle}>Select a range in Excel, then apply a style preset.</Text>
          <div className={styles.styleGrid}>
            {STYLE_PRESETS.map((item) => (
              <div key={item.preset} className={styles.styleCard}>
                <Text className={styles.styleTitle}>{item.preset}</Text>
                <Text className={shared.mutedText}>{item.description}</Text>
                <Button
                  onClick={() =>
                    void runAction(`Apply ${item.preset}`, () => applyCellStylePreset(item.preset))
                  }
                >
                  Apply Style
                </Button>
              </div>
            ))}
          </div>
        </div>

        <div className={shared.card}>
          <Text className={shared.cardTitle}>Gridlines and Freeze Panes</Text>
          <Text className={shared.cardSubtitle}>Use quick actions for common view formatting workflows.</Text>
          <div className={styles.actions}>
            <Button onClick={() => void runAction("Toggle gridlines", toggleGridlines)}>
              Toggle Gridlines
            </Button>
            <Button onClick={() => void runAction("Freeze top row", freezeTopRow)}>Freeze Top Row</Button>
            <Button onClick={() => void runAction("Freeze first column", freezeFirstColumn)}>
              Freeze First Column
            </Button>
            <Button onClick={() => void runAction("Unfreeze panes", unfreezePanes)}>Unfreeze</Button>
          </div>
        </div>

        <div className={shared.card}>
          <Text className={shared.cardTitle}>Tables and Pivot Tables</Text>
          <Text className={shared.cardSubtitle}>Apply table styles and refresh pivots on the active sheet.</Text>
          <div className={styles.actions}>
            <Select value={tableStyle} onChange={(_, data) => setTableStyle(data.value)}>
              {TABLE_STYLES.map((style) => (
                <option key={style} value={style}>
                  {style}
                </option>
              ))}
            </Select>
            <Button onClick={() => void runAction("Apply table style", () => applyTableStyle(tableStyle))}>
              Apply Table Style
            </Button>
            <Button onClick={() => void runAction("Refresh pivots", refreshPivotTables)}>
              Refresh Pivot Tables
            </Button>
          </div>
        </div>

        <div className={shared.card}>
          <Text className={shared.cardTitle}>Quick Shape Action</Text>
          <Text className={shared.cardSubtitle}>Insert rectangle and apply accent fill in one action.</Text>
          <div className={styles.actions}>
            <Button
              onClick={() =>
                void runAction("Insert accent shape", async () => {
                  await addRectangleShape();
                  await applyAccentFill();
                })
              }
            >
              Insert Accent Shape
            </Button>
          </div>
        </div>

        <div className={shared.card}>
          <Text className={shared.cardTitle}>Advanced Shape Tools (Bridge)</Text>
          <Text className={shared.cardSubtitle}>
            Open the legacy formatting panel for full shape editor capabilities while migration continues.
          </Text>
          <div className={styles.actions}>
            <Button appearance="primary" onClick={() => setShowAdvancedLegacy((prev) => !prev)}>
              {showAdvancedLegacy ? "Hide Advanced Legacy Panel" : "Show Advanced Legacy Panel"}
            </Button>
            <Button onClick={onOpenLegacy}>Open Full Legacy UI</Button>
          </div>
          {showAdvancedLegacy ? (
            <div className={styles.advancedWrap}>
              <LegacyApp initialTarget="format" bridgeMode onExitBridge={() => setShowAdvancedLegacy(false)} />
            </div>
          ) : null}
        </div>
      </div>

      {status ? <Text className={statusClass}>{status}</Text> : null}
    </div>
  );
};

export default FormatView;

