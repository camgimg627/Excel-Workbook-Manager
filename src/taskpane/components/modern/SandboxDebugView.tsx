import * as React from "react";
import { useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import {
  addRectangleShape,
  applyAccentFill,
  applyTableStyle,
  insertText,
  refreshPivotTables,
  toggleGridlines,
} from "../../taskpane";
import { useModernSharedStyles } from "./designTokens";

interface SandboxDebugViewProps {
  onOpenLegacy: () => void;
  embedded?: boolean;
}

const TABLE_STYLES = ["TableStyleMedium2", "TableStyleMedium9", "TableStyleLight11"];

const useStyles = makeStyles({
  root: { display: "grid", gap: "20px" },
  actions: { display: "flex", flexWrap: "wrap", gap: "8px" },
});

const SandboxDebugView: React.FC<SandboxDebugViewProps> = ({ onOpenLegacy, embedded = false }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [sampleText, setSampleText] = useState<string>("Workbook Manager");
  const [tableStyle, setTableStyle] = useState<string>(TABLE_STYLES[0]);

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

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Sandbox Debug</Text>
          <Text className={shared.sectionSubtitle}>Relocated quick actions for diagnostics and utility workflows.</Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Quick Actions</Text>
        <div className={styles.actions}>
          <Button onClick={() => void runAction("Refresh pivot tables", refreshPivotTables)}>Refresh Pivots</Button>
          <Button onClick={() => void runAction("Toggle gridlines", toggleGridlines)}>Toggle Gridlines</Button>
          <Button
            onClick={() =>
              void runAction("Insert rectangle + accent fill", async () => {
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
        <Text className={shared.cardTitle}>Debug Insert Text</Text>
        <Input value={sampleText} onChange={(_, data) => setSampleText(data.value)} />
        <div className={styles.actions}>
          <Button onClick={() => void runAction("Insert text", async () => insertText(sampleText))}>Insert In A1</Button>
        </div>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Table Style Utility</Text>
        <Select value={tableStyle} onChange={(_, data) => setTableStyle(data.value)}>
          {TABLE_STYLES.map((style) => (
            <option key={style} value={style}>
              {style}
            </option>
          ))}
        </Select>
        <div className={styles.actions}>
          <Button onClick={() => void runAction("Apply table style", () => applyTableStyle(tableStyle))}>
            Apply Style
          </Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
        </div>
      </div>

      {status ? (
        <Text className={statusType === "success" ? shared.successText : shared.errorText}>{status}</Text>
      ) : null}
    </div>
  );
};

export default SandboxDebugView;
