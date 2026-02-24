import * as React from "react";
import { useState } from "react";
import { Button, Field, Select, Text, makeStyles, tokens } from "@fluentui/react-components";
import {
  addRectangleShape,
  applyAccentFill,
  applyTableStyle,
  freezeFirstColumn,
  freezeTopRow,
  refreshPivotTables,
  toggleBold,
  toggleGridlines,
  unfreezePanes,
} from "../taskpane";

type StatusType = "success" | "error";

const TABLE_STYLES = ["TableStyleMedium2", "TableStyleMedium9", "TableStyleLight11"];

const useStyles = makeStyles({
  root: {
    display: "grid",
    gap: "12px",
    padding: "16px 20px",
  },
  section: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    display: "grid",
    gap: "8px",
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  row: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  statusSuccess: {
    color: tokens.colorPaletteGreenForeground1,
  },
  statusError: {
    color: tokens.colorPaletteRedForeground1,
  },
});

const FormattingTab: React.FC = () => {
  const styles = useStyles();
  const [tableStyle, setTableStyle] = useState<string>(TABLE_STYLES[0]);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<StatusType>("success");

  const runAction = async (label: string, action: () => Promise<void>) => {
    setStatus("");
    try {
      await action();
      setStatusType("success");
      setStatus(`${label} completed.`);
    } catch (error) {
      setStatusType("error");
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`${label} failed: ${message}`);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Tables</Text>
        <Field label="Table style">
          <Select value={tableStyle} onChange={(_, data) => setTableStyle(data.value)}>
            {TABLE_STYLES.map((style) => (
              <option key={style} value={style}>
                {style}
              </option>
            ))}
          </Select>
        </Field>
        <Button onClick={() => runAction("Apply table style", () => applyTableStyle(tableStyle))}>Apply style</Button>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Pivot Tables</Text>
        <Button onClick={() => runAction("Refresh pivot tables", refreshPivotTables)}>Refresh all on sheet</Button>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Text Boxes / Shapes</Text>
        <Button onClick={() => runAction("Insert rectangle", addRectangleShape)}>Insert rectangle label</Button>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Font / Style Options</Text>
        <div className={styles.row}>
          <Button onClick={() => runAction("Toggle bold", toggleBold)}>Toggle bold</Button>
          <Button onClick={() => runAction("Apply accent fill", applyAccentFill)}>Accent fill</Button>
        </div>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Gridlines</Text>
        <Button onClick={() => runAction("Toggle gridlines", toggleGridlines)}>Toggle gridlines</Button>
      </div>

      <div className={styles.section}>
        <Text className={styles.sectionTitle}>Freeze Panes</Text>
        <div className={styles.row}>
          <Button onClick={() => runAction("Freeze top row", freezeTopRow)}>Freeze top row</Button>
          <Button onClick={() => runAction("Freeze first column", freezeFirstColumn)}>Freeze first column</Button>
          <Button onClick={() => runAction("Unfreeze panes", unfreezePanes)}>Unfreeze</Button>
        </div>
      </div>

      {status ? <Text className={statusType === "success" ? styles.statusSuccess : styles.statusError}>{status}</Text> : null}
    </div>
  );
};

export default FormattingTab;
