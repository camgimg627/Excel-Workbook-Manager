import * as React from "react";
import { useState } from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
import { refreshPivotTables } from "../../taskpane";
import { useModernSharedStyles } from "./designTokens";

interface PivotsViewProps {
  onOpenLegacy: () => void;
}

const useStyles = makeStyles({
  root: {
    display: "grid",
    gap: "20px",
  },
  actions: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
});

const PivotsView: React.FC<PivotsViewProps> = ({ onOpenLegacy }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");

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
      <div>
        <Text className={shared.sectionTitle}>Pivots</Text>
        <Text className={shared.sectionSubtitle}>Refresh pivot table calculations on the active worksheet.</Text>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Pivot Table Actions</Text>
        <Text className={shared.cardSubtitle}>
          Use this action after data updates to keep pivot outputs in sync.
        </Text>
        <div className={styles.actions}>
          <Button appearance="primary" onClick={() => void runAction("Refresh pivot tables", refreshPivotTables)}>
            Refresh Pivot Tables
          </Button>
          <Button onClick={onOpenLegacy}>Open Legacy UI</Button>
        </div>
      </div>

      {status ? (
        <Text className={statusType === "success" ? shared.successText : shared.errorText}>{status}</Text>
      ) : null}
    </div>
  );
};

export default PivotsView;

