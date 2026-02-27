import * as React from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
import { useModernSharedStyles } from "./designTokens";

interface SettingsViewProps {
  onOpenLegacy: () => void;
  onResetUiPreference: () => void;
}

const useStyles = makeStyles({
  root: {
    display: "grid",
    gap: "20px",
  },
  actions: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
  list: {
    margin: "0",
    paddingLeft: "18px",
    display: "grid",
    gap: "6px",
  },
});

const SettingsView: React.FC<SettingsViewProps> = ({ onOpenLegacy, onResetUiPreference }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div>
        <Text className={shared.sectionTitle}>Settings</Text>
        <Text className={shared.sectionSubtitle}>Configure migration and UI behavior controls.</Text>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Migration Controls</Text>
        <ul className={styles.list}>
          <li>
            <Text className={shared.mutedText}>Modern UI is now the default runtime experience.</Text>
          </li>
          <li>
            <Text className={shared.mutedText}>
              Legacy UI remains available as an emergency fallback for one release cycle.
            </Text>
          </li>
          <li>
            <Text className={shared.mutedText}>Query overrides: <code>?ui=modern</code> or <code>?ui=legacy</code>.</Text>
          </li>
        </ul>
        <div className={styles.actions}>
          <Button onClick={onResetUiPreference}>Reset Saved UI Preference</Button>
          <Button onClick={onOpenLegacy}>Open Legacy UI</Button>
        </div>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Command Signals</Text>
        <Text className={shared.cardSubtitle}>
          Ribbon commands communicate with the task pane using document/runtime signal keys.
        </Text>
        <ul className={styles.list}>
          <li>
            <Text className={shared.mutedText}><code>wbm.navigation.request</code> routes target section.</Text>
          </li>
          <li>
            <Text className={shared.mutedText}><code>wbm.openFormulaEditor.request</code> routes formula actions.</Text>
          </li>
        </ul>
      </div>
    </div>
  );
};

export default SettingsView;

