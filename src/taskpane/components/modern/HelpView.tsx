import * as React from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
import { NavigationTarget } from "../../navigation";
import { useModernSharedStyles } from "./designTokens";

interface HelpViewProps {
  onNavigate: (target: NavigationTarget) => void;
  embedded?: boolean;
}

const useStyles = makeStyles({
  root: { display: "grid", gap: "20px" },
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

const HelpView: React.FC<HelpViewProps> = ({ onNavigate, embedded = false }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Help</Text>
          <Text className={shared.sectionSubtitle}>Quick guidance for the modern Workbook Manager experience.</Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Getting Started</Text>
        <ul className={styles.list}>
          <li>
            <Text className={shared.mutedText}>
              Use <strong>Names</strong> to create and manage workbook names.
            </Text>
          </li>
          <li>
            <Text className={shared.mutedText}>
              Use <strong>Tables</strong> for inline and bulk table naming.
            </Text>
          </li>
          <li>
            <Text className={shared.mutedText}>
              Use <strong>Formulas</strong> for Monaco editing, pull/apply, and testing.
            </Text>
          </li>
          <li>
            <Text className={shared.mutedText}>
              Use <strong>Format</strong> for style, table, and gridline workflows.
            </Text>
          </li>
        </ul>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Formula Shortcuts</Text>
        <ul className={styles.list}>
          <li>
            <Text className={shared.mutedText}><code>Ctrl+Shift+Alt+E</code>: Open + pull active formula</Text>
          </li>
          <li>
            <Text className={shared.mutedText}><code>Ctrl+Shift+Alt+P</code>: Pull formula</Text>
          </li>
          <li>
            <Text className={shared.mutedText}><code>Ctrl+Shift+Alt+A</code>: Apply formula</Text>
          </li>
        </ul>
      </div>

      <div className={shared.card}>
        <Text className={shared.cardTitle}>Jump To Section</Text>
        <div className={styles.actions}>
          <Button onClick={() => onNavigate("names")}>Names</Button>
          <Button onClick={() => onNavigate("tables")}>Tables</Button>
          <Button onClick={() => onNavigate("formulas")}>Formulas</Button>
          <Button onClick={() => onNavigate("format")}>Format</Button>
          <Button onClick={() => onNavigate("pivots")}>Pivots</Button>
        </div>
      </div>
    </div>
  );
};

export default HelpView;
