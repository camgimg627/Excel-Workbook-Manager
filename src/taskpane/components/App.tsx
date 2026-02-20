import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import Header from "./Header";
import TaskTabs, { TaskTab } from "./TaskTabs";
import NamesTab from "./NamesTab";
import SheetsTab from "./SheetsTab";
import FormulasTab from "./FormulasTab";
import FormattingTab from "./FormattingTab";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  tabsContainer: {
    padding: "0 20px",
  },
  contentContainer: {
    padding: "0 20px 20px",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [activeTab, setActiveTab] = React.useState<TaskTab>("names");

  const renderActiveTab = (): React.ReactNode => {
    switch (activeTab) {
      case "names":
        return <NamesTab />;
      case "sheets":
        return <SheetsTab />;
      case "formulas":
        return <FormulasTab />;
      case "formatting":
        return <FormattingTab />;
      default:
        return null;
    }
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <div className={styles.tabsContainer}>
        <TaskTabs activeTab={activeTab} onTabSelect={setActiveTab} />
      </div>
      <div className={styles.contentContainer}>{renderActiveTab()}</div>
    </div>
  );
};

export default App;
