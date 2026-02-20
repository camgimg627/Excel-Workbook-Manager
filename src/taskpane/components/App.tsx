import * as React from "react";
import { Body1Strong, Tab, TabList, makeStyles, tokens } from "@fluentui/react-components";
import NamesTab from "./NamesTab";
import SheetsTab from "./SheetsTab";
import FormulasTab from "./FormulasTab";

interface AppProps {
  title: string;
}

type WorkbookTab = "names" | "sheets" | "formulas";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    minHeight: "100vh",
    padding: tokens.spacingHorizontalL,
  },
  heading: {
    margin: 0,
  },
});

const App: React.FC<AppProps> = ({ title }) => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<WorkbookTab>("names");

  return (
    <div className={styles.root}>
      <Body1Strong as="h1" className={styles.heading}>
        {title}
      </Body1Strong>

      <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as WorkbookTab)}>
        <Tab value="names">Names</Tab>
        <Tab value="sheets">Sheets</Tab>
        <Tab value="formulas">Formulas</Tab>
      </TabList>

      {selectedTab === "names" && <NamesTab />}
      {selectedTab === "sheets" && <SheetsTab />}
      {selectedTab === "formulas" && <FormulasTab />}
    </div>
  );
};

export default App;
