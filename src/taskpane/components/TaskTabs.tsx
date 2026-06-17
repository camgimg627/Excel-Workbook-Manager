import * as React from "react";
import { Tab, TabList } from "@fluentui/react-components";

export type TaskTab = "names" | "sheets" | "formulas" | "formatting";

interface TaskTabsProps {
  activeTab: TaskTab;
  onTabSelect: (tab: TaskTab) => void;
}

const TaskTabs: React.FC<TaskTabsProps> = ({ activeTab, onTabSelect }) => {
  return (
    <TabList selectedValue={activeTab} onTabSelect={(_, data) => onTabSelect(data.value as TaskTab)}>
      <Tab value="names">Names</Tab>
      <Tab value="sheets">Sheets</Tab>
      <Tab value="formulas">Formulas</Tab>
      <Tab value="formatting">Formatting</Tab>
    </TabList>
  );
};

export default TaskTabs;
