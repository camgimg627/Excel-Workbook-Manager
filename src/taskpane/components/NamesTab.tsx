import * as React from "react";
import WorkbookDataTable, { WorkbookTableColumn } from "./WorkbookDataTable";
import { getWorkbookNames, WorkbookNameRecord } from "../workbookData";

const columns: WorkbookTableColumn[] = [
  { key: "name", label: "Name" },
  { key: "type", label: "Type" },
  { key: "value", label: "Value" },
];

const NamesTab: React.FC = () => {
  const [rows, setRows] = React.useState<WorkbookNameRecord[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    const load = async () => {
      setLoading(true);
      try {
        const workbookNames = await getWorkbookNames();
        setRows(workbookNames);
      } finally {
        setLoading(false);
      }
    };

    void load();
  }, []);

  return (
    <WorkbookDataTable
      columns={columns}
      rows={rows}
      loading={loading}
      emptyMessage="No workbook names were found."
      searchPlaceholder="Search workbook names"
    />
  );
};

export default NamesTab;
