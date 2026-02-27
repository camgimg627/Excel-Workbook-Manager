import * as React from "react";
import WorkbookDataTable, { WorkbookTableColumn } from "./WorkbookDataTable";
import { getWorkbookSheets, WorkbookSheetRecord } from "../workbookData";

const columns: WorkbookTableColumn[] = [
  { key: "name", label: "Sheet name" },
  { key: "position", label: "Position" },
  { key: "visibility", label: "Visibility" },
];

const SheetsTab: React.FC = () => {
  const [rows, setRows] = React.useState<WorkbookSheetRecord[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    const load = async () => {
      setLoading(true);
      try {
        const workbookSheets = await getWorkbookSheets();
        setRows(workbookSheets);
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
      emptyMessage="No workbook sheets were found."
      searchPlaceholder="Search sheets"
    />
  );
};

export default SheetsTab;
