import * as React from "react";
import WorkbookDataTable, { WorkbookTableColumn } from "./WorkbookDataTable";
import { getNamedFormulas, NamedFormulaRecord } from "../workbookData";

const columns: WorkbookTableColumn[] = [
  { key: "name", label: "Name" },
  { key: "formula", label: "Formula" },
  { key: "value", label: "Value" },
];

const FormulasTab: React.FC = () => {
  const [rows, setRows] = React.useState<NamedFormulaRecord[]>([]);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    const load = async () => {
      setLoading(true);
      try {
        const namedFormulas = await getNamedFormulas();
        setRows(namedFormulas);
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
      emptyMessage="No named formulas were found."
      searchPlaceholder="Search named formulas"
    />
  );
};

export default FormulasTab;
