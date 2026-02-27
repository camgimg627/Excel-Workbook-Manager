import * as React from "react";
import {
  Field,
  Input,
  Spinner,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

export interface WorkbookTableColumn {
  key: string;
  label: string;
}

interface WorkbookDataTableProps {
  columns: WorkbookTableColumn[];
  rows: object[];
  loading?: boolean;
  emptyMessage?: string;
  searchPlaceholder?: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  headerCell: {
    cursor: "pointer",
    userSelect: "none",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
  },
  emptyState: {
    textAlign: "center",
    paddingTop: tokens.spacingVerticalXXL,
    paddingBottom: tokens.spacingVerticalXXL,
    color: tokens.colorNeutralForeground3,
  },
});

const toComparableText = (value: unknown): string => {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  return String(value);
};

const WorkbookDataTable: React.FC<WorkbookDataTableProps> = ({
  columns,
  rows,
  loading = false,
  emptyMessage = "No rows to display.",
  searchPlaceholder = "Search all columns",
}) => {
  const styles = useStyles();
  const [searchText, setSearchText] = React.useState("");
  const [sortColumn, setSortColumn] = React.useState<string | undefined>(undefined);
  const [sortDirection, setSortDirection] = React.useState<"asc" | "desc">("asc");

  const filteredAndSortedRows = React.useMemo(() => {
    const normalizedQuery = searchText.trim().toLowerCase();

    const filteredRows = rows.filter((row) => {
      const record = row as Record<string, unknown>;
      if (!normalizedQuery) return true;
      return columns.some((column) => toComparableText(record[column.key]).toLowerCase().includes(normalizedQuery));
    });

    if (!sortColumn) return filteredRows;

    return [...filteredRows].sort((leftRow, rightRow) => {
      const leftRecord = leftRow as Record<string, unknown>;
      const rightRecord = rightRow as Record<string, unknown>;
      const leftValue = toComparableText(leftRecord[sortColumn]).toLowerCase();
      const rightValue = toComparableText(rightRecord[sortColumn]).toLowerCase();
      const comparison = leftValue.localeCompare(rightValue, undefined, { numeric: true, sensitivity: "base" });
      return sortDirection === "asc" ? comparison : -comparison;
    });
  }, [columns, rows, searchText, sortColumn, sortDirection]);

  const onHeaderClick = (columnKey: string) => {
    if (sortColumn === columnKey) {
      setSortDirection((currentDirection) => (currentDirection === "asc" ? "desc" : "asc"));
      return;
    }

    setSortColumn(columnKey);
    setSortDirection("asc");
  };

  return (
    <div className={styles.container}>
      <Field label="Filter rows">
        <Input
          value={searchText}
          onChange={(_, data) => setSearchText(data.value)}
          placeholder={searchPlaceholder}
          contentAfter={<Text className={styles.hint}>Global search</Text>}
        />
      </Field>

      {loading ? (
        <Spinner label="Loading workbook data..." />
      ) : filteredAndSortedRows.length === 0 ? (
        <div className={styles.emptyState}>
          <Text>{emptyMessage}</Text>
        </div>
      ) : (
        <Table aria-label="Workbook data table">
          <TableHeader>
            <TableRow>
              {columns.map((column) => {
                const isSorted = sortColumn === column.key;
                return (
                  <TableHeaderCell
                    key={column.key}
                    className={styles.headerCell}
                    onClick={() => onHeaderClick(column.key)}
                  >
                    {column.label}
                    {isSorted ? ` (${sortDirection})` : ""}
                  </TableHeaderCell>
                );
              })}
            </TableRow>
          </TableHeader>
          <TableBody>
            {filteredAndSortedRows.map((row, rowIndex) => {
              const record = row as Record<string, unknown>;
              return (
                <TableRow key={`${rowIndex}-${columns.map((column) => toComparableText(record[column.key])).join("-")}`}>
                  {columns.map((column) => (
                    <TableCell key={`${rowIndex}-${column.key}`}>
                      <TableCellLayout>{toComparableText(record[column.key])}</TableCellLayout>
                    </TableCell>
                  ))}
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      )}
    </div>
  );
};

export default WorkbookDataTable;
