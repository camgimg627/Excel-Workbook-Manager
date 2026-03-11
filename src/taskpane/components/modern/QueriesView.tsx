import * as React from "react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { Button, Text, makeStyles } from "@fluentui/react-components";
import { listWorkbookQueries, refreshWorkbookQueries, WorkbookQueryRecord } from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

/* global Office */

interface QueriesViewProps {
  onOpenLegacy: () => void;
  embedded?: boolean;
}

const useStyles = makeStyles({
  root: { display: "grid", gap: "24px" },
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  spacer: { flexGrow: 1 },
  tableWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    backgroundColor: "#fff",
    overflow: "auto",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    minWidth: "860px",
    fontSize: "12px",
  },
  th: {
    position: "sticky",
    top: 0,
    zIndex: 1,
    textAlign: "left",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F3F4F6",
    padding: "10px 8px",
    whiteSpace: "nowrap",
  },
  td: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px 8px",
    verticalAlign: "top",
  },
  muted: { color: MODERN_TOKENS.colorTextMuted },
});

const isQueryApiSupported = (): boolean => {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.14");
  } catch {
    return false;
  }
};

const QueriesView: React.FC<QueriesViewProps> = ({ onOpenLegacy, embedded = false }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [queries, setQueries] = useState<WorkbookQueryRecord[]>([]);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [busy, setBusy] = useState<boolean>(false);

  const queryApiAvailable = useMemo(() => isQueryApiSupported(), []);
  const statusClass = statusType === "success" ? shared.successText : shared.errorText;

  const load = useCallback(async () => {
    setBusy(true);
    try {
      const next = await listWorkbookQueries();
      setQueries(next);
      setStatusType("success");
      setStatus(`Loaded ${next.length.toString()} query record(s).`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatusType("error");
      setStatus(`Load failed: ${message}`);
    } finally {
      setBusy(false);
    }
  }, []);

  const refresh = useCallback(async () => {
    setBusy(true);
    try {
      const result = await refreshWorkbookQueries();
      await load();
      const warningText = result.warnings.length ? ` ${result.warnings.join(" ")}` : "";
      setStatusType(result.refreshed ? "success" : "error");
      setStatus(
        `${result.refreshed ? "Refresh request sent." : "Refresh unavailable."} Method: ${result.method}.${warningText}`
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatusType("error");
      setStatus(`Refresh failed: ${message}`);
    } finally {
      setBusy(false);
    }
  }, [load]);

  useEffect(() => {
    void load();
  }, [load]);

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Queries</Text>
          <Text className={shared.sectionSubtitle}>List and refresh workbook Power Query metadata where supported.</Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button appearance="primary" onClick={() => void refresh()} disabled={busy}>
            Refresh Queries
          </Button>
          <Button onClick={() => void load()} disabled={busy}>
            Reload List
          </Button>
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
          <div className={styles.spacer} />
          <Text className={queryApiAvailable ? shared.mutedText : shared.errorText}>
            {queryApiAvailable
              ? "ExcelApi 1.14 query metadata supported."
              : "ExcelApi 1.14 is unavailable on this host. Query list may be incomplete."}
          </Text>
        </div>
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Name</th>
              <th className={styles.th}>Loaded To</th>
              <th className={styles.th}>Data Model</th>
              <th className={styles.th}>Last Refresh</th>
              <th className={styles.th}>Rows Loaded</th>
              <th className={styles.th}>Last Error</th>
            </tr>
          </thead>
          <tbody>
            {!queries.length ? (
              <tr>
                <td className={`${styles.td} ${styles.muted}`} colSpan={6}>
                  No query metadata records returned.
                </td>
              </tr>
            ) : null}
            {queries.map((query) => (
              <tr key={query.name}>
                <td className={styles.td}>{query.name}</td>
                <td className={styles.td}>{query.loadedTo || "-"}</td>
                <td className={styles.td}>{query.loadedToDataModel ? "Yes" : "No"}</td>
                <td className={styles.td}>{query.refreshDate || "-"}</td>
                <td className={styles.td}>{query.rowsLoadedCount.toString()}</td>
                <td className={styles.td}>{query.error || "-"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {status ? <Text className={statusClass}>{status}</Text> : null}
    </div>
  );
};

export default QueriesView;
