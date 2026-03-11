import * as React from "react";
import { useMemo, useState } from "react";
import { Button, Input, Select, Text, makeStyles } from "@fluentui/react-components";
import {
  applyModelBuilderParameters,
  getCurrentSelectionAddress,
  ModelBuilderApplyResult,
  ModelBuilderParameterInput,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

interface ModelBuilderViewProps {
  onOpenLegacy: () => void;
  embedded?: boolean;
}

interface ParameterRowState extends ModelBuilderParameterInput {
  id: string;
}

const useStyles = makeStyles({
  root: { display: "grid", gap: "24px" },
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  card: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#fff",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    display: "grid",
    gap: "12px",
  },
  rowGrid: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "12px",
    backgroundColor: "#fff",
    display: "grid",
    gap: "10px",
  },
  rowHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap" },
  gridTwo: { display: "grid", gap: "10px", gridTemplateColumns: "repeat(2, minmax(0, 1fr))" },
  gridThree: { display: "grid", gap: "10px", gridTemplateColumns: "repeat(3, minmax(0, 1fr))" },
  full: { gridColumn: "1 / -1" },
  label: { fontSize: "12px", fontWeight: 600, color: MODERN_TOKENS.colorTextMuted },
  textArea: {
    width: "100%",
    minHeight: "84px",
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px 10px",
    boxSizing: "border-box",
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "12px",
    lineHeight: "1.4",
    resize: "vertical",
  },
  resultTable: { width: "100%", borderCollapse: "collapse", fontSize: "12px" },
  th: {
    textAlign: "left",
    backgroundColor: "#F3F4F6",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px",
  },
  td: {
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px",
    verticalAlign: "top",
  },
});

const buildRow = (): ParameterRowState => ({
  id: `param-${Date.now().toString()}-${Math.random().toString(16).slice(2)}`,
  label: "",
  desiredName: "",
  valueMode: "manual",
  value: "",
  listValues: "",
});

const ModelBuilderView: React.FC<ModelBuilderViewProps> = ({ onOpenLegacy, embedded = false }) => {
  const shared = useModernSharedStyles();
  const styles = useStyles();
  const [rows, setRows] = useState<ParameterRowState[]>([buildRow()]);
  const [anchorSheet, setAnchorSheet] = useState<string>("");
  const [anchorAddress, setAnchorAddress] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [status, setStatus] = useState<string>("");
  const [statusType, setStatusType] = useState<"success" | "error">("success");
  const [result, setResult] = useState<ModelBuilderApplyResult | null>(null);

  const statusClass = statusType === "success" ? shared.successText : shared.errorText;
  const hasMeaningfulRows = useMemo(
    () =>
      rows.some(
        (row) => row.label.trim() || row.desiredName.trim() || row.value.trim() || row.listValues.trim()
      ),
    [rows]
  );

  const updateRow = (id: string, patch: Partial<ParameterRowState>) => {
    setRows((prev) => prev.map((row) => (row.id === id ? { ...row, ...patch } : row)));
  };

  const removeRow = (id: string) => {
    setRows((prev) => {
      const next = prev.filter((row) => row.id !== id);
      return next.length ? next : [buildRow()];
    });
  };

  const useSelection = async () => {
    setBusy(true);
    try {
      const selection = await getCurrentSelectionAddress();
      setAnchorSheet(selection.sheet);
      setAnchorAddress(selection.address);
      setStatusType("success");
      setStatus(`Anchor set to ${selection.sheet}!${selection.address}.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatusType("error");
      setStatus(`Unable to capture selection: ${message}`);
    } finally {
      setBusy(false);
    }
  };

  const apply = async () => {
    if (!anchorSheet.trim() || !anchorAddress.trim()) {
      setStatusType("error");
      setStatus("Capture an anchor cell before applying.");
      return;
    }
    if (!hasMeaningfulRows) {
      setStatusType("error");
      setStatus("Add at least one parameter row.");
      return;
    }

    setBusy(true);
    try {
      const payloadRows: ModelBuilderParameterInput[] = rows.map((row) => ({
        label: row.label,
        desiredName: row.desiredName,
        valueMode: row.valueMode,
        value: row.value,
        listValues: row.listValues,
      }));
      const applied = await applyModelBuilderParameters({
        anchorSheet: anchorSheet.trim(),
        anchorAddress: anchorAddress.trim(),
        parameters: payloadRows,
      });
      setResult(applied);
      setStatusType("success");
      setStatus(`Applied ${applied.created.length.toString()} parameter(s).`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatusType("error");
      setStatus(`Apply failed: ${message}`);
      setResult(null);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      {!embedded ? (
        <div>
          <Text className={shared.sectionTitle}>Model Builder</Text>
          <Text className={shared.sectionSubtitle}>
            Build worksheet parameters with named value cells and list-based data validation.
          </Text>
        </div>
      ) : null}

      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button appearance="primary" onClick={() => void useSelection()} disabled={busy}>
            Use Selection As Anchor
          </Button>
          <Input
            value={anchorSheet}
            placeholder="Anchor sheet"
            onChange={(_, data) => setAnchorSheet(data.value)}
          />
          <Input
            value={anchorAddress}
            placeholder="Anchor address (example: J70)"
            onChange={(_, data) => setAnchorAddress(data.value)}
          />
          <Button onClick={onOpenLegacy}>Open Legacy View</Button>
        </div>
      </div>

      <div className={styles.card}>
        {rows.map((row, index) => (
          <div key={row.id} className={styles.rowGrid}>
            <div className={styles.rowHeader}>
              <Text className={shared.cardTitle}>Parameter {index + 1}</Text>
              <Button size="small" onClick={() => removeRow(row.id)}>
                Remove
              </Button>
            </div>
            <div className={styles.gridTwo}>
              <div>
                <Text className={styles.label}>Display Label</Text>
                <Input
                  value={row.label}
                  placeholder="Example: test Parameter 1"
                  onChange={(_, data) => updateRow(row.id, { label: data.value })}
                />
              </div>
              <div>
                <Text className={styles.label}>Desired Name</Text>
                <Input
                  value={row.desiredName}
                  placeholder="Example: testParameter1"
                  onChange={(_, data) => updateRow(row.id, { desiredName: data.value })}
                />
              </div>
            </div>
            <div className={styles.gridThree}>
              <div>
                <Text className={styles.label}>Value Mode</Text>
                <Select
                  value={row.valueMode}
                  onChange={(_, data) =>
                    updateRow(row.id, { valueMode: data.value as ParameterRowState["valueMode"] })
                  }
                >
                  <option value="manual">Manual</option>
                  <option value="formula">Formula</option>
                </Select>
              </div>
              <div className={styles.full}>
                <Text className={styles.label}>{row.valueMode === "formula" ? "Formula" : "Default Value"}</Text>
                <Input
                  value={row.value}
                  placeholder={row.valueMode === "formula" ? "=TODAY()" : "Open"}
                  onChange={(_, data) => updateRow(row.id, { value: data.value })}
                />
              </div>
            </div>
            <div>
              <Text className={styles.label}>Validation List Values (comma or newline separated)</Text>
              <textarea
                className={styles.textArea}
                value={row.listValues}
                onChange={(event) => updateRow(row.id, { listValues: event.target.value })}
                placeholder={"Open\nIn Progress\nClosed"}
              />
            </div>
          </div>
        ))}

        <div className={styles.toolbar}>
          <Button onClick={() => setRows((prev) => [...prev, buildRow()])}>+ Add Parameter</Button>
          <Button appearance="primary" onClick={() => void apply()} disabled={busy}>
            Apply Parameters
          </Button>
        </div>
      </div>

      {result ? (
        <div className={shared.card}>
          <Text className={shared.cardTitle}>Apply Summary</Text>
          <Text className={shared.mutedText}>
            Anchor: {result.anchorSheet}!{result.anchorAddress}
          </Text>
          <table className={styles.resultTable}>
            <thead>
              <tr>
                <th className={styles.th}>Label</th>
                <th className={styles.th}>Named Range</th>
                <th className={styles.th}>Value Cell</th>
                <th className={styles.th}>Validation List</th>
              </tr>
            </thead>
            <tbody>
              {result.created.map((item) => (
                <tr key={`${item.finalName}-${item.valueAddress}`}>
                  <td className={styles.td}>{item.label}</td>
                  <td className={styles.td}>{item.finalName}</td>
                  <td className={styles.td}>{item.valueAddress}</td>
                  <td className={styles.td}>{item.validationListName || "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}

      {status ? <Text className={statusClass}>{status}</Text> : null}
    </div>
  );
};

export default ModelBuilderView;

