/**
 * WatchView.tsx
 *
 * Watch Window — a persistent, live-updating panel that tracks user-curated
 * cells, named ranges, and formula outputs.
 *
 * Features:
 *  - Persists the watch list across sessions via OfficeRuntime.storage
 *  - Live-refreshes values every 3s via a polling heartbeat
 *  - Shows current value, formula text, and dependent/precedent counts
 *  - Expands dependent/precedent address lists inline per row
 *  - Clicking any address navigates Excel to that cell
 *  - Inline label editing (double-click)
 *  - Degrades gracefully on Excel Web (no dependents API — shown as "—")
 *  - Accepts a pre-seeded address via context menu signal (wbm.watch.addAddress)
 */

import * as React from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { Button, Input, Text, makeStyles } from "@fluentui/react-components";
import { Eye20Regular, Delete20Regular, Add20Regular } from "@fluentui/react-icons";
import {
  WatchItem,
  WatchItemResolved,
  getWatchWindowData,
  selectCellAddress,
} from "../../taskpane";
import { MODERN_TOKENS, useModernSharedStyles } from "./designTokens";

/* global OfficeRuntime, Office */

// ─── Signal keys (match commands.ts) ──────────────────────────────────────────
const WATCH_ADD_SIGNAL_KEY = "wbm.watch.addAddress";
const WATCH_ADD_SHEET_KEY  = "wbm.watch.addSheet";

// ─── Persistence key ──────────────────────────────────────────────────────────
const STORAGE_KEY = "wbm.watchWindow.items";

// ─── Polling interval ─────────────────────────────────────────────────────────
const POLL_INTERVAL_MS = 3000;

// ─── Tiny UUID helper (no dependency) ────────────────────────────────────────
function genId(): string {
  return `w-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

// ─── Persistence helpers ──────────────────────────────────────────────────────

async function loadPersistedItems(): Promise<WatchItem[]> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const raw = await OfficeRuntime.storage.getItem(STORAGE_KEY).catch(() => null);
      if (raw) return JSON.parse(raw) as WatchItem[];
    }
    // Fallback: document settings
    const fromSettings = Office.context?.document?.settings?.get(STORAGE_KEY) as string | null;
    if (fromSettings) return JSON.parse(fromSettings) as WatchItem[];
  } catch {
    // Non-fatal — return empty.
  }
  return [];
}

async function persistItems(items: WatchItem[]): Promise<void> {
  const serialized = JSON.stringify(items);
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(STORAGE_KEY, serialized);
    }
    if (Office.context?.document?.settings) {
      Office.context.document.settings.set(STORAGE_KEY, serialized);
      await new Promise<void>((res) =>
        Office.context.document.settings.saveAsync(() => res())
      );
    }
  } catch {
    // Best-effort persistence.
  }
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: { display: "grid", gap: "16px" },
  toolbar: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  spacer: { flexGrow: 1 },
  statusBar: {
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  dot: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    backgroundColor: "#22c55e",
    display: "inline-block",
    animation: "pulse 2s infinite",
  },
  dotStale: { backgroundColor: "#f59e0b" },

  // Table
  tableWrap: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    backgroundColor: "#fff",
    overflow: "auto",
  },
  table: {
    width: "100%",
    borderCollapse: "separate",
    borderSpacing: 0,
    fontSize: "12px",
  },
  headCell: {
    position: "sticky",
    top: 0,
    zIndex: 1,
    backgroundColor: "#F3F4F6",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px 10px",
    whiteSpace: "nowrap",
    fontWeight: 600,
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  row: {
    selectors: {
      "&:nth-child(even)": { backgroundColor: "#FCFCFD" },
      "&:hover": { backgroundColor: "#F3F8F4" },
    },
  },
  rowError: { backgroundColor: "#FFF7ED !important" },
  cell: {
    padding: "8px 10px",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
    verticalAlign: "top",
  },
  addressBtn: {
    border: "none",
    background: "transparent",
    color: MODERN_TOKENS.colorBrandStrong,
    cursor: "pointer",
    padding: 0,
    fontSize: "12px",
    fontFamily: "Consolas, 'Courier New', monospace",
    textDecorationLine: "underline",
    textAlign: "left",
  },
  formulaCode: {
    fontFamily: "Consolas, 'Courier New', monospace",
    fontSize: "11px",
    color: MODERN_TOKENS.colorText,
    display: "block",
    maxWidth: "180px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  valueBadge: {
    display: "inline-block",
    padding: "2px 6px",
    borderRadius: "4px",
    fontSize: "12px",
    fontFamily: "Consolas, 'Courier New', monospace",
    backgroundColor: "#F3F4F6",
  },
  valueBadgeError: {
    backgroundColor: "#fee2e2",
    color: "#991b1b",
  },
  depToggleBtn: {
    border: "none",
    background: "transparent",
    cursor: "pointer",
    color: MODERN_TOKENS.colorBrandStrong,
    fontSize: "12px",
    padding: "0 4px",
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  depUnavailable: {
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
  },
  depExpandRow: {
    backgroundColor: "#F8FAFF",
    borderBottom: `1px solid ${MODERN_TOKENS.colorBorder}`,
  },
  depExpandCell: {
    padding: "6px 10px 10px 32px",
  },
  depAddressWrap: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    marginTop: "4px",
  },
  depChip: {
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "4px",
    padding: "2px 8px",
    fontSize: "11px",
    fontFamily: "Consolas, 'Courier New', monospace",
    backgroundColor: "#fff",
    cursor: "pointer",
    selectors: {
      "&:hover": { backgroundColor: MODERN_TOKENS.colorBrandStrong, color: "#fff" },
    },
  },

  // Add manually form
  addForm: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr auto",
    gap: "8px",
    alignItems: "end",
  },

  // Inline label edit
  labelEditInput: { fontSize: "12px" },

  emptyState: {
    padding: "40px 24px",
    textAlign: "center",
    color: MODERN_TOKENS.colorTextMuted,
    display: "grid",
    gap: "8px",
    justifyItems: "center",
  },
  emptyIcon: { opacity: 0.35 },
});

// ─── Component ────────────────────────────────────────────────────────────────

const WatchView: React.FC = () => {
  const shared = useModernSharedStyles();
  const styles = useStyles();

  // Persisted watch list
  const [items, setItems] = useState<WatchItem[]>([]);
  // Live-resolved data per item
  const [resolved, setResolved] = useState<Map<string, WatchItemResolved>>(new Map());
  // Which rows have their dependents expanded
  const [expandedDeps, setExpandedDeps] = useState<Set<string>>(new Set());
  const [expandedPrecs, setExpandedPrecs] = useState<Set<string>>(new Set());
  // Inline label edit
  const [labelEdit, setLabelEdit] = useState<{ id: string; value: string } | null>(null);
  // Manual add form
  const [addAddress, setAddAddress] = useState("");
  const [addSheet, setAddSheet] = useState("");
  const [addLabel, setAddLabel] = useState("");
  // Status
  const [lastRefreshed, setLastRefreshed] = useState<Date | null>(null);
  const [refreshing, setRefreshing] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [dependentsAvailable, setDependentsAvailable] = useState<boolean | null>(null);

  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  // ── Load persisted items on mount ────────────────────────────────────────

  useEffect(() => {
    void (async () => {
      const loaded = await loadPersistedItems();
      setItems(loaded);
    })();
  }, []);

  // ── Context menu signal — "Add to Watch Window" ──────────────────────────
  //
  // Polls for a signal written by addToWatchWindowCommand in commands.ts.
  // When found, clears it and adds the address to the watch list.

  useEffect(() => {
    let cancelled = false;

    const consumeSignal = async () => {
      try {
        let address: string | null = null;
        let sheet: string | null = null;

        if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
          address = await OfficeRuntime.storage.getItem(WATCH_ADD_SIGNAL_KEY).catch(() => null);
          sheet   = await OfficeRuntime.storage.getItem(WATCH_ADD_SHEET_KEY).catch(() => null);
          if (address) {
            await OfficeRuntime.storage.removeItem(WATCH_ADD_SIGNAL_KEY).catch(() => undefined);
            await OfficeRuntime.storage.removeItem(WATCH_ADD_SHEET_KEY).catch(() => undefined);
          }
        } else if (Office.context?.document?.settings) {
          address = Office.context.document.settings.get(WATCH_ADD_SIGNAL_KEY) as string | null;
          sheet   = Office.context.document.settings.get(WATCH_ADD_SHEET_KEY) as string | null;
          if (address) {
            Office.context.document.settings.remove(WATCH_ADD_SIGNAL_KEY);
            Office.context.document.settings.remove(WATCH_ADD_SHEET_KEY);
            await new Promise<void>((res) =>
              Office.context.document.settings.saveAsync(() => res())
            );
          }
        }

        if (!address || !sheet || cancelled) return;

        // Strip sheet qualifier if present (e.g. "Sheet1!B5" → sheet=Sheet1, address=B5)
        const bare = address.includes("!") ? address.split("!").slice(1).join("!") : address;
        const resolvedSheet = address.includes("!") ? address.split("!")[0].replace(/^'|'$/g, "") : sheet;

        addItem(bare, resolvedSheet, `${resolvedSheet}!${bare}`);
      } catch {
        // Non-fatal.
      }
    };

    void consumeSignal();
    const interval = setInterval(() => void consumeSignal(), 700);
    return () => { cancelled = true; clearInterval(interval); };
  }, [items]); // Re-run so addItem closure sees current items

  // ── Refresh logic ─────────────────────────────────────────────────────────

  const refresh = useCallback(async (watchItems: WatchItem[]) => {
    if (watchItems.length === 0) {
      setLastRefreshed(new Date());
      return;
    }
    setRefreshing(true);
    try {
      const data = await getWatchWindowData(watchItems);
      const newMap = new Map<string, WatchItemResolved>();
      data.forEach((r) => newMap.set(r.id, r));
      setResolved(newMap);
      setLastRefreshed(new Date());

      // Determine whether the dependents API is available (check once)
      if (dependentsAvailable === null && data.length > 0) {
        setDependentsAvailable(data[0].dependentsAvailable);
      }
    } catch {
      // Non-fatal — keep stale data.
    } finally {
      setRefreshing(false);
    }
  }, [dependentsAvailable]);

  // Refresh whenever the items list changes.
  useEffect(() => {
    void refresh(items);
  }, [items]);

  // Polling heartbeat.
  useEffect(() => {
    if (pollRef.current) clearInterval(pollRef.current);
    pollRef.current = setInterval(() => void refresh(items), POLL_INTERVAL_MS);
    return () => {
      if (pollRef.current) clearInterval(pollRef.current);
    };
  }, [items, refresh]);

  // ── CRUD helpers ──────────────────────────────────────────────────────────

  const addItem = (address: string, sheet: string, label: string) => {
    const trimmedAddress = address.trim();
    const trimmedSheet   = sheet.trim();
    if (!trimmedAddress) return;

    // De-duplicate by address+sheet
    const key = `${trimmedSheet}::${trimmedAddress}`.toUpperCase();
    if (items.some((i) => `${i.sheet}::${i.address}`.toUpperCase() === key)) return;

    const next: WatchItem[] = [
      ...items,
      {
        id: genId(),
        label: label.trim() || trimmedAddress,
        sheet: trimmedSheet,
        address: trimmedAddress,
        addedAt: new Date().toISOString(),
      },
    ];
    setItems(next);
    void persistItems(next);
  };

  const removeItems = (ids: Set<string>) => {
    const next = items.filter((i) => !ids.has(i.id));
    setItems(next);
    setSelectedIds(new Set());
    void persistItems(next);
  };

  const saveLabelEdit = () => {
    if (!labelEdit) return;
    const next = items.map((i) =>
      i.id === labelEdit.id ? { ...i, label: labelEdit.value.trim() || i.address } : i
    );
    setItems(next);
    void persistItems(next);
    setLabelEdit(null);
  };

  // ── Manual add form submit ────────────────────────────────────────────────

  const handleManualAdd = () => {
    if (!addAddress.trim()) return;
    addItem(addAddress.trim(), addSheet.trim() || "Sheet1", addLabel.trim() || addAddress.trim());
    setAddAddress("");
    setAddSheet("");
    setAddLabel("");
  };

  const handleAddCurrentCell = async () => {
    try {
      // Pull current selection directly
      const data = await getWatchWindowData([]);  // returns selection info
      // We use a dedicated service function for this
      const { address: selAddress, sheet: selSheet } = await (
        await import("../../taskpane") as any
      ).getCurrentSelectionAddress();
      const bare = selAddress.includes("!") ? selAddress.split("!").slice(1).join("!") : selAddress;
      addItem(bare, selSheet, `${selSheet}!${bare}`);
    } catch {
      // Non-fatal.
    }
  };

  // ── Navigate to cell ──────────────────────────────────────────────────────

  const navigateTo = async (address: string, sheet: string) => {
    try {
      await selectCellAddress(address, sheet);
    } catch {
      // Non-fatal.
    }
  };

  // ── Toggle expand ─────────────────────────────────────────────────────────

  const toggleDeps = (id: string) => {
    setExpandedDeps((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const togglePrecs = (id: string) => {
    setExpandedPrecs((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  // ── Render helpers ────────────────────────────────────────────────────────

  const formatLastRefreshed = () => {
    if (!lastRefreshed) return "Never";
    return lastRefreshed.toLocaleTimeString();
  };

  const isErrorValue = (v: string) =>
    /^#(REF|VALUE|NAME|DIV\/0|N\/A|NUM|NULL|CALC|SPILL|CONNECT|BLOCKED|UNKNOWN)!?$/i.test(v);

  return (
    <div className={styles.root}>
      {/* ── Header ── */}
      <div>
        <Text className={shared.sectionTitle}>Watch Window</Text>
        <Text className={shared.sectionSubtitle}>
          Live-tracking for cells, named ranges, and formulas. Refreshes every 3s.
        </Text>
      </div>

      {/* ── Status bar ── */}
      <div className={styles.statusBar}>
        <span className={`${styles.dot} ${refreshing ? styles.dotStale : ""}`} />
        <span>Last refresh: {formatLastRefreshed()}</span>
        {dependentsAvailable === false ? (
          <span style={{ marginLeft: "auto", color: "#92400e", fontSize: "11px" }}>
            ⚠ Dependents/Precedents require Excel Desktop (ExcelApi 1.14+)
          </span>
        ) : null}
      </div>

      {/* ── Toolbar ── */}
      <div className={shared.card}>
        <div className={styles.toolbar}>
          <Button
            icon={<Add20Regular />}
            onClick={() => void handleAddCurrentCell()}
          >
            Add Current Cell
          </Button>
          <Button
            icon={<Delete20Regular />}
            disabled={selectedIds.size === 0}
            onClick={() => removeItems(selectedIds)}
          >
            Remove Selected ({selectedIds.size})
          </Button>
          <div className={styles.spacer} />
          <Button
            disabled={refreshing || items.length === 0}
            onClick={() => void refresh(items)}
          >
            {refreshing ? "Refreshing…" : "Refresh ↻"}
          </Button>
        </div>
      </div>

      {/* ── Watch Table ── */}
      {items.length === 0 ? (
        <div className={styles.emptyState}>
          <Eye20Regular className={styles.emptyIcon} style={{ width: 40, height: 40 }} />
          <Text className={shared.cardTitle}>No cells being watched</Text>
          <Text className={shared.mutedText}>
            Right-click any cell and choose "Add to Watch Window", or use the form below.
          </Text>
        </div>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th className={styles.headCell} style={{ width: 32 }}>
                  <input
                    type="checkbox"
                    checked={items.length > 0 && items.every((i) => selectedIds.has(i.id))}
                    onChange={() => {
                      const allSelected = items.every((i) => selectedIds.has(i.id));
                      setSelectedIds(allSelected ? new Set() : new Set(items.map((i) => i.id)));
                    }}
                  />
                </th>
                <th className={styles.headCell}>Label</th>
                <th className={styles.headCell}>Address</th>
                <th className={styles.headCell}>Value</th>
                <th className={styles.headCell}>Formula</th>
                <th className={styles.headCell} title="Direct dependents — cells that reference this one">Deps ↓</th>
                <th className={styles.headCell} title="Direct precedents — cells this one references">Precs ↑</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item) => {
                const r = resolved.get(item.id);
                const hasError = r ? isErrorValue(r.currentValue) : false;
                const rowClass = `${styles.row} ${hasError ? styles.rowError : ""}`;
                const depsExpanded = expandedDeps.has(item.id);
                const precsExpanded = expandedPrecs.has(item.id);

                return (
                  <React.Fragment key={item.id}>
                    <tr className={rowClass}>
                      {/* Checkbox */}
                      <td className={styles.cell}>
                        <input
                          type="checkbox"
                          checked={selectedIds.has(item.id)}
                          onChange={() => setSelectedIds((prev) => {
                            const next = new Set(prev);
                            if (next.has(item.id)) next.delete(item.id); else next.add(item.id);
                            return next;
                          })}
                        />
                      </td>

                      {/* Label (inline edit on double-click) */}
                      <td className={styles.cell}>
                        {labelEdit?.id === item.id ? (
                          <Input
                            className={styles.labelEditInput}
                            value={labelEdit.value}
                            autoFocus
                            onChange={(_, d) => setLabelEdit({ id: item.id, value: d.value })}
                            onBlur={saveLabelEdit}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") saveLabelEdit();
                              if (e.key === "Escape") setLabelEdit(null);
                            }}
                          />
                        ) : (
                          <span
                            title="Double-click to rename"
                            style={{ cursor: "pointer", userSelect: "none" }}
                            onDoubleClick={() => setLabelEdit({ id: item.id, value: item.label })}
                          >
                            {item.label}
                          </span>
                        )}
                      </td>

                      {/* Address — navigates Excel on click */}
                      <td className={styles.cell}>
                        <button
                          type="button"
                          className={styles.addressBtn}
                          onClick={() => void navigateTo(item.address, item.sheet)}
                          title={`Navigate to ${item.sheet}!${item.address}`}
                        >
                          {item.sheet}!{item.address}
                        </button>
                      </td>

                      {/* Value */}
                      <td className={styles.cell}>
                        {r ? (
                          <span className={`${styles.valueBadge} ${hasError ? styles.valueBadgeError : ""}`}>
                            {hasError ? `⚠ ${r.currentValue}` : r.currentValue}
                          </span>
                        ) : (
                          <span className={shared.mutedText}>…</span>
                        )}
                      </td>

                      {/* Formula */}
                      <td className={styles.cell}>
                        {r ? (
                          <span className={styles.formulaCode} title={r.formula}>
                            {r.hasFormula ? r.formula : "(value)"}
                          </span>
                        ) : (
                          <span className={shared.mutedText}>…</span>
                        )}
                      </td>

                      {/* Dependents */}
                      <td className={styles.cell}>
                        {r && r.dependentsAvailable ? (
                          r.dependentAddresses.length > 0 ? (
                            <button
                              type="button"
                              className={styles.depToggleBtn}
                              onClick={() => toggleDeps(item.id)}
                              title="Toggle dependent addresses"
                            >
                              {r.dependentCount} {depsExpanded ? "▲" : "▼"}
                            </button>
                          ) : (
                            <span className={styles.depUnavailable}>0</span>
                          )
                        ) : (
                          <span className={styles.depUnavailable} title="Requires Excel Desktop (ExcelApi 1.14+)">—</span>
                        )}
                      </td>

                      {/* Precedents */}
                      <td className={styles.cell}>
                        {r && r.dependentsAvailable ? (
                          r.precedentAddresses.length > 0 ? (
                            <button
                              type="button"
                              className={styles.depToggleBtn}
                              onClick={() => togglePrecs(item.id)}
                              title="Toggle precedent addresses"
                            >
                              {r.precedentCount} {precsExpanded ? "▲" : "▼"}
                            </button>
                          ) : (
                            <span className={styles.depUnavailable}>0</span>
                          )
                        ) : (
                          <span className={styles.depUnavailable} title="Requires Excel Desktop (ExcelApi 1.14+)">—</span>
                        )}
                      </td>
                    </tr>

                    {/* Dependents expand row */}
                    {depsExpanded && r && r.dependentAddresses.length > 0 ? (
                      <tr className={styles.depExpandRow}>
                        <td colSpan={7} className={styles.depExpandCell}>
                          <Text className={shared.mutedText} style={{ fontSize: "11px" }}>
                            Dependents — cells that reference {item.sheet}!{item.address}:
                          </Text>
                          <div className={styles.depAddressWrap}>
                            {r.dependentAddresses.map((addr) => {
                              const parts = addr.includes("!")
                                ? { sheet: addr.split("!")[0].replace(/^'|'$/g, ""), address: addr.split("!").slice(1).join("!") }
                                : { sheet: item.sheet, address: addr };
                              return (
                                <button
                                  key={addr}
                                  type="button"
                                  className={styles.depChip}
                                  onClick={() => void navigateTo(parts.address, parts.sheet)}
                                  title={`Navigate to ${addr}`}
                                >
                                  {addr}
                                </button>
                              );
                            })}
                          </div>
                        </td>
                      </tr>
                    ) : null}

                    {/* Precedents expand row */}
                    {precsExpanded && r && r.precedentAddresses.length > 0 ? (
                      <tr className={styles.depExpandRow}>
                        <td colSpan={7} className={styles.depExpandCell}>
                          <Text className={shared.mutedText} style={{ fontSize: "11px" }}>
                            Precedents — cells that {item.sheet}!{item.address} references:
                          </Text>
                          <div className={styles.depAddressWrap}>
                            {r.precedentAddresses.map((addr) => {
                              const parts = addr.includes("!")
                                ? { sheet: addr.split("!")[0].replace(/^'|'$/g, ""), address: addr.split("!").slice(1).join("!") }
                                : { sheet: item.sheet, address: addr };
                              return (
                                <button
                                  key={addr}
                                  type="button"
                                  className={styles.depChip}
                                  onClick={() => void navigateTo(parts.address, parts.sheet)}
                                  title={`Navigate to ${addr}`}
                                >
                                  {addr}
                                </button>
                              );
                            })}
                          </div>
                        </td>
                      </tr>
                    ) : null}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* ── Manual Add Form ── */}
      <div className={shared.card}>
        <Text className={shared.cardTitle} style={{ marginBottom: "8px" }}>
          Add Manually
        </Text>
        <div className={styles.addForm}>
          <div>
            <Text className={shared.mutedText}>Address</Text>
            <Input
              placeholder="B5 or A1:A10"
              value={addAddress}
              onChange={(_, d) => setAddAddress(d.value)}
              onKeyDown={(e) => { if (e.key === "Enter") handleManualAdd(); }}
            />
          </div>
          <div>
            <Text className={shared.mutedText}>Sheet</Text>
            <Input
              placeholder="Sheet1"
              value={addSheet}
              onChange={(_, d) => setAddSheet(d.value)}
              onKeyDown={(e) => { if (e.key === "Enter") handleManualAdd(); }}
            />
          </div>
          <Button
            appearance="primary"
            disabled={!addAddress.trim()}
            onClick={handleManualAdd}
            style={{ alignSelf: "end" }}
          >
            Add
          </Button>
        </div>
        <Text className={shared.mutedText} style={{ fontSize: "11px", marginTop: "4px" }}>
          Tip: You can enter a named range (e.g. <code>Revenue</code>) in the Address field — it will resolve live.
        </Text>
      </div>

      {items.length > 0 ? (
        <Text className={shared.mutedText}>{items.length} item(s) being watched</Text>
      ) : null}
    </div>
  );
};

export default WatchView;
