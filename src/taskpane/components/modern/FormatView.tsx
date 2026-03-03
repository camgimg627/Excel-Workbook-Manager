import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, Text, makeStyles } from "@fluentui/react-components";
import { Add20Regular, ArrowClockwise20Regular } from "@fluentui/react-icons";
import {
  ApplyNavigationTemplateRequest,
  CellStylePreset,
  CreateShapeBuilderShapeRequest,
  InsertableShapeType,
  NavigationDestinationOption,
  ShapeBuilderEffect,
  ShapeBuilderShapeRecord,
  SheetFormatRecord,
  activateNavigationDestination,
  applyCellStylePreset,
  applyNavigationTemplate,
  applySheetFormat,
  applySheetFormatToAllSheets,
  captureSheetFormat,
  createShapeBuilderShape,
  deleteSheetFormat,
  deleteShapeBuilderShape,
  duplicateShapeBuilderShape,
  freezeFirstColumn,
  freezeTopRow,
  listNavigationDestinations,
  listShapeBuilderShapes,
  listSheetFormats,
  openFormatEditorPopout,
  recaptureSheetFormatFromSelection,
  refreshPivotTables,
  toggleGridlines,
  unfreezePanes,
  updateShapeBuilderShape,
} from "../../taskpane";
import { MODERN_TOKENS } from "./designTokens";

interface FormatViewProps {
  onOpenLegacy: () => void;
  isPopout?: boolean;
}

type EditorTab = "shape" | "style" | "size" | "link";

const SHAPE_TYPES: Array<{ value: InsertableShapeType; label: string }> = [
  { value: "RoundedRectangle", label: "Rounded Rectangle" },
  { value: "Rectangle", label: "Rectangle" },
  { value: "Chevron", label: "Chevron" },
  { value: "Hexagon", label: "Hexagon" },
  { value: "Diamond", label: "Diamond" },
  { value: "Oval", label: "Oval" },
];

const ICONS = [
  { key: "none", label: "None", symbol: "-" },
  { key: "home", label: "Home", symbol: "⌂" },
  { key: "data", label: "Data", symbol: "▦" },
  { key: "controls", label: "Controls", symbol: "⌘" },
  { key: "reports", label: "Reports", symbol: "▤" },
  { key: "settings", label: "Settings", symbol: "⚙" },
  { key: "help", label: "Help", symbol: "?" },
];

const COLORS = ["#4679C7", "#54A75A", "#E8791B", "#CF2E2E", "#6742B5", "#5AA79D", "#EAB51C", "#8B97A5", "#1E2B3D"];

const STYLE_PRESETS: Array<{ preset: CellStylePreset; label: string }> = [
  { preset: "Input Cell", label: "Input" },
  { preset: "Parameter Cell", label: "Parameter" },
  { preset: "Header", label: "Header" },
  { preset: "Subheader", label: "Subheader" },
];

const normalizeError = (error: unknown): string => (error instanceof Error ? error.message : String(error));
const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max);

const parseCellDestination = (value: string): string | null => {
  const trimmed = value.trim();
  const match = trimmed.match(
    /^'?([^'!]+)'?!([A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?|[A-Za-z]{1,3}:[A-Za-z]{1,3}|\d+:\d+)$/i
  );
  if (!match) {
    return null;
  }
  return `cell::${match[1].trim()}::${match[2].toUpperCase()}`;
};

const styles = makeStyles({
  root: { display: "grid", gap: "16px" },
  hero: {
    borderRadius: "12px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#4579C7",
    color: "#FFFFFF",
    padding: "14px",
    display: "grid",
    gap: "10px",
    boxShadow: MODERN_TOKENS.shadowCard,
  },
  heroTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap" },
  heroTitle: { fontSize: "28px", fontWeight: 700, lineHeight: "30px" },
  heroSub: { fontSize: "14px", opacity: 0.92 },
  heroBtns: { display: "flex", gap: "8px", flexWrap: "wrap" },
  heroBtn: {
    borderRadius: "8px",
    border: "none",
    backgroundColor: "rgba(255,255,255,0.2)",
    color: "#FFFFFF",
    padding: "8px 12px",
    fontWeight: 600,
    cursor: "pointer",
  },
  addBtn: {
    borderRadius: "8px",
    border: "none",
    backgroundColor: "#FFFFFF",
    color: "#305E9D",
    padding: "10px 14px",
    fontWeight: 700,
    cursor: "pointer",
    display: "inline-flex",
    gap: "8px",
    alignItems: "center",
  },
  card: {
    borderRadius: "12px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "12px",
    display: "grid",
    gap: "10px",
  },
  row: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px", flexWrap: "wrap" },
  title: { fontSize: "13px", fontWeight: 700, textTransform: "uppercase", color: MODERN_TOKENS.colorTextMuted },
  muted: { fontSize: "12px", color: MODERN_TOKENS.colorTextMuted },
  status: { fontSize: "12px" },
  ok: { color: MODERN_TOKENS.colorBrandStrong },
  error: { color: MODERN_TOKENS.colorDanger },
  shapeList: { display: "grid", gap: "8px", maxHeight: "300px", overflowY: "auto" },
  shapeRow: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px",
    display: "grid",
    gridTemplateColumns: "88px 1fr auto",
    gap: "10px",
    alignItems: "center",
    cursor: "pointer",
  },
  selected: { backgroundColor: "#EFF4FE", border: "1px solid #A7C0E7" },
  preview: {
    width: "82px",
    height: "34px",
    borderRadius: "8px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    fontSize: "12px",
    fontWeight: 700,
    boxShadow: "0 6px 14px rgba(17,24,39,0.16)",
    overflow: "hidden",
    whiteSpace: "nowrap",
    textOverflow: "ellipsis",
  },
  itemName: { fontSize: "18px", fontWeight: 700 },
  itemSub: { fontSize: "11px", color: MODERN_TOKENS.colorTextMuted },
  itemTag: {
    marginTop: "4px",
    fontSize: "11px",
    color: "#344054",
    backgroundColor: "#F2F4F7",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "999px",
    padding: "3px 8px",
    width: "fit-content",
  },
  actionCol: { display: "grid", gap: "6px" },
  smallBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    padding: "6px 10px",
    fontSize: "12px",
    cursor: "pointer",
  },
  dangerBtn: { border: "1px solid #F2C2C2", color: MODERN_TOKENS.colorDanger },
  tabs: { display: "grid", gridTemplateColumns: "repeat(4,minmax(0,1fr))", gap: "6px", backgroundColor: "#EEF0F3", borderRadius: "10px", padding: "6px" },
  tab: { border: "none", borderRadius: "8px", backgroundColor: "transparent", color: "#6B7280", padding: "10px 6px", fontSize: "15px", fontWeight: 600, cursor: "pointer" },
  tabActive: { backgroundColor: "#FFFFFF", color: MODERN_TOKENS.colorText, boxShadow: "0 1px 2px rgba(17,24,39,0.08)" },
  panel: { display: "grid", gap: "10px" },
  label: { fontSize: "11px", fontWeight: 700, textTransform: "uppercase", color: MODERN_TOKENS.colorTextMuted },
  input: { width: "100%", boxSizing: "border-box", border: `1px solid ${MODERN_TOKENS.colorBorder}`, borderRadius: "8px", padding: "9px 10px", fontSize: "14px" },
  select: { width: "100%", boxSizing: "border-box", border: `1px solid ${MODERN_TOKENS.colorBorder}`, borderRadius: "8px", padding: "9px 10px", fontSize: "14px", backgroundColor: "#FFFFFF" },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: "8px" },
  swatches: { display: "grid", gridTemplateColumns: "repeat(5,minmax(0,1fr))", gap: "8px" },
  swatch: { height: "34px", borderRadius: "8px", border: `1px solid ${MODERN_TOKENS.colorBorder}`, cursor: "pointer" },
  swatchActive: { boxShadow: "inset 0 0 0 2px #FFFFFF, 0 0 0 2px #4679C7" },
  iconGrid: { display: "grid", gridTemplateColumns: "repeat(4,minmax(0,1fr))", gap: "8px", maxHeight: "170px", overflowY: "auto", border: `1px solid ${MODERN_TOKENS.colorBorder}`, borderRadius: "8px", padding: "8px" },
  iconBtn: { borderRadius: "8px", border: `1px solid ${MODERN_TOKENS.colorBorder}`, backgroundColor: "#FFFFFF", padding: "8px 6px", display: "grid", justifyItems: "center", gap: "2px", cursor: "pointer", fontSize: "11px" },
  iconActive: { border: "1px solid #7EA5E1", backgroundColor: "#EFF4FE" },
  radioRow: { display: "grid", gap: "8px" },
  horizontal: { display: "flex", gap: "8px", flexWrap: "wrap" },
  formatRow: { borderRadius: "8px", border: `1px solid ${MODERN_TOKENS.colorBorder}`, padding: "8px 10px", display: "grid", gridTemplateColumns: "1fr auto", gap: "8px", alignItems: "center" },
  empty: { borderRadius: "8px", border: `1px dashed ${MODERN_TOKENS.colorBorder}`, padding: "12px", color: MODERN_TOKENS.colorTextMuted, fontSize: "12px" },
});

const FormatView: React.FC<FormatViewProps> = ({ onOpenLegacy, isPopout = false }) => {
  const s = styles();
  const [shapes, setShapes] = useState<ShapeBuilderShapeRecord[]>([]);
  const [destinations, setDestinations] = useState<NavigationDestinationOption[]>([]);
  const [formats, setFormats] = useState<SheetFormatRecord[]>([]);
  const [selectedShapeId, setSelectedShapeId] = useState("");
  const [selectedFormatId, setSelectedFormatId] = useState("");
  const [newFormatName, setNewFormatName] = useState("");
  const [tab, setTab] = useState<EditorTab>("shape");
  const [draft, setDraft] = useState<ShapeBuilderShapeRecord | null>(null);
  const [iconQuery, setIconQuery] = useState("");
  const [cellRefInput, setCellRefInput] = useState("");
  const [applyTemplateAll, setApplyTemplateAll] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [statusType, setStatusType] = useState<"ok" | "error">("ok");
  const shapeHandlersRef = useRef<Array<{ remove: () => Promise<void> | void }>>([]);

  const selectedShape = useMemo(() => shapes.find((x) => x.id === selectedShapeId) ?? null, [shapes, selectedShapeId]);
  const filteredIcons = useMemo(() => {
    const q = iconQuery.trim().toLowerCase();
    return q ? ICONS.filter((x) => x.key.includes(q) || x.label.toLowerCase().includes(q)) : ICONS;
  }, [iconQuery]);

  const setOk = (message: string) => {
    setStatusType("ok");
    setStatus(message);
  };

  const setErr = (error: unknown) => {
    setStatusType("error");
    setStatus(normalizeError(error));
  };

  const refreshAll = useCallback(async () => {
    setBusy(true);
    try {
      const [shapeRecords, destinationOptions, formatRecords] = await Promise.all([
        listShapeBuilderShapes(),
        listNavigationDestinations(),
        listSheetFormats(),
      ]);
      setShapes(shapeRecords);
      setDestinations(destinationOptions);
      setFormats(formatRecords);
      setSelectedShapeId((prev) => (shapeRecords.some((x) => x.id === prev) ? prev : shapeRecords[0]?.id ?? ""));
      setSelectedFormatId((prev) => (formatRecords.some((x) => x.id === prev) ? prev : formatRecords[0]?.id ?? ""));
    } catch (error) {
      setErr(error);
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    void refreshAll();
  }, [refreshAll]);

  useEffect(() => {
    setDraft(selectedShape ? { ...selectedShape } : null);
    setIconQuery("");
    setCellRefInput("");
  }, [selectedShape]);

  const clearShapeHandlers = useCallback(() => {
    const handlers = [...shapeHandlersRef.current];
    shapeHandlersRef.current = [];
    handlers.forEach((handler) => {
      try {
        const removed = handler.remove();
        if (removed instanceof Promise) {
          void removed;
        }
      } catch {
        // ignore cleanup errors
      }
    });
  }, []);

  useEffect(() => {
    let disposed = false;

    const linkedShapes = shapes.filter((shape) => {
      if (shape.linkType === "Internal") {
        return Boolean(shape.internalDestinationId);
      }
      if (shape.linkType === "External") {
        return Boolean(shape.externalUrl);
      }
      return false;
    });

    const registerHandlers = async () => {
      if (typeof Excel === "undefined") {
        return;
      }
      clearShapeHandlers();
      if (!linkedShapes.length) {
        return;
      }

      try {
        await Excel.run(async (context) => {
          for (const shapeRecord of linkedShapes) {
            const sheet = context.workbook.worksheets.getItem(shapeRecord.sheetName);
            const shape = sheet.shapes.getItemOrNullObject(shapeRecord.shapeName);
            shape.load("name");
            await context.sync();
            if (shape.isNullObject) {
              continue;
            }

            const destinationId =
              shapeRecord.linkType === "External"
                ? `url::${shapeRecord.externalUrl}`
                : shapeRecord.internalDestinationId;
            const handler = await shape.onActivated.add(async () => {
              if (disposed) {
                return;
              }
              try {
                await activateNavigationDestination(destinationId);
              } catch {
                // best effort
              }
            });
            shapeHandlersRef.current.push(handler as { remove: () => Promise<void> | void });
          }
        });
      } catch {
        // shape activation events are host-dependent
      }
    };

    void registerHandlers();
    return () => {
      disposed = true;
      clearShapeHandlers();
    };
  }, [clearShapeHandlers, shapes]);

  const patchShape = useCallback(
    async (patch: Partial<CreateShapeBuilderShapeRequest>) => {
      if (!selectedShape) {
        return;
      }
      try {
        const updated = await updateShapeBuilderShape(selectedShape.sheetName, selectedShape.shapeName, patch);
        setShapes((prev) => prev.map((x) => (x.id === selectedShape.id ? updated : x)));
        setDraft(updated);
      } catch (error) {
        setErr(error);
      }
    },
    [selectedShape]
  );

  const addShape = async () => {
    setBusy(true);
    try {
      const created = await createShapeBuilderShape({
        shapeType: "RoundedRectangle",
        text: `Button ${shapes.length + 1}`,
        iconKey: "none",
        fillColor: "#A8D5AD",
        outlineColor: "#8CBF95",
        outlineWidth: 2,
        width: 160,
        height: 60,
        fontColor: "#1F2937",
        fontSize: 16,
        bold: true,
        italic: false,
        linkType: "None",
        internalDestinationId: "",
        externalUrl: "",
        effect: "Shadow",
      });
      setShapes((prev) => [...prev, created].sort((a, b) => a.top - b.top || a.left - b.left));
      setSelectedShapeId(created.id);
      setTab("shape");
      setOk("Shape added.");
    } catch (error) {
      setErr(error);
    } finally {
      setBusy(false);
    }
  };

  const duplicateShape = async (shape: ShapeBuilderShapeRecord) => {
    try {
      const copy = await duplicateShapeBuilderShape(shape.sheetName, shape.shapeName);
      setShapes((prev) => [...prev, copy].sort((a, b) => a.top - b.top || a.left - b.left));
      setSelectedShapeId(copy.id);
      setOk("Shape duplicated.");
    } catch (error) {
      setErr(error);
    }
  };

  const deleteShape = async (shape: ShapeBuilderShapeRecord) => {
    if (!window.confirm(`Delete "${shape.text || shape.shapeName}"?`)) {
      return;
    }
    try {
      await deleteShapeBuilderShape(shape.sheetName, shape.shapeName);
      const next = shapes.filter((x) => x.id !== shape.id);
      setShapes(next);
      setSelectedShapeId((prev) => (prev === shape.id ? next[0]?.id ?? "" : prev));
      setOk("Shape deleted.");
    } catch (error) {
      setErr(error);
    }
  };

  const applyTemplate = async (layout: "Vertical" | "Horizontal") => {
    const buttons = destinations.slice(0, 4).map((destination, index) => ({
      id: `btn-${index + 1}`,
      label: destination.label.replace(/^\w+:\s*/, "").slice(0, 28),
      destinationId: destination.id,
      fillColor: "#A8D5AD",
      outlineColor: "#8CBF95",
      fontColor: "#1F2937",
    }));
    if (!buttons.length) {
      setErr("No destinations available for a template.");
      return;
    }
    const request: ApplyNavigationTemplateRequest = {
      layout,
      panelFillColor: "#E7EAEE",
      panelOutlineColor: "#C6CED8",
      panelWidth: layout === "Vertical" ? 200 : 920,
      panelPadding: 16,
      buttonWidth: 160,
      buttonHeight: 60,
      buttonGap: 14,
      originLeft: 18,
      originTop: 32,
      buttons,
      applyToAllSheets: applyTemplateAll,
    };
    try {
      await applyNavigationTemplate(request);
      await refreshAll();
      setOk(`${layout} template applied.`);
    } catch (error) {
      setErr(error);
    }
  };

  const captureFormat = async () => {
    const name = newFormatName.trim() || `Sheet Format ${formats.length + 1}`;
    try {
      const record = await captureSheetFormat(name);
      await refreshAll();
      setSelectedFormatId(record.id);
      setNewFormatName("");
      setOk(`Captured "${record.name}".`);
    } catch (error) {
      setErr(error);
    }
  };

  return (
    <div className={s.root}>
      <div className={s.hero}>
        <div className={s.heroTop}>
          <div>
            <Text className={s.heroTitle}>Shape Builder</Text>
            <Text className={s.heroSub}>Excel Add-in</Text>
          </div>
          <div className={s.heroBtns}>
            {!isPopout ? (
              <button className={s.heroBtn} type="button" onClick={() => void openFormatEditorPopout()}>
                Pop Out
              </button>
            ) : null}
            <button className={s.heroBtn} type="button" onClick={onOpenLegacy}>
              Legacy
            </button>
          </div>
        </div>
        <div>
          <button className={s.addBtn} type="button" onClick={() => void addShape()} disabled={busy}>
            <Add20Regular /> Add Shape
          </button>
        </div>
        <div className={s.row}>
          <Text className={`${s.status} ${statusType === "error" ? s.error : s.ok}`}>{status || "Ready."}</Text>
          <Button size="small" icon={<ArrowClockwise20Regular />} onClick={() => void refreshAll()} disabled={busy}>
            Refresh
          </Button>
        </div>
      </div>

      <div className={s.card}>
        <div className={s.row}>
          <Text className={s.title}>Shape List</Text>
          <Text className={s.muted}>{shapes.length} on active sheet</Text>
        </div>
        <div className={s.shapeList}>
          {!shapes.length ? <div className={s.empty}>No shapes yet. Click Add Shape.</div> : null}
          {shapes.map((shape) => {
            const selected = shape.id === selectedShapeId;
            return (
              <div
                key={shape.id}
                className={`${s.shapeRow} ${selected ? s.selected : ""}`}
                onClick={() => setSelectedShapeId(shape.id)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    setSelectedShapeId(shape.id);
                  }
                }}
                role="button"
                tabIndex={0}
              >
                <div className={s.preview} style={{ backgroundColor: shape.fillColor, border: `${Math.max(1, shape.outlineWidth)}px solid ${shape.outlineColor}`, color: shape.fontColor }}>
                  {shape.text || "Shape"}
                </div>
                <div>
                  <Text className={s.itemName}>{shape.text || shape.shapeName}</Text>
                  <Text className={s.itemSub}>{shape.shapeType.replace(/([A-Z])/g, " $1").trim()}</Text>
                  <div className={s.itemTag}>{shape.linkType.toLowerCase()}</div>
                </div>
                <div className={s.actionCol}>
                  <button className={s.smallBtn} type="button" onClick={(event) => { event.stopPropagation(); void duplicateShape(shape); }} disabled={busy}>
                    Copy
                  </button>
                  <button className={`${s.smallBtn} ${s.dangerBtn}`} type="button" onClick={(event) => { event.stopPropagation(); void deleteShape(shape); }} disabled={busy}>
                    Delete
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div className={s.card}>
        <div className={s.row}>
          <Text className={s.title}>Editor</Text>
          <Text className={s.muted}>{selectedShape ? selectedShape.shapeName : "Select a shape"}</Text>
        </div>
        {!draft ? (
          <div className={s.empty}>Select a shape to edit.</div>
        ) : (
          <>
            <div className={s.tabs}>
              {(["shape", "style", "size", "link"] as EditorTab[]).map((item) => (
                <button key={item} className={`${s.tab} ${tab === item ? s.tabActive : ""}`} type="button" onClick={() => setTab(item)}>
                  {item === "shape" ? "Shape" : item === "style" ? "Style" : item === "size" ? "Size" : "Link"}
                </button>
              ))}
            </div>

            {tab === "shape" ? (
              <div className={s.panel}>
                <Text className={s.label}>Shape Type</Text>
                <select className={s.select} value={draft.shapeType} onChange={(event) => { const shapeType = event.target.value as InsertableShapeType; setDraft({ ...draft, shapeType }); void patchShape({ shapeType }); }}>
                  {SHAPE_TYPES.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                </select>
                <Text className={s.label}>Text</Text>
                <input className={s.input} value={draft.text} onChange={(event) => setDraft({ ...draft, text: event.target.value })} onBlur={() => void patchShape({ text: draft.text })} />
                <Text className={s.label}>Icon</Text>
                <Input value={iconQuery} placeholder="Search icons..." onChange={(_, data) => setIconQuery(data.value)} />
                <div className={s.iconGrid}>
                  {filteredIcons.map((icon) => (
                    <button key={icon.key} className={`${s.iconBtn} ${draft.iconKey === icon.key ? s.iconActive : ""}`} type="button" onClick={() => { setDraft({ ...draft, iconKey: icon.key }); void patchShape({ iconKey: icon.key }); }}>
                      <span>{icon.symbol}</span><span>{icon.label}</span>
                    </button>
                  ))}
                </div>
              </div>
            ) : null}

            {tab === "style" ? (
              <div className={s.panel}>
                <Text className={s.label}>Fill Color</Text>
                <div className={s.swatches}>
                  {COLORS.map((color) => <button key={`f-${color}`} className={`${s.swatch} ${draft.fillColor.toLowerCase() === color.toLowerCase() ? s.swatchActive : ""}`} style={{ backgroundColor: color }} type="button" onClick={() => { setDraft({ ...draft, fillColor: color }); void patchShape({ fillColor: color }); }} />)}
                </div>
                <Text className={s.label}>Outline Color</Text>
                <div className={s.swatches}>
                  {COLORS.map((color) => <button key={`o-${color}`} className={`${s.swatch} ${draft.outlineColor.toLowerCase() === color.toLowerCase() ? s.swatchActive : ""}`} style={{ backgroundColor: color }} type="button" onClick={() => { setDraft({ ...draft, outlineColor: color }); void patchShape({ outlineColor: color }); }} />)}
                </div>
                <Text className={s.label}>Outline Width: {Math.round(draft.outlineWidth)}px</Text>
                <input type="range" min={0} max={12} step={1} value={Math.round(draft.outlineWidth)} onChange={(event) => setDraft({ ...draft, outlineWidth: Number(event.target.value) })} onMouseUp={() => void patchShape({ outlineWidth: draft.outlineWidth })} onTouchEnd={() => void patchShape({ outlineWidth: draft.outlineWidth })} />
                <Text className={s.label}>Effect</Text>
                <select className={s.select} value={draft.effect} onChange={(event) => { const effect = event.target.value as ShapeBuilderEffect; setDraft({ ...draft, effect }); void patchShape({ effect }); }}>
                  <option value="None">None</option>
                  <option value="Shadow">Shadow</option>
                </select>
              </div>
            ) : null}

            {tab === "size" ? (
              <div className={s.panel}>
                <div className={s.grid2}>
                  <div><Text className={s.label}>Width</Text><input className={s.input} type="number" value={Math.round(draft.width)} onChange={(event) => setDraft({ ...draft, width: clamp(Number(event.target.value) || draft.width, 80, 1200) })} onBlur={() => void patchShape({ width: draft.width })} /></div>
                  <div><Text className={s.label}>Height</Text><input className={s.input} type="number" value={Math.round(draft.height)} onChange={(event) => setDraft({ ...draft, height: clamp(Number(event.target.value) || draft.height, 24, 700) })} onBlur={() => void patchShape({ height: draft.height })} /></div>
                </div>
                <Text className={s.label}>Width slider</Text>
                <input type="range" min={80} max={500} value={Math.round(draft.width)} onChange={(event) => setDraft({ ...draft, width: Number(event.target.value) })} onMouseUp={() => void patchShape({ width: draft.width })} onTouchEnd={() => void patchShape({ width: draft.width })} />
                <Text className={s.label}>Height slider</Text>
                <input type="range" min={24} max={220} value={Math.round(draft.height)} onChange={(event) => setDraft({ ...draft, height: Number(event.target.value) })} onMouseUp={() => void patchShape({ height: draft.height })} onTouchEnd={() => void patchShape({ height: draft.height })} />
                <div className={s.grid2}>
                  <div><Text className={s.label}>Position X</Text><input className={s.input} type="number" value={Math.round(draft.left)} onChange={(event) => setDraft({ ...draft, left: clamp(Number(event.target.value) || draft.left, 0, 5000) })} onBlur={() => void patchShape({ left: draft.left })} /></div>
                  <div><Text className={s.label}>Position Y</Text><input className={s.input} type="number" value={Math.round(draft.top)} onChange={(event) => setDraft({ ...draft, top: clamp(Number(event.target.value) || draft.top, 0, 5000) })} onBlur={() => void patchShape({ top: draft.top })} /></div>
                </div>
              </div>
            ) : null}

            {tab === "link" ? (
              <div className={s.panel}>
                <Text className={s.label}>Link Type</Text>
                <div className={s.radioRow}>
                  <label><input type="radio" checked={draft.linkType === "None"} onChange={() => { setDraft({ ...draft, linkType: "None", internalDestinationId: "", externalUrl: "" }); void patchShape({ linkType: "None", internalDestinationId: "", externalUrl: "" }); }} /> No Link</label>
                  <label><input type="radio" checked={draft.linkType === "Internal"} onChange={() => { const target = draft.internalDestinationId || destinations[0]?.id || ""; setDraft({ ...draft, linkType: "Internal", internalDestinationId: target, externalUrl: "" }); void patchShape({ linkType: "Internal", internalDestinationId: target, externalUrl: "" }); }} /> Internal (Sheet / Cell)</label>
                  <label><input type="radio" checked={draft.linkType === "External"} onChange={() => { setDraft({ ...draft, linkType: "External", internalDestinationId: "" }); void patchShape({ linkType: "External", internalDestinationId: "" }); }} /> External URL</label>
                </div>
                {draft.linkType === "Internal" ? (
                  <>
                    <Text className={s.label}>Destination</Text>
                    <select className={s.select} value={draft.internalDestinationId} onChange={(event) => { const internalDestinationId = event.target.value; setDraft({ ...draft, internalDestinationId }); void patchShape({ linkType: "Internal", internalDestinationId, externalUrl: "" }); }}>
                      {destinations.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
                    </select>
                    <Text className={s.label}>Cell Reference (e.g. Sheet2!A1)</Text>
                    <div className={s.horizontal}>
                      <input className={s.input} value={cellRefInput} onChange={(event) => setCellRefInput(event.target.value)} />
                      <button className={s.smallBtn} type="button" onClick={() => { const id = parseCellDestination(cellRefInput); if (!id) { setErr("Invalid cell reference."); return; } setDraft({ ...draft, internalDestinationId: id }); void patchShape({ linkType: "Internal", internalDestinationId: id, externalUrl: "" }); setOk("Cell reference linked."); }}>Use Cell</button>
                      <button className={s.smallBtn} type="button" onClick={() => void activateNavigationDestination(draft.internalDestinationId)}>Test</button>
                    </div>
                  </>
                ) : null}
                {draft.linkType === "External" ? (
                  <>
                    <Text className={s.label}>External URL</Text>
                    <input className={s.input} value={draft.externalUrl} onChange={(event) => setDraft({ ...draft, externalUrl: event.target.value })} onBlur={() => void patchShape({ linkType: "External", externalUrl: draft.externalUrl, internalDestinationId: "" })} />
                    <button className={s.smallBtn} type="button" onClick={() => void activateNavigationDestination(`url::${draft.externalUrl.trim()}`)}>Open URL</button>
                  </>
                ) : null}
              </div>
            ) : null}
          </>
        )}
      </div>

      <div className={s.card}>
        <div className={s.row}><Text className={s.title}>Template Quick Start</Text><Text className={s.muted}>Vertical and Horizontal</Text></div>
        <label><input type="checkbox" checked={applyTemplateAll} onChange={(event) => setApplyTemplateAll(event.target.checked)} /> Apply to all sheets</label>
        <div className={s.horizontal}>
          <button className={s.smallBtn} type="button" onClick={() => void applyTemplate("Vertical")}>Apply Vertical Template</button>
          <button className={s.smallBtn} type="button" onClick={() => void applyTemplate("Horizontal")}>Apply Horizontal Template</button>
        </div>
      </div>

      <div className={s.card}>
        <div className={s.row}><Text className={s.title}>Sheet Format Templates</Text><Text className={s.muted}>Full-sheet capture and apply</Text></div>
        <div className={s.horizontal}>
          <Input value={newFormatName} placeholder="Template name" onChange={(_, data) => setNewFormatName(data.value)} />
          <Button appearance="primary" onClick={() => void captureFormat()}>Capture Active Sheet</Button>
          <button
            className={s.smallBtn}
            type="button"
            onClick={() =>
              void (async () => {
                if (!selectedFormatId) {
                  setErr("Select a format first.");
                  return;
                }
                try {
                  await recaptureSheetFormatFromSelection(selectedFormatId);
                  await refreshAll();
                  setOk("Format recaptured.");
                } catch (error) {
                  setErr(error);
                }
              })()
            }
          >
            Recapture
          </button>
        </div>
        <div className={s.horizontal}>
          <button
            className={s.smallBtn}
            type="button"
            onClick={() =>
              void (async () => {
                if (!selectedFormatId) {
                  setErr("Select a format first.");
                  return;
                }
                try {
                  await applySheetFormat(selectedFormatId);
                  setOk("Format applied to active sheet.");
                } catch (error) {
                  setErr(error);
                }
              })()
            }
          >
            Apply Active
          </button>
          <button
            className={s.smallBtn}
            type="button"
            onClick={() =>
              void (async () => {
                if (!selectedFormatId) {
                  setErr("Select a format first.");
                  return;
                }
                try {
                  await applySheetFormatToAllSheets(selectedFormatId);
                  setOk("Format applied to all sheets.");
                } catch (error) {
                  setErr(error);
                }
              })()
            }
          >
            Apply All Sheets
          </button>
          <button
            className={`${s.smallBtn} ${s.dangerBtn}`}
            type="button"
            onClick={() =>
              void (async () => {
                if (!selectedFormatId) {
                  setErr("Select a format first.");
                  return;
                }
                if (!window.confirm("Delete selected format?")) {
                  return;
                }
                try {
                  await deleteSheetFormat(selectedFormatId);
                  await refreshAll();
                  setOk("Format deleted.");
                } catch (error) {
                  setErr(error);
                }
              })()
            }
          >
            Delete
          </button>
        </div>
        {!formats.length ? <div className={s.empty}>No formats captured yet.</div> : null}
        {formats.map((format) => (
          <div key={format.id} className={s.formatRow} style={{ backgroundColor: selectedFormatId === format.id ? "#EFF4FE" : "#FFFFFF", borderColor: selectedFormatId === format.id ? "#A7C0E7" : undefined }}>
            <div><Text>{format.name}</Text><Text className={s.muted}>Source: {format.sourceSheet} {format.sourceAddress}</Text></div>
            <button className={s.smallBtn} type="button" onClick={() => setSelectedFormatId(format.id)}>{selectedFormatId === format.id ? "Selected" : "Select"}</button>
          </div>
        ))}
      </div>

      <div className={s.card}>
        <div className={s.row}><Text className={s.title}>Formatting Utilities</Text><Text className={s.muted}>Current workbook helpers</Text></div>
        <div className={s.horizontal}>
          <button className={s.smallBtn} type="button" onClick={() => void toggleGridlines()}>Toggle Gridlines</button>
          <button className={s.smallBtn} type="button" onClick={() => void freezeTopRow()}>Freeze Top Row</button>
          <button className={s.smallBtn} type="button" onClick={() => void freezeFirstColumn()}>Freeze First Column</button>
          <button className={s.smallBtn} type="button" onClick={() => void unfreezePanes()}>Unfreeze</button>
          <button className={s.smallBtn} type="button" onClick={() => void refreshPivotTables()}>Refresh Pivots</button>
        </div>
        <div className={s.horizontal}>
          {STYLE_PRESETS.map((preset) => (
            <button key={preset.preset} className={s.smallBtn} type="button" onClick={() => void applyCellStylePreset(preset.preset)}>
              {preset.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default FormatView;
