import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, Text, makeStyles } from "@fluentui/react-components";
import {
  Add20Regular,
  ArrowClockwise20Regular,
  ChevronDown20Regular,
  ChevronUp20Regular,
  Emoji20Regular,
  EyeOff20Regular,
  Eyedropper20Regular,
  Grid20Regular,
  SlideHide20Regular,
  TableFreezeColumn20Regular,
  TableFreezeColumnAndRow20Regular,
  TableFreezeRow20Regular,
  TableResizeColumn20Regular,
  TableResizeRow20Regular,
  TextAlignCenter20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextItalic20Regular,
} from "@fluentui/react-icons";
import {
  ApplyNavigationTemplateRequest,
  CellStylePreset,
  CreateShapeBuilderShapeRequest,
  InsertableShapeType,
  NavigationDestinationOption,
  ShapeBuilderEffect,
  ShapeBuilderShapeRecord,
  ShapeBuilderTextSelectionFormatRequest,
  SheetFormatRecord,
  activateWorksheet,
  activateNavigationDestination,
  applyNavigationTemplate,
  applyCellStylePreset,
  applyFormulaToShapeAnchorCell,
  applySheetFormat,
  applySheetFormatToAllSheets,
  applyTableStyle,
  addShapeOnActiveCell,
  captureSheetFormat,
  createShapeBuilderShape,
  deleteSheetFormat,
  deleteShapeBuilderShape,
  duplicateShapeBuilderShape,
  formatShapeBuilderTextSelection,
  getNamedRangeValueText,
  freezeFirstColumn,
  freezeTopRow,
  FLUENT_ICON_NAME_CATALOG,
  hideActiveSheet,
  hideActiveSheetVeryHidden,
  alignShapeToSelection,
  listNavigationDestinations,
  SHAPE_BUILDER_ICON_OPTIONS,
  listWorkbookThemeColors,
  listShapeBuilderShapes,
  listSheetFormats,
  moveShapeToSelection,
  nudgeShape,
  NO_FILL_COLOR_TOKEN,
  openNewWorkbookWindow,
  openFormatEditorPopout,
  arrangeWorkbookWindows,
  recaptureSheetFormatFromSelection,
  refreshPivotTables,
  renameShape,
  setShapeZOrder,
  setSelectionColumnWidth,
  setSelectionRowHeight,
  toggleGridlines,
  unfreezePanes,
  updateShapeBuilderShape,
} from "../../taskpane";
import { MODERN_TOKENS } from "./designTokens";

interface FormatViewProps {
  onOpenLegacy: () => void;
  isPopout?: boolean;
  embedded?: boolean;
  selectionRequest?: {
    requestId: number;
    sheetName: string;
    shapeName: string;
  } | null;
}

type EditorTab = "shape" | "style" | "size" | "bind" | "link";
type ShapeLinkFilter = "all" | "none" | "internal" | "external";
type UtilityMenu = "view" | "freeze" | "data" | "sheet" | "sizing";
type LayoutSection = "gridFreeze" | "sheetPresentation" | "quickFormatting";
type StyleColorTarget = "fill" | "outline" | "font";
type StyleBaseline = Pick<
  ShapeBuilderShapeRecord,
  | "fillColor"
  | "outlineColor"
  | "outlineWidth"
  | "fontColor"
  | "fontSize"
  | "bold"
  | "italic"
  | "effect"
>;

const SHAPE_TYPES: Array<{ value: InsertableShapeType; label: string }> = [
  { value: "RoundedRectangle", label: "Rounded Rectangle" },
  { value: "Rectangle", label: "Rectangle" },
  { value: "Chevron", label: "Chevron" },
  { value: "Hexagon", label: "Hexagon" },
  { value: "Diamond", label: "Diamond" },
  { value: "Oval", label: "Oval" },
];
const CELL_STYLE_PRESETS: CellStylePreset[] = [
  "Input Cell",
  "Parameter Cell",
  "Header",
  "Subheader",
];
const TABLE_STYLES = ["TableStyleMedium2", "TableStyleMedium9", "TableStyleLight11"];
const UTILITY_MENU_ITEMS: Array<{ key: UtilityMenu; label: string }> = [
  { key: "view", label: "View" },
  { key: "freeze", label: "Freeze" },
  { key: "data", label: "Data" },
  { key: "sheet", label: "Sheet" },
  { key: "sizing", label: "Sizing" },
];
const COMMON_SHAPE_EMOJIS = ["😀", "🙂", "✅", "⚠️", "📌", "📎", "📊", "📍", "🚀", "🧭", "💡", "⭐"];

const DEFAULT_THEME_COLORS = [
  "#4679C7",
  "#54A75A",
  "#E8791B",
  "#CF2E2E",
  "#6742B5",
  "#5AA79D",
  "#EAB51C",
  "#8B97A5",
  "#1E2B3D",
];
const CUSTOM_ICON_KEY_PREFIX = "custom:";

const normalizeError = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);
const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max);
const isHexColor = (value: string) => /^#[0-9a-f]{6}$/i.test(value);
const toUpperHexOrFallback = (value: string, fallback: string) =>
  isHexColor(value) ? value.toUpperCase() : fallback;
const decodeCustomIconValue = (iconKey: string): string => {
  if (!iconKey.toLowerCase().startsWith(CUSTOM_ICON_KEY_PREFIX)) {
    return "";
  }
  const encoded = iconKey.slice(CUSTOM_ICON_KEY_PREFIX.length);
  try {
    return decodeURIComponent(encoded);
  } catch {
    return encoded;
  }
};
const encodeCustomIconKey = (icon: string): string =>
  `${CUSTOM_ICON_KEY_PREFIX}${encodeURIComponent(icon)}`;
const tokenizeSearchValue = (value: string): string[] =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .split(" ")
    .map((part) => part.trim())
    .filter((part) => part.length > 1);

interface EyeDropperResult {
  sRGBHex: string;
}

interface EyeDropperLike {
  open: () => Promise<EyeDropperResult>;
}

const parseCellDestination = (value: string): string | null => {
  const trimmed = value.trim();
  const match = trimmed.match(
    /^(?:'((?:[^']|'')+)'|([^'!]+))!([A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?|[A-Za-z]{1,3}:[A-Za-z]{1,3}|\d+:\d+)$/i
  );
  if (!match) {
    return null;
  }
  const quotedSheetName = match[1]?.replace(/''/g, "'");
  const unquotedSheetName = match[2];
  const sheetName = (quotedSheetName ?? unquotedSheetName ?? "").trim();
  const address = (match[3] ?? "").toUpperCase();
  if (!sheetName || !address) {
    return null;
  }
  return `cell::${sheetName}::${address}`;
};

const compareShapesBySelectionPanePosition = (
  left: ShapeBuilderShapeRecord,
  right: ShapeBuilderShapeRecord
): number =>
  right.zOrderPosition - left.zOrderPosition ||
  left.top - right.top ||
  left.left - right.left ||
  left.shapeName.localeCompare(right.shapeName);

const sortShapesBySelectionPanePosition = (
  records: ShapeBuilderShapeRecord[]
): ShapeBuilderShapeRecord[] => [...records].sort(compareShapesBySelectionPanePosition);

const styles = makeStyles({
  root: { display: "grid", gap: "10px" },
  workspaceHeader: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    boxShadow: "0 1px 2px rgba(17,24,39,0.05)",
    padding: "10px 12px",
    display: "grid",
    gap: "4px",
  },
  workspaceHeaderTop: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    flexWrap: "wrap",
  },
  workspaceTitle: { fontSize: "20px", lineHeight: "22px", fontWeight: 700 },
  workspaceSub: { fontSize: "12px", lineHeight: "16px", color: MODERN_TOKENS.colorTextMuted },
  workspaceActionBar: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "8px",
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  workspaceActionHint: {
    marginLeft: "auto",
    fontSize: "11px",
    lineHeight: "14px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  sectionStack: { display: "grid", gap: "8px", alignContent: "start" },
  sectionCard: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    overflow: "hidden",
    boxShadow: "0 1px 2px rgba(17,24,39,0.05)",
  },
  sectionHeaderBtn: {
    width: "100%",
    border: "none",
    borderRadius: 0,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    padding: "10px 12px",
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
    alignItems: "center",
    textAlign: "left",
    cursor: "pointer",
    selectors: {
      "&:hover": { backgroundColor: "#F8FAFC" },
      "&:focus-visible": {
        outline: `2px solid ${MODERN_TOKENS.colorBrand}`,
        outlineOffset: "-2px",
      },
    },
  },
  sectionHeaderMain: { display: "inline-flex", alignItems: "center", gap: "8px", minWidth: 0 },
  sectionTitle: {
    fontSize: "12px",
    lineHeight: "16px",
    fontWeight: 700,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  sectionMeta: {
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
    color: MODERN_TOKENS.colorTextMuted,
  },
  sectionCount: {
    borderRadius: "999px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
    padding: "2px 6px",
  },
  sectionChip: {
    borderRadius: "999px",
    border: "1px solid #B8CDED",
    backgroundColor: "#E9F1FD",
    color: "#194279",
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
    padding: "2px 6px",
  },
  sectionBody: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "10px",
    display: "grid",
    gap: "10px",
  },
  addBtn: {
    borderRadius: "8px",
    border: `1px solid #B8CDED`,
    backgroundColor: "#E9F1FD",
    color: "#194279",
    padding: "8px 12px",
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
  utilityDock: {
    borderRadius: "0",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "6px 8px",
    display: "grid",
    gap: "6px",
    width: "100%",
    boxSizing: "border-box",
  },
  utilityPinnedCard: {
    position: "static",
    boxShadow: "none",
  },
  utilityHeaderRow: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    alignItems: "center",
    gap: "6px",
  },
  utilityHeadingWrap: { display: "inline-flex", alignItems: "center", gap: "6px", minWidth: 0 },
  utilityTitleSmall: {
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
    letterSpacing: "0.02em",
    textTransform: "uppercase",
    color: MODERN_TOKENS.colorTextMuted,
    whiteSpace: "nowrap",
  },
  utilityActivePill: {
    borderRadius: "999px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "10px",
    lineHeight: "12px",
    fontWeight: 700,
    padding: "2px 6px",
    whiteSpace: "nowrap",
  },
  row: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    flexWrap: "wrap",
  },
  title: {
    fontSize: "13px",
    fontWeight: 700,
    textTransform: "uppercase",
    color: MODERN_TOKENS.colorTextMuted,
  },
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
  previewSvg: { width: "100%", height: "100%", display: "block" },
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
  tabs: {
    display: "grid",
    gridTemplateColumns: "repeat(5,minmax(0,1fr))",
    gap: "6px",
    backgroundColor: "#EEF0F3",
    borderRadius: "10px",
    padding: "6px",
  },
  tab: {
    border: "none",
    borderRadius: "8px",
    backgroundColor: "transparent",
    color: "#6B7280",
    padding: "10px 6px",
    fontSize: "15px",
    fontWeight: 600,
    cursor: "pointer",
  },
  tabActive: {
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    boxShadow: "0 1px 2px rgba(17,24,39,0.08)",
  },
  panel: { display: "grid", gap: "10px" },
  label: {
    fontSize: "11px",
    fontWeight: 700,
    textTransform: "uppercase",
    color: MODERN_TOKENS.colorTextMuted,
  },
  input: {
    width: "100%",
    boxSizing: "border-box",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "9px 10px",
    fontSize: "14px",
  },
  select: {
    width: "100%",
    boxSizing: "border-box",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "9px 10px",
    fontSize: "14px",
    backgroundColor: "#FFFFFF",
  },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: "8px" },
  swatches: { display: "grid", gridTemplateColumns: "repeat(6,minmax(0,1fr))", gap: "8px" },
  swatch: {
    height: "34px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    cursor: "pointer",
  },
  swatchGraphic: { width: "100%", height: "100%", display: "block", borderRadius: "inherit" },
  swatchActive: { boxShadow: "inset 0 0 0 2px #FFFFFF, 0 0 0 2px #4679C7" },
  colorCompactGroup: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
  },
  colorCompactHeader: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
    alignItems: "center",
    padding: "8px",
  },
  colorCurrentBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    color: MODERN_TOKENS.colorText,
    padding: "7px 10px",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    minHeight: "34px",
  },
  colorCurrentMain: { display: "flex", alignItems: "center", gap: "8px", minWidth: 0 },
  colorChip: {
    width: "22px",
    height: "22px",
    borderRadius: "6px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    flexShrink: 0,
  },
  colorGraphic: { width: "100%", height: "100%", display: "block" },
  colorCurrentLabel: {
    fontSize: "12px",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  iconOnlyBtn: {
    width: "34px",
    height: "34px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  iconOnlyBtnActive: { backgroundColor: "#EFF4FE", border: "1px solid #7EA5E1", color: "#173A70" },
  colorCompactBody: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  noFillSwatch: {
    borderRadius: "8px",
    border: `1px dashed ${MODERN_TOKENS.colorBorder}`,
    background: "repeating-linear-gradient(-45deg,#FFFFFF,#FFFFFF 6px,#EEF2F6 6px,#EEF2F6 12px)",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "11px",
    color: MODERN_TOKENS.colorTextMuted,
    cursor: "pointer",
    height: "34px",
  },
  toggleGroup: { display: "flex", gap: "8px", flexWrap: "wrap" },
  toggleBtnActive: { backgroundColor: "#EFF4FE", border: "1px solid #7EA5E1", color: "#173A70" },
  colorInput: {
    width: "100%",
    height: "36px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "2px 6px",
    backgroundColor: "#FFFFFF",
    cursor: "pointer",
  },
  iconGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(5,minmax(0,1fr))",
    gap: "8px",
    maxHeight: "240px",
    overflowY: "auto",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "8px",
  },
  iconBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "8px 6px",
    display: "grid",
    justifyItems: "center",
    gap: "2px",
    cursor: "pointer",
    fontSize: "11px",
  },
  iconActive: { border: "1px solid #7EA5E1", backgroundColor: "#EFF4FE" },
  textEditorCard: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "8px",
    display: "grid",
    gap: "8px",
  },
  textToolbar: { display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center" },
  textToolbarDivider: { width: "1px", height: "22px", backgroundColor: MODERN_TOKENS.colorBorder },
  textToolbarHint: { fontSize: "11px", color: MODERN_TOKENS.colorTextMuted },
  emojiPicker: {
    display: "grid",
    gridTemplateColumns: "repeat(6,minmax(0,1fr))",
    gap: "6px",
  },
  emojiBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    minHeight: "38px",
    fontSize: "20px",
    cursor: "pointer",
  },
  textArea: {
    width: "100%",
    minHeight: "76px",
    resize: "vertical",
    boxSizing: "border-box",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "8px 10px",
    fontSize: "14px",
    lineHeight: "20px",
    fontFamily: "Segoe UI, sans-serif",
    backgroundColor: "#FFFFFF",
  },
  textAreaExpanded: { minHeight: "180px" },
  textAreaAlignLeft: { textAlign: "left" },
  textAreaAlignCenter: { textAlign: "center" },
  textAreaAlignRight: { textAlign: "right" },
  compactSelect: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    height: "34px",
    padding: "0 10px",
    fontSize: "12px",
  },
  compactColorInput: {
    width: "34px",
    height: "34px",
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    padding: "2px",
    backgroundColor: "#FFFFFF",
    cursor: "pointer",
  },
  compactNumberInput: { width: "84px", padding: "7px 8px" },
  iconCount: { fontSize: "11px", color: MODERN_TOKENS.colorTextMuted },
  radioRow: { display: "grid", gap: "8px" },
  horizontal: { display: "flex", gap: "8px", flexWrap: "wrap" },
  buttonGrid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: "8px" },
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "10px",
    flexWrap: "wrap",
  },
  sectionTitleWrap: { display: "grid", gap: "2px" },
  helperCard: {
    borderRadius: "10px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#F8FAFC",
    padding: "10px",
    display: "grid",
    gap: "8px",
  },
  radioCard: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "8px 10px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  utilityMenuBar: {
    display: "flex",
    gap: "0px",
    flexWrap: "wrap",
    minWidth: 0,
    justifyContent: "center",
  },
  utilityMenuTab: {
    borderRadius: "999px",
    border: "1px solid transparent",
    backgroundColor: "#F8FAFC",
    color: MODERN_TOKENS.colorTextMuted,
    padding: "4px 8px",
    fontSize: "11px",
    fontWeight: 700,
    cursor: "pointer",
    minWidth: "auto",
    textAlign: "center",
    lineHeight: "14px",
    whiteSpace: "nowrap",
  },
  utilityMenuTabActive: {
    backgroundColor: "#FFFFFF",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    color: MODERN_TOKENS.colorText,
    boxShadow: "0 1px 2px rgba(17,24,39,0.08)",
  },
  utilityPanel: {
    borderTop: `1px solid ${MODERN_TOKENS.colorBorder}`,
    paddingTop: "6px",
    display: "grid",
    gap: "0px",
  },
  utilityActionsRow: { display: "flex", flexWrap: "wrap", gap: "8px" },
  utilityBtn: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    color: MODERN_TOKENS.colorText,
    padding: "6px 8px",
    fontSize: "11px",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: "8px",
    justifyContent: "center",
    minHeight: "30px",
    whiteSpace: "nowrap",
  },
  utilitySizeGrid: { display: "flex", flexWrap: "wrap", gap: "8px" },
  utilitySizeRow: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
    alignItems: "center",
  },
  utilitySizeInput: {
    width: "100%",
    boxSizing: "border-box",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    borderRadius: "8px",
    padding: "7px 10px",
    fontSize: "13px",
  },
  formatRow: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: "#FFFFFF",
    padding: "8px 10px",
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
    alignItems: "center",
  },
  formatRowSelected: { backgroundColor: "#EFF4FE", border: "1px solid #A7C0E7" },
  empty: {
    borderRadius: "8px",
    border: `1px dashed ${MODERN_TOKENS.colorBorder}`,
    padding: "12px",
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
  },
});

const ColorSwatchGraphic: React.FC<{ className?: string; color?: string; noFill?: boolean }> = ({
  className,
  color = "#FFFFFF",
  noFill = false,
}) => {
  const patternId = React.useId().replace(/:/g, "-");

  return (
    <svg className={className} viewBox="0 0 22 22" aria-hidden="true" focusable="false">
      {noFill ? (
        <defs>
          <pattern
            id={patternId}
            width="6"
            height="6"
            patternUnits="userSpaceOnUse"
            patternTransform="rotate(-45)"
          >
            <rect width="6" height="6" fill="#FFFFFF" />
            <rect width="3" height="6" fill="#EEF2F6" />
          </pattern>
        </defs>
      ) : null}
      <rect
        x="0"
        y="0"
        width="22"
        height="22"
        rx="6"
        fill={noFill ? `url(#${patternId})` : color}
      />
    </svg>
  );
};

const ShapePreviewGraphic: React.FC<{ className: string; shape: ShapeBuilderShapeRecord }> = ({
  className,
  shape,
}) => {
  const patternId = React.useId().replace(/:/g, "-");
  const previewText = (shape.text || "Shape").trim() || "Shape";
  const label = previewText.length > 12 ? `${previewText.slice(0, 11)}...` : previewText;
  const strokeWidth = Math.max(1, Math.min(shape.outlineWidth, 6));
  const inset = strokeWidth / 2;
  const fillIsTransparent = shape.fillColor === NO_FILL_COLOR_TOKEN;

  return (
    <svg className={className} viewBox="0 0 82 34" aria-hidden="true" focusable="false">
      {fillIsTransparent ? (
        <defs>
          <pattern
            id={patternId}
            width="6"
            height="6"
            patternUnits="userSpaceOnUse"
            patternTransform="rotate(-45)"
          >
            <rect width="6" height="6" fill="#FFFFFF" />
            <rect width="3" height="6" fill="#EEF2F6" />
          </pattern>
        </defs>
      ) : null}
      <rect
        x={inset}
        y={inset}
        width={82 - strokeWidth}
        height={34 - strokeWidth}
        rx="7"
        fill={fillIsTransparent ? `url(#${patternId})` : shape.fillColor}
        stroke={shape.outlineColor}
        strokeWidth={strokeWidth}
      />
      <text
        x="41"
        y="17"
        fill={shape.fontColor}
        fontSize="12"
        fontWeight="700"
        textAnchor="middle"
        dominantBaseline="middle"
      >
        {label}
      </text>
    </svg>
  );
};

const FormatView: React.FC<FormatViewProps> = ({
  onOpenLegacy,
  isPopout = false,
  embedded = false,
  selectionRequest = null,
}) => {
  const s = styles();
  const [shapes, setShapes] = useState<ShapeBuilderShapeRecord[]>([]);
  const [destinations, setDestinations] = useState<NavigationDestinationOption[]>([]);
  const [formats, setFormats] = useState<SheetFormatRecord[]>([]);
  const [themeColors, setThemeColors] = useState<string[]>(DEFAULT_THEME_COLORS);
  const [selectedShapeId, setSelectedShapeId] = useState("");
  const [selectedFormatId, setSelectedFormatId] = useState("");
  const [newFormatName, setNewFormatName] = useState("");
  const [tab, setTab] = useState<EditorTab>("shape");
  const [draft, setDraft] = useState<ShapeBuilderShapeRecord | null>(null);
  const [shapeFilterQuery, setShapeFilterQuery] = useState("");
  const [shapeLinkFilter, setShapeLinkFilter] = useState<ShapeLinkFilter>("all");
  const [iconQuery, setIconQuery] = useState("");
  const [customIconInput, setCustomIconInput] = useState("");
  const [cellRefInput, setCellRefInput] = useState("");
  const [namedRangeBindingInput, setNamedRangeBindingInput] = useState("");
  const [anchorFormulaInput, setAnchorFormulaInput] = useState("");
  const [tableStyleInput, setTableStyleInput] = useState(TABLE_STYLES[0]);
  const [openColorTarget, setOpenColorTarget] = useState<StyleColorTarget | null>(null);
  const [textEditorExpanded, setTextEditorExpanded] = useState(false);
  const [emojiPickerOpen, setEmojiPickerOpen] = useState(false);
  const [columnWidthInput, setColumnWidthInput] = useState("14");
  const [rowHeightInput, setRowHeightInput] = useState("18");
  const [utilityMenu, setUtilityMenu] = useState<UtilityMenu>("view");
  const [expandedSections, setExpandedSections] = useState<LayoutSection[]>(["gridFreeze"]);
  const [applyTemplateAll, setApplyTemplateAll] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [statusType, setStatusType] = useState<"ok" | "error">("ok");
  const [styleBaseline, setStyleBaseline] = useState<StyleBaseline | null>(null);
  const [confirmDeleteShapeId, setConfirmDeleteShapeId] = useState("");
  const shapeHandlersRef = useRef<Array<{ remove: () => Promise<void> | void }>>([]);
  const selectedShapeIdRef = useRef("");
  const editorCardRef = useRef<HTMLDivElement | null>(null);
  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);
  const styleBaselineShapeIdRef = useRef<string>("");

  const selectedShape = useMemo(
    () => shapes.find((x) => x.id === selectedShapeId) ?? null,
    [shapes, selectedShapeId]
  );
  const fluentIconMatches = useMemo(() => {
    const query = iconQuery.trim().toLowerCase();
    if (!query) {
      return [];
    }
    return FLUENT_ICON_NAME_CATALOG.filter((name) => name.includes(query));
  }, [iconQuery]);
  const fluentIconMatchKeywordSet = useMemo(() => {
    const set = new Set<string>();
    fluentIconMatches.forEach((name) => {
      tokenizeSearchValue(name).forEach((token) => set.add(token));
    });
    return set;
  }, [fluentIconMatches]);
  const filteredIcons = useMemo(() => {
    const q = iconQuery.trim().toLowerCase();
    if (!q) {
      return SHAPE_BUILDER_ICON_OPTIONS;
    }
    return SHAPE_BUILDER_ICON_OPTIONS.filter((icon) => {
      const directMatch =
        icon.key.toLowerCase().includes(q) ||
        icon.label.toLowerCase().includes(q) ||
        icon.symbol.toLowerCase().includes(q) ||
        icon.keywords.some((keyword) => keyword.includes(q));
      if (directMatch) {
        return true;
      }
      if (!fluentIconMatchKeywordSet.size) {
        return false;
      }
      return icon.keywords.some((keyword) => fluentIconMatchKeywordSet.has(keyword));
    });
  }, [iconQuery, fluentIconMatchKeywordSet]);
  const filteredShapes = useMemo(() => {
    const query = shapeFilterQuery.trim().toLowerCase();
    return shapes.filter((shape) => {
      if (shapeLinkFilter !== "all" && shape.linkType.toLowerCase() !== shapeLinkFilter) {
        return false;
      }
      if (!query) {
        return true;
      }
      return (
        shape.text.toLowerCase().includes(query) ||
        shape.shapeName.toLowerCase().includes(query) ||
        shape.shapeType.toLowerCase().includes(query) ||
        shape.iconKey.toLowerCase().includes(query) ||
        shape.linkType.toLowerCase().includes(query)
      );
    });
  }, [shapes, shapeFilterQuery, shapeLinkFilter]);
  const isShapeListFiltered = shapeFilterQuery.trim().length > 0 || shapeLinkFilter !== "all";
  selectedShapeIdRef.current = selectedShapeId;

  const setOk = (message: string) => {
    setStatusType("ok");
    setStatus(message);
  };

  const setErr = (error: unknown) => {
    setStatusType("error");
    setStatus(normalizeError(error));
  };

  const applyShapeRecords = useCallback(
    (shapeRecords: ShapeBuilderShapeRecord[], preferredShapeId?: string) => {
      setShapes(shapeRecords);
      setConfirmDeleteShapeId((current) =>
        current && !shapeRecords.some((shape) => shape.id === current) ? "" : current
      );
      setSelectedShapeId((prev) => {
        const nextId = preferredShapeId ?? prev;
        return shapeRecords.some((shape) => shape.id === nextId) ? nextId : (shapeRecords[0]?.id ?? "");
      });
    },
    []
  );

  const refreshShapes = useCallback(
    async (preferredShapeId?: string) => {
      setBusy(true);
      try {
        const shapeRecords = sortShapesBySelectionPanePosition(await listShapeBuilderShapes());
        applyShapeRecords(shapeRecords, preferredShapeId);
      } catch (error) {
        setErr(error);
      } finally {
        setBusy(false);
      }
    },
    [applyShapeRecords]
  );

  const refreshAll = useCallback(async () => {
    setBusy(true);
    try {
      const settle = <T,>(promise: Promise<T>) =>
        promise.then(
          (value): { status: "fulfilled"; value: T } => ({ status: "fulfilled", value }),
          (reason): { status: "rejected"; reason: unknown } => ({ status: "rejected", reason })
        );
      const [shapeResult, destinationResult, formatResult, themeResult] = await Promise.all([
        settle(listShapeBuilderShapes()),
        settle(listNavigationDestinations()),
        settle(listSheetFormats()),
        settle(listWorkbookThemeColors()),
      ]);
      const failures: string[] = [];

      if (shapeResult.status === "fulfilled") {
        applyShapeRecords(sortShapesBySelectionPanePosition(shapeResult.value));
      } else {
        setShapes([]);
        setSelectedShapeId("");
        setConfirmDeleteShapeId("");
        failures.push(`Shapes: ${normalizeError(shapeResult.reason)}`);
      }

      if (destinationResult.status === "fulfilled") {
        setDestinations(destinationResult.value);
      } else {
        setDestinations([]);
        failures.push(`Destinations: ${normalizeError(destinationResult.reason)}`);
      }

      if (formatResult.status === "fulfilled") {
        const formatRecords = formatResult.value;
        setFormats(formatRecords);
        setSelectedFormatId((prev) =>
          formatRecords.some((x) => x.id === prev) ? prev : (formatRecords[0]?.id ?? "")
        );
      } else {
        setFormats([]);
        setSelectedFormatId("");
        failures.push(`Formats: ${normalizeError(formatResult.reason)}`);
      }

      if (themeResult.status === "fulfilled") {
        setThemeColors(themeResult.value.length ? themeResult.value : DEFAULT_THEME_COLORS);
      } else {
        setThemeColors(DEFAULT_THEME_COLORS);
        failures.push(`Theme colors: ${normalizeError(themeResult.reason)}`);
      }

      if (failures.length > 0) {
        setStatusType("error");
        setStatus(failures.join(" | "));
      }
    } finally {
      setBusy(false);
    }
  }, [applyShapeRecords]);

  const refreshAndSelectShape = useCallback(
    async (sheetName: string, shapeName: string) => {
      await refreshShapes(`${sheetName}::${shapeName}`);
      setExpandedSections(["sheetPresentation"]);
      setTab("shape");
    },
    [refreshShapes]
  );

  useEffect(() => {
    void refreshAll();
  }, [refreshAll]);

  useEffect(() => {
    if (!selectionRequest) {
      return undefined;
    }

    let disposed = false;
    void (async () => {
      try {
        await activateWorksheet(selectionRequest.sheetName);
        if (disposed) {
          return;
        }
        await refreshAndSelectShape(selectionRequest.sheetName, selectionRequest.shapeName);
      } catch (error) {
        if (!disposed) {
          setErr(error);
        }
      }
    })();

    return () => {
      disposed = true;
    };
  }, [refreshAndSelectShape, selectionRequest]);

  useEffect(() => {
    setDraft(selectedShape ? { ...selectedShape } : null);
    if (!selectedShape) {
      styleBaselineShapeIdRef.current = "";
      setStyleBaseline(null);
    } else if (styleBaselineShapeIdRef.current !== selectedShape.id) {
      styleBaselineShapeIdRef.current = selectedShape.id;
      setStyleBaseline({
        fillColor: selectedShape.fillColor,
        outlineColor: selectedShape.outlineColor,
        outlineWidth: selectedShape.outlineWidth,
        fontColor: selectedShape.fontColor,
        fontSize: selectedShape.fontSize,
        bold: selectedShape.bold,
        italic: selectedShape.italic,
        effect: selectedShape.effect,
      });
    }
    setIconQuery("");
    setCustomIconInput(selectedShape ? decodeCustomIconValue(selectedShape.iconKey) : "");
    setCellRefInput("");
    setNamedRangeBindingInput("");
    setAnchorFormulaInput("");
    setOpenColorTarget(null);
    setTextEditorExpanded(false);
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

    const registerHandlers = async () => {
      if (typeof Excel === "undefined") {
        return;
      }
      clearShapeHandlers();
      if (!shapes.length) {
        return;
      }

      try {
        await Excel.run(async (context) => {
          for (const shapeRecord of shapes) {
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
              setSelectedShapeId(shapeRecord.id);
              if (
                (shapeRecord.linkType === "Internal" && !shapeRecord.internalDestinationId) ||
                (shapeRecord.linkType === "External" && !shapeRecord.externalUrl) ||
                shapeRecord.linkType === "None"
              ) {
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
      const targetShapeId = selectedShape.id;
      try {
        const updated = await updateShapeBuilderShape(
          selectedShape.sheetName,
          selectedShape.shapeName,
          patch
        );
        setShapes((prev) =>
          sortShapesBySelectionPanePosition(prev.map((x) => (x.id === targetShapeId ? updated : x)))
        );
        if (selectedShapeIdRef.current === targetShapeId) {
          setDraft((prev) => (prev && prev.id === targetShapeId ? updated : prev));
        }
      } catch (error) {
        setErr(error);
      }
    },
    [selectedShape]
  );

  const applyTextSelectionFormat = useCallback(
    async (updates: ShapeBuilderTextSelectionFormatRequest) => {
      if (!draft) {
        return;
      }
      const targetShapeId = draft.id;
      const textArea = textAreaRef.current;
      const selectionStart = textArea
        ? Math.min(textArea.selectionStart, textArea.selectionEnd)
        : 0;
      const selectionEnd = textArea ? Math.max(textArea.selectionStart, textArea.selectionEnd) : 0;
      const selectionLength = selectionEnd - selectionStart;
      if (selectionLength <= 0) {
        setErr("Select text in the editor first.");
        return;
      }

      try {
        const updated = await formatShapeBuilderTextSelection(
          draft.sheetName,
          draft.shapeName,
          selectionStart,
          selectionLength,
          updates
        );
        setShapes((prev) =>
          sortShapesBySelectionPanePosition(
            prev.map((item) => (item.id === updated.id ? updated : item))
          )
        );
        if (selectedShapeIdRef.current === targetShapeId) {
          setDraft((prev) => (prev && prev.id === targetShapeId ? updated : prev));
        }
        window.setTimeout(() => {
          const editor = textAreaRef.current;
          if (!editor) {
            return;
          }
          editor.focus();
          editor.setSelectionRange(selectionStart, selectionEnd);
        }, 0);
      } catch (error) {
        setErr(error);
      }
    },
    [draft]
  );

  const runUtilityAction = useCallback(
    async (successMessage: string, action: () => Promise<void>) => {
      try {
        await action();
        setOk(successMessage);
      } catch (error) {
        setErr(error);
      }
    },
    []
  );

  const applyFillColor = (value: string) => {
    if (!draft) {
      return;
    }
    const fillColor =
      value === NO_FILL_COLOR_TOKEN
        ? NO_FILL_COLOR_TOKEN
        : toUpperHexOrFallback(value, draft.fillColor);
    setDraft({ ...draft, fillColor });
    void patchShape({ fillColor });
  };

  const applyOutlineColor = (value: string) => {
    if (!draft) {
      return;
    }
    const outlineColor = toUpperHexOrFallback(value, draft.outlineColor);
    setDraft({ ...draft, outlineColor });
    void patchShape({ outlineColor });
  };

  const applyFontColor = (value: string) => {
    if (!draft) {
      return;
    }
    const fontColor = toUpperHexOrFallback(value, draft.fontColor);
    setDraft({ ...draft, fontColor });
    void patchShape({ fontColor });
  };

  const insertEmojiIntoText = (emoji: string) => {
    if (!draft) {
      return;
    }
    const editor = textAreaRef.current;
    const currentText = draft.text || "";
    const start = editor ? editor.selectionStart : currentText.length;
    const end = editor ? editor.selectionEnd : currentText.length;
    const nextText = `${currentText.slice(0, start)}${emoji}${currentText.slice(end)}`;
    const nextCursor = start + emoji.length;
    setDraft({ ...draft, text: nextText });
    setEmojiPickerOpen(false);
    void patchShape({ text: nextText });
    window.setTimeout(() => {
      const nextEditor = textAreaRef.current;
      if (!nextEditor) {
        return;
      }
      nextEditor.focus();
      nextEditor.setSelectionRange(nextCursor, nextCursor);
    }, 0);
    setOk(`Inserted ${emoji} into the shape text.`);
  };

  const pickColorFromScreen = async (target: StyleColorTarget) => {
    if (!draft) {
      return;
    }
    const EyeDropperCtor = (window as Window & { EyeDropper?: new () => EyeDropperLike })
      .EyeDropper;
    if (!EyeDropperCtor) {
      setErr("Eye dropper is not available in this Excel host.");
      return;
    }
    try {
      const eyedropper = new EyeDropperCtor();
      const result = await eyedropper.open();
      if (!isHexColor(result.sRGBHex)) {
        setErr("Picked color was invalid.");
        return;
      }
      if (target === "fill") {
        applyFillColor(result.sRGBHex);
      } else if (target === "outline") {
        applyOutlineColor(result.sRGBHex);
      } else {
        applyFontColor(result.sRGBHex);
      }
    } catch (error) {
      const message = normalizeError(error).toLowerCase();
      if (message.includes("abort") || message.includes("cancel")) {
        return;
      }
      setErr(error);
    }
  };

  const revertStyleChanges = () => {
    if (!draft || !styleBaseline) {
      return;
    }
    const patch: Partial<CreateShapeBuilderShapeRequest> = {
      fillColor: styleBaseline.fillColor,
      outlineColor: styleBaseline.outlineColor,
      outlineWidth: styleBaseline.outlineWidth,
      fontColor: styleBaseline.fontColor,
      fontSize: styleBaseline.fontSize,
      bold: styleBaseline.bold,
      italic: styleBaseline.italic,
      effect: styleBaseline.effect,
    };
    setDraft({
      ...draft,
      fillColor: styleBaseline.fillColor,
      outlineColor: styleBaseline.outlineColor,
      outlineWidth: styleBaseline.outlineWidth,
      fontColor: styleBaseline.fontColor,
      fontSize: styleBaseline.fontSize,
      bold: styleBaseline.bold,
      italic: styleBaseline.italic,
      effect: styleBaseline.effect,
    });
    void patchShape(patch);
    setOk("Style reverted.");
  };

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
        textHorizontalAlignment: "Left",
        textVerticalAlignment: "Middle",
        linkType: "None",
        internalDestinationId: "",
        externalUrl: "",
        effect: "None",
      });
      setShapes((prev) => sortShapesBySelectionPanePosition([...prev, created]));
      setSelectedShapeId(created.id);
      setConfirmDeleteShapeId("");
      setTab("shape");
      setOk("Shape added.");
    } catch (error) {
      setErr(error);
    } finally {
      setBusy(false);
    }
  };

  const insertShapeOnActiveCell = async () => {
    setBusy(true);
    try {
      const inserted = await addShapeOnActiveCell({
        shapeType: draft?.shapeType ?? "RoundedRectangle",
        text: draft?.text?.trim() || `Button ${shapes.length + 1}`,
        fillColor:
          draft?.fillColor && draft.fillColor !== NO_FILL_COLOR_TOKEN ? draft.fillColor : "#A8D5AD",
        outlineColor: draft?.outlineColor ?? "#8CBF95",
        fontColor: draft?.fontColor ?? "#1F2937",
      });
      await refreshAndSelectShape(inserted.sheet, inserted.name);
      setConfirmDeleteShapeId("");
      setOk(`Inserted "${inserted.name}" on ${inserted.sheet}!${inserted.anchorAddress}.`);
    } catch (error) {
      setErr(error);
    } finally {
      setBusy(false);
    }
  };

  const duplicateShape = async (shape: ShapeBuilderShapeRecord) => {
    try {
      const copy = await duplicateShapeBuilderShape(shape.sheetName, shape.shapeName);
      setShapes((prev) => sortShapesBySelectionPanePosition([...prev, copy]));
      setSelectedShapeId(copy.id);
      setConfirmDeleteShapeId("");
      setOk("Shape duplicated.");
    } catch (error) {
      setErr(error);
    }
  };

  const deleteShape = async (shape: ShapeBuilderShapeRecord) => {
    try {
      await deleteShapeBuilderShape(shape.sheetName, shape.shapeName);
      setConfirmDeleteShapeId("");
      setShapes((prev) => {
        const next = sortShapesBySelectionPanePosition(prev.filter((x) => x.id !== shape.id));
        setSelectedShapeId((current) => (current === shape.id ? (next[0]?.id ?? "") : current));
        return next;
      });
      setOk("Shape deleted.");
    } catch (error) {
      setErr(error);
    }
  };

  const editShape = (shape: ShapeBuilderShapeRecord) => {
    setSelectedShapeId(shape.id);
    setTab("shape");
    editorCardRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const renameSelectedShape = async () => {
    if (!draft || !selectedShape) {
      return;
    }
    const nextName = draft.shapeName.trim();
    if (!nextName) {
      setErr("Enter a shape name first.");
      return;
    }
    if (nextName === selectedShape.shapeName) {
      return;
    }
    try {
      await renameShape(selectedShape.sheetName, selectedShape.shapeName, nextName);
      await refreshAndSelectShape(selectedShape.sheetName, nextName);
      setConfirmDeleteShapeId("");
      setOk(`Renamed shape to "${nextName}".`);
    } catch (error) {
      setErr(error);
    }
  };

  const nudgeSelectedShape = async (dx: number, dy: number) => {
    if (!selectedShape) {
      return;
    }
    try {
      await nudgeShape(selectedShape.sheetName, selectedShape.shapeName, dx, dy);
      await refreshAndSelectShape(selectedShape.sheetName, selectedShape.shapeName);
      setConfirmDeleteShapeId("");
      setOk("Shape nudged.");
    } catch (error) {
      setErr(error);
    }
  };

  const alignSelectedShapeToSelection = async (
    alignment: "Left" | "Center" | "Right" | "Top" | "Middle" | "Bottom"
  ) => {
    if (!selectedShape) {
      return;
    }
    try {
      await alignShapeToSelection(selectedShape.sheetName, selectedShape.shapeName, alignment);
      await refreshAndSelectShape(selectedShape.sheetName, selectedShape.shapeName);
      setConfirmDeleteShapeId("");
      setOk(`Aligned shape ${alignment.toLowerCase()}.`);
    } catch (error) {
      setErr(error);
    }
  };

  const updateSelectedShapeOrder = async (
    order: "BringToFront" | "BringForward" | "SendToBack" | "SendBackward"
  ) => {
    if (!selectedShape) {
      return;
    }
    try {
      await setShapeZOrder(selectedShape.sheetName, selectedShape.shapeName, order);
      await refreshAndSelectShape(selectedShape.sheetName, selectedShape.shapeName);
      setConfirmDeleteShapeId("");
      setOk("Shape order updated.");
    } catch (error) {
      setErr(error);
    }
  };

  const moveSelectedShapeToSelection = async () => {
    if (!selectedShape) {
      return;
    }
    if (!selectedShape.anchorAddress) {
      setErr("This shape does not have an anchor yet. Insert it on the active cell first.");
      return;
    }
    try {
      await moveShapeToSelection(
        selectedShape.sheetName,
        selectedShape.shapeName,
        selectedShape.anchorAddress
      );
      await refreshAndSelectShape(selectedShape.sheetName, selectedShape.shapeName);
      setConfirmDeleteShapeId("");
      setOk("Shape moved to the current selection.");
    } catch (error) {
      setErr(error);
    }
  };

  const applyNamedRangeBinding = async () => {
    if (!draft) {
      return;
    }
    const rangeName = namedRangeBindingInput.trim();
    if (!rangeName) {
      setErr("Enter a named range first.");
      return;
    }
    try {
      const valueText = await getNamedRangeValueText(rangeName);
      setNamedRangeBindingInput(rangeName);
      setDraft({ ...draft, text: valueText });
      await patchShape({ text: valueText });
      setOk(`Copied "${rangeName}" into shape text.`);
    } catch (error) {
      setErr(error);
    }
  };

  const applyAnchorFormula = async () => {
    if (!draft) {
      return;
    }
    const formula = anchorFormulaInput.trim();
    if (!formula) {
      setErr("Enter a formula first.");
      return;
    }
    if (!draft.anchorAddress) {
      setErr("This shape does not have an anchor cell yet.");
      return;
    }
    try {
      await applyFormulaToShapeAnchorCell(
        draft.sheetName,
        draft.anchorAddress,
        formula,
        draft.shapeName
      );
      await refreshAndSelectShape(draft.sheetName, draft.shapeName);
      setConfirmDeleteShapeId("");
      setOk("Anchor-cell formula applied.");
    } catch (error) {
      setErr(error);
    }
  };

  const applyCellStyleQuickAction = async (preset: CellStylePreset) => {
    try {
      await applyCellStylePreset(preset);
      setOk(`${preset} applied.`);
    } catch (error) {
      setErr(error);
    }
  };

  const applyTableStyleQuickAction = async () => {
    try {
      await applyTableStyle(tableStyleInput);
      setOk(`Applied table style ${tableStyleInput}.`);
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

  const applyCustomIcon = () => {
    if (!draft) {
      return;
    }
    const symbol = customIconInput;
    if (!symbol) {
      setErr("Paste an icon or emoji first.");
      return;
    }
    const iconKey = encodeCustomIconKey(symbol);
    setDraft({ ...draft, iconKey });
    void patchShape({ iconKey });
    setOk("Custom icon applied.");
  };

  const applyColumnWidth = () => {
    const width = Number(columnWidthInput);
    if (!Number.isFinite(width) || width <= 0) {
      setErr("Enter a valid column width.");
      return;
    }
    void runUtilityAction(
      "Column width applied to selection (or entire sheet when a single cell is selected).",
      () => setSelectionColumnWidth(width)
    );
  };

  const applyRowHeight = () => {
    const height = Number(rowHeightInput);
    if (!Number.isFinite(height) || height <= 0) {
      setErr("Enter a valid row height.");
      return;
    }
    void runUtilityAction(
      "Row height applied to selection (or entire sheet when a single cell is selected).",
      () => setSelectionRowHeight(height)
    );
  };

  const openLayoutSection = useCallback((section: LayoutSection) => {
    setExpandedSections((prev) => (prev.includes(section) ? prev : [section]));
  }, []);

  const toggleLayoutSection = useCallback((section: LayoutSection) => {
    setExpandedSections((prev) => (prev.includes(section) ? [] : [section]));
  }, []);

  const applySelectedFormatActive = async () => {
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
  };

  const startShapeBuilder = useCallback(() => {
    setConfirmDeleteShapeId("");
    openLayoutSection("sheetPresentation");
    void addShape();
  }, [addShape, openLayoutSection]);

  const insertShapeAtSelection = useCallback(() => {
    setConfirmDeleteShapeId("");
    openLayoutSection("sheetPresentation");
    void insertShapeOnActiveCell();
  }, [insertShapeOnActiveCell, openLayoutSection]);

  const openShapeBuilderPopout = useCallback(() => {
    void openFormatEditorPopout();
  }, []);

  const renderShapeBuilderActions = (variant: "toolbar" | "card") => (
    <>
      <button
        className={variant === "toolbar" ? s.addBtn : s.smallBtn}
        type="button"
        onClick={startShapeBuilder}
        disabled={busy}
      >
        <Add20Regular /> Start Shape Builder
      </button>
      <button className={s.smallBtn} type="button" onClick={insertShapeAtSelection} disabled={busy}>
        Insert On Active Cell
      </button>
      <button className={s.smallBtn} type="button" onClick={() => void refreshShapes()} disabled={busy}>
        Refresh
      </button>
      {!isPopout ? (
        <button className={s.smallBtn} type="button" onClick={openShapeBuilderPopout}>
          Pop Out
        </button>
      ) : null}
    </>
  );

  const fillDisplayColor = draft
    ? draft.fillColor === NO_FILL_COLOR_TOKEN
      ? NO_FILL_COLOR_TOKEN
      : toUpperHexOrFallback(draft.fillColor, "#A8D5AD")
    : "#A8D5AD";
  const outlineDisplayColor = draft
    ? toUpperHexOrFallback(draft.outlineColor, "#8CBF95")
    : "#8CBF95";
  const fontDisplayColor = draft ? toUpperHexOrFallback(draft.fontColor, "#1F2937") : "#1F2937";
  const activeUtilityMenuLabel =
    UTILITY_MENU_ITEMS.find((item) => item.key === utilityMenu)?.label ?? "Utilities";
  const formattingUtilitiesCard = (
    <div className={`${s.utilityDock} ${s.utilityPinnedCard}`}>
      <div className={s.utilityHeaderRow}>
        <div className={s.utilityHeadingWrap}>
          <Text className={s.utilityTitleSmall}>Formatting Utilities</Text>
          <span className={s.utilityActivePill}>{activeUtilityMenuLabel}</span>
        </div>
        <div className={s.utilityMenuBar}>
          {UTILITY_MENU_ITEMS.map((item) => (
            <button
              key={item.key}
              className={`${s.utilityMenuTab} ${utilityMenu === item.key ? s.utilityMenuTabActive : ""}`}
              type="button"
              onClick={() => setUtilityMenu(item.key)}
            >
              {item.label}
            </button>
          ))}
        </div>
      </div>
      <div className={s.utilityPanel}>
        {utilityMenu === "view" ? (
          <>
            <div className={s.utilityActionsRow}>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() => void runUtilityAction("Gridlines toggled.", toggleGridlines)}
              >
                <Grid20Regular /> Toggle Gridlines
              </button>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() =>
                  void runUtilityAction("New workbook window opened.", openNewWorkbookWindow)
                }
              >
                <Add20Regular /> New Window
              </button>
            </div>
            <Text className={s.label}>Arrange Windows</Text>
            <Text className={s.muted}>
              Excel desktop supports Tiled, Horizontal, Vertical, and Cascade layouts.
            </Text>
            <div className={s.utilityActionsRow}>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() =>
                  void runUtilityAction("Windows arranged (tiled).", () =>
                    arrangeWorkbookWindows("Tiled")
                  )
                }
              >
                Tiled
              </button>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() =>
                  void runUtilityAction("Windows arranged (horizontal).", () =>
                    arrangeWorkbookWindows("Horizontal")
                  )
                }
              >
                Horizontal
              </button>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() =>
                  void runUtilityAction("Windows arranged (vertical).", () =>
                    arrangeWorkbookWindows("Vertical")
                  )
                }
              >
                Vertical
              </button>
              <button
                className={s.utilityBtn}
                type="button"
                onClick={() =>
                  void runUtilityAction("Windows arranged (cascade).", () =>
                    arrangeWorkbookWindows("Cascade")
                  )
                }
              >
                Cascade
              </button>
            </div>
          </>
        ) : null}
        {utilityMenu === "freeze" ? (
          <div className={s.utilityActionsRow}>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() => void runUtilityAction("Top row frozen.", freezeTopRow)}
            >
              <TableFreezeRow20Regular /> Freeze Top Row
            </button>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() => void runUtilityAction("First column frozen.", freezeFirstColumn)}
            >
              <TableFreezeColumn20Regular /> Freeze First Column
            </button>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() => void runUtilityAction("Panes unfrozen.", unfreezePanes)}
            >
              <TableFreezeColumnAndRow20Regular /> Unfreeze Panes
            </button>
          </div>
        ) : null}
        {utilityMenu === "data" ? (
          <div className={s.utilityActionsRow}>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() => void runUtilityAction("Pivot tables refreshed.", refreshPivotTables)}
            >
              <ArrowClockwise20Regular /> Refresh Pivots
            </button>
          </div>
        ) : null}
        {utilityMenu === "sheet" ? (
          <div className={s.utilityActionsRow}>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() => void runUtilityAction("Active sheet hidden.", hideActiveSheet)}
            >
              <SlideHide20Regular /> Hide Sheet
            </button>
            <button
              className={s.utilityBtn}
              type="button"
              onClick={() =>
                void runUtilityAction("Active sheet set to very hidden.", hideActiveSheetVeryHidden)
              }
            >
              <EyeOff20Regular /> Hide Very Hidden
            </button>
          </div>
        ) : null}
        {utilityMenu === "sizing" ? (
          <>
            <Text className={s.label}>Selection or Sheet Sizing</Text>
            <Text className={s.muted}>
              If only one cell is selected, sizing applies to the entire active sheet.
            </Text>
            <div className={s.utilitySizeGrid}>
              <div className={s.utilitySizeRow}>
                <input
                  className={s.utilitySizeInput}
                  type="number"
                  aria-label="Column width"
                  min={1}
                  step={0.5}
                  value={columnWidthInput}
                  onChange={(event) => setColumnWidthInput(event.target.value)}
                  placeholder="Column width"
                />
                <button
                  className={s.utilityBtn}
                  type="button"
                  onClick={applyColumnWidth}
                  title="Apply column width"
                  aria-label="Apply column width"
                >
                  <TableResizeColumn20Regular />
                </button>
              </div>
              <div className={s.utilitySizeRow}>
                <input
                  className={s.utilitySizeInput}
                  type="number"
                  aria-label="Row height"
                  min={1}
                  step={0.5}
                  value={rowHeightInput}
                  onChange={(event) => setRowHeightInput(event.target.value)}
                  placeholder="Row height"
                />
                <button
                  className={s.utilityBtn}
                  type="button"
                  onClick={applyRowHeight}
                  title="Apply row height"
                  aria-label="Apply row height"
                >
                  <TableResizeRow20Regular />
                </button>
              </div>
            </div>
          </>
        ) : null}
      </div>
    </div>
  );

  return (
    <div className={s.root}>
      {!embedded ? (
        <>
          <div className={s.workspaceHeader}>
            <div className={s.workspaceHeaderTop}>
              <Text className={s.workspaceTitle}>Layout</Text>
              <Button
                size="small"
                appearance="primary"
                onClick={() => void applySelectedFormatActive()}
                disabled={!selectedFormatId}
              >
                Apply
              </Button>
            </div>
            <Text className={s.workspaceSub}>
              Apply workbook presentation and usability improvements.
            </Text>
          </div>

          <div className={s.workspaceActionBar}>
            <Button
              size="small"
              appearance="secondary"
              onClick={() => openLayoutSection("gridFreeze")}
            >
              Gridlines and Freeze
            </Button>
            {renderShapeBuilderActions("toolbar")}
            <Button
              size="small"
              appearance="secondary"
              onClick={() => openLayoutSection("quickFormatting")}
            >
              Quick Formatting
            </Button>
            <Button size="small" onClick={onOpenLegacy}>
              Legacy
            </Button>
            <Text className={s.workspaceActionHint}>{status || "Sheet tools ready."}</Text>
          </div>
        </>
      ) : null}

      <div className={s.sectionStack}>
        <div className={s.sectionCard}>
          <button
            id="format-grid-freeze-trigger"
            className={s.sectionHeaderBtn}
            type="button"
            onClick={() => toggleLayoutSection("gridFreeze")}
            aria-expanded={expandedSections.includes("gridFreeze")}
            aria-controls="format-grid-freeze-panel"
          >
            <span className={s.sectionHeaderMain}>
              <Grid20Regular />
              <span className={s.sectionTitle}>Gridlines and Freeze Panes</span>
            </span>
            <span className={s.sectionMeta}>
              <span className={s.sectionChip}>Active</span>
              {expandedSections.includes("gridFreeze") ? (
                <ChevronUp20Regular />
              ) : (
                <ChevronDown20Regular />
              )}
            </span>
          </button>
          {expandedSections.includes("gridFreeze") ? (
            <div
              id="format-grid-freeze-panel"
              className={s.sectionBody}
              role="region"
              aria-labelledby="format-grid-freeze-trigger"
            >
              {formattingUtilitiesCard}
            </div>
          ) : null}
        </div>

        <div className={s.sectionCard}>
          <button
            id="format-sheet-presentation-trigger"
            className={s.sectionHeaderBtn}
            type="button"
            onClick={() => toggleLayoutSection("sheetPresentation")}
            aria-expanded={expandedSections.includes("sheetPresentation")}
            aria-controls="format-sheet-presentation-panel"
          >
            <span className={s.sectionHeaderMain}>
              <Emoji20Regular />
              <span className={s.sectionTitle}>Sheet Presentation</span>
            </span>
            <span className={s.sectionMeta}>
              <span className={s.sectionCount}>
                {isShapeListFiltered
                  ? `${filteredShapes.length}/${shapes.length}`
                  : `${shapes.length}`}
              </span>
              {expandedSections.includes("sheetPresentation") ? (
                <ChevronUp20Regular />
              ) : (
                <ChevronDown20Regular />
              )}
            </span>
          </button>
          {expandedSections.includes("sheetPresentation") ? (
            <div
              id="format-sheet-presentation-panel"
              className={s.sectionBody}
              role="region"
              aria-labelledby="format-sheet-presentation-trigger"
            >
              <div className={s.card}>
                <div className={s.row}>
                  <Text className={s.title}>Shape Builder</Text>
                  <Text className={`${s.status} ${statusType === "error" ? s.error : s.ok}`}>
                    {status || "Ready."}
                  </Text>
                </div>
              </div>

              <div className={s.card}>
                <div className={s.row}>
                  <Text className={s.title}>Shape List</Text>
                  <div className={s.horizontal}>
                    <Text className={s.muted}>
                      {isShapeListFiltered
                        ? `${filteredShapes.length} of ${shapes.length}`
                        : `${shapes.length}`}{" "}
                      on active sheet
                    </Text>
                    {renderShapeBuilderActions("card")}
                  </div>
                </div>
                <div className={s.grid2}>
                  <Input
                    aria-label="Filter shapes"
                    value={shapeFilterQuery}
                    placeholder="Filter by name, text, type, icon..."
                    onChange={(_, data) => setShapeFilterQuery(data.value)}
                  />
                  <select
                    className={s.select}
                    aria-label="Filter shapes by link type"
                    value={shapeLinkFilter}
                    onChange={(event) => setShapeLinkFilter(event.target.value as ShapeLinkFilter)}
                  >
                    <option value="all">All links</option>
                    <option value="none">No link</option>
                    <option value="internal">Internal link</option>
                    <option value="external">External link</option>
                  </select>
                </div>
                <div className={s.shapeList}>
                  {!shapes.length ? (
                    <div className={s.empty}>No shapes yet. Click Start Shape Builder.</div>
                  ) : null}
                  {shapes.length && !filteredShapes.length ? (
                    <div className={s.empty}>No shapes match the current filters.</div>
                  ) : null}
                  {filteredShapes.map((shape) => {
                    const selected = shape.id === selectedShapeId;
                    const isConfirmingDelete = confirmDeleteShapeId === shape.id;
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
                        <div className={s.preview}>
                          <ShapePreviewGraphic className={s.previewSvg} shape={shape} />
                        </div>
                        <div>
                          <Text className={s.itemName}>{shape.text || shape.shapeName}</Text>
                          <Text className={s.itemSub}>
                            {shape.shapeType.replace(/([A-Z])/g, " $1").trim()} • layer{" "}
                            {shape.zOrderPosition}
                          </Text>
                          <div className={s.itemTag}>{shape.linkType.toLowerCase()}</div>
                        </div>
                        <div className={s.actionCol}>
                          {isConfirmingDelete ? (
                            <>
                              <button
                                className={`${s.smallBtn} ${s.dangerBtn}`}
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  void deleteShape(shape);
                                }}
                                disabled={busy}
                              >
                                Confirm
                              </button>
                              <button
                                className={s.smallBtn}
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  setConfirmDeleteShapeId("");
                                }}
                                disabled={busy}
                              >
                                Cancel
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                className={s.smallBtn}
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  setConfirmDeleteShapeId("");
                                  editShape(shape);
                                }}
                                disabled={busy}
                              >
                                Edit
                              </button>
                              <button
                                className={s.smallBtn}
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  setConfirmDeleteShapeId("");
                                  void duplicateShape(shape);
                                }}
                                disabled={busy}
                              >
                                Copy
                              </button>
                              <button
                                className={`${s.smallBtn} ${s.dangerBtn}`}
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  setConfirmDeleteShapeId(shape.id);
                                }}
                                disabled={busy}
                              >
                                Delete
                              </button>
                            </>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className={s.card} ref={editorCardRef}>
                <div className={s.row}>
                  <Text className={s.title}>Editor</Text>
                  <Text className={s.muted}>
                    {selectedShape ? selectedShape.shapeName : "Select a shape"}
                  </Text>
                </div>
                {!draft ? (
                  <div className={s.empty}>Select a shape to edit.</div>
                ) : (
                  <>
                    <div className={s.tabs}>
                      {(["shape", "style", "size", "bind", "link"] as EditorTab[]).map((item) => (
                        <button
                          key={item}
                          className={`${s.tab} ${tab === item ? s.tabActive : ""}`}
                          type="button"
                          onClick={() => setTab(item)}
                        >
                          {item === "shape"
                            ? "Shape"
                            : item === "style"
                              ? "Style"
                              : item === "size"
                                ? "Size"
                                : item === "bind"
                                  ? "Bind"
                                  : "Link"}
                        </button>
                      ))}
                    </div>

                    {tab === "shape" ? (
                      <div className={s.panel}>
                        <Text className={s.label}>Shape Name</Text>
                        <div className={s.horizontal}>
                          <Input
                            aria-label="Shape name"
                            value={draft.shapeName}
                            onChange={(_, data) => setDraft({ ...draft, shapeName: data.value })}
                          />
                          <button
                            className={s.smallBtn}
                            type="button"
                            onClick={() => void renameSelectedShape()}
                          >
                            Rename
                          </button>
                        </div>
                        <Text className={s.muted}>
                          {draft.anchorAddress
                            ? `Anchor cell: ${draft.sheetName}!${draft.anchorAddress}`
                            : "No anchor cell yet. Use Insert On Active Cell for binding workflows."}
                        </Text>
                        <Text className={s.label}>Shape Type</Text>
                        <select
                          className={s.select}
                          aria-label="Shape type"
                          value={draft.shapeType}
                          onChange={(event) => {
                            const shapeType = event.target.value as InsertableShapeType;
                            setDraft({ ...draft, shapeType });
                            void patchShape({ shapeType });
                          }}
                        >
                          {SHAPE_TYPES.map((option) => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                        <Text className={s.label}>Text Editor</Text>
                        <div className={s.textEditorCard}>
                          <div className={s.textToolbar}>
                            <button
                              className={`${s.iconOnlyBtn} ${draft.bold ? s.iconOnlyBtnActive : ""}`}
                              type="button"
                              onClick={() => void applyTextSelectionFormat({ bold: !draft.bold })}
                              title="Bold selected text"
                              aria-label="Bold selected text"
                            >
                              <TextBold20Regular />
                            </button>
                            <button
                              className={`${s.iconOnlyBtn} ${draft.italic ? s.iconOnlyBtnActive : ""}`}
                              type="button"
                              onClick={() =>
                                void applyTextSelectionFormat({ italic: !draft.italic })
                              }
                              title="Italic selected text"
                              aria-label="Italic selected text"
                            >
                              <TextItalic20Regular />
                            </button>
                            <div className={s.textToolbarDivider} />
                            <button
                              className={`${s.iconOnlyBtn} ${
                                draft.textHorizontalAlignment === "Left" ? s.iconOnlyBtnActive : ""
                              }`}
                              type="button"
                              onClick={() => {
                                const textHorizontalAlignment = "Left";
                                setDraft({ ...draft, textHorizontalAlignment });
                                void patchShape({ textHorizontalAlignment });
                              }}
                              title="Align left"
                              aria-label="Align left"
                            >
                              <TextAlignLeft20Regular />
                            </button>
                            <button
                              className={`${s.iconOnlyBtn} ${
                                draft.textHorizontalAlignment === "Center"
                                  ? s.iconOnlyBtnActive
                                  : ""
                              }`}
                              type="button"
                              onClick={() => {
                                const textHorizontalAlignment = "Center";
                                setDraft({ ...draft, textHorizontalAlignment });
                                void patchShape({ textHorizontalAlignment });
                              }}
                              title="Align center"
                              aria-label="Align center"
                            >
                              <TextAlignCenter20Regular />
                            </button>
                            <button
                              className={`${s.iconOnlyBtn} ${
                                draft.textHorizontalAlignment === "Right" ? s.iconOnlyBtnActive : ""
                              }`}
                              type="button"
                              onClick={() => {
                                const textHorizontalAlignment = "Right";
                                setDraft({ ...draft, textHorizontalAlignment });
                                void patchShape({ textHorizontalAlignment });
                              }}
                              title="Align right"
                              aria-label="Align right"
                            >
                              <TextAlignRight20Regular />
                            </button>
                            <select
                              className={s.compactSelect}
                              aria-label="Text vertical alignment"
                              value={draft.textVerticalAlignment}
                              onChange={(event) => {
                                const textVerticalAlignment = event.target
                                  .value as ShapeBuilderShapeRecord["textVerticalAlignment"];
                                setDraft({ ...draft, textVerticalAlignment });
                                void patchShape({ textVerticalAlignment });
                              }}
                              title="Vertical alignment"
                            >
                              <option value="Top">Top</option>
                              <option value="Middle">Middle</option>
                              <option value="Bottom">Bottom</option>
                            </select>
                            <input
                              className={s.compactColorInput}
                              type="color"
                              aria-label="Text color"
                              value={toUpperHexOrFallback(draft.fontColor, "#1F2937")}
                              onChange={(event) => applyFontColor(event.target.value)}
                              title="Text color"
                            />
                            <button
                              className={s.iconOnlyBtn}
                              type="button"
                              onClick={() => void pickColorFromScreen("font")}
                              title="Eyedropper Font"
                              aria-label="Eyedropper Font"
                            >
                              <Eyedropper20Regular />
                            </button>
                            <button
                              className={s.iconOnlyBtn}
                              type="button"
                              onClick={() => {
                                textAreaRef.current?.focus();
                                setEmojiPickerOpen((prev) => !prev);
                              }}
                              title="Open emoji picker"
                              aria-label="Open emoji picker"
                            >
                              <Emoji20Regular />
                            </button>
                            <input
                              className={`${s.input} ${s.compactNumberInput}`}
                              type="number"
                              aria-label="Font size"
                              min={8}
                              max={72}
                              value={Math.round(draft.fontSize)}
                              onChange={(event) => {
                                const fontSize = clamp(
                                  Number(event.target.value) || draft.fontSize,
                                  8,
                                  72
                                );
                                setDraft({ ...draft, fontSize });
                              }}
                              onBlur={() => void patchShape({ fontSize: draft.fontSize })}
                              title="Font size"
                            />
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => setTextEditorExpanded((prev) => !prev)}
                            >
                              {textEditorExpanded ? "Compact" : "Expand"}
                            </button>
                          </div>
                          <Text className={s.textToolbarHint}>
                            Select text first for Bold/Italic. Use the emoji picker for quick
                            inserts, or paste any symbol directly into the editor.
                          </Text>
                          {emojiPickerOpen ? (
                            <div className={s.emojiPicker}>
                              {COMMON_SHAPE_EMOJIS.map((emoji) => (
                                <button
                                  key={emoji}
                                  className={s.emojiBtn}
                                  type="button"
                                  onClick={() => insertEmojiIntoText(emoji)}
                                  title={`Insert ${emoji}`}
                                  aria-label={`Insert ${emoji}`}
                                >
                                  {emoji}
                                </button>
                              ))}
                            </div>
                          ) : null}
                          <textarea
                            ref={textAreaRef}
                            className={`${s.textArea} ${textEditorExpanded ? s.textAreaExpanded : ""} ${
                              draft.textHorizontalAlignment === "Center"
                                ? s.textAreaAlignCenter
                                : draft.textHorizontalAlignment === "Right"
                                  ? s.textAreaAlignRight
                                  : s.textAreaAlignLeft
                            }`}
                            aria-label="Shape text"
                            value={draft.text}
                            onChange={(event) => setDraft({ ...draft, text: event.target.value })}
                            onBlur={() => void patchShape({ text: draft.text })}
                          />
                        </div>
                        <Text className={s.label}>Icon</Text>
                        <Input
                          aria-label="Search icons"
                          value={iconQuery}
                          placeholder="Search icons or Fluent names..."
                          onChange={(_, data) => setIconQuery(data.value)}
                        />
                        <div className={s.grid2}>
                          <Input
                            aria-label="Custom icon"
                            value={customIconInput}
                            placeholder="Paste custom icon/emoji (Windows+V)"
                            onChange={(_, data) => setCustomIconInput(data.value)}
                          />
                          <button className={s.smallBtn} type="button" onClick={applyCustomIcon}>
                            Apply Custom Icon
                          </button>
                        </div>
                        <div className={s.horizontal}>
                          <button
                            className={s.smallBtn}
                            type="button"
                            onClick={() => {
                              setCustomIconInput("");
                              setDraft({ ...draft, iconKey: "none" });
                              void patchShape({ iconKey: "none" });
                            }}
                          >
                            No Icon
                          </button>
                        </div>
                        <Text className={s.iconCount}>
                          {filteredIcons.length} icon{filteredIcons.length === 1 ? "" : "s"} shown •{" "}
                          {FLUENT_ICON_NAME_CATALOG.length.toLocaleString()} Fluent names indexed
                          {iconQuery.trim().length > 0
                            ? ` • ${fluentIconMatches.length.toLocaleString()} Fluent matches`
                            : ""}
                        </Text>
                        <div className={s.iconGrid}>
                          {filteredIcons.map((icon) => (
                            <button
                              key={icon.key}
                              className={`${s.iconBtn} ${draft.iconKey === icon.key ? s.iconActive : ""}`}
                              type="button"
                              onClick={() => {
                                setCustomIconInput("");
                                setDraft({ ...draft, iconKey: icon.key });
                                void patchShape({ iconKey: icon.key });
                              }}
                            >
                              <span>{icon.symbol || "-"}</span>
                              <span>{icon.label}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    ) : null}

                    {tab === "style" ? (
                      <div className={s.panel}>
                        <div className={s.horizontal}>
                          <button
                            className={s.smallBtn}
                            type="button"
                            onClick={revertStyleChanges}
                            disabled={!styleBaseline}
                          >
                            Revert Style
                          </button>
                        </div>
                        <Text className={s.muted}>
                          Theme swatches are pulled from this workbook's active theme.
                        </Text>
                        <Text className={s.label}>Fill</Text>
                        <div className={s.colorCompactGroup}>
                          <div className={s.colorCompactHeader}>
                            <button
                              className={s.colorCurrentBtn}
                              type="button"
                              onClick={() =>
                                setOpenColorTarget((prev) => (prev === "fill" ? null : "fill"))
                              }
                            >
                              <span className={s.colorCurrentMain}>
                                <span className={s.colorChip}>
                                  <ColorSwatchGraphic
                                    className={s.colorGraphic}
                                    color={fillDisplayColor}
                                    noFill={fillDisplayColor === NO_FILL_COLOR_TOKEN}
                                  />
                                </span>
                                <span className={s.colorCurrentLabel}>
                                  {fillDisplayColor === NO_FILL_COLOR_TOKEN
                                    ? "No Fill"
                                    : fillDisplayColor}
                                </span>
                              </span>
                              {openColorTarget === "fill" ? (
                                <ChevronUp20Regular />
                              ) : (
                                <ChevronDown20Regular />
                              )}
                            </button>
                            <button
                              className={s.iconOnlyBtn}
                              type="button"
                              onClick={() => void pickColorFromScreen("fill")}
                              title="Eyedropper Fill"
                              aria-label="Eyedropper Fill"
                            >
                              <Eyedropper20Regular />
                            </button>
                          </div>
                          {openColorTarget === "fill" ? (
                            <div className={s.colorCompactBody}>
                              <div className={s.swatches}>
                                <button
                                  type="button"
                                  className={`${s.noFillSwatch} ${draft.fillColor === NO_FILL_COLOR_TOKEN ? s.swatchActive : ""}`}
                                  onClick={() => applyFillColor(NO_FILL_COLOR_TOKEN)}
                                >
                                  No Fill
                                </button>
                                {themeColors.map((color) => (
                                  <button
                                    key={`f-${color}`}
                                    className={`${s.swatch} ${
                                      draft.fillColor !== NO_FILL_COLOR_TOKEN &&
                                      draft.fillColor.toLowerCase() === color.toLowerCase()
                                        ? s.swatchActive
                                        : ""
                                    }`}
                                    type="button"
                                    aria-label={`Apply fill color ${color}`}
                                    title={`Apply fill color ${color}`}
                                    onClick={() => applyFillColor(color)}
                                  >
                                    <ColorSwatchGraphic className={s.swatchGraphic} color={color} />
                                  </button>
                                ))}
                              </div>
                              <input
                                className={s.colorInput}
                                type="color"
                                aria-label="Custom fill color"
                                value={
                                  draft.fillColor === NO_FILL_COLOR_TOKEN
                                    ? "#FFFFFF"
                                    : toUpperHexOrFallback(draft.fillColor, "#A8D5AD")
                                }
                                onChange={(event) => applyFillColor(event.target.value)}
                              />
                            </div>
                          ) : null}
                        </div>
                        <Text className={s.label}>Outline</Text>
                        <div className={s.colorCompactGroup}>
                          <div className={s.colorCompactHeader}>
                            <button
                              className={s.colorCurrentBtn}
                              type="button"
                              onClick={() =>
                                setOpenColorTarget((prev) =>
                                  prev === "outline" ? null : "outline"
                                )
                              }
                            >
                              <span className={s.colorCurrentMain}>
                                <span className={s.colorChip}>
                                  <ColorSwatchGraphic
                                    className={s.colorGraphic}
                                    color={outlineDisplayColor}
                                  />
                                </span>
                                <span className={s.colorCurrentLabel}>{outlineDisplayColor}</span>
                              </span>
                              {openColorTarget === "outline" ? (
                                <ChevronUp20Regular />
                              ) : (
                                <ChevronDown20Regular />
                              )}
                            </button>
                            <button
                              className={s.iconOnlyBtn}
                              type="button"
                              onClick={() => void pickColorFromScreen("outline")}
                              title="Eyedropper Outline"
                              aria-label="Eyedropper Outline"
                            >
                              <Eyedropper20Regular />
                            </button>
                          </div>
                          {openColorTarget === "outline" ? (
                            <div className={s.colorCompactBody}>
                              <div className={s.swatches}>
                                {themeColors.map((color) => (
                                  <button
                                    key={`o-${color}`}
                                    className={`${s.swatch} ${
                                      draft.outlineColor.toLowerCase() === color.toLowerCase()
                                        ? s.swatchActive
                                        : ""
                                    }`}
                                    type="button"
                                    aria-label={`Apply outline color ${color}`}
                                    title={`Apply outline color ${color}`}
                                    onClick={() => applyOutlineColor(color)}
                                  >
                                    <ColorSwatchGraphic className={s.swatchGraphic} color={color} />
                                  </button>
                                ))}
                              </div>
                              <input
                                className={s.colorInput}
                                type="color"
                                aria-label="Custom outline color"
                                value={toUpperHexOrFallback(draft.outlineColor, "#8CBF95")}
                                onChange={(event) => applyOutlineColor(event.target.value)}
                              />
                            </div>
                          ) : null}
                        </div>
                        <Text className={s.label}>Font</Text>
                        <div className={s.colorCompactGroup}>
                          <div className={s.colorCompactHeader}>
                            <button
                              className={s.colorCurrentBtn}
                              type="button"
                              onClick={() =>
                                setOpenColorTarget((prev) => (prev === "font" ? null : "font"))
                              }
                            >
                              <span className={s.colorCurrentMain}>
                                <span className={s.colorChip}>
                                  <ColorSwatchGraphic
                                    className={s.colorGraphic}
                                    color={fontDisplayColor}
                                  />
                                </span>
                                <span className={s.colorCurrentLabel}>{fontDisplayColor}</span>
                              </span>
                              {openColorTarget === "font" ? (
                                <ChevronUp20Regular />
                              ) : (
                                <ChevronDown20Regular />
                              )}
                            </button>
                            <button
                              className={s.iconOnlyBtn}
                              type="button"
                              onClick={() => void pickColorFromScreen("font")}
                              title="Eyedropper Font"
                              aria-label="Eyedropper Font"
                            >
                              <Eyedropper20Regular />
                            </button>
                          </div>
                          {openColorTarget === "font" ? (
                            <div className={s.colorCompactBody}>
                              <div className={s.swatches}>
                                {themeColors.map((color) => (
                                  <button
                                    key={`font-${color}`}
                                    className={`${s.swatch} ${
                                      draft.fontColor.toLowerCase() === color.toLowerCase()
                                        ? s.swatchActive
                                        : ""
                                    }`}
                                    type="button"
                                    aria-label={`Apply font color ${color}`}
                                    title={`Apply font color ${color}`}
                                    onClick={() => applyFontColor(color)}
                                  >
                                    <ColorSwatchGraphic className={s.swatchGraphic} color={color} />
                                  </button>
                                ))}
                              </div>
                              <input
                                className={s.colorInput}
                                type="color"
                                aria-label="Custom font color"
                                value={toUpperHexOrFallback(draft.fontColor, "#1F2937")}
                                onChange={(event) => applyFontColor(event.target.value)}
                              />
                            </div>
                          ) : null}
                        </div>
                        <div className={s.grid2}>
                          <div>
                            <Text className={s.label}>Font Size</Text>
                            <input
                              className={s.input}
                              type="number"
                              aria-label="Font size"
                              min={8}
                              max={72}
                              value={Math.round(draft.fontSize)}
                              onChange={(event) => {
                                const next = clamp(
                                  Number(event.target.value) || draft.fontSize,
                                  8,
                                  72
                                );
                                setDraft({ ...draft, fontSize: next });
                              }}
                              onBlur={() => void patchShape({ fontSize: draft.fontSize })}
                            />
                          </div>
                          <div>
                            <Text className={s.label}>Font Style</Text>
                            <div className={s.toggleGroup}>
                              <button
                                className={`${s.smallBtn} ${draft.bold ? s.toggleBtnActive : ""}`}
                                type="button"
                                onClick={() => {
                                  const bold = !draft.bold;
                                  setDraft({ ...draft, bold });
                                  void patchShape({ bold });
                                }}
                              >
                                Bold
                              </button>
                              <button
                                className={`${s.smallBtn} ${draft.italic ? s.toggleBtnActive : ""}`}
                                type="button"
                                onClick={() => {
                                  const italic = !draft.italic;
                                  setDraft({ ...draft, italic });
                                  void patchShape({ italic });
                                }}
                              >
                                Italic
                              </button>
                            </div>
                          </div>
                        </div>
                        <Text className={s.label}>
                          Outline Width: {Math.round(draft.outlineWidth)}px
                        </Text>
                        <input
                          type="range"
                          aria-label="Outline width"
                          min={0}
                          max={12}
                          step={1}
                          value={Math.round(draft.outlineWidth)}
                          onChange={(event) =>
                            setDraft({ ...draft, outlineWidth: Number(event.target.value) })
                          }
                          onMouseUp={() => void patchShape({ outlineWidth: draft.outlineWidth })}
                          onTouchEnd={() => void patchShape({ outlineWidth: draft.outlineWidth })}
                        />
                        <Text className={s.label}>Effect</Text>
                        <div className={s.helperCard}>
                          <Text className={s.muted}>
                            Excel shape effect presets are not exposed by Office.js. Use Excel
                            Ribbon: Shape Effects.
                          </Text>
                          <button
                            className={s.smallBtn}
                            type="button"
                            onClick={() => {
                              if (draft.effect === "None") {
                                return;
                              }
                              setDraft({ ...draft, effect: "None" });
                              void patchShape({ effect: "None" as ShapeBuilderEffect });
                            }}
                            disabled={draft.effect === "None"}
                          >
                            Clear Effect Metadata
                          </button>
                        </div>
                      </div>
                    ) : null}

                    {tab === "size" ? (
                      <div className={s.panel}>
                        <div className={s.grid2}>
                          <div>
                            <Text className={s.label}>Width</Text>
                            <input
                              className={s.input}
                              type="number"
                              aria-label="Shape width"
                              value={Math.round(draft.width)}
                              onChange={(event) =>
                                setDraft({
                                  ...draft,
                                  width: clamp(Number(event.target.value) || draft.width, 80, 1200),
                                })
                              }
                              onBlur={() => void patchShape({ width: draft.width })}
                            />
                          </div>
                          <div>
                            <Text className={s.label}>Height</Text>
                            <input
                              className={s.input}
                              type="number"
                              aria-label="Shape height"
                              value={Math.round(draft.height)}
                              onChange={(event) =>
                                setDraft({
                                  ...draft,
                                  height: clamp(
                                    Number(event.target.value) || draft.height,
                                    24,
                                    700
                                  ),
                                })
                              }
                              onBlur={() => void patchShape({ height: draft.height })}
                            />
                          </div>
                        </div>
                        <Text className={s.label}>Width slider</Text>
                        <input
                          type="range"
                          aria-label="Shape width slider"
                          min={80}
                          max={500}
                          value={Math.round(draft.width)}
                          onChange={(event) =>
                            setDraft({ ...draft, width: Number(event.target.value) })
                          }
                          onMouseUp={() => void patchShape({ width: draft.width })}
                          onTouchEnd={() => void patchShape({ width: draft.width })}
                        />
                        <Text className={s.label}>Height slider</Text>
                        <input
                          type="range"
                          aria-label="Shape height slider"
                          min={24}
                          max={220}
                          value={Math.round(draft.height)}
                          onChange={(event) =>
                            setDraft({ ...draft, height: Number(event.target.value) })
                          }
                          onMouseUp={() => void patchShape({ height: draft.height })}
                          onTouchEnd={() => void patchShape({ height: draft.height })}
                        />
                        <div className={s.grid2}>
                          <div>
                            <Text className={s.label}>Position X</Text>
                            <input
                              className={s.input}
                              type="number"
                              aria-label="Shape position X"
                              value={Math.round(draft.left)}
                              onChange={(event) =>
                                setDraft({
                                  ...draft,
                                  left: clamp(Number(event.target.value) || draft.left, 0, 5000),
                                })
                              }
                              onBlur={() => void patchShape({ left: draft.left })}
                            />
                          </div>
                          <div>
                            <Text className={s.label}>Position Y</Text>
                            <input
                              className={s.input}
                              type="number"
                              aria-label="Shape position Y"
                              value={Math.round(draft.top)}
                              onChange={(event) =>
                                setDraft({
                                  ...draft,
                                  top: clamp(Number(event.target.value) || draft.top, 0, 5000),
                                })
                              }
                              onBlur={() => void patchShape({ top: draft.top })}
                            />
                          </div>
                        </div>
                        <div className={s.helperCard}>
                          <Text className={s.label}>Anchor and Arrange</Text>
                          <Text className={s.muted}>
                            {draft.anchorAddress
                              ? `Anchor cell: ${draft.sheetName}!${draft.anchorAddress}`
                              : "This shape is not anchored to a worksheet cell yet."}
                          </Text>
                          <div className={s.utilityActionsRow}>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void moveSelectedShapeToSelection()}
                              disabled={!draft.anchorAddress}
                            >
                              Move To Selection
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void nudgeSelectedShape(-4, 0)}
                            >
                              Nudge Left
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void nudgeSelectedShape(4, 0)}
                            >
                              Nudge Right
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void nudgeSelectedShape(0, -4)}
                            >
                              Nudge Up
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void nudgeSelectedShape(0, 4)}
                            >
                              Nudge Down
                            </button>
                          </div>
                          <div className={s.utilityActionsRow}>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Left")}
                            >
                              Align Left
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Center")}
                            >
                              Align Center
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Right")}
                            >
                              Align Right
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Top")}
                            >
                              Align Top
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Middle")}
                            >
                              Align Middle
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void alignSelectedShapeToSelection("Bottom")}
                            >
                              Align Bottom
                            </button>
                          </div>
                          <div className={s.utilityActionsRow}>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void updateSelectedShapeOrder("BringToFront")}
                            >
                              Bring To Front
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void updateSelectedShapeOrder("BringForward")}
                            >
                              Bring Forward
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void updateSelectedShapeOrder("SendBackward")}
                            >
                              Send Backward
                            </button>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void updateSelectedShapeOrder("SendToBack")}
                            >
                              Send To Back
                            </button>
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {tab === "bind" ? (
                      <div className={s.panel}>
                        <div className={s.helperCard}>
                          <Text className={s.label}>Named Range Value</Text>
                          <Text className={s.muted}>
                            Pull the current value from a named range into this shape's text.
                          </Text>
                          <div className={s.horizontal}>
                            <Input
                              aria-label="Named range for shape text"
                              value={namedRangeBindingInput}
                              placeholder="RevenueLabel"
                              onChange={(_, data) => setNamedRangeBindingInput(data.value)}
                            />
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void applyNamedRangeBinding()}
                            >
                              Apply Value
                            </button>
                          </div>
                        </div>
                        <div className={s.helperCard}>
                          <Text className={s.label}>Anchor Cell Formula</Text>
                          <Text className={s.muted}>
                            Write a formula into the shape's anchor cell and mirror the result into
                            the shape text.
                          </Text>
                          <textarea
                            className={s.textArea}
                            aria-label="Anchor cell formula"
                            value={anchorFormulaInput}
                            placeholder='=TEXT(TODAY(),"yyyy-mm-dd")'
                            onChange={(event) => setAnchorFormulaInput(event.target.value)}
                          />
                          <div className={s.horizontal}>
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() => void applyAnchorFormula()}
                              disabled={!draft.anchorAddress}
                            >
                              Apply Formula
                            </button>
                            {!draft.anchorAddress ? (
                              <Text className={s.muted}>
                                Use Insert On Active Cell to create an anchored shape first.
                              </Text>
                            ) : null}
                          </div>
                        </div>
                      </div>
                    ) : null}

                    {tab === "link" ? (
                      <div className={s.panel}>
                        <Text className={s.label}>Link Type</Text>
                        <div className={s.radioRow}>
                          <label>
                            <input
                              type="radio"
                              checked={draft.linkType === "None"}
                              onChange={() => {
                                setDraft({
                                  ...draft,
                                  linkType: "None",
                                  internalDestinationId: "",
                                  externalUrl: "",
                                });
                                void patchShape({
                                  linkType: "None",
                                  internalDestinationId: "",
                                  externalUrl: "",
                                });
                              }}
                            />{" "}
                            No Link
                          </label>
                          <label>
                            <input
                              type="radio"
                              checked={draft.linkType === "Internal"}
                              onChange={() => {
                                const target =
                                  draft.internalDestinationId || destinations[0]?.id || "";
                                setDraft({
                                  ...draft,
                                  linkType: "Internal",
                                  internalDestinationId: target,
                                  externalUrl: "",
                                });
                                void patchShape({
                                  linkType: "Internal",
                                  internalDestinationId: target,
                                  externalUrl: "",
                                });
                              }}
                            />{" "}
                            Internal (Sheet / Cell)
                          </label>
                          <label>
                            <input
                              type="radio"
                              checked={draft.linkType === "External"}
                              onChange={() => {
                                setDraft({
                                  ...draft,
                                  linkType: "External",
                                  internalDestinationId: "",
                                });
                                void patchShape({
                                  linkType: "External",
                                  internalDestinationId: "",
                                });
                              }}
                            />{" "}
                            External URL
                          </label>
                        </div>
                        {draft.linkType === "Internal" ? (
                          <>
                            <Text className={s.label}>Destination</Text>
                            <select
                              className={s.select}
                              aria-label="Internal destination"
                              value={draft.internalDestinationId}
                              onChange={(event) => {
                                const internalDestinationId = event.target.value;
                                setDraft({ ...draft, internalDestinationId });
                                void patchShape({
                                  linkType: "Internal",
                                  internalDestinationId,
                                  externalUrl: "",
                                });
                              }}
                            >
                              {destinations.map((option) => (
                                <option key={option.id} value={option.id}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                            <Text className={s.label}>Cell Reference (e.g. Sheet2!A1)</Text>
                            <div className={s.horizontal}>
                              <input
                                className={s.input}
                                aria-label="Cell reference"
                                value={cellRefInput}
                                onChange={(event) => setCellRefInput(event.target.value)}
                              />
                              <button
                                className={s.smallBtn}
                                type="button"
                                onClick={() => {
                                  const id = parseCellDestination(cellRefInput);
                                  if (!id) {
                                    setErr("Invalid cell reference.");
                                    return;
                                  }
                                  setDraft({ ...draft, internalDestinationId: id });
                                  void patchShape({
                                    linkType: "Internal",
                                    internalDestinationId: id,
                                    externalUrl: "",
                                  });
                                  setOk("Cell reference linked.");
                                }}
                              >
                                Use Cell
                              </button>
                              <button
                                className={s.smallBtn}
                                type="button"
                                onClick={() =>
                                  void (async () => {
                                    const destinationId = draft.internalDestinationId.trim();
                                    if (!destinationId) {
                                      setErr("Select an internal destination first.");
                                      return;
                                    }
                                    try {
                                      await activateNavigationDestination(destinationId);
                                    } catch (error) {
                                      setErr(error);
                                    }
                                  })()
                                }
                              >
                                Test
                              </button>
                            </div>
                          </>
                        ) : null}
                        {draft.linkType === "External" ? (
                          <>
                            <Text className={s.label}>External URL</Text>
                            <input
                              className={s.input}
                              aria-label="External URL"
                              value={draft.externalUrl}
                              onChange={(event) =>
                                setDraft({ ...draft, externalUrl: event.target.value })
                              }
                              onBlur={() =>
                                void patchShape({
                                  linkType: "External",
                                  externalUrl: draft.externalUrl,
                                  internalDestinationId: "",
                                })
                              }
                            />
                            <button
                              className={s.smallBtn}
                              type="button"
                              onClick={() =>
                                void (async () => {
                                  const url = draft.externalUrl.trim();
                                  if (!url) {
                                    setErr("Enter an external URL first.");
                                    return;
                                  }
                                  try {
                                    await activateNavigationDestination(`url::${url}`);
                                  } catch (error) {
                                    setErr(error);
                                  }
                                })()
                              }
                            >
                              Open URL
                            </button>
                          </>
                        ) : null}
                      </div>
                    ) : null}
                  </>
                )}
              </div>
            </div>
          ) : null}
        </div>

        <div className={s.sectionCard}>
          <button
            id="format-quick-formatting-trigger"
            className={s.sectionHeaderBtn}
            type="button"
            onClick={() => toggleLayoutSection("quickFormatting")}
            aria-expanded={expandedSections.includes("quickFormatting")}
            aria-controls="format-quick-formatting-panel"
          >
            <span className={s.sectionHeaderMain}>
              <TextAlignCenter20Regular />
              <span className={s.sectionTitle}>Quick Formatting Tools</span>
            </span>
            <span className={s.sectionMeta}>
              <span className={s.sectionCount}>{formats.length}</span>
              {expandedSections.includes("quickFormatting") ? (
                <ChevronUp20Regular />
              ) : (
                <ChevronDown20Regular />
              )}
            </span>
          </button>
          {expandedSections.includes("quickFormatting") ? (
            <div
              id="format-quick-formatting-panel"
              className={s.sectionBody}
              role="region"
              aria-labelledby="format-quick-formatting-trigger"
            >
              <div className={s.card}>
                <div className={s.sectionHeader}>
                  <div className={s.sectionTitleWrap}>
                    <Text className={s.title}>Cell Style Presets</Text>
                    <Text className={s.muted}>
                      Apply the core workbook presentation styles described in the add-in overview.
                    </Text>
                  </div>
                </div>
                <div className={s.helperCard}>
                  <div className={s.buttonGrid2}>
                    {CELL_STYLE_PRESETS.map((preset) => (
                      <button
                        key={preset}
                        className={s.smallBtn}
                        type="button"
                        onClick={() => void applyCellStyleQuickAction(preset)}
                      >
                        {preset}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div className={s.card}>
                <div className={s.sectionHeader}>
                  <div className={s.sectionTitleWrap}>
                    <Text className={s.title}>Table and Pivot Quick Actions</Text>
                    <Text className={s.muted}>
                      Keep workbook tables and pivot outputs consistent from the layout workspace.
                    </Text>
                  </div>
                </div>
                <div className={s.helperCard}>
                  <select
                    className={s.select}
                    aria-label="Table style"
                    value={tableStyleInput}
                    onChange={(event) => setTableStyleInput(event.target.value)}
                  >
                    {TABLE_STYLES.map((style) => (
                      <option key={style} value={style}>
                        {style}
                      </option>
                    ))}
                  </select>
                  <div className={s.utilityActionsRow}>
                    <button
                      className={s.smallBtn}
                      type="button"
                      onClick={() => void applyTableStyleQuickAction()}
                    >
                      Apply Table Style
                    </button>
                    <button
                      className={s.smallBtn}
                      type="button"
                      onClick={() =>
                        void runUtilityAction("Pivot tables refreshed.", refreshPivotTables)
                      }
                    >
                      Refresh Pivot Tables
                    </button>
                  </div>
                </div>
              </div>

              <div className={s.card}>
                <div className={s.sectionHeader}>
                  <div className={s.sectionTitleWrap}>
                    <Text className={s.title}>Template Quick Start</Text>
                    <Text className={s.muted}>
                      Mirror the new Add Shape workflow with one-click nav layout presets.
                    </Text>
                  </div>
                </div>
                <div className={s.helperCard}>
                  <label className={s.radioCard}>
                    <input
                      type="checkbox"
                      checked={applyTemplateAll}
                      onChange={(event) => setApplyTemplateAll(event.target.checked)}
                    />
                    <span>Apply template to all sheets</span>
                  </label>
                  <div className={s.buttonGrid2}>
                    <button
                      className={s.smallBtn}
                      type="button"
                      onClick={() => void applyTemplate("Vertical")}
                    >
                      Apply Vertical Template
                    </button>
                    <button
                      className={s.smallBtn}
                      type="button"
                      onClick={() => void applyTemplate("Horizontal")}
                    >
                      Apply Horizontal Template
                    </button>
                  </div>
                </div>
              </div>

              <div className={s.card}>
                <div className={s.sectionHeader}>
                  <div className={s.sectionTitleWrap}>
                    <Text className={s.title}>Sheet Format Templates</Text>
                    <Text className={s.muted}>
                      Capture once, then reapply the same sheet design everywhere.
                    </Text>
                  </div>
                </div>
                <div className={s.helperCard}>
                  <Input
                    aria-label="Template name"
                    value={newFormatName}
                    placeholder="Template name"
                    onChange={(_, data) => setNewFormatName(data.value)}
                  />
                  <Button appearance="primary" onClick={() => void captureFormat()}>
                    Capture Active Sheet
                  </Button>
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
                <div className={s.buttonGrid2}>
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
                  <div
                    key={format.id}
                    className={`${s.formatRow} ${selectedFormatId === format.id ? s.formatRowSelected : ""}`}
                  >
                    <div>
                      <Text>{format.name}</Text>
                      <Text className={s.muted}>
                        Source: {format.sourceSheet} {format.sourceAddress}
                      </Text>
                    </div>
                    <button
                      className={s.smallBtn}
                      type="button"
                      onClick={() => setSelectedFormatId(format.id)}
                    >
                      {selectedFormatId === format.id ? "Selected" : "Select"}
                    </button>
                  </div>
                ))}
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
};

export default FormatView;
