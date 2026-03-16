/**
 * nameTransforms.ts
 *
 * Utilities for building, sanitizing, and validating Excel named range names.
 * Used by the Create Named Ranges from Table modal in TablesView.tsx.
 */

// ─── Types ────────────────────────────────────────────────────────────────────

export type CaseStyle = "none" | "camelCase" | "snake_case" | "SCREAM_SNAKE_CASE";

export interface ColumnCreateSpec {
  /** The original table column name — used to look up the address in Excel. */
  columnName: string;
  /** The final, pre-validated named range name to create. */
  finalName: string;
}

export interface ColumnConfig {
  id: string;
  originalName: string;
  address: string;
  /** Whether this column is selected for named range creation. */
  included: boolean;
  /**
   * If non-empty, this value is used as the base name for this column,
   * bypassing the case style and global prefix/suffix entirely.
   */
  perColumnOverride: string;
}

export interface NamePreviewResult {
  finalName: string;
  isValid: boolean;
  isDuplicate: boolean;
}

// ─── Core transform functions ─────────────────────────────────────────────────

/**
 * Split a raw string into word tokens, handling spaces, underscores,
 * hyphens, and camelCase boundaries.
 */
function tokenize(raw: string): string[] {
  return raw
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .split(/[\s_\-]+/)
    .filter(Boolean);
}

/**
 * Apply a case style to a raw string.
 *
 * Examples (input: "Sales Region"):
 *   none              → "Sales Region"  (unchanged, sanitized separately)
 *   camelCase         → "salesRegion"
 *   snake_case        → "sales_region"
 *   SCREAM_SNAKE_CASE → "SALES_REGION"
 */
export function applyCaseStyle(raw: string, style: CaseStyle): string {
  if (style === "none") return raw;

  const words = tokenize(raw);
  if (words.length === 0) return raw;

  switch (style) {
    case "camelCase":
      return words
        .map((w, i) =>
          i === 0
            ? w.toLowerCase()
            : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
        )
        .join("");

    case "snake_case":
      return words.map((w) => w.toLowerCase()).join("_");

    case "SCREAM_SNAKE_CASE":
      return words.map((w) => w.toUpperCase()).join("_");
  }
}

/**
 * Sanitize a string so it is a legal Excel named range identifier:
 *   - Replace any character that is not a letter, digit, underscore,
 *     or period with an underscore.
 *   - If the result starts with a digit, prepend an underscore.
 *   - Collapse consecutive underscores to a single underscore.
 */
export function sanitizeForNamedRange(raw: string): string {
  let cleaned = raw.replace(/[^a-zA-Z0-9_.]/g, "_");
  if (/^[0-9]/.test(cleaned)) {
    cleaned = "_" + cleaned;
  }
  cleaned = cleaned.replace(/__+/g, "_");
  return cleaned;
}

/**
 * Build the final named range name for a single column.
 *
 * Priority:
 *   1. If perColumnOverride is non-empty, use it as the entire base
 *      (no case transform applied).
 *   2. Otherwise, apply the case style to columnName, then wrap with
 *      globalPrefix + globalSuffix.
 *
 * The result is always passed through sanitizeForNamedRange.
 */
export function buildFinalName(
  columnName: string,
  globalPrefix: string,
  globalSuffix: string,
  perColumnOverride: string,
  caseStyle: CaseStyle
): string {
  const base = perColumnOverride.trim()
    ? perColumnOverride.trim()
    : applyCaseStyle(columnName, caseStyle);

  const prefix = globalPrefix.trim();
  const suffix = globalSuffix.trim();
  const full = `${prefix}${base}${suffix}`;
  return sanitizeForNamedRange(full);
}

// ─── Validation ───────────────────────────────────────────────────────────────

/**
 * Returns true if name is a valid Excel named range identifier.
 *
 * Rules enforced:
 *   - Must not be empty.
 *   - Must start with a letter, underscore, or backslash.
 *   - Remaining characters must be letters, digits, underscores, or periods.
 *   - Must not look like a cell address (A1, R1C1-style).
 *   - Must not be the reserved single-character names "C" or "R".
 *   - Must be 255 characters or fewer.
 */
export function isValidExcelName(name: string): boolean {
  if (!name || name.length === 0) return false;
  if (name.length > 255) return false;

  // Must start with letter, underscore, or backslash
  if (!/^[a-zA-Z_\\]/.test(name)) return false;

  // Remaining characters: letters, digits, underscores, periods, backslashes
  if (!/^[a-zA-Z_\\][a-zA-Z0-9_.\\]*$/.test(name)) return false;

  // Reserved: single letters C or R
  if (/^[CRcr]$/.test(name)) return false;

  // Reject cell-address-like patterns: A1, AA123, $A$1
  if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(name)) return false;

  // Reject R1C1-style: R1C1, R123C45
  if (/^[Rr]\d+[Cc]\d+$/.test(name)) return false;

  return true;
}

// ─── Preview helper ───────────────────────────────────────────────────────────

/**
 * Compute the preview result for a single column.
 * Pure function — call in render, not stored in state.
 *
 * @param col            The column config for this row.
 * @param globalPrefix   The global prefix input value.
 * @param globalSuffix   The global suffix input value.
 * @param caseStyle      The selected case style.
 * @param existingNames  Upper-cased set of existing workbook/worksheet names.
 * @param siblingNames   Upper-cased set of final names from other included
 *                       rows in this batch — detects within-batch duplicates.
 */
export function derivePreview(
  col: ColumnConfig,
  globalPrefix: string,
  globalSuffix: string,
  caseStyle: CaseStyle,
  existingNames: Set<string>,
  siblingNames: Set<string>
): NamePreviewResult {
  if (!col.included) {
    return { finalName: "", isValid: true, isDuplicate: false };
  }

  const finalName = buildFinalName(
    col.originalName,
    globalPrefix,
    globalSuffix,
    col.perColumnOverride,
    caseStyle
  );

  const isValid = isValidExcelName(finalName);
  const upper = finalName.toUpperCase();
  const isDuplicate = existingNames.has(upper) || siblingNames.has(upper);

  return { finalName, isValid, isDuplicate };
}

/**
 * Build the full list of ColumnCreateSpec items from the current modal state,
 * ready to pass directly to createNamedRangesFromTableColumnsV2.
 *
 * Only included columns with valid, non-duplicate names are returned.
 */
export function buildColumnCreateSpecs(
  columns: ColumnConfig[],
  globalPrefix: string,
  globalSuffix: string,
  caseStyle: CaseStyle,
  existingNames: Set<string>
): ColumnCreateSpec[] {
  const specs: ColumnCreateSpec[] = [];
  const seenNames = new Set<string>();

  for (const col of columns) {
    if (!col.included) continue;

    const finalName = buildFinalName(
      col.originalName,
      globalPrefix,
      globalSuffix,
      col.perColumnOverride,
      caseStyle
    );

    const upper = finalName.toUpperCase();
    if (!isValidExcelName(finalName)) continue;
    if (existingNames.has(upper)) continue;
    if (seenNames.has(upper)) continue;

    seenNames.add(upper);
    specs.push({ columnName: col.originalName, finalName });
  }

  return specs;
}