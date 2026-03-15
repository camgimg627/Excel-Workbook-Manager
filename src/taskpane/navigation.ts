import { NAVIGATION_SIGNAL_KEY } from "../shared/signals";

export { NAVIGATION_SIGNAL_KEY };

export type NavigationTarget =
  | "names"
  | "names-create"
  | "tables"
  | "formulas"
  | "queries"
  | "model-builder"
  | "pivots"
  | "format"
  | "sandbox-debug"
  | "settings"
  | "help";

export const NAVIGATION_TARGETS: ReadonlySet<NavigationTarget> = new Set<NavigationTarget>([
  "names",
  "names-create",
  "tables",
  "formulas",
  "queries",
  "model-builder",
  "pivots",
  "format",
  "sandbox-debug",
  "settings",
  "help",
]);

export function isNavigationTarget(value: string | null | undefined): value is NavigationTarget {
  return !!value && NAVIGATION_TARGETS.has(value as NavigationTarget);
}
