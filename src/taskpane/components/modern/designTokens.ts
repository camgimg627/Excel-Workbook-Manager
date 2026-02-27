import { makeStyles } from "@fluentui/react-components";

export const MODERN_TOKENS = {
  colorBg: "#F9FAFB",
  colorSurface: "#FFFFFF",
  colorText: "#111827",
  colorTextMuted: "#6B7280",
  colorBorder: "#E5E7EB",
  colorBrand: "#2E7D32",
  colorBrandStrong: "#256729",
  colorAccent: "#F59E0B",
  colorDanger: "#B91C1C",
  shadowCard: "0 1px 2px rgba(17, 24, 39, 0.06), 0 6px 18px rgba(17, 24, 39, 0.06)",
  shadowCardHover: "0 2px 6px rgba(17, 24, 39, 0.1), 0 10px 24px rgba(17, 24, 39, 0.1)",
} as const;

export const useModernSharedStyles = makeStyles({
  page: {
    minHeight: "100vh",
    backgroundColor: MODERN_TOKENS.colorBg,
    color: MODERN_TOKENS.colorText,
  },
  card: {
    borderRadius: "8px",
    border: `1px solid ${MODERN_TOKENS.colorBorder}`,
    backgroundColor: MODERN_TOKENS.colorSurface,
    boxShadow: MODERN_TOKENS.shadowCard,
    padding: "16px",
    transitionDuration: "150ms",
    transitionProperty: "box-shadow, transform",
    transitionTimingFunction: "ease",
    selectors: {
      "&:hover": {
        boxShadow: MODERN_TOKENS.shadowCardHover,
      },
    },
  },
  cardTitle: {
    fontSize: "16px",
    fontWeight: 600,
    marginBottom: "8px",
  },
  cardSubtitle: {
    fontSize: "12px",
    color: MODERN_TOKENS.colorTextMuted,
    marginBottom: "16px",
  },
  sectionTitle: {
    fontSize: "20px",
    fontWeight: 700,
    marginBottom: "4px",
  },
  sectionSubtitle: {
    fontSize: "13px",
    color: MODERN_TOKENS.colorTextMuted,
    marginBottom: "24px",
  },
  mutedText: {
    color: MODERN_TOKENS.colorTextMuted,
    fontSize: "12px",
  },
  errorText: {
    color: MODERN_TOKENS.colorDanger,
    fontSize: "12px",
  },
  successText: {
    color: MODERN_TOKENS.colorBrandStrong,
    fontSize: "12px",
  },
});
