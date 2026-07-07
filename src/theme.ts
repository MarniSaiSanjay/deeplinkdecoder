export const theme = {
  primary: "#0078d4",
  accentGradient: "linear-gradient(135deg, #0078d4 0%, #00b7c3 100%)",
  heroGradient: "linear-gradient(135deg, #0078d4 0%, #464feb 50%, #5c2d91 100%)",
  bg: "#f5f6f7",
  cardBg: "#ffffff",
  cardBorder: "#e1e4e8",
  cardBorderLight: "#eef0f2",
  cardShadow: "0 1px 4px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)",
  cardShadowLg: "0 4px 16px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.04)",
  text: "#24292e",
  textSoft: "#586069",
  subtle: "#959da5",
  danger: "#b71c1c",
  dangerBg: "#fdecea",
  codeBg: "#0d1117",
  codeText: "#e6edf3",
  fieldBorder: "#d0d7de",
  fieldNestBg: "#f3f9fe",
  pillBg: "#e8f4fd",
  pillText: "#0078d4",
  fontStack: "'Segoe UI', system-ui, -apple-system, Roboto, sans-serif",
  monoStack: "'Cascadia Code', 'Consolas', 'SF Mono', Menlo, monospace",
} as const;

/** Badge colours keyed by decoded link surface. */
export const surfaceColors: Record<string, { bg: string; fg: string; dot: string }> = {
  M365: { bg: "#e8f1fc", fg: "#0b4a8f", dot: "#0078d4" },
  UDL: { bg: "#efe9f7", fg: "#4a2278", dot: "#5c2d91" },
  Word: { bg: "#e6effc", fg: "#123c78", dot: "#185abd" },
  "App Store": { bg: "#e9f7ee", fg: "#1b5e34", dot: "#2e9e54" },
  Other: { bg: "#eef0f2", fg: "#4a5259", dot: "#8a939b" },
};
