import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { theme } from "./theme";

// Minimal global reset applied once at bootstrap (no CSS files needed).
const style = document.createElement("style");
style.textContent = `
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  html { scroll-behavior: smooth; scrollbar-gutter: stable; }
  body {
    font-family: ${theme.fontStack};
    background:
      radial-gradient(1200px 500px at 100% -10%, rgba(70,79,235,0.08), transparent 60%),
      radial-gradient(1000px 420px at -10% 0%, rgba(0,183,195,0.08), transparent 55%),
      ${theme.bg};
    background-attachment: fixed;
    color: ${theme.text};
    -webkit-font-smoothing: antialiased;
    text-rendering: optimizeLegibility;
  }
  input, textarea, button { font-family: inherit; }
  ::selection { background: rgba(0,120,212,0.18); }

  ::-webkit-scrollbar { width: 10px; height: 10px; }
  ::-webkit-scrollbar-thumb { background: #c7ced6; border-radius: 8px; border: 2px solid transparent; background-clip: content-box; }
  ::-webkit-scrollbar-thumb:hover { background: #aab3bd; background-clip: content-box; }

  @keyframes cardIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
`;
document.head.appendChild(style);

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
