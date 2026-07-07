import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { theme } from "./theme";

// Minimal global reset applied once at bootstrap (no CSS files needed).
const style = document.createElement("style");
style.textContent = `
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: ${theme.fontStack};
    background: ${theme.bg};
    color: ${theme.text};
    -webkit-font-smoothing: antialiased;
  }
  input, textarea, button { font-family: inherit; }
`;
document.head.appendChild(style);

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
