import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// `base` must match the GitHub Pages sub-path (https://<user>.github.io/<repo>/).
// Override at build time with `VITE_BASE=/your-repo-name/ npm run build`, or
// edit the default below to match your repository name.
const base = process.env.VITE_BASE || "/DeepLinkDecoder/";

export default defineConfig({
  base,
  plugins: [react()],
  server: {
    port: 5173,
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
  },
});
