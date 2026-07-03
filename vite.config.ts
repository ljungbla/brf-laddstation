import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// `base` must match the GitHub Pages repo path so assets resolve when hosted at
// https://<user>.github.io/brf-laddstation/. Change if the repo is renamed.
export default defineConfig({
  base: "/brf-laddstation/",
  plugins: [react()],
});
