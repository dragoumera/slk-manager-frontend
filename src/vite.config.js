import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Configuration de build pour SLK Manager (interface).
// Le build produit un dossier `dist/` statique, servi ensuite par Render.
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 2000,
  },
});
