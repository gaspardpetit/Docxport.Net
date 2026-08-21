import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  root: "react",
  base: "/nested/react/",
  plugins: [react()],
  build: { outDir: "../dist/nested/react", emptyOutDir: true }
});
