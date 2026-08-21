import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";

export default defineConfig({
  root: "vue",
  base: "/nested/vue/",
  plugins: [vue()],
  build: { outDir: "../dist/nested/vue", emptyOutDir: true }
});
