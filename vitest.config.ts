import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "jsdom",
    setupFiles: ["./vitest.setup.ts"],
    alias: {
      "wasm-xlsxwriter": "../nodejs/wasm_xlsxwriter.js",
    }
  },
});
