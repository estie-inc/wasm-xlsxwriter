import { defineConfig } from "vite";

export default defineConfig({
  // Only needed for local file: linked dependencies.
  // Not required when installed from npm.
  server: {
    fs: {
      allow: ["../.."],
    },
  },
});
