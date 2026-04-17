import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "e2e",
  webServer: {
    command: "npm run start",
    port: 3000,
    reuseExistingServer: true,
  },
  use: {
    baseURL: "http://localhost:3000",
  },
});
