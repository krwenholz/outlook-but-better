import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    projects: [
      { test: { name: "unit", include: ["test/**/*.test.ts"] } },
      // Loads dist/ into real Chromium; run `mise run build` first (mise run e2e does).
      { test: { name: "e2e", include: ["test/e2e/**/*.e2e.ts"], testTimeout: 30_000 } },
    ],
  },
});
