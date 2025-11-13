import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      "@anyfile/core": resolve(rootDir, "core/src/index.ts"),
      "@anyfile/excel": resolve(rootDir, "excel/src/index.ts"),
    },
  },
  test: {
    globals: true,
    environment: "node",
    include: ["**/src/**/*.test.ts"],
  },
});

