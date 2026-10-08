import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

export default defineConfig({
  // Tests of other packages run against the core's sources, not a possibly stale dist/.
  resolve: {
    alias: {
      "@xln/core": fileURLToPath(new URL("packages/core/src/index.ts", import.meta.url)),
      "@xln/cli": fileURLToPath(new URL("packages/cli/src/main.ts", import.meta.url)),
    },
  },
  test: { include: ["packages/*/test/**/*.test.ts"] },
});
