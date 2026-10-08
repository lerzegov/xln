// Runs the probe suite in headless Chromium through @vscode/test-web.
// `--headed` shows the browser.
import { runTests } from "@vscode/test-web";
import { join } from "node:path";
import { makeFixture, root, smokeFixture } from "./fixture.mjs";

try {
  await runTests({
    browserType: "chromium",
    headless: !process.argv.includes("--headed"),
    extensionDevelopmentPath: root,
    extensionTestsPath: join(root, "dist", "web", "test", "index.js"),
    folderPath: smokeFixture("web-smoke") ?? makeFixture("web-fixture"),
    quality: "stable",
  });
} catch (err) {
  console.error(err);
  process.exit(1);
}
