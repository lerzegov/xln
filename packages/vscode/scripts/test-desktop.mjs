// Runs the probe suite in desktop VS Code through @vscode/test-electron. It downloads
// its own VS Code into .vscode-test/, so it runs alongside an open VS Code window.
import { runTests } from "@vscode/test-electron";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { makeFixture, root, smokeFixture } from "./fixture.mjs";

try {
  await runTests({
    extensionDevelopmentPath: root,
    extensionTestsPath: join(root, "dist", "node", "test", "index.js"),
    // A short user-data dir: VS Code's IPC socket path must stay under 103 characters.
    launchArgs: [smokeFixture("desktop-smoke") ?? makeFixture("desktop-fixture"), "--disable-extensions", `--user-data-dir=${mkdtempSync(join(tmpdir(), "xln-"))}`],
  });
} catch (err) {
  console.error(err);
  process.exit(1);
}
