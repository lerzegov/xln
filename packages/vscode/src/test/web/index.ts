// Entry point that @vscode/test-web loads inside the browser extension host.
// mocha's browser build installs a global `mocha` (see the build script's plugin).
import "mocha/mocha.js";
import { defineAuditSuite } from "../audit.suite.js";
import { defineBuildSuite } from "../build.suite.js";
import { defineBrowseSuite } from "../browse.suite.js";
import { defineEditorSuite } from "../editor.suite.js";
import { defineLibrarySuite } from "../library.suite.js";
import { defineProbeSuite } from "../probe.suite.js";
import { defineSmokeSuite, smokeConfig } from "../smoke.suite.js";

export async function run(): Promise<void> {
  mocha.setup({ ui: "tdd", reporter: "spec", timeout: 120000 });
  const smoke = await smokeConfig();
  if (smoke) defineSmokeSuite(smoke);
  else {
    defineProbeSuite(true);
    defineBrowseSuite();
    defineEditorSuite();
    defineLibrarySuite();
    defineAuditSuite();
    defineBuildSuite(true);
  }
  return new Promise((resolve, reject) => {
    mocha.run((failures) => (failures > 0 ? reject(new Error(`${failures} tests failed.`)) : resolve()));
  });
}
