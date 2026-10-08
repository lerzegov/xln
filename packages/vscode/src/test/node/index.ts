// Entry point that @vscode/test-electron loads inside desktop VS Code's extension host.
// mocha stays external to the bundle and is resolved from node_modules at run time.
import Mocha from "mocha";
import { defineAuditSuite } from "../audit.suite.js";
import { defineBuildSuite } from "../build.suite.js";
import { defineBrowseSuite } from "../browse.suite.js";
import { defineEditorSuite } from "../editor.suite.js";
import { defineLibrarySuite } from "../library.suite.js";
import { defineProbeSuite } from "../probe.suite.js";
import { defineSmokeSuite, smokeConfig } from "../smoke.suite.js";

export async function run(): Promise<void> {
  const mocha = new Mocha({ ui: "tdd", reporter: "spec", timeout: 120000, color: true });
  mocha.suite.emit("pre-require", globalThis, "xln.suite", mocha);
  const smoke = await smokeConfig();
  if (smoke) defineSmokeSuite(smoke);
  else {
    defineProbeSuite(false);
    defineBrowseSuite();
    defineEditorSuite();
    defineLibrarySuite();
    defineAuditSuite();
    defineBuildSuite(false);
  }
  return new Promise((resolve, reject) => {
    mocha.run((failures) => (failures > 0 ? reject(new Error(`${failures} tests failed.`)) : resolve()));
  });
}
