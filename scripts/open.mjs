#!/usr/bin/env node
// npm run open -- [folder] [--web | --vscode-dev]
//
// Opens a workbook folder with the xln extension already loaded, no clicks:
//   (default)     desktop VS Code with the extension under development; full disk access.
//   --web         VS Code for the Web served locally (@vscode/test-web) in Chrome, the same
//                 web bundle as vscode.dev. Its file system keeps changes in memory only, so
//                 the project is pulled to disk first with the CLI when it is missing.
//   --vscode-dev  the real vscode.dev: serves the extension and opens Chrome. Chrome only
//                 lets a person pick a local folder, so open it there with
//                 File > Open Recent (one click) or File > Open Folder.
// The folder defaults to the last one used (kept in .xln-open, not committed).
import { execFileSync, spawn } from "node:child_process";
import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const ext = join(root, "packages", "vscode");
const memo = join(root, ".xln-open");
const args = process.argv.slice(2);
const mode = args.includes("--web") ? "web" : args.includes("--vscode-dev") ? "vscode-dev" : "desktop";
const given = args.find((a) => !a.startsWith("--"));

let folder = given ? resolve(given.replace(/^~(?=\/|$)/, homedir())) : existsSync(memo) ? readFileSync(memo, "utf8").trim() : "";
if (!folder || !existsSync(folder)) {
  console.error("Give the folder that holds the workbook, e.g.\n\n  npm run open -- ~/Desktop/my-model\n");
  process.exit(2);
}
const dry = args.includes("--dry-run"); // print the launch instead of doing it (for tests)
if (!dry) writeFileSync(memo, folder + "\n");
const run = (cmd, argv, opts = {}) => execFileSync(cmd, argv, { stdio: "inherit", cwd: root, ...opts });
const launch = (cmd, argv) => dry ? console.log(`[dry-run] ${cmd} ${argv.join(" ")}`) : spawn(cmd, argv, { stdio: "inherit", cwd: root });
console.log(`xln: building the extension…`);
run("npm", ["run", "build", "-w", "xln", "--silent"], { stdio: ["ignore", "ignore", "inherit"] });

if (mode === "desktop") {
  console.log(`xln: opening ${folder} in VS Code with the extension loaded.`);
  if (dry) console.log(`[dry-run] code --extensionDevelopmentPath ${ext} ${folder}`);
  else run("code", ["--extensionDevelopmentPath", ext, folder]);
} else if (mode === "web") {
  // Pull every workbook that has no project yet, so browsing has something on disk.
  run("npm", ["run", "build", "-w", "@xln/core", "-w", "@xln/cli", "--if-present", "--silent"], { stdio: ["ignore", "ignore", "inherit"] });
  for (const f of readdirSync(folder)) {
    if (!/\.xls[xm]$/i.test(f) || f.startsWith("~$")) continue;
    const project = join(folder, f.replace(/\.xls[xm]$/i, ".xln"));
    if (!existsSync(project)) {
      console.log(`xln: pulling ${f} (the local web host cannot write to disk)…`);
      run("node", [join(root, "packages", "cli", "bin", "xln.js"), "pull", join(folder, f)]);
    }
  }
  console.log(`xln: opening ${folder} in VS Code for the Web (Chrome). Changes made there stay in memory. Ctrl+C to stop.`);
  launch(process.execPath, [join(root, "node_modules", "@vscode", "test-web", "out", "server", "index.js"),
    "--browserType=chromium", "--extensionDevelopmentPath", ext, folder]);
} else {
  console.log(`xln: in vscode.dev use File > Open Recent > ${folder.split("/").pop()} (or File > Open Folder).`);
  launch("npm", ["run", "serve", "-w", "xln", "--", "--open", "--quiet"]);
}
