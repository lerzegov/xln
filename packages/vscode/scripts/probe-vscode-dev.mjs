// Smoke test against the REAL vscode.dev, headless: installs this extension from
// https://localhost (Install Extension from Location), opens a folder, runs both probe
// commands and checks their output; then pulls f7_base.xlsx and browses the project
// (outline, go to definition, hover, cell usages) through the real UI.
//
// What is real: vscode.dev itself, its install-from-location path (CORS, https, Chrome's
// Local Network Access permission, granted here by Playwright), its file system provider
// over File System Access API handles, and our web bundle in its worker.
// What is simulated: the folder. A native directory picker cannot be automated, so
// window.showDirectoryPicker returns an Origin Private File System directory, which is a
// genuine FileSystemDirectoryHandle but not a folder on disk. Whether Chrome shows a
// real disk folder's "~$" files the same way is the author's manual check.
//
// Depends on vscode.dev's current UI (button labels, quick input); if it breaks after a
// vscode.dev update, fix the selectors, not the extension. Needs network and openssl.
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { chromium } from "playwright";
import { makeFixture, root } from "./fixture.mjs";
import { serve } from "./serve-https.mjs";

const PORT = 5077;
const headed = process.argv.includes("--headed");

// Throwaway self-signed certificate: the headless browser ignores certificate errors.
const certDir = join(root, ".vscode-test", "cert");
const cert = join(certDir, "cert.pem");
const key = join(certDir, "key.pem");
if (!existsSync(cert)) {
  mkdirSync(certDir, { recursive: true });
  execFileSync("openssl", ["req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", key, "-out", cert, "-days", "30", "-subj", "/CN=localhost"], { stdio: "ignore" });
}
makeFixture("vscode-dev-fixture");
const server = await serve({ port: PORT, cert, key });

const browser = await chromium.launch({ headless: !headed, args: ["--ignore-certificate-errors"] });
let failed = false;
try {
  const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  await context.grantPermissions(["local-network-access"], { origin: "https://vscode.dev" });
  await context.addInitScript((base) => {
    window.showDirectoryPicker = async () => {
      const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle("fixture", { create: true });
      for (const path of ["probe_mac.xlsx", "~$probe_mac.xlsx", "demo.xln", "book/f7_base.xlsx"]) {
        const bytes = await (await fetch(base + path.split("/").map(encodeURIComponent).join("/"))).arrayBuffer();
        const name = path.slice(path.lastIndexOf("/") + 1);
        const w = await (await dir.getFileHandle(name, { create: true })).createWritable();
        await w.write(bytes);
        await w.close();
      }
      return dir;
    };
  }, `https://localhost:${PORT}/.vscode-test/vscode-dev-fixture/`);

  const page = await context.newPage();
  const ready = async () => {
    await page.waitForSelector(".monaco-workbench", { timeout: 60000 });
    await page.waitForTimeout(5000);
  };
  const palette = async (command) => {
    await page.keyboard.press("F1");
    await page.waitForSelector(".quick-input-widget input", { state: "visible" });
    await page.keyboard.type(command);
    await page.waitForTimeout(1500);
    await page.keyboard.press("Enter");
  };

  console.log("vscode.dev: loading");
  await page.goto("https://vscode.dev", { waitUntil: "domcontentloaded" });
  await ready();

  console.log("vscode.dev: Install Extension from Location");
  await palette("Developer: Install Extension from Location");
  await page.waitForTimeout(1000);
  await page.keyboard.type(`https://localhost:${PORT}`);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(6000);
  await page.keyboard.press("Escape");

  console.log("vscode.dev: Open Folder");
  await page.getByRole("button", { name: "Open Folder" }).first().click();
  await page.waitForTimeout(3000);
  const trust = page.getByRole("button", { name: "Yes", exact: true });
  if (await trust.count()) await trust.first().click();
  await page.waitForTimeout(3000);
  await ready();

  console.log("vscode.dev: xln: Write test file");
  await palette("xln: Write test file");
  await page.waitForTimeout(3000);
  const written = await page.evaluate(async () => {
    const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle("fixture");
    return (await (await dir.getFileHandle("xln-write-test.txt")).getFile()).text();
  });

  console.log("vscode.dev: xln: Inspect workbook");
  await palette("xln: Inspect workbook");
  await page.waitForTimeout(2500);
  await page.keyboard.type("probe_mac");
  await page.waitForTimeout(800);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(3000);
  // Monaco renders spaces as U+00A0; only the last screenful of the channel is in the DOM.
  const text = async (selector) => (await page.locator(selector).first().innerText()).replace(/\u00a0/g, " ");
  const panel = await text(".panel");

  console.log("vscode.dev: xln: Pull workbook");
  await palette("xln: Pull workbook");
  await page.waitForTimeout(2500);
  await page.keyboard.type("f7_base");
  await page.waitForTimeout(800);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(5000);
  const pullPanel = await text(".panel");
  const pulled = await page
    .evaluate(async () => {
      const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle("fixture");
      const names = await (await dir.getDirectoryHandle("f7_base.xln")).getDirectoryHandle("names");
      return (await (await names.getFileHandle("_unmanaged.xln")).getFile()).text();
    })
    .catch((e) => String(e));

  console.log("vscode.dev: outline of names/_unmanaged.xln");
  await page.keyboard.press("Escape");
  await palette("Go to File");
  await page.waitForTimeout(800);
  await page.keyboard.type("_unmanaged.xln");
  await page.waitForTimeout(2000);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(3000);
  await palette("Go to Symbol in Editor");
  await page.waitForTimeout(2500);
  const outline = await text(".quick-input-list");
  await page.keyboard.type("Rate2");
  await page.waitForTimeout(800);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(800);

  console.log("vscode.dev: go to definition, hover, cell usages");
  // The cursor is on `Rate2` in `Rate2 = Rate*2;`: step onto the `Rate` it reads.
  await page.keyboard.press("End");
  for (let k = 0; k < 4; k++) await page.keyboard.press("ArrowLeft");
  await palette("Go to Definition");
  await page.waitForTimeout(2000);
  const status = await text(".statusbar");
  await palette("Show or Focus Hover");
  await page.waitForTimeout(2500);
  const hover = await text(".monaco-hover").catch(() => "");
  await page.keyboard.press("Escape");
  await palette("xln: Show cell usages");
  await page.waitForTimeout(3000);
  const usages = await text(".editor-instance .view-lines").catch(() => "");
  await page.screenshot({ path: join(root, ".vscode-test", "vscode-dev.png") });

  const checks = [
    ["write landed in the folder handle (host: vscode.dev)", /^xln write test .* from vscode\.dev/.test(written)],
    ["definedName count", panel.includes("<definedName> elements in xl/workbook.xml: 8")],
    ["lock file seen", panel.includes("Lock file ~$probe_mac.xlsx: PRESENT")],
    ["pull wrote f7_base.xln/names/_unmanaged.xln", pulled.includes("Rate2 = Rate*2;")],
    ["pull summary and model load", pullPanel.includes("7 names: 6 workbook-scoped") && pullPanel.includes("Loaded 1 project(s), 2 .xln files, 7 names")],
    ["outline lists the names", ["Fn", "Rate2", "Spl"].every((n) => outline.includes(n))],
    ["go to definition: Rate2 → Rate on line 6", status.includes("Ln 6, Col 1")],
    ["hover on Rate", hover.includes("constant · workbook scope")],
    ["cell usages of Rate", usages.includes("'S1'!A1: =Rate*2")],
  ];
  for (const [what, ok] of checks) {
    console.log(`${ok ? "PASS" : "FAIL"}  ${what}`);
    failed ||= !ok;
  }
  if (failed) {
    console.log("\nOutput panel text:\n" + panel + "\n---\n" + pullPanel);
    console.log("\nOutline:\n" + outline + "\nStatus bar:\n" + status + "\nHover:\n" + hover + "\nUsages:\n" + usages);
  }
  console.log(`Screenshot: ${join(root, ".vscode-test", "vscode-dev.png")}`);
} catch (err) {
  console.error(err);
  failed = true;
} finally {
  await browser.close();
  server.close();
}
process.exit(failed ? 1 : 0);
