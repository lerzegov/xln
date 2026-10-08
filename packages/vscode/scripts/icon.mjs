// Renders media/icon.svg to media/icon.png (128x128, transparent corners) with the
// Playwright Chromium the web tests already use, so no image toolchain is needed.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const media = fileURLToPath(new URL("../media/", import.meta.url));
const svg = readFileSync(media + "icon.svg", "utf8");
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 128, height: 128 }, deviceScaleFactor: 1 });
await page.setContent(`<html><body style="margin:0;background:transparent">${svg}</body></html>`);
await page.locator("svg").screenshot({ path: media + "icon.png", omitBackground: true });
await browser.close();
console.log("wrote", media + "icon.png");
