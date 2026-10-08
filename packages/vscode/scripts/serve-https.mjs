// Serves this extension folder at https://localhost:5443 so that vscode.dev can load it
// with "Developer: Install Extension from Location...". vscode.dev needs https with a
// certificate the browser trusts, and CORS. The certificate comes from mkcert, whose
// local CA must be installed once (`brew install mkcert && mkcert -install`).
// No dependency beyond Node: the headers below are all vscode.dev needs.
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, statSync } from "node:fs";
import { createServer } from "node:https";
import { extname, join, normalize, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { root } from "./fixture.mjs";

const TYPES = {
  ".json": "application/json",
  ".js": "text/javascript",
  ".map": "application/json",
  ".md": "text/markdown",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

/** Starts the server; resolves once it listens. `log` gets one line per request. */
export function serve({ port, cert, key, log = () => {} }) {
  const server = createServer({ cert: readFileSync(cert), key: readFileSync(key) }, (req, res) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "*");
    // Older Chrome: Private Network Access preflight for a public page fetching localhost.
    // Current Chrome asks the user instead (Local Network Access prompt).
    res.setHeader("Access-Control-Allow-Private-Network", "true");
    res.setHeader("Cache-Control", "no-store");
    if (req.method === "OPTIONS") {
      res.writeHead(204).end();
      return;
    }
    const path = normalize(join(root, decodeURIComponent(new URL(req.url ?? "/", "https://localhost").pathname)));
    const ok = (path === root || path.startsWith(root + sep)) && existsSync(path) && statSync(path).isFile();
    log(`${ok ? 200 : 404} ${req.method} ${req.url}`);
    if (!ok) {
      res.writeHead(404).end();
      return;
    }
    res.writeHead(200, { "Content-Type": TYPES[extname(path)] ?? "application/octet-stream" });
    res.end(req.method === "HEAD" ? undefined : readFileSync(path));
  });
  return new Promise((resolve) => server.listen(port, () => resolve(server)));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const port = Number(process.env.PORT ?? 5443);
  const certDir = join(root, ".certs");
  const cert = join(certDir, "localhost.pem");
  const key = join(certDir, "localhost-key.pem");
  if (!existsSync(cert) || !existsSync(key)) {
    mkdirSync(certDir, { recursive: true });
    try {
      execFileSync("mkcert", ["-cert-file", cert, "-key-file", key, "localhost", "127.0.0.1", "::1"], { stdio: "inherit" });
    } catch {
      console.error("\nmkcert is missing. Install it once, then rerun:\n\n  brew install mkcert\n  mkcert -install\n");
      process.exit(1);
    }
  }
  const quiet = process.argv.includes("--quiet");
  await serve({ port, cert, key, log: quiet ? () => {} : console.log });
  const url = `https://localhost:${port}`;
  console.log(`
Serving the xln extension at ${url}  (leave this running; Ctrl+C stops it)

Every session: vscode.dev fetches the extension's code from this server when it starts.
  1. In Chrome, https://vscode.dev (opened for you with --open).
  2. Open Folder > the folder with your workbook > View files > trust it.
  3. Right-click a .xlsx > xln commands.

First time in this Chrome profile only:
  F1 > Developer: Install Extension from Location... > ${url} > Allow "local network access".
  After a rebuild, reload the vscode.dev tab to get the new code.
`);
  if (process.argv.includes("--open")) {
    const opener = process.platform === "darwin" ? ["open", ["-a", "Google Chrome", "https://vscode.dev"]]
      : process.platform === "win32" ? ["cmd", ["/c", "start", "chrome", "https://vscode.dev"]]
      : ["xdg-open", ["https://vscode.dev"]];
    try { execFileSync(opener[0], opener[1]); } catch { console.log("Could not open Chrome; open https://vscode.dev yourself."); }
  }
}
