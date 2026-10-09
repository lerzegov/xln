// `xln-mcp [--root <folder>]...`: the server on stdio. stdout is the protocol channel, so
// everything for a person goes to stderr.

import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { createServer, Roots } from "./server.js";

const USAGE = `usage: xln-mcp [--root <folder>]...

  An MCP server (stdio) giving AI agents xln's check, names, formulas, graph, pull,
  build plan, build, verify, rename and library commands. Tools read and write only
  under the roots (a library named in a project's config is also read elsewhere, never
  written); relative paths resolve against the first. Default root: the current folder.
    --root   a folder the tools may use (repeat for several)
`;

export function parseArgs(args: string[]): { roots: string[] } | { error: string } | { help: true } {
  const roots: string[] = [];
  for (let i = 0; i < args.length; i++) {
    const a = args[i]!;
    if (a === "-h" || a === "--help") return { help: true };
    if (a === "--root") {
      const v = args[++i];
      if (v === undefined) return { error: "--root needs a folder" };
      roots.push(v);
    } else if (a.startsWith("--root=")) roots.push(a.slice("--root=".length));
    else return { error: `unknown argument '${a}'` };
  }
  return { roots: roots.length ? roots : [process.cwd()] };
}

export async function main(args: string[]): Promise<number> {
  const parsed = parseArgs(args);
  if ("help" in parsed) {
    process.stderr.write(USAGE);
    return 0;
  }
  if ("error" in parsed) {
    process.stderr.write(`xln-mcp: ${parsed.error}\n${USAGE}`);
    return 2;
  }
  try {
    // Checked once here, so a bad --root fails at launch rather than on the first call.
    new Roots(parsed.roots);
  } catch (e) {
    process.stderr.write(`xln-mcp: ${e instanceof Error ? e.message : String(e)}\n`);
    return 2;
  }
  serveStdio(() => createServer({ roots: parsed.roots }), { onerror: (e) => process.stderr.write(`xln-mcp: ${e.message}\n`) });
  return 0;
}
