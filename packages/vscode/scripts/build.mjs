// Bundles the extension twice from one source: dist/node for desktop VS Code ("main") and
// dist/web for vscode.dev ("browser"), plus the web test runner. The web bundles use
// platform "browser" with no polyfills, so any Node built-in that sneaks in fails the build.
// Both bundle @xln/core.
import * as esbuild from "esbuild";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const watch = process.argv.includes("--watch");

// @xln/core is bundled from its TypeScript sources, so a stale packages/core/dist can
// never end up in the extension (the root Vitest config does the same).
function coreFromSource() {
  return {
    name: "xln-core-from-source",
    setup(build) {
      build.onResolve({ filter: /^@xln\/core$/ }, () => ({ path: join(root, "..", "core", "src", "index.ts") }));
    },
  };
}

// The desktop bundle can run programs: it gets the Excel control that does (E7). The web
// bundles keep the stub, so no Node built-in reaches them.
function desktopExcelHost() {
  return {
    name: "xln-desktop-excel-host",
    setup(build) {
      build.onResolve({ filter: /^\.\/excelHost\.js$/ }, () => ({ path: join(root, "src", "excelHost.node.ts") }));
    },
  };
}

const common = {
  absWorkingDir: root,
  plugins: [coreFromSource()],
  bundle: true,
  format: "cjs",
  external: ["vscode"],
  sourcemap: true,
  logLevel: "info",
  target: "es2022",
};

const builds = [
  { ...common, entryPoints: ["src/extension.ts"], outfile: "dist/node/extension.js", platform: "node", plugins: [coreFromSource(), desktopExcelHost()] },
  { ...common, entryPoints: ["src/extension.ts"], outfile: "dist/web/extension.js", platform: "browser" },
  {
    ...common,
    entryPoints: ["src/test/node/index.ts"],
    outfile: "dist/node/test/index.js",
    platform: "node",
    external: ["vscode", "mocha"],
  },
  {
    ...common,
    entryPoints: ["src/test/web/index.ts"],
    outfile: "dist/web/test/index.js",
    platform: "browser",
    plugins: [coreFromSource(), mochaAsGlobal()],
  },
];

// mocha/mocha.js is a UMD file inside an ESM package, so esbuild does not wrap it and its
// `module.exports = factory()` replaces the test bundle's own exports (the `run` that
// test-web calls). Hiding module/exports/define makes it take its global branch instead.
function mochaAsGlobal() {
  return {
    name: "mocha-as-global",
    setup(build) {
      build.onLoad({ filter: /[\\/]mocha[\\/]mocha\.js$/ }, async (args) => {
        const { readFile } = await import("node:fs/promises");
        const src = await readFile(args.path, "utf8");
        return { contents: `(function (module, exports, define) {\n${src}\n}).call(globalThis);`, loader: "js" };
      });
    },
  };
}

if (watch) {
  for (const b of builds) await (await esbuild.context(b)).watch();
} else {
  await Promise.all(builds.map((b) => esbuild.build(b)));
}
