#!/usr/bin/env node
import { main } from "../dist/main.js";

// `xln formulas … | head`: the reader closed the pipe; stop quietly, as Unix tools do.
process.stdout.on("error", (e) => {
  if (e.code === "EPIPE") process.exit(0);
  throw e;
});

process.exitCode = await main(process.argv.slice(2));
