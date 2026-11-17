#!/usr/bin/env node
/*
 * SlideTeX — local HTTPS server (npm mode). See lib/app.js for the routes.
 *
 * Usage:  node server.js [--port 3000] [--http]
 * Env:    SLIDETEX_PORT, SLIDETEX_ALLOWED_ORIGINS (comma separated),
 *         SLIDETEX_LATEX (latex binary), SLIDETEX_DVISVGM (dvisvgm binary)
 */
"use strict";

const { createApp, listenLoopback, toolInfo, warmUp, buildDocument, compile } = require("./lib/app");

const args = process.argv.slice(2);
const argVal = (name, def) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : def;
};
const PORT = parseInt(argVal("--port", process.env.SLIDETEX_PORT || "3000"), 10);
const USE_HTTP = args.includes("--http");

const app = createApp({ port: PORT });

async function start() {
  let options = null;
  if (!USE_HTTP) {
    try {
      const devCerts = require("office-addin-dev-certs");
      options = await devCerts.getHttpsServerOptions(); // creates + trusts a localhost CA on first run
    } catch (e) {
      console.error("Could not obtain the localhost HTTPS certificate:", e.message || e);
      console.error("Run `npm run certs` (accept the security prompt), then retry.");
      process.exit(1);
    }
  }
  try {
    await listenLoopback(app, PORT, options);
  } catch (e) {
    console.error(`Cannot listen on port ${PORT}:`, e.message);
    process.exit(1);
  }
  const info = await toolInfo();
  console.log(`SlideTeX server: ${USE_HTTP ? "http" : "https"}://localhost:${PORT}/taskpane.html`);
  console.log(`  LaTeX   : ${info.latex ? info.latexVersion : "not found (MathJax engine only)"}`);
  console.log(`  dvisvgm : ${info.dvisvgm ? info.dvisvgmVersion : "not found"}`);
  warmUp().catch(() => {});
}

if (require.main === module) start();
module.exports = { app, buildDocument, compile };
