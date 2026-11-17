/*
 * SlideTeX — HTTP application shared by `node server.js` (npm mode) and SlideTeX.exe.
 *   createApp({ port, mountStatic, renderPng }) -> express app
 *   listenLoopback(app, port, tlsOptions|null)   -> listens on 127.0.0.1 and ::1 only
 */
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const { execFile } = require("child_process");
const express = require("express");

const LATEX_BIN = process.env.SLIDETEX_LATEX || "latex";
const DVISVGM_BIN = process.env.SLIDETEX_DVISVGM || "dvisvgm";
// First MiKTeX run builds latex.fmt and installs packages on the fly: allow time for that.
const COMPILE_TIMEOUT_MS = 180000;
const MAX_SOURCE_BYTES = 100 * 1024;
const VERSION = "1.1.1";

// ---------------------------------------------------------------- LaTeX
function run(bin, argv, opts) {
  return new Promise((resolve) => {
    execFile(bin, argv, { timeout: COMPILE_TIMEOUT_MS, maxBuffer: 16 * 1024 * 1024, windowsHide: true, ...opts },
      (err, stdout, stderr) => resolve({ err, stdout: String(stdout || ""), stderr: String(stderr || "") }));
  });
}

let toolInfoCache = null;
async function toolInfo() {
  if (toolInfoCache) return toolInfoCache;
  const l = await run(LATEX_BIN, ["--version"], { timeout: 10000 });
  const d = await run(DVISVGM_BIN, ["--version"], { timeout: 10000 });
  toolInfoCache = {
    latex: !l.err,
    dvisvgm: !d.err,
    latexVersion: l.err ? null : l.stdout.split(/\r?\n/)[0],
    dvisvgmVersion: d.err ? null : d.stdout.split(/\r?\n/)[0],
  };
  // Re-probe after a minute in case the user installs LaTeX meanwhile.
  setTimeout(() => (toolInfoCache = null), 60000).unref();
  return toolInfoCache;
}

function buildDocument({ source, mode, preamble, color }) {
  const hex = /^#?[0-9a-fA-F]{6}$/.test(color || "") ? color.replace("#", "").toUpperCase() : "000000";
  let body;
  // "dvisvgm" makes xcolor / graphicx / tikz use the dvisvgm driver.
  let classOpts = "dvisvgm,border=1pt";
  if (mode === "inline") {
    classOpts += ",varwidth";
    body = `$${source}$`;
  } else if (mode === "display") {
    classOpts += ",varwidth";
    body = `$\\displaystyle ${source}$`;
  } else {
    // "text": raw LaTeX (paragraphs, tikzpicture, tabular, ...). No varwidth so TikZ boxes stay tight.
    body = source;
  }
  return [
    `\\documentclass[${classOpts}]{standalone}`,
    "\\usepackage[T1]{fontenc}",
    "\\usepackage[utf8]{inputenc}",
    "\\usepackage{amsmath,amssymb,amsfonts}",
    "\\usepackage{xcolor}",
    preamble || "",
    "\\begin{document}",
    `\\color[HTML]{${hex}}%`,
    body,
    "\\end{document}",
    "",
  ].join("\n");
}

function parseLatexLog(log) {
  const lines = log.split(/\r?\n/);
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("!")) {
      const msg = [lines[i]];
      for (let j = i + 1; j < Math.min(lines.length, i + 8); j++) {
        if (/^l\.\d+/.test(lines[j])) { msg.push(lines[j]); break; }
      }
      out.push(msg.join("  "));
    }
  }
  if (!out.length) {
    const tail = lines.filter(Boolean).slice(-6).join("\n");
    out.push(tail || "Unknown LaTeX error");
  }
  return out.slice(0, 3).join("\n");
}

const cache = new Map(); // key -> svg (simple LRU)
function cacheGet(k) {
  if (!cache.has(k)) return null;
  const v = cache.get(k);
  cache.delete(k);
  cache.set(k, v);
  return v;
}
function cachePut(k, v) {
  cache.set(k, v);
  while (cache.size > 300) cache.delete(cache.keys().next().value);
}

async function compileOnce(params) {
  const tex = buildDocument(params);
  const key = crypto.createHash("sha256").update(tex).digest("hex");
  const hit = cacheGet(key);
  if (hit) return { ok: true, svg: hit, cached: true };

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "slidetex-"));
  try {
    fs.writeFileSync(path.join(dir, "eq.tex"), tex, "utf8");
    const env = {
      ...process.env,
      openin_any: "p", // TeX Live: no reading outside the work dir
      openout_any: "p",
      shell_escape: "f",
    };
    const l = await run(LATEX_BIN, [
      "-interaction=nonstopmode", "-halt-on-error", "-no-shell-escape", "-file-line-error-style",
      "eq.tex",
    ], { cwd: dir, env });
    const logPath = path.join(dir, "eq.log");
    const log = fs.existsSync(logPath) ? fs.readFileSync(logPath, "latin1") : l.stdout + l.stderr;
    if (l.err || !fs.existsSync(path.join(dir, "eq.dvi"))) {
      if (l.err && l.err.code === "ENOENT") return { ok: false, error: `LaTeX binary not found (${LATEX_BIN}). Install TeX Live or MiKTeX.` };
      if (l.err && l.err.killed) return { ok: false, error: "LaTeX timed out (180 s)." };
      return { ok: false, error: parseLatexLog(log) };
    }
    const d = await run(DVISVGM_BIN, [
      "--no-fonts", "--exact-bbox", "--precision=3", "--page=1", "--output=eq.svg", "eq.dvi",
    ], { cwd: dir, env });
    const svgPath = path.join(dir, "eq.svg");
    if (!fs.existsSync(svgPath)) {
      if (d.err && d.err.code === "ENOENT") return { ok: false, error: `dvisvgm not found (${DVISVGM_BIN}).` };
      return { ok: false, error: "dvisvgm failed: " + (d.stderr || d.stdout).slice(-600) };
    }
    const svg = fs.readFileSync(svgPath, "utf8");
    cachePut(key, svg);
    return { ok: true, svg };
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {});
  }
}


/*
 * Compilations run ONE AT A TIME. MiKTeX builds latex.fmt / installs packages on first use, and
 * two concurrent `latex` processes then race on the same files ("Windows API error 183").
 * Identical requests share one compilation; previews the client abandoned are skipped.
 */
let queue = Promise.resolve();
const inflight = new Map();
const TRANSIENT = /did not succeed|memory dump file|could not be found|error 183|fmt file|Fatal format file error|Permission denied|being used by another process/i;

function compile(params, isCancelled) {
  const key = crypto.createHash("sha256").update(buildDocument(params)).digest("hex");
  const waiter = isCancelled || (() => false);
  let entry = inflight.get(key);
  if (!entry) {
    entry = { waiters: new Set() };
    // Skip only when EVERY requester of this exact formula has gone away.
    const allGone = () => [...entry.waiters].every((w) => w());
    entry.task = queue.then(async () => {
      if (allGone()) return { ok: false, error: "cancelled", cancelled: true };
      let r = await compileOnce(params);
      if (!r.ok && TRANSIENT.test(r.error || "")) {
        // MiKTeX first-run hiccup (format being built / package just installed): retry once.
        await new Promise((ok) => setTimeout(ok, 1500));
        r = await compileOnce(params);
      }
      return r;
    });
    queue = entry.task.catch(() => {});
    inflight.set(key, entry);
    entry.task.finally(() => inflight.delete(key));
  }
  entry.waiters.add(waiter);
  return entry.task;
}

/** Compile a trivial formula in the background so MiKTeX prepares formats/packages early. */
async function warmUp(log = console.log) {
  const info = await toolInfo();
  if (!(info.latex && info.dvisvgm)) return null;
  const t = Date.now();
  const r = await compile({ source: "x^2", mode: "display", preamble: "", color: "#000000" });
  log(`LaTeX warm-up: ${r.ok ? "OK" : "FAILED — " + String(r.error).split("\n")[0]} (${Date.now() - t} ms)`);
  return r;
}

function defaultRenderPng(svg, px) {
  const { Resvg } = require("@resvg/resvg-js");
  return new Resvg(svg, { fitTo: { mode: "width", value: px }, background: "rgba(0,0,0,0)" }).render().asPng();
}

function createApp(opts = {}) {
  const PORT = opts.port || 3000;
  const renderPng = opts.renderPng || defaultRenderPng;
  const ALLOWED_ORIGINS = new Set(
    (process.env.SLIDETEX_ALLOWED_ORIGINS || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .concat([`https://localhost:${PORT}`, `http://localhost:${PORT}`, `https://127.0.0.1:${PORT}`])
  );
  const app = express();
  app.disable("x-powered-by");

  // ---------------------------------------------------------------- CORS
  app.use("/api", (req, res, next) => {
    const origin = req.headers.origin;
    const sameOrigin = origin && origin === `${req.protocol}://${req.headers.host}`;
    const allowed = sameOrigin || (origin && ALLOWED_ORIGINS.has(origin));
    if (origin && allowed) {
      res.setHeader("Access-Control-Allow-Origin", origin);
      res.setHeader("Vary", "Origin");
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-SlideTeX");
      // Lets an add-in hosted on a public https origin reach this local server (Chrome PNA).
      res.setHeader("Access-Control-Allow-Private-Network", "true");
    }
    if (req.method === "OPTIONS") return res.sendStatus(allowed ? 204 : 403);
    if (origin && !allowed) return res.status(403).json({ ok: false, error: "Origin not allowed" });
    next();
  });


  // ---------------------------------------------------------------- static
  app.use((req, res, next) => {
    res.setHeader("Cache-Control", "no-store");
    next();
  });
  if (opts.mountStatic) opts.mountStatic(app);
  else {
    app.use("/vendor/mathjax", express.static(path.join(__dirname, "..", "node_modules", "mathjax", "es5")));
    app.use(express.static(path.join(__dirname, "..", "public")));
  }

  // PNG rasterisation for webviews that cannot export a canvas that drew an SVG (IE11).
  app.post("/api/png", express.json({ limit: "8mb" }), async (req, res) => {
    if (req.get("X-SlideTeX") !== "1") return res.status(400).json({ ok: false, error: "Missing X-SlideTeX header" });
    const { svg, widthPt, dpi } = req.body || {};
    if (typeof svg !== "string" || !svg.startsWith("<svg")) return res.status(400).json({ ok: false, error: "Bad SVG" });
    try {
      const px = Math.min(12000, Math.max(1, Math.round((Number(widthPt) || 100) * (Number(dpi) || 300) / 72)));
      const png = await renderPng(svg, px);
      res.json({ ok: true, png: Buffer.from(png).toString("base64") });
    } catch (e) {
      res.status(500).json({ ok: false, error: "PNG: " + (e && e.message || e) });
    }
  });

  app.get("/api/health", async (req, res) => {
    const info = await toolInfo();
    res.json({ ok: true, name: "SlideTeX", version: VERSION, ...info });
  });

  app.post("/api/latex", express.json({ limit: MAX_SOURCE_BYTES * 2 }), async (req, res) => {
    if (req.get("X-SlideTeX") !== "1") return res.status(400).json({ ok: false, error: "Missing X-SlideTeX header" });
    const { source, mode, preamble, color } = req.body || {};
    if (typeof source !== "string" || !source.trim()) return res.status(400).json({ ok: false, error: "Empty source" });
    if (Buffer.byteLength(source + (preamble || "")) > MAX_SOURCE_BYTES) return res.status(413).json({ ok: false, error: "Source too large" });
    if (/\\(write18|immediate\s*\\write18)/.test(source + (preamble || ""))) return res.status(400).json({ ok: false, error: "\\write18 is not allowed" });
    try {
      let gone = false;
      res.on("close", () => { if (!res.writableEnded) gone = true; }); // client aborted (stale preview)
      const r = await compile(
        { source, mode: ["inline", "display", "text"].includes(mode) ? mode : "display", preamble: String(preamble || ""), color },
        () => gone
      );
      if (r.cancelled) return res.headersSent || gone ? undefined : res.status(409).json(r);
      res.status(r.ok ? 200 : 422).json(r);
    } catch (e) {
      res.status(500).json({ ok: false, error: String(e && e.message || e) });
    }
  });


  return app;
}

/** Listen on loopback only (IPv4 + IPv6) so "localhost" works whatever it resolves to. */
function listenLoopback(app, port, tlsOptions) {
  const make = () => (tlsOptions ? require("https").createServer(tlsOptions, app) : require("http").createServer(app));
  const listen = (host) =>
    new Promise((resolve, reject) => {
      const s = make();
      s.once("error", (e) => (host === "127.0.0.1" ? reject(e) : resolve(null)));
      s.listen(port, host, () => resolve(s));
    });
  return listen("127.0.0.1").then((a) => listen("::1").then((b) => [a, b].filter(Boolean)));
}

module.exports = { createApp, listenLoopback, toolInfo, buildDocument, compile, warmUp, VERSION };
