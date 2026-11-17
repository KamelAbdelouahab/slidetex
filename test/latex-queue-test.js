#!/usr/bin/env node
/* Checks that LaTeX compilations never overlap and that a MiKTeX first-run failure is retried. */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const state = fs.mkdtempSync(path.join(os.tmpdir(), "fake-tex-"));
process.env.SLIDETEX_FAKE_STATE = state;
process.env.SLIDETEX_LATEX = path.join(__dirname, "fake-tex", "latex");
const { compile } = require("../lib/app");

let fail = 0;
const check = (c, m) => { console.log((c ? "  ✓ " : "  ✗ ") + m); if (!c) fail++; };
(async () => {
  console.log("LaTeX queue (MiKTeX first-run simulation)");
  const jobs = ["a+b", "\\frac{1}{2}", "a+b", "x^2", "\\sqrt{y}"].map((src) => compile({ source: src, mode: "display", preamble: "", color: "#000000" }));
  let gone = false;
  const cancelled = compile({ source: "\\alpha_{stale}", mode: "display", preamble: "", color: "#000000" }, () => gone);
  gone = true; // client aborted this preview before it started
  const res = await Promise.all(jobs);
  const c = await cancelled;
  check(res.every((r) => r.ok), "all formulas compiled despite the first-run failure (retry): " + res.map((r) => (r.ok ? "ok" : r.error.slice(0, 40))).join(" | "));
  check(!fs.existsSync(path.join(state, "overlaps")), "no two latex processes ever ran at the same time");
  const runs = fs.readFileSync(path.join(state, "runs"), "utf8").trim().split("\n").length;
  check(runs === 5, `identical requests shared one compilation (latex runs: ${runs}, expected 5 = 4 distinct + 1 retry)`);
  check(c.cancelled === true, "abandoned preview skipped");
  fs.rmSync(state, { recursive: true, force: true });
  console.log(fail ? `${fail} FAILURE(S)` : "Queue tests passed.");
  process.exit(fail ? 1 : 0);
})();
