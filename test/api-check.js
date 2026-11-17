#!/usr/bin/env node
/*
 * Static check of the PowerPoint API surface used by the task pane against the official
 * typings (@types/office-js), so a method called on the wrong object (e.g.
 * Presentation.setSelectedShapes, which only exists on Slide) fails the tests even though
 * the Office mock would accept it.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const dts = fs.readFileSync(require.resolve("@types/office-js/index.d.ts"), "utf8");
const src = fs.readFileSync(path.join(__dirname, "..", "public", "taskpane.js"), "utf8");

const ns = dts.slice(dts.indexOf("declare namespace PowerPoint {"));
const ppNs = ns.slice(0, ns.indexOf("\ndeclare namespace ", 10));
function members(cls) {
  const i = ppNs.indexOf(`\n    class ${cls} extends`);
  if (i < 0) throw new Error("class not found: " + cls);
  const body = ppNs.slice(i, ppNs.indexOf("\n    }\n", i));
  const names = new Set();
  body.replace(/^\s{8}(?:readonly\s+)?([A-Za-z]+)\s*[(:?]/gm, (_, n) => names.add(n));
  return names;
}
const PRES = members("Presentation");
const SLIDE = members("Slide");
const SHAPE = members("Shape");

let fail = 0;
const check = (c, m) => { console.log((c ? "  ✓ " : "  ✗ ") + m); if (!c) fail++; };
console.log("PowerPoint API surface (vs @types/office-js)");

const presCalls = [...new Set([...src.matchAll(/ctx\.presentation\.([A-Za-z]+)/g)].map((m) => m[1]))];
for (const m of presCalls) check(PRES.has(m), `Presentation.${m} exists`);

const selCalls = [...src.matchAll(/([\w.()\[\]"' ]+)\.setSelectedShapes\(/g)].map((m) => m[1].trim());
for (const c of selCalls) check(/slides\.getItem\([^)]*\)$/.test(c) || /slide\b/i.test(c.split(".").pop()), `setSelectedShapes called on a Slide (${c})`);
check(SLIDE.has("setSelectedShapes") && !PRES.has("setSelectedShapes"), "typings: setSelectedShapes is on Slide only");

for (const m of ["setZOrder", "delete", "tags", "altTextDescription", "altTextTitle", "left", "top", "width", "height", "name", "zOrderPosition"]) {
  if (src.includes("." + m)) check(SHAPE.has(m), `Shape.${m} exists`);
}
console.log(fail ? `${fail} FAILURE(S)` : "API check passed.");
process.exit(fail ? 1 : 0);
