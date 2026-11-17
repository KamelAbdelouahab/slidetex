#!/usr/bin/env node
// Usage: node scripts/make-manifest.js [baseUrl] [outFile]
// Default: https://localhost:3000 -> manifest.xml
"use strict";
const fs = require("fs");
const path = require("path");
const base = (process.argv[2] || "https://localhost:3000").replace(/\/+$/, "");
const out = process.argv[3] || path.join(__dirname, "..", "manifest.xml");
if (!/^https:\/\//.test(base)) {
  console.error("The add-in URL must be https:// (Office requirement).");
  process.exit(1);
}
const tpl = fs.readFileSync(path.join(__dirname, "..", "templates", "manifest.template.xml"), "utf8");
fs.writeFileSync(out, tpl.split("{{BASE}}").join(base).split("{{ORIGIN}}").join(new URL(base).origin));
console.log(`Manifest written: ${out}  (base ${base})`);
