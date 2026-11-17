#!/usr/bin/env node
// Builds dist/ : a static copy of the add-in for any HTTPS host (GitHub Pages, IIS, SharePoint...).
// Usage: node scripts/build-static.js https://you.github.io/slidetex
// The built-in MathJax engine works everywhere; the "LaTeX local" engine still needs
// `npm run serve` on the user's machine (the pane calls https://localhost:3000).
"use strict";
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const root = path.join(__dirname, "..");
const base = process.argv[2];
if (!base) {
  console.error("Usage: npm run build:static -- https://your.host/path");
  process.exit(1);
}
const dist = path.join(root, "dist");
fs.rmSync(dist, { recursive: true, force: true });
fs.cpSync(path.join(root, "public"), dist, { recursive: true });
fs.cpSync(path.join(root, "node_modules", "mathjax", "es5"), path.join(dist, "vendor", "mathjax"), { recursive: true });
execFileSync(process.execPath, [path.join(__dirname, "make-manifest.js"), base, path.join(dist, "manifest.xml")], { stdio: "inherit" });
console.log("Static build ready in dist/ — upload it, then sideload dist/manifest.xml.");
