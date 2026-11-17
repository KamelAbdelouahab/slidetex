#!/usr/bin/env node
/* End-to-end tests of the task pane against the Office.js mock (headless Chromium). */
"use strict";
const path = require("path");
const fs = require("fs");
let chromium;
try { ({ chromium } = require("playwright")); } catch (_) { ({ chromium } = require(require.resolve("playwright", { paths: ["/opt/node22/lib/node_modules", "/usr/lib/node_modules", "/usr/local/lib/node_modules"] }))); }
const { app } = require("../server.js");

const MOCK = fs.readFileSync(path.join(__dirname, "office-mock.js"), "utf8");
const OUT = process.env.SLIDETEX_TEST_OUT || path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });

// SLIDETEX_TEST_URL=https://localhost:41730 runs the suite against an already running server (e.g. SlideTeX.exe).
const EXTERNAL = process.env.SLIDETEX_TEST_URL;
const BASE = (port) => EXTERNAL || `http://localhost:${port}`;
let failures = 0;
function check(cond, msg) {
  console.log((cond ? "  ✓ " : "  ✗ ") + msg);
  if (!cond) failures++;
}

async function openPane(browser, port, initScript, legacy) {
  const page = await browser.newPage({ viewport: { width: 380, height: 900 }, ignoreHTTPSErrors: true });
  page.on("pageerror", (e) => { console.log("  [pageerror]", e.message); failures++; });
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ status: 200, contentType: "application/javascript", body: (initScript || "") + "\n" + MOCK }));
  if (legacy) {
    // Force the ES5/IE11 bundle: drop module scripts, turn nomodule ones into normal scripts.
    await page.route(`${BASE(port)}/taskpane.html`, async (route) => {
      const r = await route.fetch();
      const html = (await r.text()).replace(/<script type="module"[^>]*><\/script>/g, "").replace(/ nomodule/g, "");
      route.fulfill({ response: r, body: html });
    });
  }
  await page.goto(`${BASE(port)}/taskpane.html`);
  await page.waitForFunction(() => window.SlideTeX && window.SlideTeX.S.inOffice, null, { timeout: 15000 });
  await page.evaluate(() => MathJax.startup.promise);
  return page;
}

async function typeSource(page, text) {
  await page.fill("#src", text);
  await page.waitForTimeout(400);
  await page.waitForFunction(() => !document.querySelector("#preview").classList.contains("stale"));
}

const doc = (page) => page.evaluate(() => JSON.parse(JSON.stringify(window.__mockDoc)));

(async () => {
  const server = EXTERNAL ? { close() {} } : app.listen(0);
  const port = EXTERNAL ? 0 : server.address().port;
  const browser = await chromium.launch();
  try {
    // ------------------------------------------------------------ modern PowerPoint
    console.log("Microsoft 365 (PowerPointApi 1.10, SVG)");
    let page = await openPane(browser, port);
    await page.evaluate(() => localStorage.clear());
    check((await page.textContent("#envBadge")).includes("PowerPoint"), "detected PowerPoint host");
    check(await page.isHidden("#editBanner"), "no edit banner when nothing is selected");

    await typeSource(page, "\\int_0^\\infty e^{-x^2}\\,dx = \\frac{\\sqrt{\\pi}}{2}");
    const prevSvg = await page.$eval("#preview", (el) => el.innerHTML);
    check(prevSvg.startsWith("<svg") && !prevSvg.includes("<use") && /width="[\d.]+pt"/.test(prevSvg), "MathJax preview: self-contained SVG in pt");
    check(!/currentColor/i.test(prevSvg), "no currentColor left");
    await page.screenshot({ path: path.join(OUT, "pane.png") });

    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.length === 1);
    await page.waitForTimeout(300);
    let d = await doc(page);
    let sh = d.slides[0].shapes[0];
    let meta = JSON.parse(sh.tags.SLIDETEX || "null");
    check(sh.coercion === "xmlSvg", "inserted as SVG");
    check(meta && meta.src.includes("\\int_0"), "LaTeX source stored in shape tag");
    check(sh.name.startsWith("SlideTeX:"), "shape named");
    check(sh.altTextDescription.includes("\\int_0"), "alt text set");
    check(Math.abs(sh.width - meta.w) < 0.05 && sh.width > 50, `size ${sh.width.toFixed(1)}×${sh.height.toFixed(1)} pt`);
    fs.writeFileSync(path.join(OUT, "inserted.svg"), sh.data);
    check(d.selectedShapeIds.length === 0, "selection cleared after insert");
    check((await page.textContent("#btnPrimary")) === "Insérer", "button stays 'Insérer'");

    // user moves + scales the equation, then selects it
    await page.evaluate(() => {
      const s = window.__mockDoc.slides[0].shapes[0];
      s.left = 222; s.top = 111; s.width *= 2; s.height *= 2;
      window.__mockDoc.selectedShapeIds = [s.id];
      window.__mockFireSel();
    });
    await page.waitForFunction(() => document.querySelector("#btnPrimary").textContent === "Mettre à jour", null, { timeout: 5000 });
    check((await page.inputValue("#src")).includes("\\int_0"), "selecting the shape loads its source");
    check(!(await page.isHidden("#editBanner")), "edit banner shown");

    // add a dummy shape above to test z-order preservation
    await page.evaluate(() => window.__mockDoc.slides[0].shapes.push({ id: "999", name: "Rect", tags: {}, left: 0, top: 0, width: 10, height: 10 }));
    await typeSource(page, "\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}");
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.log.some((l) => l.op === "delete"));
    await page.waitForTimeout(300);
    d = await doc(page);
    const shapes = d.slides[0].shapes;
    sh = shapes.find((s) => s.tags.SLIDETEX);
    meta = JSON.parse(sh.tags.SLIDETEX);
    check(shapes.filter((s) => s.tags.SLIDETEX).length === 1, "old picture replaced (no duplicate)");
    check(meta.src.includes("\\sum"), "tag updated with new source");
    check(sh.left === 222 && sh.top === 111, "position kept");
    check(Math.abs(sh.width / meta.w - 2) < 0.01, "user scale (×2) kept");
    check(shapes.indexOf(sh) === 0, "z-order kept (below the rectangle)");
    check((await page.textContent("#btnPrimary")) === "Insérer", "back to insert mode");

    // TeX error
    await typeSource(page, "\\frac{a}{");
    await page.waitForTimeout(300);
    check(!(await page.isHidden("#error")), "TeX error displayed: " + (await page.textContent("#error")).slice(0, 60));

    // LaTeX engine
    await page.selectOption("#engine", "latex");
    await typeSource(page, "\\mathcal{L} = \\frac{1}{2} m \\dot{q}^2");
    await page.waitForTimeout(900);
    await page.waitForFunction(() => !document.querySelector("#preview").classList.contains("stale"));
    const lprev = await page.$eval("#preview", (el) => el.innerHTML);
    check(lprev.startsWith("<svg") && !lprev.includes("<use") && (await page.textContent("#renderTime")).includes("LaTeX"), "LaTeX engine preview (dvisvgm, uses inlined)");
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.filter((s) => s.tags.SLIDETEX).length === 2);
    await page.waitForTimeout(200);

    // TikZ via LaTeX engine
    await page.selectOption("#mode", "text");
    await page.click("#preambleBlock > summary");
    await page.fill("#preamble", "\\usepackage{tikz}");
    await page.dispatchEvent("#preamble", "input");
    await typeSource(page, "\\begin{tikzpicture}\\draw[thick,->] (0,0) -- (2,1) node[right]{$\\vec u$}; \\fill[blue!40] (1,0) circle (0.3);\\end{tikzpicture}");
    await page.waitForTimeout(1000);
    await page.waitForFunction(() => !document.querySelector("#preview").classList.contains("stale"));
    check(await page.isHidden("#error"), "TikZ renders with LaTeX engine");
    await page.screenshot({ path: path.join(OUT, "pane-tikz.png") });
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.filter((s) => s.tags.SLIDETEX).length === 3);
    await page.waitForTimeout(200);
    d = await doc(page);
    fs.writeFileSync(path.join(OUT, "tikz.svg"), d.slides[0].shapes.filter((s) => s.tags.SLIDETEX).pop().data);

    // PNG format
    await page.selectOption("#engine", "mathjax");
    await page.selectOption("#mode", "display");
    await page.selectOption("#format", "png");
    await typeSource(page, "E = mc^2");
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.filter((s) => s.tags.SLIDETEX).length === 4);
    await page.waitForTimeout(200);
    d = await doc(page);
    sh = d.slides[0].shapes.filter((s) => s.tags.SLIDETEX).pop();
    check(sh.coercion === "image" && sh.data.startsWith("iVBOR"), "PNG insertion (base64 PNG)");
    fs.writeFileSync(path.join(OUT, "inserted.png"), Buffer.from(sh.data, "base64"));

    // regenerate all with new colour
    await page.selectOption("#format", "svg");
    await page.fill("#colorHex", "#1f4e79");
    await page.dispatchEvent("#colorHex", "change");
    await page.fill("#size", "40");
    await page.dispatchEvent("#size", "input");
    await page.click("#toolsBlock > summary");
    await page.check("#regenApplyStyle");
    await page.click("#btnRegenAll");
    await page.waitForFunction(() => /régénérée/.test(document.querySelector("#status").textContent), null, { timeout: 30000 });
    d = await doc(page);
    const eqs = d.slides[0].shapes.filter((s) => s.tags.SLIDETEX);
    const metas = eqs.map((s) => JSON.parse(s.tags.SLIDETEX));
    check(eqs.length === 4 && metas.every((m) => m.color === "#1f4e79" && m.size === 40), "regenerate all: 4 equations restyled");
    console.log("   status:", await page.textContent("#status"));
    await page.close();

    // ------------------------------------------------------------ old PowerPoint 2016 (no PowerPointApi, no SVG)
    console.log("PowerPoint 2016 perpetual (no PowerPointApi, PNG only)");
    page = await openPane(browser, port, "window.__MOCK_PP_LEVEL__='1.0'; window.__MOCK_SVG__=false;");
    await page.evaluate(() => localStorage.clear());
    check(await page.$eval('#format option[value="svg"]', (o) => o.disabled), "SVG option disabled");
    await typeSource(page, "a^2+b^2=c^2");
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.length === 1);
    await page.waitForTimeout(200);
    d = await doc(page);
    check(d.slides[0].shapes[0].coercion === "image", "falls back to PNG");
    const st = await page.evaluate(() => Office.context.document.settings.get("slidetex.sources"));
    check(st && st[0].src === "a^2+b^2=c^2", "source kept in document settings");
    await page.close();

    // ------------------------------------------------------------ ES5 bundle (IE11 webview path)
    console.log("Legacy ES5 bundle");
    page = await openPane(browser, port, "window.__MOCK_PP_LEVEL__='1.0'; window.__MOCK_SVG__=false;", true);
    check(await page.evaluate(() => !!document.querySelector('link[href="legacy/taskpane.css"]')), "legacy bundle active");
    await typeSource(page, "\\oint_C \\vec F \\cdot d\\vec r");
    check((await page.$eval("#preview", (el) => el.innerHTML)).startsWith("<svg"), "legacy preview renders");
    // Simulate IE11's tainted canvas: PNG must come from the server (resvg).
    await page.evaluate(() => { HTMLCanvasElement.prototype.toDataURL = function () { throw new Error("SecurityError"); }; });
    await page.click("#btnPrimary");
    await page.waitForFunction(() => window.__mockDoc.slides[0].shapes.length === 1, null, { timeout: 10000 });
    d = await doc(page);
    check(d.slides[0].shapes[0].data.startsWith("iVBOR"), "PNG produced by the server fallback");
    fs.writeFileSync(path.join(OUT, "server-png.png"), Buffer.from(d.slides[0].shapes[0].data, "base64"));
    await page.close();
  } finally {
    await browser.close();
    server.close();
  }
  console.log(failures ? `\n${failures} FAILURE(S)` : "\nAll tests passed.");
  process.exit(failures ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
