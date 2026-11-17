/*
 * SlideTeX — LaTeX equations for PowerPoint 2016+ / Microsoft 365 / PowerPoint on the web.
 *
 * Rendering : MathJax 3 (bundled, offline) or a real LaTeX install through the local server.
 * Insertion : Office common API setSelectedDataAsync (SVG, or PNG for older builds).
 * Editing   : the LaTeX source is stored in a shape tag (PowerPointApi 1.3+) so equations
 *             can be re-opened, edited and regenerated in place (PowerPointApi 1.5+).
 */
(function () {
  "use strict";

  const TAG = "SLIDETEX";
  const LS_KEY = "slidetex.settings.v1";
  const LS_HIST = "slidetex.history.v1";
  const $ = (id) => document.getElementById(id);

  // ------------------------------------------------------------------ state
  const defaults = {
    engine: "mathjax",
    mode: "display",
    size: 28,
    color: "#000000",
    format: "svg",
    dpi: "300",
    preamble: "",
    serverUrl: "",
    tab: 0,
    src: "",
  };
  const S = {
    settings: { ...defaults, ...load(LS_KEY, {}) },
    inOffice: false,
    caps: { svg: false, png: false, pp13: false, pp15: false, pp18: false, pp110: false },
    latex: { checked: false, ok: false, info: null },
    target: null, // {shapeId, slideId, meta} — selected SlideTeX shape
    loadedFrom: null, // shapeId whose source is currently in the editor
    dirty: false,
    busy: false,
    last: null, // last successful render {key, result}
    renderSeq: 0,
  };

  function load(key, def) {
    try {
      const v = localStorage.getItem(key);
      return v ? JSON.parse(v) : def;
    } catch (_) {
      return def;
    }
  }
  function save(key, val) {
    try {
      localStorage.setItem(key, JSON.stringify(val));
    } catch (_) {}
  }
  const saveSettings = debounce(() => save(LS_KEY, S.settings), 300);

  function debounce(fn, ms) {
    let t;
    return function (...a) {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(this, a), ms);
    };
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ------------------------------------------------------------------ status
  function status(msg, kind) {
    const el = $("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }
  function showError(msg) {
    const el = $("error");
    if (!msg) {
      el.hidden = true;
      el.textContent = "";
    } else {
      el.hidden = false;
      el.textContent = msg;
    }
  }

  // ------------------------------------------------------------------ rendering
  function serverBase() {
    const s = (S.settings.serverUrl || "").trim().replace(/\/+$/, "");
    if (s) return s;
    if (location.protocol.startsWith("http") && /^(localhost|127\.0\.0\.1)$/.test(location.hostname)) return location.origin;
    return "https://localhost:3000";
  }

  /** MathJax understands \newcommand & co but not \usepackage: keep only macro definitions. */
  function mathjaxPreamble(preamble) {
    const keep = /^\s*\\(newcommand|renewcommand|providecommand|def|let|DeclareMathOperator)\b/;
    return (preamble || "")
      .split(/\r?\n/)
      .map((l) => l.replace(/(^|[^\\])%.*$/, "$1"))
      .filter((l) => keep.test(l))
      .join("\n");
  }

  async function renderMathJax(p) {
    if (!window.MathJax || !MathJax.startup) throw new Error("MathJax n'est pas chargé.");
    await MathJax.startup.promise;
    let body = p.src;
    if (p.mode === "text") body = `\\text{${p.src}}`;
    const tex = mathjaxPreamble(p.preamble) + "\n" + body;
    MathJax.texReset && MathJax.texReset();
    const node = await MathJax.tex2svgPromise(tex, { display: p.mode !== "inline" });
    const svgEl = node.querySelector("svg");
    if (!svgEl) throw new Error("Rendu vide");
    const err = svgEl.querySelector('[data-mml-node="merror"]');
    if (err) {
      const t = err.getAttribute("data-mjx-error") || err.textContent || "Erreur TeX";
      throw new Error(t);
    }
    // MathJax viewBox units are 1/1000 em; 1 em = the chosen font size.
    return SlideTeXSvg.normalizeSvg(new XMLSerializer().serializeToString(svgEl), { unitToPt: p.size / 1000, color: p.color, pad: 40 });
  }

  async function renderLatex(p, signal) {
    let res;
    const slow = setTimeout(() => status("Compilation LaTeX… (au tout premier lancement, MiKTeX peut mettre une minute)"), 2500);
    try {
      res = await fetch(serverBase() + "/api/latex", {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-SlideTeX": "1" },
        body: JSON.stringify({ source: p.src, mode: p.mode, preamble: p.preamble, color: p.color }),
        signal,
      });
    } catch (e) {
      if (e && e.name === "AbortError") throw e;
      throw new Error(`Serveur LaTeX injoignable (${serverBase()}). Relancez SlideTeX.exe (ou « npm start ») ou choisissez MathJax.`);
    } finally {
      clearTimeout(slow);
    }
    const data = await res.json().catch(() => ({}));
    if (!data.ok) throw new Error(data.error || `Erreur serveur (${res.status})`);
    // dvisvgm units are TeX points at a 10 pt base font.
    // 1.5 pt padding so thick TikZ strokes / arrow tips are never clipped.
    return SlideTeXSvg.normalizeSvg(data.svg, { unitToPt: p.size / 10, color: p.color, pad: 1.5 });
  }

  function currentParams() {
    const st = S.settings;
    return {
      src: $("src").value,
      mode: st.mode,
      engine: st.engine,
      size: clamp(parseFloat(st.size) || 28, 4, 400),
      color: normHex(st.color) || "#000000",
      preamble: st.preamble || "",
    };
  }

  function paramsKey(p) {
    return JSON.stringify([p.src, p.mode, p.engine, p.size, p.color, p.engine === "latex" ? p.preamble : mathjaxPreamble(p.preamble)]);
  }

  async function render(p, signal) {
    const key = paramsKey(p);
    if (S.last && S.last.key === key) return S.last.result;
    const result = p.engine === "latex" ? await renderLatex(p, signal) : await renderMathJax(p);
    S.last = { key, result };
    return result;
  }

  async function updatePreview() {
    const p = currentParams();
    const seq = ++S.renderSeq;
    const prev = $("preview");
    if (!p.src.trim()) {
      prev.innerHTML = '<span class="muted">Aperçu</span>';
      $("dims").textContent = "";
      $("renderTime").textContent = "";
      showError(null);
      return;
    }
    prev.classList.add("stale");
    const t0 = performance.now();
    try {
      // A newer keystroke supersedes the previous LaTeX preview: abort it so the server skips it.
      if (S.previewAbort) S.previewAbort.abort();
      S.previewAbort = window.AbortController ? new AbortController() : null;
      const r = await render(p, S.previewAbort && S.previewAbort.signal);
      if (seq !== S.renderSeq) return;
      prev.innerHTML = r.svg;
      const svg = prev.querySelector("svg");
      if (svg) {
        // Show at true size (pt → CSS px) but never wider than the pane.
        svg.style.width = (r.widthPt * 96) / 72 + "px";
        svg.style.height = "auto";
      }
      $("dims").textContent = `${r.widthPt.toFixed(1)} × ${r.heightPt.toFixed(1)} pt`;
      $("renderTime").textContent = `${Math.round(performance.now() - t0)} ms · ${p.engine === "latex" ? "LaTeX" : "MathJax"}`;
      showError(null);
      if (/^Compilation LaTeX/.test($("status").textContent)) status("Prêt.");
    } catch (e) {
      if (seq !== S.renderSeq || (e && e.name === "AbortError")) return;
      showError(e.message || String(e));
    } finally {
      if (seq === S.renderSeq) prev.classList.remove("stale");
    }
  }
  const schedulePreview = debounce(updatePreview, 180);
  const schedulePreviewSlow = debounce(updatePreview, 650);
  function onSourceChanged() {
    (S.settings.engine === "latex" ? schedulePreviewSlow : schedulePreview)();
  }

  /** PNG from SVG: in-browser canvas, or the local server (resvg) when the webview can't (IE11). */
  async function toPng(r) {
    const dpi = parseInt(S.settings.dpi, 10) || 300;
    try {
      return await SlideTeXSvg.svgToPngBase64(r.svg, r.widthPt, r.heightPt, dpi);
    } catch (e) {
      if (!e.tainted) throw e;
      let res;
      try {
        res = await fetch(serverBase() + "/api/png", {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-SlideTeX": "1" },
          body: JSON.stringify({ svg: r.svg, widthPt: r.widthPt, dpi: dpi }),
        });
      } catch (_) {
        throw new Error("Conversion PNG impossible dans ce navigateur et serveur local injoignable.");
      }
      const j = await res.json();
      if (!j.ok) throw new Error(j.error || "Conversion PNG échouée");
      return j.png;
    }
  }

  // ------------------------------------------------------------------ Office helpers
  function officeAsync(fn) {
    return new Promise((resolve, reject) => {
      fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error ? r.error.message : "Erreur Office"))));
    });
  }

  function detectCaps() {
    const req = Office.context.requirements;
    const has = (s, v) => {
      try {
        return req.isSetSupported(s, v);
      } catch (_) {
        return false;
      }
    };
    S.caps = {
      svg: has("ImageCoercion", "1.2"),
      png: has("ImageCoercion", "1.1"),
      pp13: has("PowerPointApi", "1.3"),
      pp15: has("PowerPointApi", "1.5"),
      pp18: has("PowerPointApi", "1.8"),
      pp110: has("PowerPointApi", "1.10"),
    };
  }

  /**
   * Clear the shape selection on a slide (best effort, never throws).
   * The API is Slide.setSelectedShapes (PowerPointApi 1.5) — not Presentation.
   */
  async function clearShapeSelection(ctx, slideId) {
    try {
      ctx.presentation.slides.getItem(slideId).setSelectedShapes([]);
      await ctx.sync();
      return true;
    } catch (e) {
      console.warn("setSelectedShapes([]) failed, falling back to setSelectedSlides", e);
    }
    try {
      ctx.presentation.setSelectedSlides([slideId]);
      await ctx.sync();
      return true;
    } catch (e) {
      console.warn("could not clear selection", e);
      return false;
    }
  }

  /** Data stored in the shape tag. */
  function makeMeta(p, r) {
    return { v: 1, src: p.src, mode: p.mode, engine: p.engine, size: p.size, color: p.color, preamble: p.preamble, w: round2(r.widthPt), h: round2(r.heightPt) };
  }
  function parseMeta(s) {
    try {
      const m = JSON.parse(s);
      return m && typeof m.src === "string" ? m : null;
    } catch (_) {
      return null;
    }
  }

  async function getSelectionInfo() {
    if (!S.caps.pp15) return null;
    return PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items/id");
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items/id");
      await ctx.sync();
      const slideId = slides.items.length ? slides.items[0].id : null;
      if (shapes.items.length !== 1) return { slideId, target: null, count: shapes.items.length };
      const sh = shapes.items[0];
      const tag = sh.tags.getItemOrNullObject(TAG);
      tag.load("value");
      await ctx.sync();
      const meta = tag.isNullObject ? null : parseMeta(tag.value);
      return { slideId, count: 1, target: meta ? { shapeId: sh.id, slideId, meta } : null };
    });
  }

  /**
   * Insert a rendered formula on the current slide.
   * If `replace` = {shapeId, slideId, meta} the old picture is replaced, keeping position,
   * user scaling and z-order.
   */
  async function insertRendered(r, p, replace) {
    const meta = makeMeta(p, r);
    let fmt = S.settings.format;
    if (fmt === "svg" && !S.caps.svg) fmt = "png";
    let w = r.widthPt;
    let h = r.heightPt;

    let slideId = null;
    let before = null;
    let old = null;

    if (S.caps.pp15) {
      await PowerPoint.run(async (ctx) => {
        if (replace) {
          slideId = replace.slideId;
          ctx.presentation.setSelectedSlides([slideId]);
          await ctx.sync();
        }
        const slides = ctx.presentation.getSelectedSlides();
        slides.load("items/id");
        await ctx.sync();
        if (!slides.items.length) throw new Error("Sélectionnez une diapositive.");
        slideId = slides.items[0].id;
        const shapes = ctx.presentation.slides.getItem(slideId).shapes;
        shapes.load("items/id");
        if (replace) {
          const sh = shapes.getItemOrNullObject(replace.shapeId);
          sh.load(S.caps.pp18 ? "id,left,top,width,height,zOrderPosition" : "id,left,top,width,height");
          await ctx.sync();
          if (!sh.isNullObject) old = { id: sh.id, left: sh.left, top: sh.top, width: sh.width, height: sh.height, z: S.caps.pp18 ? sh.zOrderPosition : null };
        } else {
          await ctx.sync();
        }
        before = new Set(shapes.items.map((s) => s.id));
        // Deselect so the new picture never replaces / lands inside the current selection.
        await clearShapeSelection(ctx, slideId);
      });
    }

    const opts = {};
    if (old) {
      // Keep the scale the user may have applied to the previous picture.
      const sc = replace.meta && replace.meta.w ? old.width / replace.meta.w : 1;
      if (isFinite(sc) && sc > 0.05 && sc < 50) {
        w *= sc;
        h *= sc;
      }
      opts.imageLeft = old.left;
      opts.imageTop = old.top;
    }
    opts.imageWidth = w;
    opts.imageHeight = h;

    if (fmt === "svg") {
      opts.coercionType = Office.CoercionType.XmlSvg;
      await officeAsync((cb) => Office.context.document.setSelectedDataAsync(r.svg, opts, cb));
    } else {
      opts.coercionType = Office.CoercionType.Image;
      const b64 = await toPng(r);
      await officeAsync((cb) => Office.context.document.setSelectedDataAsync(b64, opts, cb));
    }

    let newId = null;
    if (S.caps.pp15 && slideId && before) {
      await sleep(60);
      newId = await PowerPoint.run(async (ctx) => {
        const shapes = ctx.presentation.slides.getItem(slideId).shapes;
        shapes.load("items/id,items/name");
        await ctx.sync();
        const fresh = shapes.items.filter((s) => !before.has(s.id));
        if (!fresh.length) return null;
        const sh = fresh[fresh.length - 1];
        sh.tags.add(TAG, JSON.stringify(meta));
        sh.name = "SlideTeX: " + p.src.replace(/\s+/g, " ").slice(0, 60);
        if (S.caps.pp110) {
          try {
            sh.altTextDescription = p.src.slice(0, 2000);
            sh.altTextTitle = "Équation LaTeX";
          } catch (_) {}
        }
        if (old) {
          sh.left = old.left;
          sh.top = old.top;
          sh.width = w;
          sh.height = h;
          const prev = shapes.getItemOrNullObject(old.id);
          prev.load("id");
          await ctx.sync();
          if (!prev.isNullObject) {
            prev.delete();
            await ctx.sync();
          }
          if (S.caps.pp18 && old.z != null) {
            // New shape is on top: send it back to the old z-order slot.
            sh.load("zOrderPosition");
            await ctx.sync();
            let steps = sh.zOrderPosition - old.z;
            while (steps-- > 0) sh.setZOrder("SendBackward");
          }
        }
        await ctx.sync();
        await clearShapeSelection(ctx, slideId);
        return sh.id;
      });
    }
    if (!S.caps.pp13) {
      // No tags on this PowerPoint build: keep the source in the document's settings.
      try {
        const list = Office.context.document.settings.get("slidetex.sources") || [];
        list.unshift({ src: p.src, mode: p.mode, engine: p.engine, t: Date.now() });
        Office.context.document.settings.set("slidetex.sources", list.slice(0, 200));
        Office.context.document.settings.saveAsync(() => {});
      } catch (_) {}
    }
    return { newId, fmt };
  }

  // ------------------------------------------------------------------ actions
  async function primaryAction(forceNew) {
    if (S.busy) return;
    const p = currentParams();
    if (!p.src.trim()) {
      status("Rien à insérer.", "err");
      return;
    }
    setBusy(true);
    try {
      status("Rendu…");
      const r = await render(p);
      showError(null);
      pushHistory(p);
      if (!S.inOffice) {
        status("Rendu OK (hors PowerPoint : utilisez Outils › Télécharger).", "ok");
        return;
      }
      const replace = !forceNew && S.target ? S.target : null;
      status(replace ? "Mise à jour de l'équation…" : "Insertion…");
      const res = await insertRendered(r, p, replace);
      S.dirty = false;
      S.target = null;
      S.loadedFrom = null;
      refreshEditUi();
      status(
        (replace ? "Équation mise à jour" : "Équation insérée") +
          (res.fmt === "png" ? " (PNG)" : " (SVG)") +
          (S.caps.pp13 ? "" : " — édition ultérieure indisponible sur cette version"),
        "ok"
      );
    } catch (e) {
      console.error(e);
      showError(e.message || String(e));
      status("Échec : " + (e.message || e), "err");
    } finally {
      setBusy(false);
    }
  }

  function setBusy(b) {
    S.busy = b;
    $("btnPrimary").disabled = b;
    $("btnInsertNew").disabled = b;
    $("btnRegenAll").disabled = b;
  }

  function loadMetaIntoEditor(meta, shapeId) {
    $("src").value = meta.src;
    S.settings.src = meta.src;
    setSetting("mode", meta.mode || "display", true);
    if (meta.engine === "latex" || meta.engine === "mathjax") setSetting("engine", meta.engine, true);
    if (meta.size) setSetting("size", meta.size, true);
    if (meta.color) setSetting("color", meta.color, true);
    if (typeof meta.preamble === "string" && meta.preamble !== S.settings.preamble && meta.preamble.trim()) {
      setSetting("preamble", meta.preamble, true);
    }
    syncControls();
    S.loadedFrom = shapeId || null;
    S.dirty = false;
    saveSettings();
    updatePreview();
  }

  function refreshEditUi() {
    const editing = !!S.target;
    $("editBanner").hidden = !editing;
    $("btnPrimary").textContent = editing ? "Mettre à jour" : "Insérer";
    $("btnInsertNew").hidden = !editing;
    if (editing) {
      const loaded = S.loadedFrom === S.target.shapeId;
      $("editBannerText").textContent = loaded ? "Modification de l'équation sélectionnée." : "Équation SlideTeX sélectionnée.";
      $("btnLoadSel").hidden = loaded;
    }
  }

  const onSelectionChanged = debounce(async () => {
    if (S.busy || !S.caps.pp15) return;
    try {
      const info = await getSelectionInfo();
      const t = info && info.target;
      S.target = t;
      if (t && (!S.dirty || !$("src").value.trim()) && S.loadedFrom !== t.shapeId) {
        loadMetaIntoEditor(t.meta, t.shapeId);
        status("Équation chargée — modifiez puis « Mettre à jour ».");
      }
      refreshEditUi();
    } catch (e) {
      console.warn("selection", e);
    }
  }, 200);

  async function regenerateAll() {
    if (!S.inOffice || !S.caps.pp15) {
      status("Nécessite PowerPoint Microsoft 365 / 2021+ (PowerPointApi 1.5).", "err");
      return;
    }
    if (S.busy) return;
    setBusy(true);
    const applyStyle = $("regenApplyStyle").checked;
    let done = 0;
    let failed = 0;
    try {
      status("Recherche des équations…");
      const found = await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items/id");
        await ctx.sync();
        const all = slides.items.map((sl) => {
          sl.shapes.load("items/id");
          return sl;
        });
        await ctx.sync();
        const tags = [];
        for (const sl of all) {
          for (const sh of sl.shapes.items) {
            const tg = sh.tags.getItemOrNullObject(TAG);
            tg.load("value");
            tags.push({ slideId: sl.id, shapeId: sh.id, tg });
          }
        }
        await ctx.sync();
        return tags.filter((x) => !x.tg.isNullObject).map((x) => ({ slideId: x.slideId, shapeId: x.shapeId, meta: parseMeta(x.tg.value) })).filter((x) => x.meta);
      });
      if (!found.length) {
        status("Aucune équation SlideTeX dans la présentation.");
        return;
      }
      const cur = currentParams();
      for (let i = 0; i < found.length; i++) {
        const f = found[i];
        status(`Régénération ${i + 1}/${found.length}…`);
        const p = {
          src: f.meta.src,
          mode: f.meta.mode || "display",
          engine: f.meta.engine || "mathjax",
          size: applyStyle ? cur.size : f.meta.size || cur.size,
          color: applyStyle ? cur.color : f.meta.color || cur.color,
          preamble: cur.preamble,
        };
        try {
          const r = await render(p);
          // When the size changes on purpose, do not keep the old user scale.
          const replace = applyStyle ? { ...f, meta: { ...f.meta, w: null } } : f;
          await insertRendered(r, p, replace);
          done++;
        } catch (e) {
          console.warn(e);
          failed++;
        }
      }
      status(`${done} équation(s) régénérée(s)` + (failed ? `, ${failed} échec(s)` : ""), failed ? "err" : "ok");
    } catch (e) {
      status("Échec : " + (e.message || e), "err");
    } finally {
      setBusy(false);
    }
  }

  // ------------------------------------------------------------------ history
  function pushHistory(p) {
    const h = load(LS_HIST, []).filter((x) => x.src !== p.src);
    h.unshift({ src: p.src, mode: p.mode, engine: p.engine });
    save(LS_HIST, h.slice(0, 30));
    renderHistory();
  }
  function renderHistory() {
    const h = load(LS_HIST, []);
    const box = $("history");
    box.innerHTML = "";
    if (!h.length) {
      box.innerHTML = '<span class="muted small">Vide.</span>';
      return;
    }
    h.forEach((item) => {
      const b = document.createElement("button");
      b.textContent = item.src.replace(/\s+/g, " ");
      b.title = item.src;
      b.onclick = () => {
        $("src").value = item.src;
        setSetting("mode", item.mode || "display");
        if (item.engine) setSetting("engine", item.engine);
        syncControls();
        markDirty();
        updatePreview();
      };
      box.appendChild(b);
    });
  }

  // ------------------------------------------------------------------ palette / examples
  function buildPalette() {
    const groups = window.SLIDETEX_SNIPPETS || [];
    const tabs = $("snippetTabs");
    const pal = $("palette");
    const show = (i) => {
      S.settings.tab = i;
      saveSettings();
      Array.from(tabs.children).forEach((b, j) => b.setAttribute("aria-selected", String(i === j)));
      pal.innerHTML = "";
      (groups[i] || groups[0]).items.forEach(([label, ins]) => {
        const b = document.createElement("button");
        b.type = "button";
        b.textContent = label;
        b.title = ins.replace("¦", "");
        b.onclick = () => insertAtCaret(ins);
        pal.appendChild(b);
      });
    };
    groups.forEach((g, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.role = "tab";
      b.textContent = g.tab;
      b.onclick = () => show(i);
      tabs.appendChild(b);
    });
    show(Math.min(S.settings.tab || 0, groups.length - 1));

    const ex = $("examples");
    (window.SLIDETEX_EXAMPLES || []).forEach((e) => {
      const b = document.createElement("button");
      b.innerHTML = "<b></b><span></span>";
      b.firstChild.textContent = e.name;
      b.lastChild.textContent = e.src.replace(/\s+/g, " ");
      b.title = e.src;
      b.onclick = () => {
        $("src").value = e.src;
        setSetting("mode", e.mode);
        if (e.engine) {
          if (e.engine === "latex" && !S.latex.ok) status("Cet exemple nécessite le moteur LaTeX local.", "err");
          setSetting("engine", e.engine);
        }
        if (e.preambleHint && S.settings.engine === "latex" && !S.settings.preamble.includes(e.preambleHint)) {
          setSetting("preamble", (S.settings.preamble ? S.settings.preamble.replace(/\s*$/, "\n") : "") + e.preambleHint);
        }
        syncControls();
        markDirty();
        updatePreview();
      };
      ex.appendChild(b);
    });
  }

  function insertAtCaret(snippet) {
    const ta = $("src");
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    const sel = ta.value.slice(start, end);
    let text = snippet;
    let caret = text.indexOf("¦");
    if (caret >= 0) {
      text = text.replace("¦", sel);
      caret += sel.length;
    } else {
      caret = text.length;
      // Separate a control word from a following letter (\alpha x, not \alphax).
      if (/^\\[a-zA-Z]+$/.test(text) && /^[a-zA-Z]/.test(ta.value.slice(end))) text += " ";
    }
    replaceRange(ta, text, start, end);
    ta.selectionStart = ta.selectionEnd = start + caret;
    ta.focus();
    markDirty();
    onSourceChanged();
  }

  function replaceRange(ta, text, start, end) {
    if (ta.setRangeText) ta.setRangeText(text, start, end, "end");
    else {
      ta.value = ta.value.slice(0, start) + text + ta.value.slice(end); // IE11
      ta.selectionStart = ta.selectionEnd = start + text.length;
    }
  }

  // ------------------------------------------------------------------ controls
  function setSetting(k, v, silent) {
    S.settings[k] = v;
    saveSettings();
    if (!silent && k === "engine") checkLatex();
  }

  function syncControls() {
    const st = S.settings;
    $("mode").value = st.mode;
    $("engine").value = st.engine;
    $("size").value = st.size;
    $("color").value = normHex(st.color) || "#000000";
    $("colorHex").value = normHex(st.color) || "#000000";
    $("format").value = st.format;
    $("dpi").value = st.dpi;
    $("preamble").value = st.preamble;
    $("serverUrl").value = st.serverUrl;
    $("dpi").disabled = st.format !== "png";
  }

  function markDirty() {
    S.dirty = true;
    S.settings.src = $("src").value;
    saveSettings();
  }

  function bindControls() {
    const ta = $("src");
    ta.addEventListener("input", () => {
      markDirty();
      onSourceChanged();
    });
    ta.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        primaryAction(false);
      } else if (e.key === "Tab" && !e.shiftKey) {
        e.preventDefault();
        replaceRange(ta, "  ", ta.selectionStart, ta.selectionEnd);
        markDirty();
      }
    });
    const onChange = (id, key, transform) =>
      $(id).addEventListener("change", () => {
        setSetting(key, transform ? transform($(id).value) : $(id).value);
        syncControls();
        updatePreview();
      });
    onChange("mode", "mode");
    onChange("engine", "engine");
    onChange("format", "format");
    onChange("dpi", "dpi");
    $("size").addEventListener("input", () => {
      const v = parseFloat($("size").value);
      if (v >= 4 && v <= 400) {
        setSetting("size", v);
        schedulePreview();
      }
    });
    $("color").addEventListener("input", () => {
      setSetting("color", $("color").value);
      $("colorHex").value = $("color").value;
      schedulePreview();
    });
    $("colorHex").addEventListener("change", () => {
      const v = normHex($("colorHex").value);
      if (v) {
        setSetting("color", v);
        syncControls();
        updatePreview();
      } else {
        $("colorHex").value = S.settings.color;
      }
    });
    $("preamble").addEventListener("input", () => {
      setSetting("preamble", $("preamble").value);
      onSourceChanged();
    });
    $("serverUrl").addEventListener("change", () => {
      setSetting("serverUrl", $("serverUrl").value.trim());
      S.latex.checked = false;
      checkLatex(true);
    });
    $("btnPrimary").onclick = () => primaryAction(false);
    $("btnInsertNew").onclick = () => primaryAction(true);
    $("btnLoadSel").onclick = () => {
      if (S.target) {
        loadMetaIntoEditor(S.target.meta, S.target.shapeId);
        refreshEditUi();
      }
    };
    $("btnRegenAll").onclick = regenerateAll;
    $("btnCheckServer").onclick = () => {
      S.latex.checked = false;
      checkLatex(true);
    };
    $("btnDownloadSvg").onclick = () => download("svg");
    $("btnDownloadPng").onclick = () => download("png");
  }

  async function download(kind) {
    try {
      const p = currentParams();
      const r = await render(p);
      let blob;
      if (kind === "svg") blob = new Blob([r.svg], { type: "image/svg+xml" });
      else {
        const b64 = await toPng(r);
        const bin = atob(b64);
        const u8 = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
        blob = new Blob([u8], { type: "image/png" });
      }
      if (window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(blob, "equation." + kind); // IE11
        return;
      }
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "equation." + kind;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        document.body.removeChild(a);
      }, 1000);
    } catch (e) {
      showError(e.message || String(e));
    }
  }

  async function checkLatex(verbose) {
    if (S.latex.checked && !verbose) return S.latex.ok;
    S.latex.checked = true;
    try {
      const ctl = new AbortController();
      const t = setTimeout(() => ctl.abort(), 4000);
      const res = await fetch(serverBase() + "/api/health", { signal: ctl.signal });
      clearTimeout(t);
      const j = await res.json();
      S.latex.ok = !!(j.latex && j.dvisvgm);
      S.latex.info = j;
      if (verbose) {
        status(S.latex.ok ? `LaTeX OK — ${j.latexVersion}` : "Serveur joignable mais LaTeX/dvisvgm introuvable (installez TeX Live ou MiKTeX).", S.latex.ok ? "ok" : "err");
      }
    } catch (e) {
      S.latex.ok = false;
      if (verbose) status(`Serveur LaTeX injoignable (${serverBase()}).`, "err");
    }
    const opt = $("engine").querySelector('option[value="latex"]');
    opt.textContent = "LaTeX local (TeX Live / MiKTeX)" + (S.latex.ok ? " ✓" : " — indisponible");
    if (S.settings.engine === "latex" && !S.latex.ok && !verbose) {
      status("LaTeX local indisponible — moteur MathJax conseillé.", "err");
    }
    return S.latex.ok;
  }

  function describeEnv() {
    const b = $("envBadge");
    const c = S.caps;
    if (!S.inOffice) {
      b.textContent = "Aperçu navigateur";
      b.title = "Ouvert hors de PowerPoint : l'insertion est désactivée.";
      $("capsInfo").textContent = "Hors PowerPoint.";
      return;
    }
    let host = "PowerPoint";
    try {
      const d = Office.context.diagnostics;
      host = `PowerPoint ${d.platform === "OfficeOnline" ? "Web" : d.platform === "Mac" ? "Mac" : "Windows"} ${d.version || ""}`.trim();
    } catch (_) {}
    b.textContent = host;
    const lvl = c.pp110 ? "1.10" : c.pp18 ? "1.8" : c.pp15 ? "1.5" : c.pp13 ? "1.3" : "—";
    b.title = `PowerPointApi ${lvl} · SVG ${c.svg ? "oui" : "non"}`;
    $("capsInfo").textContent =
      `API PowerPoint : ${lvl}. Insertion SVG : ${c.svg ? "oui" : "non (PNG utilisé)"}. ` +
      `Édition des équations : ${c.pp15 ? "oui" : c.pp13 ? "partielle" : "non (version trop ancienne : la source est gardée dans les paramètres du document)"}.`;
    if (!c.svg) {
      const o = $("format").querySelector('option[value="svg"]');
      o.disabled = true;
      o.textContent = "SVG (non pris en charge ici)";
      if (S.settings.format === "svg") setSetting("format", "png");
    }
  }

  // ------------------------------------------------------------------ utils
  function normHex(v) {
    if (!v) return null;
    let s = String(v).trim();
    if (!s.startsWith("#")) s = "#" + s;
    if (/^#[0-9a-fA-F]{3}$/.test(s)) s = "#" + s.slice(1).split("").map((c) => c + c).join("");
    return /^#[0-9a-fA-F]{6}$/.test(s) ? s.toLowerCase() : null;
  }
  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const round2 = (n) => Math.round(n * 100) / 100;

  // ------------------------------------------------------------------ boot
  function attachOffice(info) {
    if (!(info && info.host) || S.inOffice) return;
    S.inOffice = true;
    detectCaps();
    if (S.caps.pp15) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
      onSelectionChanged();
    }
    describeEnv();
    status("Prêt.");
  }

  function boot() {
    describeEnv();
    $("src").value = S.settings.src || "";
    syncControls();
    buildPalette();
    bindControls();
    renderHistory();
    refreshEditUi();
    checkLatex(false);
    status("Chargement d'Office…");
    updatePreview();
    $("src").focus();
    if (window.Office && Office.onReady) {
      Office.onReady(attachOffice);
      setTimeout(() => {
        if (!S.inOffice) status("Aperçu hors PowerPoint (insertion désactivée).");
      }, 5000);
    } else {
      status("Aperçu hors PowerPoint (office.js non chargé).");
    }
  }

  window.SlideTeX = { S, render, currentParams, primaryAction, _test: { mathjaxPreamble, normHex } };
  boot();
})();
