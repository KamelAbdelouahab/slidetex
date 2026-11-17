/*
 * SlideTeX — SVG helpers.
 * PowerPoint's SVG importer is happiest with plain <path>s, explicit colours and absolute sizes,
 * so every rendered formula goes through normalizeSvg() before insertion.
 */
(function () {
  "use strict";
  const SVG_NS = "http://www.w3.org/2000/svg";
  const XLINK_NS = "http://www.w3.org/1999/xlink";

  function parse(svgText) {
    const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
    const root = doc.documentElement;
    if (!root || root.nodeName.toLowerCase() !== "svg" || doc.getElementsByTagName("parsererror").length) {
      throw new Error("SVG invalide");
    }
    return { doc, root };
  }

  // IE11 (PowerPoint 2016/2019 perpetual) has no Element.remove() nor .children on SVG nodes.
  function removeNode(n) {
    if (n && n.parentNode) n.parentNode.removeChild(n);
  }
  function elementChildren(el) {
    return Array.prototype.filter.call(el.childNodes, (n) => n.nodeType === 1);
  }

  function hrefOf(el) {
    return el.getAttribute("href") || el.getAttributeNS(XLINK_NS, "href") || el.getAttribute("xlink:href") || "";
  }

  /** Replace every <use> by a <g> holding a copy of what it references (recursively). */
  function inlineUses(doc, root) {
    const byId = new Map();
    Array.from(root.querySelectorAll("[id]")).forEach((el) => byId.set(el.getAttribute("id"), el));
    let guard = 0;
    let uses;
    while ((uses = root.getElementsByTagName("use")).length && guard++ < 20) {
      Array.from(uses).forEach((use) => {
        const ref = hrefOf(use);
        const target = ref.startsWith("#") ? byId.get(ref.slice(1)) : null;
        const g = doc.createElementNS(SVG_NS, "g");
        const x = parseFloat(use.getAttribute("x") || "0");
        const y = parseFloat(use.getAttribute("y") || "0");
        const tf = [];
        if (use.getAttribute("transform")) tf.push(use.getAttribute("transform"));
        if (x || y) tf.push(`translate(${x} ${y})`);
        if (tf.length) g.setAttribute("transform", tf.join(" "));
        for (const a of Array.from(use.attributes)) {
          if (["x", "y", "width", "height", "href", "xlink:href", "transform", "id"].includes(a.name)) continue;
          g.setAttribute(a.name, a.value);
        }
        if (target) {
          let clone;
          if (target.nodeName === "symbol") {
            clone = doc.createElementNS(SVG_NS, "g");
            Array.from(target.childNodes).forEach((c) => clone.appendChild(c.cloneNode(true)));
          } else {
            clone = target.cloneNode(true);
          }
          clone.removeAttribute("id");
          clone.querySelectorAll && Array.from(clone.querySelectorAll("[id]")).forEach((n) => n.removeAttribute("id"));
          g.appendChild(clone);
        }
        use.parentNode.replaceChild(g, use);
      });
    }
  }

  /** Drop defs entries that nothing references any more (glyph paths after inlining). */
  function pruneDefs(root) {
    const text = new XMLSerializer().serializeToString(root);
    Array.from(root.querySelectorAll("defs")).forEach((defs) => {
      elementChildren(defs).forEach((child) => {
        const id = child.getAttribute("id");
        if (!id) return;
        const refd = text.indexOf(`url(#${id})`) >= 0 || text.indexOf(`"#${id}"`) >= 0 || text.indexOf(`'#${id}'`) >= 0;
        if (!refd) removeNode(child);
      });
      if (!elementChildren(defs).length) removeNode(defs);
    });
  }

  function stripNoise(root) {
    const all = [root].concat(Array.from(root.querySelectorAll("*")));
    for (const el of all) {
      for (const a of Array.from(el.attributes)) {
        if (a.name.startsWith("data-") || a.name === "role" || a.name === "focusable" || a.name.startsWith("aria-")) {
          el.removeAttribute(a.name);
        }
      }
    }
    Array.from(root.querySelectorAll("title, desc, metadata")).forEach(removeNode);
    // dvisvgm comment nodes
    const comments = [];
    (function walk(n) {
      Array.from(n.childNodes).forEach((c) => (c.nodeType === 8 ? comments.push(c) : c.nodeType === 1 && walk(c)));
    })(root);
    comments.forEach(removeNode);
  }

  function applyColor(root, color) {
    const all = [root].concat(Array.from(root.querySelectorAll("*")));
    for (const el of all) {
      for (const attr of ["fill", "stroke", "color"]) {
        const v = el.getAttribute(attr);
        if (v && v.toLowerCase() === "currentcolor") el.setAttribute(attr, color);
      }
      // MathJax draws glyphs with stroke-width="0": make "no outline" explicit for PowerPoint.
      if (el.getAttribute("stroke-width") === "0") {
        el.setAttribute("stroke", "none");
        el.removeAttribute("stroke-width");
      }
      const st = el.getAttribute("style");
      if (st && /currentcolor/i.test(st)) el.setAttribute("style", st.replace(/currentcolor/gi, color));
    }
    // Elements with no explicit fill inherit black; make the default explicit on the root.
    root.setAttribute("fill", color);
    root.setAttribute("stroke", root.getAttribute("stroke") || "none");
    if (root.getAttribute("stroke") === color) root.setAttribute("stroke", "none");
  }

  /**
   * @param {string} svgText
   * @param {{unitToPt:number, color:string, pad?:number}} o
   *        unitToPt : how many points one viewBox unit is worth after scaling
   *        pad      : padding in viewBox units on each side
   * @returns {{svg:string, widthPt:number, heightPt:number}}
   */
  function normalizeSvg(svgText, o) {
    const { doc, root } = parse(svgText);
    const vb = (root.getAttribute("viewBox") || "").trim().split(/[\s,]+/).map(Number);
    if (vb.length !== 4 || vb.some((n) => !isFinite(n))) throw new Error("SVG sans viewBox");
    const pad = o.pad || 0;
    const [x, y, w, h] = [vb[0] - pad, vb[1] - pad, vb[2] + 2 * pad, vb[3] + 2 * pad];
    inlineUses(doc, root);
    stripNoise(root);
    pruneDefs(root);
    applyColor(root, o.color || "#000000");
    root.removeAttribute("style");
    root.removeAttribute("version");
    root.setAttribute("xmlns", SVG_NS);
    root.setAttribute("viewBox", `${round(x)} ${round(y)} ${round(w)} ${round(h)}`);
    const widthPt = w * o.unitToPt;
    const heightPt = h * o.unitToPt;
    root.setAttribute("width", `${round(widthPt)}pt`);
    root.setAttribute("height", `${round(heightPt)}pt`);
    let out = new XMLSerializer().serializeToString(root);
    if (!/xmlns:xlink=/.test(out) && /xlink:/.test(out)) out = out.replace("<svg ", `<svg xmlns:xlink="${XLINK_NS}" `);
    return { svg: out, widthPt, heightPt };
  }

  function round(n) {
    return Math.round(n * 1000) / 1000;
  }

  /** Rasterise an SVG string to a PNG (base64, no data: prefix). */
  function svgToPngBase64(svg, widthPt, heightPt, dpi) {
    return new Promise((resolve, reject) => {
      const scale = (dpi || 300) / 72;
      const wpx = Math.max(1, Math.ceil(widthPt * scale));
      const hpx = Math.max(1, Math.ceil(heightPt * scale));
      if (wpx * hpx > 60e6) return reject(new Error("Image trop grande pour le PNG — baissez la résolution."));
      const img = new Image();
      img.onload = () => {
        const c = document.createElement("canvas");
        c.width = wpx;
        c.height = hpx;
        const ctx = c.getContext("2d");
        try {
          ctx.drawImage(img, 0, 0, wpx, hpx);
          resolve(c.toDataURL("image/png").split(",")[1]);
        } catch (e) {
          // IE11 taints canvases that drew an SVG: caller falls back to the local server.
          reject(Object.assign(new Error("canvas-tainted"), { tainted: true }));
        }
      };
      img.onerror = () => reject(new Error("Échec de la conversion PNG"));
      // Explicit pixel size so the browser rasterises at full resolution.
      const sized = svg
        .replace(/(<svg[^>]*?)\swidth="[^"]*"/, `$1 width="${wpx}"`)
        .replace(/(<svg[^>]*?)\sheight="[^"]*"/, `$1 height="${hpx}"`);
      img.src = "data:image/svg+xml;base64," + utf8ToBase64(sized);
    });
  }

  function utf8ToBase64(s) {
    const bytes = new TextEncoder().encode(s);
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }

  window.SlideTeXSvg = { normalizeSvg, svgToPngBase64, utf8ToBase64 };
})();
