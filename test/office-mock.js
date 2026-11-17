/*
 * Minimal in-memory mock of Office.js + PowerPoint JS API, enough to exercise SlideTeX's flows
 * (insert, select, edit/update, regenerate) in a headless browser. Not shipped with the add-in.
 */
(function () {
  "use strict";
  const PP_LEVEL = window.__MOCK_PP_LEVEL__ || "1.10";
  const SVG_OK = window.__MOCK_SVG__ !== false;
  let nextId = 100;
  const doc = {
    slides: [{ id: "256#", shapes: [] }, { id: "257#", shapes: [] }],
    activeSlide: 0,
    selectedShapeIds: [],
    log: [],
  };
  window.__mockDoc = doc;
  window.__mockFireSel = () => fireSel();
  const handlers = [];

  function cmpVer(a, b) {
    const pa = a.split(".").map(Number), pb = b.split(".").map(Number);
    for (let i = 0; i < 2; i++) if ((pa[i] || 0) !== (pb[i] || 0)) return (pa[i] || 0) - (pb[i] || 0);
    return 0;
  }

  function fireSel() {
    handlers.forEach((h) => setTimeout(() => h({ type: "documentSelectionChanged" }), 0));
  }

  window.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Image: "image", XmlSvg: "xmlSvg", Text: "text" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    onReady(cb) {
      setTimeout(() => cb({ host: "PowerPoint", platform: "PC" }), 50);
      return Promise.resolve({ host: "PowerPoint", platform: "PC" });
    },
    context: {
      diagnostics: { platform: "PC", version: "16.0.mock" },
      requirements: {
        isSetSupported(name, v) {
          if (name === "PowerPointApi") return cmpVer(PP_LEVEL, v || "1.1") >= 0;
          if (name === "ImageCoercion") return v === "1.1" || (SVG_OK && v === "1.2");
          return false;
        },
      },
      document: {
        settings: {
          _d: {},
          get(k) { return this._d[k]; },
          set(k, v) { this._d[k] = v; },
          saveAsync(cb) { cb && cb({ status: "succeeded" }); },
        },
        addHandlerAsync(type, h, cb) {
          handlers.push(h);
          cb && cb({ status: "succeeded" });
        },
        setSelectedDataAsync(data, opts, cb) {
          setTimeout(() => {
            const slide = doc.slides[doc.activeSlide];
            const sh = {
              id: String(nextId++),
              name: "Picture " + nextId,
              type: opts.coercionType === "xmlSvg" ? "Graphic" : "Image",
              left: opts.imageLeft != null ? opts.imageLeft : 100,
              top: opts.imageTop != null ? opts.imageTop : 100,
              width: opts.imageWidth,
              height: opts.imageHeight,
              tags: {},
              data,
              coercion: opts.coercionType,
              altTextDescription: "",
              altTextTitle: "",
            };
            slide.shapes.push(sh);
            doc.selectedShapeIds = [sh.id];
            doc.log.push({ op: "insert", id: sh.id, slide: slide.id, coercion: opts.coercionType });
            cb({ status: "succeeded", value: null });
            fireSel();
          }, 10);
        },
      },
    },
  };

  // ----------------------------------------------------------------- PowerPoint JS API mock
  // Proxy objects record operations; ctx.sync() executes them and fills in loaded values.
  function findShape(id) {
    for (const s of doc.slides) {
      const sh = s.shapes.find((x) => x.id === id);
      if (sh) return { slide: s, sh };
    }
    return null;
  }

  class Ctx {
    constructor() {
      this.q = [];
      const ctx = this;
      this.presentation = {
        getSelectedShapes: () => ctx._collection(() => doc.selectedShapeIds.map((id) => findShape(id)).filter(Boolean).map((x) => x.sh)),
        getSelectedSlides: () => ctx._slideCollection(() => [doc.slides[doc.activeSlide]]),
        setSelectedSlides: (ids) => ctx.q.push(() => {
          const i = doc.slides.findIndex((s) => s.id === ids[0]);
          if (i >= 0) doc.activeSlide = i;
          doc.selectedShapeIds = [];
        }),
        slides: null,
      };
      this.presentation.slides = ctx._slideCollection(() => doc.slides);
      this.presentation.slides.getItem = (id) => ctx._slide(() => doc.slides.find((s) => s.id === id));
    }
    _slide(getter) {
      const ctx = this;
      const o = { load() { ctx.q.push(() => { const s = getter(); o.id = s.id; }); return o; } };
      o.shapes = ctx._collection(() => getter().shapes);
      // Real API: Slide.setSelectedShapes (PowerPointApi 1.5). Presentation has no such method.
      o.setSelectedShapes = (ids) => ctx.q.push(() => {
        doc.selectedShapeIds = ids.slice();
        fireSel();
      });
      return o;
    }
    _slideCollection(getter) {
      const ctx = this;
      const col = {
        items: [],
        load() {
          ctx.q.push(() => {
            col.items = getter().map((s) => {
              const so = ctx._slide(() => s);
              so.id = s.id;
              return so;
            });
          });
          return col;
        },
      };
      return col;
    }
    _shape(getter) {
      const ctx = this;
      const o = { isNullObject: false };
      const fill = () => {
        const sh = getter();
        if (!sh) { o.isNullObject = true; return null; }
        Object.assign(o, { id: sh.id, name: sh.name, left: sh.left, top: sh.top, width: sh.width, height: sh.height, type: sh.type });
        const sl = findShape(sh.id);
        o.zOrderPosition = sl ? sl.slide.shapes.indexOf(sh) : 0;
        return sh;
      };
      o.load = () => { ctx.q.push(fill); return o; };
      o.delete = () => ctx.q.push(() => {
        const sh = getter();
        const f = sh && findShape(sh.id);
        if (f) {
          f.slide.shapes.splice(f.slide.shapes.indexOf(sh), 1);
          doc.log.push({ op: "delete", id: sh.id });
        }
      });
      o.setZOrder = (pos) => ctx.q.push(() => {
        const sh = getter();
        const f = findShape(sh.id);
        const arr = f.slide.shapes;
        const i = arr.indexOf(sh);
        if (pos === "SendBackward" && i > 0) { arr.splice(i, 1); arr.splice(i - 1, 0, sh); }
      });
      o.tags = {
        add: (k, v) => ctx.q.push(() => { getter().tags[k.toUpperCase()] = v; }),
        getItemOrNullObject: (k) => {
          const t = { isNullObject: false, value: undefined };
          t.load = () => {
            ctx.q.push(() => {
              const sh = getter();
              const v = sh && sh.tags[k.toUpperCase()];
              t.isNullObject = v === undefined;
              t.value = v;
            });
            return t;
          };
          return t;
        },
      };
      // Property writes are recorded through setters.
      return new Proxy(o, {
        set(target, prop, value) {
          if (["left", "top", "width", "height", "name", "altTextDescription", "altTextTitle"].includes(prop)) {
            ctx.q.push(() => { getter()[prop] = value; });
          }
          target[prop] = value;
          return true;
        },
      });
    }
    _collection(getter) {
      const ctx = this;
      const col = {
        items: [],
        load() {
          ctx.q.push(() => {
            col.items = getter().map((sh) => {
              const p = ctx._shape(() => sh);
              Object.assign(p, {}); // keep proxy
              const raw = sh;
              p.__raw = raw;
              return p;
            });
            col.items.forEach((p) => {
              const sh = p.__raw;
              // fill without queueing setters
              Object.defineProperty(p, "id", { value: sh.id, writable: true, configurable: true });
              Object.defineProperty(p, "name", { value: sh.name, writable: true, configurable: true });
            });
          });
          return col;
        },
        getItem: (id) => ctx._shape(() => getter().find((s) => s.id === id)),
        getItemOrNullObject: (id) => ctx._shape(() => getter().find((s) => s.id === id)),
      };
      return col;
    }
    async sync() {
      await new Promise((r) => setTimeout(r, 1));
      const q = this.q;
      this.q = [];
      q.forEach((f) => f());
    }
  }

  window.PowerPoint = {
    run: async (fn) => {
      const ctx = new Ctx();
      const r = await fn(ctx);
      await ctx.sync();
      return r;
    },
  };
})();
