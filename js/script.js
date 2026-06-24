/* =========================================================
   PETRA x DANAMON — International Banking Immersion Program
   Behaviour: language toggle, print/download, scroll reveal
   ========================================================= */
(function () {
  "use strict";

  // Flag that JS is active (enables reveal animations; degrades gracefully without JS)
  document.documentElement.classList.add("js");

  var lang = "en";

  /* ---------- Language toggle ---------- */
  var langBtn = document.getElementById("langBtn");
  var printBtns = document.querySelectorAll(".js-print");
  var pptBtns = document.querySelectorAll(".js-ppt");

  function applyLang() {
    document.documentElement.lang = lang;

    var nodes = document.querySelectorAll("[data-en]");
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      var val = el.getAttribute("data-" + lang);
      if (val !== null) el.innerHTML = val;
    }

    // Button labels
    if (langBtn) {
      langBtn.innerHTML = lang === "en" ? "🇮🇩 Bahasa Indonesia" : "🇬🇧 English";
    }
    var labelBtns = [];
    for (var a = 0; a < printBtns.length; a++) labelBtns.push(printBtns[a]);
    for (var b = 0; b < pptBtns.length; b++) labelBtns.push(pptBtns[b]);
    for (var j = 0; j < labelBtns.length; j++) {
      var pb = labelBtns[j];
      if (pb.getAttribute("data-busy") === "1") continue; // don't overwrite progress text
      var keep = pb.getAttribute("data-keep-icon");
      pb.innerHTML = pb.getAttribute("data-" + lang);
      if (keep) pb.innerHTML = keep + " " + pb.innerHTML;
    }
  }

  if (langBtn) {
    langBtn.addEventListener("click", function () {
      lang = lang === "en" ? "id" : "en";
      applyLang();
    });
  }

  /* ---------- Print / Download (Save as PDF) ---------- */
  for (var k = 0; k < printBtns.length; k++) {
    printBtns[k].addEventListener("click", function () {
      window.print();
    });
  }

  /* =========================================================
     Export to PowerPoint — native 16:9 presentation
     Each ".doc .page" section becomes one widescreen slide, rebuilt
     with native PptxGenJS text/shapes/tables/images (fully editable),
     reading content live from the DOM so it follows the language toggle.
     ========================================================= */
  var PPTX_CDN = "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";

  // Brand palette (from CSS :root), PptxGenJS hex without '#'
  var C = {
    blue: "003087", blueDk: "00205a", green: "006633",
    red: "CC0000", orange: "FF6600", yellow: "FFCC00",
    slate: "2B3445", ink: "1A1A2E", sub: "4A4A6A",
    rule: "E0E0E0", light: "F5F5F0", paper: "FAFAF7", white: "FFFFFF"
  };
  var FT = "Arial";        // body
  var FH = "Georgia";      // editorial headings (approximates Playfair)
  var W = 13.33, H = 7.5, MX = 0.62; // slide W/H + side margin (inches)

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src; s.onload = resolve;
      s.onerror = function () { reject(new Error("Failed to load " + src)); };
      document.head.appendChild(s);
    });
  }
  function ensureLibs() {
    if (typeof window.PptxGenJS !== "undefined") return Promise.resolve();
    return loadScript(PPTX_CDN);
  }
  function setBusy(btn, on, text) {
    if (!btn) return;
    if (on) {
      btn.setAttribute("data-busy", "1");
      btn.setAttribute("disabled", "disabled");
      btn.innerHTML = (btn.getAttribute("data-keep-icon") || "") + " " + text;
    } else {
      btn.removeAttribute("data-busy");
      btn.removeAttribute("disabled");
    }
  }

  // ---- DOM read helpers (current language already applied to the DOM) ----
  function txt(el) { return el ? el.textContent.replace(/\s+/g, " ").trim() : ""; }
  function texts(parent, sel) {
    var out = [], n = parent.querySelectorAll(sel);
    for (var i = 0; i < n.length; i++) { var t = txt(n[i]); if (t) out.push(t); }
    return out;
  }
  // Turn a loaded <img> into a PNG data URL (reuses already-decoded image);
  // falls back to its URL path if the canvas is tainted (e.g. file://).
  function imgRef(imgEl) {
    if (!imgEl) return null;
    try {
      var c = document.createElement("canvas");
      c.width = imgEl.naturalWidth || imgEl.width;
      c.height = imgEl.naturalHeight || imgEl.height;
      if (!c.width || !c.height) return { path: imgEl.src };
      c.getContext("2d").drawImage(imgEl, 0, 0);
      return { data: c.toDataURL("image/png") };
    } catch (e) {
      return { path: imgEl.src };
    }
  }
  function addImg(slide, ref, opts) {
    if (!ref) return;
    var o = {};
    for (var k in opts) o[k] = opts[k];
    if (ref.data) o.data = ref.data; else o.path = ref.path;
    slide.addImage(o);
  }

  function pad2(n) { return (n < 10 ? "0" : "") + n; }

  // Per-section editorial themes (accent rail + ghosted numeral tint + soft fill).
  // Indexed by slide number so the deck cycles colour with rhythm, not randomly.
  var THEMES = [
    { accent: C.blue,   ghost: "EAEFF7", soft: "EAF0F8" },
    { accent: C.green,  ghost: "EAF2EC", soft: "EFF5F1" },
    { accent: C.orange, ghost: "FCEEE0", soft: "FFF3E8" },
    { accent: C.red,    ghost: "F9E8E8", soft: "FBEEEE" }
  ];
  function themeFor(num) { return THEMES[((num || 1) - 1) % THEMES.length]; }

  // Photo library — used for the hero cover and editorial accents. Loads from
  // already-decoded DOM <img>s when present, otherwise fetches the asset fresh.
  var PHOTOS = {};
  function preloadPhotos() {
    var want = [
      ["building", "Assets/PETRA_BUILDING.png"],
      ["students", "Assets/PETRA_Students.jpg"],
      ["intl",     "Assets/PETRA_International Students.jpg"],
      ["classroom","Assets/PETRA_Class.jpg"],
      ["painting", "Assets/PETRA_Painting.jpg"],
      ["painting2","Assets/PETRA_PAINTING_2.jpg"]
    ];
    return Promise.all(want.map(function (w) {
      return new Promise(function (res) {
        var file = w[1].replace("Assets/", ""), imgs = document.images, found = null;
        for (var i = 0; i < imgs.length; i++) {
          if (decodeURIComponent(imgs[i].src).indexOf(file) >= 0 && imgs[i].complete && imgs[i].naturalWidth) {
            found = imgs[i]; break;
          }
        }
        if (found) { res([w[0], imgRef(found)]); return; }
        var im = new Image();
        im.onload = function () { res([w[0], imgRef(im)]); };
        im.onerror = function () { res([w[0], null]); };
        im.src = w[1];
      });
    })).then(function (pairs) {
      var o = {};
      for (var i = 0; i < pairs.length; i++) o[pairs[i][0]] = pairs[i][1];
      return o;
    });
  }

  function exportToPpt(btn) {
    var pages = document.querySelectorAll(".doc .page");
    if (!pages.length) return;
    setBusy(btn, true, lang === "en" ? "Building deck…" : "Menyusun…");

    // Reveal-animated elements may be at opacity 0; not needed for native build,
    // but force-show so any read of geometry is stable.
    var hidden = document.querySelectorAll(".reveal:not(.in-view)");
    for (var h = 0; h < hidden.length; h++) hidden[h].classList.add("in-view", "ppt-forced");

    var EN = lang === "en";

    ensureLibs()
      .then(function () { return preloadPhotos(); })
      .then(function (photoRefs) {
        PHOTOS = photoRefs || {};
        var pptx = new window.PptxGenJS();
        pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 (16:9)
        pptx.author = "Petra Christian University";
        pptx.company = "PETRA × DANAMON";
        pptx.title = EN
          ? "International Banking Immersion Program"
          : "Program Imersi Perbankan Internasional";

        // ---------- shared chrome ----------
        function gradBar(s, x, y, w, hh) {
          var seg = w / 3;
          s.addShape(pptx.ShapeType.rect, { x: x, y: y, w: seg, h: hh, fill: { color: C.red } });
          s.addShape(pptx.ShapeType.rect, { x: x + seg, y: y, w: seg, h: hh, fill: { color: C.orange } });
          s.addShape(pptx.ShapeType.rect, { x: x + 2 * seg, y: y, w: w - 2 * seg, h: hh, fill: { color: C.yellow } });
        }
        function footer(s, num) {
          var th = themeFor(num);
          s.addText("PETRA × DANAMON  ·  " + (EN ? "Concept Paper" : "Makalah Konsep"), { x: MX + 0.04, y: 7.08, w: 7, h: 0.3, fontFace: FT, fontSize: 8, color: C.sub, align: "left" });
          if (num) s.addText([
            { text: pad2(num), options: { bold: true, color: th.accent } },
            { text: " / 17", options: { color: C.sub } }
          ], { x: W - MX - 1.6, y: 7.08, w: 1.6, h: 0.3, fontFace: FT, fontSize: 9, align: "right" });
        }
        // Editorial section header: full-height accent rail, oversized ghosted
        // numeral bleeding off the top-right, kicker tile, serif title + rule.
        function header(s, kicker, title, num) {
          var th = themeFor(num);
          s.background = { color: C.paper };
          s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.26, h: H, fill: { color: th.accent } });
          if (num) s.addText(pad2(num), { x: W - 4.3, y: -0.75, w: 4.0, h: 3.1, align: "right", valign: "top", fontFace: FH, bold: true, fontSize: 150, color: th.ghost });
          if (kicker) {
            s.addShape(pptx.ShapeType.rect, { x: MX, y: 0.55, w: 0.15, h: 0.15, fill: { color: th.accent } });
            s.addText(kicker.toUpperCase(), { x: MX + 0.28, y: 0.44, w: W - 2 * MX - 0.28, h: 0.34, fontFace: FT, fontSize: 11, bold: true, color: th.accent, charSpacing: 2 });
          }
          s.addText(title, { x: MX, y: 0.8, w: W - 2 * MX - 0.8, h: 0.72, fontFace: FH, fontSize: 29, bold: true, color: C.blueDk });
          s.addShape(pptx.ShapeType.rect, { x: MX, y: 1.52, w: 1.1, h: 0.06, fill: { color: th.accent } });
        }
        // light card with optional colored accent strip on top + soft drop shadow
        function card(s, x, y, w, hh, stripColor) {
          s.addShape(pptx.ShapeType.roundRect, { x: x, y: y, w: w, h: hh, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.rule, width: 1 }, shadow: { type: "outer", color: "9AA0A8", blur: 5, offset: 2, angle: 90, opacity: 0.32 } });
          if (stripColor) s.addShape(pptx.ShapeType.rect, { x: x, y: y, w: w, h: 0.08, fill: { color: stripColor } });
        }

        var prog = 0, total = pages.length; // 11 doc + 3 rundown + 2 budget + 1 closing = 17
        function tick() { prog++; setBusy(btn, true, (EN ? "Building " : "Menyusun ") + Math.round((prog / total) * 100) + "%"); }

        // ================= SLIDE 1 — COVER (full-bleed photo hero) =================
        (function () {
          var p = pages[0], s = pptx.addSlide();
          s.background = { color: C.blueDk };

          // Hero photo, full bleed, with a layered dark scrim for legibility.
          var hero = PHOTOS.building || PHOTOS.intl || PHOTOS.students;
          addImg(s, hero, { x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
          s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: C.blueDk, transparency: 32 } });
          // bottom darkening band so the headline block reads cleanly
          s.addShape(pptx.ShapeType.rect, { x: 0, y: 3.5, w: W, h: H - 3.5, fill: { color: "001233", transparency: 22 } });
          gradBar(s, 0, 0, W, 0.2);
          gradBar(s, 0, H - 0.2, W, 0.2);

          // logo lockup — two white pills, top-left
          var petra = imgRef(p.querySelector(".logo-img.petra"));
          var dana = imgRef(p.querySelector(".logo-img.danamon"));
          var pillY = 0.7, pillH = 0.86;
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: pillY, w: 1.95, h: pillH, rectRadius: 0.08, fill: { color: C.white }, shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 90, opacity: 0.3 } });
          s.addShape(pptx.ShapeType.roundRect, { x: MX + 2.55, y: pillY, w: 1.95, h: pillH, rectRadius: 0.08, fill: { color: C.white }, shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 90, opacity: 0.3 } });
          addImg(s, petra, { x: MX + 0.15, y: pillY + 0.15, w: 1.65, h: pillH - 0.3, sizing: { type: "contain", w: 1.65, h: pillH - 0.3 } });
          addImg(s, dana, { x: MX + 2.7, y: pillY + 0.15, w: 1.65, h: pillH - 0.3, sizing: { type: "contain", w: 1.65, h: pillH - 0.3 } });
          s.addText("×", { x: MX + 1.95, y: pillY, w: 0.6, h: pillH, align: "center", valign: "middle", fontFace: FH, fontSize: 24, color: C.white });

          // editorial headline block, anchored bottom-left
          s.addShape(pptx.ShapeType.rect, { x: MX, y: 3.72, w: 0.85, h: 0.07, fill: { color: C.yellow } });
          s.addText(EN ? "CONCEPT PAPER · 2026" : "MAKALAH KONSEP · 2026", { x: MX + 0.98, y: 3.62, w: 7, h: 0.3, fontFace: FT, fontSize: 12, bold: true, color: C.yellow, charSpacing: 3, valign: "middle" });
          s.addText(txt(p.querySelector(".ttl1")) || "PETRA × DANAMON", { x: MX - 0.05, y: 3.95, w: 11.4, h: 1.5, align: "left", fontFace: FH, fontSize: 50, bold: true, color: C.white, lineSpacingMultiple: 0.96 });
          s.addText(txt(p.querySelector(".ttl2")), { x: MX, y: 5.42, w: 10.5, h: 0.55, align: "left", fontFace: FH, italic: true, fontSize: 21, color: "DCE6F7" });
          s.addText(txt(p.querySelector(".sub")), { x: MX, y: 5.95, w: 10.5, h: 0.32, align: "left", fontFace: FT, fontSize: 13, color: "C9D6EC", charSpacing: 1 });
          s.addText(txt(p.querySelector(".cover-tag")), { x: MX, y: 6.34, w: 9.8, h: 0.5, align: "left", fontFace: FH, italic: true, fontSize: 13, color: "EDEFF5" });
          s.addText(txt(p.querySelector(".cover-badge")), { x: MX, y: 6.98, w: W - 2 * MX, h: 0.3, align: "left", fontFace: FT, fontSize: 9.5, color: "9FB2D0", charSpacing: 1 });
          tick();
        })();

        // helper: section number from .sec-num-bg or kicker
        function kickerOf(p) { return txt(p.querySelector(".sec-kicker")); }
        function titleOf(p) { return txt(p.querySelector(".sec-title")); }

        // ================= SLIDE 2 — EXECUTIVE SUMMARY =================
        (function () {
          var p = pages[1], s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 2);
          var paras = texts(p, ".body");
          var body = paras.map(function (t, i) {
            return { text: t, options: { breakLine: true, paraSpaceAfter: 9, paraSpaceBefore: i ? 0 : 0 } };
          });
          s.addText(body, { x: MX, y: 1.78, w: 7.35, h: 4.9, valign: "top", fontFace: FT, fontSize: 11.5, color: C.ink, lineSpacingMultiple: 1.02 });

          // pull quote callout
          var pull = txt(p.querySelector(".pull p"));
          s.addShape(pptx.ShapeType.roundRect, { x: 8.25, y: 1.78, w: 4.45, h: 2.55, rectRadius: 0.06, fill: { color: "EFF5F1" }, line: { color: C.green, width: 1 } });
          s.addShape(pptx.ShapeType.rect, { x: 8.25, y: 1.78, w: 0.09, h: 2.55, fill: { color: C.green } });
          s.addText(pull, { x: 8.55, y: 1.95, w: 3.95, h: 2.2, valign: "middle", fontFace: FH, italic: true, fontSize: 15, color: C.green });

          var photo = imgRef(p.querySelector(".doc-photo img"));
          s.addShape(pptx.ShapeType.rect, { x: 8.25, y: 4.5, w: 4.45, h: 0.06, fill: { color: C.orange } });
          addImg(s, photo, { x: 8.25, y: 4.56, w: 4.45, h: 2.18, sizing: { type: "cover", w: 4.45, h: 2.18 } });
          footer(s, 2); tick();
        })();

        // ================= SLIDE 3 — AT A GLANCE =================
        (function () {
          var p = pages[2], s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 3);
          var rows = [];
          var trs = p.querySelectorAll("table.glance tr");
          for (var i = 0; i < trs.length; i++) {
            var tds = trs[i].querySelectorAll("td");
            if (tds.length < 2) continue;
            rows.push([
              { text: txt(tds[0]), options: { bold: true, color: C.green, fill: { color: "F4F8F5" }, valign: "middle" } },
              { text: txt(tds[1]), options: { color: C.ink, valign: "middle" } }
            ]);
          }
          s.addTable(rows, { x: MX, y: 1.78, w: 7.2, colW: [2.5, 4.7], fontFace: FT, fontSize: 10.5, border: { type: "solid", color: C.rule, pt: 1 }, rowH: 0.34, valign: "middle" });

          // theme block (right)
          s.addShape(pptx.ShapeType.roundRect, { x: 8.15, y: 1.78, w: 4.55, h: 4.9, rectRadius: 0.06, fill: { color: C.light }, line: { color: C.rule, width: 1 } });
          s.addShape(pptx.ShapeType.rect, { x: 8.15, y: 1.78, w: 4.55, h: 0.07, fill: { color: C.orange } });
          s.addText(EN ? "PROGRAM THEME" : "TEMA PROGRAM", { x: 8.4, y: 2.0, w: 4.1, h: 0.3, fontFace: FT, fontSize: 10, bold: true, color: C.green, charSpacing: 2 });
          s.addText(txt(p.querySelector(".t-quote")), { x: 8.4, y: 2.35, w: 4.1, h: 1.7, fontFace: FH, italic: true, fontSize: 15, color: C.blueDk, valign: "top" });
          s.addText(txt(p.querySelector(".theme-block .body")), { x: 8.4, y: 4.05, w: 4.1, h: 2.5, fontFace: FT, fontSize: 10.5, color: C.ink, valign: "top", lineSpacingMultiple: 1.03 });
          footer(s, 3); tick();
        })();

        // ================= SLIDE 4 — FOUR PILLARS =================
        (function () {
          var p = pages[3], s = pptx.addSlide();
          header(s, EN ? "Program Overview" : "Gambaran Program", titleOf(p), 4);
          var strips = [C.blue, C.green, C.orange, C.slate];
          var cards = p.querySelectorAll(".pillar-card");
          var gx = MX, gy = 1.72, gw = (W - 2 * MX - 0.4) / 2, gh = 2.42, gapx = 0.4, gapy = 0.34;
          for (var i = 0; i < cards.length && i < 4; i++) {
            var col = i % 2, row = Math.floor(i / 2);
            var x = gx + col * (gw + gapx), y = gy + row * (gh + gapy);
            card(s, x, y, gw, gh, strips[i]);
            s.addText(txt(cards[i].querySelector(".pillar-head")), { x: x + 0.2, y: y + 0.16, w: gw - 0.4, h: 0.32, fontFace: FH, bold: true, fontSize: 14, color: C.blueDk });
            s.addText(txt(cards[i].querySelector(".pillar-sub")), { x: x + 0.2, y: y + 0.5, w: gw - 0.4, h: 0.26, fontFace: FT, fontSize: 9.5, italic: true, color: C.green });
            var lis = texts(cards[i], "li");
            var bl = lis.map(function (t) { return { text: t, options: { bullet: { code: "2022", indent: 12 }, breakLine: true, paraSpaceAfter: 2 } }; });
            if (bl.length) {
              s.addText(bl, { x: x + 0.2, y: y + 0.8, w: gw - 0.4, h: gh - 0.95, valign: "top", fontFace: FT, fontSize: 9, color: C.ink });
            } else {
              s.addText(txt(cards[i].querySelector("p")), { x: x + 0.2, y: y + 0.8, w: gw - 0.4, h: gh - 0.95, valign: "top", fontFace: FT, fontSize: 9.5, color: C.ink });
            }
          }
          footer(s, 4); tick();
        })();

        // ================= SLIDE 5 — COHORT + OUTCOMES =================
        (function () {
          var p = pages[4], s = pptx.addSlide();
          header(s, EN ? "Program Overview" : "Gambaran Program", EN ? "Cohort & Participant Outcomes" : "Komposisi & Hasil Peserta", 5);
          // cohort table (left)
          var rows = [];
          var trs = p.querySelectorAll("table.data tr");
          for (var i = 0; i < trs.length; i++) {
            var cells = trs[i].querySelectorAll("th,td");
            var isHead = trs[i].querySelectorAll("th").length > 0;
            var isTot = trs[i].className.indexOf("tot") >= 0;
            var r = [];
            for (var j = 0; j < cells.length; j++) {
              r.push({ text: txt(cells[j]), options: {
                bold: isHead || isTot || j === 0,
                color: isHead ? C.white : (isTot ? C.blueDk : C.ink),
                fill: { color: isHead ? C.blue : (isTot ? "EAF0F8" : C.white) },
                align: j === 0 ? "left" : "center", valign: "middle"
              } });
            }
            rows.push(r);
          }
          s.addText(EN ? "Cohort Composition" : "Komposisi Peserta", { x: MX, y: 1.72, w: 6, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.green });
          s.addTable(rows, { x: MX, y: 2.06, w: 6.0, colW: [2.8, 1.6, 1.6], fontFace: FT, fontSize: 10, border: { type: "solid", color: C.rule, pt: 1 }, rowH: 0.36, valign: "middle" });

          // outcomes (right)
          s.addText(EN ? "Participant Outcomes" : "Hasil bagi Peserta", { x: 7.0, y: 1.72, w: 5.7, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.green });
          var outs = p.querySelectorAll(".outcome");
          var oy = 2.06;
          for (var k = 0; k < outs.length; k++) {
            var b = txt(outs[k].querySelector("b")), d = txt(outs[k].querySelector("span"));
            s.addShape(pptx.ShapeType.ellipse, { x: 7.0, y: oy + 0.04, w: 0.16, h: 0.16, fill: { color: C.orange } });
            s.addText([
              { text: b + "  ", options: { bold: true, color: C.blueDk } },
              { text: d, options: { color: C.sub } }
            ], { x: 7.28, y: oy - 0.06, w: 5.4, h: 0.78, valign: "top", fontFace: FT, fontSize: 10 });
            oy += 0.9;
          }
          footer(s, 5); tick();
        })();

        // generic "numbered subsections" slide (heading + paragraph blocks)
        function subsecSlide(p, num, opts) {
          opts = opts || {};
          var s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), num);
          var blocks = p.querySelectorAll(opts.sel || ".subsec");
          var cols = opts.cols || 2;
          var rowsN = Math.ceil(blocks.length / cols);
          var gx = MX, gy = 1.74, gw = (W - 2 * MX - (cols - 1) * 0.4) / cols, gapx = 0.4;
          var availH = 6.9 - gy, gh = (availH - (rowsN - 1) * 0.28) / rowsN;
          for (var i = 0; i < blocks.length; i++) {
            var col = i % cols, row = Math.floor(i / cols);
            var x = gx + col * (gw + gapx), y = gy + row * (gh + 0.28);
            s.addShape(pptx.ShapeType.rect, { x: x, y: y + 0.04, w: 0.07, h: gh - 0.08, fill: { color: C.orange } });
            var head = txt(blocks[i].querySelector("h3,h4"));
            var para = txt(blocks[i].querySelector("p"));
            var lab = txt(blocks[i].querySelector(".lab"));
            s.addText((lab ? lab + "  " : "") + head, { x: x + 0.22, y: y, w: gw - 0.3, h: 0.5, fontFace: FH, bold: true, fontSize: opts.headSize || 13, color: C.blueDk, valign: "top" });
            s.addText(para, { x: x + 0.22, y: y + (opts.headGap || 0.52), w: gw - 0.3, h: gh - (opts.headGap || 0.52), fontFace: FT, fontSize: opts.bodySize || 10, color: C.ink, valign: "top", lineSpacingMultiple: 1.02 });
          }
          footer(s, num);
          return s;
        }

        // ================= SLIDE 6 — WHY THIS, WHY NOW =================
        (function () { subsecSlide(pages[5], 6, { cols: 2, headSize: 13, bodySize: 10 }); tick(); })();

        // ================= SLIDE 7 — WHAT DANAMON GETS =================
        (function () { subsecSlide(pages[6], 7, { sel: ".benefit", cols: 3, headSize: 11.5, bodySize: 9, headGap: 0.62 }); tick(); })();

        // ================= SLIDE 8 — DANAMON'S VALUE AT A GLANCE =================
        (function () {
          var p = pages[7], s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 8);
          var vcards = p.querySelectorAll(".value-card");
          var cols = 3, gx = MX, gy = 1.74, gw = (W - 2 * MX - 2 * 0.34) / 3, gh = 1.55, gapx = 0.34, gapy = 0.3;
          for (var i = 0; i < vcards.length && i < 6; i++) {
            var col = i % cols, row = Math.floor(i / cols);
            var x = gx + col * (gw + gapx), y = gy + row * (gh + gapy);
            card(s, x, y, gw, gh, C.green);
            s.addShape(pptx.ShapeType.ellipse, { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, fill: { color: C.blue } });
            s.addText(String(i + 1), { x: x + 0.18, y: y + 0.22, w: 0.42, h: 0.42, align: "center", valign: "middle", fontFace: FT, bold: true, fontSize: 13, color: C.white });
            s.addText(txt(vcards[i].querySelector("h4")), { x: x + 0.72, y: y + 0.2, w: gw - 0.9, h: 0.46, fontFace: FH, bold: true, fontSize: 12.5, color: C.blueDk, valign: "middle" });
            s.addText(txt(vcards[i].querySelector("p")), { x: x + 0.2, y: y + 0.74, w: gw - 0.4, h: gh - 0.86, fontFace: FT, fontSize: 9.5, color: C.ink, valign: "top" });
          }
          // feature box
          var fy = gy + 2 * (gh + gapy);
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: fy, w: W - 2 * MX, h: 0.95, rectRadius: 0.06, fill: { color: C.blueDk } });
          s.addText("★", { x: MX + 0.25, y: fy, w: 0.7, h: 0.95, align: "center", valign: "middle", fontFace: FT, fontSize: 26, color: C.yellow });
          s.addText([
            { text: txt(p.querySelector(".vf-label")) + "  —  ", options: { bold: true, color: C.yellow } },
            { text: txt(p.querySelector(".vf-quote")), options: { color: C.white, italic: true } }
          ], { x: MX + 1.05, y: fy, w: W - 2 * MX - 1.3, h: 0.95, valign: "middle", fontFace: FT, fontSize: 12 });
          footer(s, 8); tick();
        })();

        // ================= SLIDE 9 — WHAT PETRA BRINGS =================
        (function () {
          var s = subsecSlide(pages[8], 9, { cols: 2, headSize: 12.5, bodySize: 9.5 });
          tick();
        })();

        // ================= SLIDE 10 — PARTNERSHIP ASK =================
        (function () {
          var p = pages[9], s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 10);
          var asks = p.querySelectorAll(".ask");
          var gx = MX, gy = 1.74, gw = 7.4, gapy = 0.2;
          var gh = (6.85 - gy - 3 * gapy) / 4;
          for (var i = 0; i < asks.length && i < 4; i++) {
            var y = gy + i * (gh + gapy);
            s.addShape(pptx.ShapeType.roundRect, { x: gx, y: y, w: gw, h: gh, rectRadius: 0.05, fill: { color: C.light }, line: { color: C.rule, width: 1 } });
            s.addShape(pptx.ShapeType.rect, { x: gx, y: y, w: 0.08, h: gh, fill: { color: C.green } });
            var head = txt(asks[i].querySelector("h3"));
            var para = txt(asks[i].querySelector("p"));
            var lis = texts(asks[i], "li");
            if (!para && lis.length) para = lis.join(" · ");
            s.addText(head, { x: gx + 0.22, y: y + 0.1, w: gw - 0.4, h: 0.34, fontFace: FH, bold: true, fontSize: 12, color: C.blueDk });
            s.addText(para, { x: gx + 0.22, y: y + 0.44, w: gw - 0.4, h: gh - 0.5, fontFace: FT, fontSize: 9.3, color: C.ink, valign: "top" });
          }
          // "does NOT need" box (right)
          var nx = gx + gw + 0.35, nw = W - MX - nx;
          s.addShape(pptx.ShapeType.roundRect, { x: nx, y: gy, w: nw, h: 6.85 - gy, rectRadius: 0.06, fill: { color: "FBEEEE" }, line: { color: C.red, width: 1 } });
          s.addText(txt(p.querySelector(".notneed h4")), { x: nx + 0.25, y: gy + 0.18, w: nw - 0.5, h: 0.7, fontFace: FH, bold: true, fontSize: 13, color: C.red, valign: "top" });
          var cks = p.querySelectorAll(".notneed .ck");
          var cy = gy + 1.0;
          for (var j = 0; j < cks.length; j++) {
            var sp = cks[j].querySelector("span:last-child");
            s.addText("✕", { x: nx + 0.25, y: cy, w: 0.3, h: 0.4, fontFace: FT, fontSize: 12, bold: true, color: C.red });
            s.addText(txt(sp), { x: nx + 0.6, y: cy - 0.02, w: nw - 0.85, h: 0.9, fontFace: FT, fontSize: 10, color: C.ink, valign: "top" });
            cy += 1.0;
          }
          footer(s, 10); tick();
        })();

        // ================= SLIDE 11 — PILOT YEAR FRAMING =================
        (function () {
          var p = pages[10], s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 11);
          // hero line
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: 1.74, w: W - 2 * MX, h: 1.15, rectRadius: 0.06, fill: { color: "EAF0F8" } });
          s.addText(txt(p.querySelector(".pilot-hero .tag")) || (EN ? "Year 1 = Pilot" : "Tahun 1 = Percontohan"), { x: MX + 0.25, y: 1.9, w: 2.2, h: 0.4, fontFace: FT, bold: true, fontSize: 12, color: C.white, align: "center", valign: "middle", fill: { color: C.green } });
          s.addText(txt(p.querySelector(".pilot-hero p")), { x: MX + 2.7, y: 1.84, w: W - 2 * MX - 2.95, h: 0.95, fontFace: FT, fontSize: 10.5, color: C.ink, valign: "middle" });

          var pcards = p.querySelectorAll(".pilot-card");
          var cols = 4, gx = MX, gy = 3.15, gw = (W - 2 * MX - 3 * 0.3) / 4, gh = 2.0, gapx = 0.3;
          for (var i = 0; i < pcards.length && i < 4; i++) {
            var x = gx + i * (gw + gapx);
            card(s, x, gy, gw, gh, C.orange);
            s.addText("✓", { x: x + 0.2, y: gy + 0.18, w: 0.5, h: 0.5, fontFace: FT, bold: true, fontSize: 20, color: C.green });
            s.addText(txt(pcards[i].querySelector(".pc-text")), { x: x + 0.2, y: gy + 0.7, w: gw - 0.4, h: gh - 0.85, fontFace: FT, fontSize: 10, color: C.ink, valign: "top" });
          }
          // close line
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: 5.4, w: W - 2 * MX, h: 1.4, rectRadius: 0.06, fill: { color: C.blueDk } });
          s.addShape(pptx.ShapeType.rect, { x: MX, y: 5.4, w: W - 2 * MX, h: 0.07, fill: { color: C.orange } });
          s.addText(txt(p.querySelector(".pilot-close p")), { x: MX + 0.4, y: 5.5, w: W - 2 * MX - 0.8, h: 1.2, fontFace: FH, italic: true, fontSize: 14, color: C.white, valign: "middle", align: "center" });
          footer(s, 11); tick();
        })();

        // ================= NEXT STEPS + CLOSING (built last) =================
        function nextStepsSlide() {
          var cf = document.querySelector(".doc .page .closing-foot");
          var p = cf ? cf.closest(".page") : pages[pages.length - 1];
          var s = pptx.addSlide();
          header(s, kickerOf(p), titleOf(p), 17);
          var items = p.querySelectorAll(".tl-item");
          var n = Math.min(items.length, 5);
          var gx = MX, gy = 1.74, gw = (W - 2 * MX - (n - 1) * 0.25) / n, gapx = 0.25, gh = 3.1;
          for (var i = 0; i < n; i++) {
            var x = gx + i * (gw + gapx);
            card(s, x, gy, gw, gh, C.green);
            s.addShape(pptx.ShapeType.ellipse, { x: x + gw / 2 - 0.28, y: gy + 0.22, w: 0.56, h: 0.56, fill: { color: C.green } });
            s.addText(String(i + 1), { x: x + gw / 2 - 0.28, y: gy + 0.22, w: 0.56, h: 0.56, align: "center", valign: "middle", fontFace: FT, bold: true, fontSize: 16, color: C.white });
            s.addText(txt(items[i].querySelector("h4")), { x: x + 0.15, y: gy + 0.9, w: gw - 0.3, h: 0.7, align: "center", fontFace: FH, bold: true, fontSize: 11, color: C.blueDk, valign: "top" });
            s.addText(txt(items[i].querySelector("p")), { x: x + 0.15, y: gy + 1.6, w: gw - 0.3, h: gh - 1.7, align: "center", fontFace: FT, fontSize: 8.3, color: C.sub, valign: "top" });
          }
          // one ask — photo-backed band with dark scrim for an editorial finish
          var bigAsk = txt(p.querySelector(".oneask .big"));
          var askPhoto = PHOTOS.painting || PHOTOS.classroom || PHOTOS.students;
          if (askPhoto) addImg(s, askPhoto, { x: MX, y: 5.2, w: W - 2 * MX, h: 1.55, sizing: { type: "cover", w: W - 2 * MX, h: 1.55 } });
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: 5.2, w: W - 2 * MX, h: 1.55, rectRadius: 0.06, fill: { color: C.blueDk, transparency: askPhoto ? 18 : 0 }, line: { color: C.blueDk, width: 1 } });
          gradBar(s, MX, 5.2, W - 2 * MX, 0.07);
          s.addText(EN ? "ONE SIMPLE ASK RIGHT NOW" : "SATU PERMINTAAN SEDERHANA SAAT INI", { x: MX + 0.4, y: 5.42, w: W - 2 * MX - 0.8, h: 0.3, fontFace: FT, fontSize: 10, bold: true, color: C.yellow, charSpacing: 2, align: "center" });
          s.addText([
            { text: (EN ? "We ask for " : "Kami meminta "), options: { color: "DCE6F7", fontSize: 16 } },
            { text: bigAsk, options: { color: C.white, bold: true, italic: true, fontSize: 26, fontFace: FH } }
          ], { x: MX + 0.4, y: 5.7, w: W - 2 * MX - 0.8, h: 0.9, align: "center", valign: "middle" });
          footer(s, 17); tick();
        }

        // ================= RUNDOWN SLIDES =================
        function rdDayData(dayEl) {
          var head = dayEl.querySelector(".rd-day-head");
          var dnum = head ? txt(head.querySelector(".rd-d")) : "";
          var theme = head ? txt(head.querySelector(".rd-t")) : "";
          var rows = dayEl.querySelectorAll("tr"), items = [];
          for (var i = 0; i < rows.length; i++) {
            if (rows[i].getAttribute("data-minor") === "1") continue;
            var time = txt(rows[i].querySelector(".rd-time"));
            var actEl = rows[i].querySelector(".rd-act");
            var noteEl = actEl ? actEl.querySelector(".note") : null;
            var note = noteEl ? txt(noteEl) : "";
            var main = txt(actEl);
            if (note) main = main.replace(note, "").trim();
            items.push({ time: time, act: main });
          }
          return { dnum: dnum, theme: theme, items: items };
        }
        function rundownSlide(dayEls, sub, num) {
          var s = pptx.addSlide();
          header(s, EN ? "Program Schedule" : "Jadwal Program", EN ? "Program Rundown" : "Rundown Program", num);
          s.addText(sub, { x: MX, y: 1.5, w: W - 2 * MX, h: 0.3, fontFace: FT, fontSize: 11, italic: true, color: C.sub });
          var n = dayEls.length, gap = 0.32, gx = MX, gy = 1.98, gw = (W - 2 * MX - (n - 1) * gap) / n, colH = 6.9 - gy;
          for (var i = 0; i < n; i++) {
            var d = rdDayData(dayEls[i]);
            var x = gx + i * (gw + gap);
            s.addShape(pptx.ShapeType.roundRect, { x: x, y: gy, w: gw, h: 0.74, rectRadius: 0.05, fill: { color: C.blueDk } });
            s.addShape(pptx.ShapeType.rect, { x: x, y: gy, w: gw, h: 0.06, fill: { color: C.orange } });
            s.addText(d.dnum, { x: x + 0.16, y: gy + 0.08, w: gw - 0.32, h: 0.32, fontFace: FH, bold: true, fontSize: 14, color: C.white });
            s.addText(d.theme, { x: x + 0.16, y: gy + 0.42, w: gw - 0.32, h: 0.3, fontFace: FT, fontSize: 8.3, color: "DCE6F7" });
            s.addShape(pptx.ShapeType.roundRect, { x: x, y: gy + 0.82, w: gw, h: colH - 0.82, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.rule, width: 1 } });
            var bl = [];
            for (var j = 0; j < d.items.length; j++) {
              bl.push({ text: d.items[j].time, options: { bold: true, color: C.green, fontSize: 9, breakLine: true, paraSpaceBefore: j ? 6 : 0 } });
              bl.push({ text: d.items[j].act, options: { color: C.ink, fontSize: 9.5, breakLine: true } });
            }
            s.addText(bl, { x: x + 0.18, y: gy + 0.97, w: gw - 0.36, h: colH - 1.1, valign: "top", fontFace: FT, lineSpacingMultiple: 1.0 });
          }
          footer(s, num);
        }
        var rdDays = document.querySelectorAll(".doc .rd-day");
        if (rdDays.length >= 7) {
          (function () { rundownSlide([rdDays[0], rdDays[1], rdDays[2]], EN ? "Day 0–2 · Arrival, Opening & PETRA Lectures" : "Hari 0–2 · Kedatangan, Pembukaan & Kuliah PETRA", 12); tick(); })();
          (function () { rundownSlide([rdDays[3], rdDays[4]], EN ? "Day 3–4 · Danamon Visit & Cultural Immersion" : "Hari 3–4 · Kunjungan Danamon & Imersi Budaya", 13); tick(); })();
          (function () { rundownSlide([rdDays[5], rdDays[6]], EN ? "Day 5–6 · Capstone, Closing & Departure" : "Hari 5–6 · Capstone, Closing & Keberangkatan", 14); tick(); })();
        }

        // ================= BUDGET SLIDE 16 — RAB AT A GLANCE =================
        (function () {
          var subs = document.querySelectorAll(".doc table.rab tr.sub");
          if (!subs.length) return;
          var s = pptx.addSlide();
          header(s, EN ? "Budget · RAB" : "Anggaran · RAB", EN ? "Budget at a Glance" : "Anggaran Sekilas", 15);
          var cols = 4, gx = MX, gy = 1.78, gw = (W - 2 * MX - 3 * 0.28) / 4, gh = 1.5, gapx = 0.28, gapy = 0.26;
          for (var i = 0; i < subs.length; i++) {
            var col = i % cols, row = Math.floor(i / cols);
            var x = gx + col * (gw + gapx), y = gy + row * (gh + gapy);
            card(s, x, y, gw, gh, C.green);
            s.addText(String.fromCharCode(65 + i), { x: x + 0.16, y: y + 0.14, w: 0.6, h: 0.42, fontFace: FH, bold: true, fontSize: 19, color: C.orange });
            var name = subs[i].getAttribute("data-cat-" + (EN ? "en" : "id")) || "";
            s.addText(name, { x: x + 0.16, y: y + 0.6, w: gw - 0.32, h: 0.55, fontFace: FT, bold: true, fontSize: 9.5, color: C.blueDk, valign: "top" });
            var cells = subs[i].querySelectorAll("td");
            s.addText("Rp " + txt(cells[cells.length - 1]), { x: x + 0.16, y: y + gh - 0.42, w: gw - 0.32, h: 0.3, fontFace: FT, bold: true, fontSize: 11.5, color: C.green });
          }
          var by = gy + 2 * (gh + gapy);
          s.addShape(pptx.ShapeType.roundRect, { x: MX, y: by, w: W - 2 * MX, h: 1.2, rectRadius: 0.06, fill: { color: C.blueDk } });
          gradBar(s, MX, by, W - 2 * MX, 0.07);
          function fig(tr) { var c = document.querySelectorAll(".doc " + tr + " .num"); return c.length ? txt(c[0]) : ""; }
          var seg = (W - 2 * MX) / 3;
          function totCol(x, lbl, val, hot) {
            s.addText(lbl.toUpperCase(), { x: x + 0.3, y: by + 0.22, w: seg - 0.6, h: 0.3, fontFace: FT, fontSize: 9.5, bold: true, color: hot ? C.yellow : "AFC2E0", charSpacing: 1, align: "center" });
            s.addText("Rp " + val, { x: x + 0.3, y: by + 0.52, w: seg - 0.6, h: 0.5, fontFace: FH, bold: true, fontSize: hot ? 21 : 17, color: hot ? C.white : "DCE6F7", align: "center" });
          }
          totCol(MX, EN ? "Program Subtotal" : "Subtotal Program", fig("tr.prog"), false);
          totCol(MX + seg, EN ? "Contingency 10%" : "Contingency 10%", fig("tr.cont"), false);
          totCol(MX + 2 * seg, EN ? "Grand Total" : "Grand Total", fig("tr.grand"), true);
          footer(s, 15); tick();
        })();

        // ================= BUDGET SLIDE 17 — ASSUMPTIONS & COST =================
        (function () {
          var cards = document.querySelectorAll(".doc .bud-summary .bud-card");
          var glance = document.querySelectorAll(".doc .page table.glance");
          var gp = glance.length ? glance[glance.length - 1] : null; // assumptions table (last glance)
          if (!cards.length || !gp) return;
          var s = pptx.addSlide();
          header(s, EN ? "Budget · Assumptions" : "Anggaran · Asumsi", EN ? "Assumptions & Cost per Participant" : "Asumsi & Biaya per Peserta", 16);
          // assumptions table (left)
          var rows = [], trs = gp.querySelectorAll("tr");
          for (var i = 0; i < trs.length; i++) {
            var tds = trs[i].querySelectorAll("td");
            if (tds.length < 2) continue;
            rows.push([
              { text: txt(tds[0]), options: { bold: true, color: C.green, fill: { color: "F4F8F5" }, valign: "middle" } },
              { text: txt(tds[1]), options: { color: C.ink, valign: "middle" } }
            ]);
          }
          s.addText(EN ? "Planning Parameters" : "Parameter Perencanaan", { x: MX, y: 1.72, w: 7, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.green });
          s.addTable(rows, { x: MX, y: 2.06, w: 7.0, colW: [3.4, 3.6], fontFace: FT, fontSize: 9.5, border: { type: "solid", color: C.rule, pt: 1 }, rowH: 0.3, valign: "middle" });
          // cost highlight cards (right) — Grand Total, per JP participant, per participant USD
          var pick = [];
          for (var k = 0; k < cards.length; k++) {
            if (cards[k].className.indexOf("feature") >= 0) pick.push(cards[k]);
          }
          // ensure we show grand total + the two per-participant cards
          var perJP = cards[3], perUSD = cards[4];
          var show = [pick[0] || cards[2], perJP, perUSD];
          var rx = 8.05, rw = W - MX - rx, ry = 1.72, rh = 1.55, rgap = 0.22;
          for (var m = 0; m < show.length; m++) {
            var c = show[m], y = ry + m * (rh + rgap), feat = m === 0;
            s.addShape(pptx.ShapeType.roundRect, { x: rx, y: y, w: rw, h: rh, rectRadius: 0.06, fill: { color: feat ? C.blueDk : C.white }, line: { color: feat ? C.blueDk : C.rule, width: 1 } });
            s.addShape(pptx.ShapeType.rect, { x: rx, y: y, w: rw, h: 0.07, fill: { color: feat ? C.orange : C.green } });
            s.addText(txt(c.querySelector(".bc-lbl")).toUpperCase(), { x: rx + 0.25, y: y + 0.2, w: rw - 0.5, h: 0.3, fontFace: FT, fontSize: 10, bold: true, color: feat ? C.yellow : C.green, charSpacing: 1 });
            s.addText(txt(c.querySelector(".bc-val")), { x: rx + 0.25, y: y + 0.5, w: rw - 0.5, h: 0.5, fontFace: FH, bold: true, fontSize: 24, color: feat ? C.white : C.blueDk });
            s.addText(txt(c.querySelector(".bc-sub")), { x: rx + 0.25, y: y + 1.06, w: rw - 0.5, h: 0.4, fontFace: FT, fontSize: 9.5, color: feat ? "CFE0F5" : C.sub });
          }
          footer(s, 16); tick();
        })();

        // ================= FINAL SLIDE — NEXT STEPS + CLOSING =================
        nextStepsSlide();

        return pptx.writeFile({ fileName: "PETRA-x-DANAMON-Presentation.pptx" });
      })
      .catch(function (err) {
        alert((EN ? "Export failed: " : "Ekspor gagal: ") + (err && err.message ? err.message : err));
      })
      .then(function () {
        var forced = document.querySelectorAll(".ppt-forced");
        for (var f = 0; f < forced.length; f++) forced[f].classList.remove("ppt-forced");
        setBusy(btn, false);
        applyLang();
      });
  }

  for (var pp = 0; pp < pptBtns.length; pp++) {
    pptBtns[pp].addEventListener("click", function () {
      if (this.getAttribute("data-busy") === "1") return;
      exportToPpt(this);
    });
  }

  /* ---------- Smooth scroll for "Read" buttons ---------- */
  var scrollers = document.querySelectorAll("[data-scroll-to]");
  for (var s = 0; s < scrollers.length; s++) {
    scrollers[s].addEventListener("click", function () {
      var target = document.querySelector(this.getAttribute("data-scroll-to"));
      if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  /* ---------- Scroll-reveal via IntersectionObserver ---------- */
  // Give each A4 page a gentle reveal on screen (overridden to visible for print).
  var pages = document.querySelectorAll(".doc .page");
  for (var p = 0; p < pages.length; p++) pages[p].classList.add("reveal");

  var revealEls = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window) {
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("in-view");
            io.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12, rootMargin: "0px 0px -8% 0px" }
    );
    revealEls.forEach(function (el) {
      io.observe(el);
    });
  } else {
    revealEls.forEach(function (el) {
      el.classList.add("in-view");
    });
  }

  /* ---------- Nav shadow on scroll ---------- */
  var nav = document.querySelector(".top-nav");
  function onScroll() {
    if (!nav) return;
    if (window.scrollY > 12) nav.style.boxShadow = "0 4px 18px rgba(0,0,0,.12)";
    else nav.style.boxShadow = "0 1px 6px rgba(0,0,0,.06)";
  }
  window.addEventListener("scroll", onScroll, { passive: true });

  /* ---------- Init ---------- */
  applyLang();
})();
