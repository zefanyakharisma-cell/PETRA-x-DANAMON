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

  /* ---------- Export to PowerPoint (1 slide per section/page) ---------- */
  // Each ".doc .page" (A4) becomes one slide. We rasterise each page with
  // html2canvas and place it full-bleed on an A4-portrait slide via PptxGenJS,
  // preserving the exact design and the currently selected language.
  var CDN = {
    html2canvas: "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js",
    pptxgen: "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"
  };

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src;
      s.onload = resolve;
      s.onerror = function () { reject(new Error("Failed to load " + src)); };
      document.head.appendChild(s);
    });
  }

  function ensureLibs() {
    var chain = Promise.resolve();
    if (typeof window.html2canvas === "undefined") {
      chain = chain.then(function () { return loadScript(CDN.html2canvas); });
    }
    if (typeof window.PptxGenJS === "undefined") {
      chain = chain.then(function () { return loadScript(CDN.pptxgen); });
    }
    return chain;
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

  function exportToPpt(btn) {
    var pages = document.querySelectorAll(".doc .page");
    if (!pages.length) return;

    setBusy(btn, true, lang === "en" ? "Preparing…" : "Menyiapkan…");

    // Force every reveal element visible so off-screen pages render fully.
    var hidden = document.querySelectorAll(".reveal:not(.in-view)");
    for (var h = 0; h < hidden.length; h++) hidden[h].classList.add("in-view", "ppt-forced");

    ensureLibs()
      .then(function () {
        var pptx = new window.PptxGenJS();
        // A4 portrait in inches (210 × 297 mm).
        pptx.defineLayout({ name: "A4P", width: 8.27, height: 11.69 });
        pptx.layout = "A4P";

        var slideW = 8.27, slideH = 11.69;

        function renderPage(i) {
          if (i >= pages.length) {
            return pptx.writeFile({ fileName: "PETRA-x-DANAMON-Concept-Paper.pptx" });
          }
          var pct = Math.round((i / pages.length) * 100);
          setBusy(btn, true, (lang === "en" ? "Rendering " : "Membuat ") + pct + "%");
          return window
            .html2canvas(pages[i], {
              scale: 2,
              useCORS: true,
              backgroundColor: "#ffffff",
              logging: false,
              windowWidth: document.documentElement.scrollWidth
            })
            .then(function (canvas) {
              var img = canvas.toDataURL("image/jpeg", 0.92);
              var slide = pptx.addSlide();
              slide.background = { color: "FFFFFF" };
              slide.addImage({ data: img, x: 0, y: 0, w: slideW, h: slideH });
              return renderPage(i + 1);
            });
        }

        return renderPage(0);
      })
      .catch(function (err) {
        alert((lang === "en" ? "Export failed: " : "Ekspor gagal: ") + err.message);
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
