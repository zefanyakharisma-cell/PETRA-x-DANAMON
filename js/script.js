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
    for (var j = 0; j < printBtns.length; j++) {
      var pb = printBtns[j];
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
