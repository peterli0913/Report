/* 汇报页的翻页与缩放。无依赖，直接 <script src> 引入即可。
 *
 * 页面按 1280x720 固定尺寸排版，这里按视口算缩放比 —— 和 PPT 一样精确可控，
 * 不会因为窗口大小把版面挤乱。
 */
(function () {
  "use strict";

  var deck = document.querySelector(".deck");
  if (!deck) return;
  var slides = Array.prototype.slice.call(deck.querySelectorAll(".slide"));
  if (!slides.length) return;

  var W = 1280, H = 720;
  var idx = 0;

  // 截图和嵌入场景不需要翻页控件，用 ?hideui=1 关掉
  if (/hideui/.test(location.search) || /hideui/.test(location.hash)) {
    document.body.classList.add("hide-ui");
  }

  function fit() {
    // 打印时交给 @page 处理，不要缩放
    if (window.matchMedia("print").matches) return;
    var s = Math.min(window.innerWidth / W, window.innerHeight / H);
    deck.style.transform = "scale(" + s + ")";
    deck.style.height = H * s + "px";
    // 水平居中：transform-origin 是 top center，靠 margin 撑起垂直居中
    var extra = Math.max(0, (window.innerHeight - H * s) / 2);
    deck.style.marginTop = extra + "px";
  }

  function show(i) {
    idx = Math.max(0, Math.min(slides.length - 1, i));
    slides.forEach(function (el, k) {
      el.classList.toggle("is-active", k === idx);
    });
    var bar = document.querySelector(".progress-top");
    if (bar) bar.style.width = ((idx + 1) / slides.length * 100) + "%";
    var cur = document.querySelector(".nav .cur");
    if (cur) cur.textContent = (idx + 1) + " / " + slides.length;
    if (location.hash !== "#" + (idx + 1)) {
      history.replaceState(null, "", "#" + (idx + 1));
    }
  }

  document.addEventListener("keydown", function (e) {
    if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === "PageDown" ||
        e.key === " " || e.key === "Enter") {
      show(idx + 1); e.preventDefault();
    } else if (e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "PageUp") {
      show(idx - 1); e.preventDefault();
    } else if (e.key === "Home") { show(0); }
    else if (e.key === "End") { show(slides.length - 1); }
    else if (e.key === "f" || e.key === "F") {
      if (document.fullscreenElement) document.exitFullscreen();
      else document.documentElement.requestFullscreen();
    } else if (e.key === "p" || e.key === "P") { window.print(); }
  });

  // 触屏左右滑动
  var x0 = null;
  document.addEventListener("touchstart", function (e) { x0 = e.touches[0].clientX; });
  document.addEventListener("touchend", function (e) {
    if (x0 === null) return;
    var dx = e.changedTouches[0].clientX - x0;
    if (Math.abs(dx) > 50) show(idx + (dx < 0 ? 1 : -1));
    x0 = null;
  });

  document.addEventListener("click", function (e) {
    var t = e.target.closest("[data-go]");
    if (!t) return;
    var v = t.getAttribute("data-go");
    show(v === "next" ? idx + 1 : v === "prev" ? idx - 1 : parseInt(v, 10) - 1);
  });

  window.addEventListener("resize", fit);
  window.addEventListener("beforeprint", function () {
    deck.style.transform = "none";
    deck.style.marginTop = "0";
    slides.forEach(function (el) { el.classList.add("is-active"); });
  });
  window.addEventListener("afterprint", function () { fit(); show(idx); });

  fit();
  var start = parseInt((location.hash || "").slice(1), 10);
  show(isNaN(start) ? 0 : start - 1);
})();
