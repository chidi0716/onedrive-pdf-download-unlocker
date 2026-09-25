// slides.js
// Export for PowerPoint files opened in PowerPoint for the web (including
// view-only / "block download" links). The viewer never hands the browser the
// .pptx file, but it does render every slide as ordinary HTML inside this
// frame, so we walk the slides, let each one finish rendering, and ask the
// background script to capture it at 2x (see captureSlide in background.js).
// The captures become an image-based PDF; the slide text is saved alongside.
//
// Runs in every officeapps.live.com frame (all_frames) and only activates in
// the PowerPoint viewer frame.
(function () {
  if (!/powerpoint/i.test(location.hostname)) return;
  if (window.__odpdfSlidesLoaded) return;
  window.__odpdfSlidesLoaded = true;

  const I18N = window.ODPDF_I18N;
  let lang = "en";
  const tr = (k, vars) => {
    let s = I18N ? I18N.t(lang, k) : k;
    for (const [a, b] of Object.entries(vars || {})) s = s.replace("{" + a + "}", b);
    return s;
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const nextFrame = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

  // ---- viewer structure --------------------------------------------------
  // Confirmed against PowerPoint for the web (2026-09): thumbnails are
  // role=option items with a "grid-content-thumbnail-view" class and
  // aria-selected on the current one; every slide has a
  // #PageContentSizeWrapperN container (only the current one has a size);
  // empty placeholders render their "Click to add ..." prompt inside
  // .visiblePromptTextContent; the status bar reads "Slide X of N".

  function viewPanel() {
    return document.getElementById("WACViewPanel");
  }

  // The full editor has replaced the JPEG shown while booting.
  function viewerReady() {
    const boot = document.getElementById("WireframeBootSlideImage");
    if (boot && boot.isConnected && boot.getBoundingClientRect().width > 0) return false;
    return !!viewPanel() && thumbnails().length > 0;
  }

  // Thumbnails currently in the DOM (the list may be virtualized), in order.
  function thumbnails() {
    let items = [...document.querySelectorAll('[role="option"][class*="thumbnail"]')];
    if (!items.length) items = [...document.querySelectorAll('[role="option"][aria-label^="Slide"]')];
    return items.filter((el) => el.getBoundingClientRect().width > 0);
  }

  // {current, total} from the status bar ("Slide 3 of 50"), or null.
  function slidePosition() {
    const re = /(\d+)\s*(?:of|\/|／|之|，共)\s*(\d+)/i;
    for (const el of document.querySelectorAll("[id*=Status] *, [class*=Status] *, [class*=status] *")) {
      if (el.childElementCount) continue;
      const m = re.exec((el.textContent || "").trim());
      if (m && /slide|投影片|幻灯片|スライド/i.test(el.textContent + (el.getAttribute("aria-label") || "") + (el.parentElement ? el.parentElement.textContent : ""))) {
        return { current: +m[1], total: +m[2] };
      }
    }
    return null;
  }

  function wrapperFor(index) {
    return document.getElementById("PageContentSizeWrapper" + index);
  }

  // The white slide page currently shown in the editing area.
  function currentSlideBox() {
    const panel = viewPanel();
    if (!panel) return null;
    const pb = panel.getBoundingClientRect();
    let best = null;
    for (const el of panel.querySelectorAll("div")) {
      const b = el.getBoundingClientRect();
      if (b.width < pb.width * 0.4 || b.height < pb.height * 0.3) continue;
      if (b.right > pb.right + 1 || b.bottom > pb.bottom + 1) continue;
      const ratio = b.width / b.height;
      if (ratio < 1.2 || ratio > 1.9) continue; // 4:3 .. 16:9 (and a bit)
      // smallest qualifying box = the page itself, not its wrappers
      if (!best || b.width * b.height < best.rect.width * best.rect.height) best = { el, rect: b };
    }
    return best;
  }

  // Text of one slide, read from its container (works even when the slide
  // is not the one on screen).
  function slideTextFrom(root) {
    if (!root) return "";
    const lines = [];
    for (const p of root.querySelectorAll(".Paragraph")) {
      if (p.closest(".visiblePromptTextContent")) continue;
      const t = (p.textContent || "").replace(/​/g, "").trim();
      if (t && t !== "•") lines.push(t);
    }
    return lines.join("\n");
  }

  // Hide the "Click to add ..." prompts and empty-placeholder chrome while
  // capturing, so they don't end up in the PDF.
  function hideHints(on) {
    let st = document.getElementById("__odpdf_hide_hints");
    if (on && !st) {
      st = document.createElement("style");
      st.id = "__odpdf_hide_hints";
      st.textContent = ".visiblePromptTextContent, .visiblePromptTextContent * { visibility: hidden !important; }";
      document.head.appendChild(st);
    } else if (!on && st) {
      st.remove();
    }
  }

  // Wait until the slide stops changing: no DOM mutations for `quietMs`,
  // every <img> in it decoded, fonts loaded. Gives up after `maxMs`.
  function waitForStable(el, quietMs, maxMs) {
    return new Promise((resolve) => {
      let timer = null;
      const start = performance.now();
      const done = () => {
        obs.disconnect();
        clearTimeout(timer);
        clearTimeout(cap);
        resolve(performance.now() - start);
      };
      const check = async () => {
        const imgs = [...el.querySelectorAll("img")].filter((i) => !i.complete);
        if (imgs.length) {
          await Promise.race([Promise.all(imgs.map((i) => i.decode().catch(() => {}))), sleep(maxMs)]);
        }
        if (document.fonts && document.fonts.status !== "loaded") await document.fonts.ready;
        await nextFrame();
        done();
      };
      const arm = () => {
        clearTimeout(timer);
        timer = setTimeout(check, quietMs);
      };
      const obs = new MutationObserver(arm);
      obs.observe(el, { subtree: true, childList: true, attributes: true, characterData: true });
      const cap = setTimeout(done, maxMs);
      arm();
    });
  }

  // Bring slide `index` (0-based) on screen using trusted input: click its
  // thumbnail when it's in the DOM, otherwise PageDown from the slide area.
  async function goToSlide(index) {
    const want = index + 1;
    const pos = () => slidePosition();
    const reached = () => { const p = pos(); return p ? p.current === want : false; };
    if (reached()) return true;

    for (let attempt = 0; attempt < 3 && !reached(); attempt++) {
      const thumbs = thumbnails();
      const p = pos();
      let target = null;
      if (index === 0) {
        // jump to the top of the (possibly virtualized) list first
        if (thumbs[0]) thumbs[0].scrollIntoView({ block: "start" });
        await nextFrame();
        target = thumbnails()[0];
      } else if (p) {
        const sel = thumbs.find((t) => t.getAttribute("aria-selected") === "true");
        if (sel && p.current === want - 1) {
          sel.scrollIntoView({ block: "nearest" });
          await nextFrame();
          const list = thumbnails();
          target = list[list.indexOf(list.find((t) => t.getAttribute("aria-selected") === "true")) + 1] || null;
          if (target) target.scrollIntoView({ block: "nearest" });
          await nextFrame();
        }
      }
      if (target) {
        const b = target.getBoundingClientRect();
        await send({ type: "SLIDES_INPUT", input: { kind: "click", x: b.left + b.width / 2, y: b.top + b.height / 2 } });
      } else {
        await send({ type: "SLIDES_INPUT", input: { kind: "key", key: "PageDown" } });
      }
      const start = performance.now();
      while (performance.now() - start < 3000 && !reached()) await sleep(40);
    }
    return reached();
  }

  // ---- export --------------------------------------------------------------

  function fileBase() {
    let name = (document.title || "slides").replace(/\s*[-|–].*$/, "").replace(/\.pptx?$/i, "");
    name = name.replace(/[\\/:*?"<>|]+/g, "_").trim() || "slides";
    return name;
  }

  function saveBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  function b64ToBytes(b64) {
    const s = atob(b64);
    const out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
  }

  function send(msg) {
    return new Promise((resolve) => chrome.runtime.sendMessage(msg, (r) => {
      void chrome.runtime.lastError;
      resolve(r);
    }));
  }

  async function exportSlides(withImages, setStatus) {
    const t0 = performance.now();
    const p0 = slidePosition();
    const n = p0 ? p0.total : thumbnails().length;
    if (!n) throw new Error("no slides found");
    const texts = [];
    const images = [];
    const timings = { nav: 0, render: 0, capture: 0 };

    // Text-only fast path: every slide already has a container in the DOM.
    if (!withImages) {
      let all = true;
      for (let i = 0; i < n; i++) if (!wrapperFor(i)) { all = false; break; }
      if (all) {
        for (let i = 0; i < n; i++) texts.push(`--- Slide ${i + 1} ---\n${slideTextFrom(wrapperFor(i))}`);
        saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), fileBase() + ".txt");
        return { n, secs: ((performance.now() - t0) / 1000).toFixed(1) };
      }
    }

    const begin = await send({ type: "SLIDES_BEGIN" });
    if (!begin || !begin.ok) throw new Error((begin && begin.error) || "could not start capture");
    try {
      if (withImages) hideHints(true);
      for (let i = 0; i < n; i++) {
        setStatus(tr("slidesProgress", { i: i + 1, n }));
        let t = performance.now();
        if (!(await goToSlide(i))) throw new Error("could not open slide " + (i + 1));
        timings.nav += performance.now() - t;
        t = performance.now();
        const wrap = wrapperFor(i) || viewPanel();
        await waitForStable(wrap, 250, 8000);
        timings.render += performance.now() - t;
        texts.push(`--- Slide ${i + 1} ---\n${slideTextFrom(wrapperFor(i))}`);
        if (withImages) {
          const box = currentSlideBox();
          if (!box) throw new Error("slide " + (i + 1) + " not found on screen");
          t = performance.now();
          const r = box.rect;
          const res = await send({ type: "SLIDES_CAPTURE", rect: { x: r.left, y: r.top, width: r.width, height: r.height }, scale: 2 });
          timings.capture += performance.now() - t;
          if (!res || !res.ok) throw new Error((res && res.error) || "capture failed");
          images.push(b64ToBytes(res.data));
        }
      }
    } finally {
      hideHints(false);
      await send({ type: "SLIDES_END" });
    }

    const base = fileBase();
    if (withImages) {
      const pdf = window.ODPDF_PDF.buildPdf(images, 2);
      saveBlob(new Blob([pdf], { type: "application/pdf" }), base + ".pdf");
    }
    saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), base + ".txt");
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    console.log("[odpdf] slide export", JSON.stringify({ slides: n, secs, timingsMs: Object.fromEntries(Object.entries(timings).map(([k, v]) => [k, Math.round(v)])) }));
    return { n, secs };
  }

  // ---- UI ------------------------------------------------------------------

  let bar = null;
  function buildUi() {
    if (bar) return;
    bar = document.createElement("div");
    bar.id = "__odpdf_slides";
    bar.style.cssText =
      "position:fixed;left:12px;bottom:40px;z-index:2147483647;display:flex;gap:6px;align-items:center;" +
      "font:13px/1.2 Segoe UI,system-ui,sans-serif;background:rgba(32,32,32,.92);color:#fff;padding:6px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.3)";
    const mk = (label, title) => {
      const b = document.createElement("button");
      b.textContent = label;
      b.title = title;
      b.style.cssText = "all:unset;cursor:pointer;padding:6px 10px;border-radius:6px;background:#c43e1c;color:#fff;font-weight:600";
      return b;
    };
    const pdfBtn = mk("⬇ " + tr("slidesExportPdf"), tr("slidesExportTitle") + "\n\n" + tr("slidesDebuggerNote"));
    const txtBtn = mk(tr("slidesExportText"), tr("slidesTextTitle"));
    txtBtn.style.background = "#555";
    const status = document.createElement("span");
    status.style.cssText = "padding:0 6px;white-space:nowrap";
    bar.append(pdfBtn, txtBtn, status);
    document.body.appendChild(bar);

    let busy = false;
    const run = async (withImages) => {
      if (busy) return;
      busy = true;
      pdfBtn.style.opacity = txtBtn.style.opacity = ".5";
      status.textContent = tr("slidesPreparing");
      try {
        const { n, secs } = await exportSlides(withImages, (s) => (status.textContent = s));
        status.textContent = tr(withImages ? "slidesDone" : "slidesTextDone", { n, s: secs });
      } catch (e) {
        status.textContent = tr("slidesFailed") + (e && e.message ? e.message : e);
      } finally {
        busy = false;
        pdfBtn.style.opacity = txtBtn.style.opacity = "1";
      }
    };
    pdfBtn.addEventListener("click", () => run(true));
    txtBtn.addEventListener("click", () => run(false));
  }

  function start() {
    const poll = setInterval(() => {
      if (viewerReady()) {
        clearInterval(poll);
        buildUi();
      }
    }, 1000);
  }

  if (I18N) I18N.getLang((l) => { lang = l; start(); });
  else start();
})();
