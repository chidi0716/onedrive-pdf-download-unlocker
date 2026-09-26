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
  // Visible state marker (also handy when troubleshooting a user's report
  // from DevTools): waiting -> ready.
  document.documentElement.setAttribute("data-odpdf-slides", "waiting");

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

  // {current, total} from the status bar ("Slide 3 of 50"), or null. Found by
  // text rather than by element, since the status bar re-renders freely.
  const POSITION_RE = /(?:Slide|投影片|幻灯片|スライド)\s*(\d+)\s*(?:of|\/|／|，共|共)\s*(\d+)/i;
  function slidePosition() {
    // Thumbnails carry aria-posinset (1-based) and a title like
    // "Title, Slide 2 of 4": the most reliable source, language-independent
    // for the position.
    const thumbs = thumbnails();
    const sel = thumbs.find((t) => t.getAttribute("aria-selected") === "true");
    if (sel && +sel.getAttribute("aria-posinset") > 0) {
      let total = 0;
      for (const t of thumbs) {
        const m = /(\d+)\D+(\d+)\s*$/.exec(t.getAttribute("title") || "");
        if (m) { total = +m[2]; break; }
      }
      const fromStatus = total ? null : statusPosition();
      return { current: +sel.getAttribute("aria-posinset"), total: total || (fromStatus && fromStatus.total) || thumbs.length };
    }
    return statusPosition();
  }

  function statusPosition() {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const m = POSITION_RE.exec(node.nodeValue);
      if (m && node.parentElement && node.parentElement.getClientRects().length) return { current: +m[1], total: +m[2] };
    }
    // fall back to the status bar container's combined text
    for (const el of document.querySelectorAll("[id*=Status], [class*=Status]")) {
      const m = POSITION_RE.exec(el.textContent || "");
      if (m) return { current: +m[1], total: +m[2] };
    }
    return null;
  }

  function wrapperFor(index) {
    return document.getElementById("PageContentSizeWrapper" + index);
  }

  // The slide currently shown. Each slide's container holds an <svg> exactly
  // the size of the slide (its background); use the largest one in the
  // visible container. Falls back to the smallest slide-shaped box.
  function currentSlideBox() {
    const panel = viewPanel();
    if (!panel) return null;
    const wrap = [...document.querySelectorAll("[id^=PageContentSizeWrapper]")].find((w) => w.getBoundingClientRect().width > 0);
    if (wrap) {
      let best = null;
      for (const svg of wrap.querySelectorAll("svg")) {
        const b = svg.getBoundingClientRect();
        const ratio = b.height ? b.width / b.height : 0;
        if (b.width < 200 || ratio < 1.2 || ratio > 1.9) continue;
        if (!best || b.width * b.height > best.rect.width * best.rect.height) best = { el: svg, rect: b };
      }
      if (best) return best;
    }
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
      // "Click to add ..." prompt text, and the selection overlay layer that
      // draws the dashed outline around empty placeholders (editing chrome,
      // never part of the slide itself).
      st.textContent =
        ".visiblePromptTextContent, .visiblePromptTextContent * { visibility: hidden !important; }" +
        ".ShapeSelectionOverlay { display: none !important; }";
      document.head.appendChild(st);
    } else if (!on && st) {
      st.remove();
    }
  }

  // Last time the frame fetched anything (pictures and SmartArt graphics are
  // loaded after the slide's text, so DOM quiet alone isn't enough).
  let lastResourceAt = performance.now();
  try {
    new PerformanceObserver((list) => {
      for (const e of list.getEntries()) {
        // background telemetry pings would otherwise never let us settle
        if (/RemoteUls|Telemetry|OneCollector|events\.data\.microsoft|\/collect\b|keepalive/i.test(e.name)) continue;
        lastResourceAt = Math.max(lastResourceAt, e.responseEnd || performance.now());
      }
    }).observe({ type: "resource", buffered: false });
  } catch (e) {
    // PerformanceObserver unavailable: DOM quiet only
  }

  // Wait until the slide stops changing: no DOM mutations and no network
  // fetches for `quietMs`, images decoded, fonts loaded. Gives up after `maxMs`.
  function waitForStable(el, quietMs, maxMs) {
    return new Promise((resolve) => {
      let timer = null;
      let finished = false;
      const start = performance.now();
      const done = () => {
        if (finished) return;
        finished = true;
        obs.disconnect();
        clearTimeout(timer);
        clearTimeout(cap);
        resolve(performance.now() - start);
      };
      const check = async () => {
        const sinceFetch = performance.now() - lastResourceAt;
        if (sinceFetch < quietMs + 200) {
          timer = setTimeout(check, quietMs + 200 - sinceFetch);
          return;
        }
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

  const sameRect = (a, b) => !!a && !!b && ["left", "top", "width", "height"].every((p) => Math.abs(a[p] - b[p]) < 0.5);

  // Wait until the slide hasn't moved for `quietMs` (the viewer shifts its
  // layout a few seconds after opening, when the real header replaces the
  // loading skeleton). Gives up after `maxMs`.
  async function waitLayoutQuiet(quietMs, maxMs) {
    const start = performance.now();
    let last = null;
    let since = performance.now();
    while (performance.now() - start < maxMs) {
      const box = currentSlideBox();
      const r = box && box.rect;
      if (!sameRect(r, last)) {
        last = r;
        since = performance.now();
      } else if (r && performance.now() - since >= quietMs) {
        return;
      }
      await sleep(100);
    }
  }

  // The slide box once its position has stopped moving (layout can shift,
  // e.g. when Chrome shows the "is debugging" bar).
  async function settledSlideBox() {
    let prev = null;
    for (let k = 0; k < 30; k++) {
      const box = currentSlideBox();
      if (box && prev && sameRect(box.rect, prev)) return box;
      prev = box && box.rect;
      await sleep(100);
    }
    return currentSlideBox();
  }

  function thumbByNumber(n) {
    return thumbnails().find((t) => +t.getAttribute("aria-posinset") === n) || null;
  }

  // The thumbnail list may only render the items near the viewport; scroll
  // it proportionally so thumbnail `n` of `total` gets rendered.
  async function revealThumb(n, total) {
    let t = thumbByNumber(n);
    if (t) return t;
    const any = thumbnails()[0];
    if (!any) return null;
    let sc = any.parentElement;
    while (sc && sc.scrollHeight <= sc.clientHeight + 2) sc = sc.parentElement;
    if (!sc) return null;
    for (let k = 0; k < 6 && !t; k++) {
      sc.scrollTop = Math.max(0, ((n - 1) / Math.max(1, total)) * sc.scrollHeight - sc.clientHeight / 2);
      await nextFrame();
      await sleep(60);
      t = thumbByNumber(n);
      if (!t) {
        // nudge toward it using the numbers that are rendered
        const nums = thumbnails().map((x) => +x.getAttribute("aria-posinset")).filter(Boolean);
        if (nums.length && n < Math.min(...nums)) sc.scrollTop -= sc.clientHeight / 2;
        else if (nums.length && n > Math.max(...nums)) sc.scrollTop += sc.clientHeight / 2;
        await nextFrame();
        t = thumbByNumber(n);
      }
    }
    return t;
  }

  // Bring slide `index` (0-based) on screen with a trusted click on its
  // thumbnail (found by aria-posinset), falling back to PageDown/PageUp.
  async function goToSlide(index, total) {
    const want = index + 1;
    const current = () => { const p = slidePosition(); return p ? p.current : null; };
    // Arrived when the status/thumbnail position reads `want`, OR the wanted
    // thumbnail itself is now selected (updates immediately on a real click).
    const arrived = () => {
      if (current() === want) return true;
      const t = thumbByNumber(want);
      return !!(t && t.getAttribute("aria-selected") === "true");
    };
    const tries = [];
    for (let step = 0; step < 10; step++) {
      if (arrived()) return true;
      const cur = current();
      // Alternate between clicking the thumbnail and keyboard stepping, so a
      // method that isn't taking effect doesn't stall the whole navigation.
      let input = null;
      const useClick = step % 3 !== 2; // click, click, key, click, click, key...
      if (useClick) {
        const t = await revealThumb(want, total);
        if (t) {
          t.scrollIntoView({ block: "center" });
          await sleep(150);
          const b = t.getBoundingClientRect();
          if (b.width > 4 && b.height > 4) input = { kind: "click", x: b.left + b.width / 2, y: b.top + b.height / 2 };
        }
      }
      if (!input) {
        // step toward the target one slide at a time
        const sel = thumbnails().find((x) => x.getAttribute("aria-selected") === "true");
        if (sel) { sel.scrollIntoView({ block: "center" }); await sleep(100); }
        input = { kind: "key", key: cur !== null && cur > want ? "ArrowUp" : "ArrowDown" };
      }
      tries.push(input.kind === "click" ? "click" : input.key);
      await send({ type: "SLIDES_INPUT", input });
      const start = performance.now();
      while (performance.now() - start < 4500 && !arrived()) await sleep(60);
    }
    if (arrived()) return true;
    const p = slidePosition();
    throw new Error(`could not open slide ${want} (at ${p ? p.current + "/" + p.total : "?"}, thumbs ${thumbnails().length}, tried ${tries.join(",")})`);
  }

  // ---- export --------------------------------------------------------------

  async function fileBase() {
    const r = await send({ type: "SLIDES_TAB_TITLE" });
    let name = ((r && r.title) || document.title || "slides").replace(/\s*[-|–]\s*(PowerPoint|SharePoint|OneDrive).*$/i, "").replace(/\.pptx?$/i, "");
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

  // Every page is rendered to this width regardless of how big the slide
  // happens to be on screen, so pages come out the same size and sharp.
  // 960 pt is the width of a standard 16:9 slide.
  const OUTPUT_WIDTH_PX = 2000;
  const PAGE_WIDTH_PT = 960;

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
        saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), (await fileBase()) + ".txt");
        return { n, secs: ((performance.now() - t0) / 1000).toFixed(1) };
      }
    }

    const begin = await send({ type: "SLIDES_BEGIN" });
    if (!begin || !begin.ok) throw new Error((begin && begin.error) || "could not start capture");
    // Attaching the debugger shows an infobar that resizes the page; let the
    // viewer re-layout before the first slide.
    await sleep(400);
    await waitForStable(viewPanel(), 300, 3000);
    await waitLayoutQuiet(1500, 10000);
    try {
      if (withImages) hideHints(true);
      for (let i = 0; i < n; i++) {
        setStatus(tr("slidesProgress", { i: i + 1, n }));
        let t = performance.now();
        await goToSlide(i, n);
        timings.nav += performance.now() - t;
        t = performance.now();
        // Watch the whole view panel, not just the slide: loading overlays
        // (e.g. while a SmartArt graphic renders) live outside the slide.
        await waitForStable(viewPanel(), 300, 5000);
        timings.render += performance.now() - t;
        texts.push(`--- Slide ${i + 1} ---\n${slideTextFrom(wrapperFor(i))}`);
        if (withImages) {
          // park the mouse outside the slide so no hover tooltip is captured
          const pb = viewPanel().getBoundingClientRect();
          await send({ type: "SLIDES_INPUT", input: { kind: "move", x: pb.right - 4, y: pb.bottom - 4 } });
          if (bar) bar.style.visibility = "hidden";
          let res = null;
          t = performance.now();
          // Re-measure after capturing: if the layout moved meanwhile, the
          // capture is off, so take it again.
          for (let attempt = 0; attempt < 3; attempt++) {
            const box = await settledSlideBox();
            if (!box) throw new Error("slide " + (i + 1) + " not found on screen");
            const r = box.rect;
            res = await send({ type: "SLIDES_CAPTURE", rect: { x: r.left, y: r.top, width: r.width, height: r.height }, scale: OUTPUT_WIDTH_PX / r.width });
            if (!res || !res.ok) throw new Error((res && res.error) || "capture failed");
            const after = currentSlideBox();
            if (after && sameRect(after.rect, r)) break;
          }
          timings.capture += performance.now() - t;
          if (bar) bar.style.visibility = "";
          images.push(b64ToBytes(res.data));
        }
      }
    } finally {
      hideHints(false);
      await send({ type: "SLIDES_END" });
    }

    const base = await fileBase();
    if (withImages) {
      const pdf = window.ODPDF_PDF.buildPdf(images, OUTPUT_WIDTH_PX / PAGE_WIDTH_PT);
      saveBlob(new Blob([pdf], { type: "application/pdf" }), base + ".pdf");
    }
    saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), base + ".txt");
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    console.log("[odpdf] slide export", JSON.stringify({ slides: n, secs, timingsMs: Object.fromEntries(Object.entries(timings).map(([k, v]) => [k, Math.round(v)])) }));
    return { n, secs };
  }

  // ---- UI ------------------------------------------------------------------

  // The export is started from the extension popup: the viewer grabs pointer
  // events for itself, so buttons injected into it never receive clicks. A
  // small bar in the viewer shows progress and the result.
  let bar = null;
  let statusEl = null;
  function setStatus(text) {
    if (!bar) {
      bar = document.createElement("div");
      bar.id = "__odpdf_slides";
      bar.style.cssText =
        "position:fixed;left:12px;bottom:40px;z-index:2147483647;pointer-events:none;" +
        "font:13px/1.3 Segoe UI,system-ui,sans-serif;background:rgba(32,32,32,.92);color:#fff;padding:8px 12px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.3)";
      statusEl = document.createElement("span");
      bar.appendChild(statusEl);
      document.body.appendChild(bar);
    }
    statusEl.textContent = text;
    // Also readable from DevTools when troubleshooting.
    document.documentElement.setAttribute("data-odpdf-status", text);
  }

  let busy = false;
  async function run(withImages) {
    if (busy) return;
    busy = true;
    setStatus(tr("slidesPreparing"));
    try {
      const { n, secs } = await exportSlides(withImages, setStatus);
      setStatus(tr(withImages ? "slidesDone" : "slidesTextDone", { n, s: secs }));
    } catch (e) {
      setStatus(tr("slidesFailed") + (e && e.message ? e.message : e));
    } finally {
      busy = false;
    }
  }

  let ready = false;
  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || !ready) return; // only the ready viewer frame answers
    if (msg.type === "SLIDES_PING") {
      const p = slidePosition();
      sendResponse({ ok: true, total: p ? p.total : thumbnails().length, busy });
    } else if (msg.type === "SLIDES_RUN") {
      sendResponse({ ok: true, started: !busy });
      run(!!msg.withImages);
    }
  });

  function start() {
    const poll = setInterval(() => {
      if (viewerReady()) {
        clearInterval(poll);
        ready = true;
        document.documentElement.setAttribute("data-odpdf-slides", "ready");
      }
    }, 1000);
  }

  if (I18N) I18N.getLang((l) => { lang = l; start(); });
  else start();
})();
