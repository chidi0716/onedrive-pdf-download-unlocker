// slides.js
// Export for PowerPoint files opened in PowerPoint for the web (including
// view-only / "block download" links). The viewer never hands the browser the
// .pptx file, but it does render every slide as ordinary HTML inside this
// frame, so we walk the slides, let each one finish rendering, and ask the
// background script to capture it at the chosen output width (see captureSlide
// in background.js). The captures become an image-based PDF; the slide text is
// saved alongside.
//
// What the viewer's DOM looks like (selectors, ids, heuristics) lives in
// slides-viewer.js; this file drives the export: navigation, waiting, capture,
// output and the status bar.
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
  const U = window.ODPDF_UTIL;
  const V = window.ODPDF_VIEWER;
  let lang = (I18N && I18N.DEFAULT_LANG) || "en";
  const tr = (k, vars) => {
    let s = I18N ? I18N.t(lang, k) : k;
    for (const [a, b] of Object.entries(vars || {})) s = s.replace("{" + a + "}", b);
    return s;
  };

  // Every wait in the export, in one place, so speed can be tuned here.
  const TIMING = {
    layoutQuietMs: 600, // slide box must hold still this long before the first capture
    layoutQuietMaxMs: 6000,
    slideReadyMaxMs: 5000, // cap on waiting for one slide to render
    slideAppearPollMs: 50,
    slidePollMs: 60,
    slideFloorMs: 200, // minimum time after a slide appears before capturing
    imageQuietMs: 250, // no picture fetched for this long = pictures are in
    fontWaitMs: 200,
    navSwitchMaxMs: 5000, // how long one key press may take to switch slides
    navStuckTries: 3, // key presses that moved nothing before clicking a thumbnail
    captureTries: 5, // shots per slide (it needs two identical ones)
    verifyGapMs: 120, // pause between the two shots compared
    pictureWaitMs: 5000, // max wait for a picture still showing its loading icon at capture time
    badCaptureRetryMs: 150,
    toastProbeSlides: 3, // look for the "no edit permission" toast on the first N slides
    toastSettleMs: 150,
    readyPollMs: 1000,
  };

  // The background reports this error when the user closed the browser's
  // "is debugging this browser" bar, which ends the capture session.
  const CANCELLED = "ODPDF_CANCELLED";
  function cancelledError() {
    const e = new Error("export cancelled");
    e.cancelled = true;
    return e;
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const nextFrame = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  const sameRect = (a, b) => !!a && !!b && ["left", "top", "width", "height"].every((p) => Math.abs(a[p] - b[p]) < 0.5);

  function send(msg) {
    return new Promise((resolve) => chrome.runtime.sendMessage(msg, (r) => {
      void chrome.runtime.lastError;
      resolve(r);
    }));
  }

  // Trusted (debugger) input; throws if the capture session was cancelled.
  async function input(inp) {
    const r = await send({ type: "SLIDES_INPUT", input: inp });
    if (r && !r.ok && r.error === CANCELLED) throw cancelledError();
    return r;
  }

  // Last time the frame fetched an IMAGE (a picture/graphic tile). Pictures load
  // after the slide's text, so we wait for them — but only for image fetches:
  // the viewer also makes periodic non-image background requests (autosave,
  // presence), and waiting for those to fall quiet would stall every slide.
  let lastImageFetchAt = 0;
  const IMG_FETCH_RE = /images\.ashx|\.(png|jpe?g|gif|svg|webp|emf|wmf)(\?|$)/i;
  try {
    new PerformanceObserver((list) => {
      for (const e of list.getEntries()) {
        // telemetry beacons can be image requests; they say nothing about the slide
        if (/RemoteUls|Telemetry|OneCollector|events\.data\.microsoft|\/collect\b|keepalive/i.test(e.name)) continue;
        if (e.initiatorType === "img" || IMG_FETCH_RE.test(e.name)) {
          lastImageFetchAt = Math.max(lastImageFetchAt, e.responseEnd || performance.now());
        }
      }
    }).observe({ type: "resource", buffered: false });
  } catch (e) {
    // PerformanceObserver unavailable: rely on the DOM checks alone
  }

  // Set once a font wait has timed out: some viewer fonts never report
  // "loaded", and paying the full wait on every slide adds up (50 slides x
  // 200 ms = 10 s). After the first timeout we stop waiting on fonts.
  let fontsStuck = false;

  // Wait until slide `index` is the one actually shown AND it has finished
  // rendering. Finished = the target wrapper is the visible one, every image in
  // it has loaded, no picture was fetched for a beat, AND its fingerprint hasn't
  // changed across two consecutive checks. That last part is what fixes
  // "captured a page before it loaded": a slide whose text arrives first but
  // whose picture streams in a moment later keeps changing its fingerprint (a
  // new <img>, then that <img> finishing) until it's really done. Text-only
  // slides have no pending images and settle immediately, so they stay fast.
  async function waitSlideReady(index) {
    const want = "PageContentSizeWrapper" + index;
    const start = performance.now();
    const maxMs = TIMING.slideReadyMaxMs;
    let wrap = null;
    while (performance.now() - start < maxMs) {
      wrap = V.visibleWrapper();
      if (wrap && wrap.id === want && V.wrapperHasContent(wrap)) break;
      wrap = null;
      await sleep(TIMING.slideAppearPollMs);
    }
    if (!wrap) return;
    const appeared = performance.now();
    let prevSig = null, stable = 0;
    while (performance.now() - start < maxMs) {
      const imgs = [...wrap.querySelectorAll("img")];
      // Only an image that is actually downloading counts as pending. The viewer
      // keeps src-less placeholder <img>s in every slide
      // (.SlidePictureIncrementalLoading) that never "load", and a broken or
      // blocked image is complete-but-empty for good; treating either as
      // pending made every slide wait out the full time cap.
      const pending =
        imgs.some((i) => (i.currentSrc || i.getAttribute("src")) && !i.complete) ||
        V.loadingPictures(wrap) > 0; // a picture still shows its grey "loading" icon
      const sig = V.slideSignature(wrap);
      const same = sig === prevSig;
      prevSig = sig;
      const imgQuiet = performance.now() - lastImageFetchAt > TIMING.imageQuietMs;
      if (!pending && imgQuiet && same && performance.now() - appeared > TIMING.slideFloorMs) {
        if (++stable >= 2) break;
      } else {
        stable = 0;
      }
      await sleep(TIMING.slidePollMs);
    }
    if (!fontsStuck && document.fonts && document.fonts.status !== "loaded") {
      const loaded = await Promise.race([
        document.fonts.ready.then(() => true, () => true),
        sleep(TIMING.fontWaitMs).then(() => false),
      ]);
      if (!loaded) fontsStuck = true;
    }
    await nextFrame();
  }

  // Wait until the slide hasn't moved for `quietMs` (attaching the debugger
  // shows a bar that resizes the page, and the viewer re-lays out after it).
  // Gives up after `maxMs`.
  async function waitLayoutQuiet(quietMs, maxMs) {
    const start = performance.now();
    let last = null;
    let since = performance.now();
    while (performance.now() - start < maxMs) {
      const box = V.currentSlideBox();
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

  // The slide box once the real slide background exists and has stopped moving.
  // Prefer an actual slide box (method "svg"/"div") over the derived fallback,
  // which early in the first slide's load spans the status bar: wait up to ~5s
  // for a real box before accepting the fallback, then wait for it to settle.
  async function settledSlideBox() {
    let box = null;
    for (let k = 0; k < 100; k++) { // up to ~10s to appear
      box = V.currentSlideBox();
      if (box && box.method !== "derived") break; // real slide box
      if (box && k >= 50) break; // only derived after ~5s: accept it
      box = box && box.method === "derived" ? box : null;
      await sleep(100);
    }
    if (!box) return null;
    let prev = null;
    for (let k = 0; k < 30; k++) {
      box = V.currentSlideBox() || box;
      if (box && prev && sameRect(box.rect, prev)) return box;
      prev = box && box.rect;
      await sleep(100);
    }
    return box || V.currentSlideBox();
  }

  // Scroll the thumbnail pane so `t` sits fully inside the visible viewport
  // (scrollIntoView alone can leave it clipped by the pane or off-screen).
  async function ensureThumbVisible(t) {
    const sc = V.thumbScroller();
    for (let k = 0; k < 6; k++) {
      const b = t.getBoundingClientRect();
      const vh = window.innerHeight;
      const top = sc ? Math.max(0, sc.getBoundingClientRect().top) : 0;
      const bot = sc ? Math.min(vh, sc.getBoundingClientRect().bottom) : vh;
      if (b.top >= top + 4 && b.bottom <= bot - 4) return true; // fully visible
      if (sc) sc.scrollTop += (b.top + b.height / 2) - (top + bot) / 2;
      else t.scrollIntoView({ block: "center" });
      await sleep(120);
    }
    const b = t.getBoundingClientRect();
    return b.top >= 0 && b.bottom <= window.innerHeight;
  }

  // The thumbnail list may only render the items near the viewport; scroll
  // it proportionally so thumbnail `n` of `total` gets rendered.
  async function revealThumb(n, total) {
    let t = V.thumbByNumber(n);
    if (t) return t;
    const sc = V.thumbScroller();
    if (!sc) return null;
    for (let k = 0; k < 6 && !t; k++) {
      sc.scrollTop = Math.max(0, ((n - 1) / Math.max(1, total)) * sc.scrollHeight - sc.clientHeight / 2);
      await nextFrame();
      await sleep(60);
      t = V.thumbByNumber(n);
      if (!t) {
        // nudge toward it using the numbers that are rendered
        const nums = V.thumbnails().map((x) => +x.getAttribute("aria-posinset")).filter(Boolean);
        if (nums.length && n < Math.min(...nums)) sc.scrollTop -= sc.clientHeight / 2;
        else if (nums.length && n > Math.max(...nums)) sc.scrollTop += sc.clientHeight / 2;
        await nextFrame();
        t = V.thumbByNumber(n);
      }
    }
    return t;
  }

  // Trusted click on the centre of `el`, once it sits fully on screen.
  async function clickThumb(el) {
    if (!el) return false;
    await ensureThumbVisible(el);
    const b = el.getBoundingClientRect();
    if (b.width <= 4 || b.top < 0 || b.bottom > window.innerHeight) return false;
    await input({ kind: "click", x: b.left + b.width / 2, y: b.top + b.height / 2 });
    return true;
  }

  // Give the thumbnail list keyboard focus WITHOUT any synthetic click. slides.js
  // runs inside the viewer's own frame, so a plain in-frame focus() works and
  // needs no coordinates — unlike a debugger click, which depends on the frame
  // offset and, when that was off, landed on the slide itself, entered text-edit
  // mode and made the view flip in and out of presentation. Set a tabindex if
  // the element isn't focusable, and try the selected thumbnail then the list.
  function focusThumbList() {
    if (V.thumbListHasFocus()) return true;
    for (const el of [V.selectedThumb(), V.listboxEl()]) {
      if (!el) continue;
      try {
        if (!el.hasAttribute("tabindex")) el.setAttribute("tabindex", "-1");
        el.focus({ preventScroll: true });
      } catch (e) { /* ignore */ }
      if (V.thumbListHasFocus()) return true;
    }
    return V.thumbListHasFocus();
  }

  // Bring slide `index` (0-based) on screen. The export walks the deck in
  // order, so this is almost always "next slide": one ArrowDown on the focused
  // thumbnail list, the same as a person pressing the key. Keys don't depend
  // on where anything is on screen, so they sidestep the coordinate and
  // thumbnail-scrolling problems of clicking. Clicking the target thumbnail is
  // kept as the fallback when keys stop moving the deck.
  async function goToSlide(index, total) {
    const want = index + 1;
    const current = () => { const p = V.slidePosition(); return p ? p.current : null; };
    const arrived = () => {
      if (current() === want) return true;
      const t = V.thumbByNumber(want);
      return !!(t && t.getAttribute("aria-selected") === "true");
    };
    const tries = [];
    let stuck = 0;
    for (let step = 0; step < total + 20 && !arrived(); step++) {
      const cur = current();
      const dist = cur == null ? Infinity : want - cur;
      let key;
      if (want === 1 && cur !== 1) key = "Home";
      else if (want === total && cur !== total && Math.abs(dist) > 3) key = "End";
      else key = dist > 0 ? "ArrowDown" : "ArrowUp";
      focusThumbList();
      if (stuck < TIMING.navStuckTries) {
        // Keyboard only: keys carry no coordinates, so nothing can land on the
        // slide and trip text-edit mode. This is the whole navigation path in
        // the normal case — a bare ArrowDown per slide.
        tries.push(key);
        await input({ kind: "key", key });
      } else {
        // Keys haven't moved the deck for several tries (focus refused?):
        // last-resort click on the target thumbnail to re-establish selection.
        tries.push("click");
        await clickThumb(await revealThumb(want, total));
        stuck = 0;
      }
      // Slides can take a while to switch on a slow connection; wait for the
      // position to change before deciding the input did nothing.
      const s = performance.now();
      while (performance.now() - s < TIMING.navSwitchMaxMs && !arrived() && current() === cur) await sleep(TIMING.slidePollMs);
      stuck = current() === cur ? stuck + 1 : 0;
    }
    if (arrived()) return true;
    const p = V.slidePosition();
    throw new Error(`could not open slide ${want} (at ${p ? p.current + "/" + p.total : "?"}, thumbs ${V.thumbnails().length}, tried ${tries.slice(-12).join(",")})`);
  }

  // ---- export --------------------------------------------------------------

  async function fileBase() {
    const r = await send({ type: "SLIDES_TAB_TITLE" });
    const name = ((r && r.title) || document.title || "slides")
      .replace(/\s*[-|–]\s*(PowerPoint|SharePoint|OneDrive).*$/i, "")
      .replace(/\.pptx?$/i, "");
    return U.sanitizeFilename(name) || "slides";
  }

  // Output presets the user picks in the popup. Each slide is rendered to `width`
  // pixels regardless of its on-screen size, so pages come out uniform and sharp;
  // `jpeg` is the JPEG quality. 960 pt is a standard 16:9 slide's width.
  const QUALITY = {
    standard: { width: 1600, jpeg: 80 }, // small file, fine on screen
    high: { width: 2000, jpeg: 85 },
    max: { width: 2560, jpeg: 92 }, // near-print
  };
  const PAGE_WIDTH_PT = 960;

  // Capture the slide on screen as JPEG bytes, `outputWidth` pixels wide.
  // A slide is accepted only once two consecutive shots are byte-identical.
  // The render waits above go by DOM and network signals, but the viewer can
  // still repaint a moment later (on a real 50-slide deck one table was once
  // captured missing), and identical pixels are the only proof that nothing
  // was still changing. Costs one extra shot per slide.
  let reshots = 0; // shots that differed from the one before (reported in the log)
  const sameBytes = (a, b) => {
    if (!a || !b || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  };
  async function captureCurrent(slideNo, outputWidth, jpeg) {
    let prev = null, last = null, pictureWaited = false;
    for (let attempt = 0; attempt < TIMING.captureTries; attempt++) {
      let box = V.currentSlideBox();
      if (!box || box.method === "derived") box = await settledSlideBox();
      if (!box) throw new Error("slide " + slideNo + " not found on screen");
      const r = box.rect;
      // devicePixelRatio: the debugger screenshot renders at the display's
      // DPR, so divide it out to keep the output at outputWidth.
      const res = await send({ type: "SLIDES_CAPTURE", rect: { x: r.left, y: r.top, width: r.width, height: r.height }, scale: outputWidth / r.width / (window.devicePixelRatio || 1), jpeg });
      if (!res || !res.ok) {
        if (res && res.error === CANCELLED) throw cancelledError();
        throw new Error((res && res.error) || "capture failed");
      }
      const b = U.base64ToBytes(res.data);
      // Guard against a bad capture (the first shot of an export has
      // occasionally come back not-a-JPEG): only accept real JPEG bytes,
      // otherwise re-shoot this slide rather than poison the PDF.
      if (!(b.length > 3 && b[0] === 0xff && b[1] === 0xd8)) {
        await sleep(TIMING.badCaptureRetryMs);
        continue;
      }
      const after = V.currentSlideBox();
      const moved = !(after && sameRect(after.rect, r));
      // A picture that only started loading after the slide looked finished
      // shows a grey "loading" box in a shot that is otherwise stable. Wait for
      // it to swap in, then shoot again.
      const wrap = V.visibleWrapper();
      if (wrap && !pictureWaited && V.loadingPictures(wrap) > 0) {
        pictureWaited = true; // at most one wait per slide; a picture that never loads is kept as is
        const t = performance.now();
        while (performance.now() - t < TIMING.pictureWaitMs && V.loadingPictures(wrap) > 0) await sleep(TIMING.slidePollMs);
        await nextFrame();
        prev = null;
        last = b;
        continue;
      }
      if (!moved && sameBytes(prev, b)) return b;
      if (prev) reshots++;
      prev = moved ? null : b;
      last = b;
      await sleep(TIMING.verifyGapMs);
    }
    if (!last) throw new Error("slide " + slideNo + ": capture was not a valid image");
    return last; // never settled (e.g. an animated picture): keep the latest shot
  }

  async function exportSlides(withImages, quality) {
    const q = QUALITY[quality] || QUALITY.standard;
    let outputWidth = q.width;
    const t0 = performance.now();
    const n = V.slideCount();
    if (!n) throw new Error("no slides found");
    const texts = [];
    const images = [];
    const timings = { nav: 0, render: 0, capture: 0 };
    reshots = 0;
    const secs = () => ((performance.now() - t0) / 1000).toFixed(1);

    // Text-only fast path: every slide already has a container in the DOM.
    if (!withImages) {
      let all = true;
      for (let i = 0; i < n; i++) if (!V.wrapperFor(i)) { all = false; break; }
      if (all) {
        for (let i = 0; i < n; i++) texts.push(`--- Slide ${i + 1} ---\n${V.slideTextFrom(V.wrapperFor(i))}`);
        U.saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), (await fileBase()) + ".txt");
        return { n, secs: secs() };
      }
    }

    const begin = await send({ type: "SLIDES_BEGIN" });
    if (!begin || !begin.ok) throw new Error((begin && begin.error) || "could not start capture");
    let failure = null;
    try {
      // The background has already waited for the "is debugging" bar to appear;
      // just let the slide box settle after that resize before measuring it.
      await waitLayoutQuiet(TIMING.layoutQuietMs, TIMING.layoutQuietMaxMs);
      let toastCleared = !withImages;
      if (withImages) {
        // Clamp the target width to the display's native resolution for the
        // slide. Asking the screenshot API for more pixels than the display has
        // (a high preset in a small window) makes it upscale, which some
        // browsers answer with empty/invalid data. Never exceed native.
        const b0 = V.currentSlideBox();
        const native = b0 ? Math.floor(b0.rect.width * (window.devicePixelRatio || 1)) : 0;
        if (native > 400 && native < outputWidth) outputWidth = native;
        V.hideHints(true);
        // Close the "no edit permission" toast now so it has gone by the first
        // capture; the first few slides check again in case it lands late.
        if (V.dismissNotification()) {
          toastCleared = true;
          await sleep(TIMING.toastSettleMs);
        }
      }
      // Focus the thumbnail list once up front so the export navigates by bare
      // arrow keys.
      focusThumbList();
      if (withImages) {
        // Park the mouse ONCE, off the slide, so it can't cast a hover tooltip
        // into any capture — but far from the top edge (which makes Arc show its
        // toolbar) and the bottom-right presentation button. Left-hand mid-height
        // sits over the thumbnail rail, which is outside every slide clip.
        await input({ kind: "move", x: 5, y: Math.round(window.innerHeight / 2) });
      }
      for (let i = 0; i < n; i++) {
        setStatus(tr("slidesProgress", { i: i + 1, n }));
        let t = performance.now();
        await goToSlide(i, n);
        timings.nav += performance.now() - t;
        t = performance.now();
        await waitSlideReady(i);
        timings.render += performance.now() - t;
        texts.push(`--- Slide ${i + 1} ---\n${V.slideTextFrom(V.wrapperFor(i))}`);
        if (withImages) {
          setStatus(tr("slidesProgress", { i: i + 1, n }) + " ⤵");
          t = performance.now();
          if (!toastCleared && i < TIMING.toastProbeSlides && V.dismissNotification()) {
            toastCleared = true;
            await sleep(TIMING.toastSettleMs);
          }
          // Capture via the debugger screenshot. (An in-page draw was tried but
          // rendered shapes and fonts wrong on real machines, so the screenshot —
          // faithful, though it flashes in Arc — is the reliable path.)
          if (bar) bar.style.visibility = "hidden";
          images.push(await captureCurrent(i + 1, outputWidth, q.jpeg));
          if (bar) bar.style.visibility = "";
          timings.capture += performance.now() - t;
        }
      }
    } catch (e) {
      failure = e;
    } finally {
      V.hideHints(false);
      await send({ type: "SLIDES_END" });
    }

    // On failure, still save whatever was captured (marked as partial) rather
    // than throwing away the slides that worked.
    const done = withImages ? images.length : texts.length;
    if (failure && !done) throw failure;
    const base = (await fileBase()) + (failure ? ` (1-${done} of ${n})` : "");
    if (withImages) {
      U.saveBlob(window.ODPDF_PDF.buildPdfBlob(images, outputWidth / PAGE_WIDTH_PT), base + ".pdf");
    }
    U.saveBlob(new Blob([texts.slice(0, done).join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), base + ".txt");
    console.log("[odpdf] slide export", JSON.stringify({ slides: n, done, secs: secs(), reshots, timingsMs: Object.fromEntries(Object.entries(timings).map(([k, v]) => [k, Math.round(v)])) }));
    if (failure) {
      failure.partial = done;
      failure.total = n;
      throw failure;
    }
    return { n, secs: secs() };
  }

  // ---- UI ------------------------------------------------------------------

  // The export is started from the extension popup or the on-page button
  // (content.js); the viewer grabs pointer events for itself, so buttons injected
  // into this frame never receive clicks. A small bar in the viewer shows
  // progress and the result.
  let bar = null;
  let statusEl = null;
  function ensureBar() {
    if (bar) return;
    bar = document.createElement("div");
    bar.id = "__odpdf_slides";
    bar.style.cssText =
      "position:fixed;left:12px;bottom:40px;max-width:60vw;z-index:2147483647;pointer-events:none;" +
      "font:13px/1.35 Segoe UI,system-ui,sans-serif;background:rgba(32,32,32,.92);color:#fff;padding:8px 12px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.3);white-space:pre-wrap";
    statusEl = document.createElement("span");
    bar.appendChild(statusEl);
    document.body.appendChild(bar);
  }
  function setStatus(text, kind) {
    ensureBar();
    statusEl.textContent = text;
    // Always make the bar visible: it is hidden momentarily during each capture,
    // so an error thrown mid-capture must un-hide it or the message is lost.
    bar.style.visibility = "";
    // Colour the bar so a failure is obvious at a glance.
    bar.style.background = kind === "error" ? "rgba(140,20,20,.95)" : kind === "done" ? "rgba(20,110,50,.95)" : "rgba(32,32,32,.92)";
    // Also readable from DevTools / a screenshot when troubleshooting.
    document.documentElement.setAttribute("data-odpdf-status", text);
    console.log("[odpdf] " + text);
  }

  let ready = false;
  let busy = false;

  // Tell the top page (content.js, via the background) whether a deck is here
  // and whether an export is running, so its on-page button can appear, hide
  // during an export and come back afterwards — no polling needed.
  function notifyState() {
    send({ type: "SLIDES_STATE", total: V.slideCount(), busy });
  }

  async function run(withImages, quality) {
    if (busy) return;
    busy = true;
    notifyState();
    setStatus(tr("slidesPreparing"));
    try {
      const { n, secs } = await exportSlides(withImages, quality);
      setStatus(tr(withImages ? "slidesDone" : "slidesTextDone", { n, s: secs }), "done");
    } catch (e) {
      const partial = e && e.partial ? " " + tr("slidesPartialSaved", { k: e.partial, n: e.total }) : "";
      if (e && e.cancelled) {
        setStatus(tr("slidesCancelled") + partial, "error");
      } else {
        // Show the full error (and where it happened) so a tester can report it.
        const msg = (e && e.stack) ? String(e.stack).split("\n").slice(0, 3).join("\n") : (e && e.message ? e.message : String(e));
        setStatus(tr("slidesFailed") + msg + (partial ? "\n" + partial.trim() : ""), "error");
      }
      console.error("[odpdf] export failed", e);
    } finally {
      busy = false;
      notifyState();
    }
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || !ready) return; // only the ready viewer frame answers
    if (msg.type === "SLIDES_PING") {
      sendResponse({ ok: true, total: V.slideCount(), busy });
    } else if (msg.type === "SLIDES_RUN") {
      sendResponse({ ok: true, started: !busy });
      run(!!msg.withImages, msg.quality);
    }
  });

  // Follow language changes made in the popup while the deck is open.
  try {
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area === "local" && I18N && changes[I18N.STORAGE_KEY]) {
        lang = changes[I18N.STORAGE_KEY].newValue || I18N.DEFAULT_LANG;
      }
    });
  } catch (e) {
    // storage unavailable: keep the language read at start
  }

  function start() {
    const poll = setInterval(() => {
      if (V.viewerReady()) {
        clearInterval(poll);
        ready = true;
        document.documentElement.setAttribute("data-odpdf-slides", "ready");
        notifyState();
      }
    }, TIMING.readyPollMs);
  }

  if (I18N) I18N.getLang((l) => { lang = l; start(); });
  else start();
})();
