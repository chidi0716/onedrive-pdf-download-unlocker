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

  // The slide currently shown, as {el, rect, method}. Each slide's container
  // holds an <svg> exactly the size of the slide (its background); the largest
  // slide-shaped one is the true slide rect (method "svg"). If that isn't in
  // the DOM yet we fall back to a slide-shaped div ("div"), and only as a last
  // resort derive a 16:9 rect from the wrapper ("derived").
  //
  // The distinction matters: the derived rect is taken from a container that,
  // early in the first slide's load (before the real slide SVG exists), spans
  // down into the viewer's status bar, so capturing it puts the status bar in
  // the image and shifts the slide up. settledSlideBox() therefore waits for a
  // non-derived box before the first capture instead of trusting whatever
  // currentSlideBox() returns immediately.
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
        if (!best || b.width * b.height > best.rect.width * best.rect.height) best = { el: svg, rect: b, method: "svg" };
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
      if (!best || b.width * b.height < best.rect.width * best.rect.height) best = { el, rect: b, method: "div" };
    }
    if (best) return best;
    // Last resort: derive a 16:9 slide rect centered in the visible wrapper
    // (never the whole panel — that includes the status bar), so a slide that
    // renders without a detectable background box can still be captured.
    if (wrap) {
      const hb = wrap.getBoundingClientRect();
      if (hb.width > 200 && hb.height > 120) {
        let w = hb.width, h = w * 9 / 16;
        if (h > hb.height) { h = hb.height; w = h * 16 / 9; }
        return { el: wrap, rect: { left: hb.left + (hb.width - w) / 2, top: hb.top + (hb.height - h) / 2, width: w, height: h, right: 0, bottom: 0 }, method: "derived" };
      }
    }
    return null;
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
        ".ShapeSelectionOverlay { display: none !important; }" +
        // The viewer's own status bar sits just below the slide; if a capture
        // rect is ever a touch too tall it would appear at the image's bottom.
        // Hide it (and the bottom toolbar) for the duration of the capture.
        "#WACStatusBarContainer, [id^=StatusBar], [class*=StatusBarContainer], [class*=DocumentStatusBar] { visibility: hidden !important; }" +
        // Floating notifications / callouts / tooltips (e.g. the "you don't have
        // permission to edit" toast) can overlap the slide's corner and get into
        // the capture. They're never part of the slide, so hide them while we
        // export. The capture is clipped to the slide, so this is all that can
        // land on top of it.
        ".ms-Callout, .ms-Layer, [class*=Callout], [class*=allout], [class*=otification], [class*=oastNotification], [class*=Tooltip], [class*=ooltip], [role=alert], [role=alertdialog], [role=tooltip] { visibility: hidden !important; }";
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
        if (sinceFetch < quietMs + 100) {
          timer = setTimeout(check, quietMs + 100 - sinceFetch);
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

  // Is `wrap` a rendered slide (not the blank "loading" placeholder)? A slide
  // that has switched but not yet drawn its content has an empty container with
  // just a loading placeholder; capturing it gives the broken-image blank pages
  // we saw. Treat it as rendered once it has real text, a shape/background SVG,
  // or a fully-loaded picture.
  function wrapperHasContent(wrap) {
    if (!wrap) return false;
    if (slideTextFrom(wrap).length > 0) return true;
    for (const svg of wrap.querySelectorAll("svg")) {
      const b = svg.getBoundingClientRect();
      const r = b.height ? b.width / b.height : 0;
      if (b.width >= 200 && r >= 1.2 && r <= 1.9) return true; // slide-shaped background
    }
    return [...wrap.querySelectorAll("img")].some((i) => i.complete && i.naturalWidth > 1);
  }

  // A cheap fingerprint of everything a slide has drawn so far: element count,
  // total text length, and how many images (and how many of those have finished
  // loading). When a picture streams in after the text, the image count and the
  // loaded count change, so the fingerprint keeps changing until the slide is
  // truly complete.
  function slideSignature(wrap) {
    const imgs = wrap.querySelectorAll("img");
    let loaded = 0;
    for (const i of imgs) if (i.complete && i.naturalWidth > 0) loaded++;
    return wrap.querySelectorAll("*").length + "/" + (wrap.textContent || "").length + "/" + imgs.length + "/" + loaded;
  }

  // Wait until slide `index` is the one actually shown AND it has finished
  // rendering, then capture. Finished = the target wrapper is the visible one,
  // every image in it has loaded, the network has been quiet for a beat, AND its
  // fingerprint hasn't changed across two consecutive checks. That last part is
  // what fixes "captured a page before it loaded": a slide whose text arrives
  // first but whose picture streams in a moment later keeps changing its
  // fingerprint (a new <img>, then that <img> finishing) until it's really done,
  // so we wait for it instead of firing on the text alone. Text-only slides have
  // no pending images and settle immediately, so they stay fast.
  async function waitSlideReady(index, maxMs) {
    const want = "PageContentSizeWrapper" + index;
    const start = performance.now();
    let wrap = null;
    while (performance.now() - start < maxMs) {
      wrap = [...document.querySelectorAll("[id^=PageContentSizeWrapper]")].find((w) => w.getBoundingClientRect().width > 0);
      if (wrap && wrap.id === want && wrapperHasContent(wrap)) break;
      wrap = null;
      await sleep(50);
    }
    if (!wrap) return;
    let prevSig = null, stable = 0;
    while (performance.now() - start < maxMs) {
      const imgs = [...wrap.querySelectorAll("img")];
      const pending = imgs.some((i) => !i.complete || i.naturalWidth === 0);
      const quiet = performance.now() - lastResourceAt > 180; // no tile/graphic fetch lately
      const sig = slideSignature(wrap);
      const same = sig === prevSig;
      prevSig = sig;
      if (!pending && quiet && same) {
        if (++stable >= 2) break; // confirmed done across two checks
      } else {
        stable = 0;
      }
      await sleep(90);
    }
    if (document.fonts && document.fonts.status !== "loaded") await document.fonts.ready.catch(() => {});
    await nextFrame();
  }

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

  // The slide box once the real slide background exists and has stopped moving.
  // Prefer an actual slide box (method "svg"/"div") over the derived fallback,
  // which early in the first slide's load spans the status bar: wait up to ~5s
  // for a real box before accepting the fallback, then wait for it to settle.
  async function settledSlideBox() {
    let box = null;
    for (let k = 0; k < 100; k++) { // up to ~10s to appear
      box = currentSlideBox();
      if (box && box.method !== "derived") break; // real slide box
      if (box && k >= 50) break; // only derived after ~5s: accept it
      box = box && box.method === "derived" ? box : null;
      await sleep(100);
    }
    if (!box) return null;
    let prev = null;
    for (let k = 0; k < 30; k++) {
      box = currentSlideBox() || box;
      if (box && prev && sameRect(box.rect, prev)) return box;
      prev = box && box.rect;
      await sleep(100);
    }
    return box || currentSlideBox();
  }

  function thumbByNumber(n) {
    return thumbnails().find((t) => +t.getAttribute("aria-posinset") === n) || null;
  }

  // The scrollable ancestor of the thumbnail list.
  function thumbScroller() {
    const any = thumbnails()[0];
    let sc = any ? any.parentElement : null;
    while (sc && sc.scrollHeight <= sc.clientHeight + 2) sc = sc.parentElement;
    return sc;
  }

  // Scroll the thumbnail pane so `t` sits fully inside the visible viewport
  // (scrollIntoView alone can leave it clipped by the pane or off-screen).
  async function ensureThumbVisible(t) {
    const sc = thumbScroller();
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
    let t = thumbByNumber(n);
    if (t) return t;
    const sc = thumbScroller();
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

  function selectedThumb() {
    return thumbnails().find((t) => t.getAttribute("aria-selected") === "true") || null;
  }

  // Trusted click on the centre of `el`, once it sits fully on screen.
  async function clickThumb(el) {
    if (!el) return false;
    await ensureThumbVisible(el);
    const b = el.getBoundingClientRect();
    if (b.width <= 4 || b.top < 0 || b.bottom > window.innerHeight) return false;
    await send({ type: "SLIDES_INPUT", input: { kind: "click", x: b.left + b.width / 2, y: b.top + b.height / 2 } });
    return true;
  }

  function thumbListHasFocus() {
    const a = document.activeElement;
    return !!(a && a !== document.body && a.closest('[role="listbox"], [role="option"]'));
  }

  function listboxEl() {
    const t = selectedThumb() || thumbnails()[0];
    return t ? (t.closest('[role="listbox"]') || t.parentElement) : null;
  }

  // Give the thumbnail list keyboard focus WITHOUT any synthetic click. slides.js
  // runs inside the viewer's own frame, so a plain in-frame focus() works and
  // needs no coordinates — unlike a debugger click, which depends on the frame
  // offset and, when that was off, landed on the slide itself, entered text-edit
  // mode and made the view flip in and out of presentation (the churn the user
  // saw, which also kept the caret blinking so the "settled" check never fired,
  // making it slow). Set a tabindex if the element isn't focusable, and try the
  // selected thumbnail then the listbox container.
  function focusThumbList() {
    if (thumbListHasFocus()) return true;
    for (const el of [selectedThumb(), listboxEl()]) {
      if (!el) continue;
      try {
        if (!el.hasAttribute("tabindex")) el.setAttribute("tabindex", "-1");
        el.focus({ preventScroll: true });
      } catch (e) { /* ignore */ }
      if (thumbListHasFocus()) return true;
    }
    return thumbListHasFocus();
  }

  // Bring slide `index` (0-based) on screen. The export walks the deck in
  // order, so this is almost always "next slide": one ArrowDown on the focused
  // thumbnail list, the same as a person pressing the key. Keys don't depend
  // on where anything is on screen, so they sidestep the coordinate and
  // thumbnail-scrolling problems of clicking. Clicking the target thumbnail is
  // kept for long jumps and as the fallback when keys stop moving the deck.
  async function goToSlide(index, total) {
    const want = index + 1;
    const current = () => { const p = slidePosition(); return p ? p.current : null; };
    const arrived = () => {
      if (current() === want) return true;
      const t = thumbByNumber(want);
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
      if (stuck < 3) {
        // Keyboard only: keys carry no coordinates, so nothing can land on the
        // slide and trip text-edit mode. This is the whole navigation path in
        // the normal case — a bare ArrowDown per slide.
        tries.push(key);
        await send({ type: "SLIDES_INPUT", input: { kind: "key", key } });
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
      while (performance.now() - s < 5000 && !arrived() && current() === cur) await sleep(60);
      stuck = current() === cur ? stuck + 1 : 0;
    }
    if (arrived()) return true;
    const p = slidePosition();
    throw new Error(`could not open slide ${want} (at ${p ? p.current + "/" + p.total : "?"}, thumbs ${thumbnails().length}, tried ${tries.slice(-12).join(",")})`);
  }

  // ---- export --------------------------------------------------------------

  async function fileBase() {
    const r = await send({ type: "SLIDES_TAB_TITLE" });
    let name = ((r && r.title) || document.title || "slides").replace(/\s*[-|–]\s*(PowerPoint|SharePoint|OneDrive).*$/i, "").replace(/\.pptx?$/i, "");
    name = name.replace(/[\\/:*?"<>|]+/g, "_").replace(/[\x00-\x1f\x7f]/g, "").replace(/[\s.]+$/, "").slice(0, 150).trim() || "slides";
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
  // 960 pt is the width of a standard 16:9 slide. 1600 px over a 960 pt page is
  // ~120 dpi — sharp for on-screen reading and printing, and roughly 40 %
  // fewer bytes than 2000 px, which keeps the image-based PDF from bloating.
  const OUTPUT_WIDTH_PX = 1600;
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
      // Focus the thumbnail list once up front so the whole export navigates by
      // bare arrow keys.
      focusThumbList();
      if (withImages) {
        // Park the mouse ONCE, off the slide, so it can't cast a hover tooltip
        // into any capture — but far from the top edge (which makes Arc show its
        // toolbar) and the bottom-right presentation button. Left-hand mid-height
        // sits over the thumbnail rail, which is outside every slide clip.
        await send({ type: "SLIDES_INPUT", input: { kind: "move", x: 5, y: Math.round(window.innerHeight / 2) } });
      }
      for (let i = 0; i < n; i++) {
        setStatus(tr("slidesProgress", { i: i + 1, n }));
        let t = performance.now();
        await goToSlide(i, n);
        timings.nav += performance.now() - t;
        t = performance.now();
        // Wait for the target slide to actually be shown and rendered (guards
        // against blank/duplicate captures) before capturing.
        await waitSlideReady(i, 8000);
        timings.render += performance.now() - t;
        texts.push(`--- Slide ${i + 1} ---\n${slideTextFrom(wrapperFor(i))}`);
        if (withImages) {
          setStatus(tr("slidesProgress", { i: i + 1, n }) + " ⤵");
          t = performance.now();
          // Capture via the debugger screenshot. (An in-page draw was tried but
          // rendered shapes and fonts wrong on real machines, so the screenshot —
          // faithful, though it flashes in Arc — is the reliable path.)
          if (bar) bar.style.visibility = "hidden";
          let res = null;
          for (let attempt = 0; attempt < 3; attempt++) {
            let box = currentSlideBox();
            if (!box || box.method === "derived") box = await settledSlideBox();
            if (!box) throw new Error("slide " + (i + 1) + " not found on screen");
            const r = box.rect;
            // devicePixelRatio: the debugger screenshot renders at the display's
            // DPR, so divide it out to keep the output at OUTPUT_WIDTH_PX.
            res = await send({ type: "SLIDES_CAPTURE", rect: { x: r.left, y: r.top, width: r.width, height: r.height }, scale: OUTPUT_WIDTH_PX / r.width / (window.devicePixelRatio || 1) });
            if (!res || !res.ok) throw new Error((res && res.error) || "capture failed");
            const after = currentSlideBox();
            if (after && sameRect(after.rect, r)) break;
          }
          if (bar) bar.style.visibility = "";
          timings.capture += performance.now() - t;
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
    // so an error thrown mid-capture must un-hide it or the message is lost (the
    // "progress just vanished" the user saw).
    bar.style.visibility = "";
    // Colour the bar so a failure is obvious at a glance during testing.
    bar.style.background = kind === "error" ? "rgba(140,20,20,.95)" : kind === "done" ? "rgba(20,110,50,.95)" : "rgba(32,32,32,.92)";
    // Also readable from DevTools / a screenshot when troubleshooting.
    document.documentElement.setAttribute("data-odpdf-status", text);
    console.log("[odpdf] " + text);
  }

  let busy = false;
  async function run(withImages) {
    if (busy) return;
    busy = true;
    setStatus(tr("slidesPreparing"));
    try {
      const { n, secs } = await exportSlides(withImages, setStatus);
      setStatus(tr(withImages ? "slidesDone" : "slidesTextDone", { n, s: secs }), "done");
    } catch (e) {
      // Show the full error (and where it happened) so a tester can report it.
      const msg = (e && e.stack) ? String(e.stack).split("\n").slice(0, 3).join("\n") : (e && e.message ? e.message : String(e));
      setStatus(tr("slidesFailed") + msg, "error");
      console.error("[odpdf] export failed", e);
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
