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

  function viewPanel() {
    return document.getElementById("WACViewPanel");
  }

  // The full editor has replaced the JPEG shown while booting.
  function viewerReady() {
    const boot = document.getElementById("WireframeBootSlideImage");
    if (boot && boot.isConnected && boot.getBoundingClientRect().width > 0) return false;
    return !!viewPanel() && thumbnails().length > 0;
  }

  // Slide thumbnails in the left pane, in slide order.
  function thumbnails() {
    const pane = document.querySelector("#SlidePanel, [id*=SlidePanel], [id*=ThumbnailPane], [aria-label*=humbnail]");
    const scope = pane || document;
    let items = [...scope.querySelectorAll('[role="option"], [role="tab"], [role="listitem"]')].filter((el) => {
      const b = el.getBoundingClientRect();
      return b.width > 60 && b.height > 30;
    });
    return items;
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

  function slideText(box) {
    if (!box) return "";
    const r = box.rect;
    const lines = [];
    for (const p of box.el.ownerDocument.querySelectorAll("#WACViewPanel .Paragraph, #WACViewPanel p")) {
      const b = p.getBoundingClientRect();
      if (b.width === 0 || b.bottom < r.top || b.top > r.bottom || b.right < r.left || b.left > r.right) continue;
      if (p.closest("[data-odpdf-hint]")) continue;
      const t = (p.innerText || p.textContent || "").replace(/​/g, "").trim();
      if (t && !isHintText(p)) lines.push(t);
    }
    return lines.join("\n");
  }

  // "Click to add title/subtitle/text" prompts shown for empty placeholders.
  function isHintText(el) {
    return /^(Click to add|Click to edit|按一下以新增|按一下以編輯|單擊此處添加)/i.test((el.textContent || "").trim());
  }

  function hideHints(on) {
    const panel = viewPanel();
    if (!panel) return;
    for (const el of panel.querySelectorAll("*")) {
      if (el.childElementCount === 0 && isHintText(el)) {
        // hide the whole placeholder box (text + dotted outline)
        const box = el.closest('[class*="Shape"], [class*="shape"]') || el;
        if (on) {
          if (!box.hasAttribute("data-odpdf-hint")) {
            box.setAttribute("data-odpdf-hint", box.style.visibility || "");
            box.style.visibility = "hidden";
          }
        }
      }
    }
    if (!on) {
      for (const box of panel.querySelectorAll("[data-odpdf-hint]")) {
        box.style.visibility = box.getAttribute("data-odpdf-hint");
        box.removeAttribute("data-odpdf-hint");
      }
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

  function signature(box) {
    return box ? (box.el.innerText || "").slice(0, 400) + "|" + box.el.querySelectorAll("*").length : "";
  }

  async function goToSlide(i, thumbs, prevSig) {
    const t = thumbs[i];
    t.scrollIntoView({ block: "nearest" });
    t.click();
    // wait for the main view to switch (or give up quietly after 3 s)
    const start = performance.now();
    while (performance.now() - start < 3000) {
      await nextFrame();
      const sel = t.getAttribute("aria-selected");
      const box = currentSlideBox();
      if ((sel === "true" || i === 0 || signature(box) !== prevSig) && box) return box;
      await sleep(50);
    }
    return currentSlideBox();
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
    const thumbs = thumbnails();
    const n = thumbs.length;
    if (!n) throw new Error("no slides found");
    const texts = [];
    const images = [];
    const timings = { nav: 0, render: 0, capture: 0 };

    if (withImages) {
      const r = await send({ type: "SLIDES_BEGIN" });
      if (!r || !r.ok) throw new Error((r && r.error) || "could not start capture");
    }
    let sig = "";
    try {
      for (let i = 0; i < n; i++) {
        setStatus(tr("slidesProgress", { i: i + 1, n }));
        let t = performance.now();
        const box = await goToSlide(i, thumbs, sig);
        timings.nav += performance.now() - t;
        if (!box) throw new Error("slide " + (i + 1) + " not found");
        t = performance.now();
        await waitForStable(box.el, 250, 8000);
        hideHints(true);
        await nextFrame();
        timings.render += performance.now() - t;
        const fresh = currentSlideBox() || box;
        sig = signature(fresh);
        texts.push(`--- Slide ${i + 1} ---\n${slideText(fresh)}`);
        if (withImages) {
          t = performance.now();
          const r = fresh.rect;
          const res = await send({ type: "SLIDES_CAPTURE", rect: { x: r.left, y: r.top, width: r.width, height: r.height }, scale: 2 });
          timings.capture += performance.now() - t;
          if (!res || !res.ok) throw new Error((res && res.error) || "capture failed");
          images.push(b64ToBytes(res.data));
        }
      }
    } finally {
      hideHints(false);
      if (withImages) await send({ type: "SLIDES_END" });
    }

    const base = fileBase();
    if (withImages) {
      const pdf = window.ODPDF_PDF.buildPdf(images, 2);
      saveBlob(new Blob([pdf], { type: "application/pdf" }), base + ".pdf");
    }
    saveBlob(new Blob([texts.join("\n\n") + "\n"], { type: "text/plain;charset=utf-8" }), base + ".txt");
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    console.log("[odpdf] slide export", { slides: n, secs, timingsMs: Object.fromEntries(Object.entries(timings).map(([k, v]) => [k, Math.round(v)])) });
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
