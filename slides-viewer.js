// slides-viewer.js
// Everything that depends on how PowerPoint for the web builds its page: the
// selectors, class names, ids and heuristics used to find the slides, the
// thumbnails, the current position and the notification toast. When Microsoft
// changes the viewer, this is the file to fix; slides.js only drives the export
// through the functions exported here (window.ODPDF_VIEWER).
//
// Confirmed against PowerPoint for the web (2026-09): thumbnails are
// role=option items with a "grid-content-thumbnail-view" class and
// aria-selected on the current one; every slide has a
// #PageContentSizeWrapperN container (only the current one has a size);
// empty placeholders render their "Click to add ..." prompt inside
// .visiblePromptTextContent; the status bar reads "Slide X of N"; the
// "no edit permission" toast is a Fluent callout (#BaseCallout…) holding a
// role=alertdialog with one OK button.
(function (global) {
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

  function thumbByNumber(n) {
    return thumbnails().find((t) => +t.getAttribute("aria-posinset") === n) || null;
  }

  function selectedThumb() {
    return thumbnails().find((t) => t.getAttribute("aria-selected") === "true") || null;
  }

  function listboxEl() {
    const t = selectedThumb() || thumbnails()[0];
    return t ? (t.closest('[role="listbox"]') || t.parentElement) : null;
  }

  function thumbListHasFocus() {
    const a = document.activeElement;
    return !!(a && a !== document.body && a.closest('[role="listbox"], [role="option"]'));
  }

  // The scrollable ancestor of the thumbnail list.
  function thumbScroller() {
    const any = thumbnails()[0];
    let sc = any ? any.parentElement : null;
    while (sc && sc.scrollHeight <= sc.clientHeight + 2) sc = sc.parentElement;
    return sc;
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

  // Total slide count, best effort.
  function slideCount() {
    const p = slidePosition();
    return p ? p.total : thumbnails().length;
  }

  function wrapperFor(index) {
    return document.getElementById("PageContentSizeWrapper" + index);
  }

  // The slide container currently on screen (only the shown one has a size).
  function visibleWrapper() {
    return [...document.querySelectorAll("[id^=PageContentSizeWrapper]")].find((w) => w.getBoundingClientRect().width > 0) || null;
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
  // the image and shifts the slide up. slides.js therefore waits for a
  // non-derived box before the first capture.
  function currentSlideBox() {
    const panel = viewPanel();
    if (!panel) return null;
    const wrap = visibleWrapper();
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
      // NOTE: the "no edit permission" toast is NOT CSS-hidden here. Hiding
      // any part Fluent positions (its beak, or the whole callout root) makes
      // Fluent think the show failed and re-create/re-position it forever — a
      // re-render loop that janks the browser and times out every capture.
      // dismissNotification() closes it instead.
      st.textContent =
        ".visiblePromptTextContent, .visiblePromptTextContent * { visibility: hidden !important; }" +
        ".ShapeSelectionOverlay { display: none !important; }" +
        // The viewer's own status bar sits just below the slide; hide it so a
        // slightly tall capture rect can't catch it.
        "#WACStatusBarContainer, [id^=StatusBar] { visibility: hidden !important; }";
      document.head.appendChild(st);
    } else if (!on && st) {
      st.remove();
    }
  }

  // The visible "no edit permission" toast, or null. Only one shape counts: a
  // Fluent callout containing a role=alertdialog with at most ONE button (the
  // OK). A real question dialog ("reload?", "discard changes?") has two or more
  // choices and is never matched, so we can't answer it by accident.
  function findPermissionToast() {
    for (const root of document.querySelectorAll("[id^=BaseCallout]")) {
      if (!root.getBoundingClientRect().width) continue;
      const dlg = root.querySelector("[role=alertdialog]");
      if (!dlg) continue;
      const buttons = dlg.querySelectorAll("button, [role=button]");
      if (buttons.length > 1) continue;
      return { root, dlg, button: buttons[0] || null };
    }
    return null;
  }

  // Close the toast the way a person would. First an outside click: Fluent
  // callouts are "light dismiss", so a click anywhere outside closes them (the
  // user confirmed clicking a blank area makes it go away), and that only ever
  // affects popups designed to be dismissed that way. If it is still there, press
  // its single OK button. The node is never removed by hand — React owns it.
  // Returns true if a toast was found.
  function dismissNotification() {
    const toast = findPermissionToast();
    if (!toast) return false;
    try {
      for (const type of ["mousedown", "mouseup", "click"]) {
        document.body.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
      }
    } catch (e) { /* ignore */ }
    if (toast.root.isConnected && toast.root.getBoundingClientRect().width && toast.button) {
      try { toast.button.click(); } catch (e) { /* ignore */ }
    }
    return true;
  }

  global.ODPDF_VIEWER = {
    viewPanel,
    viewerReady,
    thumbnails,
    thumbByNumber,
    selectedThumb,
    listboxEl,
    thumbListHasFocus,
    thumbScroller,
    slidePosition,
    slideCount,
    wrapperFor,
    visibleWrapper,
    currentSlideBox,
    slideTextFrom,
    wrapperHasContent,
    slideSignature,
    hideHints,
    dismissNotification,
  };
})(typeof window !== "undefined" ? window : globalThis);
