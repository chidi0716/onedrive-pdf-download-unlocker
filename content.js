// content.js
// 在 OneDrive / SharePoint 的 PDF 預覽分頁裡顯示一個下載按鈕，不需要另外
// 點開擴充功能的彈出視窗。
//
// 設計上的取捨（記錄一下，避免之後又走回頭路）：
// 一開始試過把按鈕直接「插入」(insertAdjacentElement / appendChild) 到
// OneDrive 自己的工具列 DOM 裡，結果在不同頁面上位置跑掉、樣式也很醜——
// 原因是那個工具列是 Fluent UI 的 OverflowSet 元件，由 React 控制版面，
// 我們用純 DOM 操作硬塞節點進去，不會被它的版面計算考慮在內，所以位置、
// 大小都不可靠。
//
// 改成現在這個做法：按鈕本身完全獨立於 OneDrive 的 DOM 樹之外（直接掛在
// <html> 下面），用 position:fixed + 即時計算座標的方式，「看起來」貼在
// 工具列旁邊，但實際上不會被 OneDrive 的版面邏輯影響，樣式也完全由我們
// 自己控制，所以兩種情況（有原生下載按鈕 / 沒有）看起來會是同一顆按鈕、
// 同樣的樣式，只是定位點不同，不會再有「位置不一致、醜」的問題。

(function () {
  if (window.__odpdfInjected) return;
  window.__odpdfInjected = true;

  const EXT_PATTERN = /\.(pdf|docx?|xlsx?|pptx?)$/i;

  const NATIVE_BUTTON_SELECTORS = [
    "#downloadCommand",
    '[data-automation-id="downloadCommand"]',
    '[data-automationid="downloadCommand"]',
    'button[aria-label*="下載"]',
    '[role="menuitem"][aria-label*="下載"]',
    '[role="button"][aria-label*="下載"]',
    'button[aria-label*="Download"]',
    '[role="menuitem"][aria-label*="Download"]',
    '[role="button"][aria-label*="Download"]',
    '[data-automationid="downloadButton"]',
    '[data-testid="download-button"]',
    'button[name="Download"]',
    '[data-icon-name="Download"]',
  ];

  const TOOLBAR_CONTAINER_SELECTORS = [
    ".ms-OverflowSet.ms-CommandBar-primaryCommand",
    ".ms-CommandBar-primaryCommand",
    '[data-automation-id="visibleContent"]',
    "#OneUpCommandBar",
    '[role="menubar"]',
  ];

  const ANCHOR_SEARCH_TIMEOUT_MS = 4000;
  const ANCHOR_SEARCH_INTERVAL_MS = 300;
  const REPOSITION_INTERVAL_MS = 500;

  let currentCandidate = null;
  let btnEl = null;
  let anchorEl = null; // 用來計算座標的參考元素（不會去動它的 DOM）
  let anchorMode = null; // "button" | "toolbar" | null（none，退回固定右下角）
  let dismissed = false;
  let anchorIntervalId = null; // startAnchorSearch() 的 setInterval id，切換檔案時要記得清掉
  let idleLabel = "Download PDF";
  let isBusy = false;
  let noCandidateMode = false; // 寬限期過後還沒抓到任何候選檔案時，按鈕仍會顯示，但是「未偵測到」的唯讀狀態
  let noCandidateTimerId = null;
  const NO_CANDIDATE_GRACE_MS = 3000; // 給偵測一點時間，避免一載入就馬上顯示「未偵測到」嚇到使用者

  // 語言設定：預設英文，使用者在 popup 裡切換後存在 chrome.storage.local，
  // 這裡讀出來決定按鈕文字要顯示哪種語言；popup 切換語言時也會即時同步
  // 過來（不需要重新整理頁面）。
  let LANG = (window.ODPDF_I18N && ODPDF_I18N.DEFAULT_LANG) || "en";
  function tr(key) {
    if (!window.ODPDF_I18N) return key;
    return ODPDF_I18N.t(LANG, key);
  }
  if (window.ODPDF_I18N) {
    ODPDF_I18N.getLang((lang) => {
      LANG = lang;
      idleLabel = tr("btnIdleFull");
      refreshIdleLabel();
    });
    try {
      chrome.storage.onChanged.addListener((changes, area) => {
        if (area === "local" && changes[ODPDF_I18N.STORAGE_KEY]) {
          LANG = changes[ODPDF_I18N.STORAGE_KEY].newValue || ODPDF_I18N.DEFAULT_LANG;
          if (btnEl) {
            if (noCandidateMode) {
              btnEl.title = tr("btnNotFoundTitle");
              btnEl.setAttribute("aria-label", tr("btnNotFoundTitle"));
            } else {
              btnEl.title = tr("btnTitle");
              btnEl.setAttribute("aria-label", tr("btnTitle"));
            }
            const closeEl = document.getElementById("__odpdf_close");
            if (closeEl) {
              closeEl.title = tr("closeTitle");
              closeEl.setAttribute("aria-label", tr("closeTitle"));
            }
            refreshIdleLabel();
          }
        }
      });
    } catch (e) {
      // storage 監聽失敗就靜默忽略，不影響主功能
    }
  }

  function sanitizeFilename(name) {
    name = (name || "").trim();
    name = name.replace(/[\\/:*?"<>|]/g, "_");
    name = name.replace(/[\x00-\x1f\x7f]/g, "");
    name = name.replace(/[\s.]+$/g, "");
    if (name.length > 150) name = name.slice(0, 150).trim();
    return name;
  }

  function ensurePdfExt(name) {
    if (!name) name = "document";
    if (!/\.pdf$/i.test(name)) name += ".pdf";
    return name;
  }

  function domFilenameGuess() {
    const knownSelectors = [
      '[data-automationid="title"]',
      ".od-ItemContent-fileName",
      ".od-TopBar-item.od-TopBar-title",
      "#OneUp_Title",
      '[data-automation-id="documentTitle"]',
      "header h1",
    ];
    for (const sel of knownSelectors) {
      const el = document.querySelector(sel);
      if (el && el.textContent && el.textContent.trim()) return el.textContent.trim();
    }

    const ariaEls = document.querySelectorAll("[aria-label]");
    for (const el of ariaEls) {
      const v = el.getAttribute("aria-label");
      if (v && EXT_PATTERN.test(v.trim()) && v.trim().length < 200) return v.trim();
    }

    const fuiNodes = document.querySelectorAll('[class*="fui-Text"]');
    for (const el of fuiNodes) {
      const text = el.textContent && el.textContent.trim();
      if (text && EXT_PATTERN.test(text) && text.length < 200) return text;
    }

    const all = document.querySelectorAll("body *");
    for (const el of all) {
      if (el.children && el.children.length > 0) continue;
      const text = el.textContent && el.textContent.trim();
      if (text && EXT_PATTERN.test(text) && text.length < 200) return text;
    }

    return "";
  }

  function guessFilename(candidate) {
    const domName = domFilenameGuess();
    if (domName) {
      let name = sanitizeFilename(domName);
      name = name.replace(/\s*[-|–]\s*(OneDrive|SharePoint).*$/i, "");
      return ensurePdfExt(name);
    }

    if (
      candidate.filename &&
      !/^odbtestpassthrough$/i.test(candidate.filename.replace(/\.[a-z0-9]+$/i, ""))
    ) {
      return ensurePdfExt(sanitizeFilename(candidate.filename));
    }

    let name = (document.title || "document").trim();
    name = name.replace(/\s*[-|–]\s*(OneDrive|SharePoint).*$/i, "");
    name = sanitizeFilename(name);
    return ensurePdfExt(name);
  }

  // 用版面尺寸 + computed style 判斷可見性，比 offsetParent 可靠
  // (offsetParent 在 position:fixed/sticky 元素上常常是 null，即使畫面上
  // 看得到也一樣，OneDrive 的工具列就是這種情況)。
  function isVisible(el) {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return false;
    const style = window.getComputedStyle(el);
    if (style.display === "none" || style.visibility === "hidden") return false;
    return true;
  }

  // 判斷某個元素是不是我們自己注入的按鈕/關閉鈕（或它們的子節點）。
  //
  // 抓到的真正 bug：我們自己按鈕的 aria-label／title 文字（例如英文版的
  // "Download this file with the extension..."）裡面就包含「Download」
  // 這個字，剛好符合 NATIVE_BUTTON_SELECTORS 裡
  // 'button[aria-label*="Download"]' 這條選擇器。在「頁面本身沒有原生
  // 下載按鈕」的情況下，一旦我們的按鈕被建立出來，下一次 tick() 重新搜尋
  // 錨點時，querySelector 找到的「原生下載按鈕」其實就是我們自己上一輪
  // 注入的按鈕——於是把自己當成錨點，定位成「貼在自己下面 6px」，每跑一
  // 次 tick（500ms）就再往下移動一個按鈕高度，完全不會停，這就是「不斷
  // 一直往下跑」的真正原因（之前修的「容器高度上限」只解決了另一個獨立
  // 的、一次性的量測錯誤，沒抓到這個會不斷累積的回授迴圈）。
  // 修法：搜尋原生按鈕／工具列時，明確排除我們自己注入的節點。
  function isOwnInjectedElement(el) {
    return !!(el && el.closest && el.closest("#__odpdf_btn, #__odpdf_close"));
  }

  function findNativeDownloadButton() {
    for (const sel of NATIVE_BUTTON_SELECTORS) {
      try {
        const matches = document.querySelectorAll(sel);
        for (const candidate of matches) {
          if (isOwnInjectedElement(candidate)) continue;
          let el = candidate;
          const clickable = el.closest('button, [role="menuitem"], [role="button"]');
          if (clickable) el = clickable;
          if (isOwnInjectedElement(el)) continue;
          if (isVisible(el)) return el;
        }
      } catch (e) {
        // 忽略無效選擇器
      }
    }
    return null;
  }

  function findToolbarContainer() {
    for (const sel of TOOLBAR_CONTAINER_SELECTORS) {
      try {
        const el = document.querySelector(sel);
        if (el && !isOwnInjectedElement(el) && isVisible(el)) return el;
      } catch (e) {
        // 忽略無效選擇器
      }
    }
    return null;
  }

  function injectStyle() {
    if (document.getElementById("__odpdf_style")) return;
    const style = document.createElement("style");
    style.id = "__odpdf_style";
    style.textContent = `
      /* 預設（淺色）樣式：貼著原生工具列/按鈕時用的低調樣式——白底、
         細邊框、深灰文字，跟 OneDrive 預設的淺色介面對比剛好夠看清楚，
         又不會像一塊很搶眼的色塊。

         之前這裡曾經改成用 getComputedStyle() 去抓附近原生按鈕的文字
         顏色直接套用，理論上更貼合任何配色，但忽略了背景：如果頁面是
         深色主題，抓到的文字顏色會是淺色（例如白色），套到我們固定的
         白底按鈕上就會變成「白字配白底」，完全看不到，滑過去才靠
         hover 的灰色疊色勉強看到一點輪廓——這就是上次回報「壞的、移
         上去才看得到字」的根本原因。

         現在改成更穩、更可預期的做法：不去猜「精確顏色該抓多少」，
         只判斷錨點背景是「整體偏深」還是「整體偏淺」這一個二元問題
         （見下面 isDarkBackground()），然後套用兩組我們自己準備好、
         保證對比足夠的配色（淺色版 / 深色版），而不是把任意抓到的顏色
         直接套用。這樣不管 OneDrive 是淺色或深色主題，文字永遠看得到，
         不會再出現顏色相同、互相蓋掉看不到字的情況。 */
      #__odpdf_btn {
        position: fixed;
        z-index: 2147483647;
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 3px 8px;
        height: 24px;
        box-sizing: border-box;
        cursor: pointer;
        font-family: -apple-system, "Segoe UI", "Microsoft JhengHei", "PingFang TC", sans-serif;
        font-size: 12.5px;
        font-weight: 400;
        color: #323130;
        background: rgba(255,255,255,0.95);
        border: 1px solid rgba(0,0,0,0.12);
        border-radius: 4px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.15);
        white-space: nowrap;
        user-select: none;
        transition: background 0.15s ease, color 0.15s ease, border-color 0.15s ease, opacity 0.15s ease;
      }
      #__odpdf_btn:hover { background: rgba(0,0,0,0.05); }
      #__odpdf_btn:active { background: rgba(0,0,0,0.1); }
      #__odpdf_btn.__odpdf_disabled { opacity: 0.6; cursor: default; pointer-events: none; }

      /* 深色版：當錨點所在區塊背景偏深（例如 OneDrive 的深色主題工具列）
         時套用，文字反過來用淺色，確保任何情況下文字跟背景都有足夠對比。 */
      #__odpdf_btn.__odpdf_dark {
        color: #f3f2f1;
        background: rgba(40,40,40,0.92);
        border: 1px solid rgba(255,255,255,0.18);
        box-shadow: 0 1px 3px rgba(0,0,0,0.4);
      }
      #__odpdf_btn.__odpdf_dark:hover { background: rgba(255,255,255,0.12); }
      #__odpdf_btn.__odpdf_dark:active { background: rgba(255,255,255,0.2); }

      /* 找不到任何可以依附的工具列/按鈕時，退而用的右下角浮動樣式：
         這時旁邊沒有原生按鈕可以參考樣式，所以用低調的深灰／黑色
         實心圓角樣式，盡量不刺眼，也避免使用者已經排除過的紅、黃、紫、藍。 */
      #__odpdf_btn.__odpdf_fallback {
        border-radius: 999px;
        padding: 9px 14px;
        height: auto;
        color: #fff;
        background: #3b3a39;
        border: 1px solid rgba(255,255,255,0.25);
        box-shadow: 0 4px 14px rgba(0,0,0,0.3);
      }
      #__odpdf_btn.__odpdf_fallback:hover { background: #292827; }
      #__odpdf_btn.__odpdf_fallback:active { background: #1f1e1d; }
      #__odpdf_btn .__odpdf_icon { font-size: 13px; line-height: 1; }
      /* 平常不顯示（使用者不喜歡叉叉一直在），滑鼠移到下載按鈕或叉叉上時
         才淡入。用 opacity（不是 display:none）來隱藏：
         - display:none 的元素無法被 Tab 鍵聚焦，會讓鍵盤使用者完全拿不到
           這顆關閉鈕；opacity:0 + pointer-events:none 隱藏時一樣不會擋
           到下面的滑鼠點擊，但 Tab 鍵還是能聚焦到它（聚焦時靠 JS 加上
           __odpdf_close_visible class 讓它淡入）。
         - 是否顯示交給 JS（mouseenter/mouseleave + 一小段延遲、focus/blur）
           判斷，而不是純 CSS :hover，因為按鈕跟叉叉中間有間距，純滑鼠移
           出按鈕範圍就立刻隱藏的話，滑鼠根本來不及移到叉叉上。 */
      #__odpdf_close {
        position: fixed;
        z-index: 2147483647;
        width: 16px;
        height: 16px;
        border-radius: 50%;
        background: #fff;
        color: #888;
        border: 1px solid #ddd;
        font-size: 10px;
        line-height: 1;
        cursor: pointer;
        box-shadow: 0 1px 3px rgba(0,0,0,0.2);
        display: flex;
        align-items: center;
        justify-content: center;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.15s ease;
      }
      #__odpdf_close.__odpdf_close_visible {
        opacity: 1;
        pointer-events: auto;
      }
      #__odpdf_close:hover {
        background: #f3f2f1;
        color: #444;
      }
    `;
    document.documentElement.appendChild(style);
  }

  function buildButton() {
    if (btnEl) return;
    injectStyle();

    btnEl = document.createElement("button");
    btnEl.id = "__odpdf_btn";
    btnEl.type = "button";
    btnEl.innerHTML = '<span class="__odpdf_icon" aria-hidden="true">⬇</span><span class="__odpdf_label">' + tr("btnIdleFull") + '</span>';
    btnEl.title = tr("btnTitle");
    // <button> 本身就能用 Tab 鍵聚焦、Enter/空白鍵觸發，不需要額外處理；
    // 但螢幕報讀軟體預設只會讀到圖示符號跟文字標籤，加 aria-label 讓
    // 報讀軟體讀出來的是完整說明而不是「下載」兩個字配一個箭頭符號。
    btnEl.setAttribute("aria-label", tr("btnTitle"));
    btnEl.addEventListener("click", onBtnClick);
    document.documentElement.appendChild(btnEl);

    const close = document.createElement("div");
    close.id = "__odpdf_close";
    close.textContent = "✕";
    close.title = tr("closeTitle");
    // 這顆關閉鈕是用 <div> 做的，預設不能被 Tab 鍵聚焦也不能用鍵盤觸發，
    // 補上 role="button" + tabindex + aria-label + 鍵盤事件，讓只用鍵盤
    // 操作或用螢幕報讀軟體的使用者也能關閉這顆按鈕，不是只有滑鼠可以點。
    close.setAttribute("role", "button");
    close.tabIndex = 0;
    close.setAttribute("aria-label", tr("closeTitle"));
    const dismiss = (e) => {
      e.stopPropagation();
      dismissed = true;
      btnEl.remove();
      close.remove();
      btnEl = null;
    };
    close.addEventListener("click", dismiss);
    close.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " " || e.key === "Spacebar") {
        e.preventDefault();
        dismiss(e);
      }
    });
    document.documentElement.appendChild(close);

    // 顯示/隱藏關閉鈕：滑鼠移到下載按鈕或關閉鈕本身上就淡入；移開後延遲
    // 一小段時間（300ms）才淡出，讓滑鼠有足夠時間從按鈕移到旁邊的叉叉，
    // 不會因為兩者中間有間距就一離開按鈕馬上消失、永遠點不到。鍵盤使用
    // 者 Tab 到關閉鈕時也會觸發 focus 顯示，blur 時才淡出，確保不會出
    // 現「聚焦到一個看不見的按鈕」這種情況。
    let closeHideTimer = null;
    const showClose = () => {
      if (closeHideTimer) {
        clearTimeout(closeHideTimer);
        closeHideTimer = null;
      }
      close.classList.add("__odpdf_close_visible");
    };
    const scheduleHideClose = () => {
      if (closeHideTimer) clearTimeout(closeHideTimer);
      closeHideTimer = setTimeout(() => {
        closeHideTimer = null;
        close.classList.remove("__odpdf_close_visible");
      }, 300);
    };
    btnEl.addEventListener("mouseenter", showClose);
    btnEl.addEventListener("mouseleave", scheduleHideClose);
    btnEl.addEventListener("focus", showClose);
    btnEl.addEventListener("blur", scheduleHideClose);
    close.addEventListener("mouseenter", showClose);
    close.addEventListener("mouseleave", scheduleHideClose);
    close.addEventListener("focus", showClose);
    close.addEventListener("blur", scheduleHideClose);
  }

  function setLabel(text, opts) {
    if (!btnEl) return;
    const label = btnEl.querySelector(".__odpdf_label");
    if (label) label.textContent = text;
    btnEl.classList.toggle("__odpdf_disabled", !!(opts && opts.disabled));
  }

  // 統一決定「閒置狀態」要顯示什麼文字：
  // - 還在忙（下載中/剛存檔/失敗訊息）的時候完全不要動，由呼叫端自己控制
  // - 還沒抓到任何候選檔案（noCandidateMode）：顯示「未偵測到」+ disabled
  // - 已經有候選檔案：顯示正常的下載文字
  // 集中成一個函式，是因為這個分支邏輯原本散落在 repositionButton()（每
  // 500ms 跑一次）跟兩個語言切換的地方，很容易漏改其中一處，導致「未偵測
  // 到」的文字被定時器自己蓋掉。
  function refreshIdleLabel() {
    if (isBusy) return;
    if (noCandidateMode) {
      setLabel(tr("btnNotFound"), { disabled: true });
    } else {
      setLabel(idleLabel, { disabled: false });
    }
  }

  // 從某個元素開始往上找第一個「不透明」的背景色，藉此判斷錨點所在
  // 區塊整體是偏深色還是偏淺色主題。只要這個二元判斷，不去抓精確的
  // 顏色數值——上一版直接套用抓到的文字顏色，卻沒考慮背景，深色主題
  // 下會變成「白字配白底」整個看不到，這次改成只用來挑選我們自己
  // 準備好、保證可讀的兩組配色（見上面 CSS 的 .__odpdf_dark）。
  function getBackgroundLuminance(el) {
    let node = el;
    let hops = 0;
    while (node && hops < 12) {
      try {
        const bg = window.getComputedStyle(node).backgroundColor;
        const m = bg && bg.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)/);
        if (m) {
          const alpha = m[4] !== undefined ? parseFloat(m[4]) : 1;
          if (alpha > 0.05) {
            const r = parseFloat(m[1]) / 255;
            const g = parseFloat(m[2]) / 255;
            const b = parseFloat(m[3]) / 255;
            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
          }
        }
      } catch (e) {
        // 忽略，往上一層找
      }
      node = node.parentElement;
      hops++;
    }
    return null; // 找不到任何不透明背景，當作淺色處理
  }

  function isDarkBackground(el) {
    const luminance = getBackgroundLuminance(el);
    return luminance != null && luminance < 0.45;
  }

  // 每隔一段時間重新計算一次按鈕座標：
  // - 找得到原生下載按鈕：貼在它正右邊、垂直置中對齊
  // - 找不到按鈕但找得到工具列：貼在工具列最右邊、垂直置中對齊
  // - 兩個都找不到：固定在畫面右下角（這是唯一的保底狀態）
  function repositionButton() {
    if (!btnEl) return;
    const closeEl = document.getElementById("__odpdf_close");

    if (anchorEl && isVisible(anchorEl)) {
      const rect = anchorEl.getBoundingClientRect();
      btnEl.classList.remove("__odpdf_fallback");
      btnEl.classList.toggle("__odpdf_dark", isDarkBackground(anchorEl));
      idleLabel = tr("btnIdleShort");
      refreshIdleLabel();
      // 不管是貼著原生按鈕還是貼著整條工具列，都改成貼在「下方」而不是
      // 「右邊」。原因：工具列裡的圖示通常排得很緊，右邊那一點點空隙
      // 根本不夠放一個有文字的按鈕，貼右邊一定會疊到後面的圖示（不管
      // 後面是「更多」按鈕還是其他工具）。貼在錨點正下方、靠左對齊，
      // 下面通常是文件內容區的空白，不會跟任何既有的工具列項目重疊。
      const btnWidth = btnEl.offsetWidth || 70;
      let left = rect.left;
      // 避免按鈕超出視窗右邊界
      left = Math.min(left, window.innerWidth - btnWidth - 8);
      left = Math.max(left, 4);

      // 某些頁面的工具列容器（例如 [role="menubar"]）DOM 裡會包含收合、
      // 不可見的項目（Fluent UI 的 OverflowSet 把超出寬度的按鈕收進一個
      // 隱藏列表裡，但節點還留在容器內），導致 getBoundingClientRect()
      // 量到的高度遠比畫面上看到的那一條工具列高很多——這時 rect.bottom
      // 會落到頁面中間甚至更下面，按鈕就會「自己往下跑」，而且因為跑到
      // 完全不同的區塊，旁邊的背景顏色也不一樣，連帶讓深淺色判斷跟著跑
      // 掉（這就是「顏色跟其他頁面不統一」的成因）。
      // 一般工具列實際可見的高度不會超過這個值太多，超過就視為量到了隱藏
      // 內容，改用「容器頂部 + 一個正常工具列高度」估算，不要相信過大的
      // rect.bottom。
      const MAX_REASONABLE_TOOLBAR_HEIGHT = 56;
      const anchorBottom =
        rect.height > MAX_REASONABLE_TOOLBAR_HEIGHT
          ? rect.top + MAX_REASONABLE_TOOLBAR_HEIGHT
          : rect.bottom;

      btnEl.style.top = anchorBottom + 6 + "px";
      btnEl.style.left = left + "px";
      btnEl.style.right = "";
      btnEl.style.bottom = "";
      if (closeEl) {
        closeEl.style.top = parseFloat(btnEl.style.top) - 8 + "px";
        closeEl.style.left = left + btnWidth + 4 + "px";
        closeEl.style.right = "";
        closeEl.style.bottom = "";
      }
    } else {
      btnEl.classList.add("__odpdf_fallback");
      btnEl.classList.remove("__odpdf_dark"); // 保底樣式固定用深灰實心，不需要再判斷明暗
      idleLabel = tr("btnIdleFull");
      refreshIdleLabel();
      btnEl.style.top = "";
      btnEl.style.left = "";
      btnEl.style.right = "22px";
      btnEl.style.bottom = "22px";
      if (closeEl) {
        closeEl.style.top = "";
        closeEl.style.left = "";
        closeEl.style.right = "18px";
        closeEl.style.bottom = "64px";
      }
    }
  }

  async function onBtnClick() {
    if (!currentCandidate) return;
    const candidate = currentCandidate;
    isBusy = true;
    setLabel(tr("btnDownloading"), { disabled: true });

    try {
      const filename = guessFilename(candidate);
      const resp = await chrome.runtime.sendMessage({
        type: "DOWNLOAD_FILE",
        url: candidate.url,
        filename,
        referrer: location.href,
        headers: candidate.forwardHeaders,
      });

      if (!resp || !resp.ok) {
        const reason = (resp && (resp.status || resp.statusText || resp.error)) || "未知錯誤";
        setLabel(tr("btnFailed") + reason, { disabled: false });
        setTimeout(() => { isBusy = false; setLabel(idleLabel, { disabled: false }); }, 3500);
        return;
      }

      const byteChars = atob(resp.base64);
      const byteNumbers = new Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
      const blob = new Blob([new Uint8Array(byteNumbers)], {
        type: resp.contentType || "application/pdf",
      });
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = filename;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);

      setLabel(tr("btnSaved"), { disabled: false });
      setTimeout(() => { isBusy = false; setLabel(idleLabel, { disabled: false }); }, 3500);
    } catch (e) {
      setLabel(tr("btnFailed") + (e && e.message ? e.message : String(e)), { disabled: false });
      setTimeout(() => { isBusy = false; setLabel(idleLabel, { disabled: false }); }, 3500);
    }
  }

  function startAnchorSearch() {
    const startedAt = Date.now();

    const tick = () => {
      if (dismissed) return;
      const nativeBtn = findNativeDownloadButton();
      if (nativeBtn) {
        anchorEl = nativeBtn;
        anchorMode = "button";
      } else {
        const container = findToolbarContainer();
        if (container) {
          anchorEl = container;
          anchorMode = "toolbar";
        } else if (Date.now() - startedAt >= ANCHOR_SEARCH_TIMEOUT_MS) {
          anchorEl = null;
          anchorMode = null;
        }
      }
      repositionButton();
    };

    tick();
    anchorIntervalId = setInterval(() => {
      if (dismissed) {
        clearInterval(anchorIntervalId);
        anchorIntervalId = null;
        return;
      }
      tick();
    }, REPOSITION_INTERVAL_MS);

    window.addEventListener("resize", repositionButton);
    window.addEventListener("scroll", repositionButton, true);
  }

  // 拆掉目前的按鈕跟所有相關狀態，回到「還沒偵測到任何檔案」的狀態。
  // 用在使用者於同一個分頁內切換到另一個檔案時（單頁應用程式式導覽，
  // 不會整頁重新載入，所以這個 content script 不會重新執行一次，必須
  // 自己手動重置，否則按鈕會一直指向舊檔案的下載網址）。
  function resetForNewDocument() {
    if (anchorIntervalId) {
      clearInterval(anchorIntervalId);
      anchorIntervalId = null;
    }
    if (noCandidateTimerId) {
      clearTimeout(noCandidateTimerId);
      noCandidateTimerId = null;
    }
    window.removeEventListener("resize", repositionButton);
    window.removeEventListener("scroll", repositionButton, true);

    if (btnEl) {
      btnEl.remove();
      btnEl = null;
    }
    const closeEl = document.getElementById("__odpdf_close");
    if (closeEl) closeEl.remove();

    currentCandidate = null;
    anchorEl = null;
    anchorMode = null;
    dismissed = false; // 換了一個檔案，之前「這次先不顯示」的選擇不該延續下去
    isBusy = false;
    noCandidateMode = false;

    // 換了新檔案之後，重新給一次寬限期；如果這個新檔案最後還是沒抓到
    // 候選項目，一樣要顯示「未偵測到」的按鈕，不是整頁靜悄悄沒反應。
    noCandidateTimerId = setTimeout(() => {
      noCandidateTimerId = null;
      showNoCandidateButton();
    }, NO_CANDIDATE_GRACE_MS);
  }

  // 寬限期過後還是沒有任何候選檔案時呼叫：把按鈕嵌入頁面，但用唯讀、
  // disabled 的「未偵測到 PDF」狀態顯示，而不是完全不出現。這樣使用者
  // 至少知道擴充功能「有在跑、只是這頁沒抓到東西」，而不是搞不清楚擴充
  // 功能到底有沒有作用——這也是診斷未來類似網域/偵測問題時很有用的線索
  // （使用者看得到狀態，回報問題時才會想到要講「按鈕出現了但是說未偵測到」
  // 而不是籠統地說「沒反應」）。
  function showNoCandidateButton() {
    if (dismissed || btnEl || currentCandidate) return;
    noCandidateMode = true;
    buildButton();
    startAnchorSearch();
    btnEl.title = tr("btnNotFoundTitle");
    btnEl.setAttribute("aria-label", tr("btnNotFoundTitle"));
    refreshIdleLabel();
  }

  function showButtonForCandidate(candidate) {
    if (dismissed || !candidate) return;
    if (
      currentCandidate &&
      currentCandidate.matchedBy === "keyword" &&
      candidate.matchedBy !== "keyword"
    ) {
      return;
    }

    if (noCandidateTimerId) {
      clearTimeout(noCandidateTimerId);
      noCandidateTimerId = null;
    }
    const wasNoCandidateMode = noCandidateMode;
    noCandidateMode = false;
    currentCandidate = candidate;

    if (!btnEl) {
      buildButton();
      startAnchorSearch();
    } else if (wasNoCandidateMode) {
      // 按鈕已經以「未偵測到」狀態存在（寬限期過後顯示的那顆），現在真的
      // 抓到候選檔案了，要把標題/aria-label 切回正常的下載說明文字。
      btnEl.title = tr("btnTitle");
      btnEl.setAttribute("aria-label", tr("btnTitle"));
    }
    refreshIdleLabel();
  }

  chrome.runtime.sendMessage({ type: "GET_CANDIDATES" }, (resp) => {
    const list = (resp && resp.candidates) || [];
    const best = list.find((c) => c.matchedBy === "keyword") || list[0];
    if (best) showButtonForCandidate(best);
  });

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg && msg.type === "NEW_CANDIDATE") {
      showButtonForCandidate(msg.candidate);
    } else if (msg && msg.type === "RESET_CANDIDATES") {
      resetForNewDocument();
    }
  });

  // 給偵測一個寬限期：如果這段時間內完全沒有任何候選檔案出現（不管是一
  // 開始 GET_CANDIDATES 的回應，還是之後 background.js 推送的
  // NEW_CANDIDATE），就把按鈕用「未偵測到 PDF」的唯讀狀態顯示出來，而不
  // 是讓使用者看到一片空白、猜不出擴充功能到底有沒有在運作。
  noCandidateTimerId = setTimeout(() => {
    noCandidateTimerId = null;
    showNoCandidateButton();
  }, NO_CANDIDATE_GRACE_MS);
})();
