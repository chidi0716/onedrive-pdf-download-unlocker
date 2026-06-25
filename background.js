// background.js
// Watches network requests/responses on the active tab to find the
// OneDrive/SharePoint request that actually carries the PDF bytes, and
// performs the download.
//
// Detection: that request (URL contains "passthrough", confirmed via
// DevTools) is often served with Transfer-Encoding: chunked and NO
// Content-Length header, so it can't be found by "biggest Content-Length"
// alone. Detection is keyword-first, with a large-response fallback.
//
// Download: a plain re-fetch of the URL fails because SharePoint's
// transform/passthrough endpoint requires a custom "X-SPOPacToken" header
// (a short-lived authorization token the page's own script attaches) on
// top of cookies and Referer. The browser does NOT resend that header
// automatically - it has to be captured from the original outgoing
// request and re-attached manually. We capture it via
// webRequest.onSendHeaders, then forward it (plus cookies + Referer) when
// re-fetching from this background service worker, which - unlike a
// content script - is not subject to the page's CORS policy (the same
// reason PowerShell's Invoke-WebRequest isn't CORS-restricted either).

const candidatesByTab = {}; // { [tabId]: [{url, size, time, contentType, filename, matchedBy, forwardHeaders}] }
const headersByUrl = {}; // { [url]: { [headerName]: headerValue } }
const stableFileKeyByTab = {}; // { [tabId]: string } - 「目前穩定認定」在看哪個檔案
const pendingResetTimerByTab = {}; // { [tabId]: timeoutId } - debounce 用

// 從網址裡取出可以代表「目前在看哪個檔案」的識別字串。
// onedrive.aspx 這種「檔案總管 + 內嵌預覽」頁面，常常會在載入過程中自己
// 多次用 pushState/replaceState 更新網址（例如同步資料夾路徑、預覽窗格
// 狀態），但其實還是同一個檔案；如果每次網址變動都當成「換檔案」去清掉
// 候選清單，會在偵測完成前就被自己清掉，導致按鈕永遠抓不到東西。
// 所以這裡盡量抓出「id」這個 query 參數（onedrive.aspx?id=...&parent=...
// 用的就是這個），抓不到才退回用整個路徑當識別字串。
function extractFileKey(url) {
  try {
    const u = new URL(url);
    const id = u.searchParams.get("id");
    if (id) return decodeURIComponent(id);
    return u.pathname;
  } catch (e) {
    return url;
  }
}
const MAX_CANDIDATES = 10;
const FALLBACK_MIN_SIZE = 20 * 1024; // 20 KB, only used for the generic fallback match

const URL_KEYWORDS = [
  "passthrough",
  "download.aspx",
  "getfilebycontent",
  "allowlistfiletype",
];

const EXCLUDED_TYPE_PREFIXES = [
  "text/html",
  "text/css",
  "application/javascript",
  "application/x-javascript",
  "text/javascript",
  "image/",
  "font/",
  "video/",
  "audio/",
  "text/plain",
];

// Headers worth re-attaching when we re-issue the request ourselves.
// Custom "x-..." headers and Authorization are what carry app-level auth
// tokens (like X-SPOPacToken); everything else (sec-ch-ua*, user-agent,
// cookie, etc.) is either browser-controlled or already handled separately.
function isForwardableHeader(name) {
  const n = name.toLowerCase();
  return n.startsWith("x-") || n === "authorization";
}

function getHeader(headers, name) {
  if (!headers) return "";
  const h = headers.find((x) => x.name.toLowerCase() === name);
  return h && h.value ? h.value : "";
}

function getContentLength(headers) {
  const v = getHeader(headers, "content-length");
  return v ? parseInt(v, 10) : 0;
}

function getFilenameFromHeaders(headers) {
  const value = getHeader(headers, "content-disposition");
  if (!value) return "";

  const starMatch = value.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
  if (starMatch) {
    try {
      return decodeURIComponent(starMatch[1].trim());
    } catch (e) {
      // fall through
    }
  }

  const plainMatch = value.match(/filename\s*=\s*"?([^";]+)"?/i);
  if (plainMatch) {
    return plainMatch[1].trim();
  }
  return "";
}

function isLikelyStaticAsset(url, contentType) {
  if (contentType) {
    const ct = contentType.toLowerCase();
    if (EXCLUDED_TYPE_PREFIXES.some((p) => ct.startsWith(p))) return true;
  }
  if (/\.(css|js|png|jpg|jpeg|gif|svg|woff2?|ttf|eot|ico|map)(\?|$)/i.test(url)) {
    return true;
  }
  return false;
}

// Capture request headers before they go out, so we can re-attach the
// custom auth header(s) later when re-fetching the same URL ourselves.
chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    const lowerUrl = details.url.toLowerCase();
    const isCandidateUrl = URL_KEYWORDS.some((k) => lowerUrl.includes(k));
    if (!isCandidateUrl) return;
    if (!details.requestHeaders) return;

    const captured = {};
    for (const h of details.requestHeaders) {
      if (isForwardableHeader(h.name) && h.value) {
        captured[h.name] = h.value;
      }
    }
    if (Object.keys(captured).length > 0) {
      headersByUrl[details.url] = captured;
    }
  },
  { urls: ["<all_urls>"] },
  ["requestHeaders", "extraHeaders"]
);

function addCandidate(tabId, entry) {
  if (!candidatesByTab[tabId]) candidatesByTab[tabId] = [];
  const list = candidatesByTab[tabId];
  if (list.some((c) => c.url === entry.url)) return;
  list.push(entry);
  list.sort((a, b) => {
    if (a.matchedBy !== b.matchedBy) return a.matchedBy === "keyword" ? -1 : 1;
    return b.size - a.size;
  });
  if (list.length > MAX_CANDIDATES) list.length = MAX_CANDIDATES;

  // 推送給該分頁的 content script（如果有注入的話），讓嵌入頁面的下載
  // 按鈕能即時出現/更新，不需要使用者自己點開 popup。
  //
  // 重要：這裡用 try/catch 整個包起來。chrome.tabs.sendMessage 在分頁已經
  // 關閉、或 tabId 無效時會「同步」丟出例外（不是只有透過
  // chrome.runtime.lastError 回報的非同步錯誤），如果不接住，這個例外會
  // 往外冒到 webRequest.onCompleted 的監聽器裡，可能讓 service worker
  // 中斷處理、導致之後一段時間偵測不到任何檔案（這跟之前回報的「突然抓不
  // 到部分檔案」症狀一致）。這個 try/catch 是唯一新增的保護，偵測核心邏輯
  // （上面的比對、排序）完全沒有更動。
  try {
    chrome.tabs.sendMessage(tabId, { type: "NEW_CANDIDATE", candidate: list[0] }, () => {
      // 忽略「Receiving end does not exist」之類的錯誤：代表該分頁沒有注入
      // content script（不是 OneDrive/SharePoint 網域），這是預期狀況。
      void chrome.runtime.lastError;
    });
  } catch (e) {
    // 分頁已關閉或其他原因送不出去，安靜忽略即可——候選清單已經存好了，
    // 不影響 popup 之後再用 GET_CANDIDATES 拿到結果。
  }
}

chrome.webRequest.onCompleted.addListener(
  (details) => {
    const { tabId, url, responseHeaders, statusCode, method } = details;
    if (tabId < 0) return;
    if (statusCode && (statusCode >= 400 || statusCode === 204)) return;
    if (method !== "GET" && method !== "POST") return;

    const lowerUrl = url.toLowerCase();
    const contentType = getHeader(responseHeaders, "content-type");
    const size = getContentLength(responseHeaders);
    const filename = getFilenameFromHeaders(responseHeaders);
    const forwardHeaders = headersByUrl[url] || null;

    const keywordMatch = URL_KEYWORDS.some((k) => lowerUrl.includes(k));

    if (keywordMatch) {
      addCandidate(tabId, {
        url,
        size,
        time: Date.now(),
        contentType,
        filename,
        matchedBy: "keyword",
        forwardHeaders,
      });
      return;
    }

    if (method !== "GET") return;
    if (isLikelyStaticAsset(url, contentType)) return;
    if (size < FALLBACK_MIN_SIZE) return;

    addCandidate(tabId, {
      url,
      size,
      time: Date.now(),
      contentType,
      filename,
      matchedBy: "size",
      forwardHeaders,
    });
  },
  { urls: ["<all_urls>"] },
  ["responseHeaders"]
);

chrome.tabs.onRemoved.addListener((tabId) => {
  delete candidatesByTab[tabId];
  delete stableFileKeyByTab[tabId];
  if (pendingResetTimerByTab[tabId]) {
    clearTimeout(pendingResetTimerByTab[tabId]);
    delete pendingResetTimerByTab[tabId];
  }
});

chrome.webNavigation.onCommitted.addListener((details) => {
  if (details.frameId === 0) {
    delete candidatesByTab[details.tabId];
    stableFileKeyByTab[details.tabId] = extractFileKey(details.url);
    if (pendingResetTimerByTab[details.tabId]) {
      clearTimeout(pendingResetTimerByTab[details.tabId]);
      delete pendingResetTimerByTab[details.tabId];
    }
  }
});

// OneDrive/SharePoint 在同一個分頁內切換不同檔案時（例如在檔案清單裡點
// 另一個 PDF），通常不會整頁重新載入，而是用 pushState/replaceState
// 換網址（單頁應用程式式的導覽），上面的 onCommitted 完全看不到這種
// 切換。如果不處理，舊檔案抓到的候選項目會一直留在 candidatesByTab
// 裡，新檔案打開後浮動按鈕可能還在指向「上一個檔案」的下載網址——
// 不只是視覺上不對，是真的會下載錯檔案。
// 用 onHistoryStateUpdated 偵測這種同頁切換，清掉這個分頁的候選項目，
// 並通知 content script 重置按鈕狀態（移除目前的按鈕、等新檔案的候選
// 項目進來再重新顯示）。
// 1.9.1 的版本只比對「這一次」跟「上一次」網址的識別字串，結果在某些
// 頁面（例如沒有原生下載按鈕、靠分享連結直接開啟單一檔案的那種頁面）
// 反而更容易誤判：這類頁面常常在「同一次」載入過程中，短時間內連續觸發
// 好幾次 pushState/replaceState（例如先導向到中繼網址、再正規化參數、
// 再同步預覽窗格狀態……），每一次都還是同一個檔案，但只要其中任何一次
// 的識別字串跟「前一次」不同（例如參數順序、編碼方式有差），就會被誤判
// 成「換檔案」而把剛抓到、還沒顯示出來的候選清單整個清掉——這正是
// 「預設沒有下載按鈕的頁面都抓不到」的成因：這種頁面往往是用分享連結
// 直接開啟單一檔案，初始載入時的網址正規化次數比一般「檔案總管+預覽」
// 頁面更多、更密集。
//
// 修正方式：改成 debounce（防手震）。每次網址變動只先記下「目前候選的
// 識別字串」並重新啟動一個短暫的計時器，不立刻動手清除任何東西；只有在
// 連續一段時間（600ms）都沒有再發生網址變動、整個頁面真正「安定」下來
// 之後，才去比對這個安定下來的識別字串跟「目前穩定認定」的識別字串是否
// 真的不同——真的不同才代表使用者操作換了檔案，這時才清掉候選清單、通知
// content script 重置。這樣同一次載入過程中的多次內部網址調整，不會再
// 互相打斷，只有真正的「換檔案」才會觸發重置。
chrome.webNavigation.onHistoryStateUpdated.addListener((details) => {
  if (details.frameId !== 0) return;
  const tabId = details.tabId;
  const candidateKey = extractFileKey(details.url);

  if (pendingResetTimerByTab[tabId]) {
    clearTimeout(pendingResetTimerByTab[tabId]);
  }

  pendingResetTimerByTab[tabId] = setTimeout(() => {
    delete pendingResetTimerByTab[tabId];
    const stableKey = stableFileKeyByTab[tabId];
    if (stableKey !== undefined && candidateKey === stableKey) return;

    stableFileKeyByTab[tabId] = candidateKey;
    delete candidatesByTab[tabId];
    try {
      chrome.tabs.sendMessage(tabId, { type: "RESET_CANDIDATES" }, () => {
        void chrome.runtime.lastError;
      });
    } catch (e) {
      // 分頁可能已經關閉或其他原因送不出去，安靜忽略即可
    }
  }, 600);
});

// Service workers have no DOM, so URL.createObjectURL() does not exist
// here - that is what caused "URL.createObjectURL is not a function".
// Instead we fetch the bytes here (where CORS doesn't apply and we can
// attach the captured auth header), then hand the bytes back to the
// popup as a base64 string. The popup injects a tiny script into the
// actual tab, which DOES have a DOM, to rebuild the Blob, create the
// object URL, and trigger the download there.
function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function handleDownload(msg) {
  try {
    const fetchOptions = { credentials: "include" };
    if (msg.referrer) {
      try {
        fetchOptions.referrer = msg.referrer;
      } catch (e) {
        // ignore invalid referrer value
      }
    }
    if (msg.headers && Object.keys(msg.headers).length > 0) {
      fetchOptions.headers = msg.headers;
    }

    const resp = await fetch(msg.url, fetchOptions);
    if (!resp.ok) {
      return { ok: false, status: resp.status, statusText: resp.statusText };
    }
    const buffer = await resp.arrayBuffer();
    if (!buffer || buffer.byteLength === 0) {
      return { ok: false, error: "empty response body" };
    }

    const contentType = resp.headers.get("content-type") || "application/pdf";
    const base64 = arrayBufferToBase64(buffer);
    return { ok: true, base64, contentType, size: buffer.byteLength };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg && msg.type === "GET_CANDIDATES") {
    // popup.js 會帶 tabId 過來；content.js 是從分頁內部送訊息，沒有帶
    // tabId，這時改用 sender.tab.id（背景能看到訊息是哪個分頁送來的）。
    const tabId = msg.tabId != null ? msg.tabId : sender.tab && sender.tab.id;
    sendResponse({ candidates: (tabId != null && candidatesByTab[tabId]) || [] });
    return;
  }
  if (msg && msg.type === "CLEAR_CANDIDATES") {
    delete candidatesByTab[msg.tabId];
    sendResponse({ ok: true });
    return;
  }
  if (msg && msg.type === "DOWNLOAD_FILE") {
    handleDownload(msg).then(sendResponse);
    return true; // keep the message channel open for the async response
  }
});
