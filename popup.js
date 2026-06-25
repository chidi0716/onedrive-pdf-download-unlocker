// popup.js

let LANG = ODPDF_I18N.DEFAULT_LANG;
function tr(key) {
  return ODPDF_I18N.t(LANG, key);
}

function fmtSize(bytes) {
  if (!bytes || bytes <= 0) return tr("sizeUnknown");
  if (bytes >= 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  if (bytes >= 1024) return (bytes / 1024).toFixed(0) + " KB";
  return bytes + " B";
}

// 同時相容 Windows 與 macOS 的檔名規則：
// - 移除 Windows 保留字元 \ / : * ? " < > |（macOS 允許但移除無妨）
// - 移除控制字元
// - 避免結尾留下句點或空白（Windows 不允許檔名以這些字元結尾）
// - 限制長度，避免超過任一作業系統的路徑長度限制
function sanitizeFilename(name) {
  name = (name || "").trim();
  name = name.replace(/[\\/:*?"<>|]/g, "_");
  name = name.replace(/[\x00-\x1f\x7f]/g, "");
  name = name.replace(/[\s.]+$/g, "");
  if (name.length > 150) {
    name = name.slice(0, 150).trim();
  }
  return name;
}

function ensurePdfExt(name) {
  if (!name) name = "document";
  if (!/\.pdf$/i.test(name)) name += ".pdf";
  return name;
}

async function getDomFilenameGuess(tabId) {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: () => {
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
          if (el && el.textContent && el.textContent.trim().length > 0) {
            return el.textContent.trim();
          }
        }

        const extPattern = /\.(pdf|docx?|xlsx?|pptx?)$/i;

        const fuiNodes = document.querySelectorAll('[class*="fui-Text"]');
        for (const el of fuiNodes) {
          const text = el.textContent && el.textContent.trim();
          if (text && extPattern.test(text) && text.length < 200) {
            return text;
          }
        }

        const all = document.querySelectorAll("body *");
        for (const el of all) {
          if (el.children && el.children.length > 0) continue;
          const text = el.textContent && el.textContent.trim();
          if (text && extPattern.test(text) && text.length < 200) {
            return text;
          }
        }

        return "";
      },
    });
    const result = results && results[0] ? results[0].result : "";
    return result || "";
  } catch (e) {
    return "";
  }
}

// NOTE: the Content-Disposition filename that the passthrough/transform
// endpoint itself returns is often a generic internal name (we saw
// "ODBTestPassthrough.pdf" - clearly not the real document name), so it
// can't be trusted as the primary source. The page's own title bar UI is
// far more reliable, so check that FIRST. Only fall back to the server
// header (and finally the tab title) if the DOM scan finds nothing.
async function guessFilename(tab, item) {
  const domName = await getDomFilenameGuess(tab.id);
  if (domName) {
    let name = sanitizeFilename(domName);
    name = name.replace(/\s*[-|–]\s*(OneDrive|SharePoint).*$/i, "");
    return ensurePdfExt(name);
  }

  if (item.filename && !/^odbtestpassthrough$/i.test(item.filename.replace(/\.[a-z0-9]+$/i, ""))) {
    return ensurePdfExt(sanitizeFilename(item.filename));
  }

  let name = (tab.title || "document").trim();
  name = name.replace(/\s*[-|–]\s*(OneDrive|SharePoint).*$/i, "");
  name = sanitizeFilename(name);
  return ensurePdfExt(name);
}

// Step 1: background.js fetches the bytes (privileged context, not
// subject to the page's CORS policy - same reason PowerShell's
// Invoke-WebRequest isn't CORS-restricted either), and hands them back as
// a base64 string (service workers have no DOM, so they can't create
// object URLs or trigger a download themselves).
function fetchFileBytes(url, filename, referrer, headers) {
  return chrome.runtime.sendMessage({ type: "DOWNLOAD_FILE", url, filename, referrer, headers });
}

// Step 2: rebuild the Blob and trigger the actual save INSIDE the tab,
// which has a real DOM (URL.createObjectURL, <a download>, etc.).
async function saveBytesInTab(tabId, base64, filename, contentType) {
  const results = await chrome.scripting.executeScript({
    target: { tabId },
    func: (b64, name, mime) => {
      try {
        const byteChars = atob(b64);
        const byteNumbers = new Array(byteChars.length);
        for (let i = 0; i < byteChars.length; i++) {
          byteNumbers[i] = byteChars.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: mime || "application/pdf" });
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = blobUrl;
        a.download = name;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);
        return { ok: true, size: blob.size };
      } catch (e) {
        return { ok: false, error: String(e && e.message ? e.message : e) };
      }
    },
    args: [base64, filename, contentType],
  });
  return results && results[0] ? results[0].result : { ok: false, error: "no result" };
}

async function downloadFile(tabId, url, filename, referrer, headers) {
  const fetched = await fetchFileBytes(url, filename, referrer, headers);
  if (!fetched || !fetched.ok) {
    return fetched || { ok: false, error: "no response from background" };
  }
  const saved = await saveBytesInTab(tabId, fetched.base64, filename, fetched.contentType);
  if (!saved || !saved.ok) {
    return saved || { ok: false, error: "no response from page script" };
  }
  return { ok: true, size: saved.size };
}

async function render(tab, candidates) {
  const list = document.getElementById("list");
  const empty = document.getElementById("empty");
  list.innerHTML = "";

  if (!candidates || candidates.length === 0) {
    empty.style.display = "block";
    return;
  }
  empty.style.display = "none";

  for (let idx = 0; idx < candidates.length; idx++) {
    const item = candidates[idx];
    const div = document.createElement("div");
    div.className = "item";

    const filename = await guessFilename(tab, item);
    const isKeywordMatch = item.matchedBy === "keyword";
    const hasAuth = item.forwardHeaders && Object.keys(item.forwardHeaders).length > 0;

    const rowTop = document.createElement("div");
    rowTop.className = "row-top";
    const icon = document.createElement("span");
    icon.className = "file-icon";
    icon.textContent = "📄";
    const nameEl = document.createElement("span");
    nameEl.className = "filename";
    nameEl.textContent = filename;
    rowTop.appendChild(icon);
    rowTop.appendChild(nameEl);

    const badges = document.createElement("div");
    badges.className = "badges";

    const matchBadge = document.createElement("span");
    matchBadge.className = "badge " + (isKeywordMatch ? "match" : "large");
    matchBadge.textContent = isKeywordMatch ? tr("matchKeyword") : tr("matchLarge");
    badges.appendChild(matchBadge);

    const sizeBadge = document.createElement("span");
    sizeBadge.className = "badge size";
    sizeBadge.textContent = fmtSize(item.size) + (item.contentType ? " · " + item.contentType : "");
    badges.appendChild(sizeBadge);

    const authBadge = document.createElement("span");
    authBadge.className = "badge " + (hasAuth ? "auth-ok" : "auth-bad");
    authBadge.textContent = hasAuth ? tr("authOk") : tr("authBad");
    badges.appendChild(authBadge);

    const btn = document.createElement("button");
    btn.textContent = idx === 0 ? tr("downloadBest") : tr("downloadThis");
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      btn.textContent = tr("downloading");
      try {
        const res = await downloadFile(tab.id, item.url, filename, tab.url, item.forwardHeaders);
        if (res && res.ok) {
          btn.textContent = `${tr("saved")}（${fmtSize(res.size)}）`;
        } else {
          const reason = (res && (res.status || res.statusText || res.error)) || tr("unknownError");
          btn.textContent = tr("downloadFailed") + reason;
          btn.disabled = false;
        }
      } catch (e) {
        btn.textContent = tr("downloadFailed") + (e && e.message ? e.message : String(e));
        btn.disabled = false;
      }
    });

    div.appendChild(rowTop);
    div.appendChild(badges);
    div.appendChild(btn);
    list.appendChild(div);
  }
}

function refresh() {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const tab = tabs[0];
    if (!tab) return;
    chrome.runtime.sendMessage(
      { type: "GET_CANDIDATES", tabId: tab.id },
      (resp) => {
        render(tab, (resp && resp.candidates) || []).catch(() => {});
      }
    );
  });
}

function applyStaticI18n() {
  document.documentElement.lang = LANG;
  document.getElementById("hdrTitle").innerHTML =
    tr("popupTitleLine1") + "<br/>" + tr("popupTitleLine2");
  document.getElementById("helpSummary").textContent = tr("helpSummary");
  const steps = ["helpStep1", "helpStep2", "helpStep3", "helpStep4", "helpStep5"];
  const stepsEl = document.getElementById("helpSteps");
  stepsEl.innerHTML = "";
  for (const key of steps) {
    const li = document.createElement("li");
    li.innerHTML = tr(key);
    stepsEl.appendChild(li);
  }
  document.getElementById("refresh").textContent = tr("refreshBtn");
  document.getElementById("emptyText").innerHTML = tr("emptyText");
  document.getElementById("footerText").textContent = tr("footerText");
  document.getElementById("langSelect").value = LANG;
}

function initLang() {
  ODPDF_I18N.getLang((lang) => {
    LANG = lang;
    applyStaticI18n();
    refresh();
  });

  document.getElementById("langSelect").addEventListener("change", (e) => {
    const newLang = e.target.value;
    ODPDF_I18N.setLang(newLang, () => {
      LANG = newLang;
      applyStaticI18n();
      refresh();
    });
  });
}

document.getElementById("refresh").addEventListener("click", refresh);
initLang();
