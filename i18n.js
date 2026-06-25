// i18n.js
// 共用的語言字串表，同時被 popup.html（透過 <script>）與 content.js
// （透過 manifest.json 的 content_scripts）載入，所以兩邊永遠用同一份
// 翻譯，不會各自維護一份造成不一致。
// 預設語言是英文（en），使用者可以在 popup 裡切換成中文（zh-Hant），
// 選擇會存進 chrome.storage.local，下次開啟時會記住。
(function (global) {
  const STORAGE_KEY = "odpdf_lang";
  const DEFAULT_LANG = "en";

  const STRINGS = {
    en: {
      popupTitleLine1: "OneDrive / SharePoint",
      popupTitleLine2: "PDF Downloader",
      langLabel: "Language",
      helpSummary: "How to use",
      helpStep1: "Open the OneDrive / SharePoint <b>PDF preview tab</b>",
      helpStep2: "Press <b>F5 to reload</b> that tab (the extension can only see requests made after it's active)",
      helpStep3: "Wait for the page to <b>fully load</b> (the PDF content shows up)",
      helpStep4: "Come back here and click <b>\"Rescan\"</b> below",
      helpStep5: "Pick the file to download from the list (sorted by size — <b>the first one is usually the target file</b>)",
      refreshBtn: "🔄 Rescan",
      emptyText: "No file requests detected yet.<br/>Reload the PDF preview tab, wait for it to fully load, then try again.",
      footerText: "Detected requests are only kept for the current tab and are cleared automatically when it's closed.",
      sizeUnknown: "size unknown (streamed)",
      matchKeyword: "⭐ Likely target",
      matchLarge: "📦 Large file",
      authOk: "🔑 Auth header captured",
      authBad: "⚠️ No auth header",
      downloadBest: "Download this file (best match)",
      downloadThis: "Download this file",
      downloading: "Downloading…",
      saved: "Saved ✓",
      downloadFailed: "Download failed: ",
      unknownError: "Unknown error",
      btnIdleShort: "Download",
      btnIdleFull: "Download PDF",
      btnDownloading: "Downloading…",
      btnSaved: "Saved ✓",
      btnFailed: "Failed: ",
      btnTitle: "Download this file with the extension (original filename)",
      closeTitle: "Don't show this time",
      btnNotFound: "PDF not detected",
      btnNotFoundTitle: "No downloadable file detected on this page yet. If the page just loaded, wait a moment; otherwise try reloading.",
    },
    "zh-Hant": {
      popupTitleLine1: "OneDrive / SharePoint",
      popupTitleLine2: "PDF 下載器",
      langLabel: "語言",
      helpSummary: "使用說明",
      helpStep1: "開啟 OneDrive / SharePoint 的 <b>PDF 預覽分頁</b>",
      helpStep2: "按 <b>F5 重新整理</b> 該分頁（擴充功能只能看到啟用後發出的新請求）",
      helpStep3: "等待頁面 <b>完全載入</b>（PDF 內容顯示出來）",
      helpStep4: "回到這裡點下方 <b>「重新掃描」</b>",
      helpStep5: "從清單中選擇要下載的檔案（已依大小排序，<b>第一個通常就是目標檔案</b>）",
      refreshBtn: "🔄 重新掃描",
      emptyText: "尚未偵測到任何檔案內容請求。<br/>請重新整理 PDF 預覽分頁，等待完全載入後再試一次。",
      footerText: "偵測到的請求只會保留在目前分頁，關閉分頁後會自動清除",
      sizeUnknown: "大小未知（串流）",
      matchKeyword: "⭐ 可能是目標",
      matchLarge: "📦 大型檔案",
      authOk: "🔑 已取得授權標頭",
      authBad: "⚠️ 未取得授權標頭",
      downloadBest: "下載這個檔案（最佳匹配）",
      downloadThis: "下載這個檔案",
      downloading: "下載中…",
      saved: "已儲存 ✓",
      downloadFailed: "下載失敗：",
      unknownError: "未知錯誤",
      btnIdleShort: "下載",
      btnIdleFull: "下載 PDF",
      btnDownloading: "下載中…",
      btnSaved: "已儲存 ✓",
      btnFailed: "失敗：",
      btnTitle: "用擴充功能下載這個檔案（原始檔名）",
      closeTitle: "這次先不顯示",
      btnNotFound: "未偵測到 PDF",
      btnNotFoundTitle: "這個頁面目前還沒偵測到可下載的檔案。如果頁面剛載入完，稍等一下；不然可以試著重新整理。",
    },
  };

  function getLang(callback) {
    try {
      chrome.storage.local.get([STORAGE_KEY], (res) => {
        const lang = (res && res[STORAGE_KEY]) || DEFAULT_LANG;
        callback(STRINGS[lang] ? lang : DEFAULT_LANG);
      });
    } catch (e) {
      callback(DEFAULT_LANG);
    }
  }

  function setLang(lang, callback) {
    try {
      chrome.storage.local.set({ [STORAGE_KEY]: lang }, () => {
        if (callback) callback();
      });
    } catch (e) {
      if (callback) callback();
    }
  }

  function t(lang, key) {
    const dict = STRINGS[lang] || STRINGS[DEFAULT_LANG];
    if (dict && dict[key] != null) return dict[key];
    return (STRINGS[DEFAULT_LANG] && STRINGS[DEFAULT_LANG][key]) || key;
  }

  global.ODPDF_I18N = {
    STORAGE_KEY: STORAGE_KEY,
    DEFAULT_LANG: DEFAULT_LANG,
    LANGS: Object.keys(STRINGS),
    getLang: getLang,
    setLang: setLang,
    t: t,
  };
})(typeof window !== "undefined" ? window : this);
