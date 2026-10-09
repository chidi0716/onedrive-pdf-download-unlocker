// util.js
// Helpers shared by popup.js, content.js and slides.js. Loaded before them
// (popup.html <script>, and first in each manifest content_scripts entry), the
// same way i18n.js is, so all three use one implementation instead of copies
// that slowly drift apart.
(function (global) {
  // Filename rules that work on both Windows and macOS:
  // - replace Windows-reserved characters \ / : * ? " < > | (macOS allows them,
  //   but replacing them does no harm)
  // - drop control characters
  // - no trailing dots or spaces (Windows rejects them)
  // - cap the length so no OS path limit is hit
  function sanitizeFilename(name) {
    name = (name || "").trim();
    name = name.replace(/[\\/:*?"<>|]/g, "_");
    name = name.replace(/[\x00-\x1f\x7f]/g, "");
    name = name.replace(/[\s.]+$/g, "");
    if (name.length > 150) name = name.slice(0, 150).trim();
    return name;
  }

  function base64ToBytes(b64) {
    const chars = atob(b64);
    const arr = new Uint8Array(chars.length);
    for (let i = 0; i < chars.length; i++) arr[i] = chars.charCodeAt(i);
    return arr;
  }

  function concatBytes(parts) {
    let total = 0;
    for (const p of parts) total += p.length;
    const out = new Uint8Array(total);
    let offset = 0;
    for (const p of parts) {
      out.set(p, offset);
      offset += p.length;
    }
    return out;
  }

  // Turn a DOWNLOAD_FILE response into the raw bytes. Small files arrive
  // inline as base64; large ones are pulled chunk-by-chunk (issue #1) so no
  // single runtime message approaches the ~64 MiB cap.
  async function bytesFromResponse(resp) {
    if (!resp.streamed) return base64ToBytes(resp.base64);
    const parts = [];
    for (let i = 0; i < resp.totalChunks; i++) {
      const chunk = await chrome.runtime.sendMessage({
        type: "GET_CHUNK",
        transferId: resp.transferId,
        index: i,
        chunkSize: resp.chunkSize,
      });
      if (!chunk || !chunk.ok) {
        throw new Error((chunk && chunk.error) || "chunk transfer failed");
      }
      parts.push(base64ToBytes(chunk.base64));
    }
    chrome.runtime.sendMessage(
      { type: "RELEASE_TRANSFER", transferId: resp.transferId },
      () => void chrome.runtime.lastError
    );
    return concatBytes(parts);
  }

  // Save a Blob through a temporary <a download> in the current document.
  function saveBlob(blob, filename, revokeAfterMs) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    (document.body || document.documentElement).appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), revokeAfterMs || 60000);
  }

  global.ODPDF_UTIL = {
    sanitizeFilename,
    base64ToBytes,
    concatBytes,
    bytesFromResponse,
    saveBlob,
  };
})(typeof window !== "undefined" ? window : globalThis);
