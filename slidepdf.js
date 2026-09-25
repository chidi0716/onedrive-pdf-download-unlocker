// slidepdf.js
// Minimal PDF writer: one JPEG per page, embedded as-is (DCTDecode), so no
// re-encoding and no third-party library. Used by slides.js to turn the
// captured slide images into a single PDF.
(function (global) {
  // Read width/height from the JPEG's SOFn marker.
  function jpegSize(bytes) {
    let i = 2;
    while (i + 9 < bytes.length) {
      if (bytes[i] !== 0xff) { i++; continue; }
      const marker = bytes[i + 1];
      const len = (bytes[i + 2] << 8) | bytes[i + 3];
      // SOF0..SOF15 except DHT(C4), JPG(C8), DAC(CC)
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { h: (bytes[i + 5] << 8) | bytes[i + 6], w: (bytes[i + 7] << 8) | bytes[i + 8] };
      }
      i += 2 + len;
    }
    throw new Error("not a JPEG (no SOF marker)");
  }

  // pages: array of Uint8Array (JPEG bytes). pxPerPt: how many image pixels
  // map to one PDF point (2 for a 2x capture keeps the page at on-screen size).
  function buildPdf(pages, pxPerPt) {
    const enc = new TextEncoder();
    const chunks = [];
    const offsets = [];
    let pos = 0;
    const push = (data) => {
      const b = typeof data === "string" ? enc.encode(data) : data;
      chunks.push(b);
      pos += b.length;
    };
    const obj = (n, body) => {
      offsets[n] = pos;
      push(`${n} 0 obj\n`);
      for (const part of body) push(part);
      push("\nendobj\n");
    };

    const n = pages.length;
    // Object numbers: 1 catalog, 2 pages, then 3 per page: page, image, content.
    const pageObj = (i) => 3 + i * 3;
    push("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
    obj(1, ["<< /Type /Catalog /Pages 2 0 R >>"]);
    const kids = [];
    for (let i = 0; i < n; i++) kids.push(`${pageObj(i)} 0 R`);
    obj(2, [`<< /Type /Pages /Kids [${kids.join(" ")}] /Count ${n} >>`]);

    for (let i = 0; i < n; i++) {
      const img = pages[i];
      const { w, h } = jpegSize(img);
      const pw = +(w / pxPerPt).toFixed(2);
      const ph = +(h / pxPerPt).toFixed(2);
      const p = pageObj(i);
      obj(p, [`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << /Im0 ${p + 1} 0 R >> >> /Contents ${p + 2} 0 R >>`]);
      obj(p + 1, [
        `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${img.length} >>\nstream\n`,
        img,
        "\nendstream",
      ]);
      const content = `q ${pw} 0 0 ${ph} 0 0 cm /Im0 Do Q`;
      obj(p + 2, [`<< /Length ${content.length} >>\nstream\n${content}\nendstream`]);
    }

    const total = 3 + n * 3;
    const xref = pos;
    let table = `xref\n0 ${total}\n0000000000 65535 f \n`;
    for (let k = 1; k < total; k++) table += String(offsets[k]).padStart(10, "0") + " 00000 n \n";
    push(table);
    push(`trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);

    const out = new Uint8Array(pos);
    let o = 0;
    for (const c of chunks) {
      out.set(c, o);
      o += c.length;
    }
    return out;
  }

  global.ODPDF_PDF = { buildPdf: buildPdf, jpegSize: jpegSize };
})(typeof window !== "undefined" ? window : globalThis);
