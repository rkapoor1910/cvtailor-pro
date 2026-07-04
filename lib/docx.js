// Shared DOCX builder — used on both client and server

export function escXml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

export function parseToSections(text) {
  return text.split("\n").map(raw => {
    const l = raw.trimEnd();
    if (!l.trim())            return { type: "spacer" };
    if (/^# /.test(l))       return { type: "heading", level: 1, text: l.replace(/^#+\s*/, "") };
    if (/^#{2,3} /.test(l))  return { type: "heading", level: 2, text: l.replace(/^#+\s*/, "") };
    if (/^[-*] /.test(l))    return { type: "bullet",  text: l.replace(/^[-*]\s*/, "") };
    if (/^\*\*(.+)\*\*$/.test(l)) return { type: "bold", text: l.replace(/\*\*/g, "") };
    return { type: "para", text: l };
  });
}

export function buildDocxZip(sections) {
  const body = sections.map(s => {
    if (s.type === "heading") return `<w:p><w:pPr><w:pStyle w:val="${s.level === 1 ? "Heading1" : "Heading2"}"/></w:pPr><w:r><w:t>${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type === "bullet")  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type === "bold")    return `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type === "spacer")  return `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`;
    return `<w:p><w:r><w:t xml:space="preserve">${escXml(s.text || "")}</w:t></w:r></w:p>`;
  }).join("\n");

  const enc = new TextEncoder();
  const files = {
    "[Content_Types].xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`),
    "_rels/.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`),
    "word/styles.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="60"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="2E5090"/></w:rPr></w:style></w:styles>`),
    "word/numbering.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#x2022;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`),
    "word/_rels/document.xml.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`),
  };

  const lhs = [], cd = []; let offset = 0;
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[i] = c; }
  const crc = d => { let c = 0xffffffff; for (let i = 0; i < d.length; i++) c = t[(c ^ d[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const u16 = n => [n & 0xff, (n >> 8) & 0xff];
  const u32 = n => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  const e = new TextEncoder();
  for (const [name, data] of Object.entries(files)) {
    const nb = e.encode(name), cr = crc(data);
    const lh = [0x50, 0x4b, 0x03, 0x04, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...u32(cr), ...u32(data.length), ...u32(data.length), ...u16(nb.length), 0, 0, ...nb];
    const ce = [0x50, 0x4b, 0x01, 0x02, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...u32(cr), ...u32(data.length), ...u32(data.length), ...u16(nb.length), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...u32(offset), ...nb];
    lhs.push(new Uint8Array(lh), data); cd.push(new Uint8Array(ce)); offset += lh.length + data.length;
  }
  const cdSz = cd.reduce((s, b) => s + b.length, 0);
  const eocd = [0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, ...u16(Object.keys(files).length), ...u16(Object.keys(files).length), ...u32(cdSz), ...u32(offset), 0, 0];
  const parts = [...lhs, ...cd, new Uint8Array(eocd)];
  const tot = parts.reduce((s, p) => s + p.length, 0), res = new Uint8Array(tot); let pos = 0;
  for (const p of parts) { res.set(p, pos); pos += p.length; }
  return res;
}

export function downloadDocx(arr, filename) {
  const blob = new Blob([arr], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob), a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
}
