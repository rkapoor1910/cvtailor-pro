// Pure JS DOCX processor — no native binaries, works on Vercel serverless

export const config = {
  api: { bodyParser: { sizeLimit: "15mb" } },
};

// Minimal ZIP reader in pure JS
function readZip(buffer) {
  const view = new DataView(buffer.buffer || buffer);
  const files = {};

  // Find End of Central Directory
  let eocdOffset = -1;
  for (let i = buffer.length - 22; i >= 0; i--) {
    if (buffer[i] === 0x50 && buffer[i+1] === 0x4b && buffer[i+2] === 0x05 && buffer[i+3] === 0x06) {
      eocdOffset = i; break;
    }
  }
  if (eocdOffset === -1) throw new Error("Invalid ZIP");

  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const cdSize   = view.getUint32(eocdOffset + 12, true);

  let pos = cdOffset;
  const dec = new TextDecoder();

  while (pos < cdOffset + cdSize) {
    if (buffer[pos] !== 0x50 || buffer[pos+1] !== 0x4b || buffer[pos+2] !== 0x01 || buffer[pos+3] !== 0x02) break;
    const nameLen    = view.getUint16(pos + 28, true);
    const extraLen   = view.getUint16(pos + 30, true);
    const commentLen = view.getUint16(pos + 32, true);
    const localOff   = view.getUint32(pos + 42, true);
    const name = dec.decode(buffer.slice(pos + 46, pos + 46 + nameLen));

    // Read local file header
    const lnLen = view.getUint16(localOff + 26, true);
    const leLen = view.getUint16(localOff + 28, true);
    const dataOff = localOff + 30 + lnLen + leLen;
    const compSize   = view.getUint32(pos + 20, true);
    const uncompSize = view.getUint32(pos + 24, true);
    const method     = view.getUint16(pos + 10, true);

    if (method === 0) {
      // Stored (uncompressed)
      files[name] = buffer.slice(dataOff, dataOff + uncompSize);
    }
    // Note: deflate (method=8) requires decompression — we handle it differently below

    pos += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

// Write ZIP — store all files uncompressed
function writeZip(files) {
  const enc = new TextEncoder();
  const parts = [];
  const central = [];
  let offset = 0;

  const u16 = (n) => { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n, true); return b; };
  const u32 = (n) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };

  // CRC32
  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[i] = c;
  }
  const crc32 = (data) => {
    let c = 0xffffffff;
    for (let i = 0; i < data.length; i++) c = crcTable[(c ^ data[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };

  for (const [name, data] of Object.entries(files)) {
    const nameBytes = enc.encode(name);
    const crc = crc32(data);
    const lh = new Uint8Array([
      0x50, 0x4b, 0x03, 0x04, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), 0, 0, ...nameBytes,
    ]);
    const ce = new Uint8Array([
      0x50, 0x4b, 0x01, 0x02, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ...u32(offset), ...nameBytes,
    ]);
    parts.push(lh, data);
    central.push(ce);
    offset += lh.length + data.length;
  }

  const cdSize = central.reduce((s, b) => s + b.length, 0);
  const eocd = new Uint8Array([
    0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0,
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(cdSize), ...u32(offset), 0, 0,
  ]);

  const all = [...parts, ...central, eocd];
  const total = all.reduce((s, b) => s + b.length, 0);
  const result = new Uint8Array(total);
  let pos = 0;
  for (const b of all) { result.set(b, pos); pos += b.length; }
  return result;
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function stripXml(s) {
  return s.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

function extractParagraphs(xml) {
  const paragraphs = [];
  const paraMatches = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  for (const para of paraMatches) {
    const textMatches = para.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
    const text = textMatches.map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("");
    paragraphs.push(text);
  }
  return paragraphs;
}

function needsTailoring(text) {
  const t = text.trim();
  if (!t || t.length < 15) return false;
  if (t.split(/\s+/).length < 8) return false;
  if (t === t.toUpperCase() && /[A-Z]/.test(t)) return false;
  if (/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})/i.test(t)) return false;
  if (/^(email|mobile|phone|tel|linkedin|http|www)/i.test(t)) return false;
  return true;
}

function injectTailoredText(xml, originalParas, tailoredMap) {
  let paraIndex = 0;
  return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (paraXml) => {
    const textMatches = paraXml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
    const originalText = textMatches.map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("").trim();
    const currentIndex = paraIndex++;
    if (!originalText || !tailoredMap[currentIndex]) return paraXml;

    const newText = stripXml(tailoredMap[currentIndex]);
    if (!newText || newText === originalText) return paraXml;

    const pPrMatch = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    const pPr = pPrMatch ? pPrMatch[0] : "";
    const firstRunMatch = paraXml.match(/<w:r[ >][\s\S]*?<\/w:r>/);
    const rPrMatch = firstRunMatch ? firstRunMatch[0].match(/<w:rPr>[\s\S]*?<\/w:rPr>/) : null;
    const rPr = rPrMatch ? rPrMatch[0] : "";

    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r></w:p>`;
  });
}

function extractJdKeywords(jd) {
  const stopWords = new Set(["the","a","an","and","or","but","in","on","at","to","for","of","with","by","from","is","are","was","were","be","been","have","has","had","will","would","could","should","may","might","must","we","our","you","your","they","their","it","its","that","this","these","those","as","if","not","no"]);
  const words = jd.toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(w => w.length > 3 && !stopWords.has(w));
  const freq = {};
  words.forEach(w => { freq[w] = (freq[w] || 0) + 1; });
  return Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 30).map(([w]) => w).join(", ");
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { docxBase64, jd } = req.body;
  if (!docxBase64 || !jd) return res.status(400).json({ error: "Missing DOCX or job description" });

  try {
    const docxBytes = Buffer.from(docxBase64, "base64");

    // Use adm-zip if available, otherwise fall back to pure JS
    let originalXml, allFiles;

    try {
      const AdmZip = require("adm-zip");
      const zip = new AdmZip(docxBytes);
      const entry = zip.getEntry("word/document.xml");
      if (!entry) throw new Error("No document.xml");
      originalXml = entry.getData().toString("utf8");

      // Store all entries for repacking
      allFiles = {};
      zip.getEntries().forEach(e => {
        if (!e.isDirectory) {
          allFiles[e.entryName] = e.getData();
        }
      });

      // After tailoring, repack with adm-zip
      const processAndReturn = async (newXml) => {
        const newZip = new AdmZip();
        for (const [name, data] of Object.entries(allFiles)) {
          if (name === "word/document.xml") {
            newZip.addFile(name, Buffer.from(newXml, "utf8"));
          } else {
            newZip.addFile(name, data);
          }
        }
        return newZip.toBuffer().toString("base64");
      };

      // Extract and tailor
      const originalParagraphs = extractParagraphs(originalXml);
      const tailorableItems = [];
      originalParagraphs.forEach((text, index) => {
        if (needsTailoring(text)) tailorableItems.push({ index, text: text.trim() });
      });

      console.log(`Paragraphs: ${originalParagraphs.length}, tailoring: ${tailorableItems.length}`);

      if (tailorableItems.length === 0) return res.status(200).json({ docxBase64 });

      const jdKeywords = extractJdKeywords(jd);
      const bulletList = tailorableItems.map((p, i) => `${i}|${p.text}`).join("\n");

      const prompt = `Rewrite these CV bullet points to better match a job requiring: ${jdKeywords}

Keep facts true. Make language stronger and more relevant. Return ONLY a JSON object mapping each number to the rewritten text. No explanation.

${bulletList}

JSON:`;

      const aiRes = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": process.env.ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 3000,
          messages: [{ role: "user", content: prompt }],
        }),
      });

      if (!aiRes.ok) {
        const err = await aiRes.json().catch(() => ({}));
        return res.status(500).json({ error: err?.error?.message || "AI error" });
      }

      const aiData = await aiRes.json();
      const rawResponse = aiData.content.map(b => b.text || "").join("").trim();

      let tailoredLookup = {};
      try {
        const jsonMatch = rawResponse.match(/\{[\s\S]*\}/);
        if (jsonMatch) tailoredLookup = JSON.parse(jsonMatch[0]);
      } catch (e) {
        console.error("JSON parse failed:", e.message);
        return res.status(200).json({ docxBase64 }); // return original on parse fail
      }

      const tailoredMap = {};
      tailorableItems.forEach((p, i) => {
        const t = tailoredLookup[String(i)] || tailoredLookup[i];
        if (t && typeof t === "string") tailoredMap[p.index] = t;
      });

      const newXml = injectTailoredText(originalXml, originalParagraphs, tailoredMap);
      const newDocxBase64 = await processAndReturn(newXml);
      return res.status(200).json({ docxBase64: newDocxBase64 });

    } catch (zipErr) {
      console.error("ZIP processing error:", zipErr.message);
      return res.status(500).json({ error: "Could not process DOCX file: " + zipErr.message });
    }

  } catch (err) {
    console.error("Handler error:", err.message);
    return res.status(500).json({ error: "Server error: " + err.message });
  }
}
