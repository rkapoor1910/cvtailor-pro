import AdmZip from "adm-zip";

export const config = {
  api: { bodyParser: { sizeLimit: "15mb" } },
};

function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function stripXml(s) {
  return s.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

// Extract ONLY bullet points and substantive descriptions worth tailoring
// Returns array of { index, text } where index = paragraph position in XML
function extractTailorableParagraphs(xml) {
  const all = [];
  const tailor = [];
  const paraMatches = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];

  paraMatches.forEach((para, index) => {
    const textMatches = para.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
    const text = textMatches
      .map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, ""))
      .join("").trim();
    all.push(text);

    if (!text) return;

    const words = text.split(/\s+/).filter(Boolean);
    const wordCount = words.length;

    // Only tailor substantive bullet points and descriptions:
    // - Must be at least 8 words (real sentence, not a label)
    // - Must not be ALL CAPS (section heading)
    // - Must not start with a year or month (date)
    // - Must not look like contact info
    // - Must not be a job title / company line (short bold headers)
    const isAllCaps = text === text.toUpperCase() && /[A-Z]/.test(text);
    const isDate = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})/i.test(text);
    const isContact = /^(email|mobile|phone|tel|linkedin|address|http|www)/i.test(text);
    const isTooShort = wordCount < 8;
    const looksLikeHeader = wordCount < 6 && /^[A-Z]/.test(text);

    if (!isAllCaps && !isDate && !isContact && !isTooShort && !looksLikeHeader) {
      tailor.push({ index, text });
    }
  });

  return { all, tailor };
}

// Extract top JD keywords only — we pass these to Claude instead of the full JD
function extractJdKeywords(jd) {
  // Strip common stop words, extract meaningful phrases
  const stopWords = new Set(["the","a","an","and","or","but","in","on","at","to","for","of","with","by","from","is","are","was","were","be","been","have","has","had","will","would","could","should","may","might","must","shall","we","our","you","your","they","their","it","its","that","this","these","those","as","if","then","than","so","do","does","did","not","no","nor","yet","both","either","each","all","any","few","more","most","other","some","such","into","through","during","before","after","above","below","between","out","off","over","under","again","further","once"]);

  const words = jd.toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(w => w.length > 3 && !stopWords.has(w));

  // Count frequency
  const freq = {};
  words.forEach(w => { freq[w] = (freq[w] || 0) + 1; });

  // Return top 30 keywords
  return Object.entries(freq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30)
    .map(([word]) => word)
    .join(", ");
}

function injectTailoredText(xml, tailoredMap) {
  let paraIndex = 0;
  return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (paraXml) => {
    const textMatches = paraXml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
    const originalText = textMatches
      .map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, ""))
      .join("").trim();

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

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { docxBase64, jd } = req.body;
  if (!docxBase64 || !jd) {
    return res.status(400).json({ error: "Missing DOCX or job description" });
  }

  try {
    const docxBuffer = Buffer.from(docxBase64, "base64");
    const zip = new AdmZip(docxBuffer);
    const docXmlEntry = zip.getEntry("word/document.xml");
    if (!docXmlEntry) {
      return res.status(400).json({ error: "Invalid DOCX" });
    }

    const originalXml = docXmlEntry.getData().toString("utf8");
    const { all, tailor } = extractTailorableParagraphs(originalXml);

    console.log(`Total: ${all.length} paras, tailoring: ${tailor.length}`);

    if (tailor.length === 0) {
      return res.status(200).json({ docxBase64 }); // nothing to tailor
    }

    // Extract only JD keywords — much shorter than full JD
    const jdKeywords = extractJdKeywords(jd);

    // Ultra-compact prompt — JSON in, JSON out
    // Only send the bullet texts, not the full CV structure
    const bulletList = tailor.map((p, i) => `${i}|${p.text}`).join("\n");

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

    console.log("AI raw response length:", rawResponse.length);

    // Parse JSON response — map number -> tailored text
    let tailoredLookup = {};
    try {
      const jsonMatch = rawResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        tailoredLookup = JSON.parse(jsonMatch[0]);
      }
    } catch (e) {
      console.error("JSON parse failed:", e.message);
      // Fallback: return original if parse fails
      return res.status(200).json({ docxBase64 });
    }

    // Build map of paragraph index -> tailored text
    const tailoredMap = {};
    tailor.forEach((p, i) => {
      const tailored = tailoredLookup[String(i)] || tailoredLookup[i];
      if (tailored && typeof tailored === "string") {
        tailoredMap[p.index] = tailored;
      }
    });

    console.log(`Injecting ${Object.keys(tailoredMap).length} tailored paragraphs`);

    // Inject tailored text back into original XML
    const newXml = injectTailoredText(originalXml, tailoredMap);

    zip.updateFile("word/document.xml", Buffer.from(newXml, "utf8"));
    const newDocxBuffer = zip.toBuffer();

    return res.status(200).json({
      docxBase64: newDocxBuffer.toString("base64"),
    });

  } catch (err) {
    console.error("DOCX tailor error:", err);
    return res.status(500).json({ error: "Failed: " + err.message });
  }
}
