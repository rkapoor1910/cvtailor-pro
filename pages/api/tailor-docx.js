// Lightweight server endpoint — receives pre-extracted text paragraphs,
// calls Claude, returns tailored text as JSON. No binary processing here.

export const config = {
  api: { bodyParser: { sizeLimit: "5mb" } },
};

function extractJdKeywords(jd) {
  const stopWords = new Set(["the","a","an","and","or","but","in","on","at","to","for","of","with","by","from","is","are","was","were","be","been","have","has","had","will","would","could","should","may","might","must","we","our","you","your","they","their","it","its","that","this","these","those","as","if","not","no","do","does","did"]);
  const words = jd.toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(w => w.length > 3 && !stopWords.has(w));
  const freq = {};
  words.forEach(w => { freq[w] = (freq[w] || 0) + 1; });
  return Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 30).map(([w]) => w).join(", ");
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { paragraphs, jd } = req.body;

  if (!paragraphs || !Array.isArray(paragraphs) || !jd) {
    return res.status(400).json({ error: "Missing paragraphs or job description" });
  }

  // Filter to only substantive bullet points worth tailoring
  const tailorableItems = [];
  paragraphs.forEach((text, index) => {
    const t = text.trim();
    if (!t || t.length < 15) return;
    if (t.split(/\s+/).length < 8) return;
    if (t === t.toUpperCase() && /[A-Z]/.test(t)) return;
    if (/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})/i.test(t)) return;
    if (/^(email|mobile|phone|tel|linkedin|http|www)/i.test(t)) return;
    tailorableItems.push({ index, text: t });
  });

  console.log(`Received ${paragraphs.length} paragraphs, tailoring ${tailorableItems.length}`);

  if (tailorableItems.length === 0) {
    return res.status(200).json({ tailoredMap: {} });
  }

  const jdKeywords = extractJdKeywords(jd);
  const bulletList = tailorableItems.map((p, i) => `${i}|${p.text}`).join("\n");

  const prompt = `Rewrite these CV bullet points to better match a job requiring: ${jdKeywords}

Keep all facts true. Make language stronger and more relevant to the role.
Return ONLY a valid JSON object mapping each number to the rewritten text.
No explanation, no markdown, just the JSON.

${bulletList}

JSON:`;

  try {
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
      console.error("JSON parse failed:", e.message, "Raw:", rawResponse.slice(0, 200));
      return res.status(200).json({ tailoredMap: {} });
    }

    // Build map: original paragraph index -> tailored text
    const tailoredMap = {};
    tailorableItems.forEach((p, i) => {
      const t = tailoredLookup[String(i)] || tailoredLookup[i];
      if (t && typeof t === "string" && t.trim().length > 0) {
        tailoredMap[p.index] = t.trim();
      }
    });

    console.log(`Returning ${Object.keys(tailoredMap).length} tailored paragraphs`);
    return res.status(200).json({ tailoredMap });

  } catch (err) {
    console.error("Claude error:", err.message);
    return res.status(500).json({ error: "AI call failed: " + err.message });
  }
}
