// Lightweight server endpoint — receives pre-extracted text paragraphs,
// calls Claude, returns tailored text as JSON. No binary processing here.
//
// NOTE: this route is public in middleware.js (Clerk auth is NOT enforced here).
// getAuth() below is optional — it tells us if the caller happens to be signed
// in, so we can later wire up per-user usage limits. Anonymous/trial callers
// are allowed through; add your device-fingerprint/trial-count check here if
// you want this route itself to enforce the one-free-trial-per-device limit
// (currently that's handled by /api/trial/check + /api/trial/use elsewhere).

import { getAuth } from "@clerk/nextjs/server";

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
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  // Fail fast and clearly if the key isn't configured, instead of a confusing
  // downstream fetch/auth error.
  if (!process.env.ANTHROPIC_API_KEY) {
    console.error("ANTHROPIC_API_KEY is not set");
    return res.status(500).json({ error: "Server misconfiguration: missing AI API key" });
  }

  // Optional — tells us if this is a signed-in user. Doesn't block anonymous
  // callers since this route is public in middleware.js.
  let userId = null;
  try {
    const auth = getAuth(req);
    userId = auth?.userId || null;
  } catch (e) {
    // getAuth can throw if Clerk context isn't available in some edge cases —
    // don't let that break the request, just proceed as anonymous.
  }

  const { paragraphs, jd } = req.body || {};

  if (!paragraphs || !Array.isArray(paragraphs) || paragraphs.length === 0) {
    return res.status(400).json({ error: "Missing or empty paragraphs array" });
  }
  if (!jd || typeof jd !== "string" || jd.trim().length < 10) {
    return res.status(400).json({ error: "Missing or too-short job description" });
  }

  // Filter to only substantive bullet points worth tailoring
  const tailorableItems = [];
  paragraphs.forEach((text, index) => {
    if (typeof text !== "string") return;
    const t = text.trim();
    if (!t || t.length < 15) return;
    if (t.split(/\s+/).length < 8) return;
    if (t === t.toUpperCase() && /[A-Z]/.test(t)) return;
    if (/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4})/i.test(t)) return;
    if (/^(email|mobile|phone|tel|linkedin|http|www)/i.test(t)) return;
    tailorableItems.push({ index, text: t });
  });

  console.log(`[tailor-docx] user=${userId || "anonymous"} received ${paragraphs.length} paragraphs, tailoring ${tailorableItems.length}`);

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

  // Guard against a hung upstream request leaving the user stuck on
  // "Processing..." indefinitely.
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30_000);

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
      signal: controller.signal,
    });

    clearTimeout(timeout);

    if (!aiRes.ok) {
      const err = await aiRes.json().catch(() => ({}));
      console.error(`[tailor-docx] Anthropic API error ${aiRes.status}:`, err?.error?.message);
      // Pass through the real upstream status instead of masking everything
      // as 500 — makes it obvious in the Network tab whether this is an auth
      // problem (401), rate limit (429), bad request (400), etc.
      return res.status(aiRes.status).json({ error: err?.error?.message || `AI error (${aiRes.status})` });
    }

    const aiData = await aiRes.json();
    const rawResponse = aiData.content.map(b => b.text || "").join("").trim();

    let tailoredLookup = {};
    try {
      const jsonMatch = rawResponse.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error("No JSON object found in AI response");
      tailoredLookup = JSON.parse(jsonMatch[0]);
    } catch (e) {
      console.error("[tailor-docx] JSON parse failed:", e.message, "Raw:", rawResponse.slice(0, 200));
      return res.status(502).json({ error: "AI returned an unparseable response — please try again" });
    }

    // Build map: original paragraph index -> tailored text
    const tailoredMap = {};
    tailorableItems.forEach((p, i) => {
      const t = tailoredLookup[String(i)] || tailoredLookup[i];
      if (t && typeof t === "string" && t.trim().length > 0) {
        tailoredMap[p.index] = t.trim();
      }
    });

    console.log(`[tailor-docx] user=${userId || "anonymous"} returning ${Object.keys(tailoredMap).length} tailored paragraphs`);
    return res.status(200).json({ tailoredMap });

  } catch (err) {
    clearTimeout(timeout);
    if (err.name === "AbortError") {
      console.error("[tailor-docx] Anthropic request timed out after 30s");
      return res.status(504).json({ error: "AI request timed out — please try again" });
    }
    console.error("[tailor-docx] Unexpected error:", err.message);
    return res.status(500).json({ error: "AI call failed: " + err.message });
  }
}
