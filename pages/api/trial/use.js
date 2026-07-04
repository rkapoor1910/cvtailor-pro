// pages/api/trial/use.js
// Public route — runs the one free trial tailor and marks device as used
import { supabaseAdmin } from "@/lib/supabase";

export const config = { api: { bodyParser: { sizeLimit: "10mb" } } };

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { fingerprint, cvText, jd, fileName, markOnly } = req.body;
  if (!fingerprint) return res.status(400).json({ error: "Missing fingerprint" });

  // Check trial not already used
  const { data: existing } = await supabaseAdmin
    .from("trials")
    .select("id")
    .eq("fingerprint", fingerprint)
    .single();

  if (existing)
    return res.status(403).json({ error: "Free trial already used on this device. Please sign up to continue." });

  // markOnly mode — just record usage (DOCX endpoint already called Claude)
  if (markOnly) {
    await supabaseAdmin.from("trials").insert({
      fingerprint,
      ip: req.headers["x-forwarded-for"] || req.socket.remoteAddress || "unknown",
      file_name: fileName || "cv",
      created_at: new Date().toISOString(),
    });
    return res.status(200).json({ ok: true });
  }

  if (!cvText || !jd) return res.status(400).json({ error: "Missing required fields" });

  // Call Anthropic
  const INSTRUCTIONS = `INSTRUCTIONS:
- Reorder sections and bullet points to best match the JD keywords and requirements
- Emphasise relevant experience, skills, and achievements that align with the role
- Keep ALL original content — do not fabricate anything
- IMPORTANT: Do not copy phrases directly from the job description. Rephrase and integrate relevant language naturally into the candidate's own voice and existing achievements, so it reads as authentic experience rather than a direct echo of the JD wording
- Vary sentence structure and word choice so the result does not look templated or formulaic
- Use markdown: # for name/header, ## for section headings, ### for job titles/roles
- Use bullet points (- ) for all list items
- Output the FULL tailored CV in markdown only. No preamble or explanation.

JOB DESCRIPTION:
${jd}`;

  const aiRes = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 6000,
      messages: [{ role: "user", content: `You are a professional CV writer. Here is the candidate's full CV:\n\n${cvText}\n\nTailor this CV for the job below.\n\n${INSTRUCTIONS}` }],
    }),
  });

  if (!aiRes.ok) {
    const err = await aiRes.json().catch(() => ({}));
    return res.status(500).json({ error: err?.error?.message || "AI error" });
  }

  const aiData = await aiRes.json();
  const tailored = aiData.content.map(b => b.text || "").join("");

  // Mark trial as used (store fingerprint + IP for anti-abuse)
  await supabaseAdmin.from("trials").insert({
    fingerprint,
    ip: req.headers["x-forwarded-for"] || req.socket.remoteAddress || "unknown",
    file_name: fileName || "cv",
    created_at: new Date().toISOString(),
  });

  return res.status(200).json({ tailored });
}
