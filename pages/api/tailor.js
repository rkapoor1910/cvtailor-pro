// pages/api/tailor.js
// Protected route — checks active subscription and monthly usage cap
import { getAuth } from "@clerk/nextjs/server";
import { supabaseAdmin } from "@/lib/supabase";

export const config = { api: { bodyParser: { sizeLimit: "10mb" } } };

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { userId } = getAuth(req);
  if (!userId) return res.status(401).json({ error: "Please sign in to continue." });

  const { cvText, jd, fileName, countOnly } = req.body;
  if (!cvText || !jd) return res.status(400).json({ error: "Missing CV or job description" });

  // Get user subscription
  const { data: user } = await supabaseAdmin
    .from("users")
    .select("plan, cv_limit, usage_count, usage_reset_at, stripe_customer_id, subscription_status")
    .eq("clerk_id", userId)
    .single();

  const isDev = process.env.NODE_ENV !== "production";

  if (!isDev) {
    if (!user || user.subscription_status !== "active")
      return res.status(402).json({ error: "No active subscription. Please choose a plan to continue.", redirect: "/pricing" });
  }

  // In dev mode, use generous defaults if no user record
  const cvLimit = user?.cv_limit || 999;
  const usageCount2 = user?.usage_count || 0;

  // Reset monthly usage if period has passed
  const resetAt = user?.usage_reset_at ? new Date(user.usage_reset_at) : null;
  const now = new Date();
  let usageCount = user?.usage_count || 0;

  if (user && (!resetAt || now > resetAt)) {
    const nextReset = new Date(now);
    nextReset.setMonth(nextReset.getMonth() + 1);
    await supabaseAdmin.from("users").update({
      usage_count: 0,
      usage_reset_at: nextReset.toISOString(),
    }).eq("clerk_id", userId);
    usageCount = 0;
  }

  if (!isDev && usageCount >= (user?.cv_limit || 0))
    return res.status(402).json({
      error: `Monthly limit reached (${user.cv_limit} CVs). Upgrade your plan or wait until next month.`,
      redirect: "/pricing",
    });

  // countOnly mode — just deduct usage (DOCX endpoint already called Claude)
  if (countOnly) {
    await supabaseAdmin.rpc ? null : null; // no-op placeholder
    await supabaseAdmin.from("users")
      .update({ usage_count: usageCount + 1 })
      .eq("clerk_id", userId);
    await supabaseAdmin.from("tailor_log").insert({
      clerk_id: userId, file_name: fileName || "cv", plan: user.plan,
      created_at: new Date().toISOString(),
    });
    return res.status(200).json({ ok: true });
  }

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

  // Increment usage (skip in dev if no user record)
  if (user) {
    await supabaseAdmin.from("users")
      .update({ usage_count: usageCount + 1 })
      .eq("clerk_id", userId);
  }

  // Log
  await supabaseAdmin.from("tailor_log").insert({
    clerk_id: userId,
    file_name: fileName || "cv",
    plan: user?.plan || "dev",
    created_at: new Date().toISOString(),
  });

  return res.status(200).json({
    tailored,
    usage: { used: usageCount + 1, limit: user?.cv_limit || 999 },
  });
}
