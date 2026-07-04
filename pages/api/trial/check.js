// pages/api/trial/check.js
// Public route — checks if a device fingerprint has used the free trial
import { supabaseAdmin } from "@/lib/supabase";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();
  const { fingerprint } = req.body;
  if (!fingerprint) return res.status(400).json({ error: "Missing fingerprint" });

  const { data } = await supabaseAdmin
    .from("trials")
    .select("id")
    .eq("fingerprint", fingerprint)
    .single();

  return res.status(200).json({ used: !!data });
}
