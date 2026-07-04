// pages/api/user/status.js
import { getAuth } from "@clerk/nextjs/server";
import { supabaseAdmin } from "@/lib/supabase";

export default async function handler(req, res) {
  const { userId } = getAuth(req);
  console.log("STATUS API - userId:", userId);
  if (!userId) return res.status(401).json({ error: "Unauthorised" });

  const { data: user, error } = await supabaseAdmin
    .from("users")
    .select("plan, cv_limit, usage_count, usage_reset_at, subscription_status")
    .eq("clerk_id", userId)
    .single();

  console.log("STATUS API - user:", user, "error:", error);

  if (!user) return res.status(200).json({ plan: null, status: "none" });

  return res.status(200).json({
    plan: user.plan,
    status: user.subscription_status,
    used: user.usage_count || 0,
    limit: user.cv_limit || 0,
    resetAt: user.usage_reset_at,
  });
}
