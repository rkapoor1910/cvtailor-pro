// pages/api/stripe/portal.js
import { getAuth } from "@clerk/nextjs/server";
import { stripe } from "@/lib/stripe";
import { supabaseAdmin } from "@/lib/supabase";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();
  const { userId } = getAuth(req);
  if (!userId) return res.status(401).json({ error: "Unauthorised" });

  const { data: user } = await supabaseAdmin
    .from("users").select("stripe_customer_id").eq("clerk_id", userId).single();

  if (!user?.stripe_customer_id)
    return res.status(400).json({ error: "No subscription found" });

  const session = await stripe.billingPortal.sessions.create({
    customer: user.stripe_customer_id,
    return_url: `${process.env.NEXT_PUBLIC_APP_URL}/dashboard`,
  });

  return res.status(200).json({ url: session.url });
}
