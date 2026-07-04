// pages/api/stripe/checkout.js
import { getAuth } from "@clerk/nextjs/server";
import { stripe, PLANS } from "@/lib/stripe";
import { supabaseAdmin } from "@/lib/supabase";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const { userId } = getAuth(req);
  if (!userId) return res.status(401).json({ error: "Unauthorised" });

  const { planId } = req.body;
  const plan = PLANS.find(p => p.id === planId);
  if (!plan) return res.status(400).json({ error: "Invalid plan" });

  // Get or create Stripe customer
  let { data: user } = await supabaseAdmin
    .from("users")
    .select("stripe_customer_id")
    .eq("clerk_id", userId)
    .single();

  let customerId = user?.stripe_customer_id;
  if (!customerId) {
    const customer = await stripe.customers.create({ metadata: { clerk_id: userId } });
    customerId = customer.id;
    await supabaseAdmin.from("users").upsert({
      clerk_id: userId,
      stripe_customer_id: customerId,
    }, { onConflict: "clerk_id" });
  }

  const session = await stripe.checkout.sessions.create({
    customer: customerId,
    payment_method_types: ["card"],
    line_items: [{ price: plan.priceId, quantity: 1 }],
    mode: "subscription",
    success_url: `${process.env.NEXT_PUBLIC_APP_URL}/dashboard?subscribed=true&plan=${plan.id}`,
    cancel_url: `${process.env.NEXT_PUBLIC_APP_URL}/pricing?cancelled=true`,
    metadata: { userId, planId: plan.id, cvLimit: String(plan.cvLimit) },
    subscription_data: { metadata: { userId, planId: plan.id, cvLimit: String(plan.cvLimit) } },
  });

  return res.status(200).json({ url: session.url });
}
