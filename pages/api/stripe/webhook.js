// pages/api/stripe/webhook.js
import { buffer } from "micro";
import { stripe, PLANS } from "@/lib/stripe";
import { supabaseAdmin } from "@/lib/supabase";

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  const buf = await buffer(req);
  const sig = req.headers["stripe-signature"];
  let event;
  try {
    event = stripe.webhooks.constructEvent(buf, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    return res.status(400).json({ error: `Webhook error: ${err.message}` });
  }

  const getNextReset = () => {
    const d = new Date(); d.setMonth(d.getMonth() + 1); return d.toISOString();
  };

  if (event.type === "checkout.session.completed") {
    const session = event.data.object;
    if (session.mode !== "subscription") return res.status(200).json({ received: true });

    const { userId, planId, cvLimit } = session.metadata;
    const plan = PLANS.find(p => p.id === planId);

    await supabaseAdmin.from("users").upsert({
      clerk_id: userId,
      plan: planId,
      cv_limit: parseInt(cvLimit),
      usage_count: 0,
      usage_reset_at: getNextReset(),
      subscription_status: "active",
      stripe_subscription_id: session.subscription,
      stripe_customer_id: session.customer,
      updated_at: new Date().toISOString(),
    }, { onConflict: "clerk_id" });
  }

  // Subscription renewed — reset monthly usage
  if (event.type === "invoice.paid") {
    const invoice = event.data.object;
    if (invoice.billing_reason === "subscription_cycle") {
      const sub = await stripe.subscriptions.retrieve(invoice.subscription);
      const userId = sub.metadata?.userId;
      if (userId) {
        await supabaseAdmin.from("users").update({
          usage_count: 0,
          usage_reset_at: getNextReset(),
          subscription_status: "active",
          updated_at: new Date().toISOString(),
        }).eq("clerk_id", userId);
      }
    }
  }

  // Subscription cancelled or payment failed — deactivate
  if (event.type === "customer.subscription.deleted" || event.type === "invoice.payment_failed") {
    const obj = event.data.object;
    const subId = obj.subscription || obj.id;
    if (subId) {
      await supabaseAdmin.from("users").update({
        subscription_status: "inactive",
        updated_at: new Date().toISOString(),
      }).eq("stripe_subscription_id", subId);
    }
  }

  return res.status(200).json({ received: true });
}
