import Stripe from "stripe";
export const stripe = new Stripe(process.env.STRIPE_SECRET_KEY, {
  apiVersion: "2023-10-16",
});

// Monthly subscription plans
export const PLANS = [
  {
    id: "starter",
    label: "Starter",
    price: "£3",
    pricePence: 300,
    cvLimit: 10,
    priceId: process.env.STRIPE_PRICE_STARTER,
    description: "10 tailored CVs per month",
    popular: false,
    color: "#6366f1",
  },
  {
    id: "growth",
    label: "Growth",
    price: "£5",
    pricePence: 500,
    cvLimit: 20,
    priceId: process.env.STRIPE_PRICE_GROWTH,
    description: "20 tailored CVs per month",
    popular: true,
    color: "#8b5cf6",
  },
  {
    id: "pro",
    label: "Pro",
    price: "£10",
    pricePence: 1000,
    cvLimit: 50,
    priceId: process.env.STRIPE_PRICE_PRO,
    description: "50 tailored CVs per month",
    popular: false,
    color: "#a855f7",
  },
  {
    id: "unlimited",
    label: "Unlimited",
    price: "£18",
    pricePence: 1800,
    cvLimit: 250,
    priceId: process.env.STRIPE_PRICE_UNLIMITED,
    description: "Up to 250 CVs per month",
    popular: false,
    color: "#ec4899",
  },
];
