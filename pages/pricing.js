import { useUser } from "@clerk/nextjs";
import { useRouter } from "next/router";
import { PLANS } from "@/lib/stripe";

export default function Pricing() {
  const { isSignedIn } = useUser();
  const router = useRouter();

  const handlePlan = async (planId) => {
    if (!isSignedIn) { router.push(`/sign-up?plan=${planId}`); return; }
    const res = await fetch("/api/stripe/checkout", {
      method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({ planId }),
    });
    const data = await res.json();
    if (data.url) window.location.href = data.url;
  };

  return (
    <div style={{ minHeight:"100vh", background:"linear-gradient(135deg,#0f172a,#1e1b4b)", padding:"48px 20px", fontFamily:"system-ui,sans-serif", color:"#e2e8f0" }}>
      <div style={{ maxWidth:900, margin:"0 auto", textAlign:"center" }}>
        <div style={{ display:"inline-block", background:"linear-gradient(135deg,#6366f1,#a855f7)", borderRadius:12, padding:"7px 16px", marginBottom:16, fontSize:12, fontWeight:700, letterSpacing:".1em", color:"#fff" }}>PRICING</div>
        <h1 style={{ fontSize:"clamp(28px,5vw,44px)", fontWeight:800, background:"linear-gradient(135deg,#c7d2fe,#e9d5ff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent", marginBottom:12 }}>Simple, honest pricing</h1>
        <p style={{ color:"#94a3b8", fontSize:16, marginBottom:48 }}>Monthly subscriptions. Cancel any time. Usage resets on your billing date.</p>

        {/* Free trial card */}
        <div style={{ background:"#1e293b", border:"1px solid #334155", borderRadius:16, padding:"24px 28px", marginBottom:24, display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:16 }}>
          <div style={{ textAlign:"left" }}>
            <div style={{ fontWeight:700, fontSize:18, color:"#6ee7b7", marginBottom:4 }}>✦ Free Trial</div>
            <div style={{ color:"#94a3b8", fontSize:14 }}>Try CV Tailor Pro once — no account needed. One free tailor per device.</div>
          </div>
          <a href="/trial"><button style={{ padding:"12px 28px", borderRadius:10, background:"linear-gradient(135deg,#10b981,#059669)", color:"#fff", fontWeight:700, fontSize:14, border:"none", cursor:"pointer" }}>Try for free →</button></a>
        </div>

        {/* Plan cards */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))", gap:16 }}>
          {PLANS.map(plan => (
            <div key={plan.id} style={{ background:plan.popular?"#1e1b4b":"#1e293b", border:`1px solid ${plan.popular?"#6366f1":"#334155"}`, borderRadius:16, padding:"28px 20px", position:"relative", textAlign:"center" }}>
              {plan.popular && <div style={{ position:"absolute", top:-12, left:"50%", transform:"translateX(-50%)", background:"linear-gradient(135deg,#6366f1,#8b5cf6)", borderRadius:20, padding:"4px 14px", fontSize:11, fontWeight:700, color:"#fff", whiteSpace:"nowrap" }}>MOST POPULAR</div>}
              <div style={{ fontWeight:700, fontSize:16, color:"#c7d2fe", marginBottom:8 }}>{plan.label}</div>
              <div style={{ fontSize:36, fontWeight:800, color:"#e2e8f0", marginBottom:4 }}>{plan.price}</div>
              <div style={{ fontSize:13, color:"#64748b", marginBottom:20 }}>per month</div>
              <div style={{ background:"#0f172a", borderRadius:10, padding:"12px", marginBottom:20 }}>
                <div style={{ fontSize:22, fontWeight:800, color:plan.popular?"#a5b4fc":"#e2e8f0" }}>{plan.cvLimit}</div>
                <div style={{ fontSize:12, color:"#64748b" }}>tailored CVs per month</div>
                <div style={{ fontSize:11, color:"#475569", marginTop:4 }}>= {(plan.pricePence/plan.cvLimit)}p per CV</div>
              </div>
              <button
                onClick={() => handlePlan(plan.id)}
                style={{ width:"100%", padding:"12px", borderRadius:10, border:"none", background:plan.popular?"linear-gradient(135deg,#6366f1,#8b5cf6)":"#334155", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" }}>
                Get {plan.label} →
              </button>
            </div>
          ))}
        </div>

        <div style={{ marginTop:40, color:"#64748b", fontSize:13 }}>
          <p>All plans include: Full CV tailoring · Word download · Unlimited job descriptions per session</p>
          <p style={{ marginTop:8 }}>Secure payments via Stripe · Cancel anytime from your dashboard · Usage resets monthly</p>
        </div>
        <div style={{ marginTop:20 }}>
          <a href="/" style={{ color:"#6366f1", fontSize:14 }}>← Back to home</a>
        </div>
      </div>
    </div>
  );
}
