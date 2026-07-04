import { useUser } from "@clerk/nextjs";
import { useRouter } from "next/router";
import { useEffect } from "react";

export default function Home() {
  const { isSignedIn } = useUser();
  const router = useRouter();
  useEffect(() => { if (isSignedIn) router.push("/dashboard"); }, [isSignedIn]);

  const s = {
    page: { minHeight:"100vh", background:"linear-gradient(135deg,#0f172a 0%,#1e1b4b 100%)", display:"flex", flexDirection:"column", alignItems:"center", padding:"40px 20px" },
    badge: { display:"inline-block", background:"linear-gradient(135deg,#6366f1,#a855f7)", borderRadius:12, padding:"7px 16px", marginBottom:20, fontSize:12, fontWeight:700, letterSpacing:".1em", color:"#fff" },
    h1: { fontSize:"clamp(32px,6vw,52px)", fontWeight:800, background:"linear-gradient(135deg,#c7d2fe,#e9d5ff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent", lineHeight:1.15, marginBottom:16, textAlign:"center" },
    sub: { color:"#94a3b8", fontSize:17, lineHeight:1.6, marginBottom:40, textAlign:"center", maxWidth:560 },
    btnPrimary: { padding:"14px 36px", borderRadius:12, background:"linear-gradient(135deg,#6366f1,#8b5cf6)", color:"#fff", fontWeight:700, fontSize:16, boxShadow:"0 4px 20px rgba(99,102,241,.4)", border:"none", cursor:"pointer" },
    btnSecondary: { padding:"14px 32px", borderRadius:12, border:"1px solid #334155", color:"#e2e8f0", fontWeight:600, fontSize:16, background:"transparent", cursor:"pointer" },
  };

  return (
    <div style={s.page}>
      <div style={{ textAlign:"center", maxWidth:700 }}>
        <div style={s.badge}>✦ AI-POWERED · FREE TRIAL</div>
        <h1 style={s.h1}>Land more interviews with a perfectly tailored CV</h1>
        <p style={s.sub}>Paste any job description. Our AI rewrites your CV to match it in seconds. Download as Word. No fluff, no fabrication.</p>
        <div style={{ display:"flex", gap:14, justifyContent:"center", flexWrap:"wrap", marginBottom:48 }}>
          <a href="/trial"><button style={s.btnPrimary}>Try it free — no signup →</button></a>
          <a href="/sign-up"><button style={s.btnSecondary}>Create account</button></a>
        </div>

        {/* How it works */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:14, marginBottom:48 }}>
          {[
            { icon:"📄", t:"Upload CV", d:"PDF, DOCX or TXT" },
            { icon:"📋", t:"Paste the JD", d:"Any job listing" },
            { icon:"✨", t:"AI tailors it", d:"In under 30 seconds" },
            { icon:"⬇", t:"Download", d:"Word doc, ready to send" },
          ].map(f => (
            <div key={f.t} style={{ background:"#1e293b", border:"1px solid #334155", borderRadius:14, padding:"18px 14px", textAlign:"center" }}>
              <div style={{ fontSize:26, marginBottom:8 }}>{f.icon}</div>
              <div style={{ fontWeight:700, fontSize:13, color:"#c7d2fe", marginBottom:3 }}>{f.t}</div>
              <div style={{ fontSize:12, color:"#64748b" }}>{f.d}</div>
            </div>
          ))}
        </div>

        {/* Pricing preview */}
        <div style={{ background:"#1e293b", border:"1px solid #334155", borderRadius:16, padding:"28px 24px" }}>
          <p style={{ fontSize:13, color:"#64748b", marginBottom:18, fontWeight:600, letterSpacing:".05em" }}>SIMPLE MONTHLY PRICING — CANCEL ANYTIME</p>
          <div style={{ display:"flex", gap:12, justifyContent:"center", flexWrap:"wrap" }}>
            {[
              { l:"Free Trial", p:"£0", d:"1 CV, one time" },
              { l:"Starter", p:"£3/mo", d:"10 CVs/month" },
              { l:"Growth", p:"£5/mo", d:"20 CVs/month", hot:true },
              { l:"Pro", p:"£10/mo", d:"50 CVs/month" },
              { l:"Unlimited", p:"£18/mo", d:"250 CVs/month" },
            ].map(p => (
              <div key={p.l} style={{ border:`1px solid ${p.hot?"#6366f1":"#334155"}`, borderRadius:12, padding:"14px 16px", minWidth:110, textAlign:"center", background:p.hot?"#1e1b4b":"transparent" }}>
                {p.hot && <div style={{ fontSize:10, color:"#a5b4fc", marginBottom:3, fontWeight:700 }}>POPULAR</div>}
                <div style={{ fontSize:13, color:"#94a3b8" }}>{p.l}</div>
                <div style={{ fontSize:20, fontWeight:800, color:"#c7d2fe", margin:"4px 0" }}>{p.p}</div>
                <div style={{ fontSize:11, color:"#64748b" }}>{p.d}</div>
              </div>
            ))}
          </div>
          <div style={{ marginTop:18 }}>
            <a href="/pricing"><button style={{ ...s.btnPrimary, fontSize:14, padding:"11px 28px" }}>See full pricing →</button></a>
          </div>
        </div>
      </div>
      <p style={{ marginTop:36, color:"#334155", fontSize:12 }}>No CV data stored · Powered by Claude AI · Built in the UK</p>
    </div>
  );
}
