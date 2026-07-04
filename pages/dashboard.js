import { useUser, UserButton } from "@clerk/nextjs";
import { useState, useEffect, useRef, useCallback } from "react";
import { useRouter } from "next/router";

function parsePreview(text) {
  return text.split("\n").map((line, i) => {
    if (!line.trim()) return <div key={i} style={{ height:7 }} />;
    if (/^# /.test(line))    return <h1 key={i} style={{ fontSize:19, color:"#1F3864", borderBottom:"2px solid #6366f1", paddingBottom:3, marginBottom:7 }}>{line.replace(/^# /,"")}</h1>;
    if (/^## /.test(line))   return <h2 key={i} style={{ fontSize:14, color:"#2E5090", marginTop:14, marginBottom:3 }}>{line.replace(/^## /,"")}</h2>;
    if (/^### /.test(line))  return <h3 key={i} style={{ fontSize:13, color:"#334155", marginTop:9 }}>{line.replace(/^### /,"")}</h3>;
    if (/^[-*] /.test(line)) return <div key={i} style={{ paddingLeft:18, position:"relative" }}><span style={{ position:"absolute", left:5 }}>•</span>{line.replace(/^[-*] /,"")}</div>;
    return <p key={i} style={{ margin:"3px 0" }}>{line}</p>;
  });
}

function buildDocxAndDownload(text, filename) {
  const escXml = s => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
  const sections = text.split("\n").map(raw => {
    const l = raw.trimEnd();
    if (!l.trim()) return { type:"spacer" };
    if (/^# /.test(l))       return { type:"heading", level:1, text:l.replace(/^#+\s*/,"") };
    if (/^#{2,3} /.test(l))  return { type:"heading", level:2, text:l.replace(/^#+\s*/,"") };
    if (/^[-*] /.test(l))    return { type:"bullet", text:l.replace(/^[-*]\s*/,"") };
    if (/^\*\*(.+)\*\*$/.test(l)) return { type:"bold", text:l.replace(/\*\*/g,"") };
    return { type:"para", text:l };
  });
  const body = sections.map(s => {
    if (s.type==="heading") return `<w:p><w:pPr><w:pStyle w:val="${s.level===1?"Heading1":"Heading2"}"/></w:pPr><w:r><w:t>${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type==="bullet")  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type==="bold")    return `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type==="spacer")  return `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`;
    return `<w:p><w:r><w:t xml:space="preserve">${escXml(s.text||"")}</w:t></w:r></w:p>`;
  }).join("\n");
  const enc = new TextEncoder();
  const files = {
    "[Content_Types].xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`),
    "_rels/.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`),
    "word/styles.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="60"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="2E5090"/></w:rPr></w:style></w:styles>`),
    "word/numbering.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#x2022;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`),
    "word/_rels/document.xml.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`),
  };
  const lhs=[], cd=[]; let offset=0;
  const t=new Uint32Array(256);
  for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=c&1?0xedb88320^(c>>>1):c>>>1;t[i]=c;}
  const crc=d=>{let c=0xffffffff;for(let i=0;i<d.length;i++)c=t[(c^d[i])&0xff]^(c>>>8);return(c^0xffffffff)>>>0;};
  const u16=n=>[n&0xff,(n>>8)&0xff],u32=n=>[n&0xff,(n>>8)&0xff,(n>>16)&0xff,(n>>24)&0xff];
  const e=new TextEncoder();
  for(const[name,data] of Object.entries(files)){
    const nb=e.encode(name),cr=crc(data);
    const lh=[0x50,0x4b,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(cr),...u32(data.length),...u32(data.length),...u16(nb.length),0,0,...nb];
    const ce=[0x50,0x4b,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(cr),...u32(data.length),...u32(data.length),...u16(nb.length),0,0,0,0,0,0,0,0,0,0,0,0,...u32(offset),...nb];
    lhs.push(new Uint8Array(lh),data);cd.push(new Uint8Array(ce));offset+=lh.length+data.length;
  }
  const cdSz=cd.reduce((s,b)=>s+b.length,0);
  const eocd=[0x50,0x4b,0x05,0x06,0,0,0,0,...u16(Object.keys(files).length),...u16(Object.keys(files).length),...u32(cdSz),...u32(offset),0,0];
  const parts=[...lhs,...cd,new Uint8Array(eocd)];
  const tot=parts.reduce((s,p)=>s+p.length,0),res=new Uint8Array(tot);let pos=0;
  for(const p of parts){res.set(p,pos);pos+=p.length;}
  const blob=new Blob([res],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  const url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download=filename;document.body.appendChild(a);a.click();
  setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},1000);
}

export default function Dashboard() {
  const { user } = useUser();
  const router = useRouter();
  const [userStatus, setUserStatus] = useState(null);
  const [step, setStep] = useState(1);
  const [cvFile, setCvFile] = useState(null);
  const [cvText, setCvText] = useState("");
  const [docxBase64, setDocxBase64] = useState(null);
  const [jd, setJd] = useState("");
  const [tailored, setTailored] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();

  useEffect(() => {
    fetch("/api/user/status").then(r=>r.json()).then(setUserStatus);
  }, []);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    const isDocx = /\.docx?$/i.test(file.name) || file.type.includes("wordprocessingml");
    const isTxt  = file.name.endsWith(".txt") || file.type === "text/plain";
    const isPdf  = file.name.endsWith(".pdf") || file.type === "application/pdf";
    if (!isDocx && !isTxt && !isPdf) { setError("Please upload PDF, DOCX, or TXT."); return; }
    setError(""); setCvFile(file); setDocxBase64(null);

    if (isTxt) {
      const r = new FileReader(); r.onload = e => setCvText(e.target.result); r.readAsText(file);
    } else if (isDocx) {
      const r = new FileReader();
      r.onload = e => {
        const b64 = e.target.result.split(",")[1];
        setDocxBase64(b64);
        setCvText("__DOCX_FORMAT_PRESERVE__");
      };
      r.readAsDataURL(file);
    } else if (isPdf) {
      const r = new FileReader(); r.onload = e => setCvText("__PDF__" + e.target.result.split(",")[1]); r.readAsDataURL(file);
    }
  }, []);

  const tailorCV = async () => {
    if (!cvText || !jd.trim()) return;
    setLoading(true); setError("");
    try {
      const filename = (cvFile?.name ? cvFile.name.replace(/\.[^.]+$/, "") + "_tailored" : "tailored_cv") + ".docx";

      if (cvText === "__DOCX_FORMAT_PRESERVE__" && docxBase64) {
        // Format-preserving DOCX mode
        const res = await fetch("/api/tailor-docx", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ docxBase64, jd }),
        });
        const data = await res.json();
        if (!res.ok) {
          if (data.redirect) router.push(data.redirect);
          throw new Error(data.error);
        }

        // Download returned DOCX
        const binaryStr = atob(data.docxBase64);
        const bytes = new Uint8Array(binaryStr.length);
        for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
        const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);

        setTailored("__DOCX__");
        setUserStatus(s => s ? { ...s, used: (s.used || 0) + 1 } : s);
        setStep(4);

        // Deduct usage via tailor endpoint with a lightweight call
        await fetch("/api/tailor", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ cvText: "(docx)", jd: "(processed)", fileName: cvFile?.name, countOnly: true }),
        });

      } else {
        // Plain text / PDF mode
        const res = await fetch("/api/tailor", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ cvText, jd, fileName: cvFile?.name }),
        });
        const data = await res.json();
        if (!res.ok) {
          if (data.redirect) router.push(data.redirect);
          throw new Error(data.error);
        }
        setTailored(data.tailored);
        setUserStatus(s => s ? { ...s, used: (s.used || 0) + 1 } : s);
        setStep(4);

        setTimeout(() => {
          buildDocxAndDownload(data.tailored, filename);
        }, 300);
      }

    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  };

  const manageSubscription = async () => {
    const res = await fetch("/api/stripe/portal", { method:"POST" });
    const data = await res.json();
    if (data.url) window.location.href = data.url;
  };

  const S = {
    page: { minHeight:"100vh", background:"linear-gradient(135deg,#0f172a,#1e1b4b)", fontFamily:"system-ui,sans-serif", color:"#e2e8f0" },
    nav: { display:"flex", alignItems:"center", justifyContent:"space-between", padding:"14px 24px", borderBottom:"1px solid #1e293b" },
    logo: { fontWeight:800, fontSize:18, background:"linear-gradient(135deg,#c7d2fe,#e9d5ff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" },
    main: { maxWidth:720, margin:"0 auto", padding:"28px 20px" },
    card: { background:"#1e293b", border:"1px solid #334155", borderRadius:16, padding:"22px 26px", marginBottom:16 },
    h2: { fontSize:16, fontWeight:700, color:"#c7d2fe", marginBottom:5 },
    desc: { fontSize:13, color:"#64748b", marginBottom:16 },
    dz: { border:"2px dashed #334155", borderRadius:12, padding:"32px 20px", textAlign:"center", cursor:"pointer", background:"#0f172a", transition:"all .2s" },
    ta: { width:"100%", minHeight:200, background:"#0f172a", border:"1px solid #334155", borderRadius:10, color:"#e2e8f0", fontSize:13, padding:"13px 15px", resize:"vertical", outline:"none", fontFamily:"inherit", lineHeight:1.6 },
    btnP: { padding:"11px 26px", borderRadius:10, border:"none", background:"linear-gradient(135deg,#6366f1,#8b5cf6)", color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer" },
    btnS: { padding:"11px 20px", borderRadius:10, border:"none", background:"#334155", color:"#fff", fontWeight:600, fontSize:13, cursor:"pointer" },
    btnSu: { padding:"11px 26px", borderRadius:10, border:"none", background:"linear-gradient(135deg,#10b981,#059669)", color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer" },
    err: { background:"#450a0a", border:"1px solid #b91c1c", borderRadius:10, padding:"11px 15px", color:"#fca5a5", fontSize:13, marginBottom:14 },
    prev: { background:"#fff", color:"#1e293b", borderRadius:10, padding:"22px 26px", maxHeight:400, overflowY:"auto", fontSize:13, lineHeight:1.7, fontFamily:"Calibri,Arial,sans-serif", marginBottom:16 },
  };

  const hasActivePlan = userStatus?.status === "active";
  const STEPS = ["Upload CV","Paste JD","Tailor","Download"];

  return (
    <div style={S.page}>
      <nav style={S.nav}>
        <span style={S.logo}>CV Tailor Pro</span>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          {hasActivePlan && userStatus && (
            <div style={{ background:"#1e293b", border:"1px solid #334155", borderRadius:20, padding:"5px 14px", fontSize:12 }}>
              <span style={{ color:"#94a3b8" }}>{userStatus.plan?.toUpperCase()} · </span>
              <span style={{ color:"#6ee7b7", fontWeight:700 }}>{userStatus.used || 0}/{userStatus.limit}</span>
              <span style={{ color:"#64748b" }}> this month</span>
            </div>
          )}
          {hasActivePlan && <button onClick={manageSubscription} style={{ ...S.btnS, padding:"6px 14px", fontSize:12 }}>Manage plan</button>}
          {!hasActivePlan && <a href="/pricing"><button style={{ ...S.btnP, padding:"6px 14px", fontSize:12 }}>Choose plan →</button></a>}
          <UserButton afterSignOutUrl="/" />
        </div>
      </nav>

      <div style={S.main}>
        {/* No plan banner */}
        {!hasActivePlan && (
          <div style={{ background:"linear-gradient(135deg,#1e1b4b,#1e293b)", border:"1px solid #6366f1", borderRadius:12, padding:"16px 20px", marginBottom:18, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
            <div>
              <p style={{ fontWeight:700, color:"#a5b4fc", marginBottom:3 }}>No active subscription</p>
              <p style={{ fontSize:13, color:"#64748b" }}>Choose a plan to start tailoring CVs. From £3/month.</p>
            </div>
            <a href="/pricing"><button style={S.btnP}>See plans →</button></a>
          </div>
        )}

        {/* Steps */}
        <div style={{ display:"flex", gap:10, marginBottom:18, flexWrap:"wrap" }}>
          {STEPS.map((s,i) => {
            const n=i+1, done=step>n, active=step===n;
            return <div key={s} style={{ display:"flex", alignItems:"center", gap:7 }}>
              <div style={{ width:26, height:26, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700, fontSize:12, color:"#fff", background:done?"#10b981":active?"#6366f1":"#334155" }}>{done?"✓":n}</div>
              <span style={{ fontSize:12, fontWeight:active||done?600:400, color:active?"#c7d2fe":done?"#6ee7b7":"#64748b" }}>{s}</span>
            </div>;
          })}
        </div>
        <div style={{ height:3, background:"#1e293b", borderRadius:99, marginBottom:20 }}>
          <div style={{ height:"100%", borderRadius:99, background:"linear-gradient(90deg,#6366f1,#a855f7)", width:`${((step-1)/3)*100}%`, transition:"width .5s" }} />
        </div>

        {error && <div style={S.err}>⚠ {error} {error.includes("plan") && <a href="/pricing" style={{ color:"#f87171", textDecoration:"underline", marginLeft:6 }}>Choose plan →</a>}</div>}

        {step===1 && <div style={S.card}>
          <h2 style={S.h2}>Upload your CV</h2>
          <p style={S.desc}>PDF, DOCX, or TXT</p>
          <div style={{ ...S.dz, borderColor:cvText?"#10b981":"#334155", background:cvText?"#0d2318":"#0f172a" }}
            onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files[0]);}}>
            <div style={{ fontSize:28, marginBottom:8 }}>{cvText?"✅":"📄"}</div>
            {cvFile ? <><p style={{ color:cvText?"#6ee7b7":"#fbbf24", fontWeight:600 }}>{cvFile.name}</p><p style={{ fontSize:12, color:"#475569" }}>{cvText?"Ready ✓ · Click to replace":"Extracting..."}</p></>
              : <><p style={{ color:"#94a3b8", fontWeight:500 }}>Drag & drop or click to browse</p><p style={{ fontSize:12, color:"#475569" }}>PDF · DOCX · TXT</p></>}
          </div>
          <input ref={fileRef} type="file" accept=".pdf,.docx,.doc,.txt" style={{ display:"none" }} onChange={e=>handleFile(e.target.files[0])} />
          <div style={{ display:"flex", justifyContent:"flex-end", marginTop:14 }}>
            <button style={{ ...S.btnP, opacity:(!cvFile||!cvText)?.5:1 }} disabled={!cvFile||!cvText} onClick={()=>setStep(2)}>Continue →</button>
          </div>
        </div>}

        {step===2 && <div style={S.card}>
          <h2 style={S.h2}>Paste job description</h2>
          <p style={S.desc}>Copy and paste the full job listing</p>
          <textarea style={S.ta} value={jd} onChange={e=>setJd(e.target.value)} placeholder="Paste the job description here..." />
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:8 }}>
            <span style={{ fontSize:12, color:"#475569" }}>{jd.length} chars</span>
            <div style={{ display:"flex", gap:10 }}>
              <button style={S.btnS} onClick={()=>setStep(1)}>← Back</button>
              <button style={{ ...S.btnP, opacity:jd.trim().length<50?.5:1 }} disabled={jd.trim().length<50} onClick={()=>setStep(3)}>Continue →</button>
            </div>
          </div>
        </div>}

        {step===3 && <div style={S.card}>
          <h2 style={S.h2}>Tailor your CV</h2>
          <p style={S.desc}>AI matches your CV to the job description</p>
          <div style={{ background:"#0f172a", borderRadius:10, padding:"12px 16px", border:"1px solid #1e293b", marginBottom:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:5 }}>
              <span style={{ fontSize:12, color:"#64748b" }}>CV</span><span style={{ fontSize:12, color:"#6ee7b7" }}>✓ Ready</span>
            </div>
            <p style={{ fontWeight:600, color:"#e2e8f0", fontSize:13 }}>{cvFile?.name}</p>
            <div style={{ marginTop:8, paddingTop:8, borderTop:"1px solid #1e293b" }}>
              <span style={{ fontSize:11, color:"#64748b" }}>JD PREVIEW</span>
              <p style={{ marginTop:3, color:"#94a3b8", fontSize:12 }}>{jd.slice(0,100)}…</p>
            </div>
          </div>
          {userStatus && hasActivePlan && (
            <div style={{ background:"#0f172a", borderRadius:8, padding:"10px 14px", border:"1px solid #334155", marginBottom:14, display:"flex", justifyContent:"space-between" }}>
              <span style={{ fontSize:12, color:"#94a3b8" }}>Monthly usage</span>
              <span style={{ fontSize:12, fontWeight:700, color:(userStatus.used||0) >= userStatus.limit ? "#fca5a5" : "#6ee7b7" }}>{userStatus.used||0} / {userStatus.limit} CVs used</span>
            </div>
          )}
          {loading && <div style={{ background:"#1e1b4b", borderRadius:10, padding:"14px", border:"1px solid #4338ca", marginBottom:14, textAlign:"center" }}>
            <div style={{ width:24, height:24, border:"3px solid #4338ca", borderTopColor:"#818cf8", borderRadius:"50%", animation:"spin .8s linear infinite", margin:"0 auto 8px" }} />
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
            <p style={{ color:"#a5b4fc", fontSize:13 }}>AI is tailoring your CV…</p>
          </div>}
          <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
            <button style={S.btnS} disabled={loading} onClick={()=>setStep(2)}>← Back</button>
            <button style={{ ...S.btnP, opacity:loading?.5:1 }} disabled={loading} onClick={tailorCV}>
              {loading?"Processing…":"✦ Tailor My CV"}
            </button>
          </div>
        </div>}

        {step===4 && <div style={S.card}>
          <div style={{ textAlign:"center", padding:"16px 0" }}>
            <div style={{ fontSize:44, marginBottom:14 }}>✅</div>
            <h2 style={{ ...S.h2, color:"#6ee7b7", fontSize:18, textAlign:"center" }}>Your tailored CV has downloaded</h2>
            <p style={{ ...S.desc, textAlign:"center", marginBottom:18 }}>Check your Downloads folder for the Word document.</p>
            <div style={{ display:"flex", gap:10, justifyContent:"center" }}>
              <button style={S.btnS} onClick={()=>{setStep(3);setTailored("");}}>← Re-tailor</button>
              <button style={S.btnSu} onClick={()=>buildDocxAndDownload(tailored,(cvFile?cvFile.name.replace(/\.[^.]+$/,"")+"_tailored":"tailored_cv")+".docx")}>⬇ Download again</button>
            </div>
          </div>
        </div>}
      </div>
    </div>
  );
}
