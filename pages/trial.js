import { useState, useRef, useCallback, useEffect } from "react";

// -- Device fingerprint ------------------------------------------------------
function getFingerprint() {
  if (typeof window === "undefined") return "ssr";
  try {
    const stored = localStorage.getItem("cvtailor_device_id");
    if (stored) return stored;
  } catch (e) {}
  const parts = [navigator.userAgent, navigator.language, screen.width + "x" + screen.height, new Date().getTimezoneOffset(), navigator.hardwareConcurrency || 0];
  const str = parts.join("|");
  let hash = 0;
  for (let i = 0; i < str.length; i++) { hash = ((hash << 5) - hash) + str.charCodeAt(i); hash |= 0; }
  const fp = Math.abs(hash).toString(36) + "_" + screen.width + screen.height;
  try { localStorage.setItem("cvtailor_device_id", fp); } catch (e) {}
  return fp;
}

// -- XML escape ---------------------------------------------------------------
function escXml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
}

// -- Extract paragraphs from DOCX XML -----------------------------------------
function extractParagraphsFromXml(xml) {
  const paragraphs = [];
  const paraMatches = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  for (const para of paraMatches) {
    const textMatches = para.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
    const text = textMatches.map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("");
    paragraphs.push(text);
  }
  return paragraphs;
}

// -- Inject tailored text back into DOCX XML ----------------------------------
function injectTailoredIntoXml(xml, tailoredMap) {
  let paraIndex = 0;
  return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (paraXml) => {
    const textMatches = paraXml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
    const originalText = textMatches.map(t => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("").trim();
    const idx = paraIndex++;
    if (!originalText || !tailoredMap[idx]) return paraXml;
    const newText = tailoredMap[idx].replace(/<[^>]+>/g, "").trim();
    if (!newText || newText === originalText) return paraXml;
    const pPr = (paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [""])[0];
    const firstRun = (paraXml.match(/<w:r[ >][\s\S]*?<\/w:r>/) || [""])[0];
    const rPr = (firstRun.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escXml(newText)}</w:t></w:r></w:p>`;
  });
}

// -- Pure JS ZIP builder (always writes entries as STORED/uncompressed) ------
function buildZip(files) {
  const enc = new TextEncoder();
  const lhs = [], cd = []; let offset = 0;
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[i] = c; }
  const crc = d => { let c = 0xffffffff; for (let i = 0; i < d.length; i++) c = t[(c ^ d[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const u16 = n => [n & 0xff, (n >> 8) & 0xff];
  const u32 = n => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  for (const [name, data] of Object.entries(files)) {
    const nb = enc.encode(name), cr = crc(data);
    const lh = [0x50,0x4b,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(cr),...u32(data.length),...u32(data.length),...u16(nb.length),0,0,...nb];
    const ce = [0x50,0x4b,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(cr),...u32(data.length),...u32(data.length),...u16(nb.length),0,0,0,0,0,0,0,0,0,0,0,0,...u32(offset),...nb];
    lhs.push(new Uint8Array(lh), data); cd.push(new Uint8Array(ce)); offset += lh.length + data.length;
  }
  const cdSz = cd.reduce((s, b) => s + b.length, 0);
  const eocd = [0x50,0x4b,0x05,0x06,0,0,0,0,...u16(Object.keys(files).length),...u16(Object.keys(files).length),...u32(cdSz),...u32(offset),0,0];
  const parts = [...lhs, ...cd, new Uint8Array(eocd)];
  const tot = parts.reduce((s, p) => s + p.length, 0), res = new Uint8Array(tot); let pos = 0;
  for (const p of parts) { res.set(p, pos); pos += p.length; }
  return res;
}

// -- Inflate raw DEFLATE data using the browser's built-in DecompressionStream
// (real Word-saved .docx files store their parts DEFLATE-compressed, not
// stored raw -- this was the missing piece causing empty/garbled extraction)
async function inflateRaw(bytes) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Your browser doesn't support DOCX decompression. Please use a recent version of Chrome, Edge, or Firefox.");
  }
  const ds = new DecompressionStream("deflate-raw");
  const writer = ds.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const reader = ds.readable.getReader();
  const chunks = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  const out = new Uint8Array(total);
  let pos = 0;
  for (const c of chunks) { out.set(c, pos); pos += c.length; }
  return out;
}

// -- Read all ZIP entries, decompressing DEFLATE (method 8) as needed --------
async function readZipEntries(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const dec = new TextDecoder();

  let eocdOff = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (bytes[i] === 0x50 && bytes[i+1] === 0x4b && bytes[i+2] === 0x05 && bytes[i+3] === 0x06) { eocdOff = i; break; }
  }
  if (eocdOff === -1) throw new Error("Not a valid ZIP/DOCX file");

  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const cdOff = dv.getUint32(eocdOff + 16, true);
  const cdSize = dv.getUint32(eocdOff + 12, true);

  const rawEntries = [];
  let pos = cdOff;
  while (pos < cdOff + cdSize) {
    if (bytes[pos] !== 0x50 || bytes[pos+1] !== 0x4b || bytes[pos+2] !== 0x01 || bytes[pos+3] !== 0x02) break;
    const method = dv.getUint16(pos + 10, true);
    const compSize = dv.getUint32(pos + 20, true);
    const nameLen = dv.getUint16(pos + 28, true);
    const extraLen = dv.getUint16(pos + 30, true);
    const commentLen = dv.getUint16(pos + 32, true);
    const localOff = dv.getUint32(pos + 42, true);
    const name = dec.decode(bytes.slice(pos + 46, pos + 46 + nameLen));

    const lnLen = dv.getUint16(localOff + 26, true);
    const leLen = dv.getUint16(localOff + 28, true);
    const dataOff = localOff + 30 + lnLen + leLen;
    const rawData = bytes.slice(dataOff, dataOff + compSize);

    rawEntries.push({ name, method, rawData });
    pos += 46 + nameLen + extraLen + commentLen;
  }

  const entries = {};
  for (const e of rawEntries) {
    if (e.method === 0) {
      entries[e.name] = e.rawData; // already uncompressed
    } else if (e.method === 8) {
      entries[e.name] = await inflateRaw(e.rawData); // DEFLATE -> decompress
    } else {
      throw new Error(`Unsupported ZIP compression method (${e.method}) in "${e.name}"`);
    }
  }
  return entries;
}

// -- Download helper ------------------------------------------------------------
function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob), a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
}

// -- Rebuild DOCX from scratch (for non-DOCX or fallback) --------------------
function buildFallbackDocx(text) {
  const enc = new TextEncoder();
  const sections = text.split("\n").map(raw => {
    const l = raw.trimEnd();
    if (!l.trim()) return { type:"spacer" };
    if (/^# /.test(l)) return { type:"heading", level:1, text:l.replace(/^#+\s*/,"") };
    if (/^#{2,3} /.test(l)) return { type:"heading", level:2, text:l.replace(/^#+\s*/,"") };
    if (/^[-*] /.test(l)) return { type:"bullet", text:l.replace(/^[-*]\s*/,"") };
    return { type:"para", text:l };
  });
  const body = sections.map(s => {
    if (s.type==="heading") return `<w:p><w:pPr><w:pStyle w:val="${s.level===1?"Heading1":"Heading2"}"/></w:pPr><w:r><w:t>${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type==="bullet") return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${escXml(s.text)}</w:t></w:r></w:p>`;
    if (s.type==="spacer") return `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`;
    return `<w:p><w:r><w:t xml:space="preserve">${escXml(s.text||"")}</w:t></w:r></w:p>`;
  }).join("\n");
  const files = {
    "[Content_Types].xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`),
    "_rels/.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`),
    "word/styles.xml": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`),
    "word/_rels/document.xml.rels": enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`),
  };
  return buildZip(files);
}

export default function Trial() {
  const [trialState, setTrialState] = useState("checking");
  const [step, setStep] = useState(1);
  const [cvFile, setCvFile] = useState(null);
  const [cvData, setCvData] = useState(null); // { type: 'docx'|'txt'|'pdf', paragraphs?, text?, zipEntries?, docXml? }
  const [jd, setJd] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [fp, setFp] = useState("");
  const fileRef = useRef();

  useEffect(() => {
    const fingerprint = getFingerprint();
    setFp(fingerprint);
    let localUsed = false;
    try { localUsed = localStorage.getItem("cvtailor_trial_used") === "true"; } catch (e) {}
    if (localUsed) { setTrialState("used"); return; }
    fetch("/api/trial/check", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ fingerprint }) })
      .then(r => r.json())
      .then(d => {
        if (d.used) { try { localStorage.setItem("cvtailor_trial_used","true"); } catch (e) {} setTrialState("used"); }
        else setTrialState("available");
      })
      .catch(() => setTrialState("available"));
  }, []);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    const isDocx = /\.docx?$/i.test(file.name) || file.type.includes("wordprocessingml");
    const isTxt  = file.name.endsWith(".txt") || file.type === "text/plain";
    const isPdf  = file.name.endsWith(".pdf") || file.type === "application/pdf";
    if (!isDocx && !isTxt && !isPdf) { setError("Please upload PDF, DOCX, or TXT."); return; }
    setError(""); setCvFile(file); setCvData(null);

    if (isDocx) {
      const arrayBuffer = await file.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      try {
        const entries = await readZipEntries(bytes);
        const docBytes = entries["word/document.xml"];
        if (!docBytes) throw new Error("No document.xml found — this may not be a valid Word file");
        const dec = new TextDecoder();
        const docXml = dec.decode(docBytes);
        const paragraphs = extractParagraphsFromXml(docXml);
        if (paragraphs.length === 0) throw new Error("No text paragraphs found in this document");
        setCvData({ type:"docx", paragraphs, docXml, zipEntries: entries });
      } catch (e) {
        setError("Could not read DOCX: " + e.message + ". Try saving as TXT.");
      }
    } else if (isTxt) {
      const text = await file.text();
      setCvData({ type:"txt", text });
    } else {
      const r = new FileReader();
      r.onload = e => setCvData({ type:"pdf", b64: e.target.result.split(",")[1] });
      r.readAsDataURL(file);
    }
  }, []);

  const runTrial = async () => {
    if (!cvData || !jd.trim()) return;
    setLoading(true); setError("");

    try {
      try { localStorage.setItem("cvtailor_trial_used","true"); } catch (e) {}

      const filename = (cvFile?.name ? cvFile.name.replace(/\.[^.]+$/,"") + "_tailored" : "tailored_cv") + ".docx";

      if (cvData.type === "docx") {
        const res = await fetch("/api/tailor-docx", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ paragraphs: cvData.paragraphs, jd }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "Tailoring failed");

        const tailoredMap = data.tailoredMap || {};

        const newXml = injectTailoredIntoXml(cvData.docXml, tailoredMap);
        const enc = new TextEncoder();

        // zipEntries values are now plain decompressed Uint8Array (not {data,...})
        const newFiles = {};
        for (const [name, data] of Object.entries(cvData.zipEntries)) {
          newFiles[name] = name === "word/document.xml" ? enc.encode(newXml) : data;
        }
        const newDocxBytes = buildZip(newFiles);
        downloadBytes(newDocxBytes, filename);

      } else {
        const body = cvData.type === "txt"
          ? { fingerprint: fp, cvText: cvData.text, jd, fileName: cvFile?.name }
          : { fingerprint: fp, cvText: "__PDF__" + cvData.b64, jd, fileName: cvFile?.name };

        const res = await fetch("/api/trial/use", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        const docxBytes = buildFallbackDocx(data.tailored);
        setTimeout(() => downloadBytes(docxBytes, filename), 300);
      }

      // Mark trial used in DB (fire and forget)
      fetch("/api/trial/use", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fingerprint: fp, cvText: "(processed)", jd: "(processed)", fileName: cvFile?.name, markOnly: true }),
      }).catch(() => {});

      setTrialState("done");
      setStep(4);

    } catch (e) {
      setError(e.message);
      if (e.message.includes("already used") || e.message.includes("device")) {
        try { localStorage.setItem("cvtailor_trial_used","true"); } catch (err) {}
        setTrialState("used");
      }
    } finally { setLoading(false); }
  };

  const S = {
    page: { minHeight:"100vh", background:"linear-gradient(135deg,#0f172a,#1e1b4b)", fontFamily:"system-ui,sans-serif", color:"#e2e8f0", padding:"0 0 40px" },
    nav: { display:"flex", alignItems:"center", justifyContent:"space-between", padding:"16px 28px", borderBottom:"1px solid #1e293b" },
    logo: { fontWeight:800, fontSize:18, background:"linear-gradient(135deg,#c7d2fe,#e9d5ff)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" },
    main: { maxWidth:720, margin:"0 auto", padding:"28px 20px" },
    card: { background:"#1e293b", border:"1px solid #334155", borderRadius:16, padding:"24px 28px", marginBottom:18 },
    h2: { fontSize:17, fontWeight:700, color:"#c7d2fe", marginBottom:6 },
    desc: { fontSize:13, color:"#64748b", marginBottom:18 },
    dz: { border:"2px dashed #334155", borderRadius:12, padding:"36px 20px", textAlign:"center", cursor:"pointer", background:"#0f172a", transition:"all .2s" },
    ta: { width:"100%", minHeight:200, background:"#0f172a", border:"1px solid #334155", borderRadius:10, color:"#e2e8f0", fontSize:13, padding:"13px 15px", resize:"vertical", outline:"none", fontFamily:"inherit", lineHeight:1.6 },
    btnP: { padding:"12px 28px", borderRadius:10, border:"none", background:"linear-gradient(135deg,#6366f1,#8b5cf6)", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" },
    btnS: { padding:"12px 24px", borderRadius:10, border:"none", background:"#334155", color:"#fff", fontWeight:600, fontSize:14, cursor:"pointer" },
    btnSu: { padding:"12px 28px", borderRadius:10, border:"none", background:"linear-gradient(135deg,#10b981,#059669)", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" },
    err: { background:"#450a0a", border:"1px solid #b91c1c", borderRadius:10, padding:"12px 16px", color:"#fca5a5", fontSize:13, marginBottom:16 },
  };

  if (trialState === "checking") return <div style={{ ...S.page, display:"flex", alignItems:"center", justifyContent:"center" }}><p style={{ color:"#64748b" }}>Checking…</p></div>;

  if (trialState === "used") return (
    <div style={{ ...S.page, display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ maxWidth:440, textAlign:"center", padding:20 }}>
        <div style={{ fontSize:40, marginBottom:16 }}>🔒</div>
        <h2 style={{ fontSize:22, fontWeight:700, color:"#c7d2fe", marginBottom:10 }}>Free trial already used</h2>
        <p style={{ color:"#94a3b8", marginBottom:28, lineHeight:1.6 }}>Your free trial has been used on this device. Sign up for a plan to keep tailoring CVs.</p>
        <div style={{ display:"flex", gap:12, justifyContent:"center", flexWrap:"wrap" }}>
          <a href="/pricing"><button style={S.btnP}>See pricing →</button></a>
          <a href="/sign-up"><button style={S.btnS}>Create account</button></a>
        </div>
      </div>
    </div>
  );

  const STEPS = ["Upload CV","Paste JD","Tailor","Done"];
  const ready = !!cvData;

  return (
    <div style={S.page}>
      <nav style={S.nav}>
        <a href="/"><span style={S.logo}>CV Tailor Pro</span></a>
        <div style={{ display:"flex", gap:10 }}>
          <a href="/sign-in"><button style={{ ...S.btnS, padding:"8px 16px", fontSize:13 }}>Sign in</button></a>
          <a href="/pricing"><button style={{ ...S.btnP, padding:"8px 16px", fontSize:13 }}>See plans</button></a>
        </div>
      </nav>

      <div style={S.main}>
        <div style={{ background:"linear-gradient(135deg,#1e1b4b,#1e293b)", border:"1px solid #4338ca", borderRadius:12, padding:"14px 18px", marginBottom:20, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
          <div>
            <span style={{ fontSize:14, fontWeight:700, color:"#a5b4fc" }}>✦ Free Trial</span>
            <span style={{ fontSize:13, color:"#64748b", marginLeft:10 }}>One free tailor — no signup required</span>
          </div>
          <a href="/pricing"><button style={{ ...S.btnP, padding:"8px 16px", fontSize:12 }}>Get full access →</button></a>
        </div>

        <div style={{ display:"flex", gap:10, marginBottom:20, flexWrap:"wrap" }}>
          {STEPS.map((s,i) => {
            const n=i+1, done=step>n, active=step===n;
            return <div key={s} style={{ display:"flex", alignItems:"center", gap:7 }}>
              <div style={{ width:26, height:26, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700, fontSize:12, color:"#fff", background:done?"#10b981":active?"#6366f1":"#334155" }}>{done?"✓":n}</div>
              <span style={{ fontSize:12, fontWeight:active||done?600:400, color:active?"#c7d2fe":done?"#6ee7b7":"#64748b" }}>{s}</span>
            </div>;
          })}
        </div>
        <div style={{ height:3, background:"#1e293b", borderRadius:99, marginBottom:22 }}>
          <div style={{ height:"100%", borderRadius:99, background:"linear-gradient(90deg,#6366f1,#a855f7)", width:`${((step-1)/3)*100}%`, transition:"width .5s" }} />
        </div>

        {error && <div style={S.err}>⚠ {error}</div>}

        {step===1 && <div style={S.card}>
          <h2 style={S.h2}>Step 1 — Upload your CV</h2>
          <p style={S.desc}>PDF, DOCX, or TXT format</p>
          <div style={{ ...S.dz, borderColor:ready?"#10b981":"#334155", background:ready?"#0d2318":"#0f172a" }}
            onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files[0]);}}>
            <div style={{ fontSize:30, marginBottom:10 }}>{ready?"✅":"📄"}</div>
            {cvFile ? <><p style={{ color:ready?"#6ee7b7":"#fbbf24", fontWeight:600 }}>{cvFile.name}</p><p style={{ fontSize:12, color:"#475569" }}>{ready?"Ready ✓ · Click to replace":"Processing..."}</p></>
              : <><p style={{ color:"#94a3b8", fontWeight:500 }}>Drag & drop or click to browse</p><p style={{ fontSize:13, color:"#475569" }}>PDF · DOCX · TXT</p></>}
          </div>
          <input ref={fileRef} type="file" accept=".pdf,.docx,.doc,.txt" style={{ display:"none" }} onChange={e=>handleFile(e.target.files[0])} />
          <div style={{ display:"flex", justifyContent:"flex-end", marginTop:16 }}>
            <button style={{ ...S.btnP, opacity:!ready?.5:1 }} disabled={!ready} onClick={()=>setStep(2)}>Continue →</button>
          </div>
        </div>}

        {step===2 && <div style={S.card}>
          <h2 style={S.h2}>Step 2 — Paste job description</h2>
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
          <h2 style={S.h2}>Step 3 — Tailor your CV</h2>
          <p style={S.desc}>AI matches your CV to the job description</p>
          <div style={{ background:"#0f172a", borderRadius:10, padding:"14px 18px", border:"1px solid #1e293b", marginBottom:18 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
              <span style={{ fontSize:12, color:"#64748b" }}>CV</span><span style={{ fontSize:12, color:"#6ee7b7" }}>✓ Ready</span>
            </div>
            <p style={{ fontWeight:600, color:"#e2e8f0", fontSize:14 }}>{cvFile?.name}</p>
            <div style={{ marginTop:10, paddingTop:10, borderTop:"1px solid #1e293b" }}>
              <span style={{ fontSize:12, color:"#64748b" }}>JD PREVIEW</span>
              <p style={{ marginTop:4, color:"#94a3b8", fontSize:13 }}>{jd.slice(0,100)}…</p>
            </div>
          </div>
          {loading && <div style={{ background:"#1e1b4b", borderRadius:10, padding:"14px", border:"1px solid #4338ca", marginBottom:16, textAlign:"center" }}>
            <div style={{ width:26, height:26, border:"3px solid #4338ca", borderTopColor:"#818cf8", borderRadius:"50%", animation:"spin .8s linear infinite", margin:"0 auto 8px" }} />
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
            <p style={{ color:"#a5b4fc", fontSize:13, fontWeight:600 }}>AI is tailoring your CV…</p>
            <p style={{ color:"#64748b", fontSize:12, marginTop:4 }}>Usually 5–10 seconds. Please don't close this tab.</p>
          </div>}
          <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
            <button style={S.btnS} disabled={loading} onClick={()=>setStep(2)}>← Back</button>
            <button style={{ ...S.btnP, opacity:loading?.5:1 }} disabled={loading} onClick={runTrial}>
              {loading ? "Processing…" : "✦ Tailor My CV (free)"}
            </button>
          </div>
        </div>}

        {step===4 && <div style={S.card}>
          <div style={{ textAlign:"center", padding:"20px 0" }}>
            <div style={{ fontSize:48, marginBottom:16 }}>✅</div>
            <h2 style={{ ...S.h2, color:"#6ee7b7", fontSize:20, textAlign:"center" }}>Your tailored CV has downloaded!</h2>
            <p style={{ ...S.desc, textAlign:"center", marginBottom:24 }}>Check your Downloads folder for the Word document.</p>
          </div>
          <div style={{ background:"linear-gradient(135deg,#1e1b4b,#1e293b)", border:"1px solid #6366f1", borderRadius:12, padding:"20px 22px", textAlign:"center" }}>
            <p style={{ fontWeight:700, color:"#c7d2fe", fontSize:15, marginBottom:6 }}>Want to tailor more CVs?</p>
            <p style={{ color:"#94a3b8", fontSize:13, marginBottom:16 }}>Plans start at just £3/month for 10 tailored CVs. Cancel any time.</p>
            <div style={{ display:"flex", gap:10, justifyContent:"center", flexWrap:"wrap" }}>
              <a href="/pricing"><button style={S.btnP}>See plans →</button></a>
              <a href="/sign-up"><button style={S.btnS}>Create account</button></a>
            </div>
          </div>
        </div>}
      </div>
    </div>
  );
}
