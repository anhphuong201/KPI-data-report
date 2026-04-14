import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Taxonomy ─────────────────────────────────────────────────────────────────
const TAXONOMY = [
  { id: "IRAP", label: "IRAP Projects", sub: [
    { id: "IRAP_CTO", label: "IRAP CTO Projects" },
    { id: "IRAP_OTHER", label: "Other IRAP Projects" },
  ]},
  { id: "TRICOUNCIL", label: "Tri-Council Industry Partnership", sub: [
    { id: "TC_NSERC", label: "NSERC" },
    { id: "TC_CIHR", label: "CIHR" },
    { id: "TC_SSHRC", label: "SSHRC" },
    { id: "TC_MITACS", label: "Industry Partnership Scholarships (MITACS, ICAN)" },
    { id: "TC_OTHER", label: "Other Tri-Council Industry Partnership" },
  ]},
  { id: "ACOA", label: "ACOA Projects", sub: [
    { id: "ACOA_AIF", label: "AIF" },
    { id: "ACOA_REGI", label: "REGI (RRRF)" },
    { id: "ACOA_BDP", label: "BDP" },
    { id: "ACOA_ICF", label: "ICF, PBS" },
  ]},
  { id: "PROVINCIAL", label: "Provincial Funding Programs", sub: [
    { id: "PROV_VOUCHER", label: "Vouchers" },
    { id: "PROV_OTHER", label: "Other Provincial Funding" },
  ]},
  { id: "GOV_LEVERAGED", label: "Other Industry with Government Leveraged Funding", sub: [] },
  { id: "RD_CONTRACTS", label: "R&D Contracts", sub: [
    { id: "RD_INDUSTRY", label: "Industry R&D Contracts" },
    { id: "RD_OTHER", label: "Other R&D Contracts" },
    { id: "RD_GOV", label: "Government R&D Contracts" },
  ]},
];

function allLeaves() {
  const out = [];
  for (const g of TAXONOMY) {
    if (g.sub.length === 0) out.push({ id: g.id, label: g.label, parentId: g.id, parentLabel: g.label });
    else g.sub.forEach(s => out.push({ id: s.id, label: s.label, parentId: g.id, parentLabel: g.label }));
  }
  out.push({ id: "UNCATEGORIZED", label: "Uncategorized", parentId: null, parentLabel: "" });
  return out;
}
const LEAVES = allLeaves();

function leafLabel(id) { return LEAVES.find(l => l.id === id)?.label || id; }
function leafParentLabel(id) { return LEAVES.find(l => l.id === id)?.parentLabel || ""; }

// ─── Default keywords ─────────────────────────────────────────────────────────
const DEFAULT_KEYWORDS = {
  IRAP_CTO: ["IRAP CTO", "Industrial Research Assistance Program CTO", "NRC IRAP CTO"],
  IRAP_OTHER: ["IRAP", "Industrial Research Assistance Program", "NRC IRAP", "National Research Council Canada"],
  TC_NSERC: ["Natural Sciences and Engineering Research Council", "NSERC", "Conseil de recherches en sciences naturelles", "CRSNG", "Subvention à la découverte", "Discovery Grant", "Alliance Grant"],
  TC_CIHR: ["Canadian Institutes of Health Research", "CIHR", "Instituts de recherche en santé du Canada", "IRSC", "Chaire de recherche en santé"],
  TC_SSHRC: ["Social Sciences and Humanities Research Council", "SSHRC", "Conseil de recherches en sciences humaines", "CRSH", "Insight Grant"],
  TC_MITACS: ["Mitacs", "MITACS", "ICAN", "Accélération", "Accelerate", "Elevate", "Globalink"],
  TC_OTHER: ["Tri-Council", "CFI", "Canada Foundation for Innovation", "Fondation canadienne pour l'innovation", "AARMS", "Atlantic Association for Research in the Mathematical Sciences"],
  ACOA_AIF: ["AIF", "Atlantic Innovation Fund", "Fonds d'innovation de l'Atlantique"],
  ACOA_REGI: ["REGI", "RRRF", "Regional Relief", "Regional Economic Growth"],
  ACOA_BDP: ["BDP", "Business Development Program"],
  ACOA_ICF: ["ICF", "PBS", "Innovative Communities Fund"],
  PROV_VOUCHER: ["Voucher", "Bon de service", "Innovation Voucher"],
  PROV_OTHER: ["Province", "Provincial", "Gouvernement du Nouveau-Brunswick", "Government of New Brunswick", "Government of Nova Scotia", "Government of Newfoundland and Labrador", "Government of Prince Edward Island"],
  GOV_LEVERAGED: ["SIF", "Strategic Innovation Fund", "FedDev"],
  RD_INDUSTRY: ["Industry R&D", "Research Contract", "Contrat de recherche", , "Industry Partner", "Contract"],
  RD_OTHER: ["Research Agreement", "Research Services", "Testing", "Clinical", "NGO", "Not-for-profit", "Association", "Community"],
  RD_GOV: ["Government of Canada", "Gouvernement du Canada", "Canadian Heritage", "Department of Canadian Heritage", "DND", "Agriculture Canada", "Environment Canada", "Ministère de l'Éducation", "Gouvernement du Nouveau-Brunswick"],
};

function detectCategory(agency, program, kw) {
  if (!agency) return "UNCATEGORIZED";
  const hay = (agency + " " + (program || "")).toLowerCase();
  for (const [cat, words] of Object.entries(kw)) {
    if (words.some(w => hay.includes(w.toLowerCase()))) return cat;
  }
  return "UNCATEGORIZED";
}

// ─── File parsing ─────────────────────────────────────────────────────────────
function parseSheet(data) {
  let hr = -1;
  for (let i = 0; i < data.length; i++) {
    const vals = (data[i] || []).map(v => String(v || "").toLowerCase());
    if (vals.some(v => v.includes("project title") || v.includes("titre") || v.includes("agency") || v.includes("awarded"))) { hr = i; break; }
  }
  if (hr === -1) return [];
  const headers = data[hr].map(v => String(v || "").trim());
  const rows = [];
  for (let i = hr + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(v => v === null || v === undefined || v === "")) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx] ?? ""; });
    rows.push(obj);
  }
  return rows;
}

function normalize(raw, sheet, kw) {
  const keys = Object.keys(raw);
  const find = (...terms) => {
    for (const t of terms) {
      const k = keys.find(k => k.toLowerCase().includes(t.toLowerCase()));
      if (k !== undefined) return String(raw[k] ?? "");
    }
    return "";
  };
  const agency = find("agency", "agence", "funding agency", "sponsor source");
  const program = find("program", "programme", "funding program");
  const title = find("project title", "titre du projet", "titre");
  const pi = find("principal investigator", "investigator info", "chercheur principal");
  const amount = parseFloat(String(find("awarded amount", "amount", "montant")).replace(/[^0-9.]/g, "")) || 0;
  const startDate = find("start date", "date de début", "sponsor start");
  const dept = find("department", "département", "faculty", "faculté", "principal investigator department");
  const cat = detectCategory(agency, program, kw);
  return {
    id: Math.random().toString(36).slice(2),
    institution: sheet, title, pi, agency, program, department: dept,
    amount, startDate, category: cat,
    isAtlantic: false,
    isIndustry: ["RD_INDUSTRY","IRAP_CTO","IRAP_OTHER","TC_MITACS","GOV_LEVERAGED"].includes(cat),
    include: true,
  };
}

// ─── Summary helpers ──────────────────────────────────────────────────────────
function buildSummary(included) {
  const s = {};
  for (const p of included) {
    if (!s[p.category]) s[p.category] = { count: 0, total: 0, atl: 0, atlTotal: 0 };
    s[p.category].count++;
    s[p.category].total += p.amount;
    if (p.isAtlantic) { s[p.category].atl++; s[p.category].atlTotal += p.amount; }
  }
  return s;
}

function groupSummary(sm) {
  return TAXONOMY.map(g => {
    const leaves = g.sub.length === 0 ? [{ id: g.id, label: g.label }] : g.sub;
    const subs = leaves.map(s => ({ ...s, ...(sm[s.id] || { count: 0, total: 0, atl: 0, atlTotal: 0 }) }));
    return {
      ...g, subs,
      totCount: subs.reduce((a, s) => a + s.count, 0),
      totAmount: subs.reduce((a, s) => a + s.total, 0),
      atlCount: subs.reduce((a, s) => a + s.atl, 0),
      atlTotal: subs.reduce((a, s) => a + s.atlTotal, 0),
    };
  });
}

// ─── Excel export ─────────────────────────────────────────────────────────────
function doExport(included) {
  const wb = XLSX.utils.book_new();
  const grouped = groupSummary(buildSummary(included));
  const kpiRows = [
    ["Phase 7 - KPI Metrics", "", "", "", "", ""],
    ["(To Be Reported Quarterly)", "Description", "Metric #", "Contracts with Atlantic Canadian Industry #", "Total $ Awarded", "Total $ Awarded with Atlantic Canadian Industry"],
  ];
  let gc = 0, gt = 0, ga = 0, gat = 0;
  for (const g of grouped) {
    kpiRows.push([g.label, "", g.totCount, g.atlCount, g.totAmount, g.atlTotal]);
    for (const s of g.subs) kpiRows.push(["  " + s.label, "", s.count, s.atl, s.total, s.atlTotal]);
    gc += g.totCount; gt += g.totAmount; ga += g.atlCount; gat += g.atlTotal;
  }
  kpiRows.push(["TOTAL", "", gc, ga, gt, gat]);
  const ws1 = XLSX.utils.aoa_to_sheet(kpiRows);
  ws1["!cols"] = [{ wch: 52 }, { wch: 20 }, { wch: 10 }, { wch: 36 }, { wch: 16 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws1, "KPI");

  const dh = ["Institution","Project Title","PI","Agency","Program","Department","Sub-Category","Group","Start Date","Awarded Amount","Atlantic Canada","Industry"];
  const dr = [dh, ...included.map(p => [
    p.institution, p.title, p.pi, p.agency, p.program, p.department,
    leafLabel(p.category), leafParentLabel(p.category),
    p.startDate, p.amount,
    p.isAtlantic ? "Yes" : "No", p.isIndustry ? "Yes" : "No",
  ])];
  const ws2 = XLSX.utils.aoa_to_sheet(dr);
  ws2["!cols"] = [14,40,28,36,28,26,36,32,14,16,14,10].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws2, "Projects Detail");

  const blob = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([blob], { type: "application/octet-stream" }));
  a.download = "Quarterly_Report.xlsx"; a.click();
}

// ─── Colors ───────────────────────────────────────────────────────────────────
const CC = {
  IRAP_CTO:"#185FA5", IRAP_OTHER:"#378ADD",
  TC_NSERC:"#0F6E56", TC_CIHR:"#1D9E75", TC_SSHRC:"#5DCAA5", TC_MITACS:"#993C1D", TC_OTHER:"#D85A30",
  ACOA_AIF:"#534AB7", ACOA_REGI:"#7F77DD", ACOA_BDP:"#AFA9EC", ACOA_ICF:"#5DCAA5",
  PROV_VOUCHER:"#993556", PROV_OTHER:"#D4537E",
  GOV_LEVERAGED:"#854F0B",
  RD_INDUSTRY:"#A32D2D", RD_OTHER:"#E24B4A", RD_GOV:"#F09595",
  UNCATEGORIZED:"#888780",
};
function Badge({ id }) {
  const c = CC[id] || "#888";
  return <span style={{ display:"inline-block", fontSize:11, padding:"2px 8px", borderRadius:20, background:c+"22", color:c, fontWeight:500, whiteSpace:"nowrap" }}>{leafLabel(id)}</span>;
}

// ─── App ──────────────────────────────────────────────────────────────────────
const STEPS = ["Upload","Review & Tag","Summary","Download"];

export default function App() {
  const [step, setStep] = useState(0);
  const [projects, setProjects] = useState([]);
  const [keywords, setKeywords] = useState(() => {
    try { const s = localStorage.getItem("qr_kw_v2"); return s ? JSON.parse(s) : DEFAULT_KEYWORDS; } catch { return DEFAULT_KEYWORDS; }
  });
  const [editingKw, setEditingKw] = useState(false);
  const [kwDraft, setKwDraft] = useState(null);
  const [kwInput, setKwInput] = useState({});
  const [filterCat, setFilterCat] = useState("All");
  const [sortCol, setSortCol] = useState("title");
  const [sortDir, setSortDir] = useState("asc");
  const [addRow, setAddRow] = useState(false);
  const [newProj, setNewProj] = useState({});
  const [toast, setToast] = useState("");
  const fileRef = useRef();

  const showToast = msg => { setToast(msg); setTimeout(() => setToast(""), 3000); };

  const handleFile = async e => {
    const file = e.target.files[0]; if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const all = [];
    for (const sn of wb.SheetNames) {
      const rows = parseSheet(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null }));
      rows.forEach(r => all.push(normalize(r, sn, keywords)));
    }
    setProjects(all);
    showToast(`Loaded ${all.length} projects from ${wb.SheetNames.length} sheet(s)`);
    setStep(1); e.target.value = "";
  };

  const recategorize = useCallback(kw => {
    setProjects(prev => prev.map(p => {
      const cat = detectCategory(p.agency, p.program, kw);
      return { ...p, category: cat, isIndustry: ["RD_INDUSTRY","IRAP_CTO","IRAP_OTHER","TC_MITACS","GOV_LEVERAGED"].includes(cat) };
    }));
  }, []);

  const upd = (id, field, val) => setProjects(prev => prev.map(p => p.id === id ? { ...p, [field]: val } : p));
  const included = projects.filter(p => p.include);
  const leafIds = LEAVES.map(l => l.id);

  const filtered = projects
    .filter(p => filterCat === "All" || p.category === filterCat)
    .sort((a, b) => {
      const av = a[sortCol] ?? "", bv = b[sortCol] ?? "";
      const cmp = typeof av === "number" ? av - bv : String(av).localeCompare(String(bv));
      return sortDir === "asc" ? cmp : -cmp;
    });

  const toggleSort = col => { if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc"); else { setSortCol(col); setSortDir("asc"); } };

  const addProject = () => {
    const cat = detectCategory(newProj.agency || "", newProj.program || "", keywords);
    setProjects(prev => [...prev, {
      id: Math.random().toString(36).slice(2),
      institution: newProj.institution || "", title: newProj.title || "", pi: newProj.pi || "",
      agency: newProj.agency || "", program: newProj.program || "", department: newProj.department || "",
      amount: parseFloat(newProj.amount) || 0, startDate: newProj.startDate || "",
      category: newProj.category || cat,
      isAtlantic: !!newProj.isAtlantic, isIndustry: !!newProj.isIndustry, include: true,
    }]);
    setAddRow(false); setNewProj({}); showToast("Project added");
  };

  const saveKw = () => {
    setKeywords(kwDraft);
    try { localStorage.setItem("qr_kw_v2", JSON.stringify(kwDraft)); } catch {}
    recategorize(kwDraft);
    setEditingKw(false);
    showToast("Keywords saved — projects re-categorized");
  };

  const summary = buildSummary(included);
  const grouped = groupSummary(summary);
  const grandTotal = included.reduce((s, p) => s + p.amount, 0);
  const grandAtl = included.filter(p => p.isAtlantic).reduce((s, p) => s + p.amount, 0);

  const S = {
    card: { background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:"var(--border-radius-lg)", padding:"1.25rem" },
    btn: { border:"0.5px solid var(--color-border-secondary)", background:"transparent", color:"var(--color-text-primary)", padding:"6px 14px", borderRadius:"var(--border-radius-md)", fontSize:13, cursor:"pointer" },
    btnP: { border:"none", background:"#185FA5", color:"#fff", padding:"7px 18px", borderRadius:"var(--border-radius-md)", fontSize:13, cursor:"pointer", fontWeight:500 },
    th: { padding:"8px 10px", fontSize:12, fontWeight:500, color:"var(--color-text-secondary)", textAlign:"left", borderBottom:"0.5px solid var(--color-border-tertiary)", whiteSpace:"nowrap", cursor:"pointer", userSelect:"none" },
    td: { padding:"7px 10px", fontSize:13, borderBottom:"0.5px solid var(--color-border-tertiary)", verticalAlign:"middle" },
    sel: { border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)", color:"var(--color-text-primary)", padding:"5px 8px", borderRadius:"var(--border-radius-md)", fontSize:13 },
    inp: { border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)", color:"var(--color-text-primary)", padding:"5px 8px", borderRadius:"var(--border-radius-md)", fontSize:13 },
    metric: { background:"var(--color-background-secondary)", borderRadius:"var(--border-radius-md)", padding:"0.75rem 1rem" },
  };

  return (
    <div style={{ fontFamily:"var(--font-sans)", color:"var(--color-text-primary)", maxWidth:1280, margin:"0 auto", padding:"1.5rem 1rem" }}>
      <div style={{ borderBottom:"0.5px solid var(--color-border-tertiary)", paddingBottom:"1rem", marginBottom:"1.5rem" }}>
        <h1 style={{ fontSize:22, fontWeight:500, margin:0 }}>Quarterly Report Builder</h1>
        <p style={{ fontSize:14, color:"var(--color-text-secondary)", margin:"4px 0 0" }}>Springboard Atlantic — Research Funding Consolidation</p>
      </div>

      {toast && <div style={{ position:"fixed", bottom:24, right:24, background:"#185FA5", color:"#fff", padding:"10px 20px", borderRadius:"var(--border-radius-md)", fontSize:14, zIndex:9999 }}>{toast}</div>}

      <div style={{ display:"flex", gap:0, marginBottom:"1.5rem", background:"var(--color-background-secondary)", borderRadius:"var(--border-radius-lg)", padding:4, width:"fit-content" }}>
        {STEPS.map((label, i) => (
          <button key={i} onClick={() => (projects.length > 0 || i === 0) && setStep(i)}
            style={{ border:"none", background:step===i?"var(--color-background-primary)":"transparent", color:step===i?"var(--color-text-primary)":"var(--color-text-secondary)", padding:"6px 18px", borderRadius:"var(--border-radius-md)", fontWeight:step===i?500:400, fontSize:14, cursor:"pointer" }}>
            {i+1}. {label}
          </button>
        ))}
      </div>

      {/* STEP 0 */}
      {step === 0 && (
        <div style={{ ...S.card, maxWidth:540 }}>
          <h2 style={{ fontSize:16, fontWeight:500, margin:"0 0 0.75rem" }}>Upload funding export</h2>
          <p style={{ fontSize:14, color:"var(--color-text-secondary)", lineHeight:1.65, margin:"0 0 1.25rem" }}>
            Upload the .xlsx file exported from ROMEO or your institution's research management system. Each worksheet is treated as one institution. Columns are auto-detected.
          </p>
          <div style={{ border:"1.5px dashed var(--color-border-secondary)", borderRadius:"var(--border-radius-lg)", padding:"2.5rem 2rem", textAlign:"center", marginBottom:"1rem" }}>
            <div style={{ fontSize:13, color:"var(--color-text-secondary)", marginBottom:"0.75rem" }}>Supports multi-sheet .xlsx (ROMEO, Banner, etc.)</div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
            <button style={S.btnP} onClick={() => fileRef.current.click()}>Choose file</button>
          </div>
          <p style={{ fontSize:12, color:"var(--color-text-secondary)", margin:0 }}>Auto-detects: Project Title, Agency, PI, Awarded Amount, Program, Department, Start Date</p>
          <div style={{ marginTop:"1.25rem", paddingTop:"1.25rem", borderTop:"0.5px solid var(--color-border-tertiary)" }}>
            <div style={{ fontSize:13, fontWeight:500, marginBottom:"0.4rem" }}>Category keywords</div>
            <p style={{ fontSize:13, color:"var(--color-text-secondary)", margin:"0 0 0.75rem", lineHeight:1.6 }}>
              Set up your keyword rules before uploading. Keywords control how projects are auto-categorized based on funding agency and program name.
            </p>
            <button style={S.btn} onClick={() => { setKwDraft(JSON.parse(JSON.stringify(keywords))); setKwInput({}); setEditingKw(true); }}>⚙ Edit category keywords</button>
          </div>
        </div>
      )}

      {/* STEP 1 */}
      {step === 1 && (
        <div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))", gap:10, marginBottom:"1rem" }}>
            {[["Total projects",projects.length],["Included",included.length],["Total awarded","$"+grandTotal.toLocaleString()],["Atlantic Canada",included.filter(p=>p.isAtlantic).length]].map(([l,v]) => (
              <div key={l} style={S.metric}><div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:4 }}>{l}</div><div style={{ fontSize:20, fontWeight:500 }}>{v}</div></div>
            ))}
          </div>

          <div style={{ display:"flex", gap:8, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
            <select style={S.sel} value={filterCat} onChange={e => setFilterCat(e.target.value)}>
              <option value="All">All categories</option>
              {TAXONOMY.map(g => (
                <optgroup key={g.id} label={g.label}>
                  {(g.sub.length === 0 ? [{ id:g.id, label:g.label }] : g.sub).map(s => (
                    <option key={s.id} value={s.id}>{s.label}</option>
                  ))}
                </optgroup>
              ))}
              <option value="UNCATEGORIZED">Uncategorized</option>
            </select>
            <div style={{ flex:1 }} />
            <button style={S.btn} onClick={() => { setAddRow(true); setNewProj({}); }}>+ Add project</button>
            <button style={S.btn} onClick={() => { setKwDraft(JSON.parse(JSON.stringify(keywords))); setKwInput({}); setEditingKw(true); }}>⚙ Category keywords</button>
          </div>

          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:960 }}>
              <thead>
                <tr style={{ background:"var(--color-background-secondary)" }}>
                  {[["include","✓",36],["title","Project Title",230],["pi","PI",120],["agency","Agency",160],["amount","Amount ($)",100],["category","Category",190],["isAtlantic","Atlantic",72],["isIndustry","Industry",72]].map(([col,lbl,w]) => (
                    <th key={col} style={{ ...S.th, width:w }} onClick={() => toggleSort(col)}>
                      {lbl}{sortCol===col?(sortDir==="asc"?" ↑":" ↓"):""}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {addRow && (
                  <tr style={{ background:"var(--color-background-secondary)" }}>
                    <td style={S.td}></td>
                    <td style={S.td}><input style={{ ...S.inp, width:"100%" }} placeholder="Project title" value={newProj.title||""} onChange={e => setNewProj(p => ({ ...p, title:e.target.value }))} /></td>
                    <td style={S.td}><input style={{ ...S.inp, width:"100%" }} placeholder="PI" value={newProj.pi||""} onChange={e => setNewProj(p => ({ ...p, pi:e.target.value }))} /></td>
                    <td style={S.td}><input style={{ ...S.inp, width:"100%" }} placeholder="Agency" value={newProj.agency||""} onChange={e => setNewProj(p => ({ ...p, agency:e.target.value }))} /></td>
                    <td style={S.td}><input style={{ ...S.inp, width:90 }} type="number" placeholder="Amount" value={newProj.amount||""} onChange={e => setNewProj(p => ({ ...p, amount:e.target.value }))} /></td>
                    <td style={S.td}>
                      <select style={{ ...S.sel, fontSize:11 }} value={newProj.category||"UNCATEGORIZED"} onChange={e => setNewProj(p => ({ ...p, category:e.target.value }))}>
                        {TAXONOMY.map(g => (
                          <optgroup key={g.id} label={g.label}>
                            {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                          </optgroup>
                        ))}
                        <option value="UNCATEGORIZED">Uncategorized</option>
                      </select>
                    </td>
                    <td style={{ ...S.td, textAlign:"center" }}><input type="checkbox" checked={!!newProj.isAtlantic} onChange={e => setNewProj(p => ({ ...p, isAtlantic:e.target.checked }))} /></td>
                    <td style={{ ...S.td, textAlign:"center" }}>
                      <div style={{ display:"flex", gap:6, justifyContent:"center" }}>
                        <button style={{ ...S.btnP, fontSize:11, padding:"3px 10px" }} onClick={addProject}>Save</button>
                        <button style={{ ...S.btn, fontSize:11, padding:"3px 8px" }} onClick={() => setAddRow(false)}>×</button>
                      </div>
                    </td>
                  </tr>
                )}
                {filtered.map(p => (
                  <tr key={p.id} style={{ opacity:p.include?1:0.4 }}>
                    <td style={S.td}><input type="checkbox" checked={p.include} onChange={e => upd(p.id,"include",e.target.checked)} /></td>
                    <td style={{ ...S.td, maxWidth:230 }}>
                      <div style={{ fontSize:13, lineHeight:1.4 }}>{p.title}</div>
                      {p.department && <div style={{ fontSize:11, color:"var(--color-text-secondary)", marginTop:2 }}>{p.department}</div>}
                    </td>
                    <td style={{ ...S.td, fontSize:12 }}>{p.pi}</td>
                    <td style={{ ...S.td, fontSize:12, color:"var(--color-text-secondary)" }}>{p.agency}</td>
                    <td style={{ ...S.td, textAlign:"right", fontVariantNumeric:"tabular-nums" }}>${p.amount.toLocaleString()}</td>
                    <td style={S.td}>
                      <select style={{ ...S.sel, fontSize:11, padding:"2px 6px" }} value={p.category} onChange={e => upd(p.id,"category",e.target.value)}>
                        {TAXONOMY.map(g => (
                          <optgroup key={g.id} label={g.label}>
                            {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                          </optgroup>
                        ))}
                        <option value="UNCATEGORIZED">Uncategorized</option>
                      </select>
                    </td>
                    <td style={{ ...S.td, textAlign:"center" }}><input type="checkbox" checked={p.isAtlantic} onChange={e => upd(p.id,"isAtlantic",e.target.checked)} /></td>
                    <td style={{ ...S.td, textAlign:"center" }}><input type="checkbox" checked={p.isIndustry} onChange={e => upd(p.id,"isIndustry",e.target.checked)} /></td>
                  </tr>
                ))}
                {filtered.length === 0 && <tr><td colSpan={8} style={{ ...S.td, textAlign:"center", color:"var(--color-text-secondary)", padding:"2rem" }}>No projects match the current filter.</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop:"1rem", display:"flex", gap:8 }}>
            <button style={S.btnP} onClick={() => setStep(2)}>Review summary →</button>
          </div>
        </div>
      )}

      {/* STEP 2 */}
      {step === 2 && (
        <div>
          <h2 style={{ fontSize:16, fontWeight:500, margin:"0 0 1rem" }}>Category summary — {included.length} projects</h2>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:600 }}>
              <thead>
                <tr style={{ background:"var(--color-background-secondary)" }}>
                  {["Category","Projects #","Atlantic Canada #","Total $ Awarded","Atlantic $ Awarded"].map(h => <th key={h} style={S.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {grouped.map(g => (
                  <>
                    <tr key={g.id} style={{ background:"var(--color-background-secondary)" }}>
                      <td style={{ ...S.td, fontWeight:500 }}>{g.label}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:500 }}>{g.totCount||"–"}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:500 }}>{g.atlCount||"–"}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:500 }}>{g.totAmount?"$"+g.totAmount.toLocaleString():"–"}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:500 }}>{g.atlTotal?"$"+g.atlTotal.toLocaleString():"–"}</td>
                    </tr>
                    {g.subs.map(s => (
                      <tr key={s.id}>
                        <td style={{ ...S.td, paddingLeft:28 }}><Badge id={s.id} /></td>
                        <td style={{ ...S.td, textAlign:"right" }}>{s.count||"–"}</td>
                        <td style={{ ...S.td, textAlign:"right" }}>{s.atl||"–"}</td>
                        <td style={{ ...S.td, textAlign:"right" }}>{s.total?"$"+s.total.toLocaleString():"–"}</td>
                        <td style={{ ...S.td, textAlign:"right" }}>{s.atlTotal?"$"+s.atlTotal.toLocaleString():"–"}</td>
                      </tr>
                    ))}
                  </>
                ))}
                <tr style={{ fontWeight:500, borderTop:"1.5px solid var(--color-border-primary)" }}>
                  <td style={S.td}>TOTAL</td>
                  <td style={{ ...S.td, textAlign:"right" }}>{included.length}</td>
                  <td style={{ ...S.td, textAlign:"right" }}>{included.filter(p=>p.isAtlantic).length}</td>
                  <td style={{ ...S.td, textAlign:"right" }}>${grandTotal.toLocaleString()}</td>
                  <td style={{ ...S.td, textAlign:"right" }}>${grandAtl.toLocaleString()}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div style={{ marginTop:"1rem", display:"flex", gap:8 }}>
            <button style={S.btn} onClick={() => setStep(1)}>← Back to projects</button>
            <button style={S.btnP} onClick={() => setStep(3)}>Proceed to download →</button>
          </div>
        </div>
      )}

      {/* STEP 3 */}
      {step === 3 && (
        <div style={{ ...S.card, maxWidth:480 }}>
          <h2 style={{ fontSize:16, fontWeight:500, margin:"0 0 0.75rem" }}>Download consolidated report</h2>
          <p style={{ fontSize:14, color:"var(--color-text-secondary)", lineHeight:1.65, margin:"0 0 1.25rem" }}>
            Excel file includes two sheets: <strong>KPI</strong> (reporting template with parent groups and sub-category totals) and <strong>Projects Detail</strong> (full project list).
          </p>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:"1.25rem" }}>
            {[["Included projects",included.length],["Total awarded","$"+grandTotal.toLocaleString()],["Atlantic Canada",included.filter(p=>p.isAtlantic).length],["Industry projects",included.filter(p=>p.isIndustry).length]].map(([l,v]) => (
              <div key={l} style={S.metric}><div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:4 }}>{l}</div><div style={{ fontSize:18, fontWeight:500 }}>{v}</div></div>
            ))}
          </div>
          <button style={{ ...S.btnP, width:"100%", padding:10 }} onClick={() => { doExport(included); showToast("Downloaded Quarterly_Report.xlsx"); }}>
            Download Quarterly_Report.xlsx
          </button>
        </div>
      )}

      {/* Keywords Modal */}
      {editingKw && kwDraft && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:1000, overflowY:"auto", padding:"40px 16px" }}>
          <div style={{ width:"min(740px,100%)", margin:"0 auto", background:"#ffffff", border:"1px solid #d0d0d0", borderRadius:12, padding:"1.75rem", boxShadow:"0 8px 32px rgba(0,0,0,0.18)" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.75rem" }}>
              <h2 style={{ fontSize:17, fontWeight:600, margin:0, color:"#111" }}>Category keywords</h2>
              <button style={{ border:"1px solid #ccc", background:"#f5f5f5", color:"#333", padding:"5px 14px", borderRadius:6, fontSize:13, cursor:"pointer" }} onClick={() => setEditingKw(false)}>Close</button>
            </div>
            <p style={{ fontSize:13, color:"#555", lineHeight:1.6, margin:"0 0 1.25rem", background:"#f8f8f8", border:"1px solid #e8e8e8", borderRadius:8, padding:"10px 14px" }}>
              If a project's funding agency or program contains any keyword (case-insensitive), it is assigned that sub-category. Keywords are saved locally in your browser for future sessions. Press <strong>Enter</strong> to add a keyword.
            </p>

            {TAXONOMY.map(g => (
              <div key={g.id} style={{ marginBottom:"1.5rem" }}>
                <div style={{ fontSize:13, fontWeight:600, color:"#222", marginBottom:"0.6rem", paddingBottom:"0.4rem", borderBottom:"2px solid #e0e0e0", display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ background:"#185FA5", color:"#fff", borderRadius:4, padding:"2px 8px", fontSize:11, fontWeight:500 }}>{g.label}</span>
                </div>
                {(g.sub.length === 0 ? [{ id:g.id, label:g.label }] : g.sub).map(s => (
                  <div key={s.id} style={{ marginBottom:"1rem", paddingLeft:g.sub.length>0?16:0, paddingTop:"0.5rem", paddingBottom:"0.5rem", borderBottom:"0.5px solid #ececec" }}>
                    <div style={{ marginBottom:8, fontSize:12, fontWeight:600, color:"#333" }}>{s.label}</div>
                    <div style={{ display:"flex", flexWrap:"wrap", gap:6, alignItems:"center" }}>
                      {(kwDraft[s.id]||[]).map((w,i) => (
                        <span key={i} style={{ display:"inline-flex", alignItems:"center", gap:5, background:"#eef4fc", border:"1px solid #b8d0ee", borderRadius:20, padding:"3px 12px", fontSize:12, color:"#1a4a80" }}>
                          {w}
                          <span style={{ cursor:"pointer", color:"#888", fontSize:15, lineHeight:1, fontWeight:500 }} onClick={() => setKwDraft(d => ({ ...d, [s.id]:d[s.id].filter((_,j) => j!==i) }))}>×</span>
                        </span>
                      ))}
                      <input placeholder="+ add keyword, press Enter"
                        style={{ border:"1px solid #ccc", background:"#fff", color:"#222", padding:"4px 10px", borderRadius:20, fontSize:12, minWidth:180, outline:"none" }}
                        value={kwInput[s.id]||""}
                        onChange={e => setKwInput(p => ({ ...p, [s.id]:e.target.value }))}
                        onKeyDown={e => {
                          if (e.key === "Enter" && e.target.value.trim()) {
                            const v = e.target.value.trim();
                            setKwDraft(d => ({ ...d, [s.id]:[...(d[s.id]||[]),v] }));
                            setKwInput(p => ({ ...p, [s.id]:"" }));
                          }
                        }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ))}

            <div style={{ display:"flex", gap:10, paddingTop:"1rem", borderTop:"1px solid #ddd", marginTop:"0.5rem" }}>
              <button style={{ border:"none", background:"#185FA5", color:"#fff", padding:"8px 20px", borderRadius:7, fontSize:13, cursor:"pointer", fontWeight:500 }} onClick={saveKw}>
                Save & re-categorize all projects
              </button>
              <button style={{ border:"1px solid #ccc", background:"#f5f5f5", color:"#333", padding:"8px 16px", borderRadius:7, fontSize:13, cursor:"pointer" }} onClick={() => { setKwDraft(JSON.parse(JSON.stringify(DEFAULT_KEYWORDS))); setKwInput({}); }}>
                Reset to defaults
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
