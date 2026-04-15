import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Brand colours ────────────────────────────────────────────────────────────
const B = {
  green: "#85BE00",
  greenDark: "#6a9800",
  greenLight: "#f0f7e0",
  greenMid: "#d4edaa",
  grey: "#757982",
  greyLight: "#f4f4f5",
  greyMid: "#e2e3e5",
  teal: "#00AC94",
  tealLight: "#e0f7f4",
  cyan: "#00A4EB",
  cyanLight: "#e0f3fd",
  indigo: "#495AD4",
  red: "#D32737",
  orange: "#E87828",
  white: "#ffffff",
  text: "#1a1c1e",
  textMuted: "#5a5d63",
  border: "#dfe0e2",
};

// ─── Taxonomy ─────────────────────────────────────────────────────────────────
const TAXONOMY = [
  { id: "IRAP", label: "IRAP Projects", color: B.cyan, sub: [
    { id: "IRAP_CTO", label: "IRAP CTO Projects" },
    { id: "IRAP_OTHER", label: "Other IRAP Projects" },
  ]},
  { id: "TRICOUNCIL", label: "Tri-Council Industry Partnership", color: B.green, sub: [
    { id: "TC_NSERC", label: "NSERC" },
    { id: "TC_CIHR", label: "CIHR" },
    { id: "TC_SSHRC", label: "SSHRC" },
    { id: "TC_MITACS", label: "Industry Partnership Scholarships (MITACS, ICAN)" },
    { id: "TC_OTHER", label: "Other Tri-Council Industry Partnership" },
  ]},
  { id: "ACOA", label: "ACOA Projects", color: B.teal, sub: [
    { id: "ACOA_AIF", label: "AIF" },
    { id: "ACOA_REGI", label: "REGI (RRRF)" },
    { id: "ACOA_BDP", label: "BDP" },
    { id: "ACOA_ICF", label: "ICF, PBS" },
  ]},
  { id: "PROVINCIAL", label: "Provincial Funding Programs", color: B.indigo, sub: [
    { id: "PROV_VOUCHER", label: "Vouchers" },
    { id: "PROV_OTHER", label: "Other Provincial Funding" },
  ]},
  { id: "GOV_LEVERAGED", label: "Other Industry with Government Leveraged Funding", color: B.orange, sub: [] },
  { id: "RD_CONTRACTS", label: "R&D Contracts", color: B.red, sub: [
    { id: "RD_INDUSTRY", label: "Industry R&D Contracts" },
    { id: "RD_OTHER", label: "Other R&D Contracts" },
    { id: "RD_GOV", label: "Government R&D Contracts" },
  ]},
];

// ─── Industry sectors ─────────────────────────────────────────────────────────
const INDUSTRY_SECTORS = [
  "", "Advanced Manufacturing", "Agriculture & Food (excl. seafood)",
  "Aquaculture & Seafood", "AI, Big Data & Data Analytics", "Aerospace & Defence",
  "Biotechnology", "Cannabis", "Cyber Security", "Energy",
  "Environmental & Clean Technology", "Forestry", "Health & Medical Technology",
  "Mining", "Ocean Technologies", "Oil & Gas", "Other",
];

const INDUSTRY_KEYWORDS = {
  "Advanced Manufacturing": ["manufactur", "fabricat", "machining", "automation", "CNC", "3D print", "robotics"],
  "Agriculture & Food (excl. seafood)": ["agricultur", "food", "crop", "livestock", "farm", "horticultur", "dairy", "poultry"],
  "Aquaculture & Seafood": ["aquacultur", "seafood", "fish", "shellfish", "lobster", "salmon", "mussel", "ocean harvest"],
  "AI, Big Data & Data Analytics": ["artificial intelligence", "machine learning", "big data", "data analytic", "deep learning", "NLP", "neural network", "AI"],
  "Aerospace & Defence": ["aerospace", "aeronautic", "defence", "defense", "aviat", "UAV", "drone", "satellite"],
  "Biotechnology": ["biotech", "genomic", "gene", "biolog", "molecular", "protein", "enzyme", "CRISPR"],
  "Cannabis": ["cannabis", "marijuana", "hemp", "CBD", "THC"],
  "Cyber Security": ["cyber", "security", "cryptograph", "network security", "intrusion", "firewall"],
  "Energy": ["energy", "power", "wind", "solar", "renewabl", "grid", "battery", "hydrogen", "fuel cell"],
  "Environmental & Clean Technology": ["environment", "clean tech", "cleantech", "sustainab", "emission", "carbon", "waste", "pollution", "remediat"],
  "Forestry": ["forest", "timber", "lumber", "wood", "pulp", "paper"],
  "Health & Medical Technology": ["health", "medical", "clinical", "therapeut", "diagnostic", "pharma", "drug", "patient", "hospital", "nursing", "mental health"],
  "Mining": ["mining", "mineral", "ore", "quarry", "excavat"],
  "Ocean Technologies": ["ocean", "marine", "maritime", "underwater", "coastal", "tidal", "offshore"],
  "Oil & Gas": ["oil", "gas", "petroleum", "pipeline", "refin", "hydrocarbon"],
};

function guessIndustrySector(title, agencyIndustry) {
  if (agencyIndustry && agencyIndustry.trim() && agencyIndustry.toLowerCase() !== "nan") {
    const lower = agencyIndustry.toLowerCase();
    for (const sector of INDUSTRY_SECTORS.slice(1)) {
      if (lower.includes(sector.toLowerCase().split("(")[0].trim().toLowerCase())) return sector;
    }
  }
  if (!title) return "";
  const lower = title.toLowerCase();
  for (const [sector, kws] of Object.entries(INDUSTRY_KEYWORDS)) {
    if (kws.some(k => lower.includes(k.toLowerCase()))) return sector;
  }
  return "";
}

// ─── Atlantic auto-detect ─────────────────────────────────────────────────────
const ATLANTIC_KEYWORDS = [
  "atlantic", "nova scotia", "newfoundland", "new brunswick", "prince edward island",
  "pei", "labrador", "nb", "ns", "nl", "moncton", "halifax", "fredericton",
  "charlottetown", "saint john", "cape breton", "annapolis",
];
function guessAtlantic(title, agency) {
  const hay = ((title || "") + " " + (agency || "")).toLowerCase();
  return ATLANTIC_KEYWORDS.some(k => hay.includes(k));
}

function allLeaves() {
  const out = [];
  for (const g of TAXONOMY) {
    if (g.sub.length === 0) out.push({ id: g.id, label: g.label, parentId: g.id, parentLabel: g.label, color: g.color });
    else g.sub.forEach(s => out.push({ id: s.id, label: s.label, parentId: g.id, parentLabel: g.label, color: g.color }));
  }
  out.push({ id: "UNCATEGORIZED", label: "Uncategorized", parentId: null, parentLabel: "", color: B.grey });
  return out;
}
const LEAVES = allLeaves();
function leafLabel(id) { return LEAVES.find(l => l.id === id)?.label || id; }
function leafParentLabel(id) { return LEAVES.find(l => l.id === id)?.parentLabel || ""; }
function leafColor(id) { return LEAVES.find(l => l.id === id)?.color || B.grey; }

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
  RD_INDUSTRY: ["Industry R&D", "Research Contract", "Contrat de recherche", "Industry Partner", "Contract"],
  RD_OTHER: ["Research Agreement", "Research Services", "Testing", "Clinical", "NGO", "Not-for-profit", "Association", "Community"],
  RD_GOV: ["Government of Canada", "Gouvernement du Canada", "Canadian Heritage", "Department of Canadian Heritage", "DND", "Agriculture Canada", "Environment Canada", "Ministère de l'Éducation", "Gouvernement du Nouveau-Brunswick"],
};

function detectCategory(agency, program, kw) {
  if (!agency) return "UNCATEGORIZED";
  const hay = (agency + " " + (program || "")).toLowerCase();
  for (const [cat, words] of Object.entries(kw)) {
    if (words.some(w => w && hay.includes(w.toLowerCase()))) return cat;
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
  const agencyIndustry = find("sponsor industry", "industry", "industrie");
  const cat = detectCategory(agency, program, kw);
  return {
    id: Math.random().toString(36).slice(2),
    institution: sheet, title, pi, agency, program, department: dept,
    amount, startDate, category: cat,
    isAtlantic: guessAtlantic(title, agency),
    industrySector: guessIndustrySector(title, agencyIndustry),
    include: true,
  };
}

// ─── Summary ──────────────────────────────────────────────────────────────────
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
    return { ...g, subs, totCount: subs.reduce((a,s)=>a+s.count,0), totAmount: subs.reduce((a,s)=>a+s.total,0), atlCount: subs.reduce((a,s)=>a+s.atl,0), atlTotal: subs.reduce((a,s)=>a+s.atlTotal,0) };
  });
}

// ─── Excel export ─────────────────────────────────────────────────────────────
function doExport(included) {
  const wb = XLSX.utils.book_new();
  const grouped = groupSummary(buildSummary(included));

  // KPI sheet
  const kpiRows = [
    ["Phase 7 - KPI Metrics","","","","",""],
    ["(To Be Reported Quarterly)","Description","Metric #","Contracts with Atlantic Canadian Industry #","Total $ Awarded","Total $ Awarded with Atlantic Canadian Industry"],
  ];
  let gc=0,gt=0,ga=0,gat=0;
  for (const g of grouped) {
    kpiRows.push([g.label,"",g.totCount,g.atlCount,g.totAmount,g.atlTotal]);
    for (const s of g.subs) kpiRows.push(["  "+s.label,"",s.count,s.atl,s.total,s.atlTotal]);
    gc+=g.totCount; gt+=g.totAmount; ga+=g.atlCount; gat+=g.atlTotal;
  }
  kpiRows.push(["TOTAL","",gc,ga,gt,gat]);
  const ws1 = XLSX.utils.aoa_to_sheet(kpiRows);
  ws1["!cols"] = [{wch:52},{wch:20},{wch:10},{wch:36},{wch:16},{wch:40}];
  XLSX.utils.book_append_sheet(wb, ws1, "KPI");

  // Projects Detail
  const dh = ["Institution","Project Title","PI","Agency","Program","Department","Sub-Category","Group","Start Date","Awarded Amount","Atlantic Canada","Industry Sector"];
  const dr = [dh, ...included.map(p=>[
    p.institution,p.title,p.pi,p.agency,p.program,p.department,
    leafLabel(p.category),leafParentLabel(p.category),
    p.startDate,p.amount,
    p.isAtlantic?"Yes":"No",p.industrySector||"",
  ])];
  const ws2 = XLSX.utils.aoa_to_sheet(dr);
  ws2["!cols"] = [14,40,28,36,28,26,36,32,14,16,14,24].map(wch=>({wch}));
  XLSX.utils.book_append_sheet(wb, ws2, "Projects Detail");

  // R&D by Industry sector
  const rdCats = ["RD_INDUSTRY","RD_OTHER","RD_GOV"];
  const rdLabels = ["Industry R&D Contracts","Other R&D Contracts","Government R&D Contracts"];
  const sectors = INDUSTRY_SECTORS.slice(1);
  const rdHeader = ["", ...sectors];
  const rdRows = [rdHeader];
  for (let ri = 0; ri < rdCats.length; ri++) {
    const rowData = [rdLabels[ri]];
    const catProjects = included.filter(p => p.category === rdCats[ri]);
    for (const sec of sectors) {
      rowData.push(catProjects.filter(p => p.industrySector === sec).length);
    }
    rdRows.push(rowData);
  }
  const ws3 = XLSX.utils.aoa_to_sheet(rdRows);
  ws3["!cols"] = [{wch:28},...sectors.map(()=>({wch:18}))];
  XLSX.utils.book_append_sheet(wb, ws3, "R&D by Industry");

  const blob = XLSX.write(wb, { bookType:"xlsx", type:"array" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([blob],{type:"application/octet-stream"}));
  a.download = "Quarterly_Report.xlsx"; a.click();
}

// ─── Badge component ──────────────────────────────────────────────────────────
function Badge({ id }) {
  const c = leafColor(id);
  return <span style={{ display:"inline-block", fontSize:11, padding:"2px 9px", borderRadius:20, background:c+"18", color:c, fontWeight:600, border:`1px solid ${c}44`, whiteSpace:"nowrap" }}>{leafLabel(id)}</span>;
}

// ─── Keyword modal ────────────────────────────────────────────────────────────
function KeywordModal({ kwDraft, setKwDraft, kwInput, setKwInput, onSave, onClose }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(26,28,30,0.6)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:2000, overflowY:"auto", padding:"32px 16px" }}>
      <div style={{ width:"min(760px,100%)", margin:"0 auto", background:B.white, borderRadius:14, boxShadow:"0 20px 60px rgba(0,0,0,0.25)", overflow:"hidden" }}>
        {/* Modal header */}
        <div style={{ background:B.green, padding:"1rem 1.5rem", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:B.white }}>Category Keywords</div>
            <div style={{ fontSize:12, color:"rgba(255,255,255,0.85)", marginTop:2 }}>Manage how projects are auto-categorized</div>
          </div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,0.2)", border:"none", color:B.white, width:32, height:32, borderRadius:8, fontSize:18, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}>×</button>
        </div>
        <div style={{ padding:"1.25rem 1.5rem" }}>
          <div style={{ fontSize:13, color:B.textMuted, lineHeight:1.6, marginBottom:"1.25rem", background:B.greenLight, border:`1px solid ${B.greenMid}`, borderRadius:8, padding:"10px 14px" }}>
            If a project's <strong>funding agency</strong> or <strong>program name</strong> contains any keyword (case-insensitive), it is assigned that sub-category. Saved to your browser for future sessions. Press <strong>Enter</strong> to add a keyword.
          </div>
          <div style={{ maxHeight:"55vh", overflowY:"auto", paddingRight:4 }}>
            {TAXONOMY.map(g => (
              <div key={g.id} style={{ marginBottom:"1.5rem" }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:"0.6rem", paddingBottom:"0.5rem", borderBottom:`2px solid ${g.color}44` }}>
                  <span style={{ background:g.color, color:B.white, borderRadius:5, padding:"3px 10px", fontSize:11, fontWeight:700, letterSpacing:"0.3px" }}>{g.label}</span>
                </div>
                {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s => (
                  <div key={s.id} style={{ marginBottom:"0.9rem", paddingLeft:g.sub.length>0?20:0 }}>
                    <div style={{ fontSize:12, fontWeight:600, color:B.text, marginBottom:6 }}>{s.label}</div>
                    <div style={{ display:"flex", flexWrap:"wrap", gap:6, alignItems:"center" }}>
                      {(kwDraft[s.id]||[]).map((w,i) => (
                        <span key={i} style={{ display:"inline-flex", alignItems:"center", gap:5, background:B.greyLight, border:`1px solid ${B.border}`, borderRadius:20, padding:"3px 11px", fontSize:12, color:B.text }}>
                          {w}
                          <span style={{ cursor:"pointer", color:B.red, fontSize:14, lineHeight:1, fontWeight:700 }} onClick={()=>setKwDraft(d=>({...d,[s.id]:d[s.id].filter((_,j)=>j!==i)}))}>×</span>
                        </span>
                      ))}
                      <input placeholder="+ add keyword, press Enter"
                        style={{ border:`1.5px solid ${B.border}`, background:B.white, color:B.text, padding:"4px 12px", borderRadius:20, fontSize:12, minWidth:180, outline:"none" }}
                        value={kwInput[s.id]||""}
                        onChange={e=>setKwInput(p=>({...p,[s.id]:e.target.value}))}
                        onKeyDown={e=>{
                          if (e.key==="Enter"&&e.target.value.trim()) {
                            const v=e.target.value.trim();
                            setKwDraft(d=>({...d,[s.id]:[...(d[s.id]||[]),v]}));
                            setKwInput(p=>({...p,[s.id]:""}));
                          }
                        }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </div>
          <div style={{ display:"flex", gap:10, paddingTop:"1rem", borderTop:`1px solid ${B.border}`, marginTop:"0.5rem" }}>
            <button style={{ background:B.green, color:B.white, border:"none", padding:"9px 22px", borderRadius:8, fontSize:13, cursor:"pointer", fontWeight:700 }} onClick={onSave}>
              Save & re-categorize all projects
            </button>
            <button style={{ background:B.greyLight, color:B.text, border:`1px solid ${B.border}`, padding:"9px 16px", borderRadius:8, fontSize:13, cursor:"pointer" }}
              onClick={()=>{setKwDraft(JSON.parse(JSON.stringify(DEFAULT_KEYWORDS)));setKwInput({});}}>
              Reset to defaults
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Shared top bar with keywords button ──────────────────────────────────────
function TopBar({ step, setStep, hasProjects, onKwClick }) {
  const STEPS = ["Upload","Review & Tag","Summary","Download"];
  return (
    <div style={{ background:B.white, borderBottom:`2px solid ${B.green}`, marginBottom:"1.75rem", paddingBottom:0 }}>
      <div style={{ maxWidth:1280, margin:"0 auto", padding:"0 1.25rem" }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"0.9rem 0 0" }}>
          <div style={{ display:"flex", alignItems:"center", gap:12 }}>
            <div style={{ width:38, height:38, background:B.green, borderRadius:8, display:"flex", alignItems:"center", justifyContent:"center" }}>
              <span style={{ color:B.white, fontWeight:800, fontSize:16 }}>S</span>
            </div>
            <div>
              <div style={{ fontSize:17, fontWeight:700, color:B.text, lineHeight:1.2 }}>Quarterly Report Builder</div>
              <div style={{ fontSize:12, color:B.grey }}>Springboard Atlantic — Research Funding Consolidation</div>
            </div>
          </div>
          <button onClick={onKwClick}
            style={{ display:"flex", alignItems:"center", gap:7, background:B.green, color:B.white, border:"none", padding:"8px 18px", borderRadius:8, fontSize:13, fontWeight:700, cursor:"pointer", boxShadow:"0 2px 8px rgba(133,190,0,0.35)" }}>
            <span style={{ fontSize:15 }}>⚙</span> Category Keywords
          </button>
        </div>
        <div style={{ display:"flex", gap:0, marginTop:"0.75rem" }}>
          {STEPS.map((label,i) => {
            const active = step===i;
            const done = i < step;
            return (
              <button key={i} onClick={()=>(hasProjects||i===0)&&setStep(i)}
                style={{ border:"none", background:"transparent", padding:"8px 22px", fontSize:13, fontWeight:active?700:400, color:active?B.green:done?B.grey:B.grey, cursor:"pointer", borderBottom:active?`3px solid ${B.green}`:"3px solid transparent", transition:"all 0.15s", marginBottom:-2 }}>
                <span style={{ display:"inline-flex", alignItems:"center", gap:6 }}>
                  <span style={{ width:20, height:20, borderRadius:10, background:active?B.green:done?B.greenMid:B.greyMid, color:active?B.white:B.textMuted, fontSize:11, fontWeight:700, display:"inline-flex", alignItems:"center", justifyContent:"center" }}>{i+1}</span>
                  {label}
                </span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep] = useState(0);
  const [projects, setProjects] = useState([]);
  const [keywords, setKeywords] = useState(()=>{ try{const s=localStorage.getItem("qr_kw_v2");return s?JSON.parse(s):DEFAULT_KEYWORDS;}catch{return DEFAULT_KEYWORDS;} });
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

  const showToast = (msg,type="success") => { setToast({msg,type}); setTimeout(()=>setToast(""),3200); };

  const openKw = () => { setKwDraft(JSON.parse(JSON.stringify(keywords))); setKwInput({}); setEditingKw(true); };

  const handleFile = async e => {
    const file = e.target.files[0]; if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf,{type:"array"});
    const all = [];
    for (const sn of wb.SheetNames) {
      const rows = parseSheet(XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null}));
      rows.forEach(r=>all.push(normalize(r,sn,keywords)));
    }
    setProjects(all);
    showToast(`Loaded ${all.length} projects from ${wb.SheetNames.length} sheet(s)`);
    setStep(1); e.target.value="";
  };

  const recategorize = useCallback(kw=>{
    setProjects(prev=>prev.map(p=>({...p,category:detectCategory(p.agency,p.program,kw)})));
  },[]);

  const upd = (id,field,val) => setProjects(prev=>prev.map(p=>p.id===id?{...p,[field]:val}:p));
  const included = projects.filter(p=>p.include);

  const filtered = projects
    .filter(p=>filterCat==="All"||p.category===filterCat)
    .sort((a,b)=>{
      const av=a[sortCol]??"", bv=b[sortCol]??"";
      const cmp=typeof av==="number"?av-bv:String(av).localeCompare(String(bv));
      return sortDir==="asc"?cmp:-cmp;
    });

  const toggleSort = col=>{ if(sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc"); else{setSortCol(col);setSortDir("asc");} };

  const addProject = () => {
    const cat = detectCategory(newProj.agency||"",newProj.program||"",keywords);
    setProjects(prev=>[...prev,{
      id:Math.random().toString(36).slice(2),
      institution:newProj.institution||"",title:newProj.title||"",pi:newProj.pi||"",
      agency:newProj.agency||"",program:newProj.program||"",department:newProj.department||"",
      amount:parseFloat(newProj.amount)||0,startDate:newProj.startDate||"",
      category:newProj.category||cat,
      isAtlantic:!!newProj.isAtlantic,
      industrySector:newProj.industrySector||"",
      include:true,
    }]);
    setAddRow(false); setNewProj({}); showToast("Project added");
  };

  const saveKw = () => {
    setKeywords(kwDraft);
    try{localStorage.setItem("qr_kw_v2",JSON.stringify(kwDraft));}catch{}
    recategorize(kwDraft);
    setEditingKw(false);
    showToast("Keywords saved — projects re-categorized");
  };

  const summary = buildSummary(included);
  const grouped = groupSummary(summary);
  const grandTotal = included.reduce((s,p)=>s+p.amount,0);
  const grandAtl = included.filter(p=>p.isAtlantic).reduce((s,p)=>s+p.amount,0);

  // Shared styles
  const card = { background:B.white, border:`1px solid ${B.border}`, borderRadius:12, padding:"1.25rem" };
  const btnGreen = { background:B.green, color:B.white, border:"none", padding:"8px 20px", borderRadius:8, fontSize:13, cursor:"pointer", fontWeight:700 };
  const btnOutline = { background:B.white, color:B.text, border:`1px solid ${B.border}`, padding:"7px 16px", borderRadius:8, fontSize:13, cursor:"pointer" };
  const th = { padding:"9px 12px", fontSize:12, fontWeight:700, color:B.textMuted, textAlign:"left", background:B.greyLight, borderBottom:`1px solid ${B.border}`, whiteSpace:"nowrap", cursor:"pointer", userSelect:"none" };
  const td = { padding:"8px 12px", fontSize:13, borderBottom:`1px solid ${B.greyMid}`, verticalAlign:"middle" };
  const sel = { border:`1px solid ${B.border}`, background:B.white, color:B.text, padding:"6px 10px", borderRadius:7, fontSize:13 };
  const inp = { border:`1px solid ${B.border}`, background:B.white, color:B.text, padding:"6px 10px", borderRadius:7, fontSize:13 };
  const metric = { background:B.greenLight, border:`1px solid ${B.greenMid}`, borderRadius:10, padding:"1rem 1.1rem" };

  return (
    <div style={{ fontFamily:"'Inter', 'Segoe UI', system-ui, sans-serif", color:B.text, minHeight:"100vh", background:"#f7f8f9" }}>
      <TopBar step={step} setStep={setStep} hasProjects={projects.length>0} onKwClick={openKw} />

      {toast && (
        <div style={{ position:"fixed", bottom:24, right:24, background:B.text, color:B.white, padding:"11px 22px", borderRadius:10, fontSize:13, zIndex:9999, display:"flex", alignItems:"center", gap:8, boxShadow:"0 4px 20px rgba(0,0,0,0.2)" }}>
          <span style={{ color:B.green, fontSize:16 }}>✓</span> {toast.msg}
        </div>
      )}

      <div style={{ maxWidth:1280, margin:"0 auto", padding:"0 1.25rem 3rem" }}>

        {/* STEP 0 */}
        {step===0&&(
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"1.5rem", maxWidth:860 }}>
            <div style={card}>
              <div style={{ fontSize:15, fontWeight:700, marginBottom:"0.6rem", color:B.text }}>Upload funding export</div>
              <p style={{ fontSize:13, color:B.textMuted, lineHeight:1.65, margin:"0 0 1.25rem" }}>
                Upload the .xlsx exported from ROMEO or your institution's research system. Each worksheet is treated as one institution. Columns are auto-detected.
              </p>
              <div style={{ border:`2px dashed ${B.greenMid}`, borderRadius:10, padding:"2rem", textAlign:"center", marginBottom:"0.75rem", background:B.greenLight }}>
                <div style={{ fontSize:13, color:B.textMuted, marginBottom:"0.75rem" }}>Multi-sheet .xlsx supported (ROMEO, Banner, etc.)</div>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
                <button style={btnGreen} onClick={()=>fileRef.current.click()}>Choose file</button>
              </div>
              <p style={{ fontSize:11, color:B.grey, margin:0 }}>Auto-detects: Project Title, Agency, PI, Amount, Program, Department, Start Date</p>
            </div>
            <div style={card}>
              <div style={{ fontSize:15, fontWeight:700, marginBottom:"0.6rem", color:B.text }}>How categorization works</div>
              <div style={{ fontSize:13, color:B.textMuted, lineHeight:1.7 }}>
                <p style={{ margin:"0 0 0.75rem" }}>When you upload a file, each project is automatically assigned a category based on its funding agency and program name matched against your keyword rules.</p>
                <p style={{ margin:"0 0 0.75rem" }}>Projects with <strong>Atlantic</strong>, <strong>Nova Scotia</strong>, <strong>New Brunswick</strong>, <strong>Newfoundland</strong>, or <strong>PEI</strong> in the title or agency are automatically flagged as Atlantic Canada.</p>
                <p style={{ margin:0 }}>Use the <strong style={{ color:B.green }}>⚙ Category Keywords</strong> button (top right) at any time to customize keyword rules. Changes are saved to your browser.</p>
              </div>
            </div>
          </div>
        )}

        {/* STEP 1 */}
        {step===1&&(
          <div>
            {/* Metric cards */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:12, marginBottom:"1.25rem" }}>
              {[
                {label:"Total projects",val:projects.length,color:B.grey},
                {label:"Included",val:included.length,color:B.green},
                {label:"Total awarded",val:"$"+grandTotal.toLocaleString(),color:B.teal},
                {label:"Atlantic Canada",val:included.filter(p=>p.isAtlantic).length,color:B.cyan},
              ].map(m=>(
                <div key={m.label} style={{ ...metric, borderLeft:`4px solid ${m.color}` }}>
                  <div style={{ fontSize:11, color:B.textMuted, marginBottom:4, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.5px" }}>{m.label}</div>
                  <div style={{ fontSize:22, fontWeight:800, color:m.color }}>{m.val}</div>
                </div>
              ))}
            </div>

            {/* Toolbar */}
            <div style={{ display:"flex", gap:10, marginBottom:"0.9rem", flexWrap:"wrap", alignItems:"center", background:B.white, border:`1px solid ${B.border}`, borderRadius:10, padding:"0.75rem 1rem" }}>
              <div style={{ fontSize:12, fontWeight:700, color:B.textMuted, marginRight:4 }}>FILTER:</div>
              <select style={sel} value={filterCat} onChange={e=>setFilterCat(e.target.value)}>
                <option value="All">All categories ({projects.length})</option>
                {TAXONOMY.map(g=>(
                  <optgroup key={g.id} label={g.label}>
                    {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s=>{
                      const cnt = projects.filter(p=>p.category===s.id).length;
                      return <option key={s.id} value={s.id}>{s.label} ({cnt})</option>;
                    })}
                  </optgroup>
                ))}
                <option value="UNCATEGORIZED">Uncategorized ({projects.filter(p=>p.category==="UNCATEGORIZED").length})</option>
              </select>
              <div style={{ flex:1 }} />
              <button style={btnOutline} onClick={()=>{setAddRow(true);setNewProj({});}}>+ Add project</button>
            </div>

            {/* Table */}
            <div style={{ ...card, padding:0, overflow:"hidden" }}>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", minWidth:1020 }}>
                  <thead>
                    <tr>
                      {[["include","✓",40],["title","Project Title",220],["pi","PI",120],["agency","Agency",160],["amount","Amount",95],["category","Category",180],["isAtlantic","Atlantic",80],["industrySector","Industry Sector",170]].map(([col,lbl,w])=>(
                        <th key={col} style={{...th,width:w}} onClick={()=>toggleSort(col)}>
                          {lbl}{sortCol===col?(sortDir==="asc"?" ↑":" ↓"):""}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {addRow&&(
                      <tr style={{ background:B.greenLight }}>
                        <td style={td}></td>
                        <td style={td}><input style={{...inp,width:"100%"}} placeholder="Project title" value={newProj.title||""} onChange={e=>setNewProj(p=>({...p,title:e.target.value}))}/></td>
                        <td style={td}><input style={{...inp,width:"100%"}} placeholder="PI" value={newProj.pi||""} onChange={e=>setNewProj(p=>({...p,pi:e.target.value}))}/></td>
                        <td style={td}><input style={{...inp,width:"100%"}} placeholder="Agency" value={newProj.agency||""} onChange={e=>setNewProj(p=>({...p,agency:e.target.value}))}/></td>
                        <td style={td}><input style={{...inp,width:80}} type="number" placeholder="Amount" value={newProj.amount||""} onChange={e=>setNewProj(p=>({...p,amount:e.target.value}))}/></td>
                        <td style={td}>
                          <select style={{...sel,fontSize:11}} value={newProj.category||"UNCATEGORIZED"} onChange={e=>setNewProj(p=>({...p,category:e.target.value}))}>
                            {TAXONOMY.map(g=>(
                              <optgroup key={g.id} label={g.label}>
                                {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s=><option key={s.id} value={s.id}>{s.label}</option>)}
                              </optgroup>
                            ))}
                            <option value="UNCATEGORIZED">Uncategorized</option>
                          </select>
                        </td>
                        <td style={{...td,textAlign:"center"}}>
                          <input type="checkbox" checked={!!newProj.isAtlantic} onChange={e=>setNewProj(p=>({...p,isAtlantic:e.target.checked}))}/>
                        </td>
                        <td style={td}>
                          <div style={{display:"flex",gap:6}}>
                            <select style={{...sel,fontSize:11,flex:1}} value={newProj.industrySector||""} onChange={e=>setNewProj(p=>({...p,industrySector:e.target.value}))}>
                              {INDUSTRY_SECTORS.map(s=><option key={s} value={s}>{s||"— select —"}</option>)}
                            </select>
                            <button style={{...btnGreen,fontSize:11,padding:"3px 10px"}} onClick={addProject}>Save</button>
                            <button style={{...btnOutline,fontSize:11,padding:"3px 8px"}} onClick={()=>setAddRow(false)}>×</button>
                          </div>
                        </td>
                      </tr>
                    )}
                    {filtered.map((p,ri)=>(
                      <tr key={p.id} style={{ background:ri%2===0?B.white:B.greyLight, opacity:p.include?1:0.38 }}>
                        <td style={{...td,textAlign:"center"}}>
                          <input type="checkbox" checked={p.include} onChange={e=>upd(p.id,"include",e.target.checked)}/>
                        </td>
                        <td style={{...td,maxWidth:220}}>
                          <div style={{fontSize:13,lineHeight:1.4}}>{p.title}</div>
                          {p.department&&<div style={{fontSize:11,color:B.textMuted,marginTop:2}}>{p.department}</div>}
                        </td>
                        <td style={{...td,fontSize:12}}>{p.pi}</td>
                        <td style={{...td,fontSize:12,color:B.textMuted,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.agency}</td>
                        <td style={{...td,textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:500}}>${p.amount.toLocaleString()}</td>
                        <td style={td}>
                          <select style={{...sel,fontSize:11,padding:"2px 6px"}} value={p.category} onChange={e=>upd(p.id,"category",e.target.value)}>
                            {TAXONOMY.map(g=>(
                              <optgroup key={g.id} label={g.label}>
                                {(g.sub.length===0?[{id:g.id,label:g.label}]:g.sub).map(s=><option key={s.id} value={s.id}>{s.label}</option>)}
                              </optgroup>
                            ))}
                            <option value="UNCATEGORIZED">Uncategorized</option>
                          </select>
                        </td>
                        <td style={{...td,textAlign:"center"}}>
                          <span onClick={()=>upd(p.id,"isAtlantic",!p.isAtlantic)} style={{ cursor:"pointer", display:"inline-flex", alignItems:"center", justifyContent:"center", width:26, height:26, borderRadius:6, background:p.isAtlantic?B.teal:B.greyMid, color:p.isAtlantic?B.white:B.grey, fontSize:13, fontWeight:700, userSelect:"none" }}>
                            {p.isAtlantic?"✓":"–"}
                          </span>
                        </td>
                        <td style={td}>
                          <select style={{...sel,fontSize:11,padding:"2px 6px",width:"100%"}} value={p.industrySector||""} onChange={e=>upd(p.id,"industrySector",e.target.value)}>
                            {INDUSTRY_SECTORS.map(s=><option key={s} value={s}>{s||"— none —"}</option>)}
                          </select>
                        </td>
                      </tr>
                    ))}
                    {filtered.length===0&&(
                      <tr><td colSpan={8} style={{...td,textAlign:"center",color:B.textMuted,padding:"2.5rem"}}>No projects match the current filter.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
            <div style={{marginTop:"1rem",display:"flex",gap:8}}>
              <button style={btnGreen} onClick={()=>setStep(2)}>Review summary →</button>
            </div>
          </div>
        )}

        {/* STEP 2 */}
        {step===2&&(
          <div>
            <div style={{fontSize:16,fontWeight:700,marginBottom:"1rem"}}>Category summary — {included.length} projects included</div>
            <div style={{...card,padding:0,overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead>
                  <tr>
                    {["Category","Projects #","Atlantic Canada #","Total $ Awarded","Atlantic $ Awarded"].map(h=>(
                      <th key={h} style={th}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {grouped.map(g=>(
                    <>
                      <tr key={g.id} style={{background:g.color+"14",borderLeft:`4px solid ${g.color}`}}>
                        <td style={{...td,fontWeight:700,color:B.text,paddingLeft:16}}>{g.label}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{g.totCount||"–"}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{g.atlCount||"–"}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{g.totAmount?"$"+g.totAmount.toLocaleString():"–"}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{g.atlTotal?"$"+g.atlTotal.toLocaleString():"–"}</td>
                      </tr>
                      {g.subs.map(s=>(
                        <tr key={s.id} style={{background:B.white}}>
                          <td style={{...td,paddingLeft:32}}><Badge id={s.id}/></td>
                          <td style={{...td,textAlign:"right"}}>{s.count||"–"}</td>
                          <td style={{...td,textAlign:"right"}}>{s.atl||"–"}</td>
                          <td style={{...td,textAlign:"right"}}>{s.total?"$"+s.total.toLocaleString():"–"}</td>
                          <td style={{...td,textAlign:"right"}}>{s.atlTotal?"$"+s.atlTotal.toLocaleString():"–"}</td>
                        </tr>
                      ))}
                    </>
                  ))}
                  <tr style={{background:B.greenLight,borderTop:`2px solid ${B.green}`}}>
                    <td style={{...td,fontWeight:800,color:B.text}}>TOTAL</td>
                    <td style={{...td,textAlign:"right",fontWeight:800}}>{included.length}</td>
                    <td style={{...td,textAlign:"right",fontWeight:800}}>{included.filter(p=>p.isAtlantic).length}</td>
                    <td style={{...td,textAlign:"right",fontWeight:800}}>${grandTotal.toLocaleString()}</td>
                    <td style={{...td,textAlign:"right",fontWeight:800}}>${grandAtl.toLocaleString()}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* R&D by industry mini preview */}
            <div style={{fontSize:15,fontWeight:700,margin:"1.5rem 0 0.75rem"}}>R&D contracts by industry sector</div>
            <div style={{...card,padding:0,overflow:"hidden",overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",minWidth:600}}>
                <thead>
                  <tr>
                    <th style={{...th,minWidth:180}}>Category</th>
                    {INDUSTRY_SECTORS.slice(1).map(s=><th key={s} style={{...th,fontSize:11,minWidth:90}}>{s}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {[["RD_INDUSTRY","Industry R&D Contracts"],["RD_OTHER","Other R&D Contracts"],["RD_GOV","Government R&D Contracts"]].map(([cat,label])=>(
                    <tr key={cat}>
                      <td style={{...td,fontWeight:600,fontSize:12}}>{label}</td>
                      {INDUSTRY_SECTORS.slice(1).map(sec=>{
                        const cnt=included.filter(p=>p.category===cat&&p.industrySector===sec).length;
                        return <td key={sec} style={{...td,textAlign:"center",fontWeight:cnt>0?600:400,color:cnt>0?B.green:B.textMuted}}>{cnt||"–"}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div style={{marginTop:"1rem",display:"flex",gap:8}}>
              <button style={btnOutline} onClick={()=>setStep(1)}>← Back to projects</button>
              <button style={btnGreen} onClick={()=>setStep(3)}>Proceed to download →</button>
            </div>
          </div>
        )}

        {/* STEP 3 */}
        {step===3&&(
          <div style={{maxWidth:500}}>
            <div style={card}>
              <div style={{fontSize:16,fontWeight:700,marginBottom:"0.75rem"}}>Download consolidated report</div>
              <p style={{fontSize:13,color:B.textMuted,lineHeight:1.65,margin:"0 0 1.25rem"}}>
                The Excel file includes three sheets: <strong>KPI</strong> (reporting template with parent groups and sub-category totals), <strong>Projects Detail</strong> (full project list), and <strong>R&D by Industry</strong> (R&D contracts broken down by industry sector).
              </p>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:"1.25rem"}}>
                {[
                  {label:"Included projects",val:included.length,color:B.green},
                  {label:"Total awarded",val:"$"+grandTotal.toLocaleString(),color:B.teal},
                  {label:"Atlantic Canada",val:included.filter(p=>p.isAtlantic).length,color:B.cyan},
                  {label:"With industry sector",val:included.filter(p=>p.industrySector).length,color:B.orange},
                ].map(m=>(
                  <div key={m.label} style={{...metric,borderLeft:`4px solid ${m.color}`}}>
                    <div style={{fontSize:11,color:B.textMuted,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.4px"}}>{m.label}</div>
                    <div style={{fontSize:20,fontWeight:800,color:m.color}}>{m.val}</div>
                  </div>
                ))}
              </div>
              <button style={{...btnGreen,width:"100%",padding:"11px",fontSize:14}} onClick={()=>{doExport(included);showToast("Downloaded Quarterly_Report.xlsx");}}>
                Download Quarterly_Report.xlsx
              </button>
            </div>
          </div>
        )}
      </div>

      {editingKw&&kwDraft&&(
        <KeywordModal kwDraft={kwDraft} setKwDraft={setKwDraft} kwInput={kwInput} setKwInput={setKwInput}
          onSave={saveKw} onClose={()=>setEditingKw(false)} />
      )}
    </div>
  );
}
