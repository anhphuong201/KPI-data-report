import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/xlsx.mjs";

const DEFAULT_KEYWORDS = {
  NSERC: ["Natural Sciences and Engineering Research Council", "NSERC", "Conseil de recherches en sciences naturelles et en génie", "CRSNG"],
  CIHR: ["Canadian Institutes of Health Research", "CIHR", "Instituts de recherche en santé du Canada", "IRSC"],
  SSHRC: ["Social Sciences and Humanities Research Council", "SSHRC", "Conseil de recherches en sciences humaines", "CRSH"],
  MITACS: ["Mitacs", "MITACS"],
  IRAP: ["IRAP", "Industrial Research Assistance Program", "National Research Council", "NRC"],
  ACOA: ["ACOA", "Atlantic Canada Opportunities Agency", "Agence de promotion économique du Canada atlantique", "AIF", "REGI", "BDP"],
  PROVINCIAL: ["Province", "Provincial", "Gouvernement du Nouveau-Brunswick", "Government of New Brunswick", "Government of Nova Scotia", "Government of Newfoundland and Labrador", "Government of Prince Edward Island", "PEI"],
  INDUSTRY_CONTRACT: ["Contract", "Contrat", "Industry", "Industrie", "Corporation", "Inc.", "Ltd.", "LLC"],
  OTHER_FEDERAL: ["Government of Canada", "Gouvernement du Canada", "Canadian Heritage", "Department of", "Ministère de"],
  CHARITABLE: ["Foundation", "Fondation", "Association", "Society", "Société", "Consortium"],
};

const CATEGORY_LABELS = {
  NSERC: "NSERC",
  CIHR: "CIHR",
  SSHRC: "SSHRC",
  MITACS: "Industry Partnership Scholarships (MITACS, ICAN)",
  IRAP: "IRAP Projects",
  ACOA: "ACOA Projects",
  PROVINCIAL: "Provincial Funding Programs",
  INDUSTRY_CONTRACT: "R&D Contracts – Industry",
  OTHER_FEDERAL: "Other Federal",
  CHARITABLE: "Charitable / Foundations",
  UNCATEGORIZED: "Uncategorized",
};

const TRI_COUNCIL = ["NSERC", "CIHR", "SSHRC"];

function detectCategory(agencyText, keywords) {
  if (!agencyText) return "UNCATEGORIZED";
  const lower = agencyText.toLowerCase();
  for (const [cat, words] of Object.entries(keywords)) {
    if (words.some(w => lower.includes(w.toLowerCase()))) return cat;
  }
  return "UNCATEGORIZED";
}

function parseRomeoSheet(data) {
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const vals = row.map(v => String(v || "").toLowerCase());
    if (vals.some(v => v.includes("project title") || v.includes("agency") || v.includes("titre") || v.includes("agence"))) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) return [];
  const headers = data[headerRow].map(v => String(v || "").trim());
  const rows = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(v => v === null || v === undefined || v === "")) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx] !== undefined ? row[idx] : ""; });
    rows.push(obj);
  }
  return rows;
}

function normalizeProject(rawObj, sheetName, keywords) {
  const keys = Object.keys(rawObj);
  const find = (...terms) => {
    for (const t of terms) {
      const k = keys.find(k => k.toLowerCase().includes(t.toLowerCase()));
      if (k !== undefined) return String(rawObj[k] || "");
    }
    return "";
  };
  const agency = find("agency", "agence", "funding agency", "sponsor source", "funding");
  const title = find("project title", "titre");
  const pi = find("principal investigator", "investigator info", "chercheur");
  const amount = parseFloat(String(find("awarded amount", "amount", "montant")).replace(/[^0-9.]/g, "")) || 0;
  const start = find("start date", "date de début", "sponsor start");
  const program = find("program", "programme");
  const dept = find("department", "département", "faculty", "faculté");
  const category = detectCategory(agency, keywords);
  return {
    id: Math.random().toString(36).slice(2),
    institution: sheetName,
    title,
    pi,
    agency,
    program,
    department: dept,
    amount,
    startDate: start,
    category,
    isAtlantic: false,
    isIndustry: category === "INDUSTRY_CONTRACT",
    include: true,
  };
}

const STEPS = ["Upload", "Review & Tag", "Categories", "Download"];

export default function App() {
  const [step, setStep] = useState(0);
  const [projects, setProjects] = useState([]);
  const [keywords, setKeywords] = useState(() => {
    try { const s = localStorage.getItem("reportKeywords"); return s ? JSON.parse(s) : DEFAULT_KEYWORDS; } catch { return DEFAULT_KEYWORDS; }
  });
  const [editingKw, setEditingKw] = useState(false);
  const [kwDraft, setKwDraft] = useState(null);
  const [newCatName, setNewCatName] = useState("");
  const [filterInstitution, setFilterInstitution] = useState("All");
  const [filterCategory, setFilterCategory] = useState("All");
  const [sortCol, setSortCol] = useState("title");
  const [sortDir, setSortDir] = useState("asc");
  const [addRow, setAddRow] = useState(false);
  const [newProj, setNewProj] = useState({});
  const [toast, setToast] = useState("");
  const fileRef = useRef();

  useEffect(() => {
    try { localStorage.setItem("reportKeywords", JSON.stringify(keywords)); } catch {}
  }, [keywords]);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(""), 3000); };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const allProjects = [];
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const rows = parseRomeoSheet(data);
      rows.forEach(r => allProjects.push(normalizeProject(r, sheetName, keywords)));
    }
    setProjects(allProjects);
    showToast(`Loaded ${allProjects.length} projects from ${wb.SheetNames.length} sheet(s)`);
    setStep(1);
  };

  const recategorize = useCallback((kw) => {
    setProjects(prev => prev.map(p => ({
      ...p,
      category: detectCategory(p.agency, kw),
      isIndustry: detectCategory(p.agency, kw) === "INDUSTRY_CONTRACT" ? true : p.isIndustry,
    })));
  }, []);

  const institutions = ["All", ...new Set(projects.map(p => p.institution))];
  const categories = ["All", ...Object.keys(CATEGORY_LABELS)];

  const filtered = projects.filter(p =>
    (filterInstitution === "All" || p.institution === filterInstitution) &&
    (filterCategory === "All" || p.category === filterCategory)
  ).sort((a, b) => {
    const av = a[sortCol] ?? ""; const bv = b[sortCol] ?? "";
    if (typeof av === "number") return sortDir === "asc" ? av - bv : bv - av;
    return sortDir === "asc" ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
  });

  const included = projects.filter(p => p.include);

  const toggleSort = (col) => {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  const updateProject = (id, field, val) => setProjects(prev => prev.map(p => p.id === id ? { ...p, [field]: val } : p));

  const summaryByCat = {};
  for (const p of included) {
    if (!summaryByCat[p.category]) summaryByCat[p.category] = { count: 0, total: 0, atlanticCount: 0, atlanticTotal: 0 };
    summaryByCat[p.category].count++;
    summaryByCat[p.category].total += p.amount;
    if (p.isAtlantic) { summaryByCat[p.category].atlanticCount++; summaryByCat[p.category].atlanticTotal += p.amount; }
  }

  const generateExcel = () => {
    const wb2 = XLSX.utils.book_new();

    // KPI sheet
    const rows = [
      ["Phase 7 - KPI's Metrics", "", "", "", "", ""],
      ["(To Be Reported Quarterly)", "Description", "Metric #", "Contracts with Atlantic Canadian Industry #", "Total $ Awarded", "Total $ Awarded with Atlantic Canadian Industry"],
    ];

    const catOrder = ["INDUSTRY_CONTRACT", "IRAP", "NSERC", "CIHR", "SSHRC", "MITACS", "ACOA", "PROVINCIAL", "OTHER_FEDERAL", "CHARITABLE", "UNCATEGORIZED"];
    const triCouncilRows = [];
    let grandCount = 0, grandTotal = 0, grandAtlCount = 0, grandAtlTotal = 0;

    for (const cat of catOrder) {
      const d = summaryByCat[cat] || { count: 0, total: 0, atlanticCount: 0, atlanticTotal: 0 };
      grandCount += d.count; grandTotal += d.total;
      grandAtlCount += d.atlanticCount; grandAtlTotal += d.atlanticTotal;
      const label = CATEGORY_LABELS[cat] || cat;
      if (TRI_COUNCIL.includes(cat)) {
        triCouncilRows.push([`  ${label}`, "", d.count, d.atlanticCount, d.total, d.atlanticTotal]);
      } else {
        rows.push([label, "", d.count, d.atlanticCount, d.total, d.atlanticTotal]);
        if (cat === "IRAP") {
          rows.push(["  Tri-Council Industry Partnerships", "Projects with industry – Alliance, ARD, IE (CCI)", "", "", "", ""]);
          triCouncilRows.forEach(r => rows.push(r));
          triCouncilRows.length = 0;
        }
      }
    }
    rows.push(["TOTAL", "", grandCount, grandAtlCount, grandTotal, grandAtlTotal]);
    const ws1 = XLSX.utils.aoa_to_sheet(rows);
    ws1["!cols"] = [{ wch: 48 }, { wch: 50 }, { wch: 12 }, { wch: 36 }, { wch: 16 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb2, ws1, "KPI");

    // Detail sheet
    const detailHeaders = ["Institution", "Project Title", "PI", "Agency", "Program", "Department", "Category", "Start Date", "Awarded Amount", "Atlantic Canada", "Industry"];
    const detailRows = [detailHeaders, ...included.map(p => [
      p.institution, p.title, p.pi, p.agency, p.program, p.department,
      CATEGORY_LABELS[p.category] || p.category, p.startDate, p.amount,
      p.isAtlantic ? "Yes" : "No", p.isIndustry ? "Yes" : "No"
    ])];
    const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
    ws2["!cols"] = [14, 40, 28, 36, 28, 28, 36, 14, 16, 16, 12].map(wch => ({ wch }));
    XLSX.utils.book_append_sheet(wb2, ws2, "Projects Detail");

    const blob = XLSX.write(wb2, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([blob], { type: "application/octet-stream" }));
    const a = document.createElement("a"); a.href = url; a.download = "Quarterly_Report.xlsx"; a.click();
    showToast("Downloaded Quarterly_Report.xlsx");
  };

  const saveKeywords = () => {
    setKeywords(kwDraft);
    recategorize(kwDraft);
    setEditingKw(false);
    showToast("Keywords saved & projects re-categorized");
  };

  const addNewProject = () => {
    const p = {
      id: Math.random().toString(36).slice(2),
      institution: newProj.institution || "",
      title: newProj.title || "",
      pi: newProj.pi || "",
      agency: newProj.agency || "",
      program: newProj.program || "",
      department: newProj.department || "",
      amount: parseFloat(newProj.amount) || 0,
      startDate: newProj.startDate || "",
      category: detectCategory(newProj.agency || "", keywords),
      isAtlantic: false,
      isIndustry: false,
      include: true,
    };
    setProjects(prev => [...prev, p]);
    setAddRow(false);
    setNewProj({});
    showToast("Project added");
  };

  const s = {
    app: { fontFamily: "var(--font-sans)", color: "var(--color-text-primary)", maxWidth: 1200, margin: "0 auto", padding: "1.5rem 1rem" },
    header: { borderBottom: "0.5px solid var(--color-border-tertiary)", paddingBottom: "1rem", marginBottom: "1.5rem" },
    title: { fontSize: 22, fontWeight: 500, margin: 0 },
    subtitle: { fontSize: 14, color: "var(--color-text-secondary)", margin: "4px 0 0" },
    stepper: { display: "flex", gap: 0, marginBottom: "1.5rem", background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-lg)", padding: 4, width: "fit-content" },
    stepBtn: (active) => ({ border: "none", background: active ? "var(--color-background-primary)" : "transparent", color: active ? "var(--color-text-primary)" : "var(--color-text-secondary)", padding: "6px 18px", borderRadius: "var(--border-radius-md)", fontWeight: active ? 500 : 400, fontSize: 14, cursor: "pointer", transition: "all 0.15s" }),
    card: { background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "1.25rem" },
    btn: { border: "0.5px solid var(--color-border-secondary)", background: "transparent", color: "var(--color-text-primary)", padding: "7px 16px", borderRadius: "var(--border-radius-md)", fontSize: 14, cursor: "pointer" },
    btnPrimary: { border: "none", background: "#185FA5", color: "#fff", padding: "7px 18px", borderRadius: "var(--border-radius-md)", fontSize: 14, cursor: "pointer", fontWeight: 500 },
    badge: (cat) => {
      const colors = { NSERC: "#185FA5", CIHR: "#0F6E56", SSHRC: "#534AB7", MITACS: "#993C1D", IRAP: "#3B6D11", ACOA: "#854F0B", PROVINCIAL: "#993556", INDUSTRY_CONTRACT: "#A32D2D", OTHER_FEDERAL: "#0C447C", CHARITABLE: "#5F5E5A", UNCATEGORIZED: "#444441" };
      return { display: "inline-block", fontSize: 11, padding: "2px 8px", borderRadius: 20, background: (colors[cat] || "#888") + "22", color: colors[cat] || "#888", fontWeight: 500 };
    },
    th: { padding: "8px 10px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textAlign: "left", borderBottom: "0.5px solid var(--color-border-tertiary)", whiteSpace: "nowrap", cursor: "pointer" },
    td: { padding: "7px 10px", fontSize: 13, borderBottom: "0.5px solid var(--color-border-tertiary)", verticalAlign: "top" },
    metricCard: { background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "0.75rem 1rem" },
    metricLabel: { fontSize: 12, color: "var(--color-text-secondary)", marginBottom: 4 },
    metricVal: { fontSize: 22, fontWeight: 500 },
    select: { border: "0.5px solid var(--color-border-secondary)", background: "var(--color-background-primary)", color: "var(--color-text-primary)", padding: "5px 8px", borderRadius: "var(--border-radius-md)", fontSize: 13 },
    input: { border: "0.5px solid var(--color-border-secondary)", background: "var(--color-background-primary)", color: "var(--color-text-primary)", padding: "5px 8px", borderRadius: "var(--border-radius-md)", fontSize: 13 },
  };

  return (
    <div style={s.app}>
      <div style={s.header}>
        <h1 style={s.title}>Quarterly Report Builder</h1>
        <p style={s.subtitle}>Springboard Atlantic — Research Funding Consolidation Tool</p>
      </div>

      {toast && <div style={{ position: "fixed", bottom: 24, right: 24, background: "#185FA5", color: "#fff", padding: "10px 20px", borderRadius: "var(--border-radius-md)", fontSize: 14, zIndex: 9999 }}>{toast}</div>}

      <div style={s.stepper}>
        {STEPS.map((label, i) => (
          <button key={i} style={s.stepBtn(step === i)} onClick={() => projects.length > 0 || i === 0 ? setStep(i) : null}>{i + 1}. {label}</button>
        ))}
      </div>

      {step === 0 && (
        <div style={{ ...s.card, maxWidth: 560 }}>
          <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 0.75rem" }}>Upload funding export</h2>
          <p style={{ fontSize: 14, color: "var(--color-text-secondary)", margin: "0 0 1rem", lineHeight: 1.6 }}>
            Upload the Excel file exported from ROMEO or your institutional research management system.
            Each sheet will be treated as one member institution.
          </p>
          <div style={{ border: "1.5px dashed var(--color-border-secondary)", borderRadius: "var(--border-radius-lg)", padding: "2rem", textAlign: "center", marginBottom: "1rem" }}>
            <div style={{ fontSize: 13, color: "var(--color-text-secondary)", marginBottom: "0.75rem" }}>Supports .xlsx files with one or more sheets</div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            <button style={s.btnPrimary} onClick={() => fileRef.current.click()}>Choose file</button>
          </div>
          <div style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>
            Columns auto-detected: Project Title, Agency, PI, Amount, Department, Start Date
          </div>
        </div>
      )}

      {step === 1 && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10, marginBottom: "1rem" }}>
            {[
              { label: "Total projects", val: projects.length },
              { label: "Included", val: included.length },
              { label: "Institutions", val: new Set(projects.map(p => p.institution)).size },
              { label: "Total awarded", val: "$" + included.reduce((s, p) => s + p.amount, 0).toLocaleString() },
            ].map(m => (
              <div key={m.label} style={s.metricCard}>
                <div style={s.metricLabel}>{m.label}</div>
                <div style={s.metricVal}>{m.val}</div>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", gap: 8, marginBottom: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
            <select style={s.select} value={filterInstitution} onChange={e => setFilterInstitution(e.target.value)}>
              {institutions.map(i => <option key={i}>{i}</option>)}
            </select>
            <select style={s.select} value={filterCategory} onChange={e => setFilterCategory(e.target.value)}>
              {categories.map(c => <option key={c}>{c}</option>)}
            </select>
            <div style={{ flex: 1 }} />
            <button style={s.btn} onClick={() => { setAddRow(true); setNewProj({}); }}>+ Add project</button>
            <button style={s.btn} onClick={() => { setKwDraft(JSON.parse(JSON.stringify(keywords))); setEditingKw(true); }}>Edit category keywords</button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
              <thead>
                <tr style={{ background: "var(--color-background-secondary)" }}>
                  {[["include","✓",36],["institution","Institution",110],["title","Project Title",220],["pi","PI",120],["agency","Agency",150],["amount","Amount",90],["category","Category",130],["isAtlantic","Atlantic",76],["isIndustry","Industry",76]].map(([col, label, w]) => (
                    <th key={col} style={{ ...s.th, width: w }} onClick={() => toggleSort(col)}>
                      {label} {sortCol === col ? (sortDir === "asc" ? "↑" : "↓") : ""}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {addRow && (
                  <tr style={{ background: "var(--color-background-info)" }}>
                    <td style={s.td}><input type="checkbox" checked={true} readOnly /></td>
                    {["institution","title","pi","agency","amount"].map(f => (
                      <td key={f} style={s.td} colSpan={f === "agency" ? 1 : 1}>
                        <input style={{ ...s.input, width: "100%" }} placeholder={f} value={newProj[f] || ""} onChange={e => setNewProj(p => ({ ...p, [f]: e.target.value }))} />
                      </td>
                    ))}
                    <td style={s.td}><span style={s.badge(detectCategory(newProj.agency || "", keywords))}>{detectCategory(newProj.agency || "", keywords)}</span></td>
                    <td style={s.td}><input type="checkbox" onChange={e => setNewProj(p => ({ ...p, isAtlantic: e.target.checked }))} /></td>
                    <td style={s.td}>
                      <button style={{ ...s.btnPrimary, fontSize: 12, padding: "4px 10px", marginRight: 6 }} onClick={addNewProject}>Save</button>
                      <button style={{ ...s.btn, fontSize: 12, padding: "4px 8px" }} onClick={() => setAddRow(false)}>Cancel</button>
                    </td>
                  </tr>
                )}
                {filtered.map(p => (
                  <tr key={p.id} style={{ opacity: p.include ? 1 : 0.45 }}>
                    <td style={s.td}><input type="checkbox" checked={p.include} onChange={e => updateProject(p.id, "include", e.target.checked)} /></td>
                    <td style={{ ...s.td, fontSize: 12, color: "var(--color-text-secondary)" }}>{p.institution}</td>
                    <td style={s.td}>{p.title}</td>
                    <td style={{ ...s.td, fontSize: 12 }}>{p.pi}</td>
                    <td style={{ ...s.td, fontSize: 12, color: "var(--color-text-secondary)" }}>{p.agency}</td>
                    <td style={{ ...s.td, textAlign: "right" }}>${p.amount.toLocaleString()}</td>
                    <td style={s.td}>
                      <select style={{ ...s.select, fontSize: 11, padding: "2px 6px" }} value={p.category} onChange={e => updateProject(p.id, "category", e.target.value)}>
                        {Object.keys(CATEGORY_LABELS).map(k => <option key={k} value={k}>{k}</option>)}
                      </select>
                    </td>
                    <td style={{ ...s.td, textAlign: "center" }}><input type="checkbox" checked={p.isAtlantic} onChange={e => updateProject(p.id, "isAtlantic", e.target.checked)} /></td>
                    <td style={{ ...s.td, textAlign: "center" }}><input type="checkbox" checked={p.isIndustry} onChange={e => updateProject(p.id, "isIndustry", e.target.checked)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: "1rem", display: "flex", gap: 8 }}>
            <button style={s.btnPrimary} onClick={() => setStep(2)}>Review summary →</button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div>
          <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 1rem" }}>Category summary</h2>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "var(--color-background-secondary)" }}>
                {["Category", "Projects #", "Atlantic #", "Total $ Awarded", "Atlantic $ Awarded"].map(h => (
                  <th key={h} style={s.th}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Object.entries(CATEGORY_LABELS).map(([cat, label]) => {
                const d = summaryByCat[cat] || { count: 0, total: 0, atlanticCount: 0, atlanticTotal: 0 };
                return (
                  <tr key={cat}>
                    <td style={s.td}><span style={s.badge(cat)}>{label}</span></td>
                    <td style={{ ...s.td, textAlign: "right" }}>{d.count || "–"}</td>
                    <td style={{ ...s.td, textAlign: "right" }}>{d.atlanticCount || "–"}</td>
                    <td style={{ ...s.td, textAlign: "right" }}>{d.total ? "$" + d.total.toLocaleString() : "–"}</td>
                    <td style={{ ...s.td, textAlign: "right" }}>{d.atlanticTotal ? "$" + d.atlanticTotal.toLocaleString() : "–"}</td>
                  </tr>
                );
              })}
              <tr style={{ fontWeight: 500, background: "var(--color-background-secondary)" }}>
                <td style={s.td}>TOTAL</td>
                <td style={{ ...s.td, textAlign: "right" }}>{included.length}</td>
                <td style={{ ...s.td, textAlign: "right" }}>{included.filter(p => p.isAtlantic).length}</td>
                <td style={{ ...s.td, textAlign: "right" }}>${included.reduce((s, p) => s + p.amount, 0).toLocaleString()}</td>
                <td style={{ ...s.td, textAlign: "right" }}>${included.filter(p => p.isAtlantic).reduce((s, p) => s + p.amount, 0).toLocaleString()}</td>
              </tr>
            </tbody>
          </table>
          <div style={{ marginTop: "1rem", display: "flex", gap: 8 }}>
            <button style={s.btn} onClick={() => setStep(1)}>← Back to projects</button>
            <button style={s.btnPrimary} onClick={() => setStep(3)}>Proceed to download →</button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div style={{ ...s.card, maxWidth: 480 }}>
          <h2 style={{ fontSize: 16, fontWeight: 500, margin: "0 0 0.75rem" }}>Download consolidated report</h2>
          <p style={{ fontSize: 14, color: "var(--color-text-secondary)", lineHeight: 1.6, margin: "0 0 1rem" }}>
            The Excel file will contain two sheets: <strong>KPI</strong> (reporting template with category totals) and <strong>Projects Detail</strong> (full list of included projects).
          </p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: "1.25rem" }}>
            {[
              { label: "Included projects", val: included.length },
              { label: "Total awarded", val: "$" + included.reduce((s, p) => s + p.amount, 0).toLocaleString() },
              { label: "Atlantic Canada", val: included.filter(p => p.isAtlantic).length },
              { label: "Industry", val: included.filter(p => p.isIndustry).length },
            ].map(m => (
              <div key={m.label} style={s.metricCard}>
                <div style={s.metricLabel}>{m.label}</div>
                <div style={{ fontSize: 18, fontWeight: 500 }}>{m.val}</div>
              </div>
            ))}
          </div>
          <button style={{ ...s.btnPrimary, width: "100%", padding: "10px" }} onClick={generateExcel}>Download Quarterly_Report.xlsx</button>
        </div>
      )}

      {editingKw && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, overflowY: "auto", paddingTop: 40 }}>
          <div style={{ ...s.card, width: "min(700px, 96vw)", maxHeight: "80vh", overflowY: "auto", margin: "0 auto 40px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
              <h2 style={{ fontSize: 16, fontWeight: 500, margin: 0 }}>Category keywords</h2>
              <button style={s.btn} onClick={() => setEditingKw(false)}>Close</button>
            </div>
            <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: "0 0 1rem" }}>
              If a funding agency contains any of these keywords (case-insensitive), the project is assigned to that category. Changes are saved locally for future sessions.
            </p>
            {Object.entries(kwDraft).map(([cat, words]) => (
              <div key={cat} style={{ marginBottom: "1rem", paddingBottom: "1rem", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                  <span style={s.badge(cat)}>{CATEGORY_LABELS[cat] || cat}</span>
                  <button style={{ ...s.btn, fontSize: 11, padding: "2px 8px", color: "var(--color-text-danger)" }} onClick={() => {
                    const d = { ...kwDraft }; delete d[cat]; setKwDraft(d);
                  }}>Remove category</button>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {words.map((w, i) => (
                    <span key={i} style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "var(--color-background-secondary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 20, padding: "2px 10px", fontSize: 12 }}>
                      {w}
                      <span style={{ cursor: "pointer", color: "var(--color-text-secondary)" }} onClick={() => setKwDraft(d => ({ ...d, [cat]: d[cat].filter((_, j) => j !== i) }))}>×</span>
                    </span>
                  ))}
                  <input
                    placeholder="+ add keyword"
                    style={{ ...s.input, fontSize: 12, padding: "2px 8px", minWidth: 120 }}
                    onKeyDown={e => {
                      if (e.key === "Enter" && e.target.value.trim()) {
                        setKwDraft(d => ({ ...d, [cat]: [...d[cat], e.target.value.trim()] }));
                        e.target.value = "";
                      }
                    }}
                  />
                </div>
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginBottom: "1rem" }}>
              <input style={{ ...s.input, flex: 1 }} placeholder="New category name (e.g. CFI)" value={newCatName} onChange={e => setNewCatName(e.target.value)} />
              <button style={s.btn} onClick={() => {
                if (newCatName.trim() && !kwDraft[newCatName.trim()]) {
                  setKwDraft(d => ({ ...d, [newCatName.trim()]: [] }));
                  setNewCatName("");
                }
              }}>Add category</button>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={s.btnPrimary} onClick={saveKeywords}>Save & re-categorize</button>
              <button style={s.btn} onClick={() => { setKwDraft(JSON.parse(JSON.stringify(DEFAULT_KEYWORDS))); }}>Reset to defaults</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
