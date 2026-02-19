import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Palette & constants ───────────────────────────────────────────────────
// All employees who require time tracking, grouped by team
const TEAM_DEFAULT = [
  // Design
  { name: "James Owens",           team: "Design",      rate: 50 },
  { name: "Michael Hall",          team: "Design",      rate: 50 },
  { name: "Ollie Allington",       team: "Design",      rate: 50 },
  // Engineering
  { name: "Andy Timms",            team: "Engineering", rate: 50 },
  { name: "Cameron Birtwhistle",   team: "Engineering", rate: 50 },
  { name: "David Gurjao",          team: "Engineering", rate: 50 },
  { name: "Dermot McGee",          team: "Engineering", rate: 50 },
  { name: "Ionut Dan",             team: "Engineering", rate: 50 },
  { name: "James Harris",          team: "Engineering", rate: 50 },
  { name: "Leonidas Likskendaj",   team: "Engineering", rate: 50 },
  { name: "Maciej Dulkowski",      team: "Engineering", rate: 50 },
  { name: "Reiss Lewis",           team: "Engineering", rate: 50 },
  // Presales
  { name: "Ando Goddess",          team: "Presales",    rate: 50 },
  { name: "Matt Horsfall",         team: "Presales",    rate: 50 },
  // Projects
  { name: "Arron-James Pienaar",   team: "Projects",    rate: 50 },
  { name: "Rob Elkins",            team: "Projects",    rate: 50 },
  { name: "Rupert Merryweather",   team: "Projects",    rate: 50 },
  // Service
  { name: "Darren Williams",       team: "Service",     rate: 50 },
  { name: "Sodrul Islam",          team: "Service",     rate: 50 },
];

const MONTHS = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"];

// ─── Number formatting helpers ────────────────────────────────────────────
function fmtGBP(n)  { return `£${Math.round(n).toLocaleString("en-GB")}` }
function fmtPct(n)  { return `${(n * 100).toFixed(1)}%` }
function fmtHrs(n)  { return +n.toFixed(1) }
function fmtVar(n)  { return n >= 0 ? `+${fmtGBP(n)}` : `-£${Math.abs(Math.round(n)).toLocaleString("en-GB")}` }
function fmtHrsVar(n){ return n >= 0 ? `+${n.toFixed(1)}` : n.toFixed(1) }

// ─── Excel generation ──────────────────────────────────────────────────────
function buildExcel({ timesheetRows, taskRows, team, month, year, workingDays }) {
  const wb = XLSX.utils.book_new();

  const rateMap = {};
  team.forEach(t => { rateMap[t.name] = parseFloat(t.rate) || 0; });
  const avgRate = team.reduce((s,t) => s + (parseFloat(t.rate)||0), 0) / team.length;

  // ── Build project map from task-list (budgets only) ─────────────────────
  // Use task-list for BUDGET hours (duration field)
  // Use TIMESHEET for ACTUAL hours — single source of truth
  const projBudget = {};
  const projCustomer = {};
  taskRows.forEach(r => {
    const proj = (r.project || "").trim();
    if (!proj) return;
    const dur = parseFloat(r.duration) || 0;
    // FIX: only count duration > 0 toward budget (ignores placeholder 0-duration tasks)
    if (dur > 0) {
      projBudget[proj]   = (projBudget[proj] || 0) + dur;
    }
    if (!projCustomer[proj]) projCustomer[proj] = r.customer_name || "";
  });

  // ── FIX 3: Aggregate ACTUAL hours from timesheet per project ─────────────
  const projActual = {};
  const personHrs  = {};
  const personProjHrs = {}; // for per-person-per-project breakdown
  timesheetRows.forEach(r => {
    const proj = (r.project_description || "").trim();
    const name = (r.user_name || "").trim();
    const hrs  = parseFloat(r.hours) || 0;
    if (proj) projActual[proj]  = (projActual[proj] || 0) + hrs;
    if (name) personHrs[name]   = (personHrs[name] || 0) + hrs;
    if (proj && name) {
      if (!personProjHrs[proj]) personProjHrs[proj] = {};
      personProjHrs[proj][name] = (personProjHrs[proj][name] || 0) + hrs;
    }
    // Capture customer from timesheet for projects missing from task-list
    if (proj && !projCustomer[proj]) projCustomer[proj] = "";
  });

  // ── Build unified project list ───────────────────────────────────────────
  const allProjects = [...new Set([
    ...Object.keys(projBudget),
    ...Object.keys(projActual),
  ])];

  const projMap = {};
  allProjects.forEach(proj => {
    const budget = projBudget[proj] || 0;
    projMap[proj] = {
      customer:  projCustomer[proj] || "",
      budget,
      worked:    projActual[proj] || 0,
      hasQuote:  budget > 0,
    };
  });

  const budgeted = Object.entries(projMap).filter(([,v]) => v.hasQuote);
  const noBudget = Object.entries(projMap).filter(([,v]) => !v.hasQuote);

  // ── Totals ───────────────────────────────────────────────────────────────
  const totalBudgetHrs  = budgeted.reduce((s,[,v]) => s + v.budget, 0);
  const totalActualHrs  = budgeted.reduce((s,[,v]) => s + v.worked, 0);
  const totalBudgetGBP  = totalBudgetHrs * avgRate;
  const totalActualGBP  = totalActualHrs * avgRate;
  const totalVariance   = totalBudgetGBP - totalActualGBP;
  const overallMarginPc = totalBudgetGBP > 0 ? totalVariance / totalBudgetGBP : 0;
  const availHrs        = workingDays * 8;
  const totalCapacity   = team.reduce((s,t) => s + availHrs * (parseFloat(t.rate)||0), 0);
  const totalDeployed   = Object.entries(personHrs).reduce((s,[n,h]) => s + h*(rateMap[n]||avgRate), 0);
  const totalUndeployed = totalCapacity - totalDeployed;
  const totalWorkedHrs  = Object.values(personHrs).reduce((s,h) => s+h, 0);
  const totalAvailHrs   = team.length * availHrs;
  const teamUtil        = totalAvailHrs > 0 ? totalWorkedHrs / totalAvailHrs : 0;

  // ── SHEET 1: Dashboard ───────────────────────────────────────────────────
  const dashData = [
    [`${month} ${year} — Monthly Labour Dashboard`],
    [],
    ["KPI", "Value", "Formatted"],
    ["Total Budget Hours",       fmtHrs(totalBudgetHrs),   `${fmtHrs(totalBudgetHrs)} hrs`],
    ["Total Actual Hours",       fmtHrs(totalActualHrs),   `${fmtHrs(totalActualHrs)} hrs`],
    ["Hours Variance",           fmtHrs(totalActualHrs - totalBudgetHrs), fmtHrsVar(totalActualHrs - totalBudgetHrs) + " hrs"],
    ["Labour Budget (£)",        +totalBudgetGBP.toFixed(0),  fmtGBP(totalBudgetGBP)],
    ["Actual Labour Cost (£)",   +totalActualGBP.toFixed(0),  fmtGBP(totalActualGBP)],
    ["Labour Variance (£)",      +totalVariance.toFixed(0),   fmtVar(totalVariance)],
    ["Overall Margin %",         +overallMarginPc.toFixed(4), fmtPct(overallMarginPc)],
    ["Team Utilisation %",       +teamUtil.toFixed(4),        fmtPct(teamUtil)],
    ["Total Capacity Cost (£)",  +totalCapacity.toFixed(0),   fmtGBP(totalCapacity)],
    ["Undeployed Cost (£)",      +totalUndeployed.toFixed(0), fmtGBP(totalUndeployed)],
    [],
    ["PROJECT STATUS", "", "", "", "", "", "", "", "", ""],
    ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Variance","Burn %","Budget £","Actual £","Variance £","Margin %","Status"],
    ...budgeted.map(([proj, v]) => {
      const burn   = v.budget > 0 ? v.worked / v.budget : 0;
      const bGBP   = v.budget * avgRate;
      const aGBP   = v.worked * avgRate;
      const varGBP = bGBP - aGBP;
      const margin = bGBP > 0 ? varGBP / bGBP : 0;
      const status = burn <= 0.95 ? "On Budget" : burn <= 1.10 ? "At Risk" : "Over Budget";
      return [proj, v.customer,
        fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
        fmtPct(burn), fmtGBP(bGBP), fmtGBP(aGBP), fmtVar(varGBP),
        fmtPct(margin), status];
    }),
    [],
    ["PROJECTS WITHOUT QUOTED BUDGET — excluded from margin calculations"],
    ["Project","Customer","Actual Hrs"],
    ...noBudget.map(([proj,v]) => [proj, v.customer, fmtHrs(v.worked)]),
  ];

  const wsDash = XLSX.utils.aoa_to_sheet(dashData);
  wsDash["!cols"] = [{ wch:44 },{ wch:28 },{ wch:12 },{ wch:12 },{ wch:13 },
                     { wch:10 },{ wch:13 },{ wch:13 },{ wch:13 },{ wch:11 },{ wch:13 }];
  XLSX.utils.book_append_sheet(wb, wsDash, "Dashboard");

  // ── SHEET 2: Project Profitability ───────────────────────────────────────
  const projRows = [
    [`PROJECT PROFITABILITY — ${month} ${year}`],
    [`Avg blended rate: £${avgRate.toFixed(2)}/hr  |  Actual hours sourced from timesheet export`],
    [],
    ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Variance",
     "Burn %","Budget £","Actual Cost £","Variance £","Margin £","Margin %","Status"],
    ...budgeted.map(([proj,v]) => {
      const burn   = v.budget > 0 ? v.worked / v.budget : 0;
      const bGBP   = v.budget * avgRate;
      const aGBP   = v.worked * avgRate;
      const varGBP = bGBP - aGBP;
      const margin = bGBP > 0 ? varGBP / bGBP : 0;
      const status = burn <= 0.95 ? "On Budget" : burn <= 1.10 ? "At Risk" : "Over Budget";
      return [
        proj, v.customer,
        fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
        fmtPct(burn), fmtGBP(bGBP), fmtGBP(aGBP), fmtVar(varGBP),
        fmtVar(varGBP), fmtPct(margin), status,
      ];
    }),
    ["TOTAL","",
      fmtHrs(totalBudgetHrs), fmtHrs(totalActualHrs),
      fmtHrsVar(totalActualHrs - totalBudgetHrs),
      fmtPct(totalBudgetHrs > 0 ? totalActualHrs / totalBudgetHrs : 0),
      fmtGBP(totalBudgetGBP), fmtGBP(totalActualGBP),
      fmtVar(totalVariance), fmtVar(totalVariance),
      fmtPct(overallMarginPc), "",
    ],
    [],
    ["NO QUOTED BUDGET — hours logged, excluded from margin"],
    ["Project","Customer","Actual Hrs"],
    ...noBudget.map(([proj,v]) => [proj, v.customer, fmtHrs(v.worked)]),
  ];

  const wsProj = XLSX.utils.aoa_to_sheet(projRows);
  wsProj["!cols"] = [{ wch:44},{wch:28},{wch:12},{wch:12},{wch:13},
                     {wch:10},{wch:13},{wch:14},{wch:13},{wch:13},{wch:11},{wch:13}];
  XLSX.utils.book_append_sheet(wb, wsProj, "Project Profitability");

  // ── SHEET 3: Utilisation ─────────────────────────────────────────────────
  const teamGroups = [...new Set(team.map(t => t.team))];
  const utilRows = [
    [`LABOUR UTILISATION — ${month} ${year}`],
    [`Available hours: ${workingDays} working days × 8 hrs = ${availHrs} hrs/person`],
    [],
    ["Team Member","Team","Hourly Rate (£)","Available Hrs","Hours Worked",
     "Utilisation %","Capacity Cost (£)","Labour Cost (£)","Undeployed Cost (£)"],
    ...teamGroups.flatMap(grp => {
      const members = team.filter(t => t.team === grp);
      const memberRows = members.map(t => {
        const rate    = parseFloat(t.rate) || 0;
        const worked  = personHrs[t.name] || 0;
        const util    = availHrs > 0 ? worked / availHrs : 0;
        const capCost = availHrs * rate;
        const labCost = worked * rate;
        return [
          t.name, t.team || "", fmtGBP(rate), availHrs,
          fmtHrs(worked), fmtPct(util),
          fmtGBP(capCost), fmtGBP(labCost), fmtGBP(capCost - labCost),
        ];
      });
      const grpWorked   = members.reduce((s,t) => s + (personHrs[t.name]||0), 0);
      const grpCap      = members.reduce((s,t) => s + availHrs*(parseFloat(t.rate)||0), 0);
      const grpDeployed = members.reduce((s,t) => s + (personHrs[t.name]||0)*(parseFloat(t.rate)||0), 0);
      const grpAvail    = members.length * availHrs;
      return [
        ...memberRows,
        [`— ${grp} Subtotal`, "", "", grpAvail, fmtHrs(grpWorked),
         fmtPct(grpAvail > 0 ? grpWorked / grpAvail : 0),
         fmtGBP(grpCap), fmtGBP(grpDeployed), fmtGBP(grpCap - grpDeployed)],
        [],
      ];
    }),
    ["TOTAL","","", totalAvailHrs, fmtHrs(totalWorkedHrs),
      fmtPct(teamUtil), fmtGBP(totalCapacity), fmtGBP(totalDeployed), fmtGBP(totalUndeployed)],
  ];

  const wsUtil = XLSX.utils.aoa_to_sheet(utilRows);
  wsUtil["!cols"] = [{wch:26},{wch:14},{wch:15},{wch:14},{wch:14},
                     {wch:14},{wch:20},{wch:20},{wch:20}];
  XLSX.utils.book_append_sheet(wb, wsUtil, "Utilisation");

  // ── SHEET 4: Raw Timesheet ───────────────────────────────────────────────
  const tsHeaders = ["Job No.","Task","Project","Team Member","Date","Hours","Billable","Notes"];
  const tsData = timesheetRows.map(r => [
    r.work_order_no || "",
    r.task_name || "",
    (r.project_description || "").trim(),
    r.user_name || "",
    (r.start_datetime || "").slice(0,10),
    parseFloat(r.hours) || 0,
    r.is_billable === "1" ? "Yes" : "No",
    r.notes || "",
  ]);
  const wsTs = XLSX.utils.aoa_to_sheet([tsHeaders, ...tsData]);
  wsTs["!cols"] = [{wch:12},{wch:30},{wch:42},{wch:22},{wch:12},{wch:8},{wch:9},{wch:30}];
  XLSX.utils.book_append_sheet(wb, wsTs, "Raw Timesheet");

  // ── SHEET 5: Settings snapshot ───────────────────────────────────────────
  const settingsData = [
    ["REPORT SETTINGS SNAPSHOT"],
    ["Report Month",         month],
    ["Report Year",          year],
    ["Working Days",         workingDays],
    ["Hrs per Day",          8],
    ["Available Hrs/Person", availHrs],
    ["Avg Hourly Rate",      fmtGBP(avgRate) + "/hr"],
    [],
    ["TEAM RATES"],
    ["Name","Team","Hourly Rate (£)"],
    ...team.map(t => [t.name, t.team || "", fmtGBP(parseFloat(t.rate)||0) + "/hr"]),
  ];
  const wsSet = XLSX.utils.aoa_to_sheet(settingsData);
  wsSet["!cols"] = [{wch:26},{wch:14},{wch:22}];
  XLSX.utils.book_append_sheet(wb, wsSet, "Settings");

  return wb;
}

// ─── File drop zone component ──────────────────────────────────────────────
function DropZone({ label, sublabel, accept, onFile, file, icon }) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef();

  const handleDrop = useCallback(e => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  }, [onFile]);

  return (
    <div
      onClick={() => inputRef.current.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      style={{
        border: `2px dashed ${file ? "#10b981" : dragging ? "#6366f1" : "#334155"}`,
        borderRadius: 12,
        padding: "28px 20px",
        cursor: "pointer",
        background: file ? "rgba(16,185,129,0.06)" : dragging ? "rgba(99,102,241,0.08)" : "rgba(15,23,42,0.4)",
        transition: "all 0.2s ease",
        textAlign: "center",
        position: "relative",
      }}
    >
      <input ref={inputRef} type="file" accept={accept} style={{ display:"none" }}
        onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{ fontSize: 32, marginBottom: 8 }}>{file ? "✅" : icon}</div>
      <div style={{ fontFamily:"'DM Serif Display',serif", fontSize:15, color: file ? "#10b981" : "#e2e8f0", marginBottom:4 }}>
        {file ? file.name : label}
      </div>
      <div style={{ fontSize:12, color:"#64748b", fontFamily:"'DM Mono',monospace" }}>
        {file ? `${(file.size/1024).toFixed(1)} KB` : sublabel}
      </div>
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────
export default function App() {
  const [timesheetFile, setTimesheetFile] = useState(null);
  const [taskFile,      setTaskFile]      = useState(null);
  const [team,          setTeam]          = useState(TEAM_DEFAULT);
  const [month,         setMonth]         = useState("February");
  const [year,          setYear]          = useState(2026);
  const [workingDays,   setWorkingDays]   = useState(20);
  const [status,        setStatus]        = useState("idle");
  const [errorMsg,      setErrorMsg]      = useState("");
  const [preview,       setPreview]       = useState(null);

  const parseCSV = (text) => {
    const lines = text.trim().split("\n");
    const headers = lines[0].split(",").map(h => h.replace(/"/g,"").trim());
    return lines.slice(1).map(line => {
      const vals = [];
      let cur = "", inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') { inQ = !inQ; continue; }
        if (ch === "," && !inQ) { vals.push(cur.trim()); cur = ""; continue; }
        cur += ch;
      }
      vals.push(cur.trim());
      const obj = {};
      headers.forEach((h,i) => { obj[h] = vals[i] || ""; });
      return obj;
    });
  };

  const readFile = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => res(e.target.result);
    reader.onerror = rej;
    reader.readAsText(file);
  });

  const handleGenerate = async () => {
    if (!timesheetFile || !taskFile) return;
    setStatus("parsing");
    setErrorMsg("");
    try {
      const [tsText, taskText] = await Promise.all([
        readFile(timesheetFile),
        readFile(taskFile),
      ]);
      const timesheetRows = parseCSV(tsText);
      const taskRows      = parseCSV(taskText);

      setStatus("generating");
      // Build preview summary
      const projCount = new Set(taskRows.map(r => (r.project||"").trim()).filter(Boolean)).size;
      const peopleCount = new Set(timesheetRows.map(r => r.user_name).filter(Boolean)).size;
      const totalHrs = timesheetRows.reduce((s,r) => s + (parseFloat(r.hours)||0), 0);
      setPreview({ projCount, peopleCount, totalHrs: totalHrs.toFixed(1) });

      const wb = buildExcel({ timesheetRows, taskRows, team, month, year, workingDays });

      const filename = `Labour_Report_${month}_${year}.xlsx`;
      XLSX.writeFile(wb, filename);

      setStatus("done");
    } catch (e) {
      setStatus("error");
      setErrorMsg(e.message);
    }
  };

  const updateRate = (i, val) => {
    setTeam(prev => prev.map((t,idx) => idx===i ? {...t, rate: val} : t));
  };

  const ready = timesheetFile && taskFile;

  return (
    <div style={{
      minHeight: "100vh",
      background: "#020818",
      backgroundImage: "radial-gradient(ellipse 80% 50% at 50% -20%, rgba(99,102,241,0.15), transparent)",
      fontFamily: "'DM Sans', sans-serif",
      color: "#e2e8f0",
      padding: "0 0 60px",
    }}>
      {/* Google Fonts */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

        * { box-sizing: border-box; margin: 0; padding: 0; }

        input[type=number]::-webkit-inner-spin-button { opacity: 1; }

        .rate-input {
          background: rgba(255,255,255,0.05);
          border: 1px solid #1e293b;
          border-radius: 8px;
          color: #93c5fd;
          font-family: 'DM Mono', monospace;
          font-size: 14px;
          padding: 7px 10px;
          width: 90px;
          outline: none;
          transition: border-color 0.15s;
        }
        .rate-input:focus { border-color: #6366f1; }

        select.field {
          background: rgba(255,255,255,0.05);
          border: 1px solid #1e293b;
          border-radius: 8px;
          color: #e2e8f0;
          font-family: 'DM Sans', sans-serif;
          font-size: 14px;
          padding: 9px 12px;
          outline: none;
          cursor: pointer;
          transition: border-color 0.15s;
        }
        select.field:focus { border-color: #6366f1; }
        select.field option { background: #0f172a; }

        input.field {
          background: rgba(255,255,255,0.05);
          border: 1px solid #1e293b;
          border-radius: 8px;
          color: #e2e8f0;
          font-family: 'DM Mono', monospace;
          font-size: 14px;
          padding: 9px 12px;
          outline: none;
          width: 100px;
          transition: border-color 0.15s;
        }
        input.field:focus { border-color: #6366f1; }

        .generate-btn {
          background: linear-gradient(135deg, #6366f1, #4f46e5);
          border: none;
          border-radius: 12px;
          color: white;
          cursor: pointer;
          font-family: 'DM Sans', sans-serif;
          font-size: 16px;
          font-weight: 600;
          letter-spacing: 0.3px;
          padding: 16px 40px;
          transition: all 0.2s;
          box-shadow: 0 4px 24px rgba(99,102,241,0.35);
        }
        .generate-btn:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 8px 32px rgba(99,102,241,0.5);
        }
        .generate-btn:disabled {
          opacity: 0.4;
          cursor: not-allowed;
          transform: none;
        }

        .card {
          background: rgba(15,23,42,0.7);
          border: 1px solid #1e293b;
          border-radius: 16px;
          padding: 28px;
          backdrop-filter: blur(12px);
        }

        .section-label {
          font-family: 'DM Mono', monospace;
          font-size: 11px;
          letter-spacing: 2px;
          text-transform: uppercase;
          color: #6366f1;
          margin-bottom: 16px;
        }

        .pill {
          display: inline-block;
          background: rgba(99,102,241,0.12);
          border: 1px solid rgba(99,102,241,0.25);
          border-radius: 20px;
          color: #a5b4fc;
          font-family: 'DM Mono', monospace;
          font-size: 11px;
          padding: 3px 10px;
          margin-right: 8px;
        }

        .stat-chip {
          background: rgba(255,255,255,0.04);
          border: 1px solid #1e293b;
          border-radius: 10px;
          padding: 12px 16px;
          flex: 1;
          text-align: center;
        }
      `}</style>

      {/* Header */}
      <div style={{ textAlign:"center", padding:"56px 24px 40px" }}>
        <div style={{ display:"inline-flex", alignItems:"center", gap:10,
                      background:"rgba(99,102,241,0.1)", border:"1px solid rgba(99,102,241,0.2)",
                      borderRadius:30, padding:"6px 18px", marginBottom:24 }}>
          <span style={{ width:8, height:8, borderRadius:"50%", background:"#10b981", display:"inline-block" }}/>
          <span style={{ fontFamily:"'DM Mono',monospace", fontSize:11, color:"#a5b4fc", letterSpacing:2 }}>
            LABOUR REPORTING TOOL
          </span>
        </div>
        <h1 style={{
          fontFamily: "'DM Serif Display', serif",
          fontSize: "clamp(32px, 5vw, 52px)",
          fontWeight: 400,
          lineHeight: 1.15,
          background: "linear-gradient(135deg, #e2e8f0 30%, #a5b4fc)",
          WebkitBackgroundClip: "text",
          WebkitTextFillColor: "transparent",
          marginBottom: 16,
        }}>
          Monthly Profitability<br />&amp; Utilisation Report
        </h1>
        <p style={{ color:"#64748b", fontSize:15, maxWidth:480, margin:"0 auto" }}>
          Upload your WeQuote exports, confirm settings, and download a fully formatted Excel report in seconds.
        </p>
      </div>

      <div style={{ maxWidth: 860, margin:"0 auto", padding:"0 20px", display:"flex", flexDirection:"column", gap:20 }}>

        {/* Step 1 – Upload */}
        <div className="card">
          <div className="section-label">Step 1 — Upload CSV Exports</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
            <DropZone
              label="Timesheet Entries"
              sublabel="Timesheet_Entries_YYYYMMDD.csv"
              accept=".csv"
              onFile={setTimesheetFile}
              file={timesheetFile}
              icon="⏱"
            />
            <DropZone
              label="Task List"
              sublabel="task-list.csv"
              accept=".csv"
              onFile={setTaskFile}
              file={taskFile}
              icon="📋"
            />
          </div>
        </div>

        {/* Step 2 – Period */}
        <div className="card">
          <div className="section-label">Step 2 — Report Period</div>
          <div style={{ display:"flex", gap:16, flexWrap:"wrap", alignItems:"center" }}>
            <div>
              <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>MONTH</div>
              <select className="field" value={month} onChange={e => setMonth(e.target.value)}>
                {MONTHS.map(m => <option key={m}>{m}</option>)}
              </select>
            </div>
            <div>
              <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>YEAR</div>
              <input className="field" type="number" value={year}
                onChange={e => setYear(parseInt(e.target.value)||2026)} min={2020} max={2035} />
            </div>
            <div>
              <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>WORKING DAYS</div>
              <input className="field" type="number" value={workingDays}
                onChange={e => setWorkingDays(parseInt(e.target.value)||20)} min={1} max={31} />
            </div>
            <div style={{ flex:1, minWidth:200, background:"rgba(99,102,241,0.06)",
                          border:"1px solid rgba(99,102,241,0.15)", borderRadius:10,
                          padding:"12px 16px" }}>
              <div style={{ fontSize:11, color:"#6366f1", fontFamily:"'DM Mono',monospace", marginBottom:4 }}>
                AVAILABLE HRS / PERSON
              </div>
              <div style={{ fontSize:22, fontFamily:"'DM Serif Display',serif", color:"#a5b4fc" }}>
                {workingDays * 8}
                <span style={{ fontSize:13, color:"#64748b", marginLeft:6 }}>hrs</span>
              </div>
            </div>
          </div>
        </div>

        {/* Step 3 – Team rates */}
        <div className="card">
          <div className="section-label">Step 3 — Team Hourly Cost Rates (£)</div>
          <div style={{ fontSize:12, color:"#475569", marginBottom:16, fontFamily:"'DM Mono',monospace" }}>
            19 time-tracked employees · rates default to £50/hr · update to actual cost rates before generating
          </div>
          {/* Group by team */}
          {["Design","Engineering","Presales","Projects","Service"].map(teamName => {
            const members = team.map((t,i) => ({...t, i})).filter(t => t.team === teamName);
            return (
              <div key={teamName} style={{ marginBottom:20 }}>
                <div style={{
                  fontFamily:"'DM Mono',monospace", fontSize:10, letterSpacing:2,
                  textTransform:"uppercase", color:"#6366f1", marginBottom:10,
                  paddingBottom:6, borderBottom:"1px solid #1e293b",
                }}>
                  {teamName}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(260px, 1fr))", gap:10 }}>
                  {members.map(({ name, rate, i }) => (
                    <div key={name} style={{
                      display:"flex", alignItems:"center", justifyContent:"space-between",
                      background:"rgba(255,255,255,0.03)", border:"1px solid #1e293b",
                      borderRadius:10, padding:"10px 14px",
                    }}>
                      <div style={{ fontSize:13, fontWeight:500, color:"#e2e8f0" }}>{name}</div>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <span style={{ color:"#64748b", fontSize:14 }}>£</span>
                        <input
                          className="rate-input"
                          type="number"
                          value={rate}
                          onChange={e => updateRate(i, e.target.value)}
                          min={0} max={999}
                        />
                        <span style={{ color:"#475569", fontSize:11, fontFamily:"'DM Mono',monospace" }}>/hr</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
          <div style={{ marginTop:4, fontSize:12, color:"#475569", fontFamily:"'DM Mono',monospace" }}>
            Avg blended rate: £{(team.reduce((s,t)=>s+(parseFloat(t.rate)||0),0)/team.length).toFixed(2)}/hr
            &nbsp;·&nbsp; Used for project margin calculations
          </div>
        </div>

        {/* Generate */}
        <div style={{ textAlign:"center", padding:"8px 0 4px" }}>
          {status === "error" && (
            <div style={{ background:"rgba(239,68,68,0.1)", border:"1px solid rgba(239,68,68,0.25)",
                          borderRadius:10, padding:"12px 20px", marginBottom:16,
                          color:"#fca5a5", fontSize:13, fontFamily:"'DM Mono',monospace" }}>
              ⚠ {errorMsg || "Something went wrong. Check your CSV files and try again."}
            </div>
          )}

          {status === "done" && preview && (
            <div style={{ display:"flex", gap:12, marginBottom:20, flexWrap:"wrap", justifyContent:"center" }}>
              {[
                { label:"Projects", value: preview.projCount },
                { label:"Team members", value: preview.peopleCount },
                { label:"Hours logged", value: preview.totalHrs },
                { label:"Sheets generated", value: 5 },
              ].map(s => (
                <div key={s.label} className="stat-chip" style={{ minWidth:120 }}>
                  <div style={{ fontFamily:"'DM Serif Display',serif", fontSize:24, color:"#10b981" }}>{s.value}</div>
                  <div style={{ fontSize:11, color:"#64748b", fontFamily:"'DM Mono',monospace", marginTop:2 }}>{s.label}</div>
                </div>
              ))}
            </div>
          )}

          <button
            className="generate-btn"
            disabled={!ready || status === "parsing" || status === "generating"}
            onClick={handleGenerate}
          >
            {status === "parsing"    ? "⏳ Parsing CSVs…"
           : status === "generating" ? "⚙ Building Report…"
           : status === "done"       ? "✅ Download Again"
           : "⬇ Generate & Download Report"}
          </button>

          {!ready && status === "idle" && (
            <div style={{ marginTop:12, fontSize:12, color:"#334155", fontFamily:"'DM Mono',monospace" }}>
              Upload both CSV files to continue
            </div>
          )}

          {status === "done" && (
            <div style={{ marginTop:12, fontSize:12, color:"#10b981", fontFamily:"'DM Mono',monospace" }}>
              Labour_Report_{month}_{year}.xlsx — saved to your downloads
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{ textAlign:"center", paddingTop:20, borderTop:"1px solid #0f172a" }}>
          <div style={{ fontSize:11, color:"#1e293b", fontFamily:"'DM Mono',monospace" }}>
            SHEETS GENERATED: Dashboard · Project Profitability · Utilisation · Raw Timesheet · Settings
          </div>
        </div>
      </div>
    </div>
  );
}