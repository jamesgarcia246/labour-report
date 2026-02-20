import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";

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

// ─── CSV / file helpers (outside component for stable refs) ──────────────
function parseCSV(text) {
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
}

function readFile(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => res(e.target.result);
    reader.onerror = rej;
    reader.readAsText(file);
  });
}

// ─── Excel style constants ────────────────────────────────────────────────
const XL = {
  fill: {
    title:      { type:"pattern", pattern:"solid", fgColor:{ argb:"FF0F172A" } },
    header:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FF4F46E5" } },
    subheader:  { type:"pattern", pattern:"solid", fgColor:{ argb:"FF1E293B" } },
    subtotal:   { type:"pattern", pattern:"solid", fgColor:{ argb:"FF334155" } },
    total:      { type:"pattern", pattern:"solid", fgColor:{ argb:"FF0F172A" } },
    rowEven:    { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF8FAFC" } },
    rowOdd:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFFFFFF" } },
    onBudget:   { type:"pattern", pattern:"solid", fgColor:{ argb:"FFD1FAE5" } },
    atRisk:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFEF3C7" } },
    overBudget: { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFEE2E2" } },
  },
  font: {
    title:      { name:"Calibri", bold:true,  size:14, color:{ argb:"FFE2E8F0" } },
    header:     { name:"Calibri", bold:true,  size:10, color:{ argb:"FFFFFFFF" } },
    subheader:  { name:"Calibri", bold:true,  size:10, color:{ argb:"FFA5B4FC" } },
    subtotal:   { name:"Calibri", bold:true,  size:10, color:{ argb:"FFE2E8F0" } },
    total:      { name:"Calibri", bold:true,  size:11, color:{ argb:"FFE2E8F0" } },
    body:       { name:"Calibri",             size:10, color:{ argb:"FF1E293B" } },
    bodyBold:   { name:"Calibri", bold:true,  size:10, color:{ argb:"FF1E293B" } },
    note:       { name:"Calibri", italic:true, size:9, color:{ argb:"FF6366F1" } },
    onBudget:   { name:"Calibri", bold:true,  size:10, color:{ argb:"FF065F46" } },
    atRisk:     { name:"Calibri", bold:true,  size:10, color:{ argb:"FF92400E" } },
    overBudget: { name:"Calibri", bold:true,  size:10, color:{ argb:"FF991B1B" } },
  },
  border: {
    light: {
      top:    { style:"thin",   color:{ argb:"FFE2E8F0" } },
      left:   { style:"thin",   color:{ argb:"FFE2E8F0" } },
      bottom: { style:"thin",   color:{ argb:"FFE2E8F0" } },
      right:  { style:"thin",   color:{ argb:"FFE2E8F0" } },
    },
    header: {
      top:    { style:"thin",   color:{ argb:"FF6366F1" } },
      left:   { style:"thin",   color:{ argb:"FF6366F1" } },
      bottom: { style:"medium", color:{ argb:"FF4F46E5" } },
      right:  { style:"thin",   color:{ argb:"FF6366F1" } },
    },
  },
  align: {
    left:        { vertical:"middle", horizontal:"left",   wrapText:false },
    right:       { vertical:"middle", horizontal:"right",  wrapText:false },
    center:      { vertical:"middle", horizontal:"center", wrapText:false },
    leftIndent:  { vertical:"middle", horizontal:"left",   indent:1, wrapText:false },
    rightIndent: { vertical:"middle", horizontal:"right",  indent:1, wrapText:false },
  },
};

function xlStyleRow(row, { fill, font, alignment, border, height } = {}) {
  if (height) row.height = height;
  row.eachCell({ includeEmpty: true }, cell => {
    if (fill)      cell.fill      = fill;
    if (font)      cell.font      = font;
    if (alignment) cell.alignment = alignment;
    if (border)    cell.border    = border;
  });
}

function xlTitleRow(ws, text, colCount) {
  const row = ws.addRow([text]);
  ws.mergeCells(`A${row.number}:${xlCol(colCount)}${row.number}`);
  xlStyleRow(row, { fill:XL.fill.title, font:XL.font.title, alignment:XL.align.leftIndent, height:30 });
  return row;
}

function xlSectionRow(ws, text, colCount) {
  const row = ws.addRow([text]);
  ws.mergeCells(`A${row.number}:${xlCol(colCount)}${row.number}`);
  xlStyleRow(row, { fill:XL.fill.subheader, font:XL.font.subheader, alignment:XL.align.leftIndent, height:18 });
  return row;
}

function xlHeaderRow(ws, values) {
  const row = ws.addRow(values);
  xlStyleRow(row, { fill:XL.fill.header, font:XL.font.header, alignment:XL.align.leftIndent, border:XL.border.header, height:20 });
  return row;
}

function xlDataRow(ws, values, rowIndex) {
  const row = ws.addRow(values);
  const fill = rowIndex % 2 === 0 ? XL.fill.rowEven : XL.fill.rowOdd;
  xlStyleRow(row, { fill, font:XL.font.body, alignment:XL.align.leftIndent, border:XL.border.light, height:18 });
  return row;
}

function xlTotalRow(ws, values) {
  const row = ws.addRow(values);
  xlStyleRow(row, { fill:XL.fill.total, font:XL.font.total, alignment:XL.align.leftIndent, height:22 });
  return row;
}

function xlSubtotalRow(ws, values) {
  const row = ws.addRow(values);
  xlStyleRow(row, { fill:XL.fill.subtotal, font:XL.font.subtotal, alignment:XL.align.leftIndent, height:18 });
  return row;
}

function xlStatusStyle(status) {
  if (status === "On Budget") return { fill:XL.fill.onBudget, font:XL.font.onBudget };
  if (status === "At Risk")   return { fill:XL.fill.atRisk,   font:XL.font.atRisk };
  return                               { fill:XL.fill.overBudget, font:XL.font.overBudget };
}

function xlCol(n) {
  let s = "";
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

// ─── Excel generation ──────────────────────────────────────────────────────
function buildExcel({ timesheetRows, taskRows, team, startMonth, startYear, endMonth, endYear, workingDays }) {
  const rateMap = {};
  team.forEach(t => { rateMap[t.name] = parseFloat(t.rate) || 0; });
  const avgRate = team.reduce((s,t) => s + (parseFloat(t.rate)||0), 0) / team.length;

  const rangeLabel = `${startMonth} ${startYear} — ${endMonth} ${endYear}`;

  // ── Build project map from task-list (budgets only) ──────────────────────
  const projBudget = {};
  const projCustomer = {};
  taskRows.forEach(r => {
    const proj = (r.project || "").trim();
    if (!proj) return;
    const dur = parseFloat(r.duration) || 0;
    if (dur > 0) projBudget[proj] = (projBudget[proj] || 0) + dur;
    if (!projCustomer[proj]) projCustomer[proj] = r.customer_name || "";
  });

  // ── Aggregate ACTUAL hours from timesheet per project ────────────────────
  const projActual = {};
  const personHrs  = {};
  timesheetRows.forEach(r => {
    const proj = (r.project_description || "").trim();
    const name = (r.user_name || "").trim();
    const hrs  = parseFloat(r.hours) || 0;
    if (proj) projActual[proj] = (projActual[proj] || 0) + hrs;
    if (name) personHrs[name]  = (personHrs[name]  || 0) + hrs;
    if (proj && !projCustomer[proj]) projCustomer[proj] = "";
  });

  // ── Build unified project list ───────────────────────────────────────────
  const allProjects = [...new Set([...Object.keys(projBudget), ...Object.keys(projActual)])];
  const projMap = {};
  allProjects.forEach(proj => {
    const budget = projBudget[proj] || 0;
    projMap[proj] = { customer: projCustomer[proj] || "", budget, worked: projActual[proj] || 0, hasQuote: budget > 0 };
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

  // ── Build workbook ───────────────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = "Labour Reporting Tool";
  wb.created = new Date();

  // ── SHEET 1: Dashboard ───────────────────────────────────────────────────
  const wsDash = wb.addWorksheet("Dashboard");
  wsDash.columns = [
    {width:44},{width:28},{width:12},{width:12},{width:13},
    {width:10},{width:13},{width:13},{width:13},{width:11},{width:13},
  ];

  xlTitleRow(wsDash, `${rangeLabel} — Labour Dashboard`, 11);
  wsDash.addRow([]);
  xlSectionRow(wsDash, "KEY PERFORMANCE INDICATORS", 11);

  const kpiRows = [
    ["Total Budget Hours",      `${fmtHrs(totalBudgetHrs)} hrs`],
    ["Total Actual Hours",      `${fmtHrs(totalActualHrs)} hrs`],
    ["Hours Variance",          fmtHrsVar(totalActualHrs - totalBudgetHrs) + " hrs"],
    ["Labour Budget",           fmtGBP(totalBudgetGBP)],
    ["Actual Labour Cost",      fmtGBP(totalActualGBP)],
    ["Labour Variance",         fmtVar(totalVariance)],
    ["Overall Margin %",        fmtPct(overallMarginPc)],
    ["Team Utilisation %",      fmtPct(teamUtil)],
    ["Total Capacity Cost",     fmtGBP(totalCapacity)],
    ["Undeployed Cost",         fmtGBP(totalUndeployed)],
  ];
  kpiRows.forEach(([label, val], i) => {
    const row = xlDataRow(wsDash, [label, val], i);
    row.getCell(1).font = XL.font.bodyBold;
    row.getCell(2).alignment = XL.align.rightIndent;
  });

  wsDash.addRow([]);
  xlSectionRow(wsDash, "PROJECT STATUS", 11);
  xlHeaderRow(wsDash, ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Variance","Burn %","Budget £","Actual £","Variance £","Margin %","Status"]);

  budgeted.forEach(([proj, v], i) => {
    const burn   = v.budget > 0 ? v.worked / v.budget : 0;
    const bGBP   = v.budget * avgRate;
    const aGBP   = v.worked * avgRate;
    const varGBP = bGBP - aGBP;
    const margin = bGBP > 0 ? varGBP / bGBP : 0;
    const status = burn <= 0.95 ? "On Budget" : burn <= 1.10 ? "At Risk" : "Over Budget";
    const row = xlDataRow(wsDash, [
      proj, v.customer,
      fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
      fmtPct(burn), fmtGBP(bGBP), fmtGBP(aGBP), fmtVar(varGBP), fmtPct(margin), status,
    ], i);
    const { fill, font } = xlStatusStyle(status);
    row.getCell(11).fill = fill;
    row.getCell(11).font = font;
    row.getCell(11).alignment = XL.align.center;
  });

  wsDash.addRow([]);
  xlSectionRow(wsDash, "PROJECTS WITHOUT QUOTED BUDGET — excluded from margin calculations", 11);
  xlHeaderRow(wsDash, ["Project","Customer","Actual Hrs"]);
  noBudget.forEach(([proj,v], i) => xlDataRow(wsDash, [proj, v.customer, fmtHrs(v.worked)], i));

  // ── SHEET 2: Project Profitability ───────────────────────────────────────
  const wsProj = wb.addWorksheet("Project Profitability");
  wsProj.columns = [
    {width:44},{width:28},{width:12},{width:12},{width:13},
    {width:10},{width:13},{width:14},{width:13},{width:13},{width:11},{width:13},
  ];

  xlTitleRow(wsProj, `PROJECT PROFITABILITY — ${rangeLabel}`, 12);
  const noteRow = wsProj.addRow([`Avg blended rate: £${avgRate.toFixed(2)}/hr  |  Actual hours sourced from timesheet export`]);
  wsProj.mergeCells(`A${noteRow.number}:L${noteRow.number}`);
  xlStyleRow(noteRow, { fill:XL.fill.subheader, font:XL.font.note, alignment:XL.align.leftIndent, height:16 });
  wsProj.addRow([]);

  xlHeaderRow(wsProj, ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Variance","Burn %","Budget £","Actual Cost £","Variance £","Margin £","Margin %","Status"]);

  budgeted.forEach(([proj,v], i) => {
    const burn   = v.budget > 0 ? v.worked / v.budget : 0;
    const bGBP   = v.budget * avgRate;
    const aGBP   = v.worked * avgRate;
    const varGBP = bGBP - aGBP;
    const margin = bGBP > 0 ? varGBP / bGBP : 0;
    const status = burn <= 0.95 ? "On Budget" : burn <= 1.10 ? "At Risk" : "Over Budget";
    const row = xlDataRow(wsProj, [
      proj, v.customer,
      fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
      fmtPct(burn), fmtGBP(bGBP), fmtGBP(aGBP), fmtVar(varGBP),
      fmtVar(varGBP), fmtPct(margin), status,
    ], i);
    const { fill, font } = xlStatusStyle(status);
    row.getCell(12).fill = fill;
    row.getCell(12).font = font;
    row.getCell(12).alignment = XL.align.center;
  });

  xlTotalRow(wsProj, [
    "TOTAL", "",
    fmtHrs(totalBudgetHrs), fmtHrs(totalActualHrs),
    fmtHrsVar(totalActualHrs - totalBudgetHrs),
    fmtPct(totalBudgetHrs > 0 ? totalActualHrs / totalBudgetHrs : 0),
    fmtGBP(totalBudgetGBP), fmtGBP(totalActualGBP),
    fmtVar(totalVariance), fmtVar(totalVariance),
    fmtPct(overallMarginPc), "",
  ]);

  wsProj.addRow([]);
  xlSectionRow(wsProj, "NO QUOTED BUDGET — hours logged, excluded from margin", 12);
  xlHeaderRow(wsProj, ["Project","Customer","Actual Hrs"]);
  noBudget.forEach(([proj,v], i) => xlDataRow(wsProj, [proj, v.customer, fmtHrs(v.worked)], i));

  // ── SHEET 3: Utilisation ─────────────────────────────────────────────────
  const wsUtil = wb.addWorksheet("Utilisation");
  wsUtil.columns = [
    {width:26},{width:14},{width:15},{width:14},{width:14},
    {width:14},{width:20},{width:20},{width:20},
  ];

  xlTitleRow(wsUtil, `LABOUR UTILISATION — ${rangeLabel}`, 9);
  const utilNoteRow = wsUtil.addRow([`Available hours: ${workingDays} total working days × 8 hrs = ${availHrs} hrs/person`]);
  wsUtil.mergeCells(`A${utilNoteRow.number}:I${utilNoteRow.number}`);
  xlStyleRow(utilNoteRow, { fill:XL.fill.subheader, font:XL.font.note, alignment:XL.align.leftIndent, height:16 });
  wsUtil.addRow([]);

  xlHeaderRow(wsUtil, ["Team Member","Team","Hourly Rate (£)","Available Hrs","Hours Worked","Utilisation %","Capacity Cost (£)","Labour Cost (£)","Undeployed Cost (£)"]);

  const teamGroups = [...new Set(team.map(t => t.team))];
  teamGroups.forEach(grp => {
    const members = team.filter(t => t.team === grp);
    members.forEach((t, i) => {
      const rate    = parseFloat(t.rate) || 0;
      const worked  = personHrs[t.name] || 0;
      const util    = availHrs > 0 ? worked / availHrs : 0;
      const capCost = availHrs * rate;
      const labCost = worked * rate;
      xlDataRow(wsUtil, [
        t.name, t.team || "", fmtGBP(rate), availHrs,
        fmtHrs(worked), fmtPct(util),
        fmtGBP(capCost), fmtGBP(labCost), fmtGBP(capCost - labCost),
      ], i);
    });
    const grpWorked   = members.reduce((s,t) => s + (personHrs[t.name]||0), 0);
    const grpCap      = members.reduce((s,t) => s + availHrs*(parseFloat(t.rate)||0), 0);
    const grpDeployed = members.reduce((s,t) => s + (personHrs[t.name]||0)*(parseFloat(t.rate)||0), 0);
    const grpAvail    = members.length * availHrs;
    xlSubtotalRow(wsUtil, [
      `— ${grp} Subtotal`, "", "", grpAvail, fmtHrs(grpWorked),
      fmtPct(grpAvail > 0 ? grpWorked / grpAvail : 0),
      fmtGBP(grpCap), fmtGBP(grpDeployed), fmtGBP(grpCap - grpDeployed),
    ]);
    wsUtil.addRow([]);
  });

  xlTotalRow(wsUtil, [
    "TOTAL", "", "", totalAvailHrs, fmtHrs(totalWorkedHrs),
    fmtPct(teamUtil), fmtGBP(totalCapacity), fmtGBP(totalDeployed), fmtGBP(totalUndeployed),
  ]);

  // ── SHEET 4: Raw Timesheet ───────────────────────────────────────────────
  const wsTs = wb.addWorksheet("Raw Timesheet");
  wsTs.columns = [
    {width:12},{width:30},{width:42},{width:22},{width:12},{width:8},{width:9},{width:30},
  ];

  xlHeaderRow(wsTs, ["Job No.","Task","Project","Team Member","Date","Hours","Billable","Notes"]);
  timesheetRows.forEach((r, i) => {
    xlDataRow(wsTs, [
      r.work_order_no || "",
      r.task_name || "",
      (r.project_description || "").trim(),
      r.user_name || "",
      (r.start_datetime || "").slice(0,10),
      parseFloat(r.hours) || 0,
      r.is_billable === "1" ? "Yes" : "No",
      r.notes || "",
    ], i);
  });

  // ── SHEET 5: Settings snapshot ───────────────────────────────────────────
  const wsSet = wb.addWorksheet("Settings");
  wsSet.columns = [{width:26},{width:22},{width:22}];

  xlTitleRow(wsSet, "REPORT SETTINGS SNAPSHOT", 3);
  const settingsKV = [
    ["Report Start",         `${startMonth} ${startYear}`],
    ["Report End",           `${endMonth} ${endYear}`],
    ["Total Working Days",   String(workingDays)],
    ["Hrs per Day",          "8"],
    ["Available Hrs/Person", String(availHrs)],
    ["Avg Hourly Rate",      fmtGBP(avgRate) + "/hr"],
  ];
  settingsKV.forEach(([k,v], i) => {
    const row = xlDataRow(wsSet, [k, v], i);
    row.getCell(1).font = XL.font.bodyBold;
  });

  wsSet.addRow([]);
  xlSectionRow(wsSet, "TEAM RATES", 3);
  xlHeaderRow(wsSet, ["Name","Team","Hourly Rate (£)"]);
  team.forEach((t, i) => {
    xlDataRow(wsSet, [t.name, t.team || "", fmtGBP(parseFloat(t.rate)||0) + "/hr"], i);
  });

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

// ─── Month range helper ────────────────────────────────────────────────────
function getMonthRange(startMonth, startYear, endMonth, endYear) {
  const result = [];
  let m = MONTHS.indexOf(startMonth);
  let y = parseInt(startYear);
  const endM = MONTHS.indexOf(endMonth);
  const endY = parseInt(endYear);
  while (y < endY || (y === endY && m <= endM)) {
    result.push({ month: MONTHS[m], year: y });
    m++;
    if (m > 11) { m = 0; y++; }
  }
  return result;
}

// ─── Main App ──────────────────────────────────────────────────────────────
export default function App() {
  const [timesheetFile, setTimesheetFile] = useState(null);
  const [taskFile,      setTaskFile]      = useState(null);
  const [team,          setTeam]          = useState(TEAM_DEFAULT);
  const [startMonth,    setStartMonth]    = useState("February");
  const [startYear,     setStartYear]     = useState(2026);
  const [endMonth,      setEndMonth]      = useState("February");
  const [endYear,       setEndYear]       = useState(2026);
  const [monthlyWorkingDays, setMonthlyWorkingDays] = useState({});
  const [status,        setStatus]        = useState("idle");
  const [errorMsg,      setErrorMsg]      = useState("");
  const [preview,       setPreview]       = useState(null);

  const [parsedTimesheetRows, setParsedTimesheetRows] = useState([]);
  const [parsedTaskRows,      setParsedTaskRows]      = useState([]);
  const [selectedProject,     setSelectedProject]     = useState("All Projects");

  // Parse timesheet CSV reactively whenever the file changes
  useEffect(() => {
    if (!timesheetFile) { setParsedTimesheetRows([]); return; }
    let cancelled = false;
    readFile(timesheetFile).then(text => {
      if (!cancelled) setParsedTimesheetRows(parseCSV(text));
    });
    return () => { cancelled = true; };
  }, [timesheetFile]);

  // Parse task CSV reactively; reset project selection when file changes
  useEffect(() => {
    if (!taskFile) { setParsedTaskRows([]); setSelectedProject("All Projects"); return; }
    let cancelled = false;
    readFile(taskFile).then(text => {
      if (!cancelled) {
        setParsedTaskRows(parseCSV(text));
        setSelectedProject("All Projects");
      }
    });
    return () => { cancelled = true; };
  }, [taskFile]);

  // Unique project names from the task list CSV
  const uniqueProjects = useMemo(() =>
    [...new Set(parsedTaskRows.map(r => (r.project || "").trim()).filter(Boolean))].sort(),
    [parsedTaskRows]
  );

  // Live summary for the project filter panel
  const filteredSummary = useMemo(() => {
    if (!parsedTimesheetRows.length) return { hrs: 0, count: 0 };
    const sDate = new Date(startYear, MONTHS.indexOf(startMonth), 1);
    const eDate = new Date(endYear, MONTHS.indexOf(endMonth) + 1, 0);
    let rows = parsedTimesheetRows.filter(r => {
      const dateStr = (r.start_datetime || "").slice(0, 10);
      if (!dateStr) return false;
      const d = new Date(dateStr + "T00:00:00");
      return d >= sDate && d <= eDate;
    });
    if (selectedProject !== "All Projects") {
      rows = rows.filter(r => (r.project_description || "").trim() === selectedProject);
    }
    return {
      hrs:   rows.reduce((s, r) => s + (parseFloat(r.hours) || 0), 0),
      count: rows.length,
    };
  }, [parsedTimesheetRows, startMonth, startYear, endMonth, endYear, selectedProject]);

  const monthRange = useMemo(
    () => getMonthRange(startMonth, startYear, endMonth, endYear),
    [startMonth, startYear, endMonth, endYear]
  );

  const totalWorkingDays = useMemo(
    () => monthRange.reduce((sum, { month, year }) => {
      const key = `${month} ${year}`;
      return sum + (parseInt(monthlyWorkingDays[key]) || 20);
    }, 0),
    [monthRange, monthlyWorkingDays]
  );

  const handleGenerate = async () => {
    if (!timesheetFile || !taskFile) return;
    setStatus("parsing");
    setErrorMsg("");
    try {
      // Use pre-parsed rows from state; fall back to re-parsing if not ready yet
      let allTimesheetRows = parsedTimesheetRows;
      let allTaskRows      = parsedTaskRows;
      if (!allTimesheetRows.length || !allTaskRows.length) {
        const [tsText, taskText] = await Promise.all([
          readFile(timesheetFile),
          readFile(taskFile),
        ]);
        allTimesheetRows = parseCSV(tsText);
        allTaskRows      = parseCSV(taskText);
      }

      // Filter by date range (inclusive)
      const sDate = new Date(startYear, MONTHS.indexOf(startMonth), 1);
      const eDate = new Date(endYear, MONTHS.indexOf(endMonth) + 1, 0);
      let filteredTimesheetRows = allTimesheetRows.filter(r => {
        const dateStr = (r.start_datetime || "").slice(0, 10);
        if (!dateStr) return false;
        const d = new Date(dateStr + "T00:00:00");
        return d >= sDate && d <= eDate;
      });

      // Filter by project
      let filteredTaskRows = allTaskRows;
      if (selectedProject !== "All Projects") {
        filteredTimesheetRows = filteredTimesheetRows.filter(r =>
          (r.project_description || "").trim() === selectedProject
        );
        filteredTaskRows = allTaskRows.filter(r =>
          (r.project || "").trim() === selectedProject
        );
      }

      setStatus("generating");

      // Build preview summary
      const projCount   = new Set(filteredTaskRows.map(r => (r.project||"").trim()).filter(Boolean)).size;
      const peopleCount = new Set(filteredTimesheetRows.map(r => r.user_name).filter(Boolean)).size;
      const totalHrs    = filteredTimesheetRows.reduce((s,r) => s + (parseFloat(r.hours)||0), 0);
      setPreview({ projCount, peopleCount, totalHrs: totalHrs.toFixed(1) });

      const wb = buildExcel({
        timesheetRows: filteredTimesheetRows,
        taskRows:      filteredTaskRows,
        team,
        startMonth, startYear, endMonth, endYear,
        workingDays: totalWorkingDays,
      });

      const projectLabel = selectedProject !== "All Projects"
        ? `${selectedProject.replace(/[^a-zA-Z0-9]/g, "_")}_`
        : "";
      const filename = `Labour_Report_${projectLabel}${startMonth}${startYear}_to_${endMonth}${endYear}.xlsx`;

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

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

  const projectLabel = selectedProject !== "All Projects"
    ? `${selectedProject.replace(/[^a-zA-Z0-9]/g, "_")}_`
    : "";
  const currentFilename = `Labour_Report_${projectLabel}${startMonth}${startYear}_to_${endMonth}${endYear}.xlsx`;

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
          <div style={{ display:"flex", flexDirection:"column", gap:20 }}>

            {/* Date range row */}
            <div style={{ display:"flex", gap:20, flexWrap:"wrap", alignItems:"flex-end" }}>

              {/* FROM */}
              <div>
                <div style={{ fontSize:11, color:"#6366f1", fontFamily:"'DM Mono',monospace",
                              letterSpacing:1, marginBottom:8 }}>FROM</div>
                <div style={{ display:"flex", gap:10, alignItems:"flex-end" }}>
                  <div>
                    <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>MONTH</div>
                    <select className="field" value={startMonth} onChange={e => setStartMonth(e.target.value)}>
                      {MONTHS.map(m => <option key={m}>{m}</option>)}
                    </select>
                  </div>
                  <div>
                    <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>YEAR</div>
                    <input className="field" type="number" value={startYear}
                      onChange={e => setStartYear(parseInt(e.target.value)||2026)} min={2020} max={2035} />
                  </div>
                </div>
              </div>

              {/* Arrow */}
              <div style={{ color:"#334155", fontSize:20, paddingBottom:10 }}>→</div>

              {/* TO */}
              <div>
                <div style={{ fontSize:11, color:"#6366f1", fontFamily:"'DM Mono',monospace",
                              letterSpacing:1, marginBottom:8 }}>TO</div>
                <div style={{ display:"flex", gap:10, alignItems:"flex-end" }}>
                  <div>
                    <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>MONTH</div>
                    <select className="field" value={endMonth} onChange={e => setEndMonth(e.target.value)}>
                      {MONTHS.map(m => <option key={m}>{m}</option>)}
                    </select>
                  </div>
                  <div>
                    <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>YEAR</div>
                    <input className="field" type="number" value={endYear}
                      onChange={e => setEndYear(parseInt(e.target.value)||2026)} min={2020} max={2035} />
                  </div>
                </div>
              </div>
            </div>

            {/* Per-month working days grid */}
            {monthRange.length > 0 && (
              <div>
                <div style={{ fontSize:12, color:"#64748b", marginBottom:8, fontFamily:"'DM Mono',monospace" }}>
                  WORKING DAYS PER MONTH
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(110px, 1fr))", gap:8 }}>
                  {monthRange.map(({ month, year }) => {
                    const key = `${month} ${year}`;
                    const val = monthlyWorkingDays[key] !== undefined ? monthlyWorkingDays[key] : 20;
                    return (
                      <div key={key} style={{
                        background:"rgba(255,255,255,0.03)", border:"1px solid #1e293b",
                        borderRadius:8, padding:"8px 10px",
                      }}>
                        <div style={{ fontSize:10, color:"#6366f1", fontFamily:"'DM Mono',monospace", marginBottom:6 }}>
                          {month.slice(0,3).toUpperCase()} {year}
                        </div>
                        <input
                          className="field"
                          type="number"
                          value={val}
                          onChange={e => setMonthlyWorkingDays(prev => ({
                            ...prev, [key]: parseInt(e.target.value) || 20,
                          }))}
                          min={1} max={31}
                          style={{ width:"100%" }}
                        />
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Available hrs info */}
            <div style={{ alignSelf:"flex-start", background:"rgba(99,102,241,0.06)",
                          border:"1px solid rgba(99,102,241,0.15)", borderRadius:10,
                          padding:"12px 16px" }}>
              <div style={{ fontSize:11, color:"#6366f1", fontFamily:"'DM Mono',monospace", marginBottom:4 }}>
                AVAILABLE HRS / PERSON
              </div>
              <div style={{ fontSize:22, fontFamily:"'DM Serif Display',serif", color:"#a5b4fc" }}>
                {totalWorkingDays * 8}
                <span style={{ fontSize:13, color:"#64748b", marginLeft:6 }}>hrs</span>
              </div>
              <div style={{ fontSize:11, color:"#475569", fontFamily:"'DM Mono',monospace", marginTop:4 }}>
                {totalWorkingDays} days across {monthRange.length} {monthRange.length === 1 ? "month" : "months"}
              </div>
            </div>

          </div>
        </div>

        {/* Project Filter – only visible when both CSVs are uploaded */}
        {ready && (
          <div className="card">
            <div className="section-label">Project Filter — Optional</div>
            <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
              <div>
                <div style={{ fontSize:12, color:"#64748b", marginBottom:6, fontFamily:"'DM Mono',monospace" }}>
                  PROJECT
                </div>
                <select
                  className="field"
                  value={selectedProject}
                  onChange={e => setSelectedProject(e.target.value)}
                >
                  <option value="All Projects">All Projects</option>
                  {uniqueProjects.map(p => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </div>
              <div style={{ fontSize:12, color:"#475569", fontFamily:"'DM Mono',monospace" }}>
                {filteredSummary.hrs.toFixed(1)} hrs
                &nbsp;·&nbsp;
                {filteredSummary.count} {filteredSummary.count === 1 ? "entry" : "entries"}
                {selectedProject !== "All Projects" && (
                  <span style={{ color:"#6366f1" }}> — filtered to {selectedProject}</span>
                )}
              </div>
            </div>
          </div>
        )}

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
              {currentFilename} — saved to your downloads
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
