import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import ExcelJS from "exceljs";

// ─── Constants ────────────────────────────────────────────────────────────────
const DEPARTMENTS = ["Design", "Engineering", "Presales", "Projects", "Service"];

const TEAM_DEFAULT = [
  { name: "James Owens",         dept: "Design",      rate: 50 },
  { name: "Michael Hall",        dept: "Design",      rate: 50 },
  { name: "Ollie Allington",     dept: "Design",      rate: 50 },
  { name: "Andy Timms",          dept: "Engineering", rate: 50 },
  { name: "Cameron Birtwhistle", dept: "Engineering", rate: 50 },
  { name: "David Gurjao ",       dept: "Engineering", rate: 50 },
  { name: "Dermot McGee",        dept: "Engineering", rate: 50 },
  { name: "Dan Ionut ",          dept: "Engineering", rate: 50 },
  { name: "James Harris",        dept: "Engineering", rate: 50 },
  { name: "Leonidas Likskendaj", dept: "Engineering", rate: 50 },
  { name: "Magic Dulkowski ",    dept: "Engineering", rate: 50 },
  { name: "Reiss lewis",         dept: "Engineering", rate: 50 },
  { name: "Ando Goddess",        dept: "Presales",    rate: 50 },
  { name: "Matt Horsfall",       dept: "Presales",    rate: 50 },
  { name: "Aaron-James Pienaar", dept: "Projects",    rate: 50 },
  { name: "Rob Elkins",          dept: "Projects",    rate: 50 },
  { name: "Rupert Merryweather", dept: "Projects",    rate: 50 },
  { name: "Darren Williams",     dept: "Service",     rate: 50 },
  { name: "Sodrul Islam",        dept: "Service",     rate: 50 },
];

const REPORT_TYPES = [
  { id: "all",     label: "All Projects",  icon: "◈", desc: "Full portfolio view — all projects and all team members" },
  { id: "project", label: "Project View",  icon: "◉", desc: "Filter by one or more projects — optimised for project managers" },
  { id: "person",  label: "Person View",   icon: "◎", desc: "Filter by one or more people — optimised for utilisation review" },
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const normName = s => (s || "").trim().toLowerCase().replace(/\s+/g, " ");

function parseCSV(text) {
  const lines = text.trim().split("\n");
  const headers = lines[0].split(",").map(h => h.replace(/"/g, "").trim());
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
    headers.forEach((h, i) => { obj[h] = vals[i] || ""; });
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

function fmtGBP(n)    { return `£${Math.round(n).toLocaleString("en-GB")}`; }
function fmtPct(n)    { return `${(n * 100).toFixed(1)}%`; }
function fmtHrs(n)    { return +n.toFixed(1); }
function fmtVar(n)    { return n >= 0 ? `+${fmtGBP(n)}` : `-£${Math.abs(Math.round(n)).toLocaleString("en-GB")}`; }
function gbp(n)       { return { _gbp: true, v: Math.round(n) }; }
function fmtHrsVar(n) { return n >= 0 ? `+${n.toFixed(1)}` : n.toFixed(1); }
function fmtDate(d)   { return d ? d.toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" }) : ""; }
function calcWorkingDays(start, end) {
  const s = new Date(start + "T00:00:00"), e = new Date(end + "T00:00:00");
  let count = 0, cur = new Date(s);
  while (cur <= e) { const d = cur.getDay(); if (d >= 1 && d <= 5) count++; cur.setDate(cur.getDate() + 1); }
  return count;
}

// ─── Excel Styles ─────────────────────────────────────────────────────────────
const XL = {
  fill: {
    title:      { type:"pattern", pattern:"solid", fgColor:{ argb:"FF0A0F1E" } },
    header:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FF1A2744" } },
    subheader:  { type:"pattern", pattern:"solid", fgColor:{ argb:"FF111827" } },
    subtotal:   { type:"pattern", pattern:"solid", fgColor:{ argb:"FF1E293B" } },
    total:      { type:"pattern", pattern:"solid", fgColor:{ argb:"FF0A0F1E" } },
    rowEven:    { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF8FAFC" } },
    rowOdd:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFFFFFF" } },
    onBudget:   { type:"pattern", pattern:"solid", fgColor:{ argb:"FFD1FAE5" } },
    atRisk:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFEF3C7" } },
    overBudget: { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFEE2E2" } },
    accent:     { type:"pattern", pattern:"solid", fgColor:{ argb:"FF1D4ED8" } },
  },
  font: {
    title:      { name:"Calibri", bold:true,   size:14, color:{ argb:"FFE2E8F0" } },
    header:     { name:"Calibri", bold:true,   size:10, color:{ argb:"FFE2E8F0" } },
    subheader:  { name:"Calibri", bold:true,   size:10, color:{ argb:"FF93C5FD" } },
    subtotal:   { name:"Calibri", bold:true,   size:10, color:{ argb:"FFE2E8F0" } },
    total:      { name:"Calibri", bold:true,   size:11, color:{ argb:"FFE2E8F0" } },
    body:       { name:"Calibri",              size:10, color:{ argb:"FF1E293B" } },
    bodyBold:   { name:"Calibri", bold:true,   size:10, color:{ argb:"FF1E293B" } },
    note:       { name:"Calibri", italic:true, size:9,  color:{ argb:"FF6B7280" } },
    onBudget:   { name:"Calibri", bold:true,   size:10, color:{ argb:"FF065F46" } },
    atRisk:     { name:"Calibri", bold:true,   size:10, color:{ argb:"FF92400E" } },
    overBudget: { name:"Calibri", bold:true,   size:10, color:{ argb:"FF991B1B" } },
    accent:     { name:"Calibri", bold:true,   size:10, color:{ argb:"FFFFFFFF" } },
  },
  border: {
    light: {
      top:    { style:"thin",   color:{ argb:"FFE2E8F0" } },
      bottom: { style:"thin",   color:{ argb:"FFE2E8F0" } },
      left:   { style:"thin",   color:{ argb:"FFE2E8F0" } },
      right:  { style:"thin",   color:{ argb:"FFE2E8F0" } },
    },
    header: {
      top:    { style:"thin",   color:{ argb:"FF1D4ED8" } },
      bottom: { style:"medium", color:{ argb:"FF1D4ED8" } },
      left:   { style:"thin",   color:{ argb:"FF1D4ED8" } },
      right:  { style:"thin",   color:{ argb:"FF1D4ED8" } },
    },
  },
  align: {
    left:   { vertical:"middle", horizontal:"left",   wrapText:false },
    right:  { vertical:"middle", horizontal:"right",  wrapText:false },
    center: { vertical:"middle", horizontal:"center", wrapText:false },
    li:     { vertical:"middle", horizontal:"left",   indent:1 },
    ri:     { vertical:"middle", horizontal:"right",  indent:1 },
  },
};

function xlStyleRow(row, { fill, font, alignment, border, height } = {}) {
  if (height) row.height = height;
  row.eachCell({ includeEmpty:true }, cell => {
    if (fill)      cell.fill      = fill;
    if (font)      cell.font      = font;
    if (alignment) cell.alignment = alignment;
    if (border)    cell.border    = border;
  });
}

function xlCol(n) {
  let s = "";
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function xlTitle(ws, text, cols) {
  const row = ws.addRow([text]);
  ws.mergeCells(`A${row.number}:${xlCol(cols)}${row.number}`);
  xlStyleRow(row, { fill:XL.fill.title, font:XL.font.title, alignment:XL.align.li, height:30 });
  return row;
}

function xlSection(ws, text, cols) {
  const row = ws.addRow([text]);
  ws.mergeCells(`A${row.number}:${xlCol(cols)}${row.number}`);
  xlStyleRow(row, { fill:XL.fill.subheader, font:XL.font.subheader, alignment:XL.align.li, height:18 });
  return row;
}

function xlHeader(ws, vals) {
  const row = ws.addRow(vals);
  xlStyleRow(row, { fill:XL.fill.header, font:XL.font.header, alignment:XL.align.li, border:XL.border.header, height:20 });
  return row;
}

function xlData(ws, vals, idx) {
  const rawVals = vals.map(v => (v && typeof v === 'object' && v._gbp) ? v.v : v);
  const row = ws.addRow(rawVals);
  xlStyleRow(row, {
    fill: idx % 2 === 0 ? XL.fill.rowEven : XL.fill.rowOdd,
    font: XL.font.body, alignment: XL.align.li, border: XL.border.light, height:18,
  });
  vals.forEach((v, i) => { if (v && typeof v === 'object' && v._gbp) row.getCell(i+1).numFmt = '£#,##0'; });
  return row;
}

function xlSubtotal(ws, vals) {
  const rawVals = vals.map(v => (v && typeof v === 'object' && v._gbp) ? v.v : v);
  const row = ws.addRow(rawVals);
  xlStyleRow(row, { fill:XL.fill.subtotal, font:XL.font.subtotal, alignment:XL.align.li, height:18 });
  vals.forEach((v, i) => { if (v && typeof v === 'object' && v._gbp) row.getCell(i+1).numFmt = '£#,##0'; });
  return row;
}

function xlTotal(ws, vals) {
  const rawVals = vals.map(v => (v && typeof v === 'object' && v._gbp) ? v.v : v);
  const row = ws.addRow(rawVals);
  xlStyleRow(row, { fill:XL.fill.total, font:XL.font.total, alignment:XL.align.li, height:22 });
  vals.forEach((v, i) => { if (v && typeof v === 'object' && v._gbp) row.getCell(i+1).numFmt = '£#,##0'; });
  return row;
}

function statusStyle(burn) {
  if (burn <= 0.95) return { fill:XL.fill.onBudget,   font:XL.font.onBudget,   label:"On Budget"   };
  if (burn <= 1.10) return { fill:XL.fill.atRisk,     font:XL.font.atRisk,     label:"At Risk"     };
  return              { fill:XL.fill.overBudget, font:XL.font.overBudget, label:"Over Budget" };
}

// ─── Excel Builders ───────────────────────────────────────────────────────────
function buildAllProjectsReport({ tsRows, taskRows, team, rateMap, avgRate, startDate, endDate, workingDays }) {
  const rangeLabel = `${fmtDate(startDate)} — ${fmtDate(endDate)}`;

  // Budget from task list
  const projBudget = {}, projCustomer = {};
  taskRows.forEach(r => {
    const p = (r.project || "").trim(); if (!p) return;
    const d = parseFloat(r.duration) || 0;
    if (d > 0) projBudget[p] = (projBudget[p] || 0) + d;
    if (!projCustomer[p]) projCustomer[p] = r.customer_name || "";
  });

  // Actuals from timesheet
  const projActual = {}, personHrs = {};
  tsRows.forEach(r => {
    const p = (r.project_description || "").trim();
    const n = (r.user_name || "").trim();
    const h = parseFloat(r.hours) || 0;
    if (p) projActual[p] = (projActual[p] || 0) + h;
    if (n) personHrs[n]  = (personHrs[n]  || 0) + h;
    if (p && !projCustomer[p]) projCustomer[p] = "";
  });

  const allProjs = [...new Set([...Object.keys(projBudget), ...Object.keys(projActual)])];
  const projMap = {};
  allProjs.forEach(p => {
    const b = projBudget[p] || 0;
    projMap[p] = { customer: projCustomer[p] || "", budget: b, worked: projActual[p] || 0, hasQuote: b > 0 };
  });

  const budgeted = Object.entries(projMap).filter(([,v]) => v.hasQuote)
    .sort(([,a],[,b]) => (b.budget > 0 ? b.worked/b.budget : 0) - (a.budget > 0 ? a.worked/a.budget : 0));
  const noBudget = Object.entries(projMap).filter(([,v]) => !v.hasQuote);

  const totalBudgetHrs = budgeted.reduce((s,[,v]) => s + v.budget, 0);
  const totalActualHrs = budgeted.reduce((s,[,v]) => s + v.worked, 0);
  const totalBudgetGBP = totalBudgetHrs * avgRate;
  const totalActualGBP = totalActualHrs * avgRate;
  const totalVariance  = totalBudgetGBP - totalActualGBP;
  const overallMargin  = totalBudgetGBP > 0 ? totalVariance / totalBudgetGBP : 0;
  const availHrs       = workingDays * 8;
  const totalCapacity  = team.reduce((s,t) => s + availHrs * (parseFloat(t.rate)||0), 0);
  const totalDeployed  = Object.entries(personHrs).reduce((s,[n,h]) => s + h*(rateMap[n]||avgRate), 0);
  const totalWorkedHrs = Object.values(personHrs).reduce((s,h) => s+h, 0);
  const totalAvailHrs  = team.length * availHrs;
  const teamUtil       = totalAvailHrs > 0 ? totalWorkedHrs / totalAvailHrs : 0;

  const wb = new ExcelJS.Workbook();
  wb.creator = "Labour Report v2";
  wb.created = new Date();

  // ── Sheet 1: Dashboard ──────────────────────────────────────────────────────
  const wsDash = wb.addWorksheet("Dashboard");
  wsDash.columns = [
    {width:44},{width:28},{width:12},{width:12},{width:13},
    {width:10},{width:13},{width:13},{width:13},{width:11},{width:13},
  ];
  xlTitle(wsDash, `ALL PROJECTS — Labour Dashboard  |  ${rangeLabel}`, 11);
  wsDash.addRow([]);
  xlSection(wsDash, "KEY PERFORMANCE INDICATORS", 11);
  [
    ["Report Period",       rangeLabel],
    ["Working Days",        `${workingDays} days`],
    ["Labour Variance",     gbp(totalVariance)],
    ["Overall Margin %",    fmtPct(overallMargin)],
    ["Total Budget Hours",  `${fmtHrs(totalBudgetHrs)} hrs`],
    ["Total Actual Hours",  `${fmtHrs(totalActualHrs)} hrs`],
    ["Hours Variance",      fmtHrsVar(totalActualHrs - totalBudgetHrs) + " hrs"],
    ["Labour Budget",       gbp(totalBudgetGBP)],
    ["Actual Labour Cost",  gbp(totalActualGBP)],
    ["Team Utilisation %",  fmtPct(teamUtil)],
    ["Total Capacity Cost", gbp(totalCapacity)],
    ["Undeployed Cost",     gbp(totalCapacity - totalDeployed)],
  ].forEach(([k,v], i) => {
    const r = xlData(wsDash, [k, v], i);
    r.getCell(1).font = XL.font.bodyBold;
    r.getCell(2).alignment = XL.align.ri;
  });

  wsDash.addRow([]);
  xlSection(wsDash, "PROJECT STATUS SUMMARY", 11);
  xlHeader(wsDash, ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Var","Burn %","Budget £","Actual £","Variance £","Margin %","Status"]);
  budgeted.forEach(([proj,v], i) => {
    const burn = v.budget > 0 ? v.worked / v.budget : 0;
    const bGBP = v.budget * avgRate, aGBP = v.worked * avgRate;
    const st = statusStyle(burn);
    const row = xlData(wsDash, [
      proj, v.customer,
      fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
      fmtPct(burn), gbp(bGBP), gbp(aGBP), gbp(bGBP - aGBP),
      fmtPct(bGBP > 0 ? (bGBP - aGBP)/bGBP : 0), st.label,
    ], i);
    row.getCell(11).fill = st.fill; row.getCell(11).font = st.font;
    row.getCell(11).alignment = XL.align.center;
  });

  if (noBudget.length) {
    wsDash.addRow([]);
    xlSection(wsDash, "NO QUOTED BUDGET — excluded from margin calculations", 3);
    xlHeader(wsDash, ["Project","Customer","Actual Hrs"]);
    noBudget.forEach(([p,v], i) => xlData(wsDash, [p, v.customer, fmtHrs(v.worked)], i));
  }

  // ── Sheet 2: Project Profitability ──────────────────────────────────────────
  const wsProj = wb.addWorksheet("Project Profitability");
  wsProj.columns = [
    {width:44},{width:28},{width:12},{width:12},{width:13},
    {width:10},{width:13},{width:14},{width:13},{width:11},{width:13},
  ];
  xlTitle(wsProj, `PROJECT PROFITABILITY  |  ${rangeLabel}`, 11);
  const noteR = wsProj.addRow([`Avg blended rate: £${avgRate.toFixed(2)}/hr  ·  Actuals sourced from timesheet export`]);
  wsProj.mergeCells(`A${noteR.number}:K${noteR.number}`);
  xlStyleRow(noteR, { fill:XL.fill.subheader, font:XL.font.note, alignment:XL.align.li, height:16 });
  wsProj.addRow([]);
  xlHeader(wsProj, ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Var","Burn %","Budget £","Actual £","Variance £","Margin %","Status"]);
  budgeted.forEach(([proj,v], i) => {
    const burn = v.budget > 0 ? v.worked / v.budget : 0;
    const bGBP = v.budget * avgRate, aGBP = v.worked * avgRate, varGBP = bGBP - aGBP;
    const st = statusStyle(burn);
    const row = xlData(wsProj, [
      proj, v.customer,
      fmtHrs(v.budget), fmtHrs(v.worked), fmtHrsVar(v.worked - v.budget),
      fmtPct(burn), gbp(bGBP), gbp(aGBP), gbp(varGBP),
      fmtPct(bGBP > 0 ? varGBP/bGBP : 0), st.label,
    ], i);
    row.getCell(11).fill = st.fill; row.getCell(11).font = st.font;
    row.getCell(11).alignment = XL.align.center;
  });
  xlTotal(wsProj, [
    "TOTAL","",
    fmtHrs(totalBudgetHrs), fmtHrs(totalActualHrs), fmtHrsVar(totalActualHrs - totalBudgetHrs),
    fmtPct(totalBudgetHrs > 0 ? totalActualHrs/totalBudgetHrs : 0),
    gbp(totalBudgetGBP), gbp(totalActualGBP), gbp(totalVariance),
    fmtPct(overallMargin), "",
  ]);
  if (noBudget.length) {
    wsProj.addRow([]);
    xlSection(wsProj, "NO QUOTED BUDGET — hours logged, excluded from margin", 3);
    xlHeader(wsProj, ["Project","Customer","Actual Hrs"]);
    noBudget.forEach(([p,v], i) => xlData(wsProj, [p, v.customer, fmtHrs(v.worked)], i));
  }

  // ── Sheet 3: Utilisation ────────────────────────────────────────────────────
  const wsUtil = wb.addWorksheet("Utilisation");
  wsUtil.columns = [
    {width:26},{width:14},{width:15},{width:14},{width:14},{width:14},{width:20},{width:20},{width:20},
  ];
  xlTitle(wsUtil, `LABOUR UTILISATION  |  ${rangeLabel}`, 9);
  const uNote = wsUtil.addRow([`Available hours based on ${workingDays} working days × 8 hrs/day = ${workingDays*8} hrs per person`]);
  wsUtil.mergeCells(`A${uNote.number}:I${uNote.number}`);
  xlStyleRow(uNote, { fill:XL.fill.subheader, font:XL.font.note, alignment:XL.align.li, height:16 });
  wsUtil.addRow([]);
  xlHeader(wsUtil, ["Team Member","Dept","Rate (£/hr)","Available Hrs","Hours Worked","Util %","Capacity Cost","Labour Cost","Undeployed Cost"]);
  const utilZeroHrs = [];
  DEPARTMENTS.forEach(dept => {
    const members = team.filter(t => t.dept === dept);
    const active   = members.filter(t => (personHrs[t.name] || 0) > 0);
    const inactive = members.filter(t => (personHrs[t.name] || 0) === 0);
    inactive.forEach(t => utilZeroHrs.push(t));
    active.forEach((t, i) => {
      const rate   = parseFloat(t.rate) || 0;
      const worked = personHrs[t.name] || 0;
      const cap    = availHrs * rate;
      xlData(wsUtil, [
        t.name, dept, gbp(rate), availHrs, fmtHrs(worked),
        fmtPct(availHrs > 0 ? worked/availHrs : 0),
        gbp(cap), gbp(worked*rate), gbp(cap - worked*rate),
      ], i);
    });
    const gW = members.reduce((s,t) => s+(personHrs[t.name]||0), 0);
    const gC = members.reduce((s,t) => s+availHrs*(parseFloat(t.rate)||0), 0);
    const gD = members.reduce((s,t) => s+(personHrs[t.name]||0)*(parseFloat(t.rate)||0), 0);
    const gA = members.length * availHrs;
    xlSubtotal(wsUtil, [
      `— ${dept} Subtotal`, "", "", gA, fmtHrs(gW),
      fmtPct(gA > 0 ? gW/gA : 0), gbp(gC), gbp(gD), gbp(gC-gD),
    ]);
    wsUtil.addRow([]);
  });
  xlTotal(wsUtil, [
    "TOTAL","","", totalAvailHrs, fmtHrs(totalWorkedHrs), fmtPct(teamUtil),
    gbp(totalCapacity), gbp(totalDeployed), gbp(totalCapacity - totalDeployed),
  ]);
  if (utilZeroHrs.length) {
    wsUtil.addRow([]);
    xlSection(wsUtil, "NO HOURS LOGGED THIS PERIOD", 9);
    utilZeroHrs.forEach((t, i) => {
      const rate = parseFloat(t.rate) || 0;
      const cap  = availHrs * rate;
      xlData(wsUtil, [
        t.name, t.dept, gbp(rate), availHrs, 0,
        "0.0%", gbp(cap), gbp(0), gbp(cap),
      ], i);
    });
  }

  // ── Sheet 4: Raw Timesheet ──────────────────────────────────────────────────
  const wsTs = wb.addWorksheet("Raw Timesheet");
  wsTs.columns = [
    {width:12},{width:30},{width:42},{width:22},{width:12},{width:8},{width:9},{width:30},
  ];
  xlHeader(wsTs, ["Job No.","Task","Project","Team Member","Date","Hours","Billable","Notes"]);
  tsRows.forEach((r, i) => {
    const row = xlData(wsTs, [
      r.work_order_no||"", r.task_name||"",
      (r.project_description||"").trim(), r.user_name||"",
      (r.start_datetime||"").slice(0,10), parseFloat(r.hours)||0,
      r.is_billable==="1"?"Yes":"No", r.notes||"",
    ], i);
    row.getCell(6).numFmt = "0.0";
  });

  return wb;
}

function buildProjectViewReport({ tsRows, taskRows, team, rateMap, avgRate, startDate, endDate, workingDays, selectedProjects }) {
  const rangeLabel = `${fmtDate(startDate)} — ${fmtDate(endDate)}`;
  const projLabel  = selectedProjects.length === 1 ? selectedProjects[0] : `${selectedProjects.length} Projects`;

  const projBudget = {}, projCustomer = {};
  taskRows.forEach(r => {
    const p = (r.project||"").trim(); if (!p) return;
    if (!selectedProjects.includes(p)) return;
    const d = parseFloat(r.duration) || 0;
    if (d > 0) projBudget[p] = (projBudget[p]||0) + d;
    if (!projCustomer[p]) projCustomer[p] = r.customer_name||"";
  });

  // Per-project, per-person hours
  const projPersonHrs = {}, projActual = {};
  tsRows.forEach(r => {
    const p = (r.project_description||"").trim();
    const n = (r.user_name||"").trim();
    const h = parseFloat(r.hours)||0;
    if (!p || !selectedProjects.includes(p)) return;
    projActual[p]  = (projActual[p]||0) + h;
    if (!projPersonHrs[p]) projPersonHrs[p] = {};
    projPersonHrs[p][n] = (projPersonHrs[p][n]||0) + h;
    if (!projCustomer[p]) projCustomer[p] = "";
  });

  const wb = new ExcelJS.Workbook();
  wb.creator = "Labour Report v2"; wb.created = new Date();

  // ── Sheet 1: Project Summary ──────────────────────────────────────────────
  const wsSum = wb.addWorksheet("Project Summary");
  wsSum.columns = [
    {width:44},{width:28},{width:12},{width:12},{width:13},{width:10},{width:13},{width:14},{width:13},{width:11},{width:13},
  ];
  xlTitle(wsSum, `PROJECT VIEW — ${projLabel}  |  ${rangeLabel}`, 11);
  wsSum.addRow([]);
  xlHeader(wsSum, ["Project","Customer","Budget Hrs","Actual Hrs","Hrs Var","Burn %","Budget £","Actual £","Variance £","Margin %","Status"]);

  const sortedProjects = [...selectedProjects].sort((a, b) => {
    const burnA = (projBudget[a]||0) > 0 ? (projActual[a]||0)/(projBudget[a]||0) : 0;
    const burnB = (projBudget[b]||0) > 0 ? (projActual[b]||0)/(projBudget[b]||0) : 0;
    return burnB - burnA;
  });
  let totB=0, totA=0;
  sortedProjects.forEach((proj, i) => {
    const budget = projBudget[proj]||0;
    const worked = projActual[proj]||0;
    totB += budget; totA += worked;
    const burn = budget > 0 ? worked/budget : 0;
    const bGBP = budget*avgRate, aGBP = worked*avgRate;
    const st = statusStyle(burn);
    const row = xlData(wsSum, [
      proj, projCustomer[proj]||"",
      fmtHrs(budget), fmtHrs(worked), fmtHrsVar(worked-budget),
      fmtPct(burn), gbp(bGBP), gbp(aGBP), gbp(bGBP-aGBP),
      fmtPct(bGBP>0?(bGBP-aGBP)/bGBP:0), st.label,
    ], i);
    row.getCell(11).fill=st.fill; row.getCell(11).font=st.font;
    row.getCell(11).alignment=XL.align.center;
  });
  const tBGBP = totB*avgRate, tAGBP = totA*avgRate;
  xlTotal(wsSum, [
    "TOTAL","",fmtHrs(totB),fmtHrs(totA),fmtHrsVar(totA-totB),
    fmtPct(totB>0?totA/totB:0),gbp(tBGBP),gbp(tAGBP),gbp(tBGBP-tAGBP),
    fmtPct(tBGBP>0?(tBGBP-tAGBP)/tBGBP:0),"",
  ]);

  // ── Sheet 2: Hours by Person ──────────────────────────────────────────
  const wsHbP = wb.addWorksheet("Hours by Person");
  // Collect all people who logged time on selected projects
  const peopleSet = new Set();
  selectedProjects.forEach(p => { Object.keys(projPersonHrs[p]||{}).forEach(n => peopleSet.add(n)); });
  const people = [...peopleSet].sort();

  wsHbP.columns = [
    {width:26},{width:14},
    ...selectedProjects.map(() => ({width:14})),
    {width:14},{width:14},
  ];
  xlTitle(wsHbP, `HOURS BY PERSON — ${projLabel}  |  ${rangeLabel}`, 2 + selectedProjects.length + 2);
  wsHbP.addRow([]);
  xlHeader(wsHbP, ["Team Member","Dept",...selectedProjects,"Total Hrs","Total Cost £"]);

  people.forEach((name, i) => {
    const member = team.find(t => normName(t.name) === normName(name));
    const dept   = member?.dept || "—";
    const rate   = rateMap[name] || avgRate;
    const projHrs = selectedProjects.map(p => fmtHrs(projPersonHrs[p]?.[name]||0));
    const totalH  = projHrs.reduce((s,h) => s+h, 0);
    xlData(wsHbP, [name, dept, ...projHrs, fmtHrs(totalH), gbp(totalH*rate)], i);
  });

  // Project totals row
  const projTotHrs = selectedProjects.map(p => fmtHrs(projActual[p]||0));
  const grandTotal  = projTotHrs.reduce((s,h) => s+h, 0);
  xlTotal(wsHbP, ["TOTAL","", ...projTotHrs, fmtHrs(grandTotal), gbp(grandTotal*avgRate)]);
  // ── Sheet 3: Raw Timesheet (filtered) ────────────────────────────────────
  const wsTs = wb.addWorksheet("Timesheet Detail");
  wsTs.columns = [
    {width:12},{width:30},{width:42},{width:22},{width:12},{width:8},{width:9},{width:30},
  ];
  xlHeader(wsTs, ["Job No.","Task","Project","Team Member","Date","Hours","Billable","Notes"]);
  tsRows.forEach((r,i) => {
    const row = xlData(wsTs, [
      r.work_order_no||"", r.task_name||"",
      (r.project_description||"").trim(), r.user_name||"",
      (r.start_datetime||"").slice(0,10), parseFloat(r.hours)||0,
      r.is_billable==="1"?"Yes":"No", r.notes||"",
    ], i);
    row.getCell(6).numFmt = "0.0";
  });

  return wb;
}

function buildPersonViewReport({ tsRows, taskRows, team, rateMap, avgRate, startDate, endDate, workingDays, selectedPeople }) {
  const rangeLabel  = `${fmtDate(startDate)} — ${fmtDate(endDate)}`;
  const peopleLabel = selectedPeople.length === 1 ? selectedPeople[0] : `${selectedPeople.length} People`;
  const availHrs    = workingDays * 8;

  // Build project list from task rows
  const projBudget = {}, projCustomer = {};
  taskRows.forEach(r => {
    const p = (r.project||"").trim(); if (!p) return;
    const d = parseFloat(r.duration)||0;
    if (d > 0) projBudget[p] = (projBudget[p]||0) + d;
    if (!projCustomer[p]) projCustomer[p] = r.customer_name||"";
  });

  // Per-person, per-project hours
  const personProjHrs = {}, personTotalHrs = {};
  const allProjectsForPeople = new Set();
  tsRows.forEach(r => {
    const n = (r.user_name||"").trim();
    const p = (r.project_description||"").trim();
    const h = parseFloat(r.hours)||0;
    if (!selectedPeople.includes(n)) return;
    allProjectsForPeople.add(p);
    if (!personProjHrs[n]) personProjHrs[n] = {};
    personProjHrs[n][p] = (personProjHrs[n][p]||0) + h;
    personTotalHrs[n]   = (personTotalHrs[n]||0) + h;
    if (p && !projCustomer[p]) projCustomer[p] = "";
  });

  const projects = [...allProjectsForPeople].filter(Boolean).sort();

  const wb = new ExcelJS.Workbook();
  wb.creator = "Labour Report v2"; wb.created = new Date();

  // ── Sheet 1: Person Utilisation Summary ────────────────────────────────────
  const wsSum = wb.addWorksheet("Person Summary");
  wsSum.columns = [
    {width:26},{width:14},{width:15},{width:14},{width:14},{width:14},{width:20},{width:20},{width:20},
  ];
  xlTitle(wsSum, `PERSON VIEW — ${peopleLabel}  |  ${rangeLabel}`, 9);
  wsSum.addRow([]);
  xlHeader(wsSum, ["Name","Dept","Rate (£/hr)","Available Hrs","Hours Worked","Util %","Capacity Cost","Labour Cost","Undeployed Cost"]);

  const sumZeroHrs = [];
  DEPARTMENTS.forEach(dept => {
    const members = team.filter(t => t.dept === dept && selectedPeople.includes(t.name));
    if (!members.length) return;
    const active   = members.filter(t => (personTotalHrs[t.name]||0) > 0);
    const inactive = members.filter(t => (personTotalHrs[t.name]||0) === 0);
    inactive.forEach(t => sumZeroHrs.push(t));
    active.forEach((t, i) => {
      const rate   = parseFloat(t.rate)||0;
      const worked = personTotalHrs[t.name]||0;
      const cap    = availHrs*rate;
      xlData(wsSum, [
        t.name, dept, gbp(rate), availHrs, fmtHrs(worked),
        fmtPct(availHrs>0?worked/availHrs:0),
        gbp(cap), gbp(worked*rate), gbp(cap-worked*rate),
      ], i);
    });
  });

  const totW = selectedPeople.reduce((s,n) => s+(personTotalHrs[n]||0), 0);
  const totC = selectedPeople.reduce((s,n) => {
    const m = team.find(t => normName(t.name)===normName(n));
    return s + (personTotalHrs[n]||0)*(m?parseFloat(m.rate)||avgRate:avgRate);
  }, 0);
  const totCap = selectedPeople.reduce((s,n) => {
    const m = team.find(t => normName(t.name)===normName(n));
    return s + availHrs*(m?parseFloat(m.rate)||avgRate:avgRate);
  }, 0);
  xlTotal(wsSum, [
    "TOTAL","","",selectedPeople.length*availHrs,fmtHrs(totW),
    fmtPct(selectedPeople.length*availHrs>0?totW/(selectedPeople.length*availHrs):0),
    gbp(totCap),gbp(totC),gbp(totCap-totC),
  ]);
  if (sumZeroHrs.length) {
    wsSum.addRow([]);
    xlSection(wsSum, "NO HOURS LOGGED THIS PERIOD", 9);
    sumZeroHrs.forEach((t, i) => {
      const rate = parseFloat(t.rate)||0;
      const cap  = availHrs*rate;
      xlData(wsSum, [
        t.name, t.dept, gbp(rate), availHrs, 0,
        "0.0%", gbp(cap), gbp(0), gbp(cap),
      ], i);
    });
  }

  // ── Sheet 2: Hours by Project (people as cols) ───────────────────────────
  const wsHbP = wb.addWorksheet("Hours by Project");
  wsHbP.columns = [
    {width:44},{width:28},
    ...selectedPeople.map(() => ({width:14})),
    {width:14},{width:14},
  ];
  xlTitle(wsHbP, `HOURS BY PROJECT — ${peopleLabel}  |  ${rangeLabel}`, 2+selectedPeople.length+2);
  wsHbP.addRow([]);
  xlHeader(wsHbP, ["Project","Customer",...selectedPeople,"Total Hrs","Budget Hrs"]);

  projects.forEach((proj, i) => {
    const personHrs = selectedPeople.map(n => fmtHrs(personProjHrs[n]?.[proj]||0));
    const totalH    = personHrs.reduce((s,h) => s+h, 0);
    xlData(wsHbP, [proj, projCustomer[proj]||"", ...personHrs, fmtHrs(totalH), fmtHrs(projBudget[proj]||0)], i);
  });

  const personTotals = selectedPeople.map(n => fmtHrs(personTotalHrs[n]||0));
  const grandTot     = personTotals.reduce((s,h)=>s+h,0);
  xlTotal(wsHbP, ["TOTAL","", ...personTotals, fmtHrs(grandTot),""]);

  // ── Sheet 3: Day-by-day breakdown ────────────────────────────────────────
  const wsDbd = wb.addWorksheet("Daily Breakdown");
  wsDbd.columns = [
    {width:14},{width:26},{width:14},{width:42},{width:8},{width:9},{width:30},
  ];
  xlHeader(wsDbd, ["Date","Team Member","Dept","Project","Hours","Billable","Notes"]);
  const sorted = [...tsRows].sort((a,b) =>
    (a.start_datetime||"").localeCompare(b.start_datetime||"")
  );
  sorted.forEach((r,i) => {
    const name = (r.user_name||"").trim();
    const m    = team.find(t => normName(t.name)===normName(name));
    const dRow = xlData(wsDbd, [
      (r.start_datetime||"").slice(0,10), name, m?.dept||"—",
      (r.project_description||"").trim(), parseFloat(r.hours)||0,
      r.is_billable==="1"?"Yes":"No", r.notes||"",
    ], i);
    dRow.getCell(5).numFmt = "0.0";
  });

  return wb;
}

// ─── UI Components ─────────────────────────────────────────────────────────────

function DropZone({ label, sublabel, accept, onFile, file, icon }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); const f=e.dataTransfer.files[0]; if(f) onFile(f); }}
      style={{
        border: `2px dashed ${file?"#10b981":drag?"#3b82f6":"#1e3a5f"}`,
        borderRadius: 10, padding:"24px 16px", cursor:"pointer", textAlign:"center",
        background: file?"rgba(16,185,129,0.06)":drag?"rgba(59,130,246,0.06)":"rgba(10,15,30,0.5)",
        transition:"all 0.2s",
      }}
    >
      <input ref={ref} type="file" accept={accept} style={{display:"none"}}
        onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{fontSize:28, marginBottom:8}}>{file?"✅":icon}</div>
      <div style={{fontSize:13, color:file?"#10b981":"#cbd5e1", fontWeight:600, marginBottom:4}}>
        {file ? file.name : label}
      </div>
      <div style={{fontSize:11, color:"#4b6280", fontFamily:"monospace"}}>
        {file ? `${(file.size/1024).toFixed(1)} KB` : sublabel}
      </div>
    </div>
  );
}

function MultiSelect({ options, selected, onChange, placeholder, maxHeight = 200 }) {
  const [open, setOpen] = useState(false);
  const ref = useRef();
  useEffect(() => {
    const handler = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const toggle = v => onChange(selected.includes(v) ? selected.filter(x=>x!==v) : [...selected, v]);

  return (
    <div ref={ref} style={{position:"relative"}}>
      <div
        onClick={() => setOpen(!open)}
        style={{
          background:"rgba(10,20,40,0.8)", border:"1px solid #1e3a5f",
          borderRadius:8, padding:"10px 14px", cursor:"pointer",
          display:"flex", justifyContent:"space-between", alignItems:"center",
          color: selected.length ? "#e2e8f0" : "#4b6280", fontSize:13,
          minHeight:42,
        }}
      >
        <span style={{flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>
          {selected.length === 0 ? placeholder
           : selected.length === 1 ? selected[0]
           : `${selected.length} selected`}
        </span>
        <span style={{color:"#4b6280", marginLeft:8, fontSize:10}}>{open?"▲":"▼"}</span>
      </div>
      {open && (
        <div style={{
          position:"absolute", top:"calc(100% + 4px)", left:0, right:0, zIndex:9999,
          background:"#0a1020", border:"1px solid #1e3a5f", borderRadius:8,
          maxHeight, overflowY:"auto", boxShadow:"0 8px 32px rgba(0,0,0,0.5)",
        }}>
          {options.length === 0 && (
            <div style={{padding:"10px 14px", color:"#4b6280", fontSize:12}}>No options available</div>
          )}
          {options.map(opt => (
            <div
              key={opt}
              onClick={() => toggle(opt)}
              style={{
                padding:"8px 14px", cursor:"pointer", fontSize:12,
                color: selected.includes(opt) ? "#60a5fa" : "#cbd5e1",
                background: selected.includes(opt) ? "rgba(59,130,246,0.1)" : "transparent",
                display:"flex", alignItems:"center", gap:8,
              }}
            >
              <span style={{
                width:14, height:14, borderRadius:3,
                border: selected.includes(opt) ? "none" : "1px solid #2d4a6e",
                background: selected.includes(opt) ? "#3b82f6" : "transparent",
                display:"inline-flex", alignItems:"center", justifyContent:"center",
                fontSize:9, flexShrink:0,
              }}>
                {selected.includes(opt) && "✓"}
              </span>
              {opt}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tsFile,    setTsFile]    = useState(null);
  const [taskFile,  setTaskFile]  = useState(null);
  const [tsRows,    setTsRows]    = useState([]);
  const [taskRows,  setTaskRows]  = useState([]);

  // Date range
  const [startDate, setStartDate] = useState(() => {
    const d = new Date(); d.setDate(1); return d.toISOString().slice(0,10);
  });
  const [endDate, setEndDate] = useState(() => new Date().toISOString().slice(0,10));
  const [workingDays, setWorkingDays] = useState(20);

  // Department rates
  const [deptRates, setDeptRates] = useState({ Design:50, Engineering:50, Presales:50, Projects:50, Service:50 });

  // Team with individual overrides
  const [team, setTeam] = useState(TEAM_DEFAULT.map(t => ({ ...t, override: null })));

  // Name mapping
  const [nameMap, setNameMap] = useState({});
  const [showMapping, setShowMapping] = useState(false);
  const [showAllMappings, setShowAllMappings] = useState(false);

  // Report type & filters
  const [reportType,      setReportType]      = useState("all");
  const [selectedProjects, setSelectedProjects] = useState([]);
  const [selectedPeople,   setSelectedPeople]   = useState([]);

  // Status
  const [status,   setStatus]   = useState("idle");
  const [errorMsg, setErrorMsg] = useState("");

  // Parse timesheet
  useEffect(() => {
    if (!tsFile) { setTsRows([]); return; }
    let cancelled = false;
    readFile(tsFile).then(text => { if (!cancelled) setTsRows(parseCSV(text)); });
    return () => { cancelled = true; };
  }, [tsFile]);

  // Parse task list
  useEffect(() => {
    if (!taskFile) { setTaskRows([]); return; }
    let cancelled = false;
    readFile(taskFile).then(text => { if (!cancelled) setTaskRows(parseCSV(text)); });
    return () => { cancelled = true; };
  }, [taskFile]);

  // Auto-calculate working days when date range changes
  useEffect(() => {
    if (startDate && endDate && endDate >= startDate) {
      setWorkingDays(calcWorkingDays(startDate, endDate));
    }
  }, [startDate, endDate]);

  // Apply dept rate to all team members in that dept (unless overridden)
  const handleDeptRate = (dept, val) => {
    setDeptRates(prev => ({...prev, [dept]: val}));
    setTeam(prev => prev.map(t => t.dept === dept && t.override === null ? {...t, rate: parseFloat(val)||0} : t));
  };

  const handleOverride = (i, val) => {
    setTeam(prev => prev.map((t, idx) =>
      idx === i ? {...t, rate: parseFloat(val)||0, override: val} : t
    ));
  };

  const clearOverride = (i) => {
    const dept = team[i].dept;
    setTeam(prev => prev.map((t, idx) =>
      idx === i ? {...t, rate: deptRates[dept]||50, override: null} : t
    ));
  };

  // Build rate map with name normalisation and user mapping
  const rateMap = useMemo(() => {
    const map = {};
    team.forEach(t => {
      const r = parseFloat(t.rate) || 0;
      map[t.name.trim()]           = r;
      map[normName(t.name)]        = r;
    });
    // Apply name mappings
    Object.entries(nameMap).forEach(([csvName, teamName]) => {
      if (teamName && teamName !== "__ignore__") {
        const member = team.find(t => t.name === teamName);
        if (member) {
          map[csvName]         = parseFloat(member.rate) || 0;
          map[normName(csvName)] = parseFloat(member.rate) || 0;
        }
      }
    });
    return map;
  }, [team, nameMap]);

  const avgRate = useMemo(() =>
    team.reduce((s,t) => s+(parseFloat(t.rate)||0), 0) / team.length,
    [team]
  );

  // Resolve user name from CSV using nameMap + fuzzy matching
  const resolveUser = useCallback((csvName) => {
    const trimmed = csvName.trim();
    if (nameMap[trimmed]) return nameMap[trimmed] === "__ignore__" ? null : nameMap[trimmed];
    // Exact match
    const exact = team.find(t => t.name.trim() === trimmed);
    if (exact) return exact.name;
    // Normalised match
    const norm = team.find(t => normName(t.name) === normName(trimmed));
    if (norm) return norm.name;
    return trimmed; // return as-is
  }, [nameMap, team]);

  // Apply name resolution to timesheet rows
  const resolvedTsRows = useMemo(() => {
    return tsRows.map(r => ({
      ...r,
      user_name: resolveUser((r.user_name||"").trim()) || (r.user_name||"").trim(),
    })).filter(r => {
      const n = resolveUser((r.user_name||"").trim());
      return n !== null;
    });
  }, [tsRows, resolveUser]);

  // Detect unmatched names
  const unmatchedNames = useMemo(() => {
    const teamNorms = new Set(team.map(t => normName(t.name)));
    const csvNames  = [...new Set(tsRows.map(r => (r.user_name||"").trim()).filter(Boolean))];
    return csvNames.filter(n => {
      if (nameMap[n]) return false;
      if (teamNorms.has(normName(n))) return false;
      return true;
    });
  }, [tsRows, team, nameMap]);

  // Filter rows by date range
  const filteredTsRows = useMemo(() => {
    const sD = new Date(startDate + "T00:00:00");
    const eD = new Date(endDate   + "T23:59:59");
    return resolvedTsRows.filter(r => {
      const ds = (r.start_datetime||"").slice(0,10);
      if (!ds) return false;
      const d = new Date(ds + "T00:00:00");
      return d >= sD && d <= eD;
    });
  }, [resolvedTsRows, startDate, endDate]);

  // Unique projects & people from filtered rows
  const uniqueProjects = useMemo(() =>
    [...new Set(filteredTsRows.map(r => (r.project_description||"").trim()).filter(Boolean))].sort(),
    [filteredTsRows]
  );

  const uniquePeople = useMemo(() =>
    [...new Set(filteredTsRows.map(r => (r.user_name||"").trim()).filter(Boolean))].sort(),
    [filteredTsRows]
  );

  // Live preview stats (updates with report type / filter selections)
  const previewStats = useMemo(() => {
    if (!filteredTsRows.length) return null;
    const rows = (reportType === "project" && selectedProjects.length)
      ? filteredTsRows.filter(r => selectedProjects.includes((r.project_description||"").trim()))
      : (reportType === "person" && selectedPeople.length)
      ? filteredTsRows.filter(r => selectedPeople.includes((r.user_name||"").trim()))
      : filteredTsRows;
    return {
      projCount: new Set(rows.map(r => (r.project_description||"").trim()).filter(Boolean)).size,
      pplCount:  new Set(rows.map(r => (r.user_name||"").trim()).filter(Boolean)).size,
      totalHrs:  rows.reduce((s,r) => s+(parseFloat(r.hours)||0), 0).toFixed(1),
      entries:   rows.length,
    };
  }, [filteredTsRows, reportType, selectedProjects, selectedPeople]);

  // Date range display
  const dateRangeLabel = useMemo(() => {
    if (!startDate || !endDate) return "";
    const s = new Date(startDate+"T00:00:00"), e = new Date(endDate+"T00:00:00");
    const days = Math.round((e-s)/(1000*60*60*24)) + 1;
    return `${fmtDate(s)} — ${fmtDate(e)}  (${days} calendar days)`;
  }, [startDate, endDate]);

  const tsMissingCols = useMemo(() => {
    if (!tsRows.length) return [];
    const cols = Object.keys(tsRows[0]);
    return ['user_name','project_description','hours','start_datetime'].filter(c => !cols.includes(c));
  }, [tsRows]);

  const taskMissingCols = useMemo(() => {
    if (!taskRows.length) return [];
    const cols = Object.keys(taskRows[0]);
    return ['project','duration'].filter(c => !cols.includes(c));
  }, [taskRows]);

  const ready       = tsFile && taskFile;
  const csvInvalid  = tsMissingCols.length > 0 || taskMissingCols.length > 0;
  const noEntries   = ready && tsRows.length > 0 && filteredTsRows.length === 0;
  const canGenerate = ready && !csvInvalid && !noEntries;

  const handleGenerate = async () => {
    if (!canGenerate) return;
    setStatus("generating"); setErrorMsg("");
    try {
      const sDate = new Date(startDate+"T00:00:00");
      const eDate = new Date(endDate+"T23:59:59");

      // Apply project/person filter
      let finalTsRows  = filteredTsRows;
      let finalTaskRows = taskRows;

      if (reportType === "project" && selectedProjects.length > 0) {
        finalTsRows   = filteredTsRows.filter(r => selectedProjects.includes((r.project_description||"").trim()));
        finalTaskRows = taskRows.filter(r => selectedProjects.includes((r.project||"").trim()));
      }
      if (reportType === "person" && selectedPeople.length > 0) {
        finalTsRows = filteredTsRows.filter(r => selectedPeople.includes((r.user_name||"").trim()));
      }

      const params = { tsRows:finalTsRows, taskRows:finalTaskRows, team, rateMap, avgRate, startDate:sDate, endDate:eDate, workingDays };

      let wb;
      if (reportType === "all") {
        wb = buildAllProjectsReport(params);
      } else if (reportType === "project") {
        wb = buildProjectViewReport({ ...params, selectedProjects: selectedProjects.length ? selectedProjects : uniqueProjects });
      } else {
        wb = buildPersonViewReport({ ...params, selectedPeople: selectedPeople.length ? selectedPeople : uniquePeople });
      }


      const sanitize = s => s.replace(/[\/\\:*?"<>|]/g, "").replace(/\s+/g, "_");
      let filename;
      if (reportType === "all") {
        filename = `Labour_Report_AllProjects_${startDate}_to_${endDate}.xlsx`;
      } else if (reportType === "project") {
        const eff = selectedProjects.length ? selectedProjects : uniqueProjects;
        if (eff.length === 1) {
          filename = `Labour_Report_${sanitize(eff[0]).slice(0, 30)}_${startDate}.xlsx`;
        } else if (eff.length === 2) {
          filename = `Labour_Report_${sanitize(eff[0]).slice(0, 15)}_${sanitize(eff[1]).slice(0, 15)}_${startDate}.xlsx`;
        } else {
          filename = `Labour_Report_ProjectView_${eff.length}Projects_${startDate}.xlsx`;
        }
      } else {
        const eff = selectedPeople.length ? selectedPeople : uniquePeople;
        if (eff.length === 1) {
          filename = `Labour_Report_${sanitize(eff[0])}_${startDate}.xlsx`;
        } else {
          filename = `Labour_Report_PersonView_${eff.length}People_${startDate}.xlsx`;
        }
      }

      const buf  = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url  = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url; link.download = filename;
      document.body.appendChild(link); link.click();
      document.body.removeChild(link); URL.revokeObjectURL(url);
      setStatus("done");
    } catch(e) {
      setStatus("error"); setErrorMsg(e.message);
    }
  };

  // ─── Render ───────────────────────────────────────────────────────────────
  return (
    <div style={{
      minHeight:"100vh",
      background:"#050d1a",
      backgroundImage:"radial-gradient(ellipse 70% 40% at 50% 0%, rgba(29,78,216,0.12), transparent)",
      fontFamily:"'IBM Plex Sans', 'Segoe UI', sans-serif",
      color:"#e2e8f0",
      padding:"0 0 80px",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&family=Libre+Baskerville:ital,wght@0,700;1,400&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        ::-webkit-scrollbar { width:6px; } ::-webkit-scrollbar-track { background:#0a1525; } ::-webkit-scrollbar-thumb { background:#1e3a5f; border-radius:3px; }
        .card { background:rgba(10,20,40,0.7); border:1px solid #0f2240; border-radius:12px; padding:24px; backdrop-filter:blur(8px); }
        .section-tag { font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:2.5px; text-transform:uppercase; color:#3b82f6; margin-bottom:14px; }
        .field { background:rgba(10,20,40,0.9); border:1px solid #1e3a5f; border-radius:7px; color:#e2e8f0; font-family:'IBM Plex Mono',monospace; font-size:13px; padding:9px 12px; outline:none; transition:border-color 0.15s; }
        .field:focus { border-color:#3b82f6; }
        .field-label { font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:1px; color:#4b6280; text-transform:uppercase; margin-bottom:6px; }
        .rate-input { background:rgba(10,20,40,0.9); border:1px solid #1e3a5f; border-radius:6px; color:#93c5fd; font-family:'IBM Plex Mono',monospace; font-size:13px; padding:6px 9px; width:80px; outline:none; transition:border-color 0.15s; }
        .rate-input:focus { border-color:#3b82f6; }
        .rate-input.overridden { border-color:#f59e0b; color:#fbbf24; }
        .btn-primary { background:linear-gradient(135deg,#1d4ed8,#1e40af); border:none; border-radius:10px; color:white; cursor:pointer; font-family:'IBM Plex Sans',sans-serif; font-size:15px; font-weight:600; padding:14px 36px; transition:all 0.2s; box-shadow:0 4px 20px rgba(29,78,216,0.3); letter-spacing:0.2px; }
        .btn-primary:hover:not(:disabled) { transform:translateY(-2px); box-shadow:0 6px 28px rgba(29,78,216,0.45); }
        .btn-primary:disabled { opacity:0.35; cursor:not-allowed; transform:none; }
        .btn-secondary { background:transparent; border:1px solid #1e3a5f; border-radius:8px; color:#93c5fd; cursor:pointer; font-family:'IBM Plex Sans',sans-serif; font-size:12px; padding:7px 14px; transition:all 0.15s; }
        .btn-secondary:hover { border-color:#3b82f6; color:#60a5fa; }
        .report-type-btn { background:transparent; border:1px solid #0f2240; border-radius:10px; cursor:pointer; padding:16px; text-align:left; transition:all 0.2s; width:100%; }
        .report-type-btn:hover { border-color:#1e3a5f; background:rgba(29,78,216,0.05); }
        .report-type-btn.active { border-color:#3b82f6; background:rgba(29,78,216,0.1); }
        input[type=date].field { color-scheme:dark; }
        input[type=number]::-webkit-inner-spin-button { opacity:1; }
      `}</style>

      {/* Header */}
      <div style={{textAlign:"center", padding:"52px 24px 36px"}}>
        <div style={{
          display:"inline-flex", alignItems:"center", gap:8,
          background:"rgba(29,78,216,0.08)", border:"1px solid rgba(29,78,216,0.2)",
          borderRadius:20, padding:"5px 16px", marginBottom:20,
        }}>
          <span style={{width:6,height:6,borderRadius:"50%",background:"#3b82f6",display:"inline-block"}}/>
          <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:10,color:"#60a5fa",letterSpacing:2}}>
            LABOUR REPORT v2
          </span>
        </div>
        <h1 style={{
          fontFamily:"'Libre Baskerville',serif", fontSize:"clamp(28px,4.5vw,46px)",
          fontWeight:700, lineHeight:1.2, color:"#e2e8f0", marginBottom:12,
        }}>
          Project Profitability<br />
          <span style={{fontStyle:"italic",fontWeight:400,color:"#60a5fa"}}>&amp; Labour Utilisation</span>
        </h1>
        <p style={{color:"#4b6280",fontSize:14,maxWidth:440,margin:"0 auto"}}>
          WeQuote CSV exports → formatted Excel report. Day-level filtering, team cost rates, flexible report views.
        </p>
      </div>

      <div style={{maxWidth:820,margin:"0 auto",padding:"0 20px",display:"flex",flexDirection:"column",gap:16}}>

        {/* Step 1 – Upload */}
        <div className="card">
          <div className="section-tag">01 — Upload CSV Exports</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <DropZone label="Timesheet Entries" sublabel="Timesheet_Entries_YYYYMMDD.csv" accept=".csv" onFile={setTsFile} file={tsFile} icon="⏱" />
            <DropZone label="Task List"          sublabel="task-list.csv"                  accept=".csv" onFile={setTaskFile} file={taskFile} icon="📋" />
          </div>
          {tsRows.length > 0 && (
            <div style={{marginTop:12,fontSize:12,color:"#4b6280",fontFamily:"'IBM Plex Mono',monospace"}}>
              {tsRows.length} timesheet entries loaded · {[...new Set(tsRows.map(r=>(r.project_description||"").trim()).filter(Boolean))].length} projects · {[...new Set(tsRows.map(r=>(r.user_name||"").trim()).filter(Boolean))].length} people
            </div>
          )}
          {tsFile && (
            <div style={{marginTop:12}}>
              <button className="btn-secondary" style={{fontSize:11,display:"inline-flex",alignItems:"center",gap:8}}
                onClick={() => setShowAllMappings(v => !v)}>
                {showAllMappings ? "Hide Name Mappings" : "Review Name Mappings"}
                {unmatchedNames.length > 0 && (
                  <span style={{
                    background:"rgba(245,158,11,0.15)",border:"1px solid rgba(245,158,11,0.3)",
                    borderRadius:10,padding:"1px 7px",fontSize:10,color:"#f59e0b",
                  }}>{unmatchedNames.length} unmatched</span>
                )}
              </button>
              {showAllMappings && (
                <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:6}}>
                  {[...new Set(tsRows.map(r => (r.user_name||"").trim()).filter(Boolean))].map(csvName => {
                    const explicit   = nameMap[csvName];
                    const autoMatch  = !explicit && team.find(t => t.name.trim() === csvName || normName(t.name) === normName(csvName));
                    const isUnmatched = !explicit && !autoMatch;
                    const selectVal  = explicit || (autoMatch ? autoMatch.name : "");
                    const statusPill = explicit === "__ignore__"
                      ? { label:"Ignored",      color:"#f59e0b", bg:"rgba(245,158,11,0.1)"  }
                      : explicit
                      ? { label:"Mapped",        color:"#60a5fa", bg:"rgba(59,130,246,0.1)"  }
                      : autoMatch
                      ? { label:"Auto-matched",  color:"#10b981", bg:"rgba(16,185,129,0.1)"  }
                      : { label:"Unmatched",     color:"#f87171", bg:"rgba(239,68,68,0.1)"   };
                    return (
                      <div key={csvName} style={{
                        display:"flex",alignItems:"center",gap:10,
                        background: isUnmatched ? "rgba(245,158,11,0.04)" : "rgba(10,20,40,0.4)",
                        border:`1px solid ${isUnmatched ? "rgba(245,158,11,0.15)" : "#0f2240"}`,
                        borderRadius:8,padding:"7px 12px",
                      }}>
                        <div style={{flex:"0 0 180px",fontFamily:"'IBM Plex Mono',monospace",fontSize:12,
                          color:isUnmatched?"#fbbf24":"#cbd5e1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                          {csvName}
                        </div>
                        <span style={{
                          flex:"0 0 auto",fontFamily:"'IBM Plex Mono',monospace",fontSize:10,
                          background:statusPill.bg,borderRadius:4,padding:"2px 7px",color:statusPill.color,
                        }}>{statusPill.label}</span>
                        <select
                          className="field"
                          style={{flex:1,fontSize:12,padding:"5px 10px"}}
                          value={selectVal}
                          onChange={e => {
                            const val = e.target.value;
                            if (!val) setNameMap(prev => { const n={...prev}; delete n[csvName]; return n; });
                            else setNameMap(prev => ({...prev,[csvName]:val}));
                          }}
                        >
                          <option value="">{autoMatch ? `Auto-match: ${autoMatch.name}` : "— Select team member…"}</option>
                          {team.map(t => <option key={t.name} value={t.name}>{t.name} ({t.dept})</option>)}
                          <option value="__ignore__">— Ignore this person</option>
                        </select>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}
        </div>

        {/* CSV Validation Warning */}
        {(tsMissingCols.length > 0 || taskMissingCols.length > 0) && (
          <div className="card" style={{border:"1px solid rgba(245,158,11,0.25)"}}>
            <div className="section-tag" style={{color:"#f59e0b",marginBottom:8}}>⚠ Missing Required Columns</div>
            {tsMissingCols.length > 0 && (
              <div style={{marginBottom: taskMissingCols.length > 0 ? 10 : 0}}>
                <div style={{fontSize:12,color:"#92400e",marginBottom:6}}>Timesheet CSV is missing:</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                  {tsMissingCols.map(c => (
                    <span key={c} style={{
                      fontFamily:"'IBM Plex Mono',monospace",fontSize:11,
                      background:"rgba(245,158,11,0.08)",border:"1px solid rgba(245,158,11,0.25)",
                      borderRadius:4,padding:"2px 8px",color:"#fbbf24",
                    }}>{c}</span>
                  ))}
                </div>
              </div>
            )}
            {taskMissingCols.length > 0 && (
              <div>
                <div style={{fontSize:12,color:"#92400e",marginBottom:6}}>Task List CSV is missing:</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                  {taskMissingCols.map(c => (
                    <span key={c} style={{
                      fontFamily:"'IBM Plex Mono',monospace",fontSize:11,
                      background:"rgba(245,158,11,0.08)",border:"1px solid rgba(245,158,11,0.25)",
                      borderRadius:4,padding:"2px 8px",color:"#fbbf24",
                    }}>{c}</span>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Name Mapping – shown if unmatched names */}
        {unmatchedNames.length > 0 && (
          <div className="card" style={{border:"1px solid rgba(245,158,11,0.25)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div>
                <div className="section-tag" style={{color:"#f59e0b",marginBottom:4}}>⚠ Name Mapping Required</div>
                <div style={{fontSize:12,color:"#92400e"}}>
                  {unmatchedNames.length} name{unmatchedNames.length>1?"s":""} in the timesheet CSV {unmatchedNames.length>1?"don't":"doesn't"} match the team roster.
                  Map each to a team member, or mark as ignore.
                </div>
              </div>
              <button className="btn-secondary" onClick={() => setShowMapping(!showMapping)}>
                {showMapping ? "Hide" : "Review"}
              </button>
            </div>
            {showMapping && (
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {unmatchedNames.map(csvName => (
                  <div key={csvName} style={{
                    display:"flex",alignItems:"center",gap:12,
                    background:"rgba(245,158,11,0.04)",border:"1px solid rgba(245,158,11,0.15)",
                    borderRadius:8,padding:"10px 14px",
                  }}>
                    <div style={{flex:1,fontFamily:"'IBM Plex Mono',monospace",fontSize:12,color:"#fbbf24"}}>
                      "{csvName}"
                    </div>
                    <span style={{color:"#4b6280",fontSize:12}}>→</span>
                    <select
                      className="field"
                      style={{flex:1}}
                      value={nameMap[csvName]||""}
                      onChange={e => setNameMap(prev => ({...prev,[csvName]:e.target.value}))}
                    >
                      <option value="">Select team member…</option>
                      {team.map(t => <option key={t.name} value={t.name}>{t.name} ({t.dept})</option>)}
                      <option value="__ignore__">— Ignore this person</option>
                    </select>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Step 2 – Date Range */}
        <div className="card">
          <div className="section-tag">02 — Report Period</div>
          <div style={{display:"flex",gap:20,flexWrap:"wrap",alignItems:"flex-end"}}>
            <div>
              <div className="field-label">Start Date</div>
              <input type="date" className="field" value={startDate} onChange={e => setStartDate(e.target.value)} />
            </div>
            <div style={{color:"#1e3a5f",fontSize:18,paddingBottom:8}}>→</div>
            <div>
              <div className="field-label">End Date</div>
              <input type="date" className="field" value={endDate} onChange={e => setEndDate(e.target.value)} />
            </div>
            <div>
              <div className="field-label">Working Days</div>
              <input type="number" className="field" value={workingDays} style={{width:90}}
                onChange={e => setWorkingDays(parseInt(e.target.value)||20)} min={1} max={31} />
              <div style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:10,color:"#4b6280",marginTop:4}}>auto-calculated</div>
            </div>
          </div>
          {dateRangeLabel && (
            <div style={{
              marginTop:14,padding:"10px 14px",
              background:"rgba(29,78,216,0.06)",border:"1px solid rgba(29,78,216,0.15)",
              borderRadius:8,fontFamily:"'IBM Plex Mono',monospace",fontSize:12,color:"#60a5fa",
              display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,
            }}>
              <span>{dateRangeLabel}</span>
              <span style={{color:"#1e3a5f"}}>·</span>
              <span style={{color:"#4b6280"}}>{workingDays} working days × 8 hrs = {workingDays*8} hrs/person available</span>
            </div>
          )}
          {filteredTsRows.length > 0 && (
            <div style={{marginTop:8,fontSize:11,color:"#4b6280",fontFamily:"'IBM Plex Mono',monospace"}}>
              {filteredTsRows.length} entries in range · {filteredTsRows.reduce((s,r)=>s+(parseFloat(r.hours)||0),0).toFixed(1)} hrs total
            </div>
          )}
          {noEntries && (
            <div style={{
              marginTop:10,padding:"10px 14px",
              background:"rgba(245,158,11,0.06)",border:"1px solid rgba(245,158,11,0.25)",
              borderRadius:8,fontFamily:"'IBM Plex Mono',monospace",fontSize:12,color:"#fbbf24",
            }}>
              ⚠ No timesheet entries found in this date range
            </div>
          )}
        </div>

        {/* Step 3 – Report Type */}
        <div className="card" style={{overflow:"visible",backdropFilter:"none",WebkitBackdropFilter:"none",position:"relative",zIndex:2}}>
          <div className="section-tag">03 — Report Type</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
            {REPORT_TYPES.map(rt => (
              <button
                key={rt.id}
                className={`report-type-btn ${reportType===rt.id?"active":""}`}
                onClick={() => setReportType(rt.id)}
              >
                <div style={{
                  fontSize:18,marginBottom:6,
                  color: reportType===rt.id?"#60a5fa":"#2d4a6e",
                }}>{rt.icon}</div>
                <div style={{fontSize:13,fontWeight:600,color:reportType===rt.id?"#e2e8f0":"#64748b",marginBottom:4}}>
                  {rt.label}
                </div>
                <div style={{fontSize:11,color:"#4b6280",lineHeight:1.4}}>{rt.desc}</div>
              </button>
            ))}
          </div>

          {/* Project filter */}
          {reportType === "project" && (
            <div>
              <div className="field-label">Select Projects</div>
              <MultiSelect
                options={uniqueProjects}
                selected={selectedProjects}
                onChange={setSelectedProjects}
                placeholder="Select one or more projects…"
              />
              {selectedProjects.length > 0 && (
                <div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:6}}>
                  {selectedProjects.map(p => (
                    <span key={p} style={{
                      background:"rgba(29,78,216,0.12)",border:"1px solid rgba(29,78,216,0.25)",
                      borderRadius:20,padding:"3px 10px",fontSize:11,color:"#60a5fa",
                      display:"flex",alignItems:"center",gap:6,cursor:"pointer",
                    }} onClick={() => setSelectedProjects(prev => prev.filter(x=>x!==p))}>
                      {p} <span style={{color:"#3b5475",fontSize:10}}>✕</span>
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Person filter */}
          {reportType === "person" && (
            <div>
              <div className="field-label">Select People</div>
              <MultiSelect
                options={uniquePeople}
                selected={selectedPeople}
                onChange={setSelectedPeople}
                placeholder="Select one or more team members…"
              />
              {selectedPeople.length > 0 && (
                <div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:6}}>
                  {selectedPeople.map(p => (
                    <span key={p} style={{
                      background:"rgba(29,78,216,0.12)",border:"1px solid rgba(29,78,216,0.25)",
                      borderRadius:20,padding:"3px 10px",fontSize:11,color:"#60a5fa",
                      display:"flex",alignItems:"center",gap:6,cursor:"pointer",
                    }} onClick={() => setSelectedPeople(prev => prev.filter(x=>x!==p))}>
                      {p} <span style={{color:"#3b5475",fontSize:10}}>✕</span>
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Step 4 – Rates */}
        <div className="card">
          <div className="section-tag">04 — Cost Rates (£/hr)</div>

          {/* Dept defaults */}
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:"#4b6280",marginBottom:10,fontFamily:"'IBM Plex Mono',monospace"}}>
              Department defaults — applied to all members unless individually overridden
            </div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              {DEPARTMENTS.map(dept => (
                <div key={dept} style={{
                  background:"rgba(10,20,40,0.8)",border:"1px solid #0f2240",
                  borderRadius:8,padding:"10px 14px",minWidth:140,
                }}>
                  <div style={{fontSize:10,color:"#3b82f6",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:1,marginBottom:6}}>
                    {dept.toUpperCase()}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{color:"#4b6280",fontSize:13}}>£</span>
                    <input
                      type="number" className="rate-input"
                      value={deptRates[dept]}
                      onChange={e => handleDeptRate(dept, e.target.value)}
                      min={0} max={999}
                    />
                    <span style={{color:"#4b6280",fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}>/hr</span>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Individual overrides */}
          <div>
            <div style={{fontSize:12,color:"#4b6280",marginBottom:10,fontFamily:"'IBM Plex Mono',monospace"}}>
              Individual overrides — leave blank to use department rate
            </div>
            {DEPARTMENTS.map(dept => (
              <div key={dept} style={{marginBottom:16}}>
                <div style={{
                  fontSize:10,color:"#1e3a5f",fontFamily:"'IBM Plex Mono',monospace",
                  letterSpacing:2,textTransform:"uppercase",marginBottom:8,
                  paddingBottom:5,borderBottom:"1px solid #0f2240",
                }}>
                  {dept}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:8}}>
                  {team.map((t, i) => {
                    if (t.dept !== dept) return null;
                    return (
                      <div key={t.name} style={{
                        display:"flex",alignItems:"center",justifyContent:"space-between",
                        background:"rgba(10,20,40,0.5)",border:"1px solid #0f2240",
                        borderRadius:8,padding:"8px 12px",
                      }}>
                        <div style={{fontSize:12,color:t.override!==null?"#fbbf24":"#cbd5e1",flex:1}}>
                          {t.name}
                          {t.override !== null && <span style={{fontSize:10,color:"#f59e0b",marginLeft:6}}>↑</span>}
                        </div>
                        <div style={{display:"flex",alignItems:"center",gap:4}}>
                          <span style={{color:"#4b6280",fontSize:12}}>£</span>
                          <input
                            type="number"
                            className={`rate-input${t.override!==null?" overridden":""}`}
                            value={t.rate}
                            onChange={e => handleOverride(i, e.target.value)}
                            min={0} max={999}
                          />
                          {t.override !== null && (
                            <button
                              className="btn-secondary"
                              style={{padding:"4px 8px",fontSize:10,marginLeft:2}}
                              onClick={() => clearOverride(i)}
                              title="Reset to department rate"
                            >↺</button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
            <div style={{fontSize:11,color:"#4b6280",fontFamily:"'IBM Plex Mono',monospace",marginTop:4}}>
              Avg blended rate: £{avgRate.toFixed(2)}/hr · Used where individual rate is unavailable
            </div>
          </div>
        </div>

        {/* Generate */}
        <div style={{textAlign:"center",padding:"8px 0"}}>

          {status === "error" && (
            <div style={{
              background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.2)",
              borderRadius:8,padding:"12px 18px",marginBottom:16,
              color:"#fca5a5",fontSize:12,fontFamily:"'IBM Plex Mono',monospace",textAlign:"left",
            }}>
              ⚠ {errorMsg || "Something went wrong. Check your CSV files and try again."}
            </div>
          )}

          {previewStats && (
            <div style={{marginBottom:20}}>
              <div className="section-tag" style={{marginBottom:10,textAlign:"left"}}>Report Preview</div>
              <div style={{display:"flex",gap:12,flexWrap:"wrap",justifyContent:"center"}}>
                {[
                  { label:"Projects",   value: previewStats.projCount },
                  { label:"People",     value: previewStats.pplCount  },
                  { label:"Hours",      value: previewStats.totalHrs  },
                  { label:"TS Entries", value: previewStats.entries   },
                ].map(s => (
                  <div key={s.label} style={{
                    background:"rgba(10,20,40,0.7)",border:"1px solid #0f2240",
                    borderRadius:10,padding:"12px 20px",minWidth:100,textAlign:"center",
                  }}>
                    <div style={{fontFamily:"'Libre Baskerville',serif",fontSize:22,color:"#10b981",fontWeight:700}}>{s.value}</div>
                    <div style={{fontSize:10,color:"#4b6280",fontFamily:"'IBM Plex Mono',monospace",marginTop:2}}>{s.label}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <button
            className="btn-primary"
            disabled={!canGenerate || status === "generating"}
            onClick={handleGenerate}
          >
            {status === "generating" ? "⚙ Building Report…"
             : status === "done"     ? "⬇ Download Again"
             : "⬇ Generate & Download Report"}
          </button>

          {!ready && (
            <div style={{marginTop:10,fontSize:11,color:"#1e3a5f",fontFamily:"'IBM Plex Mono',monospace"}}>
              Upload both CSV files to continue
            </div>
          )}
          {status === "done" && (
            <div style={{marginTop:10,fontSize:11,color:"#10b981",fontFamily:"'IBM Plex Mono',monospace"}}>
              Report downloaded successfully
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{textAlign:"center",paddingTop:20,borderTop:"1px solid #0a1525"}}>
          <div style={{fontSize:10,color:"#1e3a5f",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:1}}>
            ALL PROJECTS: Dashboard · Project Profitability · Utilisation · Raw Timesheet
            &nbsp;·&nbsp;
            PROJECT VIEW: Project Summary · Hours by Person · Timesheet Detail
            &nbsp;·&nbsp;
            PERSON VIEW: Person Summary · Hours by Project · Daily Breakdown
          </div>
        </div>
      </div>
    </div>
  );
}
