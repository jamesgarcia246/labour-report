# CLAUDE.md

 This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

 ## Commands

 ```bash
 npm start          # Dev server at http://localhost:3000
 npm run build      # Production build to /build
 npm test           # Run tests (interactive watch mode)
 npm test -- --watchAll=false  # Run tests once (CI mode)
 npm run deploy     # Build and deploy to GitHub Pages (jamesgarcia246.github.io/labour-report)
 ```

 ## Architecture

 This is a single-file React app (`src/App.js`) — all logic, UI, and Excel generation lives there. It is a labour/timesheet reporting tool that takes two CSV inputs and generates Excel reports.

 ### Data Flow

 1. User uploads two CSV files: a **timesheet** (hours per person per day per project) and a **task/budget** file (project budgets and planned hours).
 2. CSVs are parsed client-side (`parseCSV`, `readFile`).
 3. Name mapping is applied (`nameMap` state) to reconcile CSV names with the canonical `TEAM_DEFAULT` roster.
 4. The app computes derived metrics (burn rate, utilisation, variance) and passes everything to an Excel builder function.
 5. An `.xlsx` file is generated in-browser via `ExcelJS` and downloaded.

 ### Key Sections in App.js

 - **Constants** (line ~4): `DEPARTMENTS`, `TEAM_DEFAULT` (canonical team roster with dept/rate), `REPORT_TYPES`
 - **Helpers** (line ~35): CSV parsing, file reading, formatting (`fmtGBP`, `fmtPct`, `fmtHrs`, etc.)
 - **Excel Styles** (line ~73): `XL` style constants and helper functions (`xlStyleRow`, `xlTitle`, `xlSection`, `xlHeader`, `xlData`, `xlSubtotal`, `xlTotal`)
 - **Excel Builders** (line ~188): Three report builder functions — `buildAllProjectsReport`, `buildProjectViewReport`, `buildPersonViewReport` — each producing a multi-sheet workbook
 - **UI Components** (line ~626): `DropZone`, `MultiSelect`
 - **Main App** (line ~725): All state, parsing effects, and render

 ### Report Types

 | ID | Builder | Sheets |
 |---|---|---|
 | `all` | `buildAllProjectsReport` | Dashboard, Project Profitability, Utilisation, Raw Timesheet |
 | `project` | `buildProjectViewReport` | Project Summary, Hours by Person, Raw Timesheet (filtered) |
 | `person` | `buildPersonViewReport` | Person Utilisation Summary, Hours by Project, Day-by-day Breakdown |

 ### State Overview (Main App)

 - `tsRows` / `taskRows` — parsed CSV data
 - `team` — roster with per-person rate overrides on top of dept defaults
 - `deptRates` — default hourly rates per department
 - `nameMap` — maps CSV name strings to canonical team names
 - `reportType`, `selectedProjects`, `selectedPeople` — report filtering
 - `startDate`, `endDate`, `workingDays` — period configuration
 - `status`, `errorMsg`, `preview` — UI feedback state

 ### Dependencies of Note

 - `exceljs` — Excel file generation (not `xlsx`, despite `xlsx` being in package.json — check actual import)
 - `react-scripts` 5 (Create React App) — no custom webpack config
 - Deployed via `gh-pages` to GitHub Pages
