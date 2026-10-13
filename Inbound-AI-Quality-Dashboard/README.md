# Inbound AI Quality Dashboard v2 — Power BI recreation (11 pages, redesigned)

Reference implementation: `My Dash\Mydashboards` (newer local copy; `frontend/.../ai-quality/InboundQualityDashboard.tsx`,
`backend/.../inbound-quality/inbound-quality.service.ts`). Source table: MySQL `db_audit.call_quality_assessment`.
Original Power BI file `Quality dashbaord.pbix` was NOT modified; a copy is `Quality dashbaord - ORIGINAL BACKUP.pbix`.

## How to open
1. Close any open Power BI windows for this project, then double-click `Quality.pbip`.
2. Click **Home > Refresh** once. The data is read from `Data\quality_data.csv` (a snapshot of the 5,654 current-month audits exported
   on 13-Oct-2026), so no MySQL login is needed.
3. File > Save As > `.pbix` if you want a single file (it will then include the data).

**Live data later:** the CSV is a snapshot. To go back to live MySQL, open Transform data > `Quality` > Advanced Editor and replace the
query with the original one:
`MySQL.Database("<your-mysql-host>", "db_audit", [ReturnSingleDatabase=true, Query="SELECT * FROM call_quality_assessment WHERE CallDate >= DATE_FORMAT(CURDATE(), '%Y-%m-01') AND CallDate < DATE_FORMAT(CURDATE() + INTERVAL 1 MONTH, '%Y-%m-01')"])`
(then re-apply the column types, or copy the query from your original `Quality dashbaord.pbix`).

## Default period
Every page is filtered to the **current month** by a report-level filter on `Date[Is Current Month]` (recalculates on each refresh using TODAY()). Change or clear it in the Filters pane to look at other months.


## Pages (v2)
1 Executive Summary · 2 Quality Performance · 3 Score Components · 4 Trend Analysis · 5 Agent Performance · 6 Agent Scorecard ·
7 Parameter Analysis · 8 Scenario & Duration (incl. ACHT categorisation) · 9 Fatal Errors · 10 Customer Insights (abuse / scam / social-court) · 11 Audit Details.
Design: navy header band, accent-bar KPI cards, shared palette (navy/blue/sky/amber/green/red), process + date + agent slicers on every page, current-month default.
Extra measures in v2: Abuse Calls, Potential Scam Calls, Social Media & Court Threat Calls, Scenario Share %, Avg Call Duration, Target CQ % (95).
Score colours in cards are fixed per KPI (no threshold-based conditional formatting yet).

## Model (table `Quality` = former `Query1`)
Calculated columns: `Is Audited`, `Call Date`, `Fatal Flag`, `NonFatal Flag`, `CQ Ratio`, `Opening/Soft/Hold/Resolution/Closing Ratio`,
`Quality Band`, `Call Category`, `Duration Bucket`. Tables: `Date` (calendar, related to `Quality[Call Date]`), `Parameter` (19 QA parameters, disconnected).

## Business rules (all taken from the web code)
| Item | Rule |
|---|---|
| Audited | `quality_percentage` not blank; ClientId 375/409/475 also need `quality_percentage > 35` |
| Per-call CQ | passed parameters / non-blank parameters over 19 params (20 for ClientId 468 incl. `express_empathy`); equal weights |
| Fatal audit | `case_escalated_correctly`, `address_recorded_completely`, `correct_and_complete_information` all = 0 (none blank) |
| CQ Score % | average per-call CQ x100 over audited, non-fatal calls (SQL NULL semantics reproduced), rounded 1 dp |
| CQ Score W/O Fatal % | **identical expression in the web code** (fatal calls are always excluded from "CQ Score") |
| Bands | stored `quality_percentage`: Excellent >=98, Good 90-97.99, Average 85-89.99, Below Avg >0 and <85 |
| Score components | 1/9/3/4/2 split of the 19 params; Call Drop / Short Call scenarios count as 100% |
| Agent grade | CQ >=90 TQ, >=80 MQ, else BQ; per-call counts use stored % (TQ >=80, MQ 60-79, BQ <60) |

## Web -> Power BI mapping
| Web component | Power BI |
|---|---|
| KPI cards (CQ, W/O Fatal, Audit Count, bands) | Overview cards (`Total Audits`, `CQ Score %`, ...) |
| Last 7 days / date trend | Line chart `CQ Score %` by `Date[Date]` |
| Agent Audit Summary / Agent&Parameter | Agent Performance page: ranking table with `Agent Rank`, `Agent Grade`, TQ/MQ/BQ |
| Quality Parameters table | Parameter Analysis page: `Param Evaluated/Pass/Fail/Pass %/Fail %` |
| Fatal Analysis | Fatal Errors page: fatal cards, by date, by agent, fatal record table |
| Raw Data tab / CSV export | Audit Details page (table; use visual menu > Export data) |
| Date + process filters | `Date` range slicer + `ClientId` slicer on every page |

## Validation (Power BI measures vs independent recomputation of the web rules from the same 5,654 rows)
Same filters: all 11 processes, 1-13 Oct 2026.

| Metric | Independent recompute | Power BI | Difference | Status |
|---|---:|---:|---:|---|
| Total Audits | 3,726 | 3,726 | 0 | Pass |
| CQ Score % | 88.1 | 88.1 | 0 | Pass |
| Excellent / Good / Average / Below Avg | 1,909 / 243 / 164 / 1,336 | 1,909 / 243 / 164 / 1,336 | 0 | Pass |
| Fatal audits / Fatal % | 38 / 1.0 | 38 / 1.0 | 0 | Pass |
| Per process (11 ClientIds): audits, CQ, fatal | matched | matched | 0 | Pass |

**Not yet validated:** a comparison against the numbers shown in the live web dashboard (it needs a login and the same process/date filters).
Open the web dashboard for one process, set the same date range, and compare with this report's `ClientId` + `Date` slicers.

## Gaps and open issues
- `AgentName`, process name (`md_clients`), CDR "calls answered" and `LOB` are not in the loaded table -> agent shows as MAS ID; no Audit Coverage %.
  Add `db_masmis.AgentMaster` and `shivamgiri.md_clients` queries and CDR tables to enable these.
- `Campaign` is blank on most rows; no LOB column exists in the data.
- Processes 473 and 474 have no graded parameters this month, so their CQ Score is blank (same as the web logic).
- Web "Average Call" band label and some web texts say "Score = 0" for fatal; the code uses the three-parameter rule above (implemented).
- Fraud/compliance text columns exist but their exact flag semantics are not defined in the web code for this dashboard, so no compliance KPI was invented.
- Drill-through pages and conditional-format colour rules from the web UI were not added.
- The report pages were validated against Microsoft's report schemas (52 visuals, 0 errors, all field references resolve) but could not be
  visually inspected because the PC screen was locked during the build.

## Note for this repository
The data file `Data/quality_data.csv` (customer call data: phone numbers, transcripts, recordings) is intentionally **not** committed.
To run the project, place your own export at `Data/quality_data.csv` (same columns as `db_audit.call_quality_assessment`), or change the
`Quality` query in Power BI (Transform data > Advanced Editor) to your MySQL source. The CSV path inside the query is
`C:\Users\MAS60358\Desktop\Inbound AI Quality Dashboard v2\Data\quality_data.csv`; change it to your local path.
