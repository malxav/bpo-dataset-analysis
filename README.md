# PH Workforce Analytics & Payroll Model

A synthetic-but-realistic BPO workforce dataset, generated from scratch with Python, fed into a fully formula-driven Excel workbook that computes real Philippine payroll (SSS, PhilHealth, Pag-IBIG, BIR withholding) and surfaces workforce analytics (attrition, attendance, headcount forecasting) the way a Senior Data Analyst role in a PH BPO actually would.

This repo is two things at once:
1. **A portfolio project.** Proof I can do real, deep Excel work, not just "I know VLOOKUP." Which if I'm gonna be honest, remained only a buzzword when I first started on this project. Now I know VLOOKUP among a lot of other things!
2. **A guidebook to myself.** This README is written so that if I lose touch with Excel a year from now, I can read it top to bottom and relearn most of what's in here without opening a single file.

---

## Why I built this

I kept noticing that a large share of Data Analyst job postings here in the Philippines ask for **Excel**, not Power BI, a bit of SQL, but not Python. That's not what most online DA learning paths optimize for, so I wanted to actually sit inside that reality instead of skipping past it.

BPOs specifically stood out to me. It's one of the biggest employers of data/analyst talent in the country, and a huge amount of what they report on, attendance, attrition, payroll, shift performance, is exactly the kind of data Excel is still the primary tool for in a lot of these companies. So instead of picking a generic "sales dashboard" (the most cloned portfolio piece that exists), I built something grounded in that specific industry.

The other thing I didn't want to do: have an AI just generate a CSV of random numbers and call it a dataset. Fake-looking synthetic data with no internal logic doesn't actually test or teach anything, since there's nothing to *discover* in it. So I built my own **data generator** instead, designed so that a real pattern-hunting analyst opening these CSVs cold would find real, causally-sensible relationships (night shift correlates with attrition, Sales does the most overtime, poor attendance correlates with resignation, and so on), the kind of signal you'd expect a real HRIS export to actually contain.

This whole project was a hands-on study, a way to expose myself to what an actual analyst deliverable looks like, mistakes and all, rather than a polished tutorial clone.

---

## Repo structure

```
├── data-generator/
│   ├── 01_generate_employees.py
│   ├── 02_generate_attendance.py
│   ├── 03_generate_payroll_transactions.py
│   ├── Employees.csv
│   ├── Attendance.csv
│   ├── Payroll_Transactions.csv
│   └── _hidden_employee_traits.csv        <- internal only, see below
│
├── PH_Workforce_Analytics_Payroll_Model.xlsm
└── README.md
```

---

## Part 1: The Dataset Generator

Instead of asking a model to hallucinate a CSV, I wrote three Python scripts (`data-generator/01` through `03`, run in that order, since each depends on the previous one's output) that simulate the dataset the way a real HRIS *produces* data: as the downstream effect of individual people with individual tendencies, not as independently rolled random rows.

### The core design idea: hidden traits

Every employee gets six invisible, un-exported behavioral traits at generation time:

| Trait | Drives |
|---|---|
| `Reliability` / `AttendanceRisk` | Lateness & absence probability in Attendance.csv |
| `PerformancePotential` | Attrition hazard, pay variation, incentive sizing |
| `OvertimeWillingness` | OT probability & magnitude |
| `PromotionPotential` / `LeadershipPotential` | Likelihood of landing a senior/manager position, attrition hazard |

These traits are generated once per employee in script 01 and saved to `_hidden_employee_traits.csv`, an **internal-only file, not a deliverable**. Scripts 02 and 03 read it to keep every table correlated with the same underlying "person," the same way a real employee's actual reliability shows up consistently across their attendance record, their exit outcome, and their pay history. It's what makes "employees with poor attendance are more likely to resign" a real, discoverable pattern in the data rather than a coincidence: the same trait quietly drives both outcomes.

### What makes each table non-random

- **Employees.csv.** Position depends on Department (each department has its own job ladder: Finance has no "Agent" title, IT is mostly Day shift, Customer Service is mostly Night). Attrition is a **monthly hazard simulation**, not a single dice roll. Risk compounds from tenure, shift, department, employment type, and the hidden traits together, which is why exit timing distributions look like a real survival curve instead of noise.
- **Attendance.csv.** Lateness/absence probability scales off each employee's own `AttendanceRisk`, with added multipliers for Night shift, new-hire status (under 90 days), Monday absence bumps, and seasonal effects (rainy season sick-leave bump, December vacation-leave bump). Night shift genuinely crosses midnight in the timestamps (22:00 to 07:00), which matters later for the Night Differential formula in Excel.
- **Payroll_Transactions.csv.** `DaysWorked` is **derived from real Attendance rows**, not invented, for every month Attendance.csv actually covers (2025). For 2022 through 2024, payroll history that predates the attendance log's coverage window (a deliberate design choice mirroring how real companies often retain payroll registers far longer than raw daily swipe logs), it falls back to a realistic synthesized range. `GrossAllowance` is eligibility-gated (Night shift gets a transport allowance, Sales gets an incentive scaled by `PerformancePotential`, IT gets an on-call allowance) and legitimately blank most months for anyone not eligible for any of the three.

### Numbers that came out of the final run

- **1,200 employees**, roughly 39% cumulative attrition (target band was 35 to 45%)
- **About 183,000 attendance rows** (Jan through Dec 2025 only)
- **About 26,700 payroll transaction rows** (Jan 2022 through Dec 2025)

### A real bug I hit worth flagging
Early drafts seeded Python's RNG with tuples (`random.Random((SEED, i))`), which worked in my dev sandbox's Python version but throws `TypeError` on Python 3.11+, since `random.Random()` now only accepts `None, int, float, str, bytes, bytearray`. Fixed by switching every seed to an f-string. Small thing, but a good reminder that "it ran once" isn't the same as "it's portable."

---

## Part 2: The Excel Workbook, tab by tab

The architecture is a proper **star schema**, not a flat mega-sheet: `Employees` sits at the center with relationships out to `Attendance` and `Payroll_Transactions`, plus a generated `Calendar` date table for time intelligence. This is the single biggest thing separating this build from a beginner one; most self-taught Excel projects never leave the "one giant sheet with VLOOKUPs" stage.

| Tab | Purpose |
|---|---|
| `Employees`, `Attendance`, `Payroll_Transactions` | Raw source tables, loaded via Power Query **and** as worksheet Tables (see note below on why both) |
| `Calendar` | Generated date dimension (Jan 2022 to Dec 2025), marked as an official Date Table so DAX time-intelligence works correctly |
| `PH_Holidays` | Reference list of 2025 PH regular & special non-working holidays, used to compute Holiday Pay |
| `Statutory_Reference` | SSS, PhilHealth, Pag-IBIG, and BIR bracket tables as actual named cells and a lookup table, not numbers buried inside formulas |
| `Payroll_Engine` | The centerpiece: full Basic Pay to Net Pay computation, one row per employee-month, built entirely from formulas (see Part 3) |
| `Attrition_Analysis` | KPI strip, rolling attrition trend, cohort survival curves, Dept x Shift heatmap, a rule-based Flight Risk watchlist, and a voluntary-vs-involuntary exit reason breakdown |
| `Headcount_Forecasts` | FORECAST.ETS projection with confidence intervals, a Scenario Manager (conservative/base/aggressive hiring), and a Goal Seek example |
| `Dashboard_Executive`, `Dashboard_Payroll` | One-screen visual summaries, rebuilt via a VBA macro rather than manual Pivots (story below) |

**Why some tables are loaded to both Power Query's Data Model and as a plain worksheet Table:** Power Pivot's Data Model is invisible to normal cell formulas. XLOOKUP, SUMIFS, and COUNTIFS can't see anything that lives only in the Data Model. So anything that needs both DAX measures (for Pivots) and plain-formula lookups (for Payroll_Engine) gets loaded to both destinations from the same query. No duplication of the query itself, just two load targets.

---

## Part 3: The Formula Work (the actual point of this project)

This section is the guidebook part. If I forget how any of this works later, this is what I read first.

### `Payroll_Engine`, built in four stages, each verified before the next

**Stage 1: Basic Pay.** Monthly-rate employees get prorated against a standard 22-working-day month; Daily-rate employees get `Rate × DaysWorked` directly, since their rate already is a per-day figure and needs no proration.
```excel
=IF([@RateType]="Monthly", ([@BasicRate]/StdMonthlyDays)*[@DaysWorked], [@BasicRate]*[@DaysWorked])
```

**Stage 2: OT Pay & Night Differential.** The interesting one. PH labor law ties night differential to *actual clock hours between 10PM and 6AM*, not to which shift someone is officially labeled, so the formula reads real `ActualTimeIn`/`ActualTimeOut` values and measures the overlap with a fixed nightly window, correctly handling the case where a shift crosses midnight:
```excel
=MAX(0,
   MIN([@Date]+[@ActualTimeOut]+IF([@ActualTimeOut]<[@ActualTimeIn],1,0), [@Date]+1+TIME(6,0,0))
   - MAX([@Date]+[@ActualTimeIn], [@Date]+TIME(22,0,0))
)*24
```
This one formula alone correctly handles Day, Mid, and Night shift without any shift-specific branching. Mid shift's natural 1-hour daily overlap (10 to 11PM) and Night shift's much larger overlap both just fall out of the same math.

**Stage 3: Holiday Pay.** An array-based `SUMPRODUCT` cross-references every holiday in `PH_Holidays` against that employee's actual attendance record for the month. Regular Holidays pay out unless the employee was logged absent that exact date; Special Holidays only pay out (a 30% premium) if the employee actually showed up, matching the true "no work, no pay" rule under PH labor law:
```excel
=SUMPRODUCT(
   (PH_Holidays[HolidayType]="Regular") *
   (MONTH(PH_Holidays[Date])=MONTH([@PayPeriod])) *
   (COUNTIFS(Attendance[EmployeeID],[@EmployeeID],Attendance[Date],PH_Holidays[Date],Attendance[AbsenceFlag],"Yes")=0)
) * [@DailyRate]
```
Guarded with a `YEAR([@PayPeriod])<>2025` check, since Attendance.csv only covers 2025. Without the guard, "zero matching absence rows found" for 2022 through 2024 would be misread as "eligible," wrongly granting holiday pay for years with no attendance data to back it up.

**Stage 4: Statutory Deductions & Net Pay.** SSS, PhilHealth, and Pag-IBIG all reference **named cells** on `Statutory_Reference` (`SSS_EmployeeRate`, `PagIBIG_MaxDeduction`, and so on) rather than hardcoded numbers, so a rate change next year means editing one cell, not hunting through formulas. Withholding Tax uses a modern `LET` + `XMATCH` bracket lookup against a real BIR bracket table instead of a giant nested `IFS`, same math, more maintainable:
```excel
=LET(
  r, XMATCH([@TaxableIncome], tbl_BIR_Brackets[Bracket Floor], -1),
  base, INDEX(tbl_BIR_Brackets[Base Tax], r),
  rate, INDEX(tbl_BIR_Brackets[Rate on Excess], r),
  thresh, INDEX(tbl_BIR_Brackets[Excess Threshold], r),
  base + ([@TaxableIncome]-thresh)*rate
)
```

### `Attrition_Analysis`, the DAX and array-formula layer

- **DAX measures** live on whichever table they logically belong to. This is Excel's actual Power Pivot convention, not a disconnected "_Measures" table, which is a Power BI habit I initially carried over by mistake and had to correct. Example: attrition rate done properly as *exits divided by average headcount*, not exits divided by ending headcount:
```dax
Attrition Rate (Period) := DIVIDE([Exits in Period], [Average Headcount (Period)])
```
- **`USERELATIONSHIP`** shows up because `Calendar` has two possible join targets (`Attendance[Date]` and `Payroll_Transactions[PayPeriod]`), but a Data Model can only have one *active* relationship between two tables at a time. Any measure touching the inactive path has to invoke the other relationship explicitly:
```dax
Total Gross Allowance := 
CALCULATE(SUM(Payroll_Transactions[GrossAllowance]), USERELATIONSHIP(Calendar[Date], Payroll_Transactions[PayPeriod]))
```
- **Cohort survival analysis.** A COUNTIFS-based grid answering "of everyone hired in month X, what percent were still active at 3, 6, and 12 months," which genuinely showed the early-tenure attrition risk built into the generator once charted (3-month survival visibly lower than 12-month).
- **`FILTER`/`CHOOSECOLS`** power a live Flight Risk watchlist: active, under 6 months tenure, Night shift, plus elevated recent OT or lateness, pulled dynamically rather than kept as a static list.
- **`CUBEVALUE`** is what lets a DAX measure show up in a plain cell (like a KPI card) instead of only inside a Pivot. It's an obscure corner of Excel most people never need, but the only bridge between the Data Model and free-floating cells.

### `Headcount_Forecasts`
`FORECAST.ETS` and `FORECAST.ETS.CONFINT` produce a 12-month-out projection with a 95% confidence band, run against 48 months of actual monthly headcount. Paired with Scenario Manager (three named hiring-rate scenarios) and a Goal Seek example ("how many hires per month to hit a target headcount").

### VBA Automation: a real mid-project pivot (pun intended)
The dashboards were originally scoped as manual Pivots, Slicers, and Timelines. Partway through, the workbook started showing slicer/extension corruption on reopen, a known Excel fragility, especially across saves and version differences. Rather than keep fighting it, I rebuilt both dashboards as a single VBA macro (`Build_BPO_Dashboards`) that:
- Populates two hidden calculation sheets (`Calc_Exec`, `Calc_Payroll`) with `SUMIFS`/`COUNTIFS` rollups
- Draws native shape-based KPI cards and native charts (no Pivots at all) from scratch every time it's run
- Is fully idempotent and safe to re-run any time the source data changes, via a `Refresh Dashboards` button embedded on each dashboard sheet

This ended up more robust and more portable across Excel versions than the Pivot-based version, a real lesson in knowing when to abandon a "more standard" approach for a more resilient one.

---

## Known Limitations (things I chose to leave imperfect, on purpose)

- **Payroll currently runs monthly, not semi-monthly.** Real large PH BPOs typically cut payroll on the 1st to 15th and 16th to end-of-month schedule. A friend in the industry flagged this. The fix isn't hard, but it means regenerating `Payroll_Transactions.csv` at a semi-monthly grain and adjusting the `EOMONTH` boundary logic in `Payroll_Engine`. Deferred as a planned upgrade pass, not done yet.
- **Holiday Pay doesn't check the day before the holiday**, a real DOLE eligibility nuance (you can lose holiday pay if absent the day prior with no approved leave). Skipped as a documented simplification, since it would add real formula weight for a rule that rarely changes the outcome in this dataset.
- **The Dept x Shift attrition heatmap has real small-N noise.** A couple of cells (for example, Finance x Mid Shift) represent only 3 to 4 people, so one exit swings the percentage wildly. The Grand Total margins are the trustworthy view; interior cells with low headcount are not. Knowing when a slice is too thin to trust is a real analytical judgment call, not something to prettify away.
- **The Flight Risk Watchlist is rule-based** (fixed thresholds on tenure, OT, and lateness), not a real predictive model. Excel can't do real logistic regression without add-ins, and I didn't want to fake that it could.
- **Row-scale ceiling.** About 183,000 attendance rows is near the upper edge of comfortable Power Query/Pivot performance. A real BPO's multi-year daily swipe log would be millions of rows; past that point, the correct architecture is a database plus a real BI tool, not Excel. This project intentionally sits right at that edge on purpose, to make the ceiling visible rather than avoid it.

---

## What I'd do next

Keeping this simple rather than overpromising: the next step I'd actually take is **loading this same, already-cleaned, already-computed data into Power BI for the visualization layer**, while keeping Excel as the data-prep, formula-computation, and payroll-engine layer it's already good at. Excel proved it can carry real analytical weight here, but Power BI's DAX engine, native slicers, and publishing/sharing model are a better fit once the goal shifts from "prove I can build this" to "actually distribute this to stakeholders."

Beyond that, the deferred items from earlier are still open: the semi-monthly payroll conversion, and building a second, Power Query-fed version of `Payroll_Engine` that's auto-refreshing but less manually auditable, a real tradeoff worth showing both sides of. If I'm gonna be honest here I had to limit myself to this because it's becoming harder to input new functions due to current hardware limitations. I plan to revisit this in the future.

---

## Excel/Power Pivot skills index (for future-me)

A quick lookup table so I don't have to re-derive *why* I used something, just where to go look at it again.

| Skill | Where it's used |
|---|---|
| Power Query (Get & Transform, multi-destination loads) | Layer 1, all three raw tables |
| Star schema and relationships, including dual-relationship + `USERELATIONSHIP` | Data Model, `Calendar` to `Attendance`/`Payroll_Transactions` |
| DAX measures, correct table placement (not a disconnected measures table) | `Attrition_Analysis` |
| `XLOOKUP`, `SUMIFS`, `COUNTIFS` | `Payroll_Engine` throughout |
| `SUMPRODUCT` array logic | Holiday Pay eligibility |
| Date/time arithmetic across midnight | Night Differential |
| `LET` + `XMATCH` bracket lookups | Withholding Tax |
| `FILTER`, `CHOOSECOLS`, dynamic arrays | Flight Risk Watchlist |
| `CUBEVALUE` | KPI cards reading DAX measures outside a Pivot |
| `FORECAST.ETS` / `.CONFINT` | `Headcount_Forecasts` |
| Scenario Manager, Goal Seek, Data Tables | `Headcount_Forecasts`, `Dashboard_Payroll` OT sensitivity |
| Named ranges over hardcoded constants | `Statutory_Reference` throughout |
| VBA: idempotent macro design, native chart/shape generation, hidden calc sheets | `Build_BPO_Dashboards` |

---

## How to run this project

1. `cd data-generator`, run the three Python scripts in order: `01_generate_employees.py`, then `02_generate_attendance.py`, then `03_generate_payroll_transactions.py` (requires `pandas`, `numpy`)
2. Open `PH_Workforce_Analytics_Payroll_Model.xlsm`, enable macros when prompted
3. `Data → Refresh All` to pull the latest CSVs through Power Query
4. Click the **Refresh Dashboards** button on `Dashboard_Executive` or `Dashboard_Payroll` to rebuild both from current data
