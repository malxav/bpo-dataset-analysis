"""
03_generate_payroll_transactions.py
===================================
This script creates the monthly payroll history ledger for all employees.

DEPENDS ON
----------
- exports/Employees.csv                (Run script 01 first)
- exports/_hidden_employee_traits.csv  (The secret traits file from script 01)
- exports/Attendance.csv               (Run script 02 first)

OUTPUT
------
- exports/Payroll_Transactions.csv
  The final payroll register history.

THE TIME WINDOW & THE "DAYS WORKED" PUZZLE
------------------------------------------
* This file covers a 4-year span (Jan 2022 - Dec 2025), which is much longer 
  than the 1-year attendance file. Just like in a real company, payroll records 
  are kept forever, but detailed daily gate-swipe logs are usually cleaned out 
  after a year. (This window is expanded from 3 to 4 years to make sure the final 
  dataset hits the set target goal of 25,000 to 35,000 rows).

* Because the time frames don't perfectly match, the script calculates 'DaysWorked' 
  in two different ways:
  
  - For months inside 2025: The script counts the employee's actual, real days 
    worked directly from Attendance.csv. If you change the attendance script, 
    this number will automatically change to match.
    
  - For months before 2025 (2022-2024): Since we don't have daily swipe logs for 
    these years, the script safely estimates realistic working days instead (around 
    18-22 days for a full month, or fewer if the person was hired/quit mid-month).

WHAT IS DELIBERATELY LEFT BLANK
-------------------------------
To keep this exercise realistic, this script does NOT calculate final numbers like 
BasicPay, Overtime Pay, Taxes (SSS, PhilHealth, Pag-IBIG), or Net Pay. Those calculations 
are meant to be handled later by analysts using formulas inside the Excel Payroll Engine.

HOW ALLOWANCES WORK (Rules for data discovery)
----------------------------------------------
Instead of just dumping random amounts into 'GrossAllowance', the script 
builds the total allowance using actual business rules:

* Night Shift Allowance: Only goes to employees working the "Night" shift.
* Sales Incentives: Only goes to the "Sales" department. This is tied to an employee's 
  secret 'PerformancePotential' score, meaning top performers will naturally pull in 
  bigger bonuses. Students can spot this trend if they link payroll back to department data!
* IT On-Call Pay: Only goes to the "IT" department.
* Miscellaneous: A tiny random chance for anyone else to get a one-off minor allowance.

Each allowance is calculated independently month-by-month. If an employee doesn't 
qualify for any of them in a given month, the allowance field is left completely blank, 
which perfectly mirrors real-world payroll cycles.
"""

import os
import random
import numpy as np
import pandas as pd
from datetime import date, datetime
from calendar import monthrange

SEED = 42
random.seed(SEED)
np.random.seed(SEED)

PAYROLL_START = date(2022, 1, 1)
PAYROLL_END = date(2025, 12, 31)

# Must match 02_generate_attendance.py exactly, or the DaysWorked
# reconciliation below silently falls back to synthesis for the wrong months.
ATTENDANCE_COVERAGE_START = date(2025, 1, 1)
ATTENDANCE_COVERAGE_END = date(2025, 12, 31)

NIGHT_TRANSPORT_RANGE = (800, 1800)
NIGHT_TRANSPORT_GRANT_PROB = 0.85

SALES_INCENTIVE_RANGE = (500, 3000)
SALES_INCENTIVE_GRANT_PROB = 0.75

IT_ONCALL_RANGE = (500, 1500)
IT_ONCALL_GRANT_PROB = 0.60

MISC_ALLOWANCE_RANGE = (200, 1000)
MISC_ALLOWANCE_GRANT_PROB = 0.10


def month_range(start, end):
    months = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append(date(y, m, 1))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def load_attendance_days_worked(path="Attendance.csv"):
    """
    Returns dict {(EmployeeID, 'YYYY-MM'): days_worked} built from actual
    non-absent Attendance rows. Only covers whatever window Attendance.csv
    actually contains.
    """
    att = pd.read_csv(path, dtype=str, keep_default_na=False)
    att["ym"] = att["Date"].str.slice(0, 7)  # 'YYYY-MM'
    present = att[att["AbsenceFlag"] == "No"]
    counts = present.groupby(["EmployeeID", "ym"]).size()
    return counts.to_dict()


def synthesize_days_worked(is_partial, rng):
    if is_partial:
        return rng.randint(8, 16)
    return rng.randint(18, 22)


def compute_allowance(dept, shift, performance_potential, rng):
    total = 0.0
    got_any = False

    if shift == "Night" and rng.random() < NIGHT_TRANSPORT_GRANT_PROB:
        total += rng.uniform(*NIGHT_TRANSPORT_RANGE)
        got_any = True

    if dept == "Sales" and rng.random() < SALES_INCENTIVE_GRANT_PROB:
        lo, hi = SALES_INCENTIVE_RANGE
        # scale toward the top of the range for higher performers
        scaled_lo = lo + (hi - lo) * max(performance_potential - 0.3, 0) * 0.5
        total += rng.uniform(scaled_lo, hi)
        got_any = True

    if dept == "IT" and rng.random() < IT_ONCALL_GRANT_PROB:
        total += rng.uniform(*IT_ONCALL_RANGE)
        got_any = True

    if not got_any and rng.random() < MISC_ALLOWANCE_GRANT_PROB:
        total += rng.uniform(*MISC_ALLOWANCE_RANGE)
        got_any = True
    elif got_any and rng.random() < MISC_ALLOWANCE_GRANT_PROB * 0.5:
        # occasionally stack a small misc top-up even when another
        # allowance was already granted (e.g. Sales employee also on Night shift)
        total += rng.uniform(*MISC_ALLOWANCE_RANGE)

    return round(total, 2) if got_any else None


def main():
    os.makedirs("exports", exist_ok=True)
    os.makedirs("logs", exist_ok=True)

    employees = pd.read_csv("exports/Employees.csv", dtype=str, keep_default_na=False)
    traits_df = pd.read_csv("exports/_hidden_employee_traits.csv", dtype=str, keep_default_na=False)
    traits_df["PerformancePotential"] = traits_df["PerformancePotential"].astype(float)
    traits_lookup = traits_df.set_index("EmployeeID")["PerformancePotential"].to_dict()

    attendance_days = load_attendance_days_worked("exports/Attendance.csv")

    all_months = month_range(PAYROLL_START, PAYROLL_END)
    rows = []
    derived_count = 0
    synthesized_count = 0

    for idx, emp in employees.iterrows():
        emp_id = emp["EmployeeID"]
        dept = emp["Department"]
        shift = emp["ShiftType"]
        performance = traits_lookup.get(emp_id, 0.5)

        hire_date = datetime.strptime(emp["DateHired"], "%Y-%m-%d").date()
        exit_raw = emp["DateExited"]
        exit_date = (datetime.strptime(exit_raw, "%Y-%m-%d").date()
                     if isinstance(exit_raw, str) and exit_raw else None)

        rng = random.Random(f"{SEED}-payroll-{idx}")

        for month_start in all_months:
            month_end_day = monthrange(month_start.year, month_start.month)[1]
            month_end = date(month_start.year, month_start.month, month_end_day)

            if hire_date > month_end:
                continue
            if exit_date and exit_date < month_start:
                continue

            is_partial = (hire_date > month_start) or (exit_date is not None and exit_date < month_end)

            ym_key = f"{month_start.year:04d}-{month_start.month:02d}"
            in_attendance_window = ATTENDANCE_COVERAGE_START <= month_start <= ATTENDANCE_COVERAGE_END

            if in_attendance_window and (emp_id, ym_key) in attendance_days:
                days_worked = attendance_days[(emp_id, ym_key)]
                derived_count += 1
            else:
                days_worked = synthesize_days_worked(is_partial, rng)
                synthesized_count += 1

            # Assuming compute_allowance is defined elsewhere
            gross_allowance = compute_allowance(dept, shift, performance, rng)

            rows.append({
                "EmployeeID": emp_id,
                "PayPeriod": month_start.isoformat(),
                "DaysWorked": days_worked,
                "GrossAllowance": gross_allowance if gross_allowance is not None else "",
            })

    df = pd.DataFrame(rows)
    # Changed destination to 'exports/' to match scripts 1 and 2
    df.to_csv("exports/Payroll_Transactions.csv", index=False)

    # Open the log file and route all prints into it
    with open("logs/03.txt", "w") as f:
        print(f"Generated {len(df):,} payroll transaction rows -> exports/Payroll_Transactions.csv", file=f)
        print(f"  Unique employees represented: {df['EmployeeID'].nunique()}", file=f)
        print(f"  PayPeriod range: {df['PayPeriod'].min()} to {df['PayPeriod'].max()}", file=f)
        print(f"  DaysWorked derived from Attendance.csv: {derived_count:,} rows", file=f)
        print(f"  DaysWorked synthesized (outside Attendance coverage): {synthesized_count:,} rows", file=f)
        print(f"  Blank GrossAllowance rate: {(df['GrossAllowance']=='').mean():.2%}", file=f)
        print(file=f)
        low, high = 25000, 35000
        flag = "OK" if low <= len(df) <= high else "OUTSIDE 25k-35k TARGET BAND"
        print(f"  Row count target check: {len(df):,} rows [{flag}]", file=f)
        print(file=f)
        print("DaysWorked distribution:", file=f)
        print(df["DaysWorked"].describe(), file=f)

    print("Generation complete. Data saved to 'exports/' and execution logs saved to 'logs/03.txt'.")


if __name__ == "__main__":
    main()