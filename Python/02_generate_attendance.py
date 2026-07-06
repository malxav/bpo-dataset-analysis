"""
02_generate_attendance.py
==========================
This script creates the attendance logs (swipes, lates, absences, and overtime) 
for the year 2025. 

DEPENDS ON
----------
- exports/Employees.csv                (Run script 01 first!)
- exports/_hidden_employee_traits.csv  (The secret traits file from script 01)

OUTPUT
------
- exports/Attendance.csv
  The final attendance log file. The script reads the hidden traits to decide 
  who shows up late or calls in sick, but it never saves those traits here. 

THE TIME WINDOW (and why it only covers 2025)
----------------------------------------
* Employees.csv covers a 5-year hiring history, but this attendance file only 
  covers 12 months (Jan 2025 - Dec 2025). 
  
* According to my research, companies usually keep years of payroll history, but 
  they only keep detailed daily swipe logs for the most recent year or two. 
  Limiting this file to one year also keeps the final dataset at a clean, 
  manageable size (around 150k to 250k rows). Good for practice.

THE RULES FOR GENERATING DATA (What people can "discover")
----------------------------------------------------------
* Bad habits stick: We look up each employee's secret 'AttendanceRisk' score. 
  An employee who is naturally unreliable will consistently show up late or miss 
  days here, and they are the exact same people who were more likely to quit 
  early in script 01. This links the two files together perfectly.

* The Night Shift effect: Working the night shift automatically adds an extra 
  penalty to an employee's attendance. Even good employees are slightly more 
  prone to lateness, absences, or burning out on overtime if they work nights.

* The Rookie phase: New hires (people with less than 90 days on the job) get a 
  temporary penalty that makes them more likely to be late as they adjust to 
  the company.

* Days of the week: Mondays see a natural spike in unexplained absences. On 
  Fridays, absences shift away from "Sick Leave" and more toward planned 
  "Vacation" or "Emergency Leave".

* Seasons change: The rainy season (June to November) triggers a spike in 
  Sick Leave. December triggers a massive wave of planned Vacation Leave for 
  the holidays.

* Overtime drivers: Overtime is controlled by department needs (Sales gets the 
  most, HR/Finance gets the least) but is adjusted up or down based on the 
  employee's individual willingness to work extra hours.
"""

import os
import random
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta

SEED = 42
random.seed(SEED)
np.random.seed(SEED)

REPORT_START = date(2025, 1, 1)
REPORT_END = date(2025, 12, 31)

SHIFT_SCHEDULE = {
    "Day":   {"in": (7, 0),  "out": (16, 0)},
    "Mid":   {"in": (14, 0), "out": (23, 0)},
    "Night": {"in": (22, 0), "out": (7, 0)},   # crosses midnight
}

LEAVE_TYPES = ["Sick", "Vacation", "Emergency"]
BASE_LEAVE_WEIGHTS = {"Sick": 0.50, "Vacation": 0.30, "Emergency": 0.20}

RAINY_SEASON_MONTHS = {6, 7, 8, 9, 10, 11}   # Jun-Nov
DECEMBER = 12

BASE_ABSENCE_PROB = 0.030
BASE_LATE_PROB = 0.100

OT_BASE_PROB_BY_DEPT = {
    "Sales": 0.22, "Customer Service": 0.13, "Technical Support": 0.12,
    "Back Office": 0.08, "Workforce Management": 0.10, "IT": 0.07,
    "Finance": 0.05, "HR": 0.03,
}

NEW_HIRE_TENURE_DAYS = 90


def fmt_time(h, m):
    return f"{h:02d}:{m:02d}"


def minutes_to_hm(total_minutes):
    total_minutes = total_minutes % 1440
    return total_minutes // 60, total_minutes % 60


def generate_rest_days(seed_key):
    """Fixed pair of weekly rest days (0=Mon..6=Sun) for the employee's whole tenure."""
    rng = random.Random(seed_key)
    return set(rng.choice([
        (5, 6), (6, 0), (0, 1), (1, 2), (2, 3), (3, 4), (4, 5),
    ]))


def leave_type_weights(month, weekday):
    weights = dict(BASE_LEAVE_WEIGHTS)
    if month in RAINY_SEASON_MONTHS:
        weights["Sick"] += 0.15
    if month == DECEMBER:
        weights["Vacation"] += 0.20
    if weekday == 4:  # Friday
        weights["Vacation"] += 0.08
        weights["Emergency"] += 0.05
        weights["Sick"] = max(weights["Sick"] - 0.10, 0.05)
    return weights


def build_attendance_for_employee(emp, traits, idx):
    emp_id = emp["EmployeeID"]
    dept = emp["Department"]
    shift = emp["ShiftType"]
    hire_date = datetime.strptime(emp["DateHired"], "%Y-%m-%d").date()
    exit_raw = emp["DateExited"]
    exit_date = (datetime.strptime(exit_raw, "%Y-%m-%d").date()
                 if isinstance(exit_raw, str) and exit_raw else None)

    window_start = max(hire_date, REPORT_START)
    window_end = min(exit_date, REPORT_END) if exit_date else REPORT_END
    if window_start > window_end:
        return []

    attendance_risk = traits["AttendanceRisk"]
    ot_willingness = traits["OvertimeWillingness"]

    late_risk_mult = 0.4 + 1.4 * attendance_risk
    absence_risk_mult = 0.4 + 1.4 * attendance_risk
    late_minutes_scale = 0.7 + 0.6 * attendance_risk

    night_mult = 1.35 if shift == "Night" else 1.0
    ot_base_prob = OT_BASE_PROB_BY_DEPT.get(dept, 0.08)
    ot_prob_mult = 0.4 + 1.2 * ot_willingness
    ot_hours_mult = 0.6 + 0.8 * ot_willingness

    rest_days = generate_rest_days(f"{SEED}-restday-{idx}")
    sched = SHIFT_SCHEDULE[shift]
    sched_in_h, sched_in_m = sched["in"]
    sched_out_h, sched_out_m = sched["out"]

    rng = random.Random(f"{SEED}-attendance-{idx}")
    rows = []
    current = window_start

    while current <= window_end:
        weekday = current.weekday()
        if weekday not in rest_days:
            tenure_days = (current - hire_date).days
            is_new_hire = tenure_days < NEW_HIRE_TENURE_DAYS

            absence_p = BASE_ABSENCE_PROB * absence_risk_mult * night_mult
            if weekday == 0:  # Monday
                absence_p *= 1.30
            absence_p = min(absence_p, 0.35)

            if rng.random() < absence_p:
                weights = leave_type_weights(current.month, weekday)
                leave_type = rng.choices(list(weights.keys()), weights=list(weights.values()), k=1)[0]
                rows.append({
                    "EmployeeID": emp_id,
                    "Date": current.isoformat(),
                    "ScheduledTimeIn": fmt_time(sched_in_h, sched_in_m),
                    "ScheduledTimeOut": fmt_time(sched_out_h, sched_out_m),
                    "ActualTimeIn": "",
                    "ActualTimeOut": "",
                    "AbsenceFlag": "Yes",
                    "LeaveType": leave_type,
                    "LateMinutes": 0,
                    "OTHours": 0.0,
                })
            else:
                late_p = BASE_LATE_PROB * late_risk_mult * night_mult
                if is_new_hire:
                    late_p *= 1.40
                late_p = min(late_p, 0.6)

                late_minutes = 0
                if rng.random() < late_p:
                    raw = rng.gammavariate(2.0, 10.0 * late_minutes_scale)
                    late_minutes = int(min(max(raw, 5), 90))

                total_in_minutes = sched_in_h * 60 + sched_in_m + late_minutes
                act_in_h, act_in_m = minutes_to_hm(total_in_minutes)

                ot_hours = 0.0
                extra_minutes = 0
                ot_p = min(ot_base_prob * ot_prob_mult, 0.55)
                if rng.random() < ot_p:
                    ot_hours = round(rng.uniform(0.5, 3.0) * ot_hours_mult, 2)
                    ot_hours = min(ot_hours, 4.0)
                    extra_minutes = int(round(ot_hours * 60))
                total_out_minutes = sched_out_h * 60 + sched_out_m + extra_minutes
                act_out_h, act_out_m = minutes_to_hm(total_out_minutes)

                rows.append({
                    "EmployeeID": emp_id,
                    "Date": current.isoformat(),
                    "ScheduledTimeIn": fmt_time(sched_in_h, sched_in_m),
                    "ScheduledTimeOut": fmt_time(sched_out_h, sched_out_m),
                    "ActualTimeIn": fmt_time(act_in_h, act_in_m),
                    "ActualTimeOut": fmt_time(act_out_h, act_out_m),
                    "AbsenceFlag": "No",
                    "LeaveType": "",
                    "LateMinutes": late_minutes,
                    "OTHours": ot_hours,
                })
        current += timedelta(days=1)

    return rows


def main():
    os.makedirs("exports", exist_ok=True)
    os.makedirs("logs", exist_ok=True)

    employees = pd.read_csv("exports/Employees.csv", dtype=str, keep_default_na=False)
    traits_df = pd.read_csv("exports/_hidden_employee_traits.csv", dtype=str, keep_default_na=False)
    
    for col in ["AttendanceRisk", "OvertimeWillingness"]:
        traits_df[col] = traits_df[col].astype(float)
    traits_lookup = traits_df.set_index("EmployeeID").to_dict(orient="index")

    all_rows = []
    for idx, emp in employees.iterrows():
        traits = traits_lookup[emp["EmployeeID"]]
        all_rows.extend(build_attendance_for_employee(emp, traits, idx))

    df = pd.DataFrame(all_rows)
    df.to_csv("exports/Attendance.csv", index=False)

    with open("logs/02.txt", "w") as f:
        print(f"Generated {len(df):,} attendance rows -> exports/Attendance.csv", file=f)
        print(f"  Unique employees represented: {df['EmployeeID'].nunique()}", file=f)
        print(f"  Date range: {df['Date'].min()} to {df['Date'].max()}", file=f)
        print(f"  Overall absence rate: {(df['AbsenceFlag']=='Yes').mean():.2%}", file=f)
        present = df[df.AbsenceFlag == "No"]
        print(f"  Overall late rate (of present days): {(present['LateMinutes']>0).mean():.2%}", file=f)
        print(f"  Overall OT rate (of present days): {(present['OTHours']>0).mean():.2%}", file=f)
        print(file=f)

        merged = df.merge(employees[["EmployeeID", "Department", "ShiftType"]], on="EmployeeID")
        present_m = merged[merged.AbsenceFlag == "No"]
        print("Late rate by ShiftType:", file=f)
        print(present_m.groupby("ShiftType").apply(lambda g: (g["LateMinutes"] > 0).mean()).round(3), file=f)
        print(file=f)
        print("OT rate by Department:", file=f)
        print(present_m.groupby("Department").apply(lambda g: (g["OTHours"] > 0).mean()).round(3), file=f)
        print(file=f)
        print("Absence rate by weekday (0=Mon..6=Sun):", file=f)
        merged["weekday"] = pd.to_datetime(merged["Date"]).dt.weekday
        print(merged.groupby("weekday").apply(lambda g: (g["AbsenceFlag"] == "Yes").mean()).round(3), file=f)

    print("Generation complete. Data saved to 'exports/' and execution logs saved to 'logs/02.txt'.")


if __name__ == "__main__":
    main()
