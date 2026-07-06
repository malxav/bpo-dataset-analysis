"""
01_generate_employees.py
========================
This script builds the base workforce data. Instead of just picking random 
numbers, it links everything together logically so the data feels like a 
real-world company.

OUTPUTS
-------
1. exports/Employees.csv
   The main HR file. This looks exactly like a real data export you'd pull 
   from a company's HR database.

2. exports/_hidden_employee_traits.csv  <-- NOT TO BE SHARED WITH ANALYSTS
   This is the "secret sauce" file. It creates hidden personality traits for 
   each employee (like how reliable they are, or if they are a high flight risk). 
   
   The other scripts (attendance and payroll) read this file so they can generate 
   matching behaviors. For example, an employee with a high 'AttendanceRisk' score 
   here will automatically show up late more often in the attendance logs. 
   Hide or delete this file if you want people to find these patterns on their own!

HOW IT WORKS (Read before changing constants)
---------------------------------------------
* Everything connects: A person's department determines their role. Their role 
  then determines their shift, work location, pay scale, and age range. 
  Nothing is completely random.
  
* Skills drive promotions: We assign personality traits first. If an employee 
  has high leadership potential, the script gently pushes them into a senior or 
  manager role. This naturally creates a realistic pattern where your managers 
  tend to be your most capable people.
  
* Realistic quitting: Attrition isn't a simple coin flip. Every single month, 
  the script calculates if an employee stays or leaves based on a mix of factors 
  (night shifts, bad attendance, or low pay make them much more likely to walk out).
"""

import os
import random
import numpy as np
import pandas as pd
from datetime import date, timedelta

# =============================================================================
# CONFIG
# =============================================================================
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

N_EMPLOYEES = 1200          # spec recommended default
DATA_START = date(2021, 1, 1)
DATA_END = date(2025, 12, 31)
SNAPSHOT = date(2025, 12, 31)   # "as of" date, so anyone still employed at this date is active

TARGET_ATTRITION_LOW, TARGET_ATTRITION_HIGH = 0.35, 0.45   # sanity-check band, see bottom of file

# =============================================================================
# NAME POOLS (Common Filipino names. Gender column must match the name drawn)
# =============================================================================
MALE_FIRST_NAMES = [
    "Juan", "Jose", "Mark", "Christian", "Michael", "John", "Paulo", "Ronald",
    "Ryan", "Kevin", "Ferdinand", "Aldrin", "Jayson", "Rommel", "Bryan",
    "Angelo", "Vincent", "Joshua", "Dennis", "Allan", "Rodel", "Arnel",
    "Reynaldo", "Edgar", "Noel", "Jerico", "Marvin", "Danilo", "Romeo",
    "Emmanuel", "Rico", "Ariel", "Warren", "Leo", "Randy", "Efren", "Rey",
    "Alvin", "Jomar", "Gilbert",
]
FEMALE_FIRST_NAMES = [
    "Maria", "Ana", "Angelica", "Jasmine", "Kristine", "Mary Grace", "Nicole",
    "Cherry", "Angel", "Jhoanna", "Rachel", "Camille", "Karen", "Divine",
    "Ella", "Trisha", "Katrina", "Grace", "Michelle", "Precious", "Joanna",
    "Charmaine", "Bea", "Loraine", "Shaira", "Kimberly", "Aira", "Gemma",
    "Julienne", "Melody", "Jennilyn", "Cristina", "Abigail", "Cassandra",
    "Daphne", "Rowena", "Liza", "Marites", "Jenny", "Erlinda",
]
LAST_NAMES = [
    "Santos", "Reyes", "Cruz", "Bautista", "Ocampo", "Garcia", "Mendoza",
    "Torres", "Andrade", "Castillo", "Villanueva", "Aquino", "Del Rosario",
    "Fernandez", "Gonzales", "Ramos", "Flores", "Salazar", "Domingo", "Pascual",
    "Aguilar", "Marasigan", "Navarro", "Concepcion", "De Leon", "Manalo",
    "Rivera", "Ignacio", "Roque", "Tolentino", "Valdez", "Hernandez",
    "Lazaro", "Bernardo", "Dela Cruz", "Pineda", "Espiritu", "Gutierrez",
    "Mercado", "Villaruz", "Abad", "Sarmiento", "Panganiban", "Corpuz",
    "Macaraeg", "Buenaventura", "Lopez", "Ferrer", "Custodio", "Sison",
]

# =============================================================================
# ORG STRUCTURE
# Each department: overall headcount weight + its own position ladder.
# Position tuple = (Position title, Level, relative weight within department)
# Level in {"entry", "mid", "senior", "manager"} drives age, tenure-recency,
# employment type, and pay. NOT department, which only sets multipliers.
# =============================================================================
DEPARTMENTS = [
    "Customer Service", "Technical Support", "Sales", "Back Office",
    "Workforce Management", "IT", "Finance", "HR",
]
DEPT_WEIGHTS = [0.32, 0.20, 0.14, 0.12, 0.08, 0.06, 0.05, 0.03]  # sums up to 1.00

POSITIONS_BY_DEPT = {
    "Customer Service": [
        ("Agent", "entry", 0.68), ("Team Lead", "mid", 0.14),
        ("QA Analyst", "mid", 0.10), ("Supervisor", "senior", 0.06), # also sum up to 1.00
        ("Manager", "manager", 0.02),
    ],
    "Technical Support": [
        ("Technical Support Agent", "entry", 0.62), ("Technical Team Lead", "mid", 0.15),
        ("QA Analyst", "mid", 0.11), ("Technical Supervisor", "senior", 0.08),
        ("Technical Support Manager", "manager", 0.04),
    ],
    "Sales": [
        ("Sales Agent", "entry", 0.55), ("Senior Sales Agent", "mid", 0.20),
        ("Sales Team Lead", "mid", 0.12), ("Sales Supervisor", "senior", 0.08),
        ("Sales Manager", "manager", 0.05),
    ],
    "Back Office": [
        ("Back Office Associate", "entry", 0.60), ("Senior Back Office Associate", "mid", 0.20),
        ("Back Office Team Lead", "mid", 0.10), ("Back Office Supervisor", "senior", 0.07),
        ("Back Office Manager", "manager", 0.03),
    ],
    "Workforce Management": [
        ("WFM Analyst", "entry", 0.55), ("Senior WFM Analyst", "mid", 0.25),
        ("WFM Supervisor", "senior", 0.12), ("WFM Manager", "manager", 0.08),
    ],
    "IT": [
        ("IT Support", "entry", 0.40), ("Systems Administrator", "mid", 0.25),
        ("Network Engineer", "mid", 0.20), ("IT Supervisor", "senior", 0.10),
        ("IT Manager", "manager", 0.05),
    ],
    "Finance": [
        ("Finance Associate", "entry", 0.38), ("Senior Finance Associate", "mid", 0.32),
        ("Finance Supervisor", "senior", 0.20), ("Finance Manager", "manager", 0.10),
    ],
    "HR": [
        ("HR Associate", "entry", 0.35), ("Recruiter", "entry", 0.20),
        ("HR Specialist", "mid", 0.25), ("HR Supervisor", "senior", 0.13),
        ("HR Manager", "manager", 0.07),
    ],
}

# Derive level weights per department from the position ladder above
# derived, not hand-duplicated, so the two can never drift out of sync.
LEVEL_WEIGHTS_BY_DEPT = {}
for _dept, _positions in POSITIONS_BY_DEPT.items():
    _levels = {}
    for _, _level, _w in _positions:
        _levels[_level] = _levels.get(_level, 0.0) + _w
    LEVEL_WEIGHTS_BY_DEPT[_dept] = _levels

SHIFT_WEIGHTS_BY_DEPT = {
    "Customer Service":      {"Night": 0.65, "Mid": 0.20, "Day": 0.15}, # most BPOs have foreign, US-based clients
    "Technical Support":     {"Night": 0.45, "Mid": 0.35, "Day": 0.20},
    "Sales":                 {"Day": 0.55, "Mid": 0.30, "Night": 0.15},
    "Back Office":           {"Day": 0.60, "Mid": 0.25, "Night": 0.15},
    "Workforce Management":  {"Day": 0.50, "Mid": 0.30, "Night": 0.20},
    "IT":                    {"Day": 0.85, "Mid": 0.10, "Night": 0.05},
    "Finance":               {"Day": 0.90, "Mid": 0.08, "Night": 0.02},
    "HR":                    {"Day": 0.95, "Mid": 0.05, "Night": 0.00}, # while positions like HR do their business during the day
}

SITE_WEIGHTS_BY_DEPT = {
    "Customer Service":      {"Manila": 0.45, "Cebu": 0.35, "Davao": 0.20},
    "Technical Support":     {"Manila": 0.50, "Cebu": 0.30, "Davao": 0.20},
    "Sales":                 {"Manila": 0.60, "Cebu": 0.25, "Davao": 0.15},
    "Back Office":           {"Manila": 0.50, "Cebu": 0.30, "Davao": 0.20},
    "Workforce Management":  {"Manila": 0.60, "Cebu": 0.25, "Davao": 0.15},
    "IT":                    {"Manila": 0.75, "Cebu": 0.15, "Davao": 0.10},
    "Finance":               {"Manila": 0.85, "Cebu": 0.10, "Davao": 0.05},
    "HR":                    {"Manila": 0.85, "Cebu": 0.10, "Davao": 0.05}, # huge concentration of BPOs situated in Manila
}

EMPLOYMENT_TYPE_WEIGHTS_BY_LEVEL = {
    "entry":   {"Regular": 0.55, "Probationary": 0.25, "Project-based": 0.20},
    "mid":     {"Regular": 0.72, "Probationary": 0.14, "Project-based": 0.14},
    "senior":  {"Regular": 0.88, "Probationary": 0.08, "Project-based": 0.04},
    "manager": {"Regular": 0.97, "Probationary": 0.03, "Project-based": 0.00},  # managers are never project-based
}

AGE_AT_HIRE_RANGE = {"entry": (19, 30), "mid": (23, 38), "senior": (27, 45), "manager": (30, 50)}

# Higher levels are biased toward earlier hire years. This is what
# produces "long-tenured employees are more likely to be supervisors or
# managers" without literally simulating promotions month by month.
YEAR_WEIGHTS_BY_LEVEL = {
    "entry":   {2021: 0.15, 2022: 0.20, 2023: 0.22, 2024: 0.23, 2025: 0.20},
    "mid":     {2021: 0.28, 2022: 0.26, 2023: 0.22, 2024: 0.16, 2025: 0.08},
    "senior":  {2021: 0.40, 2022: 0.28, 2023: 0.18, 2024: 0.10, 2025: 0.04},
    "manager": {2021: 0.50, 2022: 0.28, 2023: 0.14, 2024: 0.06, 2025: 0.02},
}
MONTH_SEASONALITY_WEIGHTS = [1.4, 1.3, 1.2, 0.8, 0.7, 0.9, 1.3, 1.2, 1.1, 0.9, 0.7, 0.9]  # Jan..Dec

BASE_MONTHLY_RANGE_BY_LEVEL = {
    "entry": (16000, 21000), "mid": (21000, 29000),
    "senior": (29000, 38000), "manager": (38000, 55000),
}
DEPT_PAY_MULTIPLIER = {
    "Customer Service": 1.00, "Technical Support": 1.05, "Sales": 1.02,
    "Back Office": 0.98, "Workforce Management": 1.03, "IT": 1.15,
    "Finance": 1.10, "HR": 1.05,
}
SITE_PAY_MULTIPLIER = {"Manila": 1.10, "Cebu": 1.00, "Davao": 0.95}

RATE_TYPE_WEIGHTS_BY_EMPTYPE = {
    "Regular": {"Monthly": 0.85, "Daily": 0.15},
    "Probationary": {"Monthly": 0.60, "Daily": 0.40},
    "Project-based": {"Monthly": 0.40, "Daily": 0.60},
}

# =============================================================================
# HELPERS
# =============================================================================

def weighted_choice(weight_dict, rng=random):
    labels = list(weight_dict.keys())
    weights = list(weight_dict.values())
    return rng.choices(labels, weights=weights, k=1)[0]


def clip(x, lo=0.0, hi=1.0):
    return max(lo, min(hi, x))


def draw_hidden_traits(rng, dept):
    """
    Six latent behavioral traits per employee, each ~Beta(4,2)-shaped
    (mean ~0.67, skewed toward the higher end but with a real left tail --
    most employees are decent, a meaningful minority are not).

    Department-conditioned shifts are applied ONLY to OvertimeWillingness,
    because that's the one trait the brief explicitly ties to department
    (Sales highest OT, Finance/HR lowest). Every other trait is left as an
    individual characteristic so it creates genuine person-level variation
    rather than just re-deriving department averages.
    """
    def beta(a=4, b=2):
        return float(np.random.default_rng(rng.randint(0, 2**31 - 1)).beta(a, b))

    reliability = beta(4, 2)
    attendance_risk = clip(1 - reliability + rng.gauss(0, 0.08))
    performance_potential = clip(0.5 * reliability + 0.5 * beta(4, 2))
    promotion_potential = beta(3, 3)
    leadership_potential = clip(0.4 * promotion_potential + 0.6 * beta(3, 3))

    ot_willingness = beta(3, 3)
    if dept == "Sales":
        ot_willingness = clip(ot_willingness + 0.20)
    elif dept == "Customer Service":
        ot_willingness = clip(ot_willingness + 0.08)
    elif dept in ("Finance", "HR"):
        ot_willingness = clip(ot_willingness - 0.20)

    return {
        "Reliability": round(reliability, 4),
        "AttendanceRisk": round(attendance_risk, 4),
        "PerformancePotential": round(performance_potential, 4),
        "OvertimeWillingness": round(ot_willingness, 4),
        "PromotionPotential": round(promotion_potential, 4),
        "LeadershipPotential": round(leadership_potential, 4),
    }


def choose_level(dept, leadership_score, rng):
    """
    Picks entry/mid/senior/manager for this department, nudged (not
    overridden) by the employee's leadership/promotion score. A score of
    0.5 leaves department base weights untouched; above 0.5 tilts toward
    senior/manager, below 0.5 tilts toward entry.
    """
    base = LEVEL_WEIGHTS_BY_DEPT[dept]
    adjusted = {}
    for level, w in base.items():
        if level in ("senior", "manager"):
            mult = 1.0 + (leadership_score - 0.5) * 1.3
        elif level == "entry":
            mult = 1.0 + (0.5 - leadership_score) * 0.7
        else:
            mult = 1.0
        adjusted[level] = max(w * mult, 0.0005)
    return weighted_choice(adjusted, rng)


def choose_position(dept, level, rng):
    candidates = {pos: w for pos, lvl, w in POSITIONS_BY_DEPT[dept] if lvl == level}
    return weighted_choice(candidates, rng)


def sample_hire_date(level, rng):
    year = int(weighted_choice(YEAR_WEIGHTS_BY_LEVEL[level], rng))
    month = rng.choices(range(1, 13), weights=MONTH_SEASONALITY_WEIGHTS, k=1)[0]
    day_max = 28 if month == 2 else (30 if month in (4, 6, 9, 11) else 31)
    day = rng.randint(1, day_max)
    d = date(year, month, day)
    return min(max(d, DATA_START), date(DATA_END.year, 12, 1))


def sample_birthdate(hire_date, level, rng):
    age_lo, age_hi = AGE_AT_HIRE_RANGE[level]
    age_at_hire = rng.randint(age_lo, age_hi)
    approx_days = int(age_at_hire * 365.25 + rng.randint(-60, 60))
    return hire_date - timedelta(days=approx_days)


def sample_employment_type(level, dept, rng):
    weights = dict(EMPLOYMENT_TYPE_WEIGHTS_BY_LEVEL[level])
    if level == "entry" and dept in ("Sales", "Customer Service"):
        # account-based staffing leans more project-based at entry level
        # in these two departments specifically
        weights["Project-based"] = weights["Project-based"] + 0.05
        weights["Regular"] = max(weights["Regular"] - 0.05, 0.05)
    return weighted_choice(weights, rng)


def sample_rate_type(employment_type, rng):
    return weighted_choice(RATE_TYPE_WEIGHTS_BY_EMPTYPE[employment_type], rng)


def sample_compensation(level, dept, site, rate_type, performance_potential, rng):
    lo, hi = BASE_MONTHLY_RANGE_BY_LEVEL[level]
    monthly_equiv = rng.uniform(lo, hi)
    monthly_equiv *= DEPT_PAY_MULTIPLIER[dept] * SITE_PAY_MULTIPLIER[site]
    monthly_equiv *= (1 + (performance_potential - 0.5) * 0.10)  # small performance-linked variation
    if rate_type == "Monthly":
        return round(monthly_equiv, 2)
    else:
        daily = (monthly_equiv / 21.0) * rng.uniform(0.90, 1.10)
        return round(daily, 2)


def simulate_exit(hire_date, shift, dept, employment_type, traits, rng):
    """
    Monthly hazard simulation. Returns (exit_date, exit_reason) or (None, None)
    if the employee survives to SNAPSHOT.
    """
    attendance_risk = traits["AttendanceRisk"]
    performance = traits["PerformancePotential"]
    engagement = (traits["LeadershipPotential"] + traits["PromotionPotential"]) / 2

    current = hire_date
    tenure_months = 0
    while True:
        tenure_months += 1
        year = current.year + (current.month - 1 + 1) // 12
        month = (current.month - 1 + 1) % 12 + 1
        current = date(year, month, min(current.day, 28))
        if current > SNAPSHOT:
            return None, None

        hazard = 0.0075
        if tenure_months <= 6:
            hazard *= 2.2
        elif tenure_months <= 12:
            hazard *= 1.3
        elif tenure_months > 48:
            hazard *= 0.55  # long-tenured employees are stickier

        if shift == "Night":
            hazard *= 1.35
        if dept == "Sales":
            hazard *= 1.25
        if employment_type == "Probationary":
            hazard *= 1.5
        elif employment_type == "Project-based":
            hazard *= 1.15

        hazard *= (1 + attendance_risk * 0.8)
        hazard *= (1 + (0.5 - performance) * 0.6)
        hazard *= (1 - (engagement - 0.5) * 0.4)
        hazard = min(hazard, 0.20)

        # Project-based contracts naturally end around 6 or 12 months
        if employment_type == "Project-based" and tenure_months in (6, 12) and rng.random() < 0.35:
            return current, "End of Contract"

        if rng.random() < hazard:
            reason = pick_exit_reason(tenure_months, employment_type, attendance_risk, performance, rng)
            return current, reason


def pick_exit_reason(tenure_months, employment_type, attendance_risk, performance, rng):
    if tenure_months <= 3:
        weights = {"AWOL": 0.40, "Resignation": 0.40, "Termination": 0.20}  #   AWOL is very common in this field
        if attendance_risk > 0.6:                                           #   and much more for risk-takers
            weights["AWOL"] += 0.15             
        return weighted_choice(weights, rng)    

    if employment_type == "Project-based" and rng.random() < 0.5:
        return "End of Contract"

    if tenure_months >= 48 and rng.random() < 0.10:
        return "Retirement"

    weights = {"Resignation": 0.50, "Termination": 0.30, "AWOL": 0.15, "End of Contract": 0.05}
    if performance < 0.4 or attendance_risk > 0.65:
        weights["Termination"] += 0.15
        weights["Resignation"] -= 0.10
    return weighted_choice(weights, rng)


# =============================================================================
# MAIN GENERATION LOOP
# =============================================================================

def generate_employees(n=N_EMPLOYEES):
    public_rows = []
    hidden_rows = []

    for i in range(1, n + 1):
        emp_id = f"EMP-{i:05d}"
        emp_rng = random.Random(f"{SEED}-emp-{i}")  # deterministic per-employee stream

        gender = weighted_choice({"Male": 0.48, "Female": 0.52}, emp_rng)
        first = emp_rng.choice(MALE_FIRST_NAMES if gender == "Male" else FEMALE_FIRST_NAMES)
        last = emp_rng.choice(LAST_NAMES)
        full_name = f"{first} {last}"

        dept = weighted_choice(dict(zip(DEPARTMENTS, DEPT_WEIGHTS)), emp_rng)
        traits = draw_hidden_traits(emp_rng, dept)

        leadership_score = (traits["LeadershipPotential"] + traits["PromotionPotential"]) / 2
        level = choose_level(dept, leadership_score, emp_rng)
        position = choose_position(dept, level, emp_rng)

        hire_date = sample_hire_date(level, emp_rng)
        birth_date = sample_birthdate(hire_date, level, emp_rng)

        employment_type = sample_employment_type(level, dept, emp_rng)
        rate_type = sample_rate_type(employment_type, emp_rng)
        shift = weighted_choice(SHIFT_WEIGHTS_BY_DEPT[dept], emp_rng)
        site = weighted_choice(SITE_WEIGHTS_BY_DEPT[dept], emp_rng)

        basic_rate = sample_compensation(level, dept, site, rate_type, traits["PerformancePotential"], emp_rng)

        exit_date, exit_reason = simulate_exit(hire_date, shift, dept, employment_type, traits, emp_rng)

        public_rows.append({
            "EmployeeID": emp_id,
            "FullName": full_name,
            "Gender": gender,
            "BirthDate": birth_date.isoformat(),
            "Department": dept,
            "Site": site,
            "Position": position,
            "EmploymentType": employment_type,
            "ShiftType": shift,
            "DateHired": hire_date.isoformat(),
            "DateExited": exit_date.isoformat() if exit_date else "",
            "ExitReason": exit_reason if exit_reason else "",
            "RateType": rate_type,
            "BasicRate": basic_rate,
        })

        hidden_rows.append({
            "EmployeeID": emp_id,
            "PositionLevel": level,
            **traits,
        })

    return pd.DataFrame(public_rows), pd.DataFrame(hidden_rows)


if __name__ == "__main__":
    import os

    os.makedirs("exports", exist_ok=True)
    os.makedirs("logs", exist_ok=True)

    df_public, df_hidden = generate_employees()
    df_public.to_csv("exports/Employees.csv", index=False)
    df_hidden.to_csv("exports/_hidden_employee_traits.csv", index=False)

    n = len(df_public)
    active = (df_public["DateExited"] == "").sum()
    exited = n - active
    attrition_rate = exited / n

    with open("logs/01.txt", "w") as f:
        print(f"Generated {n} employees -> exports/Employees.csv", file=f)
        print(f"  + _hidden_employee_traits.csv (INTERNAL ONLY -- do not ship to analyst)", file=f)
        print(f"  Active: {active} ({active/n:.1%}) | Exited: {exited} ({exited/n:.1%})", file=f)
        flag = "OK" if TARGET_ATTRITION_LOW <= attrition_rate <= TARGET_ATTRITION_HIGH else "OUT OF TARGET BAND (35-45%)"
        print(f"  Attrition rate: {attrition_rate:.1%}  [{flag}]", file=f)
        print(file=f)
        print("Exit rate by ShiftType:", file=f)
        print(df_public.groupby("ShiftType")["DateExited"].apply(lambda s: (s != "").mean()).round(3), file=f)
        print(file=f)
        print("Exit rate by Department:", file=f)
        print(df_public.groupby("Department")["DateExited"].apply(lambda s: (s != "").mean()).round(3), file=f)
        print(file=f)
        print("Headcount by Department:", file=f)
        print(df_public["Department"].value_counts(), file=f)
        print(file=f)
        print("Position level distribution by Department:", file=f)
        print(df_hidden.merge(df_public[["EmployeeID", "Department"]], on="EmployeeID")
              .groupby(["Department", "PositionLevel"]).size().unstack(fill_value=0), file=f)
        print(file=f)
        print("Age at snapshot sanity check (min/max):", file=f)
        ages = (pd.Timestamp(SNAPSHOT) - pd.to_datetime(df_public["BirthDate"])).dt.days / 365.25
        print(f"  min={ages.min():.1f}, max={ages.max():.1f}", file=f)

    print("Generation complete. Data saved to 'exports/' and execution logs saved to 'logs/01.txt'.")
