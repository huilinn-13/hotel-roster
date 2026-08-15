# demo_generator.py
"""
Generate a richer demo dataset and one full roster output.

Run with:
    python demo_generator.py

It creates:
    - employees_demo.xlsx  (richer day-off / preferred-hours data)
    - generated_roster.xlsx (the auto-generated schedule)
    - printed summary stats
"""

import os
import pandas as pd
from roster_data import days, hours, STAFFING_NEEDS
from scheduler import generate_schedule
from rules_engine import check_rules

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEMO_EMPLOYEES_FILE = os.path.join(BASE_DIR, "employees_demo.xlsx")
OUTPUT_ROSTER_FILE = os.path.join(BASE_DIR, "generated_roster.xlsx")


def build_demo_employees():
    """Create a 30-person roster with realistic preferences."""
    employees = []

    # 10 Front Office staff
    fo_names = [
        ("Alicia", []),
        ("Ben", ["FB"]),
        ("Carlos", ["HK"]),
        ("Diana", []),
        ("Ethan", ["FB"]),
        ("Fiona", []),
        ("George", ["HK"]),
        ("Hannah", ["FB"]),
        ("Ian", []),
        ("Julia", ["HK"]),
    ]
    for i, (name, cross) in enumerate(fo_names, start=1):
        employees.append({"id": i, "name": f"FO_{name}", "role": "FO", "cross_train": cross})

    # 10 Housekeeping staff
    hk_names = [
        ("Kevin", ["FB"]),
        ("Luna", []),
        ("Mason", ["FO"]),
        ("Nina", []),
        ("Oscar", ["FB"]),
        ("Paula", []),
        ("Quinn", ["FO"]),
        ("Rachel", ["FB"]),
        ("Steve", []),
        ("Tina", ["FO"]),
    ]
    for i, (name, cross) in enumerate(hk_names, start=11):
        employees.append({"id": i, "name": f"HK_{name}", "role": "HK", "cross_train": cross})

    # 10 F&B staff
    fb_names = [
        ("Umar", ["HK"]),
        ("Vera", []),
        ("Will", ["FO"]),
        ("Xena", []),
        ("Yusuf", ["HK"]),
        ("Zoe", []),
        ("Adam", ["FO"]),
        ("Bella", ["HK"]),
        ("Chris", []),
        ("Dora", ["FO"]),
    ]
    for i, (name, cross) in enumerate(fb_names, start=21):
        employees.append({"id": i, "name": f"FB_{name}", "role": "FB", "cross_train": cross})

    return employees


def build_demo_preferences(employees):
    """
    Attach varied day-off and preferred-working-hour patterns.

    'Preferred hours' are modelled as OFF hours for non-preferred times.
    """
    prefs = {}

    # Helper to mark all hours except a given range as off
    def only_hours(start, end):
        return [h for h in hours if int(h) < start or int(h) > end]

    # Weekend-only workers
    weekend_only = [3, 14, 25]
    for eid in weekend_only:
        prefs[eid] = {d: None for d in days if d not in ["Sat", "Sun"]}

    # Night-shift-only preferences (work 22:00 - 06:00)
    night_only = [2, 13, 24]
    for eid in night_only:
        prefs[eid] = {d: only_hours(22, 6) for d in days}

    # Day-shift-only preferences (work 07:00 - 17:00)
    day_only = [5, 16, 27]
    for eid in day_only:
        prefs[eid] = {d: only_hours(7, 17) for d in days}

    # Specific whole days off
    prefs[1] = {"Sat": None, "Sun": None}
    prefs[4] = {"Wed": None}
    prefs[6] = {"Mon": None, "Fri": None}
    prefs[8] = {"Thu": None}
    prefs[9] = {"Tue": None, "Sun": None}
    prefs[11] = {"Mon": None}
    prefs[12] = {"Fri": None, "Sat": None}
    prefs[15] = {"Wed": None}
    prefs[17] = {"Sun": None}
    prefs[18] = {"Thu": None}
    prefs[19] = {"Sat": None}
    prefs[20] = {"Tue": None}
    prefs[21] = {"Sun": None}
    prefs[22] = {"Mon": None}
    prefs[23] = {"Fri": None}
    prefs[26] = {"Wed": None, "Sat": None}
    prefs[28] = {"Thu": None}
    prefs[29] = {"Mon": None, "Sun": None}
    prefs[30] = {"Tue": None}

    return prefs


def preferences_to_excel_cells(prefs, emp_id):
    """Convert internal preference format back to Excel cell strings."""
    cells = {}
    emp_prefs = prefs.get(emp_id, {})
    for day in days:
        if day not in emp_prefs:
            cells[day] = ""
        elif emp_prefs[day] is None:
            cells[day] = "OFF"
        else:
            cells[day] = ",".join(sorted(emp_prefs[day]))
    return cells


def build_employees_dataframe(employees, prefs):
    rows = []
    for e in employees:
        cells = preferences_to_excel_cells(prefs, e["id"])
        rows.append({
            "id": e["id"],
            "name": e["name"],
            "role": e["role"],
            "cross_train": ",".join(e.get("cross_train", [])),
            **cells,
        })
    return pd.DataFrame(rows)


def build_schedule_dataframe(schedule, employees):
    """Build a day x hour DataFrame with employee names in each cell."""
    data = []
    for day in days:
        row = {"Day": day}
        for hour in hours:
            emp_ids = schedule.get(day, {}).get(hour, [])
            names = [
                next((e["name"] for e in employees if e["id"] == eid), "???")
                for eid in emp_ids
            ]
            row[f"{hour}:00"] = ", ".join(names) if names else "—"
        data.append(row)
    return pd.DataFrame(data)


def print_summary(schedule, employees, prefs):
    total_assignments = sum(
        len(schedule[day][hour])
        for day in days
        for hour in hours
    )

    # Count hours per employee
    emp_hours = {e["id"]: 0 for e in employees}
    for day in days:
        for hour in hours:
            for eid in schedule[day][hour]:
                emp_hours[eid] += 1

    print("\n=== Demo Roster Summary ===")
    print(f"Employees: {len(employees)}")
    print(f"Total assignments (employee-hours): {total_assignments}")
    print(f"Average hours per employee: {total_assignments / len(employees):.1f}")
    print("\nTop 5 busiest employees:")
    sorted_emps = sorted(employees, key=lambda e: emp_hours[e["id"]], reverse=True)
    for e in sorted_emps[:5]:
        print(f"  {e['name']:10s} {emp_hours[e['id']]:3d} hours")

    # Weekly hours distribution
    print("\nEmployees with 0 hours ( Preferences blocked all slots ):")
    zero_hour_emps = [e for e in employees if emp_hours[e["id"]] == 0]
    for e in zero_hour_emps:
        print(f"  {e['name']}")
    if not zero_hour_emps:
        print("  None")


def main():
    employees = build_demo_employees()
    prefs = build_demo_preferences(employees)

    # Save demo employee input
    emp_df = build_employees_dataframe(employees, prefs)
    emp_df.to_excel(DEMO_EMPLOYEES_FILE, index=False)
    print(f"Saved demo employee dataset: {DEMO_EMPLOYEES_FILE}")

    # Generate schedule
    schedule = generate_schedule(employees, days, hours, prefs, STAFFING_NEEDS)

    # Validate
    errors = check_rules(schedule, employees, prefs, hours)
    print(f"\nRule check: {len(errors)} conflict(s)")
    for err in errors[:10]:
        print(f"  - {err}")

    # Save output roster
    roster_df = build_schedule_dataframe(schedule, employees)
    roster_df.to_excel(OUTPUT_ROSTER_FILE, index=False, sheet_name="Roster")
    print(f"\nSaved generated roster: {OUTPUT_ROSTER_FILE}")

    print_summary(schedule, employees, prefs)


if __name__ == "__main__":
    main()
