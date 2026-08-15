# roster_data.py
"""Default data and constants for the Hotel Smart Roster MVP."""

# 1. Define Roles
ROLES = {
    "FO": "Front Office",
    "HK": "Housekeeping",
    "FB": "Food & Beverage",
    "EN": "Engineering",
    "HR": "Human Resources",
}

# 2. Mock Employees (30 people)
# 'role' is their primary department.
# 'cross_train' means they can be scheduled in those departments too.
employees = [
    {"id": 1, "name": "Emp_01", "role": "FO", "cross_train": ["FB"]},
    {"id": 2, "name": "Emp_02", "role": "FO", "cross_train": []},
    {"id": 3, "name": "Emp_03", "role": "FO", "cross_train": ["HK"]},
    {"id": 4, "name": "Emp_04", "role": "FO", "cross_train": []},
    {"id": 5, "name": "Emp_05", "role": "FO", "cross_train": ["FB"]},
    {"id": 6, "name": "Emp_06", "role": "FO", "cross_train": []},
    {"id": 7, "name": "Emp_07", "role": "FO", "cross_train": ["HK"]},
    {"id": 8, "name": "Emp_08", "role": "FO", "cross_train": []},
    {"id": 9, "name": "Emp_09", "role": "FO", "cross_train": ["FB"]},
    {"id": 10, "name": "Emp_10", "role": "FO", "cross_train": []},
    {"id": 11, "name": "Emp_11", "role": "HK", "cross_train": ["FB"]},
    {"id": 12, "name": "Emp_12", "role": "HK", "cross_train": []},
    {"id": 13, "name": "Emp_13", "role": "HK", "cross_train": ["FO"]},
    {"id": 14, "name": "Emp_14", "role": "HK", "cross_train": []},
    {"id": 15, "name": "Emp_15", "role": "HK", "cross_train": ["FB"]},
    {"id": 16, "name": "Emp_16", "role": "HK", "cross_train": []},
    {"id": 17, "name": "Emp_17", "role": "HK", "cross_train": ["FO"]},
    {"id": 18, "name": "Emp_18", "role": "HK", "cross_train": []},
    {"id": 19, "name": "Emp_19", "role": "HK", "cross_train": ["FB"]},
    {"id": 20, "name": "Emp_20", "role": "HK", "cross_train": []},
    {"id": 21, "name": "Emp_21", "role": "FB", "cross_train": ["HK"]},
    {"id": 22, "name": "Emp_22", "role": "FB", "cross_train": []},
    {"id": 23, "name": "Emp_23", "role": "FB", "cross_train": ["FO"]},
    {"id": 24, "name": "Emp_24", "role": "FB", "cross_train": []},
    {"id": 25, "name": "Emp_25", "role": "FB", "cross_train": ["HK"]},
    {"id": 26, "name": "Emp_26", "role": "FB", "cross_train": []},
    {"id": 27, "name": "Emp_27", "role": "FB", "cross_train": ["FO"]},
    {"id": 28, "name": "Emp_28", "role": "FB", "cross_train": []},
    {"id": 29, "name": "Emp_29", "role": "FB", "cross_train": ["HK"]},
    {"id": 30, "name": "Emp_30", "role": "FB", "cross_train": []},
]

# 3. Define the planning horizon
days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# 24 one-hour slots, formatted as "HH" (00 .. 23).
hours = [f"{h:02d}" for h in range(24)]

# 4. Staffing requirements. The hotel specifies shift *start* times; the
# hourly engine expands each start into a 9-hour duty window (including the
# break), carrying overnight shifts into the next day.
def _shift_coverage(shifts, length=9):
    coverage = {h: 0 for h in hours}
    for start, count in shifts:
        for offset in range(length):
            coverage[f"{(start + offset) % 24:02d}"] += count
    return coverage


STAFFING_NEEDS = {
    "FO": _shift_coverage([(6, 2), (13, 2), (22, 1)]),
    "FB": _shift_coverage([(6, 2), (11, 2)]),
    "HK": _shift_coverage([(6, 1), (7, 1), (8, 1)]),
    "EN": _shift_coverage([(7, 1), (13, 1), (22, 1)]),
    "AT": _shift_coverage([(22, 1)]),
    # HR is covered by AC/MD rather than a separate HR department.
    "HR": _shift_coverage([(8, 1)]),
}

# 5. Staff Preferences (Who wants off?)
# Format: { employee_id: { day: None | [hours] } }
#   None       => whole day off
#   ["08","09"]=> those specific hours off
preferences = {
    1: {"Sat": None, "Sun": None},
    5: {"Wed": None},
    11: {"Mon": None},
    18: {"Fri": None, "Sat": None},
    25: {"Sun": None},
    30: {"Thu": None},
}
