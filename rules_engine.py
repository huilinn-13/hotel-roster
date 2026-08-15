# rules_engine.py
"""Rule checker for the Hotel Smart Roster (hourly version)."""


def _get_employee_name(employees, emp_id):
    """Helper to look up an employee's name by id."""
    return next((e["name"] for e in employees if e["id"] == emp_id), "Unknown")


def _is_hour_off(preferences, emp_id, day, hour):
    """Return True if the employee requested this specific hour off."""
    emp_prefs = preferences.get(emp_id, {})
    missing = object()
    day_prefs = emp_prefs.get(day, missing)
    if day_prefs is missing and isinstance(day, str) and len(day) == 10:
        try:
            from datetime import datetime
            parsed_day = datetime.strptime(day, "%Y-%m-%d")
            day_prefs = emp_prefs.get(parsed_day.strftime("%a"), missing)
            if day_prefs is missing:
                month_day = day[5:]
                for pref_day, pref_value in emp_prefs.items():
                    if isinstance(pref_day, str) and len(pref_day) == 10 and pref_day[5:] == month_day:
                        day_prefs = pref_value
                        break
        except ValueError:
            day_prefs = missing
    if day_prefs is missing:
        return False
    if day_prefs is None:
        return True  # whole day off
    if isinstance(day_prefs, list):
        return hour in day_prefs
    return False


def check_rules(schedule, employees, preferences, hours):
    """
    Validate an hourly roster.

    Schedule format: { day: { hour: [employee_ids] } }

    Returns:
        list of error/warning strings.
    """
    errors = []
    emp_ids = [e["id"] for e in employees]
    schedule_days = list(schedule.keys())

    # --- Rule 1: Max 6 consecutive working days ---
    for emp_id in emp_ids:
        emp_name = _get_employee_name(employees, emp_id)
        consecutive = 0
        for day in schedule_days:
            day_schedule = schedule.get(day, {})
            is_working = any(
                emp_id in day_schedule.get(h, []) for h in hours
            )

            if is_working:
                consecutive += 1
            else:
                consecutive = 0

            if consecutive > 6:
                errors.append(
                    f"⚠️ {emp_name} is working more than 6 consecutive days!"
                )

    # --- Rule 2: Max 9 hours per day and no duplicate hour assignments ---
    for emp_id in emp_ids:
        emp_name = _get_employee_name(employees, emp_id)
        for day in schedule_days:
            day_schedule = schedule.get(day, {})
            assigned_hours = []
            duplicates = set()
            for hour in hours:
                hour_list = day_schedule.get(hour, [])
                count = hour_list.count(emp_id)
                if count > 0:
                    assigned_hours.append(hour)
                if count > 1:
                    duplicates.add(hour)

            if duplicates:
                errors.append(
                    f"🚨 {emp_name} is assigned multiple times on {day} at "
                    f"hour(s) {', '.join(sorted(duplicates))}."
                )

            if len(assigned_hours) > 9:
                errors.append(
                    f"⏰ {emp_name} is scheduled {len(assigned_hours)} hours on "
                    f"{day} (max 9)."
                )

    # --- Rule 3: Respect Preferences (Soft Rule) ---
    for emp_id, emp_prefs in preferences.items():
        emp_name = _get_employee_name(employees, emp_id)
        for day, day_prefs in emp_prefs.items():
            day_schedule = schedule.get(day, {})
            for hour in hours:
                if emp_id in day_schedule.get(hour, []):
                    if _is_hour_off(preferences, emp_id, day, hour):
                        if day_prefs is None:
                            errors.append(
                                f"💡 {emp_name} requested {day} off but is "
                                f"scheduled at {hour}:00."
                            )
                        else:
                            errors.append(
                                f"💡 {emp_name} requested {day} {hour}:00 off "
                                f"but is scheduled."
                            )

    return errors
