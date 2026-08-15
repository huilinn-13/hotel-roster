# scheduler.py
"""Greedy auto-scheduler for the hourly hotel roster."""

import calendar
from datetime import date


def _is_hour_off(preferences, emp_id, day, hour):
    """Return True if the employee requested this specific hour off."""
    emp_prefs = preferences.get(emp_id, {})
    missing = object()
    day_prefs = emp_prefs.get(day, missing)
    if day_prefs is missing and isinstance(day, str) and len(day) == 10:
        try:
            parsed_day = __import__("datetime").datetime.strptime(day, "%Y-%m-%d")
            day_prefs = emp_prefs.get(parsed_day.strftime("%a"), missing)
            if day_prefs is missing:
                # Monthly preference files may have been uploaded in a
                # different year. Match the month/day while preserving the
                # selected working year's schedule.
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


def _can_assign(eid, day, hour, role, employees, role_schedule, daily_hours,
                blocked, max_daily_hours, preferences):
    """Check if an employee can be assigned to cover a role need this hour."""
    emp = next((e for e in employees if e["id"] == eid), None)
    if emp is None:
        return False

    # Must be primary or cross-trained for this role
    hr_cover = role == "HR" and emp["role"] in {"AC", "MD"}
    if emp["role"] != role and role not in emp.get("cross_train", []) and not hr_cover:
        return False

    # Already covering any role this hour?
    if eid in role_schedule[day][hour].get("ALL", set()):
        return False

    # Day-off / hour-off preference
    if _is_hour_off(preferences, eid, day, hour):
        return False

    # Max daily hours
    if daily_hours[eid][day] >= max_daily_hours:
        return False

    # Max consecutive days
    if eid in blocked:
        return False

    return True


def generate_schedule(
    employees,
    days,
    hours,
    preferences,
    staffing_needs,
    max_daily_hours=9,
    max_consecutive_days=6,
):
    """
    Build a valid hourly roster using a greedy algorithm.

    Internally tracks which role each employee is covering so cross-trained
    staff are not double-counted across departments.

    Returns:
        dict: { day: { hour: [employee_ids] } }
    """
    emp_ids = [e["id"] for e in employees]

    # role_schedule[day][hour][role] = set of employee ids
    role_schedule = {
        day: {
            hour: {role: set() for role in staffing_needs}
            | {"ALL": set()}
            for hour in hours
        }
        for day in days
    }

    # Track daily hours per employee
    daily_hours = {eid: {day: 0 for day in days} for eid in emp_ids}

    # Track consecutive working days ending at the previous day
    consecutive = {eid: 0 for eid in emp_ids}

    for day in days:
        blocked = {eid for eid in emp_ids if consecutive[eid] >= max_consecutive_days}

        for hour in hours:
            for role, needs in staffing_needs.items():
                need = needs.get(hour, 0)
                if need <= 0:
                    continue

                still_need = need - len(role_schedule[day][hour][role])
                if still_need <= 0:
                    continue

                # --- Phase 1: primary-role employees ---
                primary_candidates = [
                    e for e in employees
                    if e["role"] == role
                    and _can_assign(
                        e["id"], day, hour, role, employees, role_schedule,
                        daily_hours, blocked, max_daily_hours, preferences
                    )
                ]
                # Fill a person's current duty block before moving to the next
                # person. Round-robin hourly assignment creates scattered
                # cells such as 02 and 23 and makes every employee work daily.
                primary_candidates.sort(
                    key=lambda e: (-daily_hours[e["id"]][day], sum(daily_hours[e["id"]].values()))
                )

                assigned_here = []
                for e in primary_candidates:
                    if len(assigned_here) >= still_need:
                        break
                    assigned_here.append(e["id"])

                # --- Phase 2: cross-trained employees if still short ---
                if len(assigned_here) < still_need:
                    cross_candidates = [
                        e for e in employees
                        if (role in e.get("cross_train", []) or (role == "HR" and e["role"] in {"AC", "MD"}))
                        and e["id"] not in assigned_here
                        and _can_assign(
                            e["id"], day, hour, role, employees, role_schedule,
                            daily_hours, blocked, max_daily_hours, preferences
                        )
                    ]
                    cross_candidates.sort(
                        key=lambda e: (-daily_hours[e["id"]][day], sum(daily_hours[e["id"]].values()))
                    )

                    for e in cross_candidates:
                        if len(assigned_here) >= still_need:
                            break
                        assigned_here.append(e["id"])

                # Record assignments
                for eid in assigned_here:
                    role_schedule[day][hour][role].add(eid)
                    role_schedule[day][hour]["ALL"].add(eid)
                    daily_hours[eid][day] += 1

        # Update consecutive working-day counters for next day
        for eid in emp_ids:
            worked_today = any(
                eid in role_schedule[day][hour]["ALL"]
                for hour in hours
            )
            if worked_today:
                consecutive[eid] += 1
            else:
                consecutive[eid] = 0

    # Flatten to the public schedule format
    schedule = {day: {hour: [] for hour in hours} for day in days}
    for day in days:
        for hour in hours:
            assigned = set()
            for role in staffing_needs:
                assigned.update(role_schedule[day][hour][role])
            schedule[day][hour] = list(assigned)

    return schedule


def generate_month_schedule(employees, month_start, preferences, staffing_needs):
    """Generate a date-aware monthly roster while preserving sequential-day limits."""
    month_start = month_start.replace(day=1)
    date_days = [
        date(month_start.year, month_start.month, d).strftime("%Y-%m-%d")
        for d in range(1, calendar.monthrange(month_start.year, month_start.month)[1] + 1)
    ]
    return generate_schedule(
        employees,
        date_days,
        [f"{h:02d}" for h in range(24)],
        preferences,
        staffing_needs,
    )
