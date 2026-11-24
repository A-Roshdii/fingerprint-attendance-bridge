#!/usr/bin/env python3
"""
Build the attendance sheet (Excel) from the punches saved in punches.db
(by bridge.py or pull_and_export.py). The layout of the sheet is set in config.ini, section [sheet].

    python export_sheet.py 2026-10-29                 # one day
    python export_sheet.py 2026-10-27 2026-10-29      # several days

THE SHEET (no colours, nothing is added or guessed)
  - For each employee, punches less than 13 hours apart belong to ONE shift
    (13 so that 12-hour overtime shifts stay in one piece; normal rests between shifts are about 14-15 hours).
  - First punch of the shift = IN, last punch = OUT.
  - A shift is listed under the day it STARTS (its IN). So the night shift that starts at 23:00 on the
    4th and ends at 08:00 on the 5th is in the sheet of the 4th (IN dated the 4th, OUT dated the 5th).
    (config.ini, [export]  shift_belongs_to = end  -> list it under the day of its OUT instead.)
  - Only employees listed in mapping.csv are exported. Any other device ID is IGNORED
    (listed in the "Ignored" sheet of *_review.xlsx so nothing disappears silently).
    Optional, in config.ini:   [export]  unknown_devices = temporary   -> unknown IDs get a temporary code.
  - A shift with a single punch (IN or OUT forgotten, or the shift is still open) is NOT in the sheet:
    it is only listed in the Review sheet.

THE REVIEW SHEET (attendance_..._review.xlsx, sheet "Review") warns you about anything unusual:
  Forgotten punch   only one punch (nothing is guessed: it can be a forgotten IN, a forgotten OUT or an open shift)
  Late IN           arrived even one minute after the scheduled start (seconds are ignored)
  Early leave       left before the scheduled end
  Half day          only about 4.5 hours (3.5 to 5.5) between IN and OUT
  Unusual shift     does not match the shifts this employee works, or stayed more than an hour after the end (overtime?)
  Long shift        more than 14 hours (probably two shifts merged)

Shifts: 08:00-17:00, 14:00-23:00, 23:00-08:00 (next day) and the 12-hour shift (08:00-20:00 / 20:00-08:00).
The 4th column of mapping.csv ("role") tells who works what (mapping.csv: device_id,employee_code,name,role):
  rotating   changes shifts and also works nights and the 12-hour shift (all of the shifts above are normal)
  manager    may punch IN until 09:00 without being late, and a normal day is up to 9 hours after the IN
  (empty)    a regular employee: day 08:00-17:00 or middle 14:00-23:00 only. A night shift or a 12-hour shift
             is reported as unusual
"""
import configparser
import csv
import os
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime, time, timedelta

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("The openpyxl library is missing. Run:  pip install openpyxl  (needs open internet), then try again.")

FMT = "%Y-%m-%d %H:%M:%S"
TOKENS = {"code", "type", "type_name", "datetime", "date", "blank"}


def load_template(cfg):
    """How the exported sheet looks. Everything can be changed in config.ini, section [sheet]."""
    get = lambda key, default: cfg.get("sheet", key, raw=True, fallback=default)
    columns = [c.strip().lower() for c in get("columns", "code, type, datetime").split(",")]
    headers = [h.strip() for h in get("headers", "Employee Code, In/Out, Date and Time").split(",")]
    if len(columns) != len(headers) or any(c not in TOKENS for c in columns):
        sys.exit("config.ini [sheet]: 'columns' and 'headers' must have the same number of items, and the columns "
                 "can only be: %s" % ", ".join(sorted(TOKENS)))
    return {
        "columns": columns,
        "headers": headers,
        "type_name_format": get("type_name_format", "{kind}"),
        "date_format": get("date_format", "%Y-%m-%d %H:%M"),
        "header_row": 2 if get("header_row", "1").strip() == "2" else 1,
        "note": get("note", "").strip(),
        "sheet_name": get("sheet_name", "Sheet1").strip() or "Sheet1",
    }


# ---- what a normal shift looks like (change here if the working hours change) ----
DAY9 = ("day shift 08:00-17:00", time(8, 0), 9)
MIDDLE9 = ("middle shift 14:00-23:00", time(14, 0), 9)
NIGHT9 = ("night shift 23:00-08:00", time(23, 0), 9)
DAY12 = ("12-hour day shift 08:00-20:00", time(8, 0), 12)
NIGHT12 = ("12-hour night shift 20:00-08:00", time(20, 0), 12)
TYPES_ROTATING = (DAY9, MIDDLE9, NIGHT9, DAY12, NIGHT12)
TYPES_REGULAR = (DAY9, MIDDLE9)
HALF_DAY_HOURS, HALF_DAY_TOLERANCE = 4.5, 1.0   # about 4.5 hours between IN and OUT = half a day
MAX_FIT_MINUTES = 200      # IN + OUT may be at most this many minutes away from a normal shift's start + end
OVERSTAY_MINUTES = 60      # leaving more than this after the scheduled end is reported (overtime?)
MANAGER_LATEST_IN = time(9, 0)   # managers may punch IN until 09:00
MANAGER_HOURS = 9                # a manager's normal day: up to 9 hours after the IN
ROTATING_WORDS = {"rotating", "yes", "y", "1", "true", "night9", "night"}
MANAGER_WORDS = {"manager", "mgr"}


def parse_day(text):
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, f)
        except ValueError:
            pass
    sys.exit("Could not read the date '%s'. Write it like 2026-10-29" % text)


def load_mapping(path):
    """mapping.csv: device_id,employee_code,name,role  ->  (codes, names, roles) by device id."""
    codes, names, roles = {}, {}, {}
    if path and os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and row[0].strip() and not row[0].strip().lower().startswith("device"):
                    dev = row[0].strip()
                    codes[dev] = row[1].strip()
                    if len(row) > 2 and row[2].strip():
                        names[dev] = row[2].strip()
                    word = row[3].strip().lower() if len(row) > 3 else ""
                    roles[dev] = "manager" if word in MANAGER_WORDS else "rotating" if word in ROTATING_WORDS else "regular"
    return codes, names, roles


def placeholder(dev_id, taken):
    """Temporary code for a device ID that has no employee code yet."""
    if dev_id.isdigit() and len(dev_id) <= 5:
        number = "99" + dev_id.zfill(5)
        if number not in taken:
            return number
    return None


def sort_key(dev_id):
    return (0, int(dev_id)) if dev_id.isdigit() else (1, dev_id)


def save_book(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        sys.exit("Could not write %s - close it in Excel first, then run again." % path)


def to_minute(t):
    return t.replace(second=0, microsecond=0)


def best_type(first, last, types):
    """The normal shift this IN/OUT is closest to: (distance in minutes, name, scheduled start, scheduled end)."""
    best = None
    for name, start_time, length in types:
        for off in (-1, 0, 1):
            sched_start = datetime.combine(first.date() + timedelta(days=off), start_time)
            sched_end = sched_start + timedelta(hours=length)
            dist = abs((first - sched_start).total_seconds()) / 60.0 + abs((last - sched_end).total_seconds()) / 60.0
            if best is None or dist < best[0]:
                best = (dist, name, sched_start, sched_end)
    return best


def nearest_start(first, types):
    """The scheduled start closest to this IN (used when the shift looks like a half day)."""
    best = None
    for _name, start_time, _length in types:
        for off in (-1, 0, 1):
            sched = datetime.combine(first.date() + timedelta(days=off), start_time)
            if best is None or abs((first - sched).total_seconds()) < abs((first - best).total_seconds()):
                best = sched
    return best


def hm(minutes):
    return "%d h %02d min" % (minutes // 60, minutes % 60) if minutes >= 60 else "%d minute(s)" % minutes


def check_shift(first, last, role):
    """Everything unusual about one complete shift: a list of (type, details)."""
    hours = (last - first).total_seconds() / 3600.0
    span = "IN %s, OUT %s" % (first.strftime("%Y-%m-%d %H:%M"), last.strftime("%Y-%m-%d %H:%M"))
    found = []
    half = abs(hours - HALF_DAY_HOURS) <= HALF_DAY_TOLERANCE
    if half:
        found.append(("Half day", "%.1f hours between IN and OUT - %s" % (hours, span)))

    if role == "manager":
        if not (5 <= first.hour < 12):
            found.append(("Unusual shift", "A manager came at %s (a manager normally comes in the morning) - %s"
                          % (first.strftime("%H:%M"), span)))
            return found
        latest = datetime.combine(first.date(), MANAGER_LATEST_IN)
        late = int((to_minute(first) - latest).total_seconds() // 60)
        if late > 0:
            found.append(("Late IN", "Manager arrived %s, the latest is %s: %s late"
                          % (first.strftime("%H:%M"), MANAGER_LATEST_IN.strftime("%H:%M"), hm(late))))
        extra = int((to_minute(last) - (to_minute(first) + timedelta(hours=MANAGER_HOURS))).total_seconds() // 60)
        if extra > OVERSTAY_MINUTES:
            found.append(("Unusual shift", "Manager stayed %.1f hours (normal is up to %d hours after the IN, overtime?) - %s"
                          % (hours, MANAGER_HOURS, span)))
        return found

    types = TYPES_ROTATING if role == "rotating" else TYPES_REGULAR
    if role == "regular" and first.hour >= 19 and not half:
        found.append(("Unusual shift", "A night-type shift (IN %s) for an employee who does not work nights - %s"
                      % (first.strftime("%H:%M"), span)))
        return found
    dist, name, sched_start, sched_end = best_type(first, last, types)
    if half:  # a half day: judge lateness by the IN alone
        sched_start = nearest_start(first, types)
        late = max(int((to_minute(first) - sched_start).total_seconds() // 60), 0)
        if late > 0 and late < 12 * 60:
            found.append(("Late IN", "Arrived %s, scheduled %s: %s late" % (first.strftime("%H:%M"), sched_start.strftime("%H:%M"), hm(late))))
        return found
    if dist > MAX_FIT_MINUTES:
        found.append(("Unusual shift", "Does not match any shift of this employee (%.1f hours) - %s" % (hours, span)))
        return found
    late = int((to_minute(first) - sched_start).total_seconds() // 60)
    if late > 0:
        found.append(("Late IN", "Arrived %s, the %s starts at %s: %s late"
                      % (first.strftime("%H:%M"), name, sched_start.strftime("%H:%M"), hm(late))))
    early = int((sched_end - to_minute(last)).total_seconds() // 60)
    if early > 0:
        found.append(("Early leave", "Left %s, the %s ends at %s: %s early - %s"
                      % (last.strftime("%H:%M"), name, sched_end.strftime("%H:%M"), hm(early), span)))
    over = int((to_minute(last) - sched_end).total_seconds() // 60)
    if over > OVERSTAY_MINUTES:
        found.append(("Unusual shift", "Stayed %s after the end of the %s at %s (overtime?) - %s"
                      % (hm(over), name, sched_end.strftime("%H:%M"), span)))
    return found


def main():
    args = sys.argv[1:]
    if not 1 <= len(args) <= 2:
        sys.exit(__doc__)
    start, end = parse_day(args[0]), parse_day(args[-1])
    if end < start:
        sys.exit("The end date is before the start date.")
    range_end = end + timedelta(days=1)  # exclusive

    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read("config.ini", encoding="utf-8")
    db_path = cfg.get("bridge", "db_path", fallback="punches.db")
    mapping_file = cfg.get("bridge", "mapping_file", fallback="mapping.csv")
    gap = timedelta(hours=cfg.getfloat("export", "shift_gap_hours", fallback=13))
    min_shift = timedelta(minutes=cfg.getfloat("export", "min_shift_minutes", fallback=60))
    use_temp = cfg.get("export", "unknown_devices", fallback="ignore").strip().lower() == "temporary"
    by_end = cfg.get("export", "shift_belongs_to", fallback="start").strip().lower() != "start"
    max_shift = timedelta(hours=cfg.getfloat("export", "max_shift_hours", fallback=14))

    tpl = load_template(cfg)
    mapping, map_names, roles = load_mapping(mapping_file)
    if not mapping:
        sys.exit("%s was not found (or is empty). Put it in this folder first." % mapping_file)
    if not os.path.exists(db_path):
        sys.exit("%s was not found. Run this from the zk-bridge folder." % db_path)

    c = sqlite3.connect(db_path)
    c.row_factory = sqlite3.Row
    first_in_db = c.execute("SELECT MIN(punched_at) AS m FROM punches").fetchone()["m"]
    # read one day before and two days after so shifts that cross midnight are complete
    rows = c.execute(
        "SELECT device_user_id, employee_name, punched_at FROM punches"
        " WHERE punched_at >= ? AND punched_at < ? ORDER BY punched_at",
        ((start - timedelta(days=1)).strftime(FMT), (end + timedelta(days=2)).strftime(FMT)),
    ).fetchall()

    punches, device_names = defaultdict(list), {}
    for r in rows:
        punches[r["device_user_id"]].append(datetime.strptime(r["punched_at"], FMT))
        if r["employee_name"]:
            device_names[r["device_user_id"]] = r["employee_name"]
    name_of = lambda dev: map_names.get(dev) or device_names.get(dev, "")  # the name you wrote in mapping.csv wins

    good = []            # (sort key, first, employee code, first punch, last punch)
    review = []          # (employee, name, time, type, details)
    unmapped = defaultdict(int)
    temp_used = {}
    ignored = defaultdict(int)
    real_numbers = set(mapping.values())
    for dev_id, times in punches.items():
        times.sort()
        shifts = [[times[0]]]
        for t in times[1:]:
            if t - shifts[-1][-1] >= gap:
                shifts.append([t])
            else:
                shifts[-1].append(t)
        for s in shifts:
            first, last = s[0], s[-1]
            if not (start <= (last if by_end else first) < range_end):
                continue  # this shift belongs to another day
            code = mapping.get(dev_id)
            if code is None:
                if not use_temp:
                    ignored[dev_id] += 1  # not in mapping.csv = not one of your employees
                    continue
                code = placeholder(dev_id, real_numbers)
                if code is None:  # cannot even build a temporary number
                    unmapped[dev_id] += 1
                    continue
                temp_used[dev_id] = code
            name = name_of(dev_id)
            when = first.strftime("%Y-%m-%d %H:%M")
            if last - first < min_shift:   # one punch, or two punches almost at the same time
                review.append((code, name, when, "Forgotten punch",
                               "Only one punch, at %s: the IN or the OUT is missing, or the shift is still open. "
                               "NOT in the sheet" % first.strftime("%Y-%m-%d %H:%M")))
                continue
            if last - first > max_shift:
                review.append((code, name, when, "Long shift",
                               "%.1f hours, %d punches - two shifts may have been merged (overtime?), in the sheet as it is"
                               % ((last - first).total_seconds() / 3600, len(s))))
            else:
                for kind, details in check_shift(first, last, roles.get(dev_id, "regular")):
                    review.append((code, name, when, kind, details))
            good.append((sort_key(dev_id), first, code, first, last))

    for dev_id, tmp in temp_used.items():
        review.append((tmp, name_of(dev_id), "", "Temporary number",
                       "Device ID %s has no employee code in mapping.csv" % dev_id))
    for dev_id, n in unmapped.items():
        review.append((dev_id + " (device ID)", name_of(dev_id), "", "No employee code",
                       "Not in mapping.csv (%d shift%s not exported)" % (n, "" if n == 1 else "s")))

    label = args[0] if args[0] == args[-1] else "%s_to_%s" % (args[0], args[-1])
    good.sort(key=lambda x: (x[0], x[1]))

    if good:
        wb = Workbook()
        ws = wb.active
        ws.title = tpl["sheet_name"]
        ncols, hr = len(tpl["columns"]), tpl["header_row"]
        if tpl["note"]:
            ws.cell(1, ncols, tpl["note"]).font = Font(name="Calibri", size=12, bold=True)
        for i, h in enumerate(tpl["headers"], start=1):
            cell = ws.cell(hr, i, h)
            cell.font = Font(name="Calibri", size=12, bold=True)
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[cell.column_letter].width = max(14, len(h) + 4)
        r = hr + 1
        for _, _, code, first, last in good:
            for kind, when in (("IN", first), ("OUT", last)):
                values = {"code": code, "type": kind, "type_name": tpl["type_name_format"].format(kind=kind),
                          "datetime": when.strftime(tpl["date_format"]), "date": when.strftime("%Y-%m-%d"), "blank": None}
                for col, token in enumerate(tpl["columns"], start=1):
                    cell = ws.cell(r, col, values[token])  # everything is written as text
                    cell.font = Font(name="Calibri", size=12)
                    cell.alignment = Alignment(horizontal="center")
                r += 1
        ws.freeze_panes = "A%d" % (hr + 1)
        ws.auto_filter.ref = "A%d:%s%d" % (hr, get_column_letter(ncols), r - 1)
        out_path = "attendance_%s.xlsx" % label
        save_book(wb, out_path)
        print("Done: %d shift(s) = %d rows  ->  %s" % (len(good), len(good) * 2, out_path))
    else:
        print("No complete shifts found for these days.")

    if review or ignored:
        order = {"Forgotten punch": 0, "Late IN": 1, "Early leave": 2, "Half day": 3, "Unusual shift": 4, "Long shift": 5}
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "Review"
        ws2.append(["Employee", "Name", "Time", "Type", "Details"])
        for row in sorted(review, key=lambda x: (order.get(x[3], 9), str(x[0]), x[2])):
            ws2.append(list(row))
        for col, w in zip("ABCDE", (24, 30, 18, 18, 110)):
            ws2.column_dimensions[col].width = w
        for cell in ws2[1]:
            cell.font = Font(bold=True)
        ws3 = wb2.create_sheet("Ignored")
        ws3.append(["Device number", "Name on device", "Shifts not exported"])
        for dev_id in sorted(ignored, key=sort_key):
            ws3.append([dev_id, device_names.get(dev_id, ""), ignored[dev_id]])
        for col, w in zip("ABC", (16, 30, 20)):
            ws3.column_dimensions[col].width = w
        for cell in ws3[1]:
            cell.font = Font(bold=True)
        rpath = "attendance_%s_review.xlsx" % label
        save_book(wb2, rpath)
        if review:
            kinds = defaultdict(int)
            for row in review:
                kinds[row[3]] += 1
            print("Needs your review: %d item(s)  ->  %s  (sheet 'Review')" % (len(review), rpath))
            print("  " + ", ".join("%s: %d" % (k, v) for k, v in sorted(kinds.items(), key=lambda kv: order.get(kv[0], 9))))
        if ignored:
            print("Ignored %d device number(s) that are not in mapping.csv (listed in sheet 'Ignored' of %s)"
                  % (len(ignored), rpath))

    if first_in_db and start < parse_day(first_in_db[:10]) + timedelta(days=1):
        print("Note: punches.db only starts at %s, so days before/around it can be incomplete."
              % first_in_db)


if __name__ == "__main__":
    main()
