"""
Shows every punch of ONE employee that is saved in punches.db (read-only).

    python show_punches.py 1 2026-03-10 2026-03-12
                           ^device number or 7-digit code   ^from      ^to (included)
"""
import csv
import os
import sqlite3
import sys
from datetime import datetime, timedelta

if len(sys.argv) != 4:
    sys.exit(__doc__)
who = sys.argv[1].strip()
try:
    start = datetime.strptime(sys.argv[2], "%Y-%m-%d")
    end = datetime.strptime(sys.argv[3], "%Y-%m-%d") + timedelta(days=1)
except ValueError:
    sys.exit("Write the dates like 2026-11-04")

device_id, code, name = who, "", ""
if os.path.exists("mapping.csv"):
    with open("mapping.csv", newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if len(row) >= 2 and who in (row[0].strip(), row[1].strip()):
                device_id, code = row[0].strip(), row[1].strip()
                name = row[2].strip() if len(row) > 2 else ""

c = sqlite3.connect("punches.db")
rows = c.execute("SELECT punched_at FROM punches WHERE device_user_id = ? AND punched_at >= ? AND punched_at < ?"
                 " ORDER BY punched_at", (device_id, start.strftime("%Y-%m-%d %H:%M:%S"),
                                          end.strftime("%Y-%m-%d %H:%M:%S"))).fetchall()
print("Device number %s | code %s | %s" % (device_id, code or "-", name or "-"))
print("Punches saved in punches.db from %s to %s: %d" % (sys.argv[2], sys.argv[3], len(rows)))
for (t,) in rows:
    print("  ", t)
