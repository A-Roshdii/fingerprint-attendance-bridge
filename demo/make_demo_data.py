"""
Try the whole tool WITHOUT a fingerprint device.

    python demo/make_demo_data.py

It invents 11 employees and one day of punches (a normal day, a late arrival, a half day, a night shift,
a forgotten punch, managers, overtime...), saves them in demo_data/punches.db and then runs export_sheet.py on
them. Open demo_data/attendance_2026-03-12.xlsx and demo_data/attendance_2026-03-12_review.xlsx.
Everything here is made up: no real people, no real company.
"""
import csv
import os
import subprocess
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import bridge  # noqa: E402  (reuses the same database code as the real tool)

OUT = os.path.join(ROOT, "demo_data")
os.makedirs(OUT, exist_ok=True)
for name in ("punches.db", "punches.db-wal", "punches.db-shm"):
    path = os.path.join(OUT, name)
    if os.path.exists(path):
        os.remove(path)

employees = [  # device number, employee code, name, role
    (1, "1000001", "Employee 01", ""), (2, "1000002", "Employee 02", ""), (3, "1000003", "Employee 03", ""),
    (4, "1000004", "Employee 04", "rotating"), (5, "1000005", "Employee 05", ""), (6, "1000006", "Employee 06", ""),
    (7, "1000007", "Employee 07", ""), (8, "1000008", "Employee 08", "manager"), (9, "1000009", "Employee 09", "manager"),
    (10, "1000010", "Employee 10", ""), (11, "1000011", "Employee 11", ""),
]
with open(os.path.join(OUT, "mapping.csv"), "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["device_id", "employee_code", "name", "role"])
    writer.writerows(employees)
with open(os.path.join(OUT, "config.ini"), "w", encoding="utf-8") as f:
    f.write("[device]\nip = 192.168.1.201\n[bridge]\ndb_path = punches.db\nmapping_file = mapping.csv\n")

db = bridge.open_db(os.path.join(OUT, "punches.db"))


def punch(dev, day, hour, minute, second=0):
    bridge.save_punch(db, str(dev), datetime(2026, 3, day, hour, minute, second), 0, "Employee %02d" % dev)


punch(1, 12, 7, 58, 12); punch(1, 12, 17, 5, 40)       # normal day shift
punch(2, 12, 8, 7, 3); punch(2, 12, 17, 10, 22)        # arrived 7 minutes late
punch(3, 12, 7, 56, 30); punch(3, 12, 12, 30, 15)      # half day (about 4.5 hours)
punch(4, 12, 22, 56, 8); punch(4, 13, 8, 3, 50)        # rotating employee: night shift across midnight
punch(5, 12, 13, 55, 40); punch(5, 12, 23, 2, 11)      # middle shift
punch(6, 12, 8, 0, 20); punch(6, 12, 16, 20, 5)        # left 40 minutes early
punch(7, 12, 8, 2, 17)                                 # forgotten punch: only one punch
punch(8, 12, 8, 52, 9); punch(8, 12, 17, 30, 33)       # manager: 08:52 is fine (managers may come until 09:00)
punch(9, 12, 9, 12, 41); punch(9, 12, 17, 40, 2)       # manager: after 09:00 -> late
punch(10, 12, 7, 57, 5); punch(10, 12, 20, 1, 44)      # a 12-hour day: overtime for a regular employee
punch(11, 12, 7, 54, 59); punch(11, 12, 17, 3, 3)      # normal day shift
db.close()

print("Demo data created in", OUT)
subprocess.run([sys.executable, os.path.join(ROOT, "export_sheet.py"), "2026-03-12"], cwd=OUT, check=False)
print("\nOpen the two Excel files in:", OUT)
