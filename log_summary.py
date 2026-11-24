"""
Summarise the connection problems found in bridge.log (read-only, changes nothing).

    python log_summary.py            # reads bridge.log in this folder
    python log_summary.py other.log
"""
import re
import sys
from datetime import datetime

path = sys.argv[1] if len(sys.argv) > 1 else "bridge.log"
LINE = re.compile(r"^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d),\d+ (\w+) (.*)$")

outages, start, punches, retries = [], None, 0, 0
try:
    f = open(path, encoding="utf-8", errors="replace")
except FileNotFoundError:
    sys.exit("%s was not found. Run this from the zk-bridge folder." % path)

with f:
    for raw in f:
        m = LINE.match(raw.strip())
        if not m:
            continue
        t = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
        msg = m.group(3)
        if "device connection problem" in msg:
            retries += 1
            if start is None:
                start = t
        elif msg.startswith("connected"):
            if start is not None:
                outages.append((start, t))
                start = None
        elif msg.startswith("punch "):
            punches += 1
if start is not None:
    outages.append((start, None))


def fmt(seconds):
    m, s = divmod(int(seconds), 60)
    return "%d min %02d s" % (m, s)


print("Punches captured in this log : %d" % punches)
print("Connection problems (outages): %d" % len(outages))
total = 0
for a, b in outages:
    if b is None:
        print("  %s  ->  still not connected" % a.strftime("%Y-%m-%d %H:%M:%S"))
    else:
        secs = (b - a).total_seconds()
        total += secs
        print("  %s  ->  %s   (%s)" % (a.strftime("%Y-%m-%d %H:%M:%S"), b.strftime("%H:%M:%S"), fmt(secs)))
if outages:
    print("Total time without connection: %s" % fmt(total))
    print("Note: each outage really began a little BEFORE its first line here, so the real gaps are slightly longer.")
