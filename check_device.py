"""
Diagnostic only - reads from the ZKTeco device, changes nothing.
Stop bridge.py first (Ctrl + C), then run:  python check_device.py
"""
import configparser
from datetime import datetime

from zk import ZK

cfg = configparser.ConfigParser()
cfg.read("config.ini", encoding="utf-8")
ip = cfg["device"]["ip"]
port = int(cfg["device"].get("port", 4370))

print("Connecting to %s:%s ..." % (ip, port))
conn = ZK(ip, port=port, timeout=15).connect()
print("Connected.")
print("This computer's clock:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

try:
    print("Device clock         :", conn.get_time())
except Exception as e:
    print("Could not read the device clock:", e)

records = None
try:
    records = conn.get_attendance()
except Exception as e:
    print("ERROR while reading punches from the device:", e)

if records is not None:
    print("Total punches stored on the device:", len(records))
    if records:
        records.sort(key=lambda r: r.timestamp)
        today = datetime.now().date()
        print("Oldest punch:", records[0].timestamp)
        print("Newest punch:", records[-1].timestamp)
        print("Punches dated today:", sum(1 for r in records if r.timestamp.date() == today))
        print("Last 10 punches:")
        for r in records[-10:]:
            print("   user=%s  at=%s" % (r.user_id, r.timestamp))

conn.disconnect()
print("Done.")
