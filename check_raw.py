"""
Diagnostic only (reads from the device, changes nothing).
Shows what the device REALLY sends when we ask for its stored punches.

Stop bridge.py first (Ctrl + C), then run:   python check_raw.py
It takes about 2 minutes. Send me everything it prints.
"""
import configparser
from struct import unpack

from zk import ZK, const

cfg = configparser.ConfigParser()
cfg.read("config.ini", encoding="utf-8")
ip = cfg["device"]["ip"]
port = int(cfg["device"].get("port", 4370))

print("Connecting to %s:%s ..." % (ip, port))
conn = ZK(ip, port=port, timeout=60).connect()
print("Connected.")

try:
    conn.read_sizes()
    print("Punches the device says it holds :", getattr(conn, "records", "?"))
    print("Punch capacity of the device     :", getattr(conn, "rec_cap", "?"))
    print("Users on the device              :", getattr(conn, "users", "?"))
except Exception as e:
    print("Could not read the device sizes:", e)

try:
    print("Downloading the stored punches (about 2 minutes) ...")
    data, size = conn.read_with_buffer(const.CMD_ATTLOG_RRQ)
    print("Bytes received                   :", len(data))
    total = unpack("I", data[:4])[0] if len(data) >= 4 else 0
    records = getattr(conn, "records", 0) or 0
    print("Size announced in the data header:", total)
    if records and total:
        rec_size = total / records
        print("Bytes per punch                  :", rec_size)
        body = data[4:]
        n = int(rec_size)
        if n > 0:
            blank = sum(1 for i in range(0, len(body) - n + 1, n) if not any(body[i:i + n]))
            print("Punches that are all zero bytes  :", blank)
            print("First punch (hex)                :", body[:n].hex())
            print("A punch from the middle (hex)    :", body[(len(body) // n // 2) * n:][:n].hex())
            print("Last punch (hex)                 :", body[-n:].hex())
except Exception as e:
    print("ERROR while downloading the punches:", repr(e))

conn.disconnect()
print("Done.")
