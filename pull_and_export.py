#!/usr/bin/env python3
"""
Morning routine in ONE command (the bridge does NOT need to be running):

    python pull_and_export.py 2026-11-03               # one day
    python pull_and_export.py 2026-11-03 2026-11-04    # several days

1. Connects to the ZKTeco device and downloads ALL punches stored on it (about 2 minutes,
   the device cannot be asked for one day only).
2. Decodes them with its own decoder and checks the data is sane (retries if the download is bad).
3. Keeps the punches around the requested days and saves them in punches.db.
4. Builds the attendance sheet for the requested days (same as export_sheet.py).

Stop bridge.py first if it is running (the device accepts one connection at a time).
"""
import configparser
import struct
import sys
from datetime import datetime, timedelta

import bridge          # reuses open_db() and save_punch()
import export_sheet    # reuses parse_day() and the sheet builder
from zk import ZK, const

RECORD = 40            # bytes per stored punch on this device
ATTEMPTS = 3


def decode_time(raw):
    """ZKTeco packs date+time in 4 bytes (31-day months, years since 2000)."""
    t = struct.unpack("<I", raw.ljust(4, b"\x00"))[0]
    second = t % 60
    t //= 60
    minute = t % 60
    t //= 60
    hour = t % 24
    t //= 24
    day = t % 31 + 1
    t //= 31
    month = t % 12 + 1
    t //= 12
    return datetime(t + 2000, month, day, hour, minute, second)


def decode_punches(data, expected_records):
    """Returns (valid punches [(user_id, datetime, punch_type)], number of broken records)."""
    if len(data) < 4:
        raise ValueError("the device sent no data")
    total = struct.unpack("<I", data[:4])[0]
    if expected_records and total != expected_records * RECORD:
        raise ValueError("unexpected punch size (%s bytes for %s punches), this decoder expects %d bytes each"
                         % (total, expected_records, RECORD))
    body = data[4:4 + total]
    good, broken = [], 0
    for i in range(0, len(body) - RECORD + 1, RECORD):
        _uid, user_raw, _status, ts_raw, punch, _pad = struct.unpack("<H24sB4sB8s", body[i:i + RECORD])
        user_id = user_raw.split(b"\x00")[0].decode("ascii", errors="ignore").strip()
        try:
            ts = decode_time(ts_raw)
        except ValueError:
            broken += 1
            continue
        if not user_id or not (2020 <= ts.year <= 2100):
            broken += 1
            continue
        good.append((user_id, ts, punch))
    return good, broken


def download(ip, port):
    """Download + decode, retrying when the data looks broken. Returns (punches, names)."""
    last_error = "unknown"
    for attempt in range(1, ATTEMPTS + 1):
        conn = None
        try:
            print("Attempt %d/%d: connecting to %s:%s ..." % (attempt, ATTEMPTS, ip, port))
            conn = ZK(ip, port=port, timeout=60).connect()  # long timeout: the download is big
            conn.read_sizes()
            expected = getattr(conn, "records", 0) or 0
            print("The device says it holds %d punches. Downloading (about 2 minutes) ..." % expected)
            data, _size = conn.read_with_buffer(const.CMD_ATTLOG_RRQ)
            punches, broken = decode_punches(data, expected)
            if not punches or broken > max(50, len(punches) // 100):
                raise ValueError("%d of %d records are broken" % (broken, len(punches) + broken))
            try:
                names = {str(u.user_id): (u.name or "").strip() for u in conn.get_users()}
            except Exception as e:  # names are only a nicety
                print("Could not read the employee names (%s), continuing without them." % e)
                names = {}
            print("Decoded %d punches (%d broken records skipped)." % (len(punches), broken))
            return punches, names
        except Exception as e:
            last_error = str(e)
            print("  Problem: %s" % last_error)
        finally:
            if conn:
                try:
                    conn.disconnect()
                except Exception:
                    pass
    sys.exit("Could not get good data from the device after %d attempts (last problem: %s)." % (ATTEMPTS, last_error))


def main():
    args = sys.argv[1:]
    if not 1 <= len(args) <= 2:
        sys.exit(__doc__)
    start, end = export_sheet.parse_day(args[0]), export_sheet.parse_day(args[-1])
    if end < start:
        sys.exit("The end date is before the start date.")

    cfg = configparser.ConfigParser()
    cfg.read("config.ini", encoding="utf-8")
    ip = cfg["device"]["ip"]
    port = int(cfg["device"].get("port", 4370))
    db_path = cfg.get("bridge", "db_path", fallback="punches.db")

    punches, names = download(ip, port)
    newest = max(p[1] for p in punches)
    print("Newest punch on the device: %s" % newest.strftime("%Y-%m-%d %H:%M:%S"))
    if newest < end + timedelta(days=1):
        print("Note: the device has no punches after that time yet, so the last shift of the "
              "range may still be open (OUT missing).")

    # same window the sheet builder reads: one day before, two days after
    lo, hi = start - timedelta(days=1), end + timedelta(days=2)
    c = bridge.open_db(db_path)
    kept = new = 0
    for user_id, ts, punch in punches:
        if lo <= ts < hi:
            kept += 1
            if bridge.save_punch(c, user_id, ts, punch, names.get(user_id, "")):
                new += 1
    print("Punches around the requested days: %d (%d new in punches.db)" % (kept, new))
    print("-" * 60)

    sys.argv = ["export_sheet.py"] + args
    export_sheet.main()


if __name__ == "__main__":
    main()
