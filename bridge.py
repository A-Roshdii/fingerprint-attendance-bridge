#!/usr/bin/env python3
"""
ZKTeco fingerprint device -> attendance system bridge

1. Listens to the ZKTeco device in real time (live capture).
2. Saves every punch in a local SQLite queue first (nothing gets lost).
3. A sender thread pushes queued punches to the attendance system, retrying on failure.

The default output mode is "dry_run": punches are written to a CSV file only.
To push them to your own system, implement send_to_hr_system() (see the TODO inside it).

Usage:
    python bridge.py                # real run (needs the device on the network)
    python bridge.py --simulate     # fake punches, to test without a device
"""
import argparse
import configparser
import csv
import logging
import os
import random
import sqlite3
import threading
import time
from datetime import datetime, timedelta

log = logging.getLogger("bridge")
FMT = "%Y-%m-%d %H:%M:%S"


# ---------------------------------------------------------------- database
def open_db(path):
    c = sqlite3.connect(path, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    c.execute(
        """CREATE TABLE IF NOT EXISTS punches(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_user_id TEXT NOT NULL,
            punched_at TEXT NOT NULL,
            punch_type INTEGER,
            status TEXT NOT NULL DEFAULT 'pending',
            attempts INTEGER NOT NULL DEFAULT 0,
            last_error TEXT,
            created_at TEXT NOT NULL,
            sent_at TEXT,
            UNIQUE(device_user_id, punched_at))"""
    )
    cols = [r["name"] for r in c.execute("PRAGMA table_info(punches)")]
    if "employee_name" not in cols:  # upgrade an older punches.db
        try:
            c.execute("ALTER TABLE punches ADD COLUMN employee_name TEXT")
        except sqlite3.OperationalError:  # the other thread already added it
            pass
    c.commit()
    return c


def save_punch(c, user_id, ts, punch_type, name=""):
    """Insert a punch. Returns True if it is new (duplicates are ignored)."""
    cur = c.execute(
        "INSERT OR IGNORE INTO punches(device_user_id, employee_name, punched_at,"
        " punch_type, created_at) VALUES (?,?,?,?,?)",
        (str(user_id), name or "", ts.strftime(FMT), punch_type,
         datetime.now().strftime(FMT)),
    )
    c.commit()
    return cur.rowcount == 1


# ------------------------------------------------------------ sending to the attendance system
def load_mapping(path):
    """Optional CSV: device_id,employee_code  (if the numbers differ)."""
    mapping = {}
    if path and os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and not row[0].lower().startswith("device"):
                    mapping[row[0].strip()] = row[1].strip()
    return mapping


def send_to_hr_system(rec, cfg, mapping):
    """Send ONE punch to the attendance system. Raise an exception on failure (it will retry)."""
    employee = mapping.get(rec["device_user_id"], rec["device_user_id"])
    payload = {
        "employee_id": employee,
        "employee_name": rec["employee_name"] or "",
        "timestamp": rec["punched_at"],
        "punch_type": rec["punch_type"],  # device value: 0 = in, 1 = out (usually)
    }
    mode = cfg["output"]["mode"]

    if mode == "dry_run":
        path = cfg["output"]["dry_run_file"]
        new_file = not os.path.exists(path)
        # utf-8-sig on a NEW file so Excel shows Arabic names correctly
        with open(path, "a", newline="", encoding="utf-8-sig" if new_file else "utf-8") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(["employee_id", "employee_name", "timestamp", "punch_type"])
            w.writerow([payload["employee_id"], payload["employee_name"],
                        payload["timestamp"], payload["punch_type"]])
        return

    # TODO: fill this in with the API of your attendance system, e.g.:
    #   requests.post(url, json=payload, headers={"Authorization": "..."}, timeout=10)
    #   then check the response status and raise if it is not successful.
    raise NotImplementedError("This output mode is not implemented yet (mode = %s)" % mode)


def sender_loop(cfg, stop):
    c = open_db(cfg["bridge"]["db_path"])
    mapping = load_mapping(cfg["bridge"].get("mapping_file"))
    delay = 1
    while not stop.is_set():
        rows = c.execute(
            "SELECT * FROM punches WHERE status='pending' ORDER BY punched_at LIMIT 50"
        ).fetchall()
        if not rows:
            stop.wait(1)
            continue
        for r in rows:
            try:
                send_to_hr_system(r, cfg, mapping)
                c.execute(
                    "UPDATE punches SET status='sent', sent_at=?, attempts=attempts+1 WHERE id=?",
                    (datetime.now().strftime(FMT), r["id"]),
                )
                c.commit()
                log.info("sent  user=%s name=%s at=%s", r["device_user_id"],
                         r["employee_name"] or "?", r["punched_at"])
                delay = 1
            except Exception as e:
                c.execute(
                    "UPDATE punches SET attempts=attempts+1, last_error=? WHERE id=?",
                    (str(e), r["id"]),
                )
                c.commit()
                delay = min(delay * 2, 60)
                log.warning("send failed (%s) - retry in %ss", e, delay)
                stop.wait(delay)
                break  # keep order: retry the same punch first


# ------------------------------------------------------------- device side
def load_names(conn):
    """Read employee names stored on the device: {user_id: name}."""
    try:
        return {str(u.user_id): (u.name or "").strip() for u in conn.get_users()}
    except Exception as e:
        log.warning("could not read employee names from device: %s", e)
        return {}


def backfill(conn, c, cfg, names):
    """After (re)connecting, queue punches we missed while we were offline."""
    last = c.execute("SELECT MAX(punched_at) AS m FROM punches").fetchone()["m"]
    if last:
        cutoff = datetime.strptime(last, FMT)
    else:
        days = int(cfg["bridge"].get("initial_backfill_days", 0))
        cutoff = datetime.now() - timedelta(days=days)
    log.info("backfill: reading the punches stored on the device (can take a minute)...")
    records = conn.get_attendance()
    added = 0
    for a in records:
        name = names.get(str(a.user_id), "")
        if a.timestamp >= cutoff and save_punch(c, a.user_id, a.timestamp, a.punch, name):
            added += 1
    log.info("backfill done: device has %d punches, queued %d missed ones", len(records), added)


def device_loop(cfg, stop):
    from zk import ZK  # pip install pyzk

    c = open_db(cfg["bridge"]["db_path"])
    ip = cfg["device"]["ip"]
    port = int(cfg["device"].get("port", 4370))
    while not stop.is_set():
        conn = None
        try:
            log.info("connecting to device %s:%s ...", ip, port)
            conn = ZK(ip, port=port, timeout=10).connect()
            log.info("connected")
            names = load_names(conn)
            log.info("loaded %d employee names from device", len(names))
            if cfg.getboolean("bridge", "backfill", fallback=True):
                backfill(conn, c, cfg, names)
            else:
                log.info("backfill skipped (backfill = no in config.ini) - listening now")
            for att in conn.live_capture():
                if stop.is_set():
                    break
                if att is None:  # timeout tick, no punch
                    continue
                name = names.get(str(att.user_id), "")
                if save_punch(c, att.user_id, att.timestamp, att.punch, name):
                    log.info("punch user=%s name=%s at=%s", att.user_id, name or "?", att.timestamp)
        except Exception as e:
            log.warning("device connection problem: %s - retrying in 5s", e)
            stop.wait(5)
        finally:
            if conn:
                try:
                    conn.disconnect()
                except Exception:
                    pass


def simulate_loop(cfg, stop):
    """Fake punches so we can test the whole pipeline without the device."""
    c = open_db(cfg["bridge"]["db_path"])
    while not stop.is_set():
        uid = random.randint(1, 5)
        if save_punch(c, uid, datetime.now(), random.choice([0, 1]), "Test %d" % uid):
            log.info("simulated punch user=%s", uid)
        stop.wait(2)


# --------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.ini")
    ap.add_argument("--simulate", action="store_true", help="fake punches, no device needed")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler("bridge.log", encoding="utf-8")],
    )
    cfg = configparser.ConfigParser()
    cfg.read(args.config, encoding="utf-8")

    stop = threading.Event()
    source = simulate_loop if args.simulate else device_loop
    threads = [
        threading.Thread(target=source, args=(cfg, stop), daemon=True),
        threading.Thread(target=sender_loop, args=(cfg, stop), daemon=True),
    ]
    for t in threads:
        t.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("stopping...")
        stop.set()


if __name__ == "__main__":
    main()
