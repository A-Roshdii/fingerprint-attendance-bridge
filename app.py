#!/usr/bin/env python3
"""
A small window for the daily routine - no typing of commands.

Run  python app.py  (or make a desktop shortcut to  pythonw app.py ).
Pick the first and last day, press the green button, wait about 2 minutes: the sheet
and the review file appear in this folder and can be opened with one click.

It simply runs  pull_and_export.py <from> <to>  for you, so everything works exactly as before.
"""
import os
import queue
import re
import subprocess
import sys
import threading
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def parse_date(text):
    text = text.strip()
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(text, f)
        except ValueError:
            pass
    return None


def python_exe():
    """python.exe (pythonw.exe has no console output, so the child process must use python.exe)."""
    exe = sys.executable
    if os.path.basename(exe).lower() == "pythonw.exe":
        candidate = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.exists(candidate):
            return candidate
    return exe


def build_command(start, end):
    return [python_exe(), "-u", os.path.join(HERE, "pull_and_export.py"),
            start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")]


def run_job(start, end, on_line):
    """Runs pull_and_export.py; every output line goes to on_line(). Returns (exit code, files it created)."""
    flags = 0x08000000 if os.name == "nt" else 0   # CREATE_NO_WINDOW: no black window
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen(build_command(start, end), cwd=HERE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace", env=env, creationflags=flags)
    files = []
    for line in proc.stdout:
        line = line.rstrip("\r\n")
        on_line(line)
        m = re.search(r"->\s+(attendance_\S+\.xlsx)", line)
        if m:
            files.append(m.group(1))
    return proc.wait(), files


def open_path(path):
    if hasattr(os, "startfile"):
        os.startfile(path)  # Windows: opens with the default program (Excel)
    else:
        subprocess.Popen(["xdg-open", path])


def main():
    import tkinter as tk
    from tkinter import scrolledtext

    root = tk.Tk()
    root.title("Attendance sheet")
    root.geometry("780x560")
    root.minsize(640, 420)
    pad = {"padx": 10, "pady": 6}

    tk.Label(root, text="Pull the punches from the device and build the attendance sheet",
             font=("Segoe UI", 12, "bold")).pack(anchor="w", **pad)

    row = tk.Frame(root)
    row.pack(anchor="w", **pad)
    yesterday = datetime.now() - timedelta(days=1)
    start_var = tk.StringVar(value=yesterday.strftime("%Y-%m-%d"))
    end_var = tk.StringVar(value=yesterday.strftime("%Y-%m-%d"))
    tk.Label(row, text="From day:").grid(row=0, column=0)
    tk.Entry(row, textvariable=start_var, width=12, font=("Segoe UI", 11)).grid(row=0, column=1, padx=6)
    tk.Label(row, text="To day:").grid(row=0, column=2, padx=(14, 0))
    tk.Entry(row, textvariable=end_var, width=12, font=("Segoe UI", 11)).grid(row=0, column=3, padx=6)
    tk.Label(row, text="(like 2026-11-10)", fg="gray").grid(row=0, column=4, padx=6)

    def set_days(first, last):
        start_var.set(first.strftime("%Y-%m-%d"))
        end_var.set(last.strftime("%Y-%m-%d"))

    quick = tk.Frame(root)
    quick.pack(anchor="w", padx=10)
    tk.Button(quick, text="Yesterday", command=lambda: set_days(yesterday, yesterday)).pack(side="left", padx=(0, 6))
    tk.Button(quick, text="Last 7 days",
              command=lambda: set_days(datetime.now() - timedelta(days=7), yesterday)).pack(side="left")

    status = tk.StringVar(value="Ready. The night shift of the last day ends the next morning, so pull after 9:00.")
    run_button = tk.Button(root, text="Pull from the device and make the sheet", bg="#2e7d32", fg="white",
                           font=("Segoe UI", 11, "bold"), height=2)
    run_button.pack(fill="x", **pad)
    tk.Label(root, textvariable=status, anchor="w", fg="#444").pack(fill="x", padx=10)

    bottom = tk.Frame(root)
    bottom.pack(side="bottom", anchor="w", **pad)   # packed first, so the buttons never get cut off

    box = scrolledtext.ScrolledText(root, height=16, font=("Consolas", 10), state="disabled")
    box.pack(fill="both", expand=True, **pad)
    sheet_button = tk.Button(bottom, text="Open the sheet", state="disabled")
    review_button = tk.Button(bottom, text="Open the review file", state="disabled")
    folder_button = tk.Button(bottom, text="Open the folder", command=lambda: open_path(HERE))
    for b in (sheet_button, review_button, folder_button):
        b.pack(side="left", padx=(0, 8))

    messages = queue.Queue()
    found = {"sheet": None, "review": None}

    def log(line):
        box.configure(state="normal")
        box.insert("end", line + "\n")
        box.see("end")
        box.configure(state="disabled")

    def worker(start, end):
        try:
            code, files = run_job(start, end, lambda line: messages.put(("line", line)))
            messages.put(("done", (code, files)))
        except Exception as e:  # e.g. python or the script was not found
            messages.put(("line", "PROBLEM: %s" % e))
            messages.put(("done", (1, [])))

    def poll():
        try:
            while True:
                kind, data = messages.get_nowait()
                if kind == "line":
                    log(data)
                else:
                    code, files = data
                    run_button.configure(state="normal")
                    for name in files:
                        found["review" if name.endswith("_review.xlsx") else "sheet"] = name
                    sheet_button.configure(state="normal" if found["sheet"] else "disabled")
                    review_button.configure(state="normal" if found["review"] else "disabled")
                    status.set("Finished." if code == 0 else "It did not finish - read the lines above (a screenshot helps).")
        except queue.Empty:
            pass
        root.after(150, poll)

    def start_run():
        start, end = parse_date(start_var.get()), parse_date(end_var.get())
        if start is None or end is None:
            status.set("Write the days like 2026-11-10")
            return
        if end < start:
            status.set("The last day is before the first day")
            return
        found["sheet"] = found["review"] = None
        sheet_button.configure(state="disabled")
        review_button.configure(state="disabled")
        run_button.configure(state="disabled")
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.configure(state="disabled")
        status.set("Working... about 2 minutes. Please do not close this window.")
        threading.Thread(target=worker, args=(start, end), daemon=True).start()

    run_button.configure(command=start_run)
    sheet_button.configure(command=lambda: open_path(os.path.join(HERE, found["sheet"])))
    review_button.configure(command=lambda: open_path(os.path.join(HERE, found["review"])))
    root.after(150, poll)
    root.mainloop()


if __name__ == "__main__":
    main()
