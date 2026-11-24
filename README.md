# Fingerprint Attendance Bridge

Turn the raw punches stored on a **ZKTeco fingerprint device** into a ready-to-import **attendance sheet (Excel)**, with a **review file** that flags everything unusual: late arrivals, early leaves, half days, forgotten punches and overtime.

No cloud, no subscription: it runs on one PC inside your network, and the data never leaves it.

> **Status:** working project, tested on one real device and one real workflow. Not yet tested on other device models or firmware. See [Limitations](#limitations).

## What problem does it solve?

Every day someone exports the punches from the device, decides who is IN and who is OUT, handles night shifts and forgotten punches, builds the sheet by hand and uploads it to the attendance system. That is slow and easy to get wrong. This tool does it in one click (about two minutes, most of it spent downloading the punches from the device).

```
fingerprint device  ->  download + decode  ->  punches.db  ->  shift rules  ->  attendance_<day>.xlsx
                                                                           \->  attendance_<day>_review.xlsx
```

## Features

- Reads **all punches stored on the device** with its own decoder, checks the data is sane and retries if the download is bad.
- Groups punches into **shifts** (day, middle, night across midnight, 12-hour overtime): **first punch = IN, last punch = OUT**.
- Builds the sheet in **your system's import layout** (columns, labels and date format are set in `config.ini`).
- A **review file** reports, never guesses: forgotten punch, late IN (even one minute), early leave, half day (about 4.5 hours), unusual or very long shifts.
- **Roles**: managers may arrive until 09:00; "rotating" employees can work nights and 12-hour shifts.
- Small **window app** (`app.py`) so non-technical staff can pull a day with one click.
- **Demo mode** with invented employees: try everything without a device.
- Employees not listed in `mapping.csv` are ignored (and listed in the review file, so nothing disappears silently).

## Quick start without a device (demo)

```
pip install -r requirements.txt
python demo/make_demo_data.py
```

It invents 11 employees and one day of punches, runs the exporter and writes two Excel files in `demo_data/`.

## Use with a real device

1. Install Python 3.10+ and the libraries: `pip install -r requirements.txt`
2. Copy `config.example.ini` to `config.ini` and set the device `ip` (device menu: Comm. > Ethernet). Give the device a fixed IP in your router.
3. Copy `mapping.example.csv` to `mapping.csv` and list your employees:

   | column | meaning |
   |---|---|
   | `device_id` | the user number on the fingerprint device |
   | `employee_code` | the employee code in your attendance system |
   | `name` | shown in the review file |
   | `role` | empty (regular), `manager` or `rotating` |

4. Check the connection: `ping <device ip>`, then `python check_device.py`.
5. Pull a day and build the sheet (stop `bridge.py` first, the device accepts one connection at a time):

   ```
   python pull_and_export.py 2026-03-12            # one day
   python pull_and_export.py 2026-03-10 2026-03-15 # several days, one sheet
   ```

   or just run `python app.py` for the window.

Pull a day **after its last shift has ended** (a night shift ends the next morning).

## The output

| file | what it is |
|---|---|
| `attendance_<day>.xlsx` | the sheet to import: one IN row and one OUT row per shift |
| `attendance_<day>_review.xlsx` | sheet **Review**: what needs your attention. Sheet **Ignored**: devices not in `mapping.csv` |

Review types: `Forgotten punch`, `Late IN`, `Early leave`, `Half day`, `Unusual shift`, `Long shift`.

## Rules applied

- Punches **less than 13 hours apart** belong to one shift (so a 12-hour overtime shift stays in one piece; normal rests are longer).
- **First punch = IN, last punch = OUT.** Extra punches in between are ignored.
- A shift belongs to the **day it starts** (`shift_belongs_to = end` changes this).
- A shift with a **single punch** (or two within an hour) is **not exported and not guessed**: it is reported as `Forgotten punch`.
- Shifts: day 08:00-17:00, middle 14:00-23:00, night 23:00-08:00, 12-hour day 08:00-20:00 and 12-hour night 20:00-08:00. Regular employees only work day and middle; `rotating` employees work all of them.
- **Late IN**: any minute after the scheduled start (seconds are ignored). **Managers** may arrive until 09:00, and a normal day is up to 9 hours after the IN.
- **Half day**: 3.5 to 5.5 hours between IN and OUT. **Early leave**: OUT before the scheduled end. More than an hour after the end, or a shift that fits no normal pattern, is `Unusual shift`. More than 14 hours is `Long shift`.

The shift times are constants at the top of `export_sheet.py`; the other limits are in `config.ini`.

## Configuration

Everything is in `config.ini` (see `config.example.ini`, every option is commented). The `[sheet]` section controls the layout of the Excel file, for example for a system that imports 8 columns with the date as `YYYYMMDDHHMM`:

```ini
[sheet]
columns = code, type_name, type, datetime, blank, blank, blank, blank
headers = Employee ID, Type Name, Type, Event Date, Gate, Remark 1, Remark 2, Remark 3
type_name_format = {kind} > {kind}
date_format = %Y%m%d%H%M
header_row = 2
note = My company import template
```

## Project layout

| file | purpose |
|---|---|
| `pull_and_export.py` | the daily routine: download the punches, decode, build the sheet |
| `export_sheet.py` | the shift rules and the Excel writer |
| `app.py` | small window around `pull_and_export.py` |
| `bridge.py` | optional live listener: saves each punch the moment it happens |
| `check_device.py`, `check_raw.py` | diagnostics for the device connection and data |
| `show_punches.py` | list the saved punches of one employee |
| `log_summary.py` | summarise connection problems from `bridge.log` |
| `demo/make_demo_data.py` | invented data to try everything |

Libraries: [pyzk](https://github.com/fananimi/pyzk) (talks to the device), `openpyxl` (Excel), `requests` (reserved for an API output), and the Python standard library (`sqlite3`, `csv`, `configparser`, `tkinter`).

## Limitations

- Tested with **one device** (TCP port 4370, 40-byte punch records). Other models or firmware may differ; the decoder refuses data it does not understand instead of guessing.
- The device cannot be asked for one day only, so every pull downloads **all** stored punches (about two minutes for ~95,000).
- Device memory is limited. When it is full the device may stop recording or overwrite old punches: archive it with whoever administers the device. **Do not delete punches yourself.**
- Shift times are fixed constants for now; other schedules need a small edit in `export_sheet.py`.
- The sheet layout is configurable, but uploading it to your attendance system is up to you.

## Troubleshooting

- **`can't reach device (ping ...)`**: the PC is not on the device's network. Run `ping <ip>`.
- **Download looks empty or broken**: `pull_and_export.py` uses a 60-second timeout and retries three times. Run `python check_raw.py` and look at the numbers.
- **Excel says it cannot write the file**: close the sheet in Excel and run again.

## Privacy

`punches.db`, `config.ini`, `mapping.csv` and the generated sheets contain employee data and your network address. They are listed in `.gitignore`: never publish them.

## Disclaimer

Not affiliated with ZKTeco. ZKTeco is a trademark of its owner; the name is used only to describe compatibility.

## Credits

Built with the help of an AI assistant (Claude).

## License

All rights reserved for now (portfolio project). A license will be added.
