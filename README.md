<div align="center">

# ⏱️ Worker Time List Generator 3

### Multilingual desktop timesheet for Windows and Python

**Work hours · PL / NO / EN · PDF · CSV · Autosave · Overnight shifts**

</div>

Worker Time List Generator 3 is a complete modernization of the original Tkinter timesheet. The old V1/V2 and separate English scripts are replaced by one maintainable application with automatic language detection, safer local persistence and a modern dark-blue UI.

## Highlights

- one multilingual application: **Polish, Norwegian and English**
- system language is detected at first launch; unsupported languages fall back to English
- date/time validation with one-click field repair
- client/address cleanup without the old leading-dot problem
- configurable break deduction: 0 / 30 / 45 / 60 minutes
- overnight shifts, for example 22:00 → 06:00
- edit, delete and duplicate existing rows
- automatic local autosave in the user's application-data folder
- CSV import/export
- professional landscape PDF export with totals and per-client summary
- PDF → PNG conversion utility
- keyboard shortcuts: `Ctrl+N` add/update, `Ctrl+S` PDF, `Ctrl+Shift+S` CSV, `Delete` remove
- dedicated application icon
- tested on Python 3.10–3.14 in CI
- automated Windows EXE + portable ZIP + SHA256 release pipeline

## Install from source

```bash
git clone https://github.com/Swir/Worker-Time-list-generator.git
cd Worker-Time-list-generator
python -m venv .venv
.venv\Scripts\activate
pip install -e .
python main.py
```

Linux/macOS activation:

```bash
source .venv/bin/activate
```

## Data and privacy

The program works offline. Timesheet autosave data is stored in the current user's application-data directory, not inside the Git repository or beside the executable.

## PDF workflow

A shift can be entered using several common date formats. Internally the application normalizes dates to `YYYY-MM-DD` and times to `HH:MM`. PDF export uses a landscape A4 layout, calculates the overall total and includes per-client totals.

## CSV format

Exports contain `work_date, client, start, end, break_minutes, duration, note`. The same format can be imported back into the application.

## Release build

GitHub Actions tests the project, generates the Windows `.ico`, builds `WorkerTimeList.exe`, smoke-tests the executable and publishes the EXE, portable ZIP and SHA256 files.

## Project layout

```text
src/worker_time_list/
├── app.py
├── i18n.py
├── models.py
├── pdf_export.py
├── storage.py
└── summary.py
tests/
assets/
tools/
```

## Author

Developed by **Swir** — https://github.com/Swir
