<div align="center">

<img src="assets/worker-time-list.svg" alt="Worker Time List Generator icon" width="128" height="128">

# ⏱️ Worker Time List Generator 3.1

### Multilingual desktop timesheet for Windows and Python

**Work hours · PL / NO / EN · PDF · CSV · Autosave · Overnight shifts**

</div>

Worker Time List Generator 3.1 keeps the modern modular application while restoring useful workflows from the original V1/V2 programs that were lost during the first rewrite. The old separate scripts are still replaced by one maintainable application, but classic table-title, undo, total-summary and PDF→JPG behavior are available again.

## Highlights

- one multilingual application: **Polish, Norwegian and English**
- system language is detected at first launch; unsupported languages fall back to English
- date/time validation with one-click field repair
- client/address cleanup without the old leading-dot problem
- configurable break deduction: 0 / 30 / 45 / 60 minutes
- overnight shifts, for example 22:00 → 06:00
- edit, delete and duplicate existing rows
- **Undo / Ctrl+Z** for recent table mutations, restored from the classic V1 workflow
- automatic local autosave in the user's application-data folder
- separate **Employee** and **PDF / table title** fields, restoring the classic custom-header workflow
- quick **Show total** dialog plus live total in the main window
- CSV import/export
- professional landscape PDF export with totals, notes, per-client summary and localized PL/NO/EN column labels
- **PDF → PNG** and restored **PDF → JPG** conversion
- keyboard shortcuts: `Ctrl+N` add/update, `Ctrl+S` PDF, `Ctrl+Shift+S` CSV, `Ctrl+Z` undo, `Delete` remove
- dedicated Worker Time List application icon shown above, embedded in Windows builds and applied to the desktop window
- tested on Python 3.10–3.14 in CI
- Windows GUI startup smoke test plus packaged EXE GUI smoke test
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

Check the installed version without opening the GUI:

```bash
python main.py --version
```

## Data and privacy

The program works offline. Timesheet autosave data is stored in the current user's application-data directory, not inside the Git repository or beside the executable. Existing v3.0 state files are migrated automatically to the v3.1 state schema.

## PDF workflow

A shift can be entered using several common date formats. Internally the application normalizes dates to `YYYY-MM-DD` and times to `HH:MM`. The optional **PDF / table title** appears as the document heading; the employee name can be shown below it. PDF export uses landscape A4, includes the overall total and can include per-client totals and notes.

The original application could convert generated PDFs to JPG. v3.1 restores that workflow while keeping the newer PNG export too.

## CSV format

Exports contain `work_date, client, start, end, break_minutes, duration, note`. The same format can be imported back into the application.

## Release build

GitHub Actions tests the project, generates the custom Windows `.ico` and PNG artwork, starts the real GUI on Windows, builds `WorkerTimeList.exe`, runs version and packaged-GUI smoke tests, then publishes the EXE, portable ZIP and SHA256 files.

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

## Regression-audit notes

Compared with the historical `worker v1.py` / `worker v2.py` line, v3.1 deliberately restores the useful user-facing behaviors (undo, table header/title, total dialog and JPG conversion) without bringing back duplicated language programs or repository-local runtime data. Autosave makes the old "unsaved data will be lost" exit warning unnecessary.

## Author

Developed by **Swir** — https://github.com/Swir
