from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from platformdirs import user_data_dir

from .models import WorkEntry

APP_NAME = "WorkerTimeListGenerator"
APP_AUTHOR = "Swir"


@dataclass(slots=True)
class AppState:
    employee: str
    language: str
    entries: list[WorkEntry]


def data_dir() -> Path:
    path = Path(user_data_dir(APP_NAME, APP_AUTHOR))
    path.mkdir(parents=True, exist_ok=True)
    return path


def state_path() -> Path:
    return data_dir() / "timesheet.json"


def load_state(default_language: str = "en") -> AppState:
    path = state_path()
    if not path.exists():
        return AppState("", default_language, [])
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        entries = [WorkEntry.from_dict(item) for item in payload.get("entries", [])]
        return AppState(str(payload.get("employee", "")), str(payload.get("language", default_language)), entries)
    except Exception:
        backup = path.with_suffix(".broken.json")
        try:
            path.replace(backup)
        except OSError:
            pass
        return AppState("", default_language, [])


def save_state(state: AppState) -> None:
    payload = {"schema": 1, "employee": state.employee.strip(), "language": state.language, "entries": [entry.to_dict() for entry in state.entries]}
    path = state_path()
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def export_csv(path: str | Path, entries: Iterable[WorkEntry]) -> None:
    with Path(path).open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=("work_date", "client", "start", "end", "break_minutes", "duration", "note"))
        writer.writeheader()
        for entry in entries:
            writer.writerow({"work_date": entry.work_date, "client": entry.client, "start": entry.start, "end": entry.end, "break_minutes": entry.break_minutes, "duration": entry.duration_hhmm, "note": entry.note})


def import_csv(path: str | Path) -> list[WorkEntry]:
    items: list[WorkEntry] = []
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if not row:
                continue
            items.append(WorkEntry(work_date=row.get("work_date") or row.get("date") or "", client=row.get("client") or row.get("customer") or "", start=row.get("start") or "", end=row.get("end") or "", break_minutes=int(row.get("break_minutes") or row.get("break") or 0), note=row.get("note") or ""))
    return items
