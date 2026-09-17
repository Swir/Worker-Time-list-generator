import json

from worker_time_list.models import WorkEntry
from worker_time_list import storage
from worker_time_list.storage import AppState, export_csv, import_csv


def test_csv_roundtrip(tmp_path):
    original = [
        WorkEntry("2026-09-17", "Client A", "07:00", "15:30", 30, "note"),
        WorkEntry("2026-09-18", "Client B", "22:00", "06:00", 0, ""),
    ]
    target = tmp_path / "hours.csv"
    export_csv(target, original)
    loaded = import_csv(target)
    assert [item.to_dict() for item in loaded] == [item.to_dict() for item in original]


def test_state_roundtrip_preserves_pdf_title(tmp_path, monkeypatch):
    target = tmp_path / "timesheet.json"
    monkeypatch.setattr(storage, "state_path", lambda: target)
    state = AppState("Thomas", "September 2026", "no", [WorkEntry("2026-09-17", "Client", "07:00", "15:30", 30, "")])
    storage.save_state(state)
    loaded = storage.load_state("en")
    assert loaded.employee == "Thomas"
    assert loaded.pdf_title == "September 2026"
    assert loaded.language == "no"
    assert len(loaded.entries) == 1


def test_v3_state_without_pdf_title_migrates_cleanly(tmp_path, monkeypatch):
    target = tmp_path / "timesheet.json"
    monkeypatch.setattr(storage, "state_path", lambda: target)
    target.write_text(json.dumps({"schema": 1, "employee": "Worker", "language": "pl", "entries": []}), encoding="utf-8")
    loaded = storage.load_state("en")
    assert loaded.employee == "Worker"
    assert loaded.pdf_title == ""
    assert loaded.language == "pl"
