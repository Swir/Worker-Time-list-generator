from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from typing import Any

DATE_FORMAT = "%Y-%m-%d"
TIME_FORMAT = "%H:%M"


@dataclass(slots=True)
class WorkEntry:
    work_date: str
    client: str
    start: str
    end: str
    break_minutes: int = 30
    note: str = ""

    def __post_init__(self) -> None:
        self.work_date = normalize_date(self.work_date)
        self.start = normalize_time(self.start)
        self.end = normalize_time(self.end)
        self.client = self.client.strip()
        self.note = self.note.strip()
        if not self.client:
            raise ValueError("client_required")
        if self.break_minutes < 0 or self.break_minutes > 360:
            raise ValueError("invalid_break")
        if self.duration_minutes <= 0:
            raise ValueError("invalid_duration")

    @property
    def duration_minutes(self) -> int:
        start_dt, end_dt = shift_datetimes(self.work_date, self.start, self.end)
        minutes = int((end_dt - start_dt).total_seconds() // 60) - int(self.break_minutes)
        return minutes

    @property
    def duration_hhmm(self) -> str:
        return minutes_to_hhmm(self.duration_minutes)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["duration_minutes"] = self.duration_minutes
        return payload

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "WorkEntry":
        return cls(
            work_date=str(payload.get("work_date", "")),
            client=str(payload.get("client", "")),
            start=str(payload.get("start", "")),
            end=str(payload.get("end", "")),
            break_minutes=int(payload.get("break_minutes", 0)),
            note=str(payload.get("note", "")),
        )


def normalize_date(value: str) -> str:
    value = value.strip()
    formats = ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y")
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).strftime(DATE_FORMAT)
        except ValueError:
            continue
    raise ValueError("invalid_date")


def normalize_time(value: str) -> str:
    value = value.strip().replace(".", ":")
    for fmt in ("%H:%M", "%H%M"):
        try:
            return datetime.strptime(value, fmt).strftime(TIME_FORMAT)
        except ValueError:
            continue
    if ":" in value:
        hour, minute = value.split(":", 1)
        if hour.isdigit() and minute.isdigit():
            try:
                return datetime.strptime(f"{int(hour):02d}:{int(minute):02d}", "%H:%M").strftime(TIME_FORMAT)
            except ValueError:
                pass
    raise ValueError("invalid_time")


def shift_datetimes(work_date: str, start: str, end: str) -> tuple[datetime, datetime]:
    day = datetime.strptime(normalize_date(work_date), DATE_FORMAT)
    start_t = datetime.strptime(normalize_time(start), TIME_FORMAT).time()
    end_t = datetime.strptime(normalize_time(end), TIME_FORMAT).time()
    start_dt = datetime.combine(day.date(), start_t)
    end_dt = datetime.combine(day.date(), end_t)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
    return start_dt, end_dt


def minutes_to_hhmm(minutes: int) -> str:
    if minutes < 0:
        raise ValueError("negative_minutes")
    hours, mins = divmod(int(minutes), 60)
    return f"{hours}:{mins:02d}"


def today_iso() -> str:
    return date.today().strftime(DATE_FORMAT)
