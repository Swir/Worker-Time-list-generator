from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Iterable

from .models import WorkEntry, minutes_to_hhmm


@dataclass(frozen=True, slots=True)
class TimesheetSummary:
    entry_count: int
    total_minutes: int
    client_minutes: dict[str, int]

    @property
    def total_hhmm(self) -> str:
        return minutes_to_hhmm(self.total_minutes)


def summarize(entries: Iterable[WorkEntry]) -> TimesheetSummary:
    items = list(entries)
    per_client: dict[str, int] = defaultdict(int)
    total = 0
    for entry in items:
        minutes = entry.duration_minutes
        total += minutes
        per_client[entry.client] += minutes
    return TimesheetSummary(len(items), total, dict(sorted(per_client.items())))
