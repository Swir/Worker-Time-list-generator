import pytest
from worker_time_list.models import WorkEntry, minutes_to_hhmm, normalize_date, normalize_time

def test_normalizes_common_date_and_time_formats():
    assert normalize_date("17.09.2026")=="2026-09-17"; assert normalize_date("17-09-2026")=="2026-09-17"; assert normalize_time("7:05")=="07:05"; assert normalize_time("0730")=="07:30"
def test_regular_shift_with_break():
    e=WorkEntry("2026-09-17","Client","07:00","15:30",30); assert e.duration_minutes==480; assert e.duration_hhmm=="8:00"
def test_overnight_shift(): assert WorkEntry("2026-09-17","Night","22:00","06:00",30).duration_minutes==450
def test_invalid_break_rejected():
    with pytest.raises(ValueError,match="invalid_break"): WorkEntry("2026-09-17","Client","08:00","10:00",999)
def test_minutes_to_hhmm(): assert minutes_to_hhmm(485)=="8:05"
