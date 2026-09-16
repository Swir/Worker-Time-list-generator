from worker_time_list.models import WorkEntry
from worker_time_list.summary import summarize

def test_summary_totals_and_clients():
    entries=[WorkEntry("2026-09-17","A","08:00","12:00",0),WorkEntry("2026-09-18","A","08:00","10:00",0),WorkEntry("2026-09-18","B","10:00","12:30",30)]
    result=summarize(entries); assert result.entry_count==3; assert result.total_minutes==480; assert result.client_minutes=={"A":360,"B":120}
