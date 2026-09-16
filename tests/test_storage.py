from worker_time_list.models import WorkEntry
from worker_time_list.storage import export_csv, import_csv

def test_csv_roundtrip(tmp_path):
    original=[WorkEntry("2026-09-17","Client A","07:00","15:30",30,"note"),WorkEntry("2026-09-18","Client B","22:00","06:00",0,"")]; target=tmp_path/"hours.csv"; export_csv(target,original); loaded=import_csv(target); assert [i.to_dict() for i in loaded]==[i.to_dict() for i in original]
