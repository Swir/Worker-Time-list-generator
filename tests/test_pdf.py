from worker_time_list.models import WorkEntry
from worker_time_list.pdf_export import PdfOptions, export_pdf

def test_pdf_export(tmp_path):
    target=tmp_path/"hours.pdf"; export_pdf(target,[WorkEntry("2026-09-17","Client","07:00","15:30",30,"Site work")],"Test Worker",PdfOptions(include_summary=True,include_notes=True)); assert target.exists(); assert target.stat().st_size>1000
