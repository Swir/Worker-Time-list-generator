from worker_time_list.models import WorkEntry
from worker_time_list.pdf_export import PdfOptions, export_pdf, pdf_to_jpg, pdf_to_png


def sample_entry() -> WorkEntry:
    return WorkEntry("2026-09-17", "Client", "07:00", "15:30", 30, "Site work")


def test_pdf_export_with_title_language_and_notes(tmp_path):
    target = tmp_path / "hours.pdf"
    export_pdf(
        target,
        [sample_entry()],
        "Test Worker",
        PdfOptions(include_summary=True, include_notes=True, title="September timesheet", language="no"),
    )
    assert target.exists()
    assert target.stat().st_size > 1000


def test_pdf_can_be_converted_to_png_and_jpg(tmp_path):
    target = tmp_path / "hours.pdf"
    export_pdf(target, [sample_entry()], "Test Worker")
    png_dir = tmp_path / "png"
    jpg_dir = tmp_path / "jpg"
    png_files = pdf_to_png(target, png_dir)
    jpg_files = pdf_to_jpg(target, jpg_dir, quality=90)
    assert len(png_files) == 1 and png_files[0].suffix == ".png" and png_files[0].stat().st_size > 100
    assert len(jpg_files) == 1 and jpg_files[0].suffix == ".jpg" and jpg_files[0].stat().st_size > 100
