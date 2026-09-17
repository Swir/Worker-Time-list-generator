from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .models import WorkEntry, minutes_to_hhmm
from .summary import summarize


@dataclass(frozen=True, slots=True)
class PdfOptions:
    columns: tuple[str, ...] = ("work_date", "client", "start", "end", "break_minutes", "duration")
    include_summary: bool = True
    include_notes: bool = False
    title: str = ""
    language: str = "en"


HEADERS = {
    "en": {"work_date": "Date", "client": "Client / address", "start": "Start", "end": "End", "break_minutes": "Break", "duration": "Work", "note": "Note", "total": "TOTAL", "summary": "Summary"},
    "pl": {"work_date": "Data", "client": "Klient / adres", "start": "Start", "end": "Koniec", "break_minutes": "Przerwa", "duration": "Praca", "note": "Notatka", "total": "RAZEM", "summary": "Podsumowanie"},
    "no": {"work_date": "Dato", "client": "Kunde / adresse", "start": "Start", "end": "Slutt", "break_minutes": "Pause", "duration": "Arbeid", "note": "Notat", "total": "TOTALT", "summary": "Sammendrag"},
}


def _headers(language: str) -> dict[str, str]:
    return HEADERS.get(language, HEADERS["en"])


def _cell(entry: WorkEntry, key: str) -> str:
    if key == "duration":
        return entry.duration_hhmm
    if key == "break_minutes":
        return f"{entry.break_minutes} min"
    return str(getattr(entry, key))


def _safe_text(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def export_pdf(path: str | Path, entries: Iterable[WorkEntry], employee: str = "", options: PdfOptions | None = None) -> None:
    opts = options or PdfOptions()
    items = list(entries)
    if not items:
        raise ValueError("no_entries")
    columns = list(opts.columns)
    if opts.include_notes and "note" not in columns:
        columns.append("note")
    if not columns:
        columns = ["work_date", "client", "duration"]

    labels = _headers(opts.language)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleBlue", parent=styles["Title"], textColor=colors.HexColor("#246A93"), alignment=TA_CENTER, fontSize=18, leading=22)
    subtitle_style = ParagraphStyle("Subtitle", parent=styles["BodyText"], textColor=colors.HexColor("#4A6072"), alignment=TA_CENTER, fontSize=10, leading=13)
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=8.5, leading=11)
    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm, title=opts.title.strip() or "Worker Time List", author="Swir")

    heading = opts.title.strip() or employee.strip() or "Worker Time List"
    story: list[object] = [Paragraph(_safe_text(heading), title_style)]
    if opts.title.strip() and employee.strip():
        story.extend([Spacer(1, 1.5 * mm), Paragraph(_safe_text(employee.strip()), subtitle_style)])
    story.append(Spacer(1, 5 * mm))

    data: list[list[object]] = [[labels.get(key, key) for key in columns]]
    for entry in items:
        row: list[object] = []
        for key in columns:
            value = _cell(entry, key)
            row.append(Paragraph(_safe_text(value), small) if key in {"client", "note"} else value)
        data.append(row)

    summary = summarize(items)
    total_row = []
    for index, key in enumerate(columns):
        total_row.append(labels["total"] if index == 0 else summary.total_hhmm if key == "duration" else "")
    data.append(total_row)

    available_width = landscape(A4)[0] - 24 * mm
    weights = {"work_date": 1.0, "client": 2.5, "start": 0.8, "end": 0.8, "break_minutes": 0.8, "duration": 0.9, "note": 2.0}
    total_weight = sum(weights.get(key, 1.0) for key in columns)
    widths = [available_width * weights.get(key, 1.0) / total_weight for key in columns]
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#153A5B")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DFF6FF")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#829AB1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (1, 1), (1, -2), "LEFT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F5FAFD")]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(table)

    if opts.include_summary and summary.client_minutes:
        rows = [[labels["client"], labels["total"]]] + [
            [Paragraph(_safe_text(client), small), minutes_to_hhmm(minutes)]
            for client, minutes in summary.client_minutes.items()
        ]
        summary_table = Table(rows, colWidths=[available_width * 0.75, available_width * 0.25], repeatRows=1)
        summary_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#153A5B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#829AB1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (1, 0), (1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.extend([Spacer(1, 6 * mm), KeepTogether([Paragraph(labels["summary"], styles["Heading2"]), summary_table])])
    doc.build(story)


def pdf_to_images(pdf_path: str | Path, output_dir: str | Path, image_format: str = "png", dpi: int = 160, jpeg_quality: int = 95) -> list[Path]:
    import pymupdf

    fmt = image_format.lower().strip()
    if fmt not in {"png", "jpg", "jpeg"}:
        raise ValueError("unsupported_image_format")
    extension = "jpg" if fmt in {"jpg", "jpeg"} else "png"
    source = pymupdf.open(str(pdf_path))
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []
    matrix = pymupdf.Matrix(dpi / 72, dpi / 72)
    try:
        for index, page in enumerate(source):
            target = destination / f"page_{index + 1:02d}.{extension}"
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            if extension == "png":
                pixmap.save(str(target))
            else:
                from PIL import Image

                image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                image.save(target, "JPEG", quality=max(1, min(100, int(jpeg_quality))), optimize=True)
            outputs.append(target)
    finally:
        source.close()
    return outputs


def pdf_to_png(pdf_path: str | Path, output_dir: str | Path, dpi: int = 160) -> list[Path]:
    return pdf_to_images(pdf_path, output_dir, "png", dpi=dpi)


def pdf_to_jpg(pdf_path: str | Path, output_dir: str | Path, dpi: int = 160, quality: int = 95) -> list[Path]:
    return pdf_to_images(pdf_path, output_dir, "jpg", dpi=dpi, jpeg_quality=quality)
