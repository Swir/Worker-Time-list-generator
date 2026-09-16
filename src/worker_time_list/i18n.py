from __future__ import annotations

import locale

SUPPORTED = ("en", "pl", "no")

STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "title": "Worker Time List Generator", "employee": "Employee / PDF title", "date": "Date", "client": "Client / address",
        "start": "Start", "end": "End", "break": "Break", "note": "Note", "add": "Add entry", "update": "Update",
        "clear": "Clear", "delete": "Delete", "duplicate": "Duplicate", "export_pdf": "Export PDF", "export_csv": "Export CSV",
        "import_csv": "Import CSV", "pdf_images": "PDF → images", "settings": "Timesheet", "language": "Language", "total": "Total",
        "entries": "entries", "ready": "Ready", "saved": "Saved", "error": "Error", "warning": "Warning",
        "client_required": "Client / address is required.", "invalid_date": "Invalid date. Use YYYY-MM-DD, DD-MM-YYYY, DD.MM.YYYY or DD/MM/YYYY.",
        "invalid_time": "Invalid time. Use HH:MM.", "invalid_break": "Break must be between 0 and 360 minutes.",
        "invalid_duration": "Break is longer than the shift.", "select_row": "Select a row first.",
        "confirm_delete": "Delete the selected entry?", "pdf_done": "PDF exported successfully.", "csv_done": "CSV exported successfully.",
        "csv_imported": "CSV imported successfully.", "pdf_images_done": "PDF pages exported as PNG images.",
        "no_entries": "There are no entries to export.", "fix": "Fix fields", "fixed": "Fields normalized where possible.",
        "by": "by Swir",
    },
    "pl": {
        "title": "Generator listy czasu pracy", "employee": "Pracownik / tytuł PDF", "date": "Data", "client": "Klient / adres",
        "start": "Start", "end": "Koniec", "break": "Przerwa", "note": "Notatka", "add": "Dodaj wpis", "update": "Aktualizuj",
        "clear": "Wyczyść", "delete": "Usuń", "duplicate": "Duplikuj", "export_pdf": "Eksport PDF", "export_csv": "Eksport CSV",
        "import_csv": "Import CSV", "pdf_images": "PDF → obrazy", "settings": "Lista czasu", "language": "Język", "total": "Razem",
        "entries": "wpisów", "ready": "Gotowe", "saved": "Zapisano", "error": "Błąd", "warning": "Uwaga",
        "client_required": "Klient / adres jest wymagany.", "invalid_date": "Nieprawidłowa data. Użyj YYYY-MM-DD, DD-MM-YYYY, DD.MM.YYYY lub DD/MM/YYYY.",
        "invalid_time": "Nieprawidłowa godzina. Użyj HH:MM.", "invalid_break": "Przerwa musi mieć od 0 do 360 minut.",
        "invalid_duration": "Przerwa jest dłuższa niż zmiana.", "select_row": "Najpierw wybierz wiersz.",
        "confirm_delete": "Usunąć wybrany wpis?", "pdf_done": "PDF został wyeksportowany.", "csv_done": "CSV został wyeksportowany.",
        "csv_imported": "CSV został zaimportowany.", "pdf_images_done": "Strony PDF zapisano jako obrazy PNG.",
        "no_entries": "Brak wpisów do eksportu.", "fix": "Napraw pola", "fixed": "Pola zostały znormalizowane tam, gdzie było to możliwe.",
        "by": "by Swir",
    },
    "no": {
        "title": "Arbeidstidsliste", "employee": "Ansatt / PDF-tittel", "date": "Dato", "client": "Kunde / adresse",
        "start": "Start", "end": "Slutt", "break": "Pause", "note": "Notat", "add": "Legg til", "update": "Oppdater",
        "clear": "Tøm", "delete": "Slett", "duplicate": "Dupliser", "export_pdf": "Eksporter PDF", "export_csv": "Eksporter CSV",
        "import_csv": "Importer CSV", "pdf_images": "PDF → bilder", "settings": "Timeliste", "language": "Språk", "total": "Totalt",
        "entries": "oppføringer", "ready": "Klar", "saved": "Lagret", "error": "Feil", "warning": "Advarsel",
        "client_required": "Kunde / adresse er påkrevd.", "invalid_date": "Ugyldig dato. Bruk YYYY-MM-DD, DD-MM-YYYY, DD.MM.YYYY eller DD/MM/YYYY.",
        "invalid_time": "Ugyldig tid. Bruk HH:MM.", "invalid_break": "Pause må være mellom 0 og 360 minutter.",
        "invalid_duration": "Pausen er lengre enn skiftet.", "select_row": "Velg en rad først.",
        "confirm_delete": "Slette valgt oppføring?", "pdf_done": "PDF eksportert.", "csv_done": "CSV eksportert.",
        "csv_imported": "CSV importert.", "pdf_images_done": "PDF-sidene ble eksportert som PNG-bilder.",
        "no_entries": "Ingen oppføringer å eksportere.", "fix": "Rett felt", "fixed": "Feltene ble normalisert der det var mulig.",
        "by": "by Swir",
    },
}


def detect_language() -> str:
    try:
        current = locale.getlocale()[0]
    except Exception:
        current = None
    if current:
        code = current.lower().split("_", 1)[0].split("-", 1)[0]
        if code in {"nb", "nn"}:
            code = "no"
        if code in SUPPORTED:
            return code
    return "en"


def tr(language: str, key: str) -> str:
    lang = language if language in STRINGS else "en"
    return STRINGS.get(lang, STRINGS["en"]).get(key, STRINGS["en"].get(key, key))
