# Worker Time List Generator

A Python desktop timesheet and work-hours calculator for recording shifts, customer/address details, breaks and total working time. The application can export the completed timesheet to PDF and convert PDF pages to JPG images.

## Features

- Record work date, customer/address, start time and end time
- Optional 30-minute break deduction
- Automatic work-time calculation
- Running total of worked hours and minutes
- Export the work table to PDF
- Custom PDF table header
- Convert PDF pages to JPG
- Tkinter desktop GUI
- Two included versions: `worker v1.py` and `worker v2.py`

## Screens / data fields

The application works with the following fields:

- Date
- Customer / Address
- Start time
- End time
- Break
- Working time
- Total hours

The current V2 interface uses Norwegian labels such as `Dato`, `Kunde/Adresse`, `Starttid`, `Sluttid`, `Tok pause` and `Arbeidstid`.

## Requirements

- Python 3
- Tkinter
- Pillow
- ReportLab
- PyMuPDF

Install the Python dependencies:

```bash
pip install Pillow reportlab PyMuPDF
```

## Run

Clone or download the repository, then start one of the versions:

```bash
python "worker v2.py"
```

or:

```bash
python "worker v1.py"
```

## PDF export

The program can save the entered work table and total working time as a PDF document. V2 can also open a PDF and export each page as a JPG image.

## Use cases

Useful for employees, contractors and small teams that need a simple desktop work-hours log without a web service or account.

## Search keywords

`python timesheet` `work hours calculator` `employee timesheet` `working time tracker` `tkinter timesheet` `work time calculator python` `timesheet pdf export` `work hours pdf` `norwegian timesheet` `arbeidstid` `timeliste` `python desktop app`

## Author

Created by Swir.
