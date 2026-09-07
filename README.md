<div align="center">

# ⏱️ Worker Time List Generator

### Desktop Timesheet & Work-Hours Calculator with PDF Export

**Python • Tkinter • Work Hours • Breaks • PDF • JPG • EN / NO**

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-2ea44f)
![PDF](https://img.shields.io/badge/Export-PDF-red)
![PyMuPDF](https://img.shields.io/badge/PDF%20to-JPG-ff9800)
![Languages](https://img.shields.io/badge/Languages-EN%20%7C%20NO-ff4fa3)

</div>

---

## 🚀 About

**Worker Time List Generator** is a Python desktop timesheet application for recording work dates, customer or address information, start/end times, breaks and total worked time.

The program automatically calculates shift duration, can subtract a 30-minute break, keeps a running total and exports the completed work table to PDF. It also includes a PDF-to-JPG conversion utility.

The repository now contains a dedicated **English version** in addition to the original Norwegian-oriented versions.

It is designed for users searching for a **Python timesheet app**, **work hours calculator**, **employee time tracker**, **Tkinter timesheet**, **work log PDF generator**, **timeliste app** or a lightweight offline work-hours tracker.

---

## ✨ Features

| Feature | Description |
|---|---|
| 📅 Work-date entry | Record each shift date |
| 📍 Customer / address | Save job or workplace information |
| 🕒 Start & end times | Enter shift start and finish |
| ☕ Break deduction | Optional 30-minute break subtraction |
| 🧮 Automatic calculation | Calculates worked hours and minutes |
| ➕ Running total | Shows total accumulated work time |
| 📄 PDF export | Export the complete table to PDF |
| 🏷️ Custom PDF title | Add employee name or custom table header |
| 🖼️ PDF to JPG | Convert PDF pages into JPG images |
| 💾 Ctrl+S | Quick PDF-save shortcut in the desktop app |
| 🌍 English version | Dedicated `worker_v2_EN.py` build |

---

## 🌍 Versions

| Version | Language / Purpose |
|---|---|
| `worker_v2_EN.py` | 🇬🇧 Full English version |
| `worker v2.py` | 🇳🇴 Norwegian-oriented V2 |
| `worker v1.py` | Legacy / decimal-time variant |

---

## 📋 Requirements

- Python 3.x
- Tkinter
- Pillow
- ReportLab
- PyMuPDF

Install dependencies:

```bash
pip install Pillow reportlab PyMuPDF
```

---

## 📦 Installation

```bash
git clone https://github.com/Swir/Worker-Time-list-generator.git
cd Worker-Time-list-generator
pip install Pillow reportlab PyMuPDF
```

Run the English version:

```bash
python worker_v2_EN.py
```

Run the original V2:

```bash
python "worker v2.py"
```

---

## 🧠 Work-Time Example

```text
Start: 07:00
End:   15:30
Break: 30 minutes
-----------------
Work:  8:00
```

The English version also handles shifts that cross midnight.

---

## 📄 PDF Workflow

```text
Work entries
     │
     ▼
Timesheet table
     │
     ▼
PDF export
     │
     └────► optional PDF → JPG conversion
```

---

## 🔍 Discoverability

`python timesheet` • `work hours calculator` • `employee timesheet app` • `work time tracker python` • `tkinter timesheet` • `timesheet pdf export` • `work log generator` • `working hours calculator` • `timeliste python` • `arbeidstid app` • `offline timesheet`

---

## 👨‍💻 Author

Developed by **Swir** — [@Swir](https://github.com/Swir)

<div align="center">

### ⏱️ Track the shift. Calculate the hours. Export the PDF.

⭐ **Star the repository if it helps with your work logs!**

</div>
