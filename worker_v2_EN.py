import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.simpledialog import askstring
from datetime import datetime, timedelta

from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import fitz  # PyMuPDF

styles = getSampleStyleSheet()


class EmployeeTimesheet:
    def __init__(self):
        self.data = []
        self.total_minutes = 0
        self.table_header = ""

    @staticmethod
    def time_to_minutes(value):
        hours, minutes = map(int, value.split(":"))
        return hours * 60 + minutes

    @staticmethod
    def calculate_work_time(start, end, took_break):
        if end < start:
            end += timedelta(days=1)
        minutes = int((end - start).total_seconds() // 60)
        if took_break:
            minutes -= 30
        return max(minutes, 0)

    def add_entry(self, date, customer_address, start_time, end_time, took_break, work_time):
        entry = {
            "Date": date,
            "Customer / Address": customer_address,
            "Start": start_time,
            "End": end_time,
            "Break": took_break,
            "Work time": work_time,
        }
        self.data.append(entry)
        self.total_minutes += self.time_to_minutes(work_time)

    def total_work_time(self):
        return sum(self.time_to_minutes(row["Work time"]) for row in self.data)

    def table_data(self):
        rows = [["Date", "Customer / Address", "Start", "End", "Break", "Work time"]]
        rows.extend([list(row.values()) for row in self.data])
        hours, minutes = divmod(self.total_minutes, 60)
        rows.append(["Total", "", "", "", "", f"{hours}:{minutes:02d}"])
        return rows

    def save_pdf(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=letter)
        table = Table(self.table_data())
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
            ("BACKGROUND", (0, 1), (-1, -2), colors.white),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("BACKGROUND", (-1, -1), (-1, -1), colors.green),
            ("TEXTCOLOR", (-1, -1), (-1, -1), colors.white),
        ]))

        elements = [table]
        if self.table_header:
            elements.insert(0, Paragraph(self.table_header, styles["Title"]))
        doc.build(elements)

    @staticmethod
    def pdf_to_jpg(pdf_filename, output_folder, quality=100):
        document = fitz.open(pdf_filename)
        for page_number in range(document.page_count):
            page = document[page_number]
            pixmap = page.get_pixmap()
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            image.save(
                os.path.join(output_folder, f"page_{page_number + 1}.jpg"),
                "JPEG",
                quality=quality,
            )
        document.close()


class TimesheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Worker Time List Generator - English")
        self.root.geometry("1050x780")
        self.timesheet = EmployeeTimesheet()

        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Control-s>", self.save_pdf)

    def build_ui(self):
        top = tk.Frame(self.root)
        top.pack(side=tk.TOP, pady=10, fill=tk.X)

        tk.Label(top, text="Enter work information", font=("Arial", 14, "bold")).grid(
            row=0, column=0, pady=5, padx=10, sticky="w"
        )

        labels = ["Date (DD-MM-YYYY):", "Customer / Address:", "Start time (HH:MM):", "End time (HH:MM):"]
        self.entries = []
        for row, label in enumerate(labels, start=1):
            tk.Label(top, text=label, font=("Arial", 11)).grid(row=row, column=0, pady=5, padx=10, sticky="w")
            entry = tk.Entry(top, width=28, font=("Arial", 11))
            entry.grid(row=row, column=1, pady=5, padx=10, sticky="w")
            self.entries.append(entry)

        self.date_entry, self.customer_entry, self.start_entry, self.end_entry = self.entries

        self.break_var = tk.BooleanVar()
        tk.Checkbutton(top, text="Subtract 30-minute break", variable=self.break_var, font=("Arial", 11)).grid(
            row=5, column=0, columnspan=2, pady=5, padx=10, sticky="w"
        )

        tk.Label(top, text="Employee / table title:", font=("Arial", 11)).grid(row=6, column=0, pady=5, padx=10, sticky="w")
        self.name_entry = tk.Entry(top, width=28, font=("Arial", 11))
        self.name_entry.grid(row=6, column=1, pady=5, padx=10, sticky="w")

        logo_path = os.path.join(os.path.dirname(sys.argv[0] if getattr(sys, "frozen", False) else __file__), "logo.png")
        if os.path.exists(logo_path):
            try:
                image = Image.open(logo_path)
                image.thumbnail((220, 220))
                self.logo_image = ImageTk.PhotoImage(image)
                tk.Label(top, image=self.logo_image).grid(row=1, column=5, rowspan=6, padx=20)
            except Exception:
                pass

        buttons = tk.Frame(top)
        buttons.grid(row=7, column=0, columnspan=6, pady=12)
        tk.Button(buttons, text="Add entry", command=self.add_entry).pack(side=tk.LEFT, padx=4)
        tk.Button(buttons, text="Show total", command=self.show_total).pack(side=tk.LEFT, padx=4)
        tk.Button(buttons, text="Save PDF", command=self.save_pdf).pack(side=tk.LEFT, padx=4)
        tk.Button(buttons, text="PDF to JPG", command=self.convert_pdf).pack(side=tk.LEFT, padx=4)
        tk.Button(buttons, text="Set table title", command=self.set_table_header).pack(side=tk.LEFT, padx=4)
        tk.Button(buttons, text="Clear", command=self.clear_table).pack(side=tk.LEFT, padx=4)

        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        columns = ("Date", "Customer / Address", "Start", "End", "Break", "Work time")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, anchor="center", width=145)
        self.tree.column("Customer / Address", width=250)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.total_label = tk.Label(
            self.root,
            text="Total work time: 0 h 00 min",
            font=("Arial", 13, "bold"),
            fg="blue",
        )
        self.total_label.pack(pady=(0, 12))

    @staticmethod
    def parse_time(value):
        for fmt in ("%H:%M",):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                pass
        raise ValueError("Invalid time format. Use HH:MM, for example 07:30 or 16:00.")

    @staticmethod
    def valid_date(value):
        try:
            datetime.strptime(value, "%d-%m-%Y")
            return True
        except ValueError:
            return False

    def add_entry(self):
        date = self.date_entry.get().strip()
        customer = self.customer_entry.get().strip()
        start_text = self.start_entry.get().strip()
        end_text = self.end_entry.get().strip()

        if not self.valid_date(date):
            messagebox.showerror("Invalid date", "Use DD-MM-YYYY format.")
            return

        try:
            start = self.parse_time(start_text)
            end = self.parse_time(end_text)
            minutes = self.timesheet.calculate_work_time(start, end, self.break_var.get())
            hours, mins = divmod(minutes, 60)
            work_time = f"{hours}:{mins:02d}"
            break_text = "Yes" if self.break_var.get() else "No"

            self.timesheet.add_entry(date, customer, start_text, end_text, break_text, work_time)
            self.tree.insert("", "end", values=(date, customer, start_text, end_text, break_text, work_time))
            self.update_total()
        except ValueError as exc:
            messagebox.showerror("Invalid input", str(exc))

    def update_total(self):
        total = self.timesheet.total_work_time()
        hours, minutes = divmod(total, 60)
        self.total_label.config(text=f"Total work time: {hours} h {minutes:02d} min")

    def show_total(self):
        total = self.timesheet.total_work_time()
        hours, minutes = divmod(total, 60)
        messagebox.showinfo("Total work time", f"{hours} h {minutes:02d} min")

    def save_pdf(self, event=None):
        if not self.timesheet.data:
            messagebox.showwarning("No data", "Add at least one entry before exporting.")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filename:
            try:
                self.timesheet.save_pdf(filename)
                messagebox.showinfo("Saved", "Timesheet saved as PDF.")
            except Exception as exc:
                messagebox.showerror("Export error", str(exc))

    def convert_pdf(self):
        pdf_filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_filename:
            return
        output_folder = filedialog.askdirectory()
        if not output_folder:
            return
        try:
            self.timesheet.pdf_to_jpg(pdf_filename, output_folder)
            messagebox.showinfo("Finished", "PDF pages converted to JPG.")
        except Exception as exc:
            messagebox.showerror("Conversion error", str(exc))

    def set_table_header(self):
        header = self.name_entry.get().strip()
        if not header:
            header = askstring("Table title", "Enter a title for the PDF table:") or ""
        self.timesheet.table_header = header
        if header:
            messagebox.showinfo("Table title", f"Table title set to: {header}")

    def clear_table(self):
        if self.timesheet.data and not messagebox.askyesno("Clear table", "Remove all current entries?"):
            return
        self.timesheet = EmployeeTimesheet()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_total()

    def on_close(self):
        if self.timesheet.data:
            if not messagebox.askyesno("Exit", "Close the program? Unsaved entries will be lost."):
                return
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TimesheetApp(root)
    root.mainloop()
