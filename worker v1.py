import os
from tkinter import Tk, Label, Button, filedialog, ttk, messagebox, Entry
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import tkinter as tk
import sys
import fitz  # PyMuPDF
from copy import deepcopy  # Dodano import do obsługi głębokiej kopii

styles = getSampleStyleSheet()


class WorkingHoursTable:
    def __init__(self):
        self.data = []
        self.sum_hours = 0
        self.table_header = ""
        self.history = []  # Dodano pole do przechowywania historii stanów tabeli

    def add_record(self, date, client_address, working_hours):
        # Zapisz aktualny stan tabeli przed dodaniem nowego wpisu
        self.history.append(deepcopy(self.data))

        entry = {
            "Dato": date,
            "Kunde/Adresse": client_address,
            "Arbeidstid": working_hours
        }
        self.data.append(entry)
        self.sum_hours += float(working_hours)

    def calculate_total_working_hours(self):
        total_hours = sum(float(entry["Arbeidstid"]) for entry in self.data)
        return total_hours

    def prepare_table_data(self):
        table_data = [["", self.table_header, ""], ["Dato", "Kunde/Adresse", "Arbeidstid"]]
        for entry in self.data:
            table_data.append(list(entry.values()))
        sum_row = ["", "Sum Timer", f"{self.sum_hours:.2f}"]
        table_data.append(sum_row)
        return table_data

    def add_table_style(self, table):
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -2), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)])

        green_style = TableStyle([
            ('BACKGROUND', (-1, -1), (-1, -1), colors.green),
            ('TEXTCOLOR', (-1, -1), (-1, -1), colors.white)])

        table.setStyle(style)
        table.setStyle(green_style)

    def create_header(self):
        return Paragraph(self.table_header, styles['Title'])

    def save_table_to_pdf(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=letter)
        data = self.prepare_table_data()
        table = Table(data)
        self.add_table_style(table)

        elements = [table]
        if self.table_header:
            elements.insert(0, self.create_header())

        doc.build(elements)

    def convert_pdf_to_jpg(self, pdf_filename, output_folder, quality=95):
        try:
            pdf_document = fitz.open(pdf_filename)
            number_of_pages = pdf_document.page_count

            for page_number in range(number_of_pages):
                page = pdf_document[page_number]
                pixmap = page.get_pixmap()
                image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                image_path = os.path.join(output_folder, f"page_{page_number + 1}.jpg")

                try:
                    image.save(image_path, "JPEG", quality=quality)
                except Exception as e:
                    print(f"Error saving page {page_number + 1}: {e}")

            messagebox.showinfo('Success', 'PDF converted to JPG.')
        except Exception as e:
            messagebox.showerror('Error', f'Error converting PDF to JPG: {e}')

    def on_exit(self):
        if self.data:
            response = messagebox.askyesno('Warning', 'Are you sure you want to exit? Unsaved data will be lost.')
            if not response:
                return
        self.master.destroy()


class EmployeeProgram:
    def __init__(self, master):
        self.master = master
        master.title("WojThmas Soft beta 1.1 Gamon Patrol version")
        master.geometry("1005x800")

        self.table_choice = tk.IntVar()
        self.table_choice.set(1)  # Domyślnie wybierz tabelę "Working Hours v1"

        self.working_hours_table = WorkingHoursTable()

        self.author_information = "Program developed by: [Wojciech K. and Thomas O. Polish-Norwegian grammar mistakes are intentional :) Program version 1.1 for patrol gamers]"

        self.create_interface()

    def create_interface(self):
        top_frame = tk.Frame(self.master)
        top_frame.pack(side=tk.TOP, pady=10, fill=tk.X)

        tk.Label(top_frame, text='Enter information:', font=("Arial", 14, "bold")).grid(row=0, column=0, pady=5, padx=10)

        labels = ['Date (DD-MM-YYYY):', 'Client/Address:', 'Working Hours (hours):']
        entries = [tk.Entry(top_frame, font=("Arial", 12)) for _ in range(len(labels))]

        for i, label in enumerate(labels):
            tk.Label(top_frame, text=label, font=("Arial", 12)).grid(row=i + 1, column=0, pady=5, padx=10, sticky="w")
            entries[i].grid(row=i + 1, column=1, pady=5, padx=10, sticky="w")

        self.date_entry, self.client_address_entry, self.working_hours_entry = entries

        self.header_entry = tk.Entry(top_frame, font=("Arial", 12))
        tk.Label(top_frame, text='Table Header:', font=("Arial", 12)).grid(row=len(labels) + 1, column=0, pady=5, padx=10,
                                                                         sticky="w")
        self.header_entry.grid(row=len(labels) + 1, column=1, pady=5, padx=10, sticky="w")

        # Dodanie logo
        logo_path = "logo.png"  # Zmienić na nazwę pliku PNG z logo
        if getattr(sys, 'frozen', False):
            # Aplikacja jest zamrożona (skompilowana do .exe)
            logo_full_path = os.path.join(os.path.dirname(sys.argv[0]), logo_path)
        else:
            # Aplikacja nie jest zamrożona
            logo_full_path = os.path.join(os.path.dirname(__file__), logo_path)

        if os.path.exists(logo_full_path):
            self.logo_image = Image.open(logo_full_path)
            self.logo_image.thumbnail((250, 250))
            self.logo_image = ImageTk.PhotoImage(self.logo_image)
            logo_label = tk.Label(top_frame, image=self.logo_image)
            logo_label.grid(row=2, column=6, rowspan=6, padx=(200, 10), sticky="e")
        else:
            print(f"File not found: {logo_full_path}")

        buttons_frame = tk.Frame(top_frame)
        buttons_frame.grid(row=len(labels) + 4, column=0, columnspan=2, pady=10)

        tk.Radiobutton(top_frame, text="Working Hours v1", variable=self.table_choice, value=1,
                       font=("Arial", 12)).grid(row=len(labels) + 2, column=0, pady=5, padx=10, sticky="w")

        buttons_frame = tk.Frame(top_frame)
        buttons_frame.grid(row=len(labels) + 3, column=0, columnspan=2, pady=10)

        tk.Button(buttons_frame, text='Add to Table', command=self.add_to_table,
                  font=("Arial", 12)).grid(row=0, column=0)
        tk.Button(buttons_frame, text='Save to PDF', command=self.save_table_to_pdf,
                  font=("Arial", 12)).grid(row=0, column=2)
        tk.Button(buttons_frame, text='Convert PDF to JPG', command=self.convert_pdf_to_jpg,
                  font=("Arial", 12)).grid(row=0, column=3)
        tk.Button(buttons_frame, text='Info', command=self.show_info, font=("Arial", 12)).grid(row=0, column=4)
        tk.Button(buttons_frame, text='Undo', command=self.undo,
                  font=("Arial", 12)).grid(row=0, column=5)

        self.create_table()

    def create_table(self):
        table_frame = tk.Frame(self.master)
        table_frame.pack(side=tk.LEFT, pady=10, fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        style.configure("Treeview", font=("Arial", 10))

        self.tree = ttk.Treeview(table_frame, columns=(
        "Date", "Client/Address", "Working Hours"), show="headings", selectmode="browse")

        column_headers = ["Date", "Client/Address", "Working Hours"]
        for header in column_headers:
            self.tree.heading(header, text=header)

        column_widths = [100, 100, 150]
        for i, width in enumerate(column_widths):
            self.tree.column(column_headers[i], width=width)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.master.bind("<Configure>", lambda event: self.update_columns(event))
        self.master.bind("<Control-s>", lambda event: self.save_table_to_pdf())

        self.working_hours_table.row_height = int(self.tree.winfo_reqheight() / 30)

        self.label_total_hours = tk.Label(table_frame, text='Total working hours: 0 hours 0 minutes',
                                          font=("Arial", 14, "bold"), fg="blue")
        self.label_total_hours.pack(side=tk.RIGHT, pady=10, padx=10)

    def update_columns(self, event=None):
        tree_width = self.tree.winfo_width()
        total_width = sum([self.tree.column(c)["width"] for c in self.tree["columns"]])
        factor = tree_width / total_width if total_width > 0 else 1

        for c in self.tree["columns"]:
            self.tree.column(c, width=int(self.tree.column(c)["width"] * factor))

    def save_table_to_pdf(self, event=None):
        try:
            file_name = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if file_name:
                self.working_hours_table.table_header = self.header_entry.get()
                self.working_hours_table.save_table_to_pdf(file_name)
                messagebox.showinfo('Success', 'Table and sum saved as PDF.')
        except Exception as e:
            messagebox.showerror('Error', f'Error saving PDF: {e}')

    def convert_pdf_to_jpg(self):
        try:
            pdf_file_name = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if not pdf_file_name:
                return

            output_folder = filedialog.askdirectory()
            if not output_folder:
                return

            self.working_hours_table.convert_pdf_to_jpg(pdf_file_name, output_folder)
            messagebox.showinfo('Success', 'PDF converted to JPG.')
        except Exception as e:
            messagebox.showerror('Error', f'Error converting PDF to JPG: {e}')

    def add_to_table(self):
        date, client_address, working_hours = [entry.get() for entry in
                                               (self.date_entry, self.client_address_entry, self.working_hours_entry)]

        try:
            self.working_hours_table.add_record(date, client_address, working_hours)
            self.update_table(self.working_hours_table)
            self.show_total_working_hours()
        except ValueError as e:
            messagebox.showerror('Error', str(e))

    def show_total_working_hours(self):
        total_hours = self.working_hours_table.calculate_total_working_hours()
        self.label_total_hours.config(
            text=f'Total working hours: {total_hours:.2f} hours')

    def show_info(self):
        messagebox.showinfo('Information', self.author_information)

    def update_table(self, table_object):
        self.tree.delete(*self.tree.get_children())
        data = table_object.prepare_table_data()
        for entry in data:
            self.tree.insert("", "end", values=entry)

    def undo(self):
        if len(self.working_hours_table.history) > 1:
            # Pobierz poprzednią wersję tabeli z historii
            previous_state = self.working_hours_table.history.pop()
            self.working_hours_table.data = previous_state
            self.update_table(self.working_hours_table)
            self.show_total_working_hours()

    def on_exit(self):
        if self.working_hours_table.data:
            response = messagebox.askyesno('Warning', 'Are you sure you want to exit? Unsaved data will be lost.')
            if not response:
                return
        self.master.destroy()


if __name__ == "__main__":
    root = Tk()
    program = EmployeeProgram(root)
    root.protocol("WM_DELETE_WINDOW", program.on_exit)
    root.mainloop()

