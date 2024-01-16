import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime, timedelta
import fitz  # PyMuPDF
from tkinter.simpledialog import askstring

styles = getSampleStyleSheet()


class Ansatt:
    def __init__(self):
        self.data = []
        self.sum_timer = 0
        self.table_header = ""

    def legg_til_opptegnelse(self, dato, kunde_adresse, start_tid, slutt_tid, tok_pause, arbeidstid):
        oppføring = {
            "Dato": dato,
            "Kunde/Adresse": kunde_adresse,
            "Starttid": start_tid,
            "Sluttid": slutt_tid,
            "Tok pause": tok_pause,
            "Arbeidstid": arbeidstid
        }
        self.data.append(oppføring)
        self.sum_timer += self.tid_til_minutter(arbeidstid)

    def beregn_total_arbeidstid(self):
        total_tid = sum(self.tid_til_minutter(oppføring["Arbeidstid"]) for oppføring in self.data)
        return total_tid

    def tid_til_minutter(self, tid):
        try:
            timer, minutter = map(int, tid.split(":"))
            return timer * 60 + minutter
        except ValueError:
            raise ValueError("Ugyldig tidsformat. Bruk formatet HH:MM.") from None

    def beregn_arbeidstid(self, start, slutt, tok_pause):
        slutt -= timedelta(minutes=30) if tok_pause else timedelta(minutes=0)
        return (slutt - start).total_seconds() // 60

    def zapisz_tabele_do_pdf(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=letter)
        data = self.prepare_table_data()
        table = Table(data)
        self.apply_table_style(table)

        elements = [table]
        if self.table_header:
            elements.insert(0, self.create_header())

        doc.build(elements)

    def prepare_table_data(self):
        table_data = [["Dato", "Kunde/Adresse", "Starttid", "Sluttid", "Tok pause", "Arbeidstid"]]
        for oppføring in self.data:
            table_data.append(list(oppføring.values()))
        sum_row = ["Sum Timer", "", "", "", "", f"{int(self.sum_timer // 60)}:{int(self.sum_timer % 60):02d}"]
        table_data.append(sum_row)
        return table_data

    def apply_table_style(self, table):
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

    def konwertuj_pdf_do_jpg(self, pdf_filename, output_folder, quality=100):
        pdf_document = fitz.open(pdf_filename)
        total_pages = pdf_document.page_count

        for page_number in range(total_pages):
            page = pdf_document[page_number]
            pixmap = page.get_pixmap()
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            image_path = os.path.join(output_folder, f"page_{page_number + 1}.jpg")

            try:
                image.save(image_path, "JPEG", quality=quality)
            except Exception as e:
                print(f"Błąd podczas zapisywania strony {page_number + 1}: {e}")

    def on_close(self):
        if self.data:
            response = messagebox.askyesno('Uwaga', 'Czy na pewno chcesz zamknąć program? Niezapisane dane zostaną utracone.')
            if not response:
                return
        root.destroy()


class AnsattProgram:
    def __init__(self, master):
        self.master = master
        master.title("WojThmas Soft beta 1.0")
        master.geometry("1005x800")

        self.ansatt = Ansatt()
        self.author_info = "Program created by: [Wojciech K. i Thomas O. Polish-Norwegian grammatical errors were left on purpose:)]"

        self.create_interface()

    def create_interface(self):
        øverste_ramme = tk.Frame(self.master)
        øverste_ramme.pack(side=tk.TOP, pady=10, fill=tk.X)

        tk.Label(øverste_ramme, text='Skriv inn informasjon:', font=("Arial", 14, "bold")).grid(row=0, column=0, pady=5, padx=10)

        etiketter = ['Dato (DD-MM-YYYY):', 'Kunde/Adresse:', 'Starttid:', 'Sluttid:']
        inndata = [tk.Entry(øverste_ramme, font=("Arial", 12)) for _ in range(len(etiketter))]

        for i, etikett in enumerate(etiketter):
            tk.Label(øverste_ramme, text=etikett, font=("Arial", 12)).grid(row=i + 1, column=0, pady=5, padx=10, sticky="w")
            inndata[i].grid(row=i + 1, column=1, pady=5, padx=10, sticky="w")

        self.inndata_dato, self.inndata_kunde_adresse, self.inndata_start, self.inndata_slutt = inndata

        self.pause_var = tk.BooleanVar()
        tk.Checkbutton(øverste_ramme, text='Tok du pause?', variable=self.pause_var, font=("Arial", 12)).grid(
            row=len(etiketter) + 1, column=0, pady=5, padx=10, sticky="w")

        self.etikett_navn = tk.Label(øverste_ramme, text='Navn:', font=("Arial", 12))
        self.etikett_navn.grid(row=len(etiketter) + 3, column=0, pady=5, padx=10, sticky="w")
        self.inndata_navn = tk.Entry(øverste_ramme, font=("Arial", 12))
        self.inndata_navn.grid(row=len(etiketter) + 3, column=1, pady=5, padx=10, sticky="w")

        # Dodanie loga
        logo_path = "logo.png"  # Zmienić na nazwę pliku PNG z logo
        if getattr(sys, 'frozen', False):
            # The application is frozen (compiled to .exe)
            logo_full_path = os.path.join(os.path.dirname(sys.argv[0]), logo_path)
        else:
            # The application is not frozen
            logo_full_path = os.path.join(os.path.dirname(__file__), logo_path)

        if os.path.exists(logo_full_path):
            self.logo_image = Image.open(logo_full_path)
            self.logo_image.thumbnail((250, 250))
            self.logo_image = ImageTk.PhotoImage(self.logo_image)
            logo_label = tk.Label(øverste_ramme, image=self.logo_image)
            logo_label.grid(row=2, column=6, rowspan=6, padx=0)
        else:
            print(f"File not found: {logo_full_path}")

        knapper_ramme = tk.Frame(øverste_ramme)
        knapper_ramme.grid(row=len(etiketter) + 4, column=0, columnspan=2, pady=10)

        tk.Button(knapper_ramme, text='Legg til i tabellen', command=self.legg_til_i_tabellen,
                  font=("Arial", 12)).grid(row=0, column=0)
        tk.Button(knapper_ramme, text='Vis totalt antall timer', command=self.vis_totalt_antall_timer,
                  font=("Arial", 12)).grid(row=0, column=1)
        tk.Button(knapper_ramme, text='Save to PDF', command=self.zapisz_tabele_do_pdf,
                  font=("Arial", 12)).grid(row=0, column=2)
        tk.Button(knapper_ramme, text='Konvertuj PDF do JPG', command=self.konwertuj_pdf_do_jpg,
                  font=("Arial", 12)).grid(row=0, column=3)
        tk.Button(knapper_ramme, text='Set Table Header', command=self.set_table_header,
                  font=("Arial", 12)).grid(row=0, column=4)
        tk.Button(knapper_ramme, text='Info', command=self.show_info, font=("Arial", 12)).grid(row=0, column=5)  # Dodany przycisk "Info"

        self.opprett_ark()

    def opprett_ark(self):
        ark_ramme = tk.Frame(self.master)
        ark_ramme.pack(side=tk.LEFT, pady=10, fill=tk.BOTH, expand=True)

        stil = ttk.Style()
        stil.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        stil.configure("Treeview", font=("Arial", 10))

        self.tre = ttk.Treeview(ark_ramme, columns=(
        "Dato", "Kunde/Adresse", "Starttid", "Sluttid", "Tok pause", "Arbeidstid"), show="headings", selectmode="browse")

        kolonneoverskrifter = ["Dato", "Kunde/Adresse", "Starttid", "Sluttid", "Tok pause", "Arbeidstid"]
        for overskrift in kolonneoverskrifter:
            self.tre.heading(overskrift, text=overskrift)

        kolonne_bredder = [100, 100, 150, 150, 100, 150]
        for i, bredde in enumerate(kolonne_bredder):
            self.tre.column(kolonneoverskrifter[i], width=bredde)

        self.tre.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(ark_ramme, orient="vertical", command=self.tre.yview)
        scrollbar.pack(side="right", fill="y")
        self.tre.configure(yscrollcommand=scrollbar.set)

        self.master.bind("<Configure>", lambda event: self.resize_columns(event))
        self.master.bind("<Control-s>", lambda event: self.zapisz_tabele_do_pdf())

        self.ansatt.row_height = int(self.tre.winfo_reqheight() / 30)

        self.etikett_totalt_timer = tk.Label(ark_ramme, text='Totalt antall timer: 0 timer 0 minutter',
                                            font=("Arial", 14, "bold"), fg="blue")
        self.etikett_totalt_timer.pack(side=tk.RIGHT, pady=10, padx=10)

    def resize_columns(self, event=None):
        tree_width = self.tre.winfo_width()
        total_width = sum([self.tre.column(c)["width"] for c in self.tre["columns"]])
        factor = tree_width / total_width if total_width > 0 else 1

        for c in self.tre["columns"]:
            self.tre.column(c, width=int(self.tre.column(c)["width"] * factor))

    def zapisz_tabele_do_pdf(self, event=None):
        try:
            filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if filename:
                self.ansatt.zapisz_tabele_do_pdf(filename)
                messagebox.showinfo('Suksess', 'Tabell og sum Tabell lagret som PDF.')
        except Exception as e:
            messagebox.showerror('Feil', str(e))

    def konwertuj_pdf_do_jpg(self):
        try:
            pdf_filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if pdf_filename:
                output_folder = filedialog.askdirectory()
                self.ansatt.konwertuj_pdf_do_jpg(pdf_filename, output_folder)
                messagebox.showinfo('Suksess', 'PDF konvertert til JPG.')
        except Exception as e:
            messagebox.showerror('Feil', str(e))

    def legg_til_i_tabellen(self):
        dato, kunde_adresse, start_tid, slutt_tid = [inndata.get() for inndata in
                                                     [self.inndata_dato, self.inndata_kunde_adresse, self.inndata_start,
                                                      self.inndata_slutt]]

        if not self.sjekk_datoformat(dato):
            messagebox.showerror('Feil', 'Ugyldig datoformat. Bruk formatet DD-MM-YYYY.')
            return

        try:
            start = self.parse_tid(start_tid)
            slutt = self.parse_tid(slutt_tid)

            arbeidstid = self.ansatt.beregn_arbeidstid(start, slutt, self.pause_var.get())
            arbeidstid_str = f"{int(arbeidstid // 60)}:{int(arbeidstid % 60):02d}"

            self.ansatt.legg_til_opptegnelse(dato, kunde_adresse, start_tid, slutt_tid,
                                             'Ja' if self.pause_var.get() else 'Nei', arbeidstid_str)

            verdier = (dato, kunde_adresse, start_tid, slutt_tid,
                       'Ja' if self.pause_var.get() else 'Nei', arbeidstid_str)

            self.tre.insert("", "end", values=verdier)

            totalt_timer = self.ansatt.beregn_total_arbeidstid()
            timer, minutter = divmod(totalt_timer, 60)
            self.etikett_totalt_timer.config(
                text=f"Totalt antall timer: {timer} timer {minutter} minutter",
                font=("Arial", 12, "bold"))

            messagebox.showinfo('Suksess', 'Oppføring lagt til i tabellen.')

        except ValueError as e:
            messagebox.showerror('Feil', str(e))

    def sjekk_datoformat(self, dato):
        try:
            datetime.strptime(dato, "%d-%m-%Y")
            return True
        except ValueError:
            return False

    def parse_tid(self, tid):
        try:
            return datetime.strptime(tid, "%H:%M")
        except ValueError:
            raise ValueError('Ugyldig tidsformat. Bruk formatet HH:MM eller H:MM.') from None

    def vis_totalt_antall_timer(self):
        if not self.ansatt.data:
            messagebox.showwarning('Oppmerksomhet', 'Ingen data å vise totalt antall timer.')
            return

        totalt_timer = self.ansatt.beregn_total_arbeidstid()
        timer, minutter = divmod(totalt_timer, 60)
        messagebox.showinfo('Totalt antall timer',
                            f"Totalt antall timer: {timer} timer {minutter} minutter")

    def set_table_header(self):
        header = self.inndata_navn.get()
        self.ansatt.table_header = header
        messagebox.showinfo('Suksess', f'Table header set to: {header}')

    def show_info(self):
        messagebox.showinfo('Info', self.author_info)

    def on_close(self):
        self.ansatt.on_close()


if __name__ == "__main__":
    root = tk.Tk()
    app = AnsattProgram(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()

