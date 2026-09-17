from __future__ import annotations

import copy
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, YES

from .i18n import SUPPORTED, detect_language, tr
from .models import WorkEntry, normalize_date, normalize_time, today_iso
from .pdf_export import PdfOptions, export_pdf, pdf_to_jpg, pdf_to_png
from .storage import export_csv, import_csv, load_state, save_state
from .summary import summarize


@dataclass(slots=True)
class FormValues:
    work_date: str
    client: str
    start: str
    end: str
    break_minutes: int
    note: str


def _asset_path(name: str) -> Path:
    frozen_root = getattr(sys, "_MEIPASS", None)
    if frozen_root:
        return Path(frozen_root) / "assets" / name
    return Path(__file__).resolve().parents[2] / "assets" / name


class WorkerTimeApp:
    def __init__(self) -> None:
        self.initial_language = detect_language()
        self.state = load_state(self.initial_language)
        if self.state.language not in SUPPORTED:
            self.state.language = self.initial_language
        self.language = self.state.language
        self.root = ttk.Window(themename="darkly")
        self.root.title(f"{tr(self.language, 'title')} 3.1")
        self.root.geometry("1220x790")
        self.root.minsize(920, 640)
        self._apply_icon()

        self.selected_index: int | None = None
        self._history: list[list[WorkEntry]] = []
        self.employee_var = ttk.StringVar(value=self.state.employee)
        self.pdf_title_var = ttk.StringVar(value=self.state.pdf_title)
        self.date_var = ttk.StringVar(value=today_iso())
        self.client_var = ttk.StringVar()
        self.start_var = ttk.StringVar(value="07:00")
        self.end_var = ttk.StringVar(value="15:30")
        self.break_var = ttk.StringVar(value="30")
        self.note_var = ttk.StringVar()
        self.language_var = ttk.StringVar(value=self.language)
        self.status_var = ttk.StringVar(value=tr(self.language, "ready"))
        self.total_var = ttk.StringVar()

        self._build()
        self._refresh_tree()
        self._update_total()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.bind("<Control-n>", lambda _e: self.add_or_update())
        self.root.bind("<Control-s>", lambda _e: self.export_pdf_dialog())
        self.root.bind("<Control-Shift-S>", lambda _e: self.export_csv_dialog())
        self.root.bind("<Control-z>", lambda _e: self.undo())
        self.root.bind("<Delete>", lambda _e: self.delete_selected())

    def _apply_icon(self) -> None:
        ico = _asset_path("worker-time-list.ico")
        if ico.exists():
            try:
                self.root.iconbitmap(str(ico))
            except Exception:
                pass

    def _build(self) -> None:
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill=X)
        ttk.Label(top, text=tr(self.language, "title"), font=("Segoe UI", 20, "bold"), bootstyle="info").pack(side=LEFT)
        ttk.Combobox(top, textvariable=self.language_var, values=list(SUPPORTED), width=5, state="readonly").pack(side=RIGHT, padx=(8, 0))
        ttk.Label(top, text=tr(self.language, "language")).pack(side=RIGHT)
        self.language_var.trace_add("write", self._on_language_change)

        content = ttk.Panedwindow(self.root, orient="horizontal")
        content.pack(fill=BOTH, expand=YES, padx=12, pady=(0, 10))
        form = ttk.Labelframe(content, text=tr(self.language, "settings"), padding=14)
        table_frame = ttk.Frame(content, padding=(10, 0, 0, 0))
        content.add(form, weight=0)
        content.add(table_frame, weight=1)

        fields = [
            ("employee", self.employee_var),
            ("pdf_title", self.pdf_title_var),
            ("date", self.date_var),
            ("client", self.client_var),
            ("start", self.start_var),
            ("end", self.end_var),
            ("note", self.note_var),
        ]
        self.entries: dict[str, ttk.Entry] = {}
        for row, (key, var) in enumerate(fields):
            ttk.Label(form, text=tr(self.language, key)).grid(row=row * 2, column=0, sticky="w", pady=(2, 3))
            entry = ttk.Entry(form, textvariable=var, width=32)
            entry.grid(row=row * 2 + 1, column=0, sticky="ew", pady=(0, 8))
            self.entries[key] = entry

        break_row = len(fields) * 2
        ttk.Label(form, text=tr(self.language, "break")).grid(row=break_row, column=0, sticky="w", pady=(2, 3))
        ttk.Combobox(form, textvariable=self.break_var, values=["0", "30", "45", "60"], state="readonly", width=30).grid(row=break_row + 1, column=0, sticky="ew", pady=(0, 10))

        action_row = break_row + 2
        self.primary_button = ttk.Button(form, text=tr(self.language, "add"), bootstyle="success", command=self.add_or_update)
        self.primary_button.grid(row=action_row, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "undo"), bootstyle="warning-outline", command=self.undo).grid(row=action_row + 1, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "fix"), bootstyle="info-outline", command=self.fix_fields).grid(row=action_row + 2, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "clear"), bootstyle="secondary-outline", command=self.clear_form).grid(row=action_row + 3, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "show_total"), bootstyle="secondary-outline", command=self.show_total).grid(row=action_row + 4, column=0, sticky="ew", pady=3)
        ttk.Separator(form).grid(row=action_row + 5, column=0, sticky="ew", pady=9)
        ttk.Button(form, text=tr(self.language, "export_pdf"), command=self.export_pdf_dialog).grid(row=action_row + 6, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "export_csv"), command=self.export_csv_dialog).grid(row=action_row + 7, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "import_csv"), command=self.import_csv_dialog).grid(row=action_row + 8, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "pdf_images"), command=self.pdf_images_dialog).grid(row=action_row + 9, column=0, sticky="ew", pady=3)
        ttk.Button(form, text=tr(self.language, "pdf_jpg"), command=self.pdf_jpg_dialog).grid(row=action_row + 10, column=0, sticky="ew", pady=3)

        columns = ("date", "client", "start", "end", "break", "work", "note")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse", bootstyle="info")
        for col, key in zip(columns, ("date", "client", "start", "end", "break", "total", "note")):
            self.tree.heading(col, text=tr(self.language, key))
        self.tree.column("date", width=100, anchor="center")
        self.tree.column("client", width=240)
        self.tree.column("start", width=75, anchor="center")
        self.tree.column("end", width=75, anchor="center")
        self.tree.column("break", width=75, anchor="center")
        self.tree.column("work", width=80, anchor="center")
        self.tree.column("note", width=190)
        self.tree.pack(fill=BOTH, expand=YES)
        self.tree.bind("<<TreeviewSelect>>", self._select_row)
        self.tree.bind("<Double-1>", self._select_row)

        row_actions = ttk.Frame(table_frame)
        row_actions.pack(fill=X, pady=(8, 0))
        ttk.Button(row_actions, text=tr(self.language, "delete"), bootstyle="danger-outline", command=self.delete_selected).pack(side=LEFT)
        ttk.Button(row_actions, text=tr(self.language, "duplicate"), bootstyle="secondary-outline", command=self.duplicate_selected).pack(side=LEFT, padx=6)
        ttk.Button(row_actions, text=tr(self.language, "undo"), bootstyle="warning-outline", command=self.undo).pack(side=LEFT)
        ttk.Label(row_actions, textvariable=self.total_var, font=("Segoe UI", 13, "bold"), bootstyle="info").pack(side=RIGHT)

        footer = ttk.Frame(self.root, padding=(12, 4, 12, 10))
        footer.pack(fill=X)
        ttk.Label(footer, textvariable=self.status_var).pack(side=LEFT)
        author = ttk.Label(footer, text=tr(self.language, "by"), cursor="hand2", bootstyle="secondary")
        author.pack(side=RIGHT)
        author.bind("<Button-1>", lambda _e: webbrowser.open("https://github.com/Swir"))

    def _entry_from_form(self) -> WorkEntry:
        return WorkEntry(
            self.date_var.get(),
            self.client_var.get(),
            self.start_var.get(),
            self.end_var.get(),
            int(self.break_var.get() or 0),
            self.note_var.get(),
        )

    def _mark_validation(self, bad_key: str | None = None) -> None:
        for key, widget in self.entries.items():
            widget.configure(bootstyle="danger" if key == bad_key else "default")

    def _remember_history(self) -> None:
        self._history.append(copy.deepcopy(self.state.entries))
        if len(self._history) > 50:
            del self._history[0]

    def add_or_update(self) -> None:
        try:
            entry = self._entry_from_form()
        except ValueError as exc:
            key = str(exc)
            bad = "date" if key == "invalid_date" else "start" if key == "invalid_time" else "client" if key == "client_required" else None
            self._mark_validation(bad)
            messagebox.showerror(tr(self.language, "error"), tr(self.language, key), parent=self.root)
            return
        self._mark_validation()
        self._remember_history()
        if self.selected_index is None:
            self.state.entries.append(entry)
        else:
            self.state.entries[self.selected_index] = entry
        self._persist()
        self.clear_form(True)
        self._refresh_tree()
        self._update_total()
        self.status_var.set(tr(self.language, "saved"))

    def clear_form(self, keep_employee: bool = True) -> None:
        self.selected_index = None
        self.date_var.set(today_iso())
        self.client_var.set("")
        self.start_var.set("07:00")
        self.end_var.set("15:30")
        self.break_var.set("30")
        self.note_var.set("")
        if not keep_employee:
            self.employee_var.set("")
        self.primary_button.configure(text=tr(self.language, "add"))
        self._mark_validation()

    def fix_fields(self) -> None:
        try:
            self.date_var.set(normalize_date(self.date_var.get()))
        except ValueError:
            self.date_var.set(today_iso())
        for var in (self.start_var, self.end_var):
            try:
                var.set(normalize_time(var.get()))
            except ValueError:
                pass
        self.client_var.set(self.client_var.get().strip().lstrip("•·. ").strip())
        self.status_var.set(tr(self.language, "fixed"))

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, entry in enumerate(self.state.entries):
            self.tree.insert("", END, iid=str(index), values=(entry.work_date, entry.client, entry.start, entry.end, f"{entry.break_minutes} min", entry.duration_hhmm, entry.note))

    def _select_row(self, _event=None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        self.selected_index = int(selected[0])
        entry = self.state.entries[self.selected_index]
        self.date_var.set(entry.work_date)
        self.client_var.set(entry.client)
        self.start_var.set(entry.start)
        self.end_var.set(entry.end)
        self.break_var.set(str(entry.break_minutes))
        self.note_var.set(entry.note)
        self.primary_button.configure(text=tr(self.language, "update"))

    def delete_selected(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning(tr(self.language, "warning"), tr(self.language, "select_row"), parent=self.root)
            return
        if not messagebox.askyesno(tr(self.language, "warning"), tr(self.language, "confirm_delete"), parent=self.root):
            return
        self._remember_history()
        del self.state.entries[self.selected_index]
        self._persist()
        self.clear_form()
        self._refresh_tree()
        self._update_total()

    def duplicate_selected(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning(tr(self.language, "warning"), tr(self.language, "select_row"), parent=self.root)
            return
        self._remember_history()
        entry = self.state.entries[self.selected_index]
        self.state.entries.append(WorkEntry(entry.work_date, entry.client, entry.start, entry.end, entry.break_minutes, entry.note))
        self._persist()
        self._refresh_tree()
        self._update_total()

    def undo(self) -> None:
        if not self._history:
            self.status_var.set(tr(self.language, "nothing_to_undo"))
            return
        self.state.entries = self._history.pop()
        self._persist()
        self.clear_form()
        self._refresh_tree()
        self._update_total()
        self.status_var.set(tr(self.language, "saved"))

    def _update_total(self) -> None:
        summary = summarize(self.state.entries)
        self.total_var.set(f"{tr(self.language, 'total')}: {summary.total_hhmm} · {summary.entry_count} {tr(self.language, 'entries')}")

    def show_total(self) -> None:
        summary = summarize(self.state.entries)
        text = tr(self.language, "total_message").format(total=summary.total_hhmm, count=summary.entry_count)
        messagebox.showinfo(tr(self.language, "total"), text, parent=self.root)

    def _persist(self) -> None:
        self.state.employee = self.employee_var.get().strip()
        self.state.pdf_title = self.pdf_title_var.get().strip()
        self.state.language = self.language
        save_state(self.state)

    def export_pdf_dialog(self) -> None:
        if not self.state.entries:
            messagebox.showwarning(tr(self.language, "warning"), tr(self.language, "no_entries"), parent=self.root)
            return
        path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile="work-hours.pdf")
        if not path:
            return
        try:
            options = PdfOptions(include_summary=True, include_notes=True, title=self.pdf_title_var.get().strip(), language=self.language)
            export_pdf(path, self.state.entries, self.employee_var.get(), options)
            self._persist()
            messagebox.showinfo(tr(self.language, "saved"), tr(self.language, "pdf_done"), parent=self.root)
        except Exception as exc:
            messagebox.showerror(tr(self.language, "error"), str(exc), parent=self.root)

    def export_csv_dialog(self) -> None:
        if not self.state.entries:
            messagebox.showwarning(tr(self.language, "warning"), tr(self.language, "no_entries"), parent=self.root)
            return
        path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="work-hours.csv")
        if path:
            try:
                export_csv(path, self.state.entries)
                messagebox.showinfo(tr(self.language, "saved"), tr(self.language, "csv_done"), parent=self.root)
            except Exception as exc:
                messagebox.showerror(tr(self.language, "error"), str(exc), parent=self.root)

    def import_csv_dialog(self) -> None:
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("CSV", "*.csv")])
        if path:
            try:
                imported = import_csv(path)
                if imported:
                    self._remember_history()
                    self.state.entries.extend(imported)
                    self._persist()
                    self._refresh_tree()
                    self._update_total()
                messagebox.showinfo(tr(self.language, "saved"), tr(self.language, "csv_imported"), parent=self.root)
            except Exception as exc:
                messagebox.showerror(tr(self.language, "error"), str(exc), parent=self.root)

    def _choose_pdf_and_folder(self) -> tuple[str, str] | None:
        source = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF", "*.pdf")])
        if not source:
            return None
        destination = filedialog.askdirectory(parent=self.root)
        if not destination:
            return None
        return source, destination

    def pdf_images_dialog(self) -> None:
        selected = self._choose_pdf_and_folder()
        if not selected:
            return
        try:
            pdf_to_png(*selected)
            messagebox.showinfo(tr(self.language, "saved"), tr(self.language, "pdf_images_done"), parent=self.root)
        except Exception as exc:
            messagebox.showerror(tr(self.language, "error"), str(exc), parent=self.root)

    def pdf_jpg_dialog(self) -> None:
        selected = self._choose_pdf_and_folder()
        if not selected:
            return
        try:
            pdf_to_jpg(*selected)
            messagebox.showinfo(tr(self.language, "saved"), tr(self.language, "pdf_jpg_done"), parent=self.root)
        except Exception as exc:
            messagebox.showerror(tr(self.language, "error"), str(exc), parent=self.root)

    def _on_language_change(self, *_args) -> None:
        new_language = self.language_var.get()
        if new_language in SUPPORTED and new_language != self.language:
            self.language = new_language
            self.state.language = new_language
            self._persist()
            self.root.after_idle(self._restart)

    def _restart(self) -> None:
        self.root.destroy()
        WorkerTimeApp().run()

    def _close(self) -> None:
        self._persist()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def run() -> None:
    WorkerTimeApp().run()
