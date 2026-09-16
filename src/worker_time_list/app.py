from __future__ import annotations

import webbrowser
from dataclasses import dataclass
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, YES

from .i18n import SUPPORTED, detect_language, tr
from .models import WorkEntry, normalize_date, normalize_time, today_iso
from .pdf_export import PdfOptions, export_pdf, pdf_to_png
from .storage import export_csv, import_csv, load_state, save_state
from .summary import summarize


@dataclass(slots=True)
class FormValues:
    work_date: str; client: str; start: str; end: str; break_minutes: int; note: str


class WorkerTimeApp:
    def __init__(self) -> None:
        self.initial_language = detect_language(); self.state = load_state(self.initial_language)
        if self.state.language not in SUPPORTED: self.state.language = self.initial_language
        self.language = self.state.language
        self.root = ttk.Window(themename="darkly"); self.root.title(f"{tr(self.language,'title')} 3.0"); self.root.geometry("1180x760"); self.root.minsize(900,620)
        self.selected_index: int | None = None
        self.employee_var=ttk.StringVar(value=self.state.employee); self.date_var=ttk.StringVar(value=today_iso()); self.client_var=ttk.StringVar(); self.start_var=ttk.StringVar(value="07:00"); self.end_var=ttk.StringVar(value="15:30"); self.break_var=ttk.StringVar(value="30"); self.note_var=ttk.StringVar(); self.language_var=ttk.StringVar(value=self.language); self.status_var=ttk.StringVar(value=tr(self.language,"ready")); self.total_var=ttk.StringVar()
        self._build(); self._refresh_tree(); self._update_total(); self.root.protocol("WM_DELETE_WINDOW",self._close)
        self.root.bind("<Control-n>",lambda _e:self.add_or_update()); self.root.bind("<Control-s>",lambda _e:self.export_pdf_dialog()); self.root.bind("<Control-Shift-S>",lambda _e:self.export_csv_dialog()); self.root.bind("<Delete>",lambda _e:self.delete_selected())

    def _build(self) -> None:
        top=ttk.Frame(self.root,padding=12); top.pack(fill=X); ttk.Label(top,text=tr(self.language,"title"),font=("Segoe UI",20,"bold"),bootstyle="info").pack(side=LEFT)
        ttk.Combobox(top,textvariable=self.language_var,values=list(SUPPORTED),width=5,state="readonly").pack(side=RIGHT,padx=(8,0)); ttk.Label(top,text=tr(self.language,"language")).pack(side=RIGHT); self.language_var.trace_add("write",self._on_language_change)
        content=ttk.Panedwindow(self.root,orient="horizontal"); content.pack(fill=BOTH,expand=YES,padx=12,pady=(0,10)); form=ttk.Labelframe(content,text=tr(self.language,"settings"),padding=14); table_frame=ttk.Frame(content,padding=(10,0,0,0)); content.add(form,weight=0); content.add(table_frame,weight=1)
        fields=[("employee",self.employee_var),("date",self.date_var),("client",self.client_var),("start",self.start_var),("end",self.end_var),("note",self.note_var)]; self.entries={}
        for row,(key,var) in enumerate(fields):
            ttk.Label(form,text=tr(self.language,key)).grid(row=row*2,column=0,sticky="w",pady=(2,3)); entry=ttk.Entry(form,textvariable=var,width=32); entry.grid(row=row*2+1,column=0,sticky="ew",pady=(0,8)); self.entries[key]=entry
        break_row=len(fields)*2; ttk.Label(form,text=tr(self.language,"break")).grid(row=break_row,column=0,sticky="w",pady=(2,3)); ttk.Combobox(form,textvariable=self.break_var,values=["0","30","45","60"],state="readonly",width=30).grid(row=break_row+1,column=0,sticky="ew",pady=(0,10))
        action_row=break_row+2; self.primary_button=ttk.Button(form,text=tr(self.language,"add"),bootstyle="success",command=self.add_or_update); self.primary_button.grid(row=action_row,column=0,sticky="ew",pady=3)
        ttk.Button(form,text=tr(self.language,"fix"),bootstyle="info-outline",command=self.fix_fields).grid(row=action_row+1,column=0,sticky="ew",pady=3); ttk.Button(form,text=tr(self.language,"clear"),bootstyle="secondary-outline",command=self.clear_form).grid(row=action_row+2,column=0,sticky="ew",pady=3); ttk.Separator(form).grid(row=action_row+3,column=0,sticky="ew",pady=9)
        ttk.Button(form,text=tr(self.language,"export_pdf"),command=self.export_pdf_dialog).grid(row=action_row+4,column=0,sticky="ew",pady=3); ttk.Button(form,text=tr(self.language,"export_csv"),command=self.export_csv_dialog).grid(row=action_row+5,column=0,sticky="ew",pady=3); ttk.Button(form,text=tr(self.language,"import_csv"),command=self.import_csv_dialog).grid(row=action_row+6,column=0,sticky="ew",pady=3); ttk.Button(form,text=tr(self.language,"pdf_images"),command=self.pdf_images_dialog).grid(row=action_row+7,column=0,sticky="ew",pady=3)
        columns=("date","client","start","end","break","work","note"); self.tree=ttk.Treeview(table_frame,columns=columns,show="headings",selectmode="browse",bootstyle="info")
        for col,key in zip(columns,("date","client","start","end","break","total","note")): self.tree.heading(col,text=tr(self.language,key))
        self.tree.column("date",width=100,anchor="center"); self.tree.column("client",width=240); self.tree.column("start",width=75,anchor="center"); self.tree.column("end",width=75,anchor="center"); self.tree.column("break",width=75,anchor="center"); self.tree.column("work",width=80,anchor="center"); self.tree.column("note",width=190); self.tree.pack(fill=BOTH,expand=YES); self.tree.bind("<<TreeviewSelect>>",self._select_row); self.tree.bind("<Double-1>",self._select_row)
        row_actions=ttk.Frame(table_frame); row_actions.pack(fill=X,pady=(8,0)); ttk.Button(row_actions,text=tr(self.language,"delete"),bootstyle="danger-outline",command=self.delete_selected).pack(side=LEFT); ttk.Button(row_actions,text=tr(self.language,"duplicate"),bootstyle="secondary-outline",command=self.duplicate_selected).pack(side=LEFT,padx=6); ttk.Label(row_actions,textvariable=self.total_var,font=("Segoe UI",13,"bold"),bootstyle="info").pack(side=RIGHT)
        footer=ttk.Frame(self.root,padding=(12,4,12,10)); footer.pack(fill=X); ttk.Label(footer,textvariable=self.status_var).pack(side=LEFT); author=ttk.Label(footer,text=tr(self.language,"by"),cursor="hand2",bootstyle="secondary"); author.pack(side=RIGHT); author.bind("<Button-1>",lambda _e:webbrowser.open("https://github.com/Swir"))

    def _entry_from_form(self) -> WorkEntry:
        return WorkEntry(self.date_var.get(),self.client_var.get(),self.start_var.get(),self.end_var.get(),int(self.break_var.get() or 0),self.note_var.get())
    def _mark_validation(self,bad_key=None) -> None:
        for key,widget in self.entries.items(): widget.configure(bootstyle="danger" if key==bad_key else "default")
    def add_or_update(self) -> None:
        try: entry=self._entry_from_form()
        except ValueError as exc:
            key=str(exc); bad="date" if key=="invalid_date" else "start" if key=="invalid_time" else "client" if key=="client_required" else None; self._mark_validation(bad); messagebox.showerror(tr(self.language,"error"),tr(self.language,key),parent=self.root); return
        self._mark_validation(); self.state.entries.append(entry) if self.selected_index is None else self.state.entries.__setitem__(self.selected_index,entry); self.state.employee=self.employee_var.get().strip(); self._persist(); self.clear_form(True); self._refresh_tree(); self._update_total(); self.status_var.set(tr(self.language,"saved"))
    def clear_form(self,keep_employee=True) -> None:
        self.selected_index=None; self.date_var.set(today_iso()); self.client_var.set(""); self.start_var.set("07:00"); self.end_var.set("15:30"); self.break_var.set("30"); self.note_var.set("");
        if not keep_employee: self.employee_var.set("")
        self.primary_button.configure(text=tr(self.language,"add")); self._mark_validation()
    def fix_fields(self) -> None:
        try: self.date_var.set(normalize_date(self.date_var.get()))
        except ValueError: self.date_var.set(today_iso())
        for var in (self.start_var,self.end_var):
            try: var.set(normalize_time(var.get()))
            except ValueError: pass
        self.client_var.set(self.client_var.get().strip().lstrip("•·. ").strip()); self.status_var.set(tr(self.language,"fixed"))
    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index,e in enumerate(self.state.entries): self.tree.insert("",END,iid=str(index),values=(e.work_date,e.client,e.start,e.end,f"{e.break_minutes} min",e.duration_hhmm,e.note))
    def _select_row(self,_event=None) -> None:
        selected=self.tree.selection()
        if not selected:return
        self.selected_index=int(selected[0]); e=self.state.entries[self.selected_index]; self.date_var.set(e.work_date); self.client_var.set(e.client); self.start_var.set(e.start); self.end_var.set(e.end); self.break_var.set(str(e.break_minutes)); self.note_var.set(e.note); self.primary_button.configure(text=tr(self.language,"update"))
    def delete_selected(self) -> None:
        if self.selected_index is None: messagebox.showwarning(tr(self.language,"warning"),tr(self.language,"select_row"),parent=self.root); return
        if not messagebox.askyesno(tr(self.language,"warning"),tr(self.language,"confirm_delete"),parent=self.root): return
        del self.state.entries[self.selected_index]; self._persist(); self.clear_form(); self._refresh_tree(); self._update_total()
    def duplicate_selected(self) -> None:
        if self.selected_index is None: messagebox.showwarning(tr(self.language,"warning"),tr(self.language,"select_row"),parent=self.root); return
        e=self.state.entries[self.selected_index]; self.state.entries.append(WorkEntry(e.work_date,e.client,e.start,e.end,e.break_minutes,e.note)); self._persist(); self._refresh_tree(); self._update_total()
    def _update_total(self) -> None:
        s=summarize(self.state.entries); self.total_var.set(f"{tr(self.language,'total')}: {s.total_hhmm} · {s.entry_count} {tr(self.language,'entries')}")
    def _persist(self) -> None:
        self.state.employee=self.employee_var.get().strip(); self.state.language=self.language; save_state(self.state)
    def export_pdf_dialog(self) -> None:
        if not self.state.entries: messagebox.showwarning(tr(self.language,"warning"),tr(self.language,"no_entries"),parent=self.root); return
        path=filedialog.asksaveasfilename(parent=self.root,defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile="work-hours.pdf")
        if not path:return
        try: export_pdf(path,self.state.entries,self.employee_var.get(),PdfOptions(include_summary=True,include_notes=True)); self._persist(); messagebox.showinfo(tr(self.language,"saved"),tr(self.language,"pdf_done"),parent=self.root)
        except Exception as exc: messagebox.showerror(tr(self.language,"error"),str(exc),parent=self.root)
    def export_csv_dialog(self) -> None:
        if not self.state.entries: messagebox.showwarning(tr(self.language,"warning"),tr(self.language,"no_entries"),parent=self.root); return
        path=filedialog.asksaveasfilename(parent=self.root,defaultextension=".csv",filetypes=[("CSV","*.csv")],initialfile="work-hours.csv")
        if path:
            try: export_csv(path,self.state.entries); messagebox.showinfo(tr(self.language,"saved"),tr(self.language,"csv_done"),parent=self.root)
            except Exception as exc: messagebox.showerror(tr(self.language,"error"),str(exc),parent=self.root)
    def import_csv_dialog(self) -> None:
        path=filedialog.askopenfilename(parent=self.root,filetypes=[("CSV","*.csv")])
        if path:
            try: self.state.entries.extend(import_csv(path)); self._persist(); self._refresh_tree(); self._update_total(); messagebox.showinfo(tr(self.language,"saved"),tr(self.language,"csv_imported"),parent=self.root)
            except Exception as exc: messagebox.showerror(tr(self.language,"error"),str(exc),parent=self.root)
    def pdf_images_dialog(self) -> None:
        source=filedialog.askopenfilename(parent=self.root,filetypes=[("PDF","*.pdf")]);
        if not source:return
        destination=filedialog.askdirectory(parent=self.root)
        if not destination:return
        try: pdf_to_png(source,destination); messagebox.showinfo(tr(self.language,"saved"),tr(self.language,"pdf_images_done"),parent=self.root)
        except Exception as exc: messagebox.showerror(tr(self.language,"error"),str(exc),parent=self.root)
    def _on_language_change(self,*_args) -> None:
        new=self.language_var.get()
        if new in SUPPORTED and new!=self.language: self.language=new; self.state.language=new; self._persist(); self.root.after_idle(self._restart)
    def _restart(self) -> None:
        self.root.destroy(); WorkerTimeApp().run()
    def _close(self) -> None:
        self._persist(); self.root.destroy()
    def run(self) -> None: self.root.mainloop()


def run() -> None: WorkerTimeApp().run()
