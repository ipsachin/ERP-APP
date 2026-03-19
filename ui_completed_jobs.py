# ui_completed_jobs.py

import tkinter as tk
from tkinter import ttk
from ui_common import BasePage

def norm_text(v): return str(v or "").strip()
def safe_float(v):
    try: return float(v or 0)
    except: return 0.0

class CompletedJobsPage(BasePage):
    PAGE_NAME = "completed_jobs"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self._map = []
        self.search_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        top = ttk.Frame(frame)
        top.pack(fill="x")

        ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left")
        ttk.Label(top, text="Completed Jobs", style="Title.TLabel").pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right")

        body = ttk.Panedwindow(frame, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=1)
        body.add(right, weight=3)

        ttk.Entry(left, textvariable=self.search_var).pack(fill="x")

        self.listbox = tk.Listbox(left)
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        self.text = tk.Text(right)
        self.text.pack(fill="both", expand=True)

    def refresh_page(self):
        if not self.require_workbook():
            return
        self.refresh()

    def refresh(self):
        self.listbox.delete(0, tk.END)
        self._map.clear()
        self.text.delete("1.0", "end")

        data = self.app.services.completed_jobs.list_completed_jobs()

        for r in data:
            self.listbox.insert(tk.END, f"{r.project_code} | {r.product_name}")
            self._map.append(r.snapshot_id)

        if data:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self._on_select()

    def _on_select(self, e=None):
        if not self.listbox.curselection():
            return

        sid = self._map[self.listbox.curselection()[0]]
        lines = self.app.services.completed_jobs.get_completed_job_lines(sid)

        out = ["COMPLETED JOB\n"]

        for l in lines:
            if l.line_type == "ASSEMBLY":
                out.append(f"{l.code} | Qty {l.qty} | Hours {l.hours}")
            else:
                out.append(f"{l.description} | {l.part_number} | Qty {l.qty} | ${l.line_total}")

        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(out))
