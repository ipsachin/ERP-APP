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
        self._records_by_snapshot_id = {}
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
        self._records_by_snapshot_id.clear()
        self.text.delete("1.0", "end")

        data = self.app.services.completed_jobs.list_completed_jobs()

        for r in data:
            self.listbox.insert(tk.END, f"{r.project_code} | {r.product_name}")
            self._map.append(r.snapshot_id)
            self._records_by_snapshot_id[r.snapshot_id] = r

        if data:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self._on_select()

    def _on_select(self, e=None):
        if not self.listbox.curselection():
            return

        sid = self._map[self.listbox.curselection()[0]]
        record = self._records_by_snapshot_id.get(sid)
        lines = self.app.services.completed_jobs.get_completed_job_lines(sid)

        assembly_lines = [l for l in lines if l.line_type == "ASSEMBLY"]
        part_lines = [l for l in lines if l.line_type != "ASSEMBLY"]
        out = [
            "COMPLETED JOB SUMMARY",
            "---------------------",
        ]

        if record is not None:
            out += [
                f"Project Code: {record.project_code}",
                f"Client: {record.client_name}",
                f"Quote Ref: {record.quote_ref}",
                f"Product Code: {record.product_code}",
                f"Product Name: {record.product_name or 'N/A'}",
                f"Completed On: {record.completed_on}",
                "",
                f"Assemblies: {len(assembly_lines)}",
                f"Parts: {len(part_lines)}",
                f"Labour Hours: {safe_float(record.labour_hours):g}",
                f"Parts Total: ${safe_float(record.parts_total):,.2f}",
                f"Grand Total: ${safe_float(record.grand_total):,.2f}",
            ]
            if norm_text(getattr(record, "notes", "")):
                out += ["", "Notes:", norm_text(record.notes)]

        if assembly_lines:
            out += ["", "ASSEMBLY BREAKDOWN", "------------------"]
            for l in assembly_lines:
                hours = safe_float(getattr(l, "hours", 0))
                qty = safe_float(getattr(l, "qty", 0))
                label = norm_text(l.description) or norm_text(l.code) or "Assembly"
                out.append(f"{label} | Qty {qty:g} | Hours {hours:g}")

        if part_lines:
            out += ["", "PARTS BREAKDOWN", "---------------"]
            for l in part_lines:
                qty = safe_float(getattr(l, "qty", 0))
                total = safe_float(getattr(l, "line_total", 0))
                desc = norm_text(l.description) or "Part"
                part_no = norm_text(l.part_number)
                row = f"{desc} | Qty {qty:g}"
                if part_no:
                    row += f" | {part_no}"
                row += f" | ${total:,.2f}"
                out.append(row)

        if not lines:
            out += ["", "No snapshot lines were found for this completed job."]

        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(out))
