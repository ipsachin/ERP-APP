# ============================================================
# ui_parts.py
# Parts manager page for Liquimech ERP Desktop App
# Back button + editable unit price
# ============================================================

from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk

from app_config import AppConfig
from ui_common import BasePage, treeview_clear, show_warning, show_info, show_error


def norm_text(v):
    return str(v or "").strip()


def safe_float(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


class PartsPage(BasePage):
    PAGE_NAME = "parts"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self.search_var = tk.StringVar()
        self.price_var = tk.StringVar()
        self.parts_rows = []
        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        top = ttk.Frame(wrapper)
        top.pack(fill="x", pady=(0, 8))

        ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left")
        ttk.Label(top, text="Parts Manager", style="Title.TLabel").pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right")

        search_row = ttk.Frame(wrapper)
        search_row.pack(fill="x", pady=(0, 8))
        ttk.Label(search_row, text="Search").pack(side="left")
        ent = ttk.Entry(search_row, textvariable=self.search_var)
        ent.pack(side="left", fill="x", expand=True, padx=8)
        ent.bind("<KeyRelease>", lambda e: self.refresh_parts())
        ttk.Button(search_row, text="Open Selected", command=self.open_selected_part).pack(side="left")

        add_card = ttk.LabelFrame(wrapper, text="Add New Part", padding=10)
        add_card.pack(fill="x", pady=(0, 8))

        self.new_part_name_var = tk.StringVar()
        self.new_part_number_var = tk.StringVar()
        self.new_qty_var = tk.StringVar(value="0")
        self.new_soh_var = tk.StringVar(value="0")
        self.new_supplier_var = tk.StringVar()
        self.new_lead_var = tk.StringVar(value="0")
        self.new_price_var = tk.StringVar(value="0")
        self.new_notes_var = tk.StringVar()

        ttk.Label(add_card, text="Part Name").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_part_name_var).grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Part Number").grid(row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_part_number_var).grid(row=0, column=3, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Qty").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_qty_var).grid(row=1, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="SOH Qty").grid(row=1, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_soh_var).grid(row=1, column=3, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Supplier").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_supplier_var).grid(row=2, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Lead Time").grid(row=2, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_lead_var).grid(row=2, column=3, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Unit Price").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_price_var).grid(row=3, column=1, sticky="ew", padx=4, pady=4)

        ttk.Label(add_card, text="Notes").grid(row=3, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(add_card, textvariable=self.new_notes_var).grid(row=3, column=3, sticky="ew", padx=4, pady=4)

        ttk.Button(add_card, text="Add Part", command=self.add_new_part).grid(
            row=0, column=4, rowspan=4, padx=8, pady=4, sticky="ns"
        )

        add_card.columnconfigure(1, weight=1)
        add_card.columnconfigure(3, weight=1)

        body = ttk.Panedwindow(wrapper, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=3)
        body.add(right, weight=2)

        cols = ("PartName", "PartNumber", "Qty", "SOHQty", "Supplier", "LeadTime", "UnitPrice", "Source")
        self.parts_tree = ttk.Treeview(left, columns=cols, show="headings", height=20)
        widths = {
            "PartName": 240,
            "PartNumber": 150,
            "Qty": 70,
            "SOHQty": 80,
            "Supplier": 160,
            "LeadTime": 80,
            "UnitPrice": 90,
            "Source": 140,
        }
        for c in cols:
            self.parts_tree.heading(c, text=c)
            self.parts_tree.column(c, width=widths[c], anchor="w")
        ysb = ttk.Scrollbar(left, orient="vertical", command=self.parts_tree.yview)
        self.parts_tree.configure(yscrollcommand=ysb.set)
        self.parts_tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.parts_tree.bind("<<TreeviewSelect>>", lambda e: self.show_part_details())
        self.parts_tree.bind("<Double-1>", lambda e: self.show_part_details())

        detail_card = ttk.LabelFrame(right, text="Part Details", padding=10)
        detail_card.pack(fill="both", expand=True)

        form = ttk.Frame(detail_card)
        form.pack(fill="x", pady=(0, 8))

        ttk.Label(form, text="Unit Price").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(form, textvariable=self.price_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(form, text="Save Price", command=self.save_selected_price).grid(row=0, column=2, padx=(8, 0), pady=4)
        form.columnconfigure(1, weight=1)

        note = ttk.Label(
            detail_card,
            text="Price is saved into component Notes as UnitPrice=<value> and used by quote/summary patches.",
            style="Sub.TLabel",
            wraplength=320,
            justify="left",
        )
        note.pack(anchor="w", pady=(0, 8))

        self.detail_text = tk.Text(detail_card, wrap="word")
        dscroll = ttk.Scrollbar(detail_card, orient="vertical", command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=dscroll.set)
        self.detail_text.pack(side="left", fill="both", expand=True)
        dscroll.pack(side="right", fill="y")
        self.detail_text.configure(state="disabled")

    def add_new_part(self):
        name = norm_text(self.new_part_name_var.get())
        part_number = norm_text(self.new_part_number_var.get())

        if not name:
            show_warning("Missing Data", "Part Name is required.")
            return

        repo = getattr(self.app, "repo", None)
        if repo is None:
            show_error("Save Error", "Repository not available.")
            return

        qty = safe_float(self.new_qty_var.get())
        soh_qty = safe_float(self.new_soh_var.get())
        supplier = norm_text(self.new_supplier_var.get())
        lead_time = norm_text(self.new_lead_var.get())
        price = safe_float(self.new_price_var.get())
        notes = norm_text(self.new_notes_var.get())

        if price:
            notes = f"{notes} | UnitPrice={price:.2f}" if notes else f"UnitPrice={price:.2f}"

        existing = repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
        new_id = f"CMP-{len(existing)+1:05d}"

        row = {
            "ComponentID": new_id,
            "OwnerType": "PART",
            "OwnerCode": "PARTS_MASTER",
            "ComponentName": name,
            "Qty": qty,
            "SOHQty": soh_qty,
            "PreferredSupplier": supplier,
            "LeadTimeDays": lead_time,
            "PartNumber": part_number,
            "Notes": notes,
            "CreatedOn": "",
            "UpdatedOn": "",
        }

        repo.append_dict(AppConfig.SHEET_COMPONENTS, row)

        try:
            repo.save_workbook()
        except Exception:
            pass

        self.new_part_name_var.set("")
        self.new_part_number_var.set("")
        self.new_qty_var.set("0")
        self.new_soh_var.set("0")
        self.new_supplier_var.set("")
        self.new_lead_var.set("0")
        self.new_price_var.set("0")
        self.new_notes_var.set("")

        show_info("Part Added", f"New part added: {name}")
        self.refresh_parts()
    def refresh_page(self):
        if not self.require_workbook():
            treeview_clear(self.parts_tree)
            return
        self.refresh_parts()

    def _extract_unit_price(self, notes: str) -> float:
        notes = norm_text(notes)
        m = re.search(r'UnitPrice\s*=\s*([0-9]+(?:\.[0-9]+)?)', notes, re.I)
        if m:
            return safe_float(m.group(1))
        return 0.0

    def _merge_unit_price(self, notes: str, value: float) -> str:
        notes = norm_text(notes)
        tag = f"UnitPrice={value:.2f}"
        if not notes:
            return tag
        if re.search(r'UnitPrice\s*=\s*[0-9]+(?:\.[0-9]+)?', notes, re.I):
            return re.sub(r'UnitPrice\s*=\s*[0-9]+(?:\.[0-9]+)?', tag, notes, flags=re.I)
        return notes + " | " + tag

    def _collect_parts(self):
        rows = []

        try:
            if hasattr(self.app.services, "parts") and hasattr(self.app.services.parts, "list_parts"):
                for p in self.app.services.parts.list_parts():
                    notes = norm_text(getattr(p, "notes", ""))
                    rows.append({
                        "part_name": norm_text(getattr(p, "part_name", "")) or norm_text(getattr(p, "component_name", "")),
                        "part_number": norm_text(getattr(p, "part_number", "")) or norm_text(getattr(p, "part_code", "")),
                        "qty": safe_float(getattr(p, "qty", 0)),
                        "soh_qty": safe_float(getattr(p, "soh_qty", 0)),
                        "supplier": norm_text(getattr(p, "preferred_supplier", "")) or norm_text(getattr(p, "supplier", "")),
                        "lead_time": norm_text(getattr(p, "lead_time_days", "")),
                        "unit_price": self._extract_unit_price(notes),
                        "source": "Parts",
                        "notes": notes,
                    })
                return rows
        except Exception:
            pass

        try:
            modules = self.app.services.modules.list_modules()
            for mod in modules:
                comps = self.app.services.modules.get_module_components(mod.module_code)
                for c in comps:
                    notes = norm_text(getattr(c, "notes", ""))
                    rows.append({
                        "part_name": norm_text(getattr(c, "component_name", "")),
                        "part_number": norm_text(getattr(c, "part_number", "")),
                        "qty": safe_float(getattr(c, "qty", 0)),
                        "soh_qty": safe_float(getattr(c, "soh_qty", 0)),
                        "supplier": norm_text(getattr(c, "preferred_supplier", "")),
                        "lead_time": norm_text(getattr(c, "lead_time_days", "")),
                        "unit_price": self._extract_unit_price(notes),
                        "source": f"Assembly {mod.module_code}",
                        "notes": notes,
                    })
        except Exception:
            pass

        try:
            products = self.app.services.products.list_products()
            for prod in products:
                if hasattr(self.app.services.products, "get_product_bundle"):
                    bundle = self.app.services.products.get_product_bundle(prod.product_code)
                    for c in getattr(bundle, "direct_components", []) or []:
                        notes = norm_text(getattr(c, "notes", ""))
                        rows.append({
                            "part_name": norm_text(getattr(c, "component_name", "")),
                            "part_number": norm_text(getattr(c, "part_number", "")),
                            "qty": safe_float(getattr(c, "qty", 0)),
                            "soh_qty": safe_float(getattr(c, "soh_qty", 0)),
                            "supplier": norm_text(getattr(c, "preferred_supplier", "")),
                            "lead_time": norm_text(getattr(c, "lead_time_days", "")),
                            "unit_price": self._extract_unit_price(notes),
                            "source": f"Product {prod.product_code}",
                            "notes": notes,
                        })
        except Exception:
            pass

        return rows

    def refresh_parts(self):
        treeview_clear(self.parts_tree)
        self.parts_rows = self._collect_parts()

        q = norm_text(self.search_var.get()).lower()
        for idx, r in enumerate(self.parts_rows):
            hay = f'{r["part_name"]} {r["part_number"]} {r["supplier"]} {r["source"]}'.lower()
            if q and q not in hay:
                continue
            price_text = f'${r["unit_price"]:,.2f}' if r["unit_price"] else ""
            self.parts_tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    r["part_name"],
                    r["part_number"],
                    f'{r["qty"]:.2f}' if r["qty"] else "",
                    f'{r["soh_qty"]:.2f}' if r["soh_qty"] else "",
                    r["supplier"],
                    r["lead_time"],
                    price_text,
                    r["source"],
                )
            )

    def _selected_row(self):
        sel = self.parts_tree.selection()
        if not sel:
            return None
        try:
            return self.parts_rows[int(sel[0])]
        except Exception:
            return None

    def show_part_details(self):
        row = self._selected_row()
        if not row:
            return
        self.price_var.set(f'{row["unit_price"]:.2f}' if row["unit_price"] else "")
        lines = [
            "PART DETAILS",
            "------------",
            f'Part Name: {row["part_name"]}',
            f'Part Number: {row["part_number"] or "-"}',
            f'Qty: {row["qty"]:.2f}',
            f'SOH Qty: {row["soh_qty"]:.2f}',
            f'Supplier: {row["supplier"] or "-"}',
            f'Lead Time: {row["lead_time"] or "-"}',
            f'Unit Price: ${row["unit_price"]:,.2f}' if row["unit_price"] else "Unit Price: -",
            f'Source: {row["source"]}',
        ]
        if row["notes"]:
            lines += ["", "Notes:", row["notes"]]

        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.configure(state="disabled")

    def save_selected_price(self):
        row = self._selected_row()
        if not row:
            show_warning("No Selection", "Select a part first.")
            return

        try:
            price = safe_float(self.price_var.get())
        except Exception:
            show_warning("Invalid Price", "Enter a valid number.")
            return

        try:
            repo = getattr(self.app, "repo", None)
            if repo is None:
                show_error("Save Error", "Repository not available.")
                return

            rows = repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
            changed = 0
            for r in rows:
                pn = norm_text(r.get("PartNumber"))
                name = norm_text(r.get("ComponentName"))
                if (row["part_number"] and pn == row["part_number"]) or (not row["part_number"] and name == row["part_name"]):
                    notes = norm_text(r.get("Notes"))
                    new_notes = self._merge_unit_price(notes, price)
                    comp_id = norm_text(r.get("ComponentID"))
                    if comp_id:
                        ok = repo.update_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", comp_id, {"Notes": new_notes})
                        if ok:
                            changed += 1

            if hasattr(repo, "save_workbook"):
                try:
                    repo.save_workbook()
                except Exception:
                    pass

            if changed == 0:
                show_warning("No Rows Updated", "No matching component rows were found to store the price.")
                return

            show_info("Price Saved", f"Updated unit price on {changed} component row(s).")
            self.refresh_parts()
        except Exception as exc:
            show_error("Save Error", str(exc))

    def open_selected_part(self):
        row = self._selected_row()
        if not row:
            show_warning("No Selection", "Select a part first.")
            return
        self.show_part_details()
