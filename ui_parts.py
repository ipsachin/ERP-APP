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
        self.edit_part_name_var = tk.StringVar()
        self.edit_part_number_var = tk.StringVar()
        self.edit_qty_var = tk.StringVar()
        self.edit_soh_var = tk.StringVar()
        self.edit_supplier_var = tk.StringVar()
        self.edit_lead_var = tk.StringVar()
        self.edit_price_var = tk.StringVar()
        self.edit_notes_var = tk.StringVar()
        self.parts_rows = []
        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=8)
        wrapper.pack(fill="both", expand=True)

        top = ttk.Frame(wrapper)
        top.pack(fill="x", pady=(0, 5))

        ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left")
        ttk.Label(top, text="Parts Manager", style="Title.TLabel").pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right")

        search_row = ttk.Frame(wrapper)
        search_row.pack(fill="x", pady=(0, 5))
        ttk.Label(search_row, text="Search").pack(side="left")
        ent = ttk.Entry(search_row, textvariable=self.search_var)
        ent.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(search_row, text="Open Selected", command=self.open_selected_part).pack(side="left")

        add_card = ttk.LabelFrame(wrapper, text="Add New Part", padding=6)
        add_card.pack(fill="x", pady=(0, 5))

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

        detail_card = ttk.LabelFrame(right, text="Part Details", padding=6)
        detail_card.pack(fill="both", expand=True)

        form = ttk.Frame(detail_card)
        form.pack(fill="x", pady=(0, 5))

        fields = [
            ("Part Name", self.edit_part_name_var),
            ("Part Number", self.edit_part_number_var),
            ("Qty", self.edit_qty_var),
            ("SOH Qty", self.edit_soh_var),
            ("Supplier", self.edit_supplier_var),
            ("Lead Time", self.edit_lead_var),
            ("Unit Price", self.edit_price_var),
            ("Notes", self.edit_notes_var),
        ]
        for idx, (label, var) in enumerate(fields):
            r = idx // 2
            c = (idx % 2) * 2
            ttk.Label(form, text=label).grid(row=r, column=c, sticky="w", padx=(0, 4), pady=2)
            ttk.Entry(form, textvariable=var).grid(row=r, column=c + 1, sticky="ew", padx=(0, 8), pady=2)

        btn_row = ttk.Frame(form)
        btn_row.grid(row=(len(fields) + 1) // 2, column=0, columnspan=4, sticky="ew", pady=(5, 0))
        ttk.Button(btn_row, text="Save Part", command=self.save_selected_part).pack(side="left", fill="x", expand=True, padx=(0, 3))
        ttk.Button(btn_row, text="Delete Part", command=self.delete_selected_part).pack(side="left", fill="x", expand=True, padx=(3, 0))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)

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
            "UnitPrice": price,
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
                    unit_price = safe_float(getattr(p, "unit_price", 0)) or self._extract_unit_price(notes)
                    rows.append({
                        "component_id": norm_text(getattr(p, "component_id", "")),
                        "owner_type": norm_text(getattr(p, "owner_type", "")),
                        "owner_code": norm_text(getattr(p, "owner_code", "")),
                        "part_name": norm_text(getattr(p, "part_name", "")) or norm_text(getattr(p, "component_name", "")),
                        "part_number": norm_text(getattr(p, "part_number", "")) or norm_text(getattr(p, "part_code", "")),
                        "qty": safe_float(getattr(p, "qty", 0)),
                        "soh_qty": safe_float(getattr(p, "soh_qty", 0)),
                        "supplier": norm_text(getattr(p, "preferred_supplier", "")) or norm_text(getattr(p, "supplier", "")),
                        "lead_time": norm_text(getattr(p, "lead_time_days", "")),
                        "unit_price": unit_price,
                        "source": "Parts",
                        "notes": notes,
                    })
                return rows
        except Exception:
            pass

        try:
            repo = getattr(self.app, "repo", None)
            if repo is not None:
                for r in repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS):
                    if norm_text(r.get("OwnerType")) != "PART":
                        continue
                    notes = norm_text(r.get("Notes"))
                    unit_price = safe_float(r.get("UnitPrice")) or self._extract_unit_price(notes)
                    rows.append({
                        "component_id": norm_text(r.get("ComponentID")),
                        "owner_type": norm_text(r.get("OwnerType")),
                        "owner_code": norm_text(r.get("OwnerCode")),
                        "part_name": norm_text(r.get("ComponentName")),
                        "part_number": norm_text(r.get("PartNumber")),
                        "qty": safe_float(r.get("Qty")),
                        "soh_qty": safe_float(r.get("SOHQty")),
                        "supplier": norm_text(r.get("PreferredSupplier")),
                        "lead_time": norm_text(r.get("LeadTimeDays")),
                        "unit_price": unit_price,
                        "source": "Parts",
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
        self.edit_part_name_var.set(row["part_name"])
        self.edit_part_number_var.set(row["part_number"])
        self.edit_qty_var.set(f'{row["qty"]:.2f}' if row["qty"] else "")
        self.edit_soh_var.set(f'{row["soh_qty"]:.2f}' if row["soh_qty"] else "")
        self.edit_supplier_var.set(row["supplier"])
        self.edit_lead_var.set(str(row["lead_time"] or ""))
        self.edit_price_var.set(f'{row["unit_price"]:.2f}' if row["unit_price"] else "")
        self.edit_notes_var.set(row["notes"])
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
        self.edit_price_var.set(self.price_var.get())
        self.save_selected_part()

    def save_selected_part(self):
        row = self._selected_row()
        if not row:
            show_warning("No Selection", "Select a part first.")
            return

        comp_id = norm_text(row.get("component_id"))
        if not comp_id:
            show_warning("No Selection", "Could not identify the selected part row.")
            return

        try:
            repo = getattr(self.app, "repo", None)
            if repo is None:
                show_error("Save Error", "Repository not available.")
                return

            part_name = norm_text(self.edit_part_name_var.get())
            if not part_name:
                show_warning("Missing Data", "Part Name is required.")
                return

            notes = norm_text(self.edit_notes_var.get())
            price = safe_float(self.edit_price_var.get())
            if price:
                notes = self._merge_unit_price(notes, price)

            updates = {
                "ComponentName": part_name,
                "Qty": safe_float(self.edit_qty_var.get()),
                "SOHQty": safe_float(self.edit_soh_var.get()),
                "PreferredSupplier": norm_text(self.edit_supplier_var.get()),
                "LeadTimeDays": norm_text(self.edit_lead_var.get()),
                "UnitPrice": price,
                "PartNumber": norm_text(self.edit_part_number_var.get()),
                "Notes": notes,
            }
            if "UpdatedOn" in AppConfig.COMPONENT_HEADERS:
                updates["UpdatedOn"] = ""

            changed = repo.update_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", comp_id, updates)

            if hasattr(repo, "save_workbook"):
                try:
                    repo.save_workbook()
                except Exception:
                    pass

            if not changed:
                show_warning("No Rows Updated", "The selected part row was not found.")
                return

            show_info("Part Saved", f"Updated part: {part_name}")
            self.refresh_parts()
        except Exception as exc:
            show_error("Save Error", str(exc))

    def _part_matches_row(self, candidate, row):
        if norm_text(candidate.get("OwnerType")) == "PART":
            return False
        comp_id = norm_text(row.get("component_id"))
        if comp_id and f"SourcePartID={comp_id}" in norm_text(candidate.get("Notes")):
            return True
        part_number = norm_text(row.get("part_number"))
        part_name = norm_text(row.get("part_name"))
        candidate_part_number = norm_text(candidate.get("PartNumber"))
        candidate_name = norm_text(candidate.get("ComponentName"))
        if part_number and candidate_part_number:
            return candidate_part_number.lower() == part_number.lower()
        return bool(part_name and candidate_name and candidate_name.lower() == part_name.lower())

    def _find_part_associations(self, row):
        return [item["display"] for item in self._find_part_association_rows(row)]

    def _find_part_association_rows(self, row):
        associations = []
        repo = getattr(self.app, "repo", None)
        if repo is None:
            return associations

        try:
            rows = repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
        except Exception:
            return associations

        for candidate in rows:
            if not self._part_matches_row(candidate, row):
                continue
            owner_type = norm_text(candidate.get("OwnerType")) or "UNKNOWN"
            owner_code = norm_text(candidate.get("OwnerCode")) or "-"
            component_name = norm_text(candidate.get("ComponentName")) or norm_text(row.get("part_name"))
            associations.append({
                "component_id": norm_text(candidate.get("ComponentID")),
                "display": f"{owner_type}: {owner_code} ({component_name})",
            })

        deduped = {}
        for association in associations:
            key = association["component_id"] or association["display"]
            deduped[key] = association
        return sorted(deduped.values(), key=lambda item: item["display"])

    def _confirm_delete_part(self, row, associations):
        dialog = tk.Toplevel(self)
        dialog.title("Delete Part")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.resizable(False, False)

        result = {"delete": False, "delete_associations": False}
        understands_var = tk.BooleanVar(value=False)
        delete_associations_var = tk.BooleanVar(value=False)

        body = ttk.Frame(dialog, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(
            body,
            text=f'Delete part "{row["part_name"]}" from the parts database?',
            style="Title.TLabel",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            body,
            text="This removes the master part record.",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        if associations:
            ttk.Label(
                body,
                text="This part appears to be used by:",
                wraplength=520,
                justify="left",
            ).pack(anchor="w")

            assoc_text = tk.Text(body, height=min(8, max(3, len(associations))), width=70, wrap="word")
            assoc_text.pack(fill="x", pady=(4, 8))
            assoc_text.insert("1.0", "\n".join(f' - {item["display"]}' for item in associations[:20]))
            if len(associations) > 20:
                assoc_text.insert("end", f"\n - ...and {len(associations) - 20} more")
            assoc_text.configure(state="disabled")

            ttk.Checkbutton(
                body,
                text="Also delete this part from all associated assemblies, products, and projects.",
                variable=delete_associations_var,
            ).pack(anchor="w", pady=(0, 8))
        else:
            ttk.Label(
                body,
                text="No associated assembly, product, or project BOM rows were found.",
                wraplength=520,
                justify="left",
            ).pack(anchor="w", pady=(0, 8))

        confirm_check = ttk.Checkbutton(
            body,
            text="I understand this deletion cannot be undone.",
            variable=understands_var,
        )
        confirm_check.pack(anchor="w", pady=(0, 12))

        btn_row = ttk.Frame(body)
        btn_row.pack(fill="x")

        delete_btn = ttk.Button(btn_row, text="Delete", state="disabled")
        cancel_btn = ttk.Button(btn_row, text="Cancel")
        cancel_btn.pack(side="right", padx=(8, 0))
        delete_btn.pack(side="right")

        def _sync_delete_state(*_args):
            delete_btn.configure(state="normal" if understands_var.get() else "disabled")

        def _cancel():
            dialog.destroy()

        def _delete():
            result["delete"] = True
            result["delete_associations"] = bool(delete_associations_var.get())
            dialog.destroy()

        understands_var.trace_add("write", _sync_delete_state)
        delete_btn.configure(command=_delete)
        cancel_btn.configure(command=_cancel)
        dialog.protocol("WM_DELETE_WINDOW", _cancel)

        dialog.update_idletasks()
        parent = self.winfo_toplevel()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - dialog.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - dialog.winfo_height()) // 2)
        dialog.geometry(f"+{x}+{y}")
        dialog.wait_window()
        return result

    def delete_selected_part(self):
        row = self._selected_row()
        if not row:
            show_warning("No Selection", "Select a part first.")
            return

        comp_id = norm_text(row.get("component_id"))
        if not comp_id:
            show_warning("No Selection", "Could not identify the selected part row.")
            return

        associations = self._find_part_association_rows(row)
        confirmation = self._confirm_delete_part(row, associations)
        if not confirmation["delete"]:
            return

        try:
            repo = getattr(self.app, "repo", None)
            if repo is None:
                show_error("Delete Error", "Repository not available.")
                return

            associated_deleted = 0
            if confirmation["delete_associations"]:
                for association in associations:
                    assoc_id = norm_text(association.get("component_id"))
                    if assoc_id and repo.delete_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", assoc_id):
                        associated_deleted += 1

            deleted = repo.delete_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", comp_id)
            if hasattr(repo, "save_workbook"):
                try:
                    repo.save_workbook()
                except Exception:
                    pass

            if not deleted:
                show_warning("Not Deleted", "The selected part row was not found.")
                return

            message = f'Deleted part: {row["part_name"]}'
            if confirmation["delete_associations"]:
                message += f"\nDeleted associated BOM rows: {associated_deleted}"
            show_info("Part Deleted", message)
            self._clear_part_editor()
            self.refresh_parts()
        except Exception as exc:
            show_error("Delete Error", str(exc))

    def _clear_part_editor(self):
        self.price_var.set("")
        self.edit_part_name_var.set("")
        self.edit_part_number_var.set("")
        self.edit_qty_var.set("")
        self.edit_soh_var.set("")
        self.edit_supplier_var.set("")
        self.edit_lead_var.set("")
        self.edit_price_var.set("")
        self.edit_notes_var.set("")
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.configure(state="disabled")

    def open_selected_part(self):
        row = self._selected_row()
        if not row:
            show_warning("No Selection", "Select a part first.")
            return
        self.show_part_details()
