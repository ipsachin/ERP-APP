# ============================================================
# ui_products.py
# Product builder page for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog
from typing import Dict, List, Optional, Tuple

from app_config import AppConfig
from models import ProductRecord, ProductModuleLinkRecord, ProductDocumentRecord, WorkOrderRecord
from ui_common import (
    BasePage,
    DraggableListbox,
    attach_tooltip,
    treeview_clear,
    set_combobox_values,
    show_warning,
    show_error,
    make_readonly_text,
    get_text_value,
    set_text_readonly,
    open_file_with_default_app,
    Validators,
    Dialogs,
)


def norm_text(value) -> str:
    return str(value or "").strip()


class ProductPage(BasePage):
    PAGE_NAME = "products"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        self.current_workorder_id: Optional[str] = None
        self.current_product_document_id: Optional[str] = None

        # ----------------------------------------------------
        # Product vars
        # ----------------------------------------------------
        self.product_code_var = tk.StringVar()
        self.quote_ref_var = tk.StringVar()
        self.product_name_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.revision_var = tk.StringVar(value="R0")
        self.status_var_local = tk.StringVar(value="Draft")

        # ----------------------------------------------------
        # Product module builder vars
        # ----------------------------------------------------
        self.available_module_var = tk.StringVar()
        self.module_dependency_var = tk.StringVar()

        # ----------------------------------------------------
        # Product doc vars
        # ----------------------------------------------------
        self.doc_section_var = tk.StringVar()
        self.doc_type_var = tk.StringVar(value="Other")
        self.doc_instruction_var = tk.StringVar()

        # ----------------------------------------------------
        # Workorder vars
        # ----------------------------------------------------
        self.workorder_name_var = tk.StringVar()
        self.workorder_stage_var = tk.StringVar()
        self.workorder_owner_var = tk.StringVar()
        self.workorder_due_var = tk.StringVar()
        self.workorder_status_var = tk.StringVar(value="Open")
        self.workorder_notes_var = tk.StringVar()

        self.product_select_var = tk.StringVar()

        self._build_ui()

    # ========================================================
    # UI
    # ========================================================

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        self._build_topbar(wrapper)

        self.product_summary_label = ttk.Label(
            wrapper,
            text="No product selected",
            style="Title.TLabel"
        )
        self.product_summary_label.pack(anchor="w", pady=(0, 10))

        paned = ttk.Panedwindow(wrapper, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, padding=6)
        right = ttk.Frame(paned, padding=6)

        paned.add(left, weight=2)
        paned.add(right, weight=3)

        self._build_left_panel(left)
        self._build_right_panel(right)


    def _build_topbar(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 8))

        left = ttk.Frame(top)
        left.pack(side="left", fill="x", expand=True)

        ttk.Button(left, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)

        ttk.Label(left, text="Open Product").pack(side="left", padx=(12, 4))

        self.product_select_combo = ttk.Combobox(
            left,
            textvariable=self.product_select_var,
            state="readonly",
            width=55
        )
        self.product_select_combo.pack(side="left", padx=4)
        self.product_select_combo.bind("<<ComboboxSelected>>", lambda e: self.open_selected_product_from_dropdown())

        ttk.Button(left, text="Open", command=self.open_selected_product_from_dropdown).pack(side="left", padx=4)

        right = ttk.Frame(top)
        right.pack(side="right")

        ttk.Button(right, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)
        ttk.Button(right, text="Email Quote PDF", command=self.email_product_pdf).pack(side="right", padx=2)
        ttk.Button(right, text="Generate Quote PDF", command=self.generate_product_pdf).pack(side="right", padx=2)


    def refresh_product_selector(self):
        if not self.app.workbook_manager.has_workbook():
            self.product_select_combo["values"] = []
            return

        records = self.app.services.products.list_products()
        values = [
            f"{p.product_code} | {p.quote_ref} | {p.product_name}"
            for p in records
        ]
        set_combobox_values(self.product_select_combo, values, keep_current=True)

        selected_code = getattr(self.app, "selected_product_code", "").strip()
        if selected_code:
            for p in records:
                if p.product_code == selected_code:
                    self.product_select_var.set(f"{p.product_code} | {p.quote_ref} | {p.product_name}")
                    break

    def open_selected_product_from_dropdown(self):
        selected = self.product_select_var.get().strip()
        if not selected:
            return

        product_code = selected.split("|")[0].strip()
        self.app.set_selected_product(product_code)
        self.refresh_page()

    def _build_left_panel(self, parent):
        self._build_product_details_card(parent)
        self._build_product_module_builder_card(parent)
        self._build_product_document_card(parent)
        self._build_workorder_card(parent)

    def _build_right_panel(self, parent):
        tabs = ttk.Notebook(parent)
        tabs.pack(fill="both", expand=True)

        self.modules_tab = ttk.Frame(tabs)
        self.docs_tab = ttk.Frame(tabs)
        self.workorders_tab = ttk.Frame(tabs)
        self.summary_tab = ttk.Frame(tabs)

        tabs.add(self.modules_tab, text="Product Modules")
        tabs.add(self.docs_tab, text="Instruction Manuals / Docs")
        tabs.add(self.workorders_tab, text="Work Orders")
        tabs.add(self.summary_tab, text="Summary")

        self._build_modules_tab(self.modules_tab)
        self._build_docs_tab(self.docs_tab)
        self._build_workorders_tab(self.workorders_tab)
        self._build_summary_tab(self.summary_tab)

    # --------------------------------------------------------
    # Left cards
    # --------------------------------------------------------

    def _build_product_details_card(self, parent):
        card = ttk.LabelFrame(parent, text="Product Details", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Product Code").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.product_code_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Quote Ref").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.quote_ref_var).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Product Name").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.product_name_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Description").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.description_var).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Revision").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.revision_var).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.status_var_local,
            values=AppConfig.PRODUCT_STATUSES,
            state="readonly"
        ).grid(row=5, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        save_btn = ttk.Button(btn_row, text="Save Product", command=self.save_product)
        save_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(save_btn, "Create a new product or update the selected product.")

        del_btn = ttk.Button(btn_row, text="Delete Product", command=self.delete_product)
        del_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(del_btn, "Delete selected product and linked product-level records.")

        card.columnconfigure(1, weight=1)

    def _build_product_module_builder_card(self, parent):
        card = ttk.LabelFrame(parent, text="Build Product from Modules", style="Card.TLabelframe", padding=12)
        card.pack(fill="both", expand=True, pady=6)

        row1 = ttk.Frame(card)
        row1.pack(fill="x", pady=(0, 6))

        ttk.Label(row1, text="Available Module").pack(side="left")
        self.available_module_combo = ttk.Combobox(
            row1,
            textvariable=self.available_module_var,
            state="readonly",
            width=38
        )
        self.available_module_combo.pack(side="left", padx=6, fill="x", expand=True)

        add_btn = ttk.Button(row1, text="Add Module", command=self.add_module_to_product)
        add_btn.pack(side="left", padx=2)
        attach_tooltip(add_btn, "Add selected module to this product.")

        rem_btn = ttk.Button(row1, text="Remove Selected", command=self.remove_selected_module)
        rem_btn.pack(side="left", padx=2)

        row2 = ttk.Frame(card)
        row2.pack(fill="x", pady=(0, 6))

        ttk.Label(row2, text="Dependency Module").pack(side="left")
        self.module_dependency_combo = ttk.Combobox(
            row2,
            textvariable=self.module_dependency_var,
            state="readonly",
            width=38
        )
        self.module_dependency_combo.pack(side="left", padx=6, fill="x", expand=True)

        dep_btn = ttk.Button(row2, text="Set Dependency", command=self.set_selected_module_dependency)
        dep_btn.pack(side="left", padx=2)
        attach_tooltip(dep_btn, "Set dependency for selected product module.")

        qty_btn = ttk.Button(row2, text="Set Qty", command=self.set_selected_module_qty)
        qty_btn.pack(side="left", padx=2)

        info = ttk.Label(
            card,
            text="Drag to reorder module build sequence. Double-click list item to open module page.",
            style="Sub.TLabel"
        )
        info.pack(anchor="w", pady=(0, 4))

        lb_wrap = ttk.Frame(card)
        lb_wrap.pack(fill="both", expand=True)

        self.module_order_listbox = DraggableListbox(
            lb_wrap,
            selectmode=tk.SINGLE,
            font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
            activestyle="none"
        )
        self.module_order_listbox.pack(side="left", fill="both", expand=True)

        self.module_order_listbox.bind("<Double-1>", lambda e: self.open_module_from_list())

        sb = ttk.Scrollbar(lb_wrap, orient="vertical", command=self.module_order_listbox.yview)
        self.module_order_listbox.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        btn_row = ttk.Frame(card)
        btn_row.pack(fill="x", pady=(6, 0))

        ttk.Button(btn_row, text="Save Order", command=self.save_module_order).pack(side="left", padx=2)
        ttk.Button(btn_row, text="Move Up", command=lambda: self.move_selected_module(-1)).pack(side="left", padx=2)
        ttk.Button(btn_row, text="Move Down", command=lambda: self.move_selected_module(1)).pack(side="left", padx=2)

    def _build_product_document_card(self, parent):
        card = ttk.LabelFrame(parent, text="Product Documents", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Section Name").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.doc_section_var).grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Document Type").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.doc_type_var,
            values=AppConfig.DOC_TYPES,
            state="readonly"
        ).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Instruction Text").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.doc_instruction_var).grid(row=2, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Add File", command=self.add_product_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Open Selected", command=self.open_selected_product_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_product_document).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    def _build_workorder_card(self, parent):
        card = ttk.LabelFrame(parent, text="Product Work Order / Workflow", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Work Order Name").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.workorder_name_var).grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Stage").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.workorder_stage_var,
            values=AppConfig.WORKORDER_STAGES,
            state="readonly"
        ).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Owner").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.workorder_owner_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Due Date").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.workorder_due_var).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.workorder_status_var,
            values=AppConfig.WORKORDER_STATUSES,
            state="readonly"
        ).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Notes").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.workorder_notes_var).grid(row=5, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Add Work Order", command=self.add_workorder).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Update Selected", command=self.update_selected_workorder).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_workorder).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Clear", command=self.clear_workorder_editor).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    # --------------------------------------------------------
    # Right tabs
    # --------------------------------------------------------

    def _build_modules_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("Order", "ModuleCode", "ModuleName", "Qty", "Dependency", "Description")
        self.product_modules_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Order", 70),
            ("ModuleCode", 260),
            ("ModuleName", 220),
            ("Qty", 70),
            ("Dependency", 220),
            ("Description", 380),
        ]:
            self.product_modules_tree.heading(col, text=col)
            self.product_modules_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.product_modules_tree.yview)
        self.product_modules_tree.configure(yscrollcommand=sb.set)

        self.product_modules_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_docs_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("SectionName", "DocName", "DocType", "FilePath", "AddedOn")
        self.product_document_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("SectionName", 150),
            ("DocName", 220),
            ("DocType", 150),
            ("FilePath", 420),
            ("AddedOn", 160),
        ]:
            self.product_document_tree.heading(col, text=col)
            self.product_document_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.product_document_tree.yview)
        self.product_document_tree.configure(yscrollcommand=sb.set)

        self.product_document_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_workorders_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("Name", "Stage", "Owner", "DueDate", "Status", "Notes")
        self.workorder_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Name", 220),
            ("Stage", 150),
            ("Owner", 160),
            ("DueDate", 120),
            ("Status", 120),
            ("Notes", 320),
        ]:
            self.workorder_tree.heading(col, text=col)
            self.workorder_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.workorder_tree.yview)
        self.workorder_tree.configure(yscrollcommand=sb.set)

        self.workorder_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.workorder_tree.bind("<ButtonRelease-1>", lambda e: self.load_selected_workorder())
        self.workorder_tree.bind("<Double-1>", lambda e: self.load_selected_workorder())

    def _build_summary_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        self.summary_text = make_readonly_text(wrap, height=26)
        self.summary_text.pack(fill="both", expand=True)

    # ========================================================
    # Page refresh
    # ========================================================

    def refresh_page(self):
        if not self.require_workbook():
            return

        product_code = getattr(self.app, "selected_product_code", "").strip()
        if not product_code:
            self._clear_all_views()
            self.product_summary_label.config(text="No product selected")
            return

        try:
            bundle = self.app.services.products.get_product_bundle(product_code)
            if not bundle.product:
                self._clear_all_views()
                self.product_summary_label.config(text="Product not found")
                return

            self._load_product_into_form(bundle.product)
            self._load_module_links(bundle)
            self._load_product_documents(bundle.product_documents or [])
            self._load_workorders(bundle.workorders or [])
            self._load_summary(bundle)
            self._refresh_available_modules()
            self._refresh_dependency_modules(bundle.module_links or [])

            self.product_summary_label.config(
                text=f"🧩 {bundle.product.product_name} | Quote: {bundle.product.quote_ref} | Rev: {bundle.product.revision} | Code: {bundle.product.product_code}"
            )
            self.set_status(f"Loaded product: {bundle.product.product_code}")

        except Exception as exc:
            show_error("Product Refresh Error", str(exc))

    def _clear_all_views(self):
        self.product_code_var.set("")
        self.quote_ref_var.set("")
        self.product_name_var.set("")
        self.description_var.set("")
        self.revision_var.set("R0")
        self.status_var_local.set("Draft")

        treeview_clear(self.product_modules_tree)
        treeview_clear(self.product_document_tree)
        treeview_clear(self.workorder_tree)
        set_text_readonly(self.summary_text, "")
        self.module_order_listbox.delete(0, tk.END)

        self.clear_workorder_editor()

    def _load_product_into_form(self, product: ProductRecord):
        self.product_code_var.set(product.product_code)
        self.quote_ref_var.set(product.quote_ref)
        self.product_name_var.set(product.product_name)
        self.description_var.set(product.description)
        self.revision_var.set(product.revision or "R0")
        self.status_var_local.set(product.status or "Draft")

    def _load_module_links(self, bundle):
        treeview_clear(self.product_modules_tree)
        self.module_order_listbox.delete(0, tk.END)

        module_name_lookup = {}
        module_desc_lookup = {}
        for m in (bundle.modules or []):
            module_name_lookup[m.module_code] = m.module_name
            module_desc_lookup[m.module_code] = m.description

        for link in (bundle.module_links or []):
            dep_display = link.dependency_module_code or ""
            self.product_modules_tree.insert(
                "",
                "end",
                values=(
                    link.module_order,
                    link.module_code,
                    module_name_lookup.get(link.module_code, ""),
                    link.module_qty,
                    dep_display,
                    module_desc_lookup.get(link.module_code, ""),
                ),
                tags=(link.link_id,),
            )
            self.module_order_listbox.insert(
                tk.END,
                self._format_module_list_item(link.module_order, link.module_qty, link.module_code)
            )

    def _load_product_documents(self, docs: List[ProductDocumentRecord]):
        treeview_clear(self.product_document_tree)
        for doc in docs:
            self.product_document_tree.insert(
                "",
                "end",
                values=(
                    doc.section_name,
                    doc.doc_name,
                    doc.doc_type,
                    doc.file_path,
                    doc.added_on,
                ),
                tags=(doc.prod_doc_id,),
            )

    def _load_workorders(self, workorders: List[WorkOrderRecord]):
        treeview_clear(self.workorder_tree)
        for w in workorders:
            self.workorder_tree.insert(
                "",
                "end",
                values=(
                    w.workorder_name,
                    w.stage,
                    w.owner,
                    w.due_date,
                    w.status,
                    w.notes,
                ),
                tags=(w.workorder_id,),
            )

    def _load_summary(self, bundle):
        product = bundle.product
        links = bundle.module_links or []
        docs = bundle.product_documents or []
        workorders = bundle.workorders or []
        tasks_by_module = bundle.tasks_by_module or {}

        lines = [
            f"Product Code: {product.product_code}",
            f"Quote Ref: {product.quote_ref}",
            f"Product Name: {product.product_name}",
            f"Description: {product.description}",
            f"Revision: {product.revision}",
            f"Status: {product.status}",
            "",
            f"Assigned Modules: {len(links)}",
            f"Product Documents: {len(docs)}",
            f"Work Orders: {len(workorders)}",
            f"Aggregated Product Hours: {float(bundle.total_hours or 0.0):.2f}",
            "",
            "Module Breakdown:",
        ]

        if not links:
            lines.append(" - No modules assigned.")
        else:
            for link in links:
                module_tasks = tasks_by_module.get(link.module_code, [])
                module_hours = sum(float(t.estimated_hours or 0.0) for t in module_tasks)
                dep_text = f" | Depends on: {link.dependency_module_code}" if norm_text(link.dependency_module_code) else ""
                lines.append(
                    f" {link.module_order:02d}. {link.module_code} | Qty {link.module_qty} | Hours {module_hours * link.module_qty:.2f}{dep_text}"
                )
                for t in module_tasks:
                    indent = "   - " if norm_text(t.parent_task_id) else "   * "
                    lines.append(f"{indent}{t.task_name} ({float(t.estimated_hours or 0.0):.2f} hrs)")
                lines.append("")

        set_text_readonly(self.summary_text, "\n".join(lines))

    def _refresh_available_modules(self):
        modules = self.app.services.modules.list_modules()
        available = [m.module_code for m in modules]
        set_combobox_values(self.available_module_combo, available, keep_current=True)

    def _refresh_dependency_modules(self, links: List[ProductModuleLinkRecord]):
        module_codes = [l.module_code for l in links]
        set_combobox_values(self.module_dependency_combo, module_codes, keep_current=True)

    # ========================================================
    # Helpers
    # ========================================================

    def _selected_product_code(self) -> str:
        return self.product_code_var.get().strip() or getattr(self.app, "selected_product_code", "").strip()

    def _format_module_list_item(self, order_num: int, qty: int, module_code: str) -> str:
        return f"{order_num:02d} | Qty {qty} | {module_code}"

    def _parse_module_list_item(self, text: str) -> Tuple[int, int, str]:
        parts = [p.strip() for p in str(text).split("|")]
        order_num = 0
        qty = 1
        module_code = ""

        if len(parts) >= 1:
            try:
                order_num = int(parts[0])
            except Exception:
                order_num = 0
        if len(parts) >= 2 and parts[1].startswith("Qty"):
            try:
                qty = int(parts[1].replace("Qty", "").strip())
            except Exception:
                qty = 1
        if len(parts) >= 3:
            module_code = parts[2].strip()

        return order_num, qty, module_code

    def _selected_module_list_item(self) -> str:
        sel = self.module_order_listbox.curselection()
        if not sel:
            return ""
        return self.module_order_listbox.get(sel[0])

    # ========================================================
    # Product CRUD
    # ========================================================

    def save_product(self):
        if not self.require_workbook():
            return

        try:
            quote_ref = self.quote_ref_var.get().strip()
            product_name = Validators.require_text(self.product_name_var.get(), "Product name")
            description = self.description_var.get().strip()
            revision = self.revision_var.get().strip() or "R0"
            status = self.status_var_local.get().strip() or "Draft"
            existing_code = self.product_code_var.get().strip() or None

            new_code = self.app.services.products.create_or_update_product(
                quote_ref=quote_ref,
                product_name=product_name,
                description=description,
                revision=revision,
                status=status,
                existing_product_code=existing_code,
            )

            self.app.set_selected_product(new_code)
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Product saved: {new_code}")

        except Exception as exc:
            show_error("Save Product Error", str(exc))

    def delete_product(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "No product selected.")
            return

        if not tk.messagebox.askyesno(
            "Delete Product",
            "Delete this product?\n\nThis will also delete product module links, product docs, and product work orders."
        ):
            return

        try:
            self.app.services.products.delete_product(product_code, delete_docs_files=False)
            self.app.set_selected_product("")
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Product deleted: {product_code}")
        except Exception as exc:
            show_error("Delete Product Error", str(exc))

    # ========================================================
    # Product modules
    # ========================================================

    def add_module_to_product(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Save or select a product first.")
            return

        module_code = self.available_module_var.get().strip()
        if not module_code:
            show_warning("No Module", "Select a module first.")
            return

        try:
            self.app.services.products.add_module_to_product(
                product_code=product_code,
                module_code=module_code,
                qty=1,
                dependency_module_code="",
                notes="",
            )
            self.refresh_page()
            self.set_status(f"Module added to product: {module_code}")
        except Exception as exc:
            show_error("Add Module Error", str(exc))

    def remove_selected_module(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        item = self._selected_module_list_item()
        if not product_code or not item:
            show_warning("No Selection", "Select a product module first.")
            return

        _order_num, _qty, module_code = self._parse_module_list_item(item)

        if not tk.messagebox.askyesno("Remove Module", f"Remove module from product?\n\n{module_code}"):
            return

        try:
            self.app.services.products.remove_module_from_product(product_code, module_code)
            self.refresh_page()
            self.set_status(f"Module removed: {module_code}")
        except Exception as exc:
            show_error("Remove Module Error", str(exc))

    def set_selected_module_qty(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        item = self._selected_module_list_item()
        if not product_code or not item:
            show_warning("No Selection", "Select a product module first.")
            return

        _order_num, current_qty, module_code = self._parse_module_list_item(item)

        qty = Dialogs.ask_int(
            "Set Module Quantity",
            f"Enter quantity for:\n{module_code}",
            minvalue=1,
            initialvalue=current_qty or 1
        )
        if qty is None:
            return

        try:
            self.app.services.products.set_module_qty(product_code, module_code, qty)
            self.refresh_page()
            self.set_status(f"Module quantity updated: {module_code} -> {qty}")
        except Exception as exc:
            show_error("Set Quantity Error", str(exc))

    def set_selected_module_dependency(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        item = self._selected_module_list_item()
        if not product_code or not item:
            show_warning("No Selection", "Select a product module first.")
            return

        _order_num, _qty, module_code = self._parse_module_list_item(item)
        dependency_module_code = self.module_dependency_var.get().strip()

        if dependency_module_code == module_code:
            show_warning("Invalid Dependency", "A module cannot depend on itself.")
            return

        try:
            self.app.services.products.set_module_dependency(
                product_code=product_code,
                module_code=module_code,
                dependency_module_code=dependency_module_code,
            )
            self.refresh_page()
            self.set_status(f"Dependency updated for {module_code}")
        except Exception as exc:
            show_error("Set Dependency Error", str(exc))

    def move_selected_module(self, direction: int):
        sel = self.module_order_listbox.curselection()
        if not sel:
            return

        idx = sel[0]
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= self.module_order_listbox.size():
            return

        txt = self.module_order_listbox.get(idx)
        self.module_order_listbox.delete(idx)
        self.module_order_listbox.insert(new_idx, txt)
        self.module_order_listbox.selection_clear(0, tk.END)
        self.module_order_listbox.selection_set(new_idx)

    def save_module_order(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Select a product first.")
            return

        try:
            ordered_items: List[Tuple[str, int]] = []
            for i in range(self.module_order_listbox.size()):
                item = self.module_order_listbox.get(i)
                _order_num, qty, module_code = self._parse_module_list_item(item)
                if module_code:
                    ordered_items.append((module_code, qty))

            self.app.services.products.save_module_order(product_code, ordered_items)
            self.refresh_page()
            self.set_status("Product module order saved.")
        except Exception as exc:
            show_error("Save Order Error", str(exc))

    def open_module_from_list(self):
        item = self._selected_module_list_item()
        if not item:
            return
        _order_num, _qty, module_code = self._parse_module_list_item(item)
        if not module_code:
            return
        self.app.set_selected_module(module_code)
        self.show_page("modules")

    # ========================================================
    # Product docs
    # ========================================================

    def add_product_document(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Save or select a product first.")
            return

        file_path = filedialog.askopenfilename(
            title="Select Product Document",
            filetypes=[
                ("Documents", "*.pdf *.dwg *.dxf *.step *.stp *.sldprt *.sldasm *.doc *.docx *.xls *.xlsx *.png *.jpg *.jpeg"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            self.app.services.products.add_product_document(
                product_code=product_code,
                source_file_path=file_path,
                section_name=self.doc_section_var.get().strip(),
                doc_type=self.doc_type_var.get().strip() or "Other",
                instruction_text=self.doc_instruction_var.get().strip(),
                copy_file=True,
            )
            self.refresh_page()
            self.set_status("Product document added successfully.")
        except Exception as exc:
            show_error("Add Product Document Error", str(exc))

    def _get_selected_product_document_info(self):
        sel = self.product_document_tree.selection()
        if not sel:
            return None, None

        tags = self.product_document_tree.item(sel[0], "tags")
        vals = self.product_document_tree.item(sel[0], "values")

        doc_id = tags[0] if tags else ""
        file_path = vals[3] if vals and len(vals) > 3 else ""
        return doc_id, file_path

    def open_selected_product_document(self):
        doc_id, file_path = self._get_selected_product_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a product document first.")
            return

        try:
            file_path = self.app.services.products.resolve_product_document_open_path(doc_id, str(file_path))
            open_file_with_default_app(str(file_path))
        except Exception as exc:
            show_error("Open Product Document Error", str(exc))

    def delete_selected_product_document(self):
        if not self.require_workbook():
            return

        doc_id, _file_path = self._get_selected_product_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a product document first.")
            return

        if not tk.messagebox.askyesno("Delete Product Document", "Delete the selected product document record?"):
            return

        try:
            self.app.services.products.delete_product_document(doc_id, delete_file=False)
            self.refresh_page()
            self.set_status("Product document deleted successfully.")
        except Exception as exc:
            show_error("Delete Product Document Error", str(exc))

    # ========================================================
    # Workorders
    # ========================================================

    def clear_workorder_editor(self):
        self.current_workorder_id = None
        self.workorder_name_var.set("")
        self.workorder_stage_var.set("")
        self.workorder_owner_var.set("")
        self.workorder_due_var.set("")
        self.workorder_status_var.set("Open")
        self.workorder_notes_var.set("")

    def add_workorder(self):
        if not self.require_workbook():
            return

        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Save or select a product first.")
            return

        try:
            workorder_name = Validators.require_text(self.workorder_name_var.get(), "Work order name")

            self.app.services.products.add_product_workorder(
                product_code=product_code,
                workorder_name=workorder_name,
                stage=self.workorder_stage_var.get().strip(),
                owner=self.workorder_owner_var.get().strip(),
                due_date=self.workorder_due_var.get().strip(),
                status=self.workorder_status_var.get().strip() or "Open",
                notes=self.workorder_notes_var.get().strip(),
            )

            self.clear_workorder_editor()
            self.refresh_page()
            self.set_status("Work order added successfully.")
        except Exception as exc:
            show_error("Add Work Order Error", str(exc))

    def load_selected_workorder(self):
        sel = self.workorder_tree.selection()
        if not sel:
            return

        tags = self.workorder_tree.item(sel[0], "tags")
        if not tags:
            return

        workorder_id = tags[0]
        product_code = self._selected_product_code()
        workorders = self.app.services.products.get_product_workorders(product_code)

        selected = None
        for w in workorders:
            if w.workorder_id == workorder_id:
                selected = w
                break

        if not selected:
            return

        self.current_workorder_id = selected.workorder_id
        self.workorder_name_var.set(selected.workorder_name)
        self.workorder_stage_var.set(selected.stage)
        self.workorder_owner_var.set(selected.owner)
        self.workorder_due_var.set(selected.due_date)
        self.workorder_status_var.set(selected.status or "Open")
        self.workorder_notes_var.set(selected.notes)

    def update_selected_workorder(self):
        if not self.require_workbook():
            return

        if not self.current_workorder_id:
            show_warning("No Selection", "Select a work order first.")
            return

        try:
            workorder_name = Validators.require_text(self.workorder_name_var.get(), "Work order name")

            self.app.services.products.update_workorder(
                self.current_workorder_id,
                {
                    "WorkOrderName": workorder_name,
                    "Stage": self.workorder_stage_var.get().strip(),
                    "Owner": self.workorder_owner_var.get().strip(),
                    "DueDate": self.workorder_due_var.get().strip(),
                    "Status": self.workorder_status_var.get().strip() or "Open",
                    "Notes": self.workorder_notes_var.get().strip(),
                }
            )

            self.refresh_page()
            self.set_status("Work order updated successfully.")
        except Exception as exc:
            show_error("Update Work Order Error", str(exc))

    def delete_selected_workorder(self):
        if not self.require_workbook():
            return

        if not self.current_workorder_id:
            show_warning("No Selection", "Select a work order first.")
            return

        if not tk.messagebox.askyesno("Delete Work Order", "Delete the selected work order?"):
            return

        try:
            self.app.services.products.delete_workorder(self.current_workorder_id)
            self.clear_workorder_editor()
            self.refresh_page()
            self.set_status("Work order deleted successfully.")
        except Exception as exc:
            show_error("Delete Work Order Error", str(exc))

    # ========================================================
    # PDF / email hooks
    # ========================================================

    def generate_product_pdf(self):
        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Select a product first.")
            return

        if not hasattr(self.app, "reports"):
            show_warning("Not Ready", "Report service is not wired yet. We will connect it when reports.py and main.py are added.")
            return

        try:
            self.app.reports.generate_product_quote_dialog(product_code)
        except Exception as exc:
            show_error("Generate Product PDF Error", str(exc))

    def email_product_pdf(self):
        product_code = self._selected_product_code()
        if not product_code:
            show_warning("No Product", "Select a product first.")
            return

        if not hasattr(self.app, "reports") or not hasattr(self.app, "mailer"):
            show_warning("Not Ready", "Mailer/report services are not wired yet. We will connect them in the final files.")
            return

        try:
            self.app.reports.email_product_quote_dialog(product_code)
        except Exception as exc:
            show_error("Email Product PDF Error", str(exc))

# ============================================================
# PATCH: scrollable product page + direct product parts + searchable combos
# ============================================================
from models import ComponentRecord


def _erp_bind_filterable_combobox(combo, values_getter):
    def _apply(event=None):
        try:
            all_values = [str(v) for v in values_getter()]
            typed = combo.get().strip().lower()
            if typed:
                filtered = [v for v in all_values if typed in v.lower()]
            else:
                filtered = all_values
            combo["values"] = filtered
        except Exception:
            pass
    combo.bind("<KeyRelease>", _apply, add="+")
    combo.bind("<Button-1>", _apply, add="+")
    return combo


def _patched_product_build_ui(self):
    if not hasattr(self, 'part_name_var'):
        self.part_name_var = tk.StringVar()
        self.part_number_var = tk.StringVar()
        self.part_qty_var = tk.StringVar(value='1')
        self.part_soh_var = tk.StringVar(value='0')
        self.part_supplier_var = tk.StringVar()
        self.part_lead_time_var = tk.StringVar(value='0')
        self.part_notes_var = tk.StringVar()
        self.current_product_part_id = None

    wrapper = ttk.Frame(self, padding=14)
    wrapper.pack(fill="both", expand=True)
    self._build_topbar(wrapper)
    self.product_summary_label = ttk.Label(wrapper, text="No product selected", style="Title.TLabel")
    self.product_summary_label.pack(anchor="w", pady=(0, 10))

    paned = ttk.Panedwindow(wrapper, orient="horizontal")
    paned.pack(fill="both", expand=True)

    left_host = ttk.Frame(paned)
    right = ttk.Frame(paned, padding=6)
    paned.add(left_host, weight=2)
    paned.add(right, weight=3)

    self.left_canvas = tk.Canvas(left_host, highlightthickness=0, bg=AppConfig.COLOR_BG)
    self.left_scrollbar = ttk.Scrollbar(left_host, orient="vertical", command=self.left_canvas.yview)
    self.left_frame = ttk.Frame(self.left_canvas)
    self.left_frame.bind("<Configure>", lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all")))
    self.canvas_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
    self.left_canvas.bind("<Configure>", lambda e: self.left_canvas.itemconfig(self.canvas_window, width=e.width))
    self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)
    self.left_canvas.pack(side="left", fill="both", expand=True)
    self.left_scrollbar.pack(side="right", fill="y")

    self._build_left_panel(self.left_frame)
    self._build_right_panel(right)

    def _on_mousewheel(event):
        try:
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
    self.left_canvas.bind_all("<MouseWheel>", _on_mousewheel)


def _patched_product_topbar(self, parent):
    top = ttk.Frame(parent)
    top.pack(fill="x", pady=(0, 8))
    left = ttk.Frame(top)
    left.pack(side="left", fill="x", expand=True)
    ttk.Button(left, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)
    ttk.Label(left, text="Open Product").pack(side="left", padx=(12, 4))
    self.product_select_combo = ttk.Combobox(left, textvariable=self.product_select_var, state="normal", width=55)
    self.product_select_combo.pack(side="left", padx=4)
    self.product_select_combo.bind("<<ComboboxSelected>>", lambda e: self.open_selected_product_from_dropdown())
    ttk.Button(left, text="Open", command=self.open_selected_product_from_dropdown).pack(side="left", padx=4)
    _erp_bind_filterable_combobox(self.product_select_combo, lambda: getattr(self.product_select_combo, '_all_values', self.product_select_combo.cget('values')))

    right = ttk.Frame(top)
    right.pack(side="right")
    ttk.Button(right, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)
    ttk.Button(right, text="Email Quote PDF", command=self.email_product_pdf).pack(side="right", padx=2)
    ttk.Button(right, text="Generate Quote PDF", command=self.generate_product_pdf).pack(side="right", padx=2)


def _patched_refresh_product_selector(self):
    if not self.app.workbook_manager.has_workbook():
        self.product_select_combo["values"] = []
        self.product_select_combo._all_values = []
        return
    records = self.app.services.products.list_products()
    values = [f"{p.product_code} | {p.quote_ref} | {p.product_name}" for p in records]
    self.product_select_combo["values"] = values
    self.product_select_combo._all_values = values
    selected_code = getattr(self.app, "selected_product_code", "").strip()
    if selected_code:
        for p in records:
            if p.product_code == selected_code:
                self.product_select_var.set(f"{p.product_code} | {p.quote_ref} | {p.product_name}")
                break


def _patched_build_left_panel(self, parent):
    self._build_product_details_card(parent)
    self._build_product_module_builder_card(parent)
    self._build_product_parts_card(parent)
    self._build_product_document_card(parent)
    self._build_workorder_card(parent)


def _patched_build_product_module_builder_card(self, parent):
    card = ttk.LabelFrame(parent, text="Build Product Scope (Assemblies + Parts)", style="Card.TLabelframe", padding=12)
    card.pack(fill="both", expand=True, pady=6)

    row1 = ttk.Frame(card)
    row1.pack(fill="x", pady=(0, 6))
    ttk.Label(row1, text="Available Assembly").pack(side="left")
    self.available_module_combo = ttk.Combobox(row1, textvariable=self.available_module_var, state="normal", width=38)
    self.available_module_combo.pack(side="left", padx=6, fill="x", expand=True)
    ttk.Button(row1, text="Add Assembly", command=self.add_module_to_product).pack(side="left", padx=2)
    ttk.Button(row1, text="Remove Selected", command=self.remove_selected_module).pack(side="left", padx=2)

    row2 = ttk.Frame(card)
    row2.pack(fill="x", pady=(0, 6))
    ttk.Label(row2, text="Dependency Assembly").pack(side="left")
    self.module_dependency_combo = ttk.Combobox(row2, textvariable=self.module_dependency_var, state="normal", width=38)
    self.module_dependency_combo.pack(side="left", padx=6, fill="x", expand=True)
    ttk.Button(row2, text="Set Dependency", command=self.set_selected_module_dependency).pack(side="left", padx=2)
    ttk.Button(row2, text="Set Qty", command=self.set_selected_module_qty).pack(side="left", padx=2)

    ttk.Label(card, text="Drag to reorder assemblies. Use the Direct Product Parts section below to add pump / sensor / valve / loose stock directly to the product.", style="Sub.TLabel").pack(anchor="w", pady=(0,4))
    lb_wrap = ttk.Frame(card)
    lb_wrap.pack(fill="both", expand=True)
    self.module_order_listbox = DraggableListbox(lb_wrap, selectmode=tk.SINGLE, font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL), activestyle="none")
    self.module_order_listbox.pack(side="left", fill="both", expand=True)
    self.module_order_listbox.bind("<Double-1>", lambda e: self.open_module_from_list())
    sb = ttk.Scrollbar(lb_wrap, orient="vertical", command=self.module_order_listbox.yview)
    self.module_order_listbox.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    btn_row = ttk.Frame(card)
    btn_row.pack(fill="x", pady=(6, 0))
    ttk.Button(btn_row, text="Save Order", command=self.save_module_order).pack(side="left", padx=2)
    ttk.Button(btn_row, text="Move Up", command=lambda: self.move_selected_module(-1)).pack(side="left", padx=2)
    ttk.Button(btn_row, text="Move Down", command=lambda: self.move_selected_module(1)).pack(side="left", padx=2)

    _erp_bind_filterable_combobox(self.available_module_combo, lambda: getattr(self.available_module_combo, '_all_values', self.available_module_combo.cget('values')))
    _erp_bind_filterable_combobox(self.module_dependency_combo, lambda: getattr(self.module_dependency_combo, '_all_values', self.module_dependency_combo.cget('values')))


def _patched_build_product_parts_card(self, parent):
    card = ttk.LabelFrame(parent, text="Direct Product Parts", style="Card.TLabelframe", padding=12)
    card.pack(fill="x", pady=6)
    fields = [
        ("Part Name", self.part_name_var),
        ("Part Number", self.part_number_var),
        ("Qty", self.part_qty_var),
        ("SOH Qty", self.part_soh_var),
        ("Supplier", self.part_supplier_var),
        ("Lead Time Days", self.part_lead_time_var),
        ("Notes", self.part_notes_var),
    ]
    for r, (label, var) in enumerate(fields):
        ttk.Label(card, text=label).grid(row=r, column=0, sticky='w', pady=4)
        ttk.Entry(card, textvariable=var).grid(row=r, column=1, sticky='ew', padx=6)
    btn = ttk.Frame(card)
    btn.grid(row=len(fields), column=0, columnspan=2, sticky='ew', pady=(10,0))
    ttk.Button(btn, text='Add Direct Part', command=self.add_part_to_product).pack(side='left', fill='x', expand=True, padx=2)
    ttk.Button(btn, text='Remove Selected Part', command=self.remove_selected_product_part).pack(side='left', fill='x', expand=True, padx=2)
    card.columnconfigure(1, weight=1)


def _patched_build_right_panel(self, parent):
    tabs = ttk.Notebook(parent)
    tabs.pack(fill="both", expand=True)
    self.modules_tab = ttk.Frame(tabs)
    self.parts_tab = ttk.Frame(tabs)
    self.all_parts_tab = ttk.Frame(tabs)
    self.docs_tab = ttk.Frame(tabs)
    self.workorders_tab = ttk.Frame(tabs)
    self.summary_tab = ttk.Frame(tabs)
    tabs.add(self.modules_tab, text="Assemblies")
    tabs.add(self.parts_tab, text="Direct Parts")
    tabs.add(self.all_parts_tab, text="All Parts")
    tabs.add(self.docs_tab, text="Instruction Manuals / Docs")
    tabs.add(self.workorders_tab, text="Work Orders")
    tabs.add(self.summary_tab, text="Summary")
    self._build_modules_tab(self.modules_tab)
    self._build_product_parts_tab(self.parts_tab)
    self._build_all_product_parts_tab(self.all_parts_tab)
    self._build_docs_tab(self.docs_tab)
    self._build_workorders_tab(self.workorders_tab)
    self._build_summary_tab(self.summary_tab)


def _patched_build_modules_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill="both", expand=True)
    cols = ("Order", "ModuleCode", "ModuleName", "Qty", "Dependency", "Description")
    self.product_modules_tree = ttk.Treeview(wrap, columns=cols, show="headings")
    for col, width in [("Order",70),("ModuleCode",260),("ModuleName",220),("Qty",70),("Dependency",220),("Description",380)]:
        self.product_modules_tree.heading(col, text=col)
        self.product_modules_tree.column(col, width=width, anchor='w')
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.product_modules_tree.xview)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.product_modules_tree.yview)
    self.product_modules_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.product_modules_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


def _patched_build_product_parts_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    cols = ('PartName','PartNumber','Qty','SOH','Supplier','LeadTime','Notes')
    self.product_parts_tree = ttk.Treeview(wrap, columns=cols, show='headings')
    for col, width, title in [
        ('PartName',220,'Part Name'),('PartNumber',140,'Part Number'),('Qty',70,'Qty'),('SOH',70,'SOH'),('Supplier',160,'Supplier'),('LeadTime',90,'Lead Time'),('Notes',260,'Notes')]:
        self.product_parts_tree.heading(col, text=title)
        self.product_parts_tree.column(col, width=width, anchor='w')
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.product_parts_tree.xview)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.product_parts_tree.yview)
    self.product_parts_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.product_parts_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


def _build_all_product_parts_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    cols = ('Source', 'Assembly', 'PartName', 'PartNumber', 'Qty', 'SOH', 'Supplier', 'LeadTime', 'Notes')
    self.all_product_parts_tree = ttk.Treeview(wrap, columns=cols, show='headings')
    for col, width, title in [
        ('Source', 110, 'Source'),
        ('Assembly', 220, 'Assembly'),
        ('PartName', 220, 'Part Name'),
        ('PartNumber', 140, 'Part Number'),
        ('Qty', 80, 'Qty'),
        ('SOH', 70, 'SOH'),
        ('Supplier', 160, 'Supplier'),
        ('LeadTime', 90, 'Lead Time'),
        ('Notes', 260, 'Notes'),
    ]:
        self.all_product_parts_tree.heading(col, text=title)
        self.all_product_parts_tree.column(col, width=width, anchor='w')
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.all_product_parts_tree.xview)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.all_product_parts_tree.yview)
    self.all_product_parts_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.all_product_parts_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


def _patched_refresh_available_modules(self):
    modules = self.app.services.modules.list_modules()
    available = [m.module_code for m in modules]
    self.available_module_combo['values'] = available
    self.available_module_combo._all_values = available


def _patched_refresh_dependency_modules(self, links):
    module_codes = [l.module_code for l in links]
    self.module_dependency_combo['values'] = module_codes
    self.module_dependency_combo._all_values = module_codes


def _patched_clear_all_views(self):
    self.product_code_var.set('')
    self.quote_ref_var.set('')
    self.product_name_var.set('')
    self.description_var.set('')
    self.revision_var.set('R0')
    self.status_var_local.set('Draft')
    treeview_clear(self.product_modules_tree)
    if hasattr(self, 'product_parts_tree'):
        treeview_clear(self.product_parts_tree)
    if hasattr(self, 'all_product_parts_tree'):
        treeview_clear(self.all_product_parts_tree)
    treeview_clear(self.product_document_tree)
    treeview_clear(self.workorder_tree)
    set_text_readonly(self.summary_text, '')
    self.module_order_listbox.delete(0, tk.END)
    for var in [self.part_name_var,self.part_number_var,self.part_qty_var,self.part_soh_var,self.part_supplier_var,self.part_lead_time_var,self.part_notes_var]:
        var.set('')
    self.part_qty_var.set('1')
    self.part_soh_var.set('0')
    self.part_lead_time_var.set('0')
    self.clear_workorder_editor()


def _patched_load_product_parts(self, parts):
    if not hasattr(self, 'product_parts_tree'):
        return
    treeview_clear(self.product_parts_tree)
    for p in parts or []:
        self.product_parts_tree.insert('', 'end', values=(p.component_name, p.part_number, p.qty, p.soh_qty, p.preferred_supplier, p.lead_time_days, p.notes), tags=(p.component_id,))


def _load_all_product_parts(self, bundle, direct_parts):
    if not hasattr(self, 'all_product_parts_tree'):
        return
    treeview_clear(self.all_product_parts_tree)

    module_service = getattr(self.app.services, 'modules', None)
    get_module_components = getattr(module_service, 'get_module_components', None)
    if not callable(get_module_components):
        return

    module_name_map = {
        norm_text(getattr(mod, 'module_code', '')): norm_text(getattr(mod, 'module_name', ''))
        for mod in (getattr(bundle, 'modules', []) or [])
    }

    for link in (getattr(bundle, 'module_links', []) or []):
        module_code = norm_text(getattr(link, 'module_code', ''))
        module_name = module_name_map.get(module_code, module_code)
        module_qty = getattr(link, 'module_qty', 1) or 1
        try:
            module_qty = float(module_qty)
        except Exception:
            module_qty = 1.0
        try:
            components = get_module_components(module_code)
        except Exception:
            components = []
        for comp in components or []:
            try:
                qty = float(getattr(comp, 'qty', 0) or 0) * module_qty
                soh_qty = float(getattr(comp, 'soh_qty', 0) or 0)
            except Exception:
                qty = 0.0
                soh_qty = 0.0
            self.all_product_parts_tree.insert(
                '',
                'end',
                values=(
                    'Assembly',
                    f"{module_code} | {module_name}",
                    getattr(comp, 'component_name', ''),
                    getattr(comp, 'part_number', ''),
                    f"{qty:g}",
                    f"{soh_qty:g}",
                    getattr(comp, 'preferred_supplier', ''),
                    getattr(comp, 'lead_time_days', ''),
                    getattr(comp, 'notes', ''),
                ),
                tags=(getattr(comp, 'component_id', ''),),
            )

    for part in direct_parts or []:
        try:
            qty = float(getattr(part, 'qty', 0) or 0)
            soh_qty = float(getattr(part, 'soh_qty', 0) or 0)
        except Exception:
            qty = 0.0
            soh_qty = 0.0
        self.all_product_parts_tree.insert(
            '',
            'end',
            values=(
                'Direct',
                '',
                getattr(part, 'component_name', ''),
                getattr(part, 'part_number', ''),
                f"{qty:g}",
                f"{soh_qty:g}",
                getattr(part, 'preferred_supplier', ''),
                getattr(part, 'lead_time_days', ''),
                getattr(part, 'notes', ''),
            ),
            tags=(getattr(part, 'component_id', ''),),
        )


def _patched_refresh_page(self):
    if not self.require_workbook():
        return
    self.refresh_product_selector()
    product_code = getattr(self.app, 'selected_product_code', '').strip()
    if not product_code:
        self._clear_all_views()
        self.product_summary_label.config(text='No product selected')
        return
    try:
        bundle = self.app.services.products.get_product_bundle(product_code)
        if not bundle.product:
            self._clear_all_views()
            self.product_summary_label.config(text='Product not found')
            return
        self._load_product_into_form(bundle.product)
        self._load_module_links(bundle)
        self._load_product_parts(self.app.services.products.get_product_parts(product_code))
        self._load_product_documents(bundle.product_documents or [])
        self._load_workorders(bundle.workorders or [])
        self._load_summary(bundle)
        self._refresh_available_modules()
        self._refresh_dependency_modules(bundle.module_links or [])
        self.product_summary_label.config(text=f"🧩 {bundle.product.product_name} | Quote: {bundle.product.quote_ref} | Rev: {bundle.product.revision} | Code: {bundle.product.product_code}")
        self.set_status(f"Loaded product: {bundle.product.product_code}")
    except Exception as exc:
        show_error('Product Refresh Error', str(exc))


def _patched_load_summary(self, bundle):
    product = bundle.product
    links = bundle.module_links or []
    docs = bundle.product_documents or []
    workorders = bundle.workorders or []
    tasks_by_module = bundle.tasks_by_module or {}
    parts = self.app.services.products.get_product_parts(product.product_code)
    lines = [
        f"Product Code: {product.product_code}",
        f"Quote Ref: {product.quote_ref}",
        f"Product Name: {product.product_name}",
        f"Description: {product.description}",
        f"Revision: {product.revision}",
        f"Status: {product.status}",
        "",
        f"Assigned Assemblies: {len(links)}",
        f"Direct Product Parts: {len(parts)}",
        f"Product Documents: {len(docs)}",
        f"Work Orders: {len(workorders)}",
        f"Aggregated Product Hours: {float(bundle.total_hours or 0.0):.2f}",
        "",
        "Assembly Breakdown:",
    ]
    if not links:
        lines.append(' - No assemblies assigned.')
    else:
        for link in links:
            module_tasks = tasks_by_module.get(link.module_code, [])
            module_hours = sum(float(t.estimated_hours or 0.0) for t in module_tasks)
            dep_text = f" | Depends on: {link.dependency_module_code}" if norm_text(link.dependency_module_code) else ''
            lines.append(f" {link.module_order:02d}. {link.module_code} | Qty {link.module_qty} | Hours {module_hours * link.module_qty:.2f}{dep_text}")
    lines.append('')
    lines.append('Direct Product Parts:')
    if not parts:
        lines.append(' - No direct parts assigned.')
    else:
        for p in parts:
            lines.append(f" - {p.component_name} | PN {p.part_number or '-'} | Qty {p.qty} | Supplier {p.preferred_supplier or '-'} | Lead {p.lead_time_days}d")
    set_text_readonly(self.summary_text, '\n'.join(lines))


def _patched_add_part_to_product(self):
    if not self.require_workbook():
        return
    product_code = self._selected_product_code()
    if not product_code:
        show_warning('No Product', 'Save or select a product first.')
        return
    try:
        self.app.services.products.add_product_part(
            product_code=product_code,
            component_name=Validators.require_text(self.part_name_var.get(), 'Part name'),
            qty=float(self.part_qty_var.get().strip() or '1'),
            soh_qty=float(self.part_soh_var.get().strip() or '0'),
            preferred_supplier=self.part_supplier_var.get().strip(),
            lead_time_days=int(float(self.part_lead_time_var.get().strip() or '0')),
            part_number=self.part_number_var.get().strip(),
            notes=self.part_notes_var.get().strip(),
        )
        self.refresh_page()
        self.set_status('Direct product part added.')
    except Exception as exc:
        show_error('Add Direct Part Error', str(exc))


def _patched_remove_selected_product_part(self):
    if not self.require_workbook():
        return
    sel = getattr(self, 'product_parts_tree').selection()
    if not sel:
        show_warning('No Selection', 'Select a direct product part first.')
        return
    tags = self.product_parts_tree.item(sel[0], 'tags')
    comp_id = tags[0] if tags else ''
    if not comp_id:
        show_warning('No Selection', 'Could not identify selected part.')
        return
    try:
        self.app.services.modules.delete_component(comp_id)
        self.refresh_page()
        self.set_status('Direct product part removed.')
    except Exception as exc:
        show_error('Remove Direct Part Error', str(exc))


ProductPage._build_ui = _patched_product_build_ui
ProductPage._build_topbar = _patched_product_topbar
ProductPage.refresh_product_selector = _patched_refresh_product_selector
ProductPage._build_left_panel = _patched_build_left_panel
ProductPage._build_product_module_builder_card = _patched_build_product_module_builder_card
ProductPage._build_product_parts_card = _patched_build_product_parts_card
ProductPage._build_right_panel = _patched_build_right_panel
ProductPage._build_modules_tab = _patched_build_modules_tab
ProductPage._build_product_parts_tab = _patched_build_product_parts_tab
ProductPage._build_all_product_parts_tab = _build_all_product_parts_tab
ProductPage._refresh_available_modules = _patched_refresh_available_modules
ProductPage._refresh_dependency_modules = _patched_refresh_dependency_modules
ProductPage._clear_all_views = _patched_clear_all_views
ProductPage._load_product_parts = _patched_load_product_parts
ProductPage._load_all_product_parts = _load_all_product_parts
ProductPage.refresh_page = _patched_refresh_page
ProductPage._load_summary = _patched_load_summary
ProductPage.add_part_to_product = _patched_add_part_to_product
ProductPage.remove_selected_product_part = _patched_remove_selected_product_part


# ============================================================
# V3 PATCH: searchable product code + quote ref + better summary scroll
# ============================================================

def _v3_collect_quote_refs(page):
    refs = []
    try:
        for p in page.app.services.products.list_products():
            if norm_text(getattr(p, 'quote_ref', '')):
                refs.append(norm_text(p.quote_ref))
    except Exception:
        pass
    try:
        for m in page.app.services.modules.list_modules():
            if norm_text(getattr(m, 'quote_ref', '')):
                refs.append(norm_text(m.quote_ref))
    except Exception:
        pass
    try:
        for o in page.app.services.projects.list_projects():
            if norm_text(getattr(o, 'quote_ref', '')):
                refs.append(norm_text(o.quote_ref))
    except Exception:
        pass
    return sorted(set(refs))


def _v3_build_product_details_card(self, parent):
    card = ttk.LabelFrame(parent, text="Product Details", style="Card.TLabelframe", padding=12)
    card.pack(fill="x", pady=6)

    ttk.Label(card, text="Product Code").grid(row=0, column=0, sticky="w", pady=4)
    self.product_code_combo = ttk.Combobox(card, textvariable=self.product_code_var, state="normal")
    self.product_code_combo.grid(row=0, column=1, sticky="ew", padx=6)
    self.product_code_combo.bind("<<ComboboxSelected>>", lambda e: self._open_product_from_details_code())

    ttk.Label(card, text="Quote Ref").grid(row=1, column=0, sticky="w", pady=4)
    self.quote_ref_combo = ttk.Combobox(card, textvariable=self.quote_ref_var, state="normal")
    self.quote_ref_combo.grid(row=1, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Product Name").grid(row=2, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.product_name_var).grid(row=2, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Description").grid(row=3, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.description_var).grid(row=3, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Revision").grid(row=4, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.revision_var).grid(row=4, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Status").grid(row=5, column=0, sticky="w", pady=4)
    ttk.Combobox(card, textvariable=self.status_var_local, values=AppConfig.PRODUCT_STATUSES, state="readonly").grid(row=5, column=1, sticky="ew", padx=6)

    btn_row = ttk.Frame(card)
    btn_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
    ttk.Button(btn_row, text="Save Product", command=self.save_product).pack(side="left", fill="x", expand=True, padx=2)
    ttk.Button(btn_row, text="Delete Product", command=self.delete_product).pack(side="left", fill="x", expand=True, padx=2)

    card.columnconfigure(1, weight=1)
    _erp_bind_filterable_combobox(self.product_code_combo, lambda: getattr(self.product_code_combo, '_all_values', self.product_code_combo.cget('values')))
    _erp_bind_filterable_combobox(self.quote_ref_combo, lambda: getattr(self.quote_ref_combo, '_all_values', self.quote_ref_combo.cget('values')))


def _v3_refresh_product_selector(self):
    _patched_refresh_product_selector(self)
    try:
        records = self.app.services.products.list_products() if self.app.workbook_manager.has_workbook() else []
        code_values = [f"{p.product_code} | {p.product_name}" for p in records]
        if hasattr(self, 'product_code_combo'):
            self.product_code_combo['values'] = code_values
            self.product_code_combo._all_values = code_values
        quote_values = _v3_collect_quote_refs(self)
        if hasattr(self, 'quote_ref_combo'):
            self.quote_ref_combo['values'] = quote_values
            self.quote_ref_combo._all_values = quote_values
    except Exception:
        pass


def _v3_open_product_from_details_code(self):
    text = norm_text(self.product_code_var.get())
    if not text:
        return
    code = text.split('|')[0].strip()
    self.app.set_selected_product(code)
    self.refresh_page()


def _v3_build_summary_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    self.summary_text = make_readonly_text(wrap, height=26)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.summary_text.yview)
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.summary_text.xview)
    self.summary_text.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set, wrap='none')
    self.summary_text.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


ProductPage._build_product_details_card = _v3_build_product_details_card
ProductPage.refresh_product_selector = _v3_refresh_product_selector
ProductPage._open_product_from_details_code = _v3_open_product_from_details_code
ProductPage._build_summary_tab = _v3_build_summary_tab



# ============================================================
# Final patch: full-column scroll + searchable direct part dropdowns
# ============================================================

def _final_bind_mousewheel_recursive(self, widget, target_canvas):
    def _on_mousewheel(event):
        try:
            if getattr(event, "num", None) == 4:
                target_canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                target_canvas.yview_scroll(1, "units")
            else:
                delta = getattr(event, "delta", 0)
                if delta:
                    target_canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        except Exception:
            pass
        return "break"

    try:
        widget.bind("<MouseWheel>", _on_mousewheel, add="+")
        widget.bind("<Button-4>", _on_mousewheel, add="+")
        widget.bind("<Button-5>", _on_mousewheel, add="+")
    except Exception:
        pass

    try:
        for child in widget.winfo_children():
            _final_bind_mousewheel_recursive(self, child, target_canvas)
    except Exception:
        pass


def _final_build_ui(self):
    wrapper = ttk.Frame(self, padding=14)
    wrapper.pack(fill="both", expand=True)

    self._build_topbar(wrapper)

    self.product_summary_label = ttk.Label(
        wrapper,
        text="No product selected",
        style="Title.TLabel"
    )
    self.product_summary_label.pack(anchor="w", pady=(0, 10))

    paned = ttk.Panedwindow(wrapper, orient="horizontal")
    paned.pack(fill="both", expand=True)

    left_host = ttk.Frame(paned)
    right = ttk.Frame(paned, padding=6)

    paned.add(left_host, weight=2)
    paned.add(right, weight=3)

    self.left_canvas = tk.Canvas(left_host, highlightthickness=0, bg=AppConfig.COLOR_BG)
    self.left_scrollbar = ttk.Scrollbar(left_host, orient="vertical", command=self.left_canvas.yview)
    self.left_frame = ttk.Frame(self.left_canvas)

    self.left_frame.bind(
        "<Configure>",
        lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
    )

    self.canvas_window = self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")

    def _resize(event):
        try:
            self.left_canvas.itemconfig(self.canvas_window, width=event.width)
        except Exception:
            pass

    self.left_canvas.bind("<Configure>", _resize)
    self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)

    self.left_canvas.pack(side="left", fill="both", expand=True)
    self.left_scrollbar.pack(side="right", fill="y")

    self._build_left_panel(self.left_frame)
    self._build_right_panel(right)

    _final_bind_mousewheel_recursive(self, self.left_frame, self.left_canvas)
    _final_bind_mousewheel_recursive(self, self.left_canvas, self.left_canvas)


def _final_get_all_available_parts(self):
    parts = []
    seen = set()
    try:
        repo = getattr(self.app, "repo", None)
        if repo is not None:
            rows = []
            if hasattr(repo, "read_sheet_as_dicts"):
                rows = repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
            elif hasattr(repo, "list_dicts"):
                rows = repo.list_dicts(AppConfig.SHEET_COMPONENTS)
            for row in rows or []:
                part_no = str(row.get("PartNumber", "") or row.get("part_number", "") or "").strip()
                part_name = str(row.get("ComponentName", "") or row.get("component_name", "") or "").strip()
                supplier = str(row.get("PreferredSupplier", "") or row.get("preferred_supplier", "") or "").strip()
                soh = str(row.get("SOHQty", "") or row.get("soh_qty", "") or row.get("StockOnHand", "") or "").strip()
                lead = str(row.get("LeadTimeDays", "") or row.get("lead_time_days", "") or "").strip()
                notes = str(row.get("Notes", "") or row.get("notes", "") or "").strip()
                key = (part_no.lower(), part_name.lower())
                if key in seen:
                    continue
                seen.add(key)
                if part_no or part_name:
                    parts.append({
                        "part_number": part_no,
                        "part_name": part_name,
                        "supplier": supplier,
                        "soh": soh,
                        "lead": lead,
                        "notes": notes,
                    })
    except Exception:
        pass
    return parts


def _final_filter_combobox_values(self, combo, full_values, typed_text):
    typed = str(typed_text or "").strip().lower()
    if typed:
        vals = [v for v in full_values if typed in str(v).lower()]
    else:
        vals = list(full_values)
    combo["values"] = vals


def _final_refresh_part_dropdowns(self):
    all_parts = _final_get_all_available_parts(self)
    self._all_available_parts = all_parts

    name_values = sorted({p["part_name"] for p in all_parts if p["part_name"]})
    number_values = sorted({p["part_number"] for p in all_parts if p["part_number"]})

    if hasattr(self, "part_name_combo"):
        self.part_name_combo["values"] = name_values
        self.part_name_combo._all_values = name_values
    if hasattr(self, "part_number_combo"):
        self.part_number_combo["values"] = number_values
        self.part_number_combo._all_values = number_values


def _final_on_part_name_selected(self, event=None):
    name = self.part_name_var.get().strip().lower()
    for p in getattr(self, "_all_available_parts", []):
        if p["part_name"].strip().lower() == name:
            self.part_number_var.set(p["part_number"])
            if not self.part_supplier_var.get().strip():
                self.part_supplier_var.set(p["supplier"])
            if not self.part_soh_var.get().strip():
                self.part_soh_var.set(p["soh"])
            if not self.part_lead_time_var.get().strip():
                self.part_lead_time_var.set(p["lead"])
            if not self.part_notes_var.get().strip():
                self.part_notes_var.set(p["notes"])
            break


def _final_on_part_number_selected(self, event=None):
    number = self.part_number_var.get().strip().lower()
    for p in getattr(self, "_all_available_parts", []):
        if p["part_number"].strip().lower() == number:
            self.part_name_var.set(p["part_name"])
            if not self.part_supplier_var.get().strip():
                self.part_supplier_var.set(p["supplier"])
            if not self.part_soh_var.get().strip():
                self.part_soh_var.set(p["soh"])
            if not self.part_lead_time_var.get().strip():
                self.part_lead_time_var.set(p["lead"])
            if not self.part_notes_var.get().strip():
                self.part_notes_var.set(p["notes"])
            break


def _final_part_name_keyrelease(self, event=None):
    vals = getattr(self.part_name_combo, "_all_values", list(self.part_name_combo.cget("values")))
    _final_filter_combobox_values(self, self.part_name_combo, vals, self.part_name_var.get())


def _final_part_number_keyrelease(self, event=None):
    vals = getattr(self.part_number_combo, "_all_values", list(self.part_number_combo.cget("values")))
    _final_filter_combobox_values(self, self.part_number_combo, vals, self.part_number_var.get())


def _final_build_product_parts_card(self, parent):
    if not hasattr(self, 'part_name_var'):
        self.part_name_var = tk.StringVar()
        self.part_number_var = tk.StringVar()
        self.part_qty_var = tk.StringVar(value='1')
        self.part_soh_var = tk.StringVar()
        self.part_supplier_var = tk.StringVar()
        self.part_lead_time_var = tk.StringVar(value='0')
        self.part_notes_var = tk.StringVar()

    card = ttk.LabelFrame(parent, text="Direct Product Parts", style="Card.TLabelframe", padding=12)
    card.pack(fill="x", pady=6)

    ttk.Label(card, text="Part Name").grid(row=0, column=0, sticky='w', pady=4)
    self.part_name_combo = ttk.Combobox(card, textvariable=self.part_name_var, state="normal")
    self.part_name_combo.grid(row=0, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="Part Number").grid(row=1, column=0, sticky='w', pady=4)
    self.part_number_combo = ttk.Combobox(card, textvariable=self.part_number_var, state="normal")
    self.part_number_combo.grid(row=1, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="Qty").grid(row=2, column=0, sticky='w', pady=4)
    ttk.Entry(card, textvariable=self.part_qty_var).grid(row=2, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="SOH Qty").grid(row=3, column=0, sticky='w', pady=4)
    ttk.Entry(card, textvariable=self.part_soh_var).grid(row=3, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="Supplier").grid(row=4, column=0, sticky='w', pady=4)
    ttk.Entry(card, textvariable=self.part_supplier_var).grid(row=4, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="Lead Time Days").grid(row=5, column=0, sticky='w', pady=4)
    ttk.Entry(card, textvariable=self.part_lead_time_var).grid(row=5, column=1, sticky='ew', padx=6)

    ttk.Label(card, text="Notes").grid(row=6, column=0, sticky='w', pady=4)
    ttk.Entry(card, textvariable=self.part_notes_var).grid(row=6, column=1, sticky='ew', padx=6)

    btn = ttk.Frame(card)
    btn.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(10,0))
    ttk.Button(btn, text='Add Direct Part', command=self.add_part_to_product).pack(side='left', fill='x', expand=True, padx=2)
    ttk.Button(btn, text='Remove Selected Part', command=self.remove_selected_product_part).pack(side='left', fill='x', expand=True, padx=2)

    card.columnconfigure(1, weight=1)

    _final_refresh_part_dropdowns(self)
    self.part_name_combo.bind("<<ComboboxSelected>>", lambda e: _final_on_part_name_selected(self, e))
    self.part_number_combo.bind("<<ComboboxSelected>>", lambda e: _final_on_part_number_selected(self, e))
    self.part_name_combo.bind("<KeyRelease>", lambda e: _final_part_name_keyrelease(self, e))
    self.part_number_combo.bind("<KeyRelease>", lambda e: _final_part_number_keyrelease(self, e))


def _final_refresh_page(self):
    if not self.require_workbook():
        return
    try:
        self.refresh_product_selector()
        _final_refresh_part_dropdowns(self)
    except Exception:
        pass

    product_code = getattr(self.app, "selected_product_code", "").strip()
    if not product_code:
        self._clear_all_views()
        self.product_summary_label.config(text="No product selected")
        return

    try:
        bundle = self.app.services.products.get_product_bundle(product_code)
        if not bundle.product:
            self._clear_all_views()
            self.product_summary_label.config(text="Product not found")
            return

        self._load_product_into_form(bundle.product)
        self._load_module_links(bundle)
        try:
            direct_parts = self.app.services.products.get_product_parts(product_code)
            self._load_product_parts(direct_parts)
            self._load_all_product_parts(bundle, direct_parts)
        except Exception:
            pass
        self._load_product_documents(bundle.product_documents or [])
        self._load_workorders(bundle.workorders or [])
        try:
            self._load_summary(bundle)
        except Exception:
            try:
                product = bundle.product
                parts_count = len(direct_parts) if 'direct_parts' in locals() else 0
                set_text_readonly(
                    self.summary_text,
                    "\n".join([
                        f"Product Code: {getattr(product, 'product_code', '')}",
                        f"Quote Ref: {getattr(product, 'quote_ref', '')}",
                        f"Product Name: {getattr(product, 'product_name', '')}",
                        f"Description: {getattr(product, 'description', '')}",
                        f"Revision: {getattr(product, 'revision', '')}",
                        f"Status: {getattr(product, 'status', '')}",
                        "",
                        f"Assigned Assemblies: {len(getattr(bundle, 'module_links', []) or [])}",
                        f"Direct Product Parts: {parts_count}",
                        f"Product Documents: {len(getattr(bundle, 'product_documents', []) or [])}",
                        f"Work Orders: {len(getattr(bundle, 'workorders', []) or [])}",
                        f"Aggregated Product Hours: {float(getattr(bundle, 'total_hours', 0.0) or 0.0):.2f}",
                    ])
                )
            except Exception:
                pass
        self._refresh_available_modules()
        self._refresh_dependency_modules(bundle.module_links or [])
        try:
            _final_refresh_part_dropdowns(self)
        except Exception:
            pass
        try:
            if not get_text_value(self.summary_text):
                product = bundle.product
                set_text_readonly(
                    self.summary_text,
                    "\n".join([
                        f"Product Code: {getattr(product, 'product_code', '')}",
                        f"Quote Ref: {getattr(product, 'quote_ref', '')}",
                        f"Product Name: {getattr(product, 'product_name', '')}",
                        f"Status: {getattr(product, 'status', '')}",
                    ])
                )
        except Exception:
            pass

        self.product_summary_label.config(
            text=f"🧩 {bundle.product.product_name} | Quote: {bundle.product.quote_ref} | Rev: {bundle.product.revision} | Code: {bundle.product.product_code}"
        )
        self.set_status(f"Loaded product: {bundle.product.product_code}")
    except Exception as exc:
        show_error('Product Refresh Error', str(exc))


ProductPage._build_ui = _final_build_ui
ProductPage._build_product_parts_card = _final_build_product_parts_card
ProductPage.refresh_page = _final_refresh_page
ProductPage._bind_mousewheel_recursive = _final_bind_mousewheel_recursive
ProductPage._refresh_part_dropdowns = _final_refresh_part_dropdowns
ProductPage._on_part_name_selected = _final_on_part_name_selected
ProductPage._on_part_number_selected = _final_on_part_number_selected
