# ============================================================
# ui_home.py
# Home / dashboard page for Liquimech ERP Desktop App
# Cleaned version with fixed manager navigation
# ============================================================

from __future__ import annotations

from pathlib import Path

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

import tkinter as tk
from tkinter import ttk, filedialog

from app_config import AppConfig
from ui_common import (
    BasePage,
    NavStrip,
    treeview_clear,
    set_combobox_values,
    show_warning,
    show_error,
    show_info,
)

BASE_DIR = Path(__file__).resolve().parent


class HomePage(BasePage):
    PAGE_NAME = "home"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        self.selected_module_var = tk.StringVar()
        self.selected_product_var = tk.StringVar()
        self.selected_project_var = tk.StringVar()

        self.module_search_var = tk.StringVar()
        self.product_search_var = tk.StringVar()
        self.project_search_var = tk.StringVar()

        self.create_quote_ref_var = tk.StringVar()
        self.create_name_var = tk.StringVar()
        self.create_desc_var = tk.StringVar()
        self.create_type_var = tk.StringVar(value="Assembly")

        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=16)
        wrapper.pack(fill="both", expand=True)

        self._build_header(wrapper)
        self._build_stats_row(wrapper)
        self._build_quick_actions(wrapper)
        self._build_navigation(wrapper)
        self._build_recent_tabs(wrapper)

    def _build_header(self, parent):
        header = ttk.Frame(parent)
        header.pack(fill="x", pady=(0, 10))

        left = ttk.Frame(header)
        left.pack(side="left", fill="x", expand=True)

        brand_block = ttk.Frame(left)
        brand_block.pack(anchor="w")

        self.logo_label = ttk.Label(brand_block, background=AppConfig.COLOR_BG)
        self.logo_label.pack(anchor="w", pady=(0, 4))

        self._load_company_logo()

        ttk.Label(
            left,
            text="Project Delivery Suite",
            style="Title.TLabel"
        ).pack(anchor="w", pady=(4, 0))

        ttk.Label(
            left,
            text="Parts, assemblies, products, live orders, and scheduling in one desktop system.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        right = ttk.Frame(header)
        right.pack(side="right")

        self.workbook_label = ttk.Label(
            right,
            text="No workbook selected",
            style="Sub.TLabel"
        )
        self.workbook_label.pack(anchor="e", pady=(0, 6))

        btn_row = ttk.Frame(right)
        btn_row.pack(anchor="e")

        ttk.Button(btn_row, text="Create Workbook", command=self.create_workbook).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Open Workbook", command=self.open_workbook).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Refresh Dashboard", command=self.refresh_page).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Open Job Cards Board", command=lambda: self.show_page("jobcards")).pack(side="left", padx=4)

    def _load_company_logo(self):
        if Image is None or ImageTk is None:
            self.logo_label.config(text="")
            return

        possible_paths = [
            BASE_DIR / "logo.png",
            BASE_DIR / "assets" / "logo.png",
            BASE_DIR / "assets" / "liquimech_logo.png",
            BASE_DIR / "assets" / "liquimech_icon.png",
        ]

        logo_path = None
        for p in possible_paths:
            if p.exists():
                logo_path = p
                break

        if not logo_path:
            self.logo_label.config(text="")
            return

        try:
            img = Image.open(logo_path)
            img.thumbnail((150, 80), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(img)
            self.logo_label.config(image=self.logo_photo)
        except Exception:
            self.logo_label.config(text="")

    def _build_stats_row(self, parent):
        stats = ttk.Frame(parent)
        stats.pack(fill="x", pady=(0, 20))

        self.parts_card = ttk.Button(
            stats,
            text="Parts\n0\nStock, supplier, lead time",
            command=self.open_parts_manager,
            style="Metric.TButton"
        )
        self.parts_card.pack(side="left", fill="x", expand=True, padx=12, ipady=10)

        self.modules_card = ttk.Button(
            stats,
            text="Assemblies\n0\nBuild definitions",
            command=self.go_to_assembly_manager,
            style="Metric.TButton"
        )
        self.modules_card.pack(side="left", fill="x", expand=True, padx=12, ipady=10)

        self.products_card = ttk.Button(
            stats,
            text="Products\n0\nConfigured products",
            command=self.go_to_products,
            style="Metric.TButton"
        )
        self.products_card.pack(side="left", fill="x", expand=True, padx=12, ipady=10)

        self.projects_card = ttk.Button(
            stats,
            text="Live Orders\n0\nExecution jobs",
            command=self.go_to_projects,
            style="Metric.TButton"
        )
        self.projects_card.pack(side="left", fill="x", expand=True, padx=12, ipady=10)

    def _build_quick_actions(self, parent):
        card = ttk.LabelFrame(parent, text="Quick Create / Open", style="Card.TLabelframe", padding=14)
        card.pack(fill="x", pady=(0, 12))

        top = ttk.Frame(card)
        top.pack(fill="x", pady=(0, 8))

        module_box = ttk.Frame(top)
        module_box.pack(side="left", fill="x", expand=True, padx=4)

        ttk.Label(module_box, text="Select Assembly").pack(anchor="w")
        self.module_combo = ttk.Combobox(module_box, textvariable=self.selected_module_var, state="readonly")
        self.module_combo.pack(fill="x", pady=(4, 0))
        self.module_combo.bind("<Double-Button-1>", lambda e: self.open_selected_module())

        mod_btns = ttk.Frame(module_box)
        mod_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(mod_btns, text="Open Assembly", command=self.open_selected_module).pack(side="left", padx=2)
        ttk.Button(mod_btns, text="Go to Assembly Manager", command=self.go_to_assembly_manager).pack(side="left", padx=2)

        product_box = ttk.Frame(top)
        product_box.pack(side="left", fill="x", expand=True, padx=4)

        ttk.Label(product_box, text="Select Product").pack(anchor="w")
        self.product_combo = ttk.Combobox(product_box, textvariable=self.selected_product_var, state="readonly")
        self.product_combo.pack(fill="x", pady=(4, 0))
        self.product_combo.bind("<Double-Button-1>", lambda e: self.open_selected_product())

        prod_btns = ttk.Frame(product_box)
        prod_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(prod_btns, text="Open Product", command=self.open_selected_product).pack(side="left", padx=2)
        ttk.Button(prod_btns, text="Go to Products", command=self.go_to_products).pack(side="left", padx=2)

        project_box = ttk.Frame(top)
        project_box.pack(side="left", fill="x", expand=True, padx=4)

        ttk.Label(project_box, text="Select Live Order").pack(anchor="w")
        self.project_combo = ttk.Combobox(project_box, textvariable=self.selected_project_var, state="readonly")
        self.project_combo.pack(fill="x", pady=(4, 0))
        self.project_combo.bind("<Double-Button-1>", lambda e: self.open_selected_project())

        prj_btns = ttk.Frame(project_box)
        prj_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(prj_btns, text="Open Live Order", command=self.open_selected_project).pack(side="left", padx=2)
        ttk.Button(prj_btns, text="Go to Live Orders", command=self.go_to_projects).pack(side="left", padx=2)

        create_card = ttk.LabelFrame(card, text="Quick Create", style="Card.TLabelframe", padding=12)
        create_card.pack(fill="x", pady=(8, 0))

        ttk.Label(create_card, text="Type").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.create_type_combo = ttk.Combobox(
            create_card,
            textvariable=self.create_type_var,
            values=["Assembly", "Product", "Order"],
            state="readonly",
            width=14,
        )
        self.create_type_combo.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(create_card, text="Quote Ref").grid(row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(create_card, textvariable=self.create_quote_ref_var).grid(row=0, column=3, sticky="ew", padx=4, pady=4)

        ttk.Label(create_card, text="Name").grid(row=0, column=4, sticky="w", padx=4, pady=4)
        ttk.Entry(create_card, textvariable=self.create_name_var).grid(row=0, column=5, sticky="ew", padx=4, pady=4)

        ttk.Label(create_card, text="Description").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(create_card, textvariable=self.create_desc_var).grid(row=1, column=1, columnspan=5, sticky="ew", padx=4, pady=4)

        ttk.Button(create_card, text="Create Now", command=self.quick_create).grid(row=0, column=6, rowspan=2, sticky="ns", padx=8, pady=4)

        create_card.columnconfigure(3, weight=1)
        create_card.columnconfigure(5, weight=1)

    def _build_navigation(self, parent):
        nav_card = ttk.LabelFrame(parent, text="Navigation", style="Card.TLabelframe", padding=12)
        nav_card.pack(fill="x", pady=(0, 12))

        NavStrip(
            nav_card,
            buttons=[
                ("Assembly Manager", self.go_to_assembly_manager),
                ("Products", self.go_to_products),
                ("Live Orders", self.go_to_projects),
                ("Scheduling", lambda: self.show_page("scheduler")),
                ("Dependencies", lambda: self.show_page("dependencies")),
                ("Job Cards", lambda: self.show_page("jobcards")),
                ("Completed Jobs", lambda: self.show_page("completed_jobs")),
            ]
        ).pack(anchor="w")

    def _build_recent_tabs(self, parent):
        tabs = ttk.Notebook(parent)
        tabs.pack(fill="both", expand=True)

        self.modules_tab = ttk.Frame(tabs)
        self.products_tab = ttk.Frame(tabs)
        self.projects_tab = ttk.Frame(tabs)

        tabs.add(self.modules_tab, text="Recent Assemblies")
        tabs.add(self.products_tab, text="Recent Products")
        tabs.add(self.projects_tab, text="Recent Live Orders")

        self._build_recent_modules_tab(self.modules_tab)
        self._build_recent_products_tab(self.products_tab)
        self._build_recent_projects_tab(self.projects_tab)

    def _build_recent_modules_tab(self, parent):
        search_row = ttk.Frame(parent, padding=8)
        search_row.pack(fill="x")

        ttk.Label(search_row, text="Search").pack(side="left")
        ent = ttk.Entry(search_row, textvariable=self.module_search_var, width=40)
        ent.pack(side="left", padx=8)
        ent.bind("<KeyRelease>", lambda e: self.refresh_recent_modules())

        ttk.Button(search_row, text="Open Selected", command=self.open_recent_module).pack(side="left", padx=4)

        frame = ttk.Frame(parent, padding=(8, 0, 8, 8))
        frame.pack(fill="both", expand=True)

        cols = ("AssemblyCode", "QuoteRef", "AssemblyName", "Status", "Hours", "UpdatedOn")
        self.recent_modules_tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        for col, width in [
            ("AssemblyCode", 280),
            ("QuoteRef", 130),
            ("AssemblyName", 220),
            ("Status", 120),
            ("Hours", 90),
            ("UpdatedOn", 180),
        ]:
            self.recent_modules_tree.heading(col, text=col)
            self.recent_modules_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.recent_modules_tree.yview)
        self.recent_modules_tree.configure(yscrollcommand=sb.set)

        self.recent_modules_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.recent_modules_tree.bind("<Double-1>", lambda e: self.open_recent_module())

    def _build_recent_products_tab(self, parent):
        search_row = ttk.Frame(parent, padding=8)
        search_row.pack(fill="x")

        ttk.Label(search_row, text="Search").pack(side="left")
        ent = ttk.Entry(search_row, textvariable=self.product_search_var, width=40)
        ent.pack(side="left", padx=8)
        ent.bind("<KeyRelease>", lambda e: self.refresh_recent_products())

        ttk.Button(search_row, text="Open Selected", command=self.open_recent_product).pack(side="left", padx=4)

        frame = ttk.Frame(parent, padding=(8, 0, 8, 8))
        frame.pack(fill="both", expand=True)

        cols = ("ProductCode", "QuoteRef", "ProductName", "Revision", "Status", "UpdatedOn")
        self.recent_products_tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        for col, width in [
            ("ProductCode", 320),
            ("QuoteRef", 130),
            ("ProductName", 250),
            ("Revision", 90),
            ("Status", 120),
            ("UpdatedOn", 180),
        ]:
            self.recent_products_tree.heading(col, text=col)
            self.recent_products_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.recent_products_tree.yview)
        self.recent_products_tree.configure(yscrollcommand=sb.set)

        self.recent_products_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.recent_products_tree.bind("<Double-1>", lambda e: self.open_recent_product())

    def _build_recent_projects_tab(self, parent):
        search_row = ttk.Frame(parent, padding=8)
        search_row.pack(fill="x")

        ttk.Label(search_row, text="Search").pack(side="left")
        ent = ttk.Entry(search_row, textvariable=self.project_search_var, width=40)
        ent.pack(side="left", padx=8)
        ent.bind("<KeyRelease>", lambda e: self.refresh_recent_projects())

        ttk.Button(search_row, text="Open Selected", command=self.open_recent_project).pack(side="left", padx=4)

        frame = ttk.Frame(parent, padding=(8, 0, 8, 8))
        frame.pack(fill="both", expand=True)

        cols = ("ProjectCode", "QuoteRef", "ProjectName", "ClientName", "Status", "DueDate", "UpdatedOn")
        self.recent_projects_tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        for col, width in [
            ("ProjectCode", 280),
            ("QuoteRef", 120),
            ("ProjectName", 220),
            ("ClientName", 180),
            ("Status", 120),
            ("DueDate", 110),
            ("UpdatedOn", 170),
        ]:
            self.recent_projects_tree.heading(col, text=col)
            self.recent_projects_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.recent_projects_tree.yview)
        self.recent_projects_tree.configure(yscrollcommand=sb.set)

        self.recent_projects_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.recent_projects_tree.bind("<Double-1>", lambda e: self.open_recent_project())

    def create_workbook(self):
        AppConfig.ensure_directories()
        path = filedialog.asksaveasfilename(
            title="Create ERP Workbook",
            defaultextension=".xlsx",
            initialdir=str(AppConfig.DATA_DIR),
            initialfile=AppConfig.DEFAULT_WORKBOOK_NAME,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return

        try:
            self.app.workbook_manager.create_workbook(path)
            self.workbook_label.config(text=path)
            self.set_status("Workbook created successfully.")
            self.refresh_page()
            show_info("Workbook Created", f"Workbook created:\n{path}")
        except Exception as exc:
            show_error("Create Workbook Error", str(exc))

    def open_workbook(self):
        AppConfig.ensure_directories()
        path = filedialog.askopenfilename(
            title="Open ERP Workbook",
            initialdir=str(AppConfig.DATA_DIR),
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return

        try:
            self.app.workbook_manager.open_workbook(path)
            self.workbook_label.config(text=path)
            self.set_status("Workbook loaded.")
            self.refresh_page()
        except Exception as exc:
            show_error("Open Workbook Error", str(exc))

    def quick_create(self):
        if not self.require_workbook():
            return

        obj_type = self.create_type_var.get().strip()
        quote_ref = self.create_quote_ref_var.get().strip()
        name = self.create_name_var.get().strip()
        desc = self.create_desc_var.get().strip()

        if not name:
            show_warning("Missing Data", f"{obj_type} name is required.")
            return

        try:
            if obj_type == "Assembly":
                code = self.app.services.modules.create_or_update_module(
                    quote_ref=quote_ref,
                    module_name=name,
                    description=desc,
                )
                self.app.set_selected_module(code)
                self.set_status(f"Assembly created: {code}")
                self.refresh_page()
                self.show_page("modules")

            elif obj_type == "Product":
                code = self.app.services.products.create_or_update_product(
                    quote_ref=quote_ref,
                    product_name=name,
                    description=desc,
                )
                self.app.set_selected_product(code)
                self.set_status(f"Product created: {code}")
                self.refresh_page()
                self.show_page("products")

            elif obj_type == "Order":
                code = self.app.services.projects.create_or_update_project(
                    quote_ref=quote_ref,
                    project_name=name,
                    description=desc,
                )
                self.app.set_selected_project(code)
                self.set_status(f"Live order created: {code}")
                self.refresh_page()
                self.show_page("projects")

            else:
                show_warning("Unknown Type", f"Unsupported create type: {obj_type}")
                return

            self.create_quote_ref_var.set("")
            self.create_name_var.set("")
            self.create_desc_var.set("")

        except Exception as exc:
            show_error("Quick Create Error", str(exc))

    def open_parts_manager(self):
        # No dedicated ui_parts.py was uploaded in your current project.
        # Route to modules page for now.
        self.show_page("modules")

    def go_to_assembly_manager(self):
        self.selected_module_var.set("")
        try:
            self.app.set_selected_module("")
        except Exception:
            pass
        self.show_page("modules")

    def go_to_products(self):
        self.selected_product_var.set("")
        try:
            self.app.set_selected_product("")
        except Exception:
            pass
        self.show_page("products")

    def go_to_projects(self):
        self.selected_project_var.set("")
        try:
            self.app.set_selected_project("")
        except Exception:
            pass
        self.show_page("projects")

    def open_selected_module(self):
        code = self.selected_module_var.get().strip()
        if not code:
            show_warning("No Selection", "Select an assembly first.")
            return
        self.app.set_selected_module(code)
        self.show_page("modules")

    def open_selected_product(self):
        code = self.selected_product_var.get().strip()
        if not code:
            show_warning("No Selection", "Select a product first.")
            return
        self.app.set_selected_product(code)
        self.show_page("products")

    def open_selected_project(self):
        code = self.selected_project_var.get().strip()
        if not code:
            show_warning("No Selection", "Select a live order first.")
            return
        self.app.set_selected_project(code)
        self.show_page("projects")

    def _get_selected_tree_value(self, tree: ttk.Treeview, index: int = 0) -> str:
        sel = tree.selection()
        if not sel:
            return ""
        values = tree.item(sel[0], "values")
        if not values or index >= len(values):
            return ""
        return str(values[index]).strip()

    def open_recent_module(self):
        code = self._get_selected_tree_value(self.recent_modules_tree, 0)
        if not code:
            show_warning("No Selection", "Select an assembly first.")
            return
        self.app.set_selected_module(code)
        self.show_page("modules")

    def open_recent_product(self):
        code = self._get_selected_tree_value(self.recent_products_tree, 0)
        if not code:
            show_warning("No Selection", "Select a product first.")
            return
        self.app.set_selected_product(code)
        self.show_page("products")

    def open_recent_project(self):
        code = self._get_selected_tree_value(self.recent_projects_tree, 0)
        if not code:
            show_warning("No Selection", "Select a live order first.")
            return
        self.app.set_selected_project(code)
        self.show_page("projects")

    def refresh_page(self):
        path = self.app.workbook_manager.workbook_path
        self.workbook_label.config(text=path or "No workbook selected")

        if not self.app.workbook_manager.has_workbook():
            self._reset_dashboard()
            return

        try:
            modules = self.app.services.modules.list_modules()
            products = self.app.services.products.list_products()
            projects = self.app.services.projects.list_projects()
            self.refresh_counts(modules=modules, products=products, projects=projects)
            self.refresh_combos(modules=modules, products=products, projects=projects)
            self.refresh_recent_modules(modules=modules)
            self.refresh_recent_products(products=products)
            self.refresh_recent_projects(projects=projects)
        except Exception as exc:
            show_error("Dashboard Refresh Error", str(exc))

    def _reset_dashboard(self):
        self.parts_card.config(text="Parts\n0\nStock, supplier, lead time")
        self.modules_card.config(text="Assemblies\n0\nBuild definitions")
        self.products_card.config(text="Products\n0\nConfigured products")
        self.projects_card.config(text="Live Orders\n0\nExecution jobs")

        set_combobox_values(self.module_combo, [], keep_current=False)
        set_combobox_values(self.product_combo, [], keep_current=False)
        set_combobox_values(self.project_combo, [], keep_current=False)

        treeview_clear(self.recent_modules_tree)
        treeview_clear(self.recent_products_tree)
        treeview_clear(self.recent_projects_tree)

    def refresh_counts(self, modules=None, products=None, projects=None):
        parts_count = 0
        try:
            if hasattr(self.app.services, "parts") and hasattr(self.app.services.parts, "list_parts"):
                parts_count = len(self.app.services.parts.list_parts())
            else:
                modules = modules if modules is not None else self.app.services.modules.list_modules()
                for m in modules:
                    if hasattr(self.app.services.modules, "get_module_components"):
                        parts_count += len(self.app.services.modules.get_module_components(m.module_code))
        except Exception:
            parts_count = 0

        modules = modules if modules is not None else self.app.services.modules.list_modules()
        products = products if products is not None else self.app.services.products.list_products()
        projects = projects if projects is not None else self.app.services.projects.list_projects()

        self.parts_card.config(text=f"Parts\n{parts_count}\nStock, supplier, lead time")
        self.modules_card.config(text=f"Assemblies\n{len(modules)}\nBuild definitions")
        self.products_card.config(text=f"Products\n{len(products)}\nConfigured products")
        self.projects_card.config(text=f"Live Orders\n{len(projects)}\nExecution jobs")

    def refresh_combos(self, modules=None, products=None, projects=None):
        modules = modules if modules is not None else self.app.services.modules.list_modules()
        products = products if products is not None else self.app.services.products.list_products()
        projects = projects if projects is not None else self.app.services.projects.list_projects()

        set_combobox_values(self.module_combo, [m.module_code for m in modules], keep_current=False)
        set_combobox_values(self.product_combo, [p.product_code for p in products], keep_current=False)
        set_combobox_values(self.project_combo, [p.project_code for p in projects], keep_current=False)

        if getattr(self.app, "selected_module_code", ""):
            self.selected_module_var.set(self.app.selected_module_code)
        if getattr(self.app, "selected_product_code", ""):
            self.selected_product_var.set(self.app.selected_product_code)
        if getattr(self.app, "selected_project_code", ""):
            self.selected_project_var.set(self.app.selected_project_code)

    def refresh_recent_modules(self, modules=None):
        treeview_clear(self.recent_modules_tree)
        q = self.module_search_var.get().strip().lower()

        records = modules if modules is not None and not q else self.app.services.modules.list_modules(search_text=q)
        for m in records[:50]:
            self.recent_modules_tree.insert(
                "",
                "end",
                values=(
                    m.module_code,
                    m.quote_ref,
                    m.module_name,
                    m.status,
                    f"{float(m.estimated_hours or 0.0):.2f}",
                    m.updated_on,
                )
            )

    def refresh_recent_products(self, products=None):
        treeview_clear(self.recent_products_tree)
        q = self.product_search_var.get().strip().lower()

        records = products if products is not None and not q else self.app.services.products.list_products(search_text=q)
        for p in records[:50]:
            self.recent_products_tree.insert(
                "",
                "end",
                values=(
                    p.product_code,
                    p.quote_ref,
                    p.product_name,
                    p.revision,
                    p.status,
                    p.updated_on,
                )
            )

    def open_parts_manager(self):
        self.show_page("parts")
        
    def refresh_recent_projects(self, projects=None):
        treeview_clear(self.recent_projects_tree)
        q = self.project_search_var.get().strip().lower()

        records = projects if projects is not None and not q else self.app.services.projects.list_projects(search_text=q)
        for p in records[:50]:
            self.recent_projects_tree.insert(
                "",
                "end",
                values=(
                    p.project_code,
                    p.quote_ref,
                    p.project_name,
                    p.client_name,
                    p.status,
                    p.due_date,
                    p.updated_on,
                )
            )
