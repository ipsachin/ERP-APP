# ============================================================
# ui_projects.py
# Live project management page for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog
from typing import Dict, List, Optional, Tuple

from app_config import AppConfig
from models import (
    ProjectRecord,
    ProjectModuleLinkRecord,
    ProjectTaskRecord,
    ProjectDocumentRecord,
    WorkOrderRecord,
)
from ui_common import (
    BasePage,
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


class ProjectPage(BasePage):
    PAGE_NAME = "projects"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        self.current_project_module_link_id: Optional[str] = None
        self.current_project_task_id: Optional[str] = None
        self.current_project_document_id: Optional[str] = None
        self.current_workorder_id: Optional[str] = None

        # ----------------------------------------------------
        # Project vars
        # ----------------------------------------------------
        self.project_code_var = tk.StringVar()
        self.quote_ref_var = tk.StringVar()
        self.project_name_var = tk.StringVar()
        self.client_name_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.linked_product_var = tk.StringVar()
        self.status_var_local = tk.StringVar(value="Planned")
        self.start_date_var = tk.StringVar()
        self.due_date_var = tk.StringVar()

        # ----------------------------------------------------
        # Project module vars
        # ----------------------------------------------------
        self.direct_module_var = tk.StringVar()
        self.project_module_dependency_var = tk.StringVar()
        self.project_module_stage_var = tk.StringVar(value="Not Started")
        self.project_module_status_var = tk.StringVar(value="Not Started")
        self.project_module_notes_var = tk.StringVar()

        # ----------------------------------------------------
        # Project task vars
        # ----------------------------------------------------
        self.project_task_select_var = tk.StringVar()
        self.project_task_module_var = tk.StringVar()
        self.project_task_name_var = tk.StringVar()
        self.project_task_department_var = tk.StringVar()
        self.project_task_hours_var = tk.StringVar()
        self.project_task_stage_var = tk.StringVar()
        self.project_task_status_var = tk.StringVar(value="Not Started")
        self.project_task_dependency_var = tk.StringVar()
        self.project_task_assigned_to_var = tk.StringVar()
        self.project_task_notes_var = tk.StringVar()

        # ----------------------------------------------------
        # Project doc vars
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

        self.project_select_top_var = tk.StringVar()

        self._build_ui()

    # ========================================================
    # UI
    # ========================================================

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        self._build_topbar(wrapper)

        # ADD THIS BACK
        self.project_summary_label = ttk.Label(
            wrapper,
            text="No live project selected",
            style="Title.TLabel"
        )
        self.project_summary_label.pack(anchor="w", pady=(0, 10))

        paned = ttk.Panedwindow(wrapper, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # LEFT SCROLLABLE
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

        self.canvas_window = self.left_canvas.create_window(
            (0, 0),
            window=self.left_frame,
            anchor="nw"
        )

        def resize(event):
            self.left_canvas.itemconfig(self.canvas_window, width=event.width)

        self.left_canvas.bind("<Configure>", resize)
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)

        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_scrollbar.pack(side="right", fill="y")

        self._build_left_panel(self.left_frame)
        self._build_right_panel(right)

        self._bind_mousewheel()


    def _bind_mousewheel(self):
        def _on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.left_canvas.bind_all("<MouseWheel>", _on_mousewheel)


    def _build_topbar(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 8))

        left = ttk.Frame(top)
        left.pack(side="left", fill="x", expand=True)

        ttk.Button(left, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)

        ttk.Label(left, text="Open Project").pack(side="left", padx=(12, 4))

        self.project_select_top_combo = ttk.Combobox(
            left,
            textvariable=self.project_select_top_var,
            state="readonly",
            width=60
        )
        self.project_select_top_combo.pack(side="left", padx=4)
        self.project_select_top_combo.bind("<<ComboboxSelected>>", lambda e: self.open_selected_project_from_dropdown())

        ttk.Button(left, text="Open", command=self.open_selected_project_from_dropdown).pack(side="left", padx=4)

        right = ttk.Frame(top)
        right.pack(side="right")

        ttk.Button(right, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)
        ttk.Button(right, text="Email Project PDF", command=self.email_project_pdf).pack(side="right", padx=2)
        ttk.Button(right, text="Generate Project PDF", command=self.generate_project_pdf).pack(side="right", padx=2)



    def refresh_project_selector(self):
        if not self.app.workbook_manager.has_workbook():
            self.project_select_top_combo["values"] = []
            return

        records = self.app.services.projects.list_projects()
        values = [
            f"{p.project_code} | {p.quote_ref} | {p.project_name}"
            for p in records
        ]
        set_combobox_values(self.project_select_top_combo, values, keep_current=True)

        selected_code = getattr(self.app, "selected_project_code", "").strip()
        if selected_code:
            for p in records:
                if p.project_code == selected_code:
                    self.project_select_top_var.set(f"{p.project_code} | {p.quote_ref} | {p.project_name}")
                    break


    def open_selected_project_from_dropdown(self):
        selected = self.project_select_top_var.get().strip()
        if not selected:
            return

        project_code = selected.split("|")[0].strip()
        self.app.set_selected_project(project_code)
        self.refresh_page()

    def _build_left_panel(self, parent):
        self._build_project_browser_card(parent)
        self._build_project_details_card(parent)
        self._build_project_module_builder_card(parent)
        self._build_project_task_editor_card(parent)
        self._build_project_document_card(parent)
        self._build_workorder_card(parent)

    def _build_right_panel(self, parent):
        tabs = ttk.Notebook(parent)
        tabs.pack(fill="both", expand=True)

        self.modules_tab = ttk.Frame(tabs)
        self.tasks_tab = ttk.Frame(tabs)
        self.docs_tab = ttk.Frame(tabs)
        self.workorders_tab = ttk.Frame(tabs)
        self.summary_tab = ttk.Frame(tabs)

        tabs.add(self.modules_tab, text="Project Modules")
        tabs.add(self.tasks_tab, text="Project Tasks")
        tabs.add(self.docs_tab, text="Project Docs")
        tabs.add(self.workorders_tab, text="Work Orders")
        tabs.add(self.summary_tab, text="Summary")

        self._build_modules_tab(self.modules_tab)
        self._build_tasks_tab(self.tasks_tab)
        self._build_docs_tab(self.docs_tab)
        self._build_workorders_tab(self.workorders_tab)
        self._build_summary_tab(self.summary_tab)


    def _build_project_browser_card(self, parent):
        card = ttk.LabelFrame(parent, text="Project Browser", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        tabs = ttk.Notebook(card)
        tabs.pack(fill="both", expand=True)

        live_tab = ttk.Frame(tabs)
        completed_tab = ttk.Frame(tabs)
        preview_tab = ttk.Frame(tabs)

        tabs.add(live_tab, text="Live Projects")
        tabs.add(completed_tab, text="Completed Projects")
        tabs.add(preview_tab, text="Task Preview")

        # Live projects tree
        live_wrap = ttk.Frame(live_tab, padding=4)
        live_wrap.pack(fill="both", expand=True)

        self.live_projects_tree = ttk.Treeview(
            live_wrap,
            columns=("ProjectName", "Client", "Status", "DueDate"),
            show="headings",
            height=6
        )
        for col, width in [
            ("ProjectName", 220),
            ("Client", 160),
            ("Status", 110),
            ("DueDate", 100),
        ]:
            self.live_projects_tree.heading(col, text=col)
            self.live_projects_tree.column(col, width=width, anchor="w")

        live_sb = ttk.Scrollbar(live_wrap, orient="vertical", command=self.live_projects_tree.yview)
        self.live_projects_tree.configure(yscrollcommand=live_sb.set)
        self.live_projects_tree.pack(side="left", fill="both", expand=True)
        live_sb.pack(side="right", fill="y")

        self.live_projects_tree.bind("<ButtonRelease-1>", lambda e: self.load_project_from_browser(live_only=True))
        self.live_projects_tree.bind("<Double-1>", lambda e: self.load_project_from_browser(live_only=True))

        # Completed projects tree
        comp_wrap = ttk.Frame(completed_tab, padding=4)
        comp_wrap.pack(fill="both", expand=True)

        self.completed_projects_tree = ttk.Treeview(
            comp_wrap,
            columns=("ProjectName", "Client", "Status", "DueDate"),
            show="headings",
            height=6
        )
        for col, width in [
            ("ProjectName", 220),
            ("Client", 160),
            ("Status", 110),
            ("DueDate", 100),
        ]:
            self.completed_projects_tree.heading(col, text=col)
            self.completed_projects_tree.column(col, width=width, anchor="w")

        comp_sb = ttk.Scrollbar(comp_wrap, orient="vertical", command=self.completed_projects_tree.yview)
        self.completed_projects_tree.configure(yscrollcommand=comp_sb.set)
        self.completed_projects_tree.pack(side="left", fill="both", expand=True)
        comp_sb.pack(side="right", fill="y")

        self.completed_projects_tree.bind("<ButtonRelease-1>", lambda e: self.load_project_from_browser(live_only=False))
        self.completed_projects_tree.bind("<Double-1>", lambda e: self.load_project_from_browser(live_only=False))

        # Task preview tree
        preview_wrap = ttk.Frame(preview_tab, padding=4)
        preview_wrap.pack(fill="both", expand=True)

        self.project_preview_tasks_tree = ttk.Treeview(
            preview_wrap,
            columns=("ModuleCode", "TaskName", "Status"),
            show="headings",
            height=8
        )
        for col, width in [
            ("ModuleCode", 140),
            ("TaskName", 240),
            ("Status", 120),
        ]:
            self.project_preview_tasks_tree.heading(col, text=col)
            self.project_preview_tasks_tree.column(col, width=width, anchor="w")

        preview_sb = ttk.Scrollbar(preview_wrap, orient="vertical", command=self.project_preview_tasks_tree.yview)
        self.project_preview_tasks_tree.configure(yscrollcommand=preview_sb.set)
        self.project_preview_tasks_tree.pack(side="left", fill="both", expand=True)
        preview_sb.pack(side="right", fill="y")
    # --------------------------------------------------------
    # Left cards
    # --------------------------------------------------------

    def _build_project_details_card(self, parent):
        card = ttk.LabelFrame(parent, text="Live Project Details", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Project Code").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_code_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Quote Ref").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.quote_ref_var).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Project Name").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_name_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Client Name").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.client_name_var).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Location").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.location_var).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Description").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.description_var).grid(row=5, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Linked Product").grid(row=6, column=0, sticky="w", pady=4)
        self.linked_product_combo = ttk.Combobox(
            card,
            textvariable=self.linked_product_var,
            state="readonly"
        )
        self.linked_product_combo.grid(row=6, column=1, sticky="ew", padx=6)

        prod_btn_row = ttk.Frame(card)
        prod_btn_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        ttk.Button(prod_btn_row, text="Attach Product", command=self.attach_product_to_project).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(prod_btn_row, text="Rebuild Modules", command=self.rebuild_modules_from_product).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(prod_btn_row, text="Rebuild Tasks", command=self.rebuild_project_tasks).pack(side="left", fill="x", expand=True, padx=2)

        ttk.Label(card, text="Status").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.status_var_local,
            values=AppConfig.PROJECT_STATUSES,
            state="readonly"
        ).grid(row=8, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Start Date").grid(row=9, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.start_date_var).grid(row=9, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Due Date").grid(row=10, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.due_date_var).grid(row=10, column=1, sticky="ew", padx=6)

        # btn_row = ttk.Frame(card)
        # btn_row.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        # save_btn = ttk.Button(btn_row, text="Save Project", command=self.save_project)
        # save_btn.pack(side="left", fill="x", expand=True, padx=2)
        # attach_tooltip(save_btn, "Create a new project or update the selected project.")

        # del_btn = ttk.Button(btn_row, text="Delete Project", command=self.delete_project)
        # del_btn.pack(side="left", fill="x", expand=True, padx=2)
        # attach_tooltip(del_btn, "Delete selected project and linked project-level records.")

        btn_row = ttk.Frame(card)
        btn_row.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        save_btn = ttk.Button(btn_row, text="Save Project", command=self.save_project)
        save_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(save_btn, "Create a new project or update the selected project.")

        complete_btn = ttk.Button(btn_row, text="Mark Completed", command=self.mark_project_completed)
        complete_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(complete_btn, "Mark this live project as completed and move it to completed projects list.")

        del_btn = ttk.Button(btn_row, text="Delete Project", command=self.delete_project)
        del_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(del_btn, "Delete selected project and linked project-level records.")

        card.columnconfigure(1, weight=1)


    def mark_project_completed(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return

        if not tk.messagebox.askyesno(
            "Complete Project",
            "Mark this project as completed?\n\nIt will move from Live Projects to Completed Projects."
        ):
            return

        try:
            self.app.services.projects.create_or_update_project(
                quote_ref=self.quote_ref_var.get().strip(),
                project_name=self.project_name_var.get().strip(),
                client_name=self.client_name_var.get().strip(),
                location=self.location_var.get().strip(),
                description=self.description_var.get().strip(),
                linked_product_code=self.linked_product_var.get().strip(),
                status="Completed",
                start_date=self.start_date_var.get().strip(),
                due_date=self.due_date_var.get().strip(),
                existing_project_code=project_code,
            )
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Project marked completed: {project_code}")
        except Exception as exc:
            show_error("Complete Project Error", str(exc))


    def load_project_from_browser(self, live_only=True):
        tree = self.live_projects_tree if live_only else self.completed_projects_tree
        sel = tree.selection()
        if not sel:
            return

        tags = tree.item(sel[0], "tags")
        if not tags:
            return

        project_code = tags[0]
        self.app.set_selected_project(project_code)
        self.refresh_page()


    def refresh_project_browser(self):
        treeview_clear(self.live_projects_tree)
        treeview_clear(self.completed_projects_tree)
        treeview_clear(self.project_preview_tasks_tree)

        projects = self.app.services.projects.list_projects()

        live_projects = [p for p in projects if norm_text(p.status) != "Completed"]
        completed_projects = [p for p in projects if norm_text(p.status) == "Completed"]

        for p in live_projects:
            self.live_projects_tree.insert(
                "",
                "end",
                values=(p.project_name, p.client_name, p.status, p.due_date),
                tags=(p.project_code,)
            )

        for p in completed_projects:
            self.completed_projects_tree.insert(
                "",
                "end",
                values=(p.project_name, p.client_name, p.status, p.due_date),
                tags=(p.project_code,)
            )

        selected_project = self._selected_project_code()
        if selected_project:
            tasks = self.app.services.projects.get_project_tasks(selected_project)
            for t in tasks[:30]:
                self.project_preview_tasks_tree.insert(
                    "",
                    "end",
                    values=(t.module_code, t.task_name, t.status)
                )   


    def _build_project_module_builder_card(self, parent):
        card = ttk.LabelFrame(parent, text="Project Module Planner", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Direct Module").grid(row=0, column=0, sticky="w", pady=4)
        self.direct_module_combo = ttk.Combobox(
            card,
            textvariable=self.direct_module_var,
            state="readonly"
        )
        self.direct_module_combo.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Button(card, text="Add Direct Module", command=self.add_direct_module).grid(row=0, column=2, sticky="ew", padx=4)

        ttk.Label(card, text="Dependency").grid(row=1, column=0, sticky="w", pady=4)
        self.project_module_dependency_combo = ttk.Combobox(
            card,
            textvariable=self.project_module_dependency_var,
            state="readonly"
        )
        self.project_module_dependency_combo.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Button(card, text="Set Selected Dependency", command=self.set_selected_project_module_dependency).grid(row=1, column=2, sticky="ew", padx=4)

        ttk.Label(card, text="Stage").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.project_module_stage_var,
            values=AppConfig.MODULE_EXEC_STAGES,
            state="readonly"
        ).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.project_module_status_var,
            values=AppConfig.TASK_STATUSES,
            state="readonly"
        ).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Notes").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_module_notes_var).grid(row=4, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Update Selected Module", command=self.update_selected_project_module).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Remove Selected Module", command=self.remove_selected_project_module).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)
    

    def _build_project_task_editor_card(self, parent):
        card = ttk.LabelFrame(parent, text="Project Task Execution Editor", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Select Module").grid(row=0, column=0, sticky="w", pady=4)
        self.project_task_module_combo = ttk.Combobox(
            card,
            textvariable=self.project_task_module_var,
            state="readonly"
        )
        self.project_task_module_combo.grid(row=0, column=1, sticky="ew", padx=6)
        self.project_task_module_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_module_task_dropdown())

        ttk.Label(card, text="Select Task").grid(row=1, column=0, sticky="w", pady=4)
        self.project_task_select_combo = ttk.Combobox(
            card,
            textvariable=self.project_task_select_var,
            state="readonly"
        )
        self.project_task_select_combo.grid(row=1, column=1, sticky="ew", padx=6)
        self.project_task_select_combo.bind("<<ComboboxSelected>>", lambda e: self.load_project_task_from_dropdown())

        ttk.Label(card, text="Task Name").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_task_name_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Department").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.project_task_department_var,
            values=AppConfig.DEPARTMENTS,
            state="readonly"
        ).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Hours").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_task_hours_var).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Stage").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.project_task_stage_var,
            values=AppConfig.MODULE_EXEC_STAGES,
            state="readonly"
        ).grid(row=5, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.project_task_status_var,
            values=AppConfig.TASK_STATUSES,
            state="readonly"
        ).grid(row=6, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Dependency Task").grid(row=7, column=0, sticky="w", pady=4)
        self.project_task_dependency_combo = ttk.Combobox(
            card,
            textvariable=self.project_task_dependency_var,
            state="readonly"
        )
        self.project_task_dependency_combo.grid(row=7, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Assigned To").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_task_assigned_to_var).grid(row=8, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Notes").grid(row=9, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.project_task_notes_var).grid(row=9, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Update Selected Task", command=self.update_selected_project_task).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected Task", command=self.delete_selected_project_task).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Clear", command=self.clear_project_task_editor).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    def refresh_project_task_module_dropdown(self, module_links):
        module_codes = [m.module_code for m in module_links]
        set_combobox_values(self.project_task_module_combo, module_codes, keep_current=True)

    def refresh_module_task_dropdown(self):
        project_code = self._selected_project_code()
        selected_module = self.project_task_module_var.get().strip()

        if not project_code or not selected_module:
            self.project_task_select_combo["values"] = []
            self.project_task_select_var.set("")
            return

        tasks = self.app.services.projects.get_project_tasks(project_code)
        module_tasks = [t for t in tasks if norm_text(t.module_code) == selected_module]

        values = [t.task_name for t in module_tasks]
        set_combobox_values(self.project_task_select_combo, values, keep_current=False)

        # Also refresh dependency dropdown for same project
        dep_values = [f"{t.module_code} | {t.task_name}" for t in tasks]
        set_combobox_values(self.project_task_dependency_combo, dep_values, keep_current=True)

    def load_project_task_from_dropdown(self):
        project_code = self._selected_project_code()
        selected_module = self.project_task_module_var.get().strip()
        selected_task_name = self.project_task_select_var.get().strip()

        if not project_code or not selected_module or not selected_task_name:
            return

        tasks = self.app.services.projects.get_project_tasks(project_code)

        for t in tasks:
            if norm_text(t.module_code) == selected_module and norm_text(t.task_name) == selected_task_name:
                self.current_project_task_id = t.project_task_id
                self._load_project_task_into_editor(t)
                break

    def _load_project_task_into_editor(self, task):
        self.project_task_name_var.set(task.task_name)
        self.project_task_module_var.set(task.module_code)
        self.project_task_select_var.set(task.task_name)
        self.project_task_department_var.set(task.department)
        self.project_task_hours_var.set(str(task.estimated_hours))
        self.project_task_stage_var.set(task.stage)
        self.project_task_status_var.set(task.status)
        self.project_task_assigned_to_var.set(task.assigned_to)
        self.project_task_notes_var.set(task.notes)

        dep_display = ""
        if norm_text(task.dependency_task_id):
            tasks = self.app.services.projects.get_project_tasks(self._selected_project_code())
            for t in tasks:
                if t.project_task_id == task.dependency_task_id:
                    dep_display = f"{t.module_code} | {t.task_name}"
                    break
            if not dep_display:
                dep_display = task.dependency_task_id

        self.project_task_dependency_var.set(dep_display)

    def _build_project_document_card(self, parent):
        card = ttk.LabelFrame(parent, text="Project Documents", style="Card.TLabelframe", padding=12)
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
        ttk.Button(btn_row, text="Add File", command=self.add_project_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Open Selected", command=self.open_selected_project_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_project_document).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    def _build_workorder_card(self, parent):
        card = ttk.LabelFrame(parent, text="Project Work Orders", style="Card.TLabelframe", padding=12)
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

        cols = ("Order", "ModuleCode", "SourceType", "SourceCode", "Qty", "Stage", "Status", "Dependency", "Notes")
        self.project_modules_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Order", 60),
            ("ModuleCode", 220),
            ("SourceType", 110),
            ("SourceCode", 220),
            ("Qty", 60),
            ("Stage", 130),
            ("Status", 120),
            ("Dependency", 200),
            ("Notes", 240),
        ]:
            self.project_modules_tree.heading(col, text=col)
            self.project_modules_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.project_modules_tree.yview)
        self.project_modules_tree.configure(yscrollcommand=sb.set)

        self.project_modules_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.project_modules_tree.bind("<ButtonRelease-1>", lambda e: self.load_selected_project_module())
        self.project_modules_tree.bind("<Double-1>", lambda e: self.load_selected_project_module())

    def _build_tasks_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("ModuleCode", "TaskName", "Department", "Hours", "Stage", "Status", "AssignedTo", "DependencyTaskID", "Notes")
        self.project_tasks_tree = ttk.Treeview(wrap, columns=cols, show="headings")
        self.project_tasks_tree.bind("<ButtonRelease-1>", self.on_project_task_select)
        self.project_tasks_tree.bind("<Double-1>", self.on_project_task_select)

        for col, width in [
            ("ModuleCode", 180),
            ("TaskName", 240),
            ("Department", 130),
            ("Hours", 70),
            ("Stage", 130),
            ("Status", 120),
            ("AssignedTo", 140),
            ("DependencyTaskID", 220),
            ("Notes", 240),
        ]:
            self.project_tasks_tree.heading(col, text=col)
            self.project_tasks_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.project_tasks_tree.yview)
        self.project_tasks_tree.configure(yscrollcommand=sb.set)

        self.project_tasks_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.project_tasks_tree.bind("<ButtonRelease-1>", lambda e: self.load_selected_project_task())
        self.project_tasks_tree.bind("<Double-1>", lambda e: self.load_selected_project_task())

    
    
    
    def _build_docs_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("SectionName", "DocName", "DocType", "FilePath", "AddedOn")
        self.project_document_tree = ttk.Treeview(wrap, columns=cols, show="headings")
        self.project_document_tree.bind("<Double-1>", lambda e: self.open_selected_project_document())

        for col, width in [
            ("SectionName", 150),
            ("DocName", 220),
            ("DocType", 150),
            ("FilePath", 420),
            ("AddedOn", 160),
        ]:
            self.project_document_tree.heading(col, text=col)
            self.project_document_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.project_document_tree.yview)
        self.project_document_tree.configure(yscrollcommand=sb.set)

        self.project_document_tree.pack(side="left", fill="both", expand=True)
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

    def on_project_task_select(self, event=None):
        sel = self.project_tasks_tree.selection()
        if not sel:
            return

        item = self.project_tasks_tree.item(sel[0])
        tags = item.get("tags", [])

        if not tags:
            return

        task_id = tags[0]

        tasks = self.app.services.projects.get_project_tasks(self._selected_project_code())
        for t in tasks:
            if t.project_task_id == task_id:
                self._load_task_into_editor(t)
                break


    def _load_task_into_editor(self, task):
        self.task_name_var.set(task.task_name)
        self.task_module_var.set(task.module_code)
        self.task_department_var.set(task.department)
        self.task_hours_var.set(str(task.estimated_hours))
        self.task_stage_var.set(task.stage)
        self.task_status_var.set(task.status)
        self.task_assigned_var.set(task.assigned_to)
        self.task_dependency_var.set(task.dependency_task_id)
        self.task_notes_var.set(task.notes)
    # ========================================================
    # Page refresh
    # ========================================================

    def refresh_page(self):
        if not self.require_workbook():
            return
        
        self.refresh_project_selector()
        project_code = getattr(self.app, "selected_project_code", "").strip()
        if not project_code:
            self._clear_all_views()
            self.project_summary_label.config(text="No live project selected")
            return

        try:
            bundle = self.app.services.projects.get_project_bundle(project_code)
            if not bundle.project:
                self._clear_all_views()
                self.project_summary_label.config(text="Project not found")
                return

            self._load_project_into_form(bundle.project)
            self._load_project_modules(bundle.module_links or [])
            self._load_project_tasks(bundle.project_tasks or [])
            self._load_project_documents(bundle.project_documents or [])
            self._load_workorders(bundle.workorders or [])
            self._load_summary(bundle)
            if hasattr(self, "live_projects_tree"):
                 self.refresh_project_browser()
            self._refresh_products_combo()
            self._refresh_modules_combo()
            self._refresh_project_module_dependency_combo(bundle.module_links or [])
            self._refresh_project_task_dependency_combo(bundle.project_tasks or [])
            self.refresh_project_task_module_dropdown(bundle.module_links or [])
            self.refresh_module_task_dropdown()

            self.project_summary_label.config(
                text=f"🏗 {bundle.project.project_name} | Client: {bundle.project.client_name} | Status: {bundle.project.status} | Code: {bundle.project.project_code}"
            )
            self.set_status(f"Loaded project: {bundle.project.project_code}")

        except Exception as exc:
            show_error("Project Refresh Error", str(exc))

    def _clear_all_views(self):
        self.project_code_var.set("")
        self.quote_ref_var.set("")
        self.project_name_var.set("")
        self.client_name_var.set("")
        self.location_var.set("")
        self.description_var.set("")
        self.linked_product_var.set("")
        self.status_var_local.set("Planned")
        self.start_date_var.set("")
        self.due_date_var.set("")

        treeview_clear(self.project_modules_tree)
        treeview_clear(self.project_tasks_tree)
        treeview_clear(self.project_document_tree)
        treeview_clear(self.workorder_tree)

        set_text_readonly(self.summary_text, "")
        self.clear_project_task_editor()
        self.clear_workorder_editor()

    def _load_project_into_form(self, project: ProjectRecord):
        self.project_code_var.set(project.project_code)
        self.quote_ref_var.set(project.quote_ref)
        self.project_name_var.set(project.project_name)
        self.client_name_var.set(project.client_name)
        self.location_var.set(project.location)
        self.description_var.set(project.description)
        self.linked_product_var.set(project.linked_product_code)
        self.status_var_local.set(project.status or "Planned")
        self.start_date_var.set(project.start_date)
        self.due_date_var.set(project.due_date)

    def _load_project_modules(self, links: List[ProjectModuleLinkRecord]):
        treeview_clear(self.project_modules_tree)
        for link in links:
            self.project_modules_tree.insert(
                "",
                "end",
                values=(
                    link.module_order,
                    link.module_code,
                    link.source_type,
                    link.source_code,
                    link.module_qty,
                    link.stage,
                    link.status,
                    link.dependency_module_code,
                    link.notes,
                ),
                tags=(link.link_id,),
            )

    def _load_project_tasks(self, tasks: List[ProjectTaskRecord]):
        treeview_clear(self.project_tasks_tree)
        for t in tasks:
            self.project_tasks_tree.insert(
                "",
                "end",
                values=(
                    t.module_code,
                    t.task_name,
                    t.department,
                    f"{float(t.estimated_hours or 0.0):.2f}",
                    t.stage,
                    t.status,
                    t.assigned_to,
                    t.dependency_task_id,
                    t.notes,
                ),
                tags=(t.project_task_id,),
            )

    def _load_project_documents(self, docs: List[ProjectDocumentRecord]):
        treeview_clear(self.project_document_tree)
        for doc in docs:
            self.project_document_tree.insert(
                "",
                "end",
                values=(
                    doc.section_name,
                    doc.doc_name,
                    doc.doc_type,
                    doc.file_path,
                    doc.added_on,
                ),
                tags=(doc.project_doc_id,),
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
        project = bundle.project
        module_links = bundle.module_links or []
        project_tasks = bundle.project_tasks or []
        project_docs = bundle.project_documents or []
        workorders = bundle.workorders or []

        blockers = []
        try:
            blockers = self.app.services.scheduler.get_open_blockers_for_project(project.project_code)
        except Exception:
            blockers = []

        lines = [
            f"Project Code: {project.project_code}",
            f"Quote Ref: {project.quote_ref}",
            f"Project Name: {project.project_name}",
            f"Client Name: {project.client_name}",
            f"Location: {project.location}",
            f"Description: {project.description}",
            f"Linked Product: {project.linked_product_code or '-'}",
            f"Status: {project.status}",
            f"Start Date: {project.start_date or '-'}",
            f"Due Date: {project.due_date or '-'}",
            "",
            f"Project Modules: {len(module_links)}",
            f"Project Tasks: {len(project_tasks)}",
            f"Project Documents: {len(project_docs)}",
            f"Work Orders: {len(workorders)}",
            f"Aggregated Execution Hours: {float(bundle.total_hours or 0.0):.2f}",
            "",
            "Module Execution Status:",
        ]

        if not module_links:
            lines.append(" - No modules loaded.")
        else:
            for link in module_links:
                dep_text = f" | Depends on: {link.dependency_module_code}" if norm_text(link.dependency_module_code) else ""
                lines.append(
                    f" {link.module_order:02d}. {link.module_code} | {link.source_type} | Qty {link.module_qty} | Stage {link.stage} | Status {link.status}{dep_text}"
                )

        lines.append("")
        lines.append("Open Blockers:")

        if not blockers:
            lines.append(" - None")
        else:
            for b in blockers:
                if b.get("type") == "TASK":
                    lines.append(
                        f" - TASK: {b.get('task_name')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})"
                    )
                elif b.get("type") == "MODULE":
                    lines.append(
                        f" - MODULE: {b.get('module_code')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})"
                    )

        set_text_readonly(self.summary_text, "\n".join(lines))

    def _refresh_products_combo(self):
        products = self.app.services.products.list_products()
        set_combobox_values(self.linked_product_combo, [p.product_code for p in products], keep_current=True)

    def _refresh_modules_combo(self):
        modules = self.app.services.modules.list_modules()
        set_combobox_values(self.direct_module_combo, [m.module_code for m in modules], keep_current=True)

    def _refresh_project_module_dependency_combo(self, module_links: List[ProjectModuleLinkRecord]):
        set_combobox_values(
            self.project_module_dependency_combo,
            [m.module_code for m in module_links],
            keep_current=True
        )

    def _refresh_project_task_dependency_combo(self, tasks: List[ProjectTaskRecord]):
        display_values = [f"{t.module_code} | {t.task_name}" for t in tasks]
        set_combobox_values(self.project_task_dependency_combo, display_values, keep_current=True)

    def _refresh_project_task_module_combo(self, module_links: List[ProjectModuleLinkRecord]):
        set_combobox_values(
            self.project_task_module_combo,
            [m.module_code for m in module_links],
            keep_current=True
        )

    # ========================================================
    # Helpers
    # ========================================================

    def _selected_project_code(self) -> str:
        return self.project_code_var.get().strip() or getattr(self.app, "selected_project_code", "").strip()

    def _selected_project_task_dependency_id(self) -> str:
        display = self.project_task_dependency_var.get().strip()
        if not display:
            return ""
        tasks = self.app.services.projects.get_project_tasks(self._selected_project_code())
        for t in tasks:
            if display == f"{t.module_code} | {t.task_name}":
                return t.project_task_id
        return ""

    # ========================================================
    # Project CRUD
    # ========================================================

    def save_project(self):
        if not self.require_workbook():
            return

        try:
            quote_ref = self.quote_ref_var.get().strip()
            project_name = Validators.require_text(self.project_name_var.get(), "Project name")
            client_name = self.client_name_var.get().strip()
            location = self.location_var.get().strip()
            description = self.description_var.get().strip()
            linked_product_code = self.linked_product_var.get().strip()
            status = self.status_var_local.get().strip() or "Planned"
            start_date = self.start_date_var.get().strip()
            due_date = self.due_date_var.get().strip()
            existing_code = self.project_code_var.get().strip() or None

            new_code = self.app.services.projects.create_or_update_project(
                quote_ref=quote_ref,
                project_name=project_name,
                client_name=client_name,
                location=location,
                description=description,
                linked_product_code=linked_product_code,
                status=status,
                start_date=start_date,
                due_date=due_date,
                existing_project_code=existing_code,
            )

            self.app.set_selected_project(new_code)
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Project saved: {new_code}")

        except Exception as exc:
            show_error("Save Project Error", str(exc))

    def delete_project(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "No project selected.")
            return

        if not tk.messagebox.askyesno(
            "Delete Project",
            "Delete this project?\n\nThis will also delete project modules, tasks, documents, and project work orders."
        ):
            return

        try:
            self.app.services.projects.delete_project(project_code, delete_docs_files=False)
            self.app.set_selected_project("")
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Project deleted: {project_code}")
        except Exception as exc:
            show_error("Delete Project Error", str(exc))

    # ========================================================
    # Product attach / rebuild
    # ========================================================

    def attach_product_to_project(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        product_code = self.linked_product_var.get().strip()

        if not project_code:
            show_warning("No Project", "Save or select a project first.")
            return
        if not product_code:
            show_warning("No Product", "Select a product first.")
            return

        try:
            self.app.services.projects.attach_product(
                project_code=project_code,
                product_code=product_code,
                rebuild_modules=True,
                rebuild_tasks=True,
            )
            self.refresh_page()
            self.set_status(f"Product attached to project: {product_code}")
        except Exception as exc:
            show_error("Attach Product Error", str(exc))

    def rebuild_modules_from_product(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return

        try:
            self.app.services.projects.rebuild_project_modules_from_product(project_code)
            self.refresh_page()
            self.set_status("Project modules rebuilt from linked product.")
        except Exception as exc:
            show_error("Rebuild Modules Error", str(exc))

    def rebuild_project_tasks(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return

        if not tk.messagebox.askyesno(
            "Rebuild Project Tasks",
            "Rebuild project tasks from current project modules?\n\nThis will replace existing project execution tasks."
        ):
            return

        try:
            self.app.services.projects.populate_project_tasks_from_modules(project_code, clear_existing=True)
            self.refresh_page()
            self.set_status("Project tasks rebuilt from modules.")
        except Exception as exc:
            show_error("Rebuild Tasks Error", str(exc))

    # ========================================================
    # Project modules
    # ========================================================

    def add_direct_module(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        module_code = self.direct_module_var.get().strip()

        if not project_code:
            show_warning("No Project", "Save or select a project first.")
            return
        if not module_code:
            show_warning("No Module", "Select a module first.")
            return

        try:
            self.app.services.projects.add_direct_module(
                project_code=project_code,
                module_code=module_code,
                qty=1,
                stage=self.project_module_stage_var.get().strip() or "Not Started",
                status=self.project_module_status_var.get().strip() or "Not Started",
                dependency_module_code=self.project_module_dependency_var.get().strip(),
                notes=self.project_module_notes_var.get().strip(),
            )
            self.refresh_page()
            self.set_status(f"Direct module added: {module_code}")
        except Exception as exc:
            show_error("Add Direct Module Error", str(exc))

    def load_selected_project_module(self):
        sel = self.project_modules_tree.selection()
        if not sel:
            return

        tags = self.project_modules_tree.item(sel[0], "tags")
        if not tags:
            return

        link_id = tags[0]
        project_code = self._selected_project_code()
        links = self.app.services.projects.get_project_module_links(project_code)

        selected = None
        for l in links:
            if l.link_id == link_id:
                selected = l
                break

        if not selected:
            return

        self.current_project_module_link_id = selected.link_id
        self.direct_module_var.set(selected.module_code)
        self.project_module_dependency_var.set(selected.dependency_module_code)
        self.project_module_stage_var.set(selected.stage)
        self.project_module_status_var.set(selected.status)
        self.project_module_notes_var.set(selected.notes)

    def update_selected_project_module(self):
        if not self.require_workbook():
            return

        if not self.current_project_module_link_id:
            show_warning("No Selection", "Select a project module first.")
            return

        dep_code = self.project_module_dependency_var.get().strip()
        selected_module_code = self.direct_module_var.get().strip()

        if dep_code and dep_code == selected_module_code:
            show_warning("Invalid Dependency", "A module cannot depend on itself.")
            return

        try:
            self.app.services.projects.update_project_module_status(
                self.current_project_module_link_id,
                stage=self.project_module_stage_var.get().strip(),
                status=self.project_module_status_var.get().strip(),
                notes=self.project_module_notes_var.get().strip(),
            )

            # dependency update handled directly against repo through update_row_by_key_name
            self.app.services.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                self.current_project_module_link_id,
                {
                    "DependencyModuleCode": dep_code,
                    "UpdatedOn": self.app.services.repo.load_workbook_safe and ""  # harmless placeholder replaced below
                }
            )
            # overwrite UpdatedOn cleanly
            self.app.services.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                self.current_project_module_link_id,
                {
                    "DependencyModuleCode": dep_code,
                    "UpdatedOn": __import__("storage").now_str(),
                }
            )

            self.refresh_page()
            self.set_status("Project module updated successfully.")

        except Exception as exc:
            show_error("Update Project Module Error", str(exc))

    def set_selected_project_module_dependency(self):
        if not self.require_workbook():
            return

        if not self.current_project_module_link_id:
            show_warning("No Selection", "Select a project module first.")
            return

        dep_code = self.project_module_dependency_var.get().strip()
        selected_module_code = self.direct_module_var.get().strip()

        if dep_code and dep_code == selected_module_code:
            show_warning("Invalid Dependency", "A module cannot depend on itself.")
            return

        try:
            self.app.services.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                self.current_project_module_link_id,
                {
                    "DependencyModuleCode": dep_code,
                    "UpdatedOn": __import__("storage").now_str(),
                }
            )
            self.refresh_page()
            self.set_status("Project module dependency updated.")
        except Exception as exc:
            show_error("Set Dependency Error", str(exc))

    def remove_selected_project_module(self):
        if not self.require_workbook():
            return

        if not self.current_project_module_link_id:
            show_warning("No Selection", "Select a project module first.")
            return

        if not tk.messagebox.askyesno("Remove Project Module", "Remove the selected project module?"):
            return

        try:
            self.app.services.repo.delete_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                self.current_project_module_link_id
            )
            # Rebuild project tasks so removed module task rows disappear
            self.app.services.projects.populate_project_tasks_from_modules(self._selected_project_code(), clear_existing=True)
            self.current_project_module_link_id = None
            self.refresh_page()
            self.set_status("Project module removed.")
        except Exception as exc:
            show_error("Remove Project Module Error", str(exc))

    # ========================================================
    # Project task execution
    # ========================================================

    def clear_project_task_editor(self):

        self.current_project_task_id = None
        self.project_task_select_var.set("")
        self.project_task_name_var.set("")
        self.project_task_module_var.set("")
        self.project_task_department_var.set("")
        self.project_task_hours_var.set("")
        self.project_task_stage_var.set("")
        self.project_task_status_var.set("Not Started")
        self.project_task_dependency_var.set("")
        self.project_task_assigned_to_var.set("")
        self.project_task_notes_var.set("")
        self.project_task_select_combo["values"] = []
        # self.current_project_task_id = None
        # self.project_task_name_var.set("")
        # self.project_task_module_var.set("")
        # self.project_task_department_var.set("")
        # self.project_task_hours_var.set("")
        # self.project_task_stage_var.set("")
        # self.project_task_status_var.set("Not Started")
        # self.project_task_dependency_var.set("")
        # self.project_task_assigned_to_var.set("")
        # self.project_task_notes_var.set("")

    def load_selected_project_task(self):
        sel = self.project_tasks_tree.selection()
        if not sel:
            return

        tags = self.project_tasks_tree.item(sel[0], "tags")
        if not tags:
            return

        task_id = tags[0]
        project_code = self._selected_project_code()
        tasks = self.app.services.projects.get_project_tasks(project_code)

        selected = None
        for t in tasks:
            if t.project_task_id == task_id:
                selected = t
                break

        if not selected:
            return

        self.current_project_task_id = selected.project_task_id
        self.project_task_name_var.set(selected.task_name)
        self.project_task_module_var.set(selected.module_code)
        self.project_task_department_var.set(selected.department)
        self.project_task_hours_var.set(f"{float(selected.estimated_hours or 0.0):.2f}")
        self.project_task_stage_var.set(selected.stage)
        self.project_task_status_var.set(selected.status or "Not Started")
        self.project_task_assigned_to_var.set(selected.assigned_to)
        self.project_task_notes_var.set(selected.notes)
        self.refresh_module_task_dropdown()
        self._load_project_task_into_editor(selected)

        dep_display = ""
        if norm_text(selected.dependency_task_id):
            for t in tasks:
                if t.project_task_id == selected.dependency_task_id:
                    dep_display = f"{t.module_code} | {t.task_name}"
                    break
        self.project_task_dependency_var.set(dep_display)

    def update_selected_project_task(self):
        if not self.require_workbook():
            return

        if not self.current_project_task_id:
            show_warning("No Selection", "Select a project task first.")
            return

        dep_id = self._selected_project_task_dependency_id()
        if dep_id and dep_id == self.current_project_task_id:
            show_warning("Invalid Dependency", "A task cannot depend on itself.")
            return

        try:
            # editable execution fields first
            self.app.services.projects.update_project_task_status(
                project_task_id=self.current_project_task_id,
                stage=self.project_task_stage_var.get().strip(),
                status=self.project_task_status_var.get().strip(),
                assigned_to=self.project_task_assigned_to_var.get().strip(),
                notes=self.project_task_notes_var.get().strip(),
            )

            # then editable descriptive fields
            self.app.services.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                self.current_project_task_id,
                {
                    "TaskName": self.project_task_name_var.get().strip(),
                    "ModuleCode": self.project_task_module_var.get().strip(),
                    "Department": self.project_task_department_var.get().strip(),
                    "EstimatedHours": Validators.parse_float(self.project_task_hours_var.get(), "Hours", default=0.0),
                    "DependencyTaskID": dep_id,
                    "UpdatedOn": __import__("storage").now_str(),
                }
            )

            self.refresh_page()
            self.set_status("Project task updated successfully.")

        except Exception as exc:
            show_error("Update Project Task Error", str(exc))

    def delete_selected_project_task(self):
        if not self.require_workbook():
            return

        if not self.current_project_task_id:
            show_warning("No Selection", "Select a project task first.")
            return

        if not tk.messagebox.askyesno("Delete Project Task", "Delete the selected project task?"):
            return

        try:
            self.app.services.projects.delete_project_task(self.current_project_task_id)
            self.clear_project_task_editor()
            self.refresh_page()
            self.set_status("Project task deleted successfully.")
        except Exception as exc:
            show_error("Delete Project Task Error", str(exc))

    # ========================================================
    # Project docs
    # ========================================================

    def add_project_document(self):
        if not self.require_workbook():
            return

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Save or select a project first.")
            return

        file_path = filedialog.askopenfilename(
            title="Select Project Document",
            filetypes=[
                ("Documents", "*.pdf *.dwg *.dxf *.step *.stp *.sldprt *.sldasm *.doc *.docx *.xls *.xlsx *.png *.jpg *.jpeg"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            self.app.services.projects.add_project_document(
                project_code=project_code,
                source_file_path=file_path,
                section_name=self.doc_section_var.get().strip(),
                doc_type=self.doc_type_var.get().strip() or "Other",
                instruction_text=self.doc_instruction_var.get().strip(),
                copy_file=True,
            )
            self.refresh_page()
            self.set_status("Project document added successfully.")
        except Exception as exc:
            show_error("Add Project Document Error", str(exc))

    def _get_selected_project_document_info(self):
        sel = self.project_document_tree.selection()
        if not sel:
            return None, None

        tags = self.project_document_tree.item(sel[0], "tags")
        vals = self.project_document_tree.item(sel[0], "values")

        doc_id = tags[0] if tags else ""
        file_path = vals[3] if vals and len(vals) > 3 else ""
        return doc_id, file_path

    def open_selected_project_document(self):
        doc_id, file_path = self._get_selected_project_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a project document first.")
            return

        try:
            file_path = self.app.services.projects.resolve_project_document_open_path(doc_id, str(file_path))
            open_file_with_default_app(str(file_path))
        except Exception as exc:
            show_error("Open Project Document Error", str(exc))

    def delete_selected_project_document(self):
        if not self.require_workbook():
            return

        doc_id, _file_path = self._get_selected_project_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a project document first.")
            return

        if not tk.messagebox.askyesno("Delete Project Document", "Delete the selected project document record?"):
            return

        try:
            self.app.services.projects.delete_project_document(doc_id, delete_file=False)
            self.refresh_page()
            self.set_status("Project document deleted successfully.")
        except Exception as exc:
            show_error("Delete Project Document Error", str(exc))

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

        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Save or select a project first.")
            return

        try:
            workorder_name = Validators.require_text(self.workorder_name_var.get(), "Work order name")

            self.app.services.projects.add_project_workorder(
                project_code=project_code,
                workorder_name=workorder_name,
                stage=self.workorder_stage_var.get().strip(),
                owner=self.workorder_owner_var.get().strip(),
                due_date=self.workorder_due_var.get().strip(),
                status=self.workorder_status_var.get().strip() or "Open",
                notes=self.workorder_notes_var.get().strip(),
            )

            self.clear_workorder_editor()
            self.refresh_page()
            self.set_status("Project work order added successfully.")
        except Exception as exc:
            show_error("Add Project Work Order Error", str(exc))

    def load_selected_workorder(self):
        sel = self.workorder_tree.selection()
        if not sel:
            return

        tags = self.workorder_tree.item(sel[0], "tags")
        if not tags:
            return

        workorder_id = tags[0]
        project_code = self._selected_project_code()
        workorders = self.app.services.projects.get_project_workorders(project_code)

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
            self.set_status("Project work order updated successfully.")
        except Exception as exc:
            show_error("Update Project Work Order Error", str(exc))

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
            self.set_status("Project work order deleted successfully.")
        except Exception as exc:
            show_error("Delete Project Work Order Error", str(exc))

    # ========================================================
    # PDF / email hooks
    # ========================================================

    def generate_project_pdf(self):
        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return

        if not hasattr(self.app, "reports"):
            show_warning("Not Ready", "Report service is not wired yet. We will connect it when reports.py and main.py are added.")
            return

        try:
            self.app.reports.generate_project_report_dialog(project_code)
        except Exception as exc:
            show_error("Generate Project PDF Error", str(exc))

    def email_project_pdf(self):
        project_code = self._selected_project_code()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return

        if not hasattr(self.app, "reports") or not hasattr(self.app, "mailer"):
            show_warning("Not Ready", "Mailer/report services are not wired yet. We will connect them in the final files.")
            return

        try:
            self.app.reports.email_project_report_dialog(project_code)
        except Exception as exc:
            show_error("Email Project PDF Error", str(exc))

# ============================================================
# PATCH: live orders wording + searchable combos + tracker tab + parts tab
# ============================================================
from datetime import datetime as _dt


def _erp_bind_filterable_combobox(combo, values_getter):
    def _set_values(filter_text=""):
        try:
            all_values = [str(v) for v in values_getter()]
            typed = str(filter_text or "").strip().lower()
            combo['values'] = [v for v in all_values if typed in v.lower()] if typed else all_values
        except Exception:
            pass

    def _apply(event=None):
        _set_values(combo.get())

    def _reset_dropdown(event=None):
        _set_values("")

    combo.configure(postcommand=_reset_dropdown)
    combo.bind('<KeyRelease>', _apply, add='+')
    combo.bind('<Button-1>', _reset_dropdown, add='+')
    combo.bind('<FocusIn>', _reset_dropdown, add='+')
    return combo


def _patched_project_topbar(self, parent):
    top = ttk.Frame(parent)
    top.pack(fill='x', pady=(0, 8))
    left = ttk.Frame(top)
    left.pack(side='left', fill='x', expand=True)
    ttk.Button(left, text='← Back to Dashboard', command=lambda: self.show_page('home')).pack(side='left', padx=2)
    ttk.Label(left, text='Open Live Order').pack(side='left', padx=(12, 4))
    self.project_select_top_combo = ttk.Combobox(left, textvariable=self.project_select_top_var, state='normal', width=60)
    self.project_select_top_combo.pack(side='left', padx=4)
    self.project_select_top_combo.bind('<<ComboboxSelected>>', lambda e: self.open_selected_project_from_dropdown())
    ttk.Button(left, text='Open', command=self.open_selected_project_from_dropdown).pack(side='left', padx=4)
    _erp_bind_filterable_combobox(self.project_select_top_combo, lambda: getattr(self.project_select_top_combo, '_all_values', self.project_select_top_combo.cget('values')))
    right = ttk.Frame(top)
    right.pack(side='right')
    ttk.Button(right, text='Refresh', command=self.refresh_page).pack(side='right', padx=2)
    ttk.Button(right, text='Email Order PDF', command=self.email_project_pdf).pack(side='right', padx=2)
    ttk.Button(right, text='Generate Order PDF', command=self.generate_project_pdf).pack(side='right', padx=2)


def _patched_refresh_project_selector(self):
    if not self.app.workbook_manager.has_workbook():
        self.project_select_top_combo['values'] = []
        self.project_select_top_combo._all_values = []
        return
    records = self.app.services.projects.list_projects()
    values = [f"{p.project_code} | {p.quote_ref} | {p.project_name}" for p in records]
    self.project_select_top_combo['values'] = values
    self.project_select_top_combo._all_values = values
    selected_code = getattr(self.app, 'selected_project_code', '').strip()
    if selected_code:
        for p in records:
            if p.project_code == selected_code:
                self.project_select_top_var.set(f"{p.project_code} | {p.quote_ref} | {p.project_name}")
                break


def _patched_build_right_panel(self, parent):
    tabs = ttk.Notebook(parent)
    tabs.pack(fill='both', expand=True)
    self.modules_tab = ttk.Frame(tabs)
    self.parts_tab = ttk.Frame(tabs)
    self.all_parts_tab = ttk.Frame(tabs)
    self.tasks_tab = ttk.Frame(tabs)
    self.tracker_tab = ttk.Frame(tabs)
    self.docs_tab = ttk.Frame(tabs)
    self.workorders_tab = ttk.Frame(tabs)
    self.summary_tab = ttk.Frame(tabs)
    tabs.add(self.modules_tab, text='Order Assemblies')
    tabs.add(self.parts_tab, text='Order Parts')
    tabs.add(self.all_parts_tab, text='All Parts')
    tabs.add(self.tasks_tab, text='Department Tasks')
    tabs.add(self.tracker_tab, text='Live Order Tracker')
    tabs.add(self.docs_tab, text='Order Docs')
    tabs.add(self.workorders_tab, text='Job Cards')
    tabs.add(self.summary_tab, text='Summary')
    self._build_modules_tab(self.modules_tab)
    self._build_order_parts_tab(self.parts_tab)
    self._build_all_order_parts_tab(self.all_parts_tab)
    self._build_tasks_tab(self.tasks_tab)
    self._build_order_tracker_tab(self.tracker_tab)
    self._build_docs_tab(self.docs_tab)
    self._build_workorders_tab(self.workorders_tab)
    self._build_summary_tab(self.summary_tab)


def _patched_build_order_parts_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    cols = ('PartName','PartNumber','Qty','SOH','Supplier','LeadTime','Notes')
    self.order_parts_tree = ttk.Treeview(wrap, columns=cols, show='headings')
    for col, width, title in [
        ('PartName',220,'Part Name'),('PartNumber',140,'Part Number'),('Qty',70,'Qty'),('SOH',70,'SOH'),('Supplier',160,'Supplier'),('LeadTime',90,'Lead Time'),('Notes',260,'Notes')]:
        self.order_parts_tree.heading(col, text=title)
        self.order_parts_tree.column(col, width=width, anchor='w')
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.order_parts_tree.xview)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.order_parts_tree.yview)
    self.order_parts_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.order_parts_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


def _build_all_order_parts_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    cols = ('Source', 'Assembly', 'PartName', 'PartNumber', 'Qty', 'SOH', 'Supplier', 'LeadTime', 'Notes')
    self.all_order_parts_tree = ttk.Treeview(wrap, columns=cols, show='headings')
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
        self.all_order_parts_tree.heading(col, text=title)
        self.all_order_parts_tree.column(col, width=width, anchor='w')
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.all_order_parts_tree.xview)
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.all_order_parts_tree.yview)
    self.all_order_parts_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.all_order_parts_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')


def _patched_build_order_tracker_tab(self, parent):
    wrap = ttk.Frame(parent, padding=8)
    wrap.pack(fill='both', expand=True)
    cols = ('Assembly','Automation','AssemblyDept','Fabrication','Procurement','Software','Mechanical','Logistics','Overall','DueDate')
    self.order_tracker_tree = ttk.Treeview(wrap, columns=cols, show='headings')
    titles = {
        'Assembly':'Assembly', 'Automation':'Automation', 'AssemblyDept':'Assembly', 'Fabrication':'Fabrication',
        'Procurement':'Procurement', 'Software':'Software', 'Mechanical':'Mechanical', 'Logistics':'Logistics/Quote',
        'Overall':'Overall', 'DueDate':'Due Date'
    }
    widths = {'Assembly':220,'Automation':110,'AssemblyDept':110,'Fabrication':110,'Procurement':120,'Software':100,'Mechanical':110,'Logistics':120,'Overall':90,'DueDate':100}
    for c in cols:
        self.order_tracker_tree.heading(c, text=titles[c])
        self.order_tracker_tree.column(c, width=widths[c], anchor='center' if c!='Assembly' else 'w')
    ysb = ttk.Scrollbar(wrap, orient='vertical', command=self.order_tracker_tree.yview)
    xsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.order_tracker_tree.xview)
    self.order_tracker_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    self.order_tracker_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    xsb.pack(side='bottom', fill='x')
    self.order_tracker_tree.tag_configure('green', background='#d9f2d9')
    self.order_tracker_tree.tag_configure('orange', background='#ffe5bf')
    self.order_tracker_tree.tag_configure('red', background='#ffd6d6')
    self.order_tracker_tree.tag_configure('white', background='white')


def _patched_build_project_browser_card(self, parent):
    card = ttk.LabelFrame(parent, text='Live Order Browser', style='Card.TLabelframe', padding=12)
    card.pack(fill='x', pady=6)
    tabs = ttk.Notebook(card)
    tabs.pack(fill='both', expand=True)
    live_tab = ttk.Frame(tabs)
    completed_tab = ttk.Frame(tabs)
    preview_tab = ttk.Frame(tabs)
    tabs.add(live_tab, text='Live Orders')
    tabs.add(completed_tab, text='Completed Orders')
    tabs.add(preview_tab, text='Task Preview')
    live_wrap = ttk.Frame(live_tab, padding=4)
    live_wrap.pack(fill='both', expand=True)
    self.live_projects_tree = ttk.Treeview(live_wrap, columns=('ProjectName','Client','Status','DueDate'), show='headings', height=6)
    for col, width, title in [('ProjectName',220,'Order Name'),('Client',160,'Client'),('Status',110,'Status'),('DueDate',100,'Due Date')]:
        self.live_projects_tree.heading(col, text=title)
        self.live_projects_tree.column(col, width=width, anchor='w')
    ysb = ttk.Scrollbar(live_wrap, orient='vertical', command=self.live_projects_tree.yview)
    self.live_projects_tree.configure(yscrollcommand=ysb.set)
    self.live_projects_tree.pack(side='left', fill='both', expand=True)
    ysb.pack(side='right', fill='y')
    self.live_projects_tree.bind('<Double-1>', lambda e: self.load_project_from_browser())
    self.live_projects_tree.bind('<ButtonRelease-1>', lambda e: self.load_project_from_browser())

    completed_wrap = ttk.Frame(completed_tab, padding=4)
    completed_wrap.pack(fill='both', expand=True)
    self.completed_projects_tree = ttk.Treeview(completed_wrap, columns=('ProjectName','Client','CompletedOn'), show='headings', height=5)
    for col, width, title in [('ProjectName',240,'Completed Order'),('Client',160,'Client'),('CompletedOn',140,'Completed / Updated')]:
        self.completed_projects_tree.heading(col, text=title)
        self.completed_projects_tree.column(col, width=width, anchor='w')
    y2 = ttk.Scrollbar(completed_wrap, orient='vertical', command=self.completed_projects_tree.yview)
    self.completed_projects_tree.configure(yscrollcommand=y2.set)
    self.completed_projects_tree.pack(side='left', fill='both', expand=True)
    y2.pack(side='right', fill='y')

    prev = ttk.Frame(preview_tab, padding=6)
    prev.pack(fill='both', expand=True)
    self.project_preview_text = make_readonly_text(prev, height=10)
    self.project_preview_text.pack(fill='both', expand=True)


def _patched_refresh_project_browser(self):
    if not hasattr(self, 'live_projects_tree'):
        return
    treeview_clear(self.live_projects_tree)
    treeview_clear(self.completed_projects_tree)
    projects = self.app.services.projects.list_projects()
    live = [p for p in projects if norm_text(p.status).lower() != 'completed']
    completed = [p for p in projects if norm_text(p.status).lower() == 'completed']
    for p in live:
        self.live_projects_tree.insert('', 'end', values=(p.project_name, p.client_name, p.status, p.due_date), tags=(p.project_code,))
    for p in completed:
        self.completed_projects_tree.insert('', 'end', values=(p.project_name, p.client_name, p.updated_on), tags=(p.project_code,))


def _patched_refresh_combos(self):
    products = self.app.services.products.list_products()
    product_values = [p.product_code for p in products]
    set_combobox_values(self.linked_product_combo, product_values, keep_current=True)
    self.linked_product_combo._all_values = product_values
    modules = self.app.services.modules.list_modules()
    module_values = [m.module_code for m in modules]
    set_combobox_values(self.direct_module_combo, module_values, keep_current=True)
    self.direct_module_combo._all_values = module_values
    _erp_bind_filterable_combobox(self.linked_product_combo, lambda: getattr(self.linked_product_combo, '_all_values', []))
    _erp_bind_filterable_combobox(self.direct_module_combo, lambda: getattr(self.direct_module_combo, '_all_values', []))
    _erp_bind_filterable_combobox(self.project_module_dependency_combo, lambda: getattr(self.project_module_dependency_combo, '_all_values', []))
    _erp_bind_filterable_combobox(self.project_task_module_combo, lambda: getattr(self.project_task_module_combo, '_all_values', []))
    _erp_bind_filterable_combobox(self.project_task_dependency_combo, lambda: getattr(self.project_task_dependency_combo, '_all_values', []))


def _patched_refresh_project_module_dependency_combo(self, module_links):
    vals = [m.module_code for m in module_links]
    set_combobox_values(self.project_module_dependency_combo, vals, keep_current=True)
    self.project_module_dependency_combo._all_values = vals


def _patched_refresh_project_task_dependency_combo(self, tasks):
    display_values = [f"{t.project_task_id} | {t.module_code} | {t.task_name}" for t in tasks]
    set_combobox_values(self.project_task_dependency_combo, display_values, keep_current=True)
    self.project_task_dependency_combo._all_values = display_values


def _patched_refresh_project_task_module_dropdown(self, module_links):
    module_codes = [m.module_code for m in module_links]
    set_combobox_values(self.project_task_module_combo, module_codes, keep_current=True)
    self.project_task_module_combo._all_values = module_codes


def _patched_load_order_parts(self, parts):
    if not hasattr(self, 'order_parts_tree'):
        return
    treeview_clear(self.order_parts_tree)
    for p in parts or []:
        self.order_parts_tree.insert('', 'end', values=(p.component_name, p.part_number, p.qty, p.soh_qty, p.preferred_supplier, p.lead_time_days, p.notes), tags=(p.component_id,))


def _load_all_order_parts(self, bundle):
    if not hasattr(self, 'all_order_parts_tree'):
        return
    treeview_clear(self.all_order_parts_tree)

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
            self.all_order_parts_tree.insert(
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

    for part in (getattr(bundle, 'project_parts', []) or []):
        try:
            qty = float(getattr(part, 'qty', 0) or 0)
            soh_qty = float(getattr(part, 'soh_qty', 0) or 0)
        except Exception:
            qty = 0.0
            soh_qty = 0.0
        self.all_order_parts_tree.insert(
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


def _status_from_workorder(status, due_date):
    st = norm_text(status).lower()
    if st == 'completed':
        return 'green'
    due = None
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            due = _dt.strptime(norm_text(due_date), fmt).date()
            break
        except Exception:
            pass
    if not due:
        return 'white'
    today = _dt.today().date()
    if due < today:
        return 'red'
    if (due - today).days <= 1:
        return 'orange'
    return 'green'


def _patched_load_order_tracker(self, bundle):
    if not hasattr(self, 'order_tracker_tree'):
        return
    treeview_clear(self.order_tracker_tree)
    workorders = bundle.workorders or []
    dept_map = {}
    for w in workorders:
        if not norm_text(w.notes).startswith('AUTO_TRACKER'):
            continue
        parts = [p.strip() for p in norm_text(w.notes).split('|')]
        module_code = ''
        for part in parts:
            if part.startswith('MODULE='):
                module_code = part.split('=',1)[1].strip()
                break
        if not module_code:
            module_code = norm_text(w.workorder_name).split('|')[0].strip()
        dept = norm_text(w.stage or w.owner)
        dept_map.setdefault(module_code, {})[dept] = w
    for link in bundle.module_links or []:
        row = {'Automation':'','AssemblyDept':'','Fabrication':'','Procurement':'','Software':'','Mechanical':'','Logistics':'','Overall':'Open','DueDate':''}
        overall_tag = 'white'
        module_due = ''
        seen_statuses = []
        for key, dept_names in {
            'Automation':['automation','electrical'],
            'AssemblyDept':['assembly','operations'],
            'Fabrication':['fabrication'],
            'Procurement':['procurement'],
            'Software':['software'],
            'Mechanical':['mechanical','ga drawing'],
            'Logistics':['logistics','quote/order','hr'],
        }.items():
            display = '-'
            tag = 'white'
            for dept_name, wo in dept_map.get(link.module_code, {}).items():
                dlow = dept_name.lower()
                if any(token in dlow for token in dept_names):
                    display = norm_text(wo.status) or 'Open'
                    tag = _status_from_workorder(wo.status, wo.due_date)
                    if not module_due and wo.due_date:
                        module_due = wo.due_date
                    seen_statuses.append(display.lower())
                    break
            row[key] = display
            overall_tag = 'red' if tag == 'red' else ('orange' if tag == 'orange' and overall_tag != 'red' else overall_tag)
        if seen_statuses and all(s == 'completed' for s in seen_statuses):
            row['Overall'] = 'Completed'
            overall_tag = 'green'
        elif overall_tag == 'orange':
            row['Overall'] = 'Due Soon'
        elif overall_tag == 'red':
            row['Overall'] = 'Overdue'
        else:
            row['Overall'] = 'On Track'
        row['DueDate'] = module_due
        self.order_tracker_tree.insert('', 'end', values=(link.module_code, row['Automation'], row['AssemblyDept'], row['Fabrication'], row['Procurement'], row['Software'], row['Mechanical'], row['Logistics'], row['Overall'], row['DueDate']), tags=(overall_tag,))


def _patched_refresh_page(self):
    if not self.require_workbook():
        return
    self.refresh_project_selector()
    project_code = getattr(self.app, 'selected_project_code', '').strip()
    if not project_code:
        self._clear_all_views()
        self.project_summary_label.config(text='No live order selected')
        return
    try:
        bundle = self.app.services.projects.get_project_bundle(project_code)
        if not bundle.project:
            self._clear_all_views()
            self.project_summary_label.config(text='Order not found')
            return
        self._load_project_into_form(bundle.project)
        self._load_project_modules(bundle.module_links or [])
        self._load_project_tasks(bundle.project_tasks or [])
        self._load_order_parts(getattr(bundle, 'project_parts', []))
        self._load_project_documents(bundle.project_documents or [])
        self._load_workorders(bundle.workorders or [])
        try:
            self._load_summary(bundle)
        except Exception:
            try:
                project = bundle.project
                set_text_readonly(
                    self.summary_text,
                    "\n".join([
                        f"Order Code: {getattr(project, 'project_code', '')}",
                        f"Quote Ref: {getattr(project, 'quote_ref', '')}",
                        f"Order Name: {getattr(project, 'project_name', '')}",
                        f"Client Name: {getattr(project, 'client_name', '')}",
                        f"Location: {getattr(project, 'location', '')}",
                        f"Description: {getattr(project, 'description', '')}",
                        f"Linked Product: {getattr(project, 'linked_product_code', '')}",
                        f"Status: {getattr(project, 'status', '')}",
                        f"Start Date: {getattr(project, 'start_date', '')}",
                        f"Due Date: {getattr(project, 'due_date', '')}",
                        "",
                        f"Assemblies Loaded: {len(getattr(bundle, 'module_links', []) or [])}",
                        f"Direct Parts Loaded: {len(getattr(bundle, 'project_parts', []) or [])}",
                        f"Department Tasks: {len(getattr(bundle, 'project_tasks', []) or [])}",
                        f"Documents: {len(getattr(bundle, 'project_documents', []) or [])}",
                        f"Job Cards: {len(getattr(bundle, 'workorders', []) or [])}",
                        f"Total Labour Hours: {float(getattr(bundle, 'total_hours', 0.0) or 0.0):.2f}",
                    ])
                )
            except Exception:
                pass
        try:
            self._load_all_order_parts(bundle)
        except Exception:
            pass
        self._load_order_tracker(bundle)
        if hasattr(self, 'live_projects_tree'):
            self.refresh_project_browser()
        self._refresh_products_combo()
        self._refresh_modules_combo()
        self._refresh_project_module_dependency_combo(bundle.module_links or [])
        self._refresh_project_task_dependency_combo(bundle.project_tasks or [])
        self.refresh_project_task_module_dropdown(bundle.module_links or [])
        self.refresh_module_task_dropdown()
        try:
            if not get_text_value(self.summary_text):
                project = bundle.project
                set_text_readonly(
                    self.summary_text,
                    "\n".join([
                        f"Order Code: {getattr(project, 'project_code', '')}",
                        f"Quote Ref: {getattr(project, 'quote_ref', '')}",
                        f"Order Name: {getattr(project, 'project_name', '')}",
                        f"Status: {getattr(project, 'status', '')}",
                    ])
                )
        except Exception:
            pass
        self.project_summary_label.config(text=f"📦 {bundle.project.project_name} | Quote: {bundle.project.quote_ref} | Client: {bundle.project.client_name} | Order Code: {bundle.project.project_code}")
        self.set_status(f"Loaded live order: {bundle.project.project_code}")
    except Exception as exc:
        show_error('Project Refresh Error', str(exc))


def _patched_load_summary(self, bundle):
    project = bundle.project
    module_links = bundle.module_links or []
    project_tasks = bundle.project_tasks or []
    project_docs = bundle.project_documents or []
    workorders = bundle.workorders or []
    project_parts = getattr(bundle, 'project_parts', []) or []
    blockers = []
    try:
        blockers = self.app.services.scheduler.get_open_blockers_for_project(project.project_code)
    except Exception:
        blockers = []
    lines = [
        f"Order Code: {project.project_code}",
        f"Quote Ref: {project.quote_ref}",
        f"Order Name: {project.project_name}",
        f"Client Name: {project.client_name}",
        f"Location: {project.location}",
        f"Description: {project.description}",
        f"Linked Product: {project.linked_product_code or '-'}",
        f"Status: {project.status}",
        f"Start Date: {project.start_date or '-'}",
        f"Due Date: {project.due_date or '-'}",
        '',
        f"Order Assemblies: {len(module_links)}",
        f"Order Parts: {len(project_parts)}",
        f"Department Tasks: {len(project_tasks)}",
        f"Order Documents: {len(project_docs)}",
        f"Job Cards: {len(workorders)}",
        f"Aggregated Labour Hours: {float(bundle.total_hours or 0.0):.2f}",
        '',
        'Assembly Execution Status:',
    ]
    if not module_links:
        lines.append(' - No assemblies loaded.')
    else:
        for link in module_links:
            dep_text = f" | Depends on: {link.dependency_module_code}" if norm_text(link.dependency_module_code) else ''
            lines.append(f" {link.module_order:02d}. {link.module_code} | {link.source_type} | Qty {link.module_qty} | Stage {link.stage} | Status {link.status}{dep_text}")
    lines.append('')
    lines.append('Direct / Product Parts:')
    if not project_parts:
        lines.append(' - None')
    else:
        for p in project_parts[:20]:
            lines.append(f" - {p.component_name} | PN {p.part_number or '-'} | Qty {p.qty} | Supplier {p.preferred_supplier or '-'}")
        if len(project_parts) > 20:
            lines.append(f" - ... plus {len(project_parts)-20} more")
    lines.append('')
    lines.append('Open Blockers:')
    if not blockers:
        lines.append(' - None')
    else:
        for b in blockers:
            if b.get('type') == 'TASK':
                lines.append(f" - TASK: {b.get('task_name')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})")
            else:
                lines.append(f" - ASSEMBLY: {b.get('module_code')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})")
    set_text_readonly(self.summary_text, '\n'.join(lines))


ProjectPage._build_topbar = _patched_project_topbar
ProjectPage.refresh_project_selector = _patched_refresh_project_selector
ProjectPage._build_right_panel = _patched_build_right_panel
ProjectPage._build_order_parts_tab = _patched_build_order_parts_tab
ProjectPage._build_order_tracker_tab = _patched_build_order_tracker_tab
ProjectPage._build_project_browser_card = _patched_build_project_browser_card
ProjectPage.refresh_project_browser = _patched_refresh_project_browser
ProjectPage._refresh_products_combo = _patched_refresh_combos
ProjectPage._refresh_modules_combo = _patched_refresh_combos
ProjectPage._refresh_project_module_dependency_combo = _patched_refresh_project_module_dependency_combo
ProjectPage._refresh_project_task_dependency_combo = _patched_refresh_project_task_dependency_combo
ProjectPage._refresh_project_task_module_combo = _patched_refresh_project_task_module_dropdown
ProjectPage._load_order_parts = _patched_load_order_parts
ProjectPage._load_order_tracker = _patched_load_order_tracker
ProjectPage.refresh_page = _patched_refresh_page
ProjectPage._load_summary = _patched_load_summary


# ============================================================
# V3 PATCH: live order details dropdowns + scrollable summary + cost summary
# ============================================================

def _v3_collect_quote_refs_project(page):
    refs = []
    for getter in (
        lambda: page.app.services.products.list_products(),
        lambda: page.app.services.modules.list_modules(),
        lambda: page.app.services.projects.list_projects(),
    ):
        try:
            for rec in getter():
                q = norm_text(getattr(rec, 'quote_ref', ''))
                if q:
                    refs.append(q)
        except Exception:
            pass
    return sorted(set(refs))


def _v3_project_details_card(self, parent):
    card = ttk.LabelFrame(parent, text="Live Order Details", style="Card.TLabelframe", padding=12)
    card.pack(fill="x", pady=6)

    ttk.Label(card, text="Order Code").grid(row=0, column=0, sticky="w", pady=4)
    self.project_code_combo = ttk.Combobox(card, textvariable=self.project_code_var, state="normal")
    self.project_code_combo.grid(row=0, column=1, sticky="ew", padx=6)
    self.project_code_combo.bind("<<ComboboxSelected>>", lambda e: self._open_order_from_details_code())

    ttk.Label(card, text="Quote Ref").grid(row=1, column=0, sticky="w", pady=4)
    self.quote_ref_combo = ttk.Combobox(card, textvariable=self.quote_ref_var, state="normal")
    self.quote_ref_combo.grid(row=1, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Order Name").grid(row=2, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.project_name_var).grid(row=2, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Client Name").grid(row=3, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.client_name_var).grid(row=3, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Location").grid(row=4, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.location_var).grid(row=4, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Description").grid(row=5, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.description_var).grid(row=5, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Linked Product").grid(row=6, column=0, sticky="w", pady=4)
    self.linked_product_combo = ttk.Combobox(card, textvariable=self.linked_product_var, state="normal")
    self.linked_product_combo.grid(row=6, column=1, sticky="ew", padx=6)

    prod_btn_row = ttk.Frame(card)
    prod_btn_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 2))
    ttk.Button(prod_btn_row, text="Attach Product", command=self.attach_product_to_project).pack(side="left", fill="x", expand=True, padx=2)
    ttk.Button(prod_btn_row, text="Rebuild Assemblies", command=self.rebuild_modules_from_product).pack(side="left", fill="x", expand=True, padx=2)
    ttk.Button(prod_btn_row, text="Rebuild Tasks", command=self.rebuild_project_tasks).pack(side="left", fill="x", expand=True, padx=2)

    ttk.Label(card, text="Status").grid(row=8, column=0, sticky="w", pady=4)
    ttk.Combobox(card, textvariable=self.status_var_local, values=AppConfig.PROJECT_STATUSES, state="readonly").grid(row=8, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Start Date").grid(row=9, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.start_date_var).grid(row=9, column=1, sticky="ew", padx=6)

    ttk.Label(card, text="Due Date").grid(row=10, column=0, sticky="w", pady=4)
    ttk.Entry(card, textvariable=self.due_date_var).grid(row=10, column=1, sticky="ew", padx=6)

    btn_row = ttk.Frame(card)
    btn_row.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(10, 0))
    ttk.Button(btn_row, text="Save Order", command=self.save_project).pack(side="left", fill="x", expand=True, padx=2)
    ttk.Button(btn_row, text="Mark Completed", command=self.mark_project_completed).pack(side="left", fill="x", expand=True, padx=2)
    ttk.Button(btn_row, text="Delete Order", command=self.delete_project).pack(side="left", fill="x", expand=True, padx=2)

    card.columnconfigure(1, weight=1)
    _erp_bind_filterable_combobox(self.project_code_combo, lambda: getattr(self.project_code_combo, '_all_values', self.project_code_combo.cget('values')))
    _erp_bind_filterable_combobox(self.quote_ref_combo, lambda: getattr(self.quote_ref_combo, '_all_values', self.quote_ref_combo.cget('values')))
    _erp_bind_filterable_combobox(self.linked_product_combo, lambda: getattr(self.linked_product_combo, '_all_values', self.linked_product_combo.cget('values')))


def _v3_refresh_combos(self):
    _patched_refresh_combos(self)
    try:
        projects = self.app.services.projects.list_projects() if self.app.workbook_manager.has_workbook() else []
        values = [f"{p.project_code} | {p.project_name}" for p in projects]
        if hasattr(self, 'project_code_combo'):
            self.project_code_combo['values'] = values
            self.project_code_combo._all_values = values
        qvals = _v3_collect_quote_refs_project(self)
        if hasattr(self, 'quote_ref_combo'):
            self.quote_ref_combo['values'] = qvals
            self.quote_ref_combo._all_values = qvals
        if hasattr(self, 'linked_product_combo'):
            self.linked_product_combo._all_values = list(self.linked_product_combo.cget('values'))
    except Exception:
        pass


def _v3_open_order_from_details_code(self):
    text = norm_text(self.project_code_var.get())
    if not text:
        return
    code = text.split('|')[0].strip()
    self.app.set_selected_project(code)
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


def _v3_lookup_part_price(self, comp):
    # Try dedicated Parts sheet first; fall back to optional UnitPrice on Components row if present.
    candidates = [norm_text(getattr(comp, 'part_number', '')), norm_text(getattr(comp, 'component_name', ''))]
    try:
        parts_rows = self.app.repo.list_dicts('Parts')
    except Exception:
        parts_rows = []
    for row in parts_rows:
        pn = norm_text(row.get('PartNumber'))
        name = norm_text(row.get('PartName'))
        if any(c and c == pn for c in candidates) or any(c and c == name for c in candidates):
            try:
                return float(row.get('UnitPrice') or 0)
            except Exception:
                return 0.0
    try:
        comp_rows = self.app.repo.list_dicts('Components')
        for row in comp_rows:
            if norm_text(row.get('ComponentID')) == norm_text(getattr(comp, 'component_id', '')):
                return float(row.get('UnitPrice') or 0)
    except Exception:
        pass
    return 0.0


def _v3_load_summary(self, bundle):
    project = bundle.project
    module_links = bundle.module_links or []
    project_tasks = bundle.project_tasks or []
    project_docs = bundle.project_documents or []
    workorders = bundle.workorders or []
    project_parts = getattr(bundle, 'project_parts', []) or []
    blockers = []
    try:
        blockers = self.app.services.scheduler.get_open_blockers_for_project(project.project_code)
    except Exception:
        blockers = []
    total_parts_cost = 0.0
    for p in project_parts:
        try:
            total_parts_cost += float(getattr(p, 'qty', 0) or 0) * float(_v3_lookup_part_price(self, p) or 0)
        except Exception:
            pass
    total_labour = float(bundle.total_hours or 0.0)
    lines = [
        f"Order Code: {project.project_code}",
        f"Quote Ref: {project.quote_ref}",
        f"Order Name: {project.project_name}",
        f"Client Name: {project.client_name}",
        f"Location: {project.location}",
        f"Description: {project.description}",
        f"Linked Product: {project.linked_product_code or '-'}",
        f"Status: {project.status}",
        f"Start Date: {project.start_date or '-'}",
        f"Due Date: {project.due_date or '-'}",
        '',
        f"Assemblies Loaded: {len(module_links)}",
        f"Direct Parts Loaded: {len(project_parts)}",
        f"Department Tasks: {len(project_tasks)}",
        f"Documents: {len(project_docs)}",
        f"Job Cards: {len(workorders)}",
        f"Total Labour Hours: {total_labour:.2f}",
        f"Estimated Parts Cost: ${total_parts_cost:,.2f}",
        '',
        'Order Visibility Summary:',
    ]
    if module_links:
        for link in module_links:
            dep_text = f" | Depends on: {link.dependency_module_code}" if norm_text(link.dependency_module_code) else ''
            lines.append(f" {link.module_order:02d}. {link.module_code} | Qty {link.module_qty} | Stage {link.stage} | Status {link.status}{dep_text}")
    else:
        lines.append(' - No assemblies loaded.')
    lines += ['', 'Direct Parts:']
    if project_parts:
        for p in project_parts[:30]:
            unit = _v3_lookup_part_price(self, p)
            ext = (float(getattr(p, 'qty', 0) or 0) * float(unit or 0))
            lines.append(f" - {p.component_name} | PN {p.part_number or '-'} | Qty {p.qty} | Unit ${unit:,.2f} | Ext ${ext:,.2f}")
        if len(project_parts) > 30:
            lines.append(f" - ... plus {len(project_parts)-30} more")
    else:
        lines.append(' - None')
    lines += ['', 'Open Blockers:']
    if blockers:
        for b in blockers:
            if b.get('type') == 'TASK':
                lines.append(f" - TASK: {b.get('task_name')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})")
            else:
                lines.append(f" - ASSEMBLY: {b.get('module_code')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})")
    else:
        lines.append(' - None')
    set_text_readonly(self.summary_text, '\n'.join(lines))


ProjectPage._build_project_details_card = _v3_project_details_card
ProjectPage._refresh_products_combo = _v3_refresh_combos
ProjectPage._refresh_modules_combo = _v3_refresh_combos
ProjectPage._open_order_from_details_code = _v3_open_order_from_details_code
ProjectPage._build_summary_tab = _v3_build_summary_tab
ProjectPage._load_summary = _v3_load_summary
ProjectPage._build_all_order_parts_tab = _build_all_order_parts_tab
ProjectPage._load_all_order_parts = _load_all_order_parts


# Import-safe aliases
LiveOrderPage = ProjectPage
OrderPage = ProjectPage
__all__ = ["ProjectPage", "LiveOrderPage", "OrderPage"]
