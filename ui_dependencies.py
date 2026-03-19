# ============================================================
# ui_dependencies.py
# Dependency visibility / blocker analysis page for Liquimech ERP
# ============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional

from app_config import AppConfig
from ui_common import (
    BasePage,
    set_combobox_values,
    treeview_clear,
    show_warning,
    show_error,
    make_readonly_text,
    set_text_readonly,
)


def norm_text(value) -> str:
    return str(value or "").strip()


class DependenciesPage(BasePage):
    PAGE_NAME = "dependencies"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        self.project_select_var = tk.StringVar()
        self.blocker_type_var = tk.StringVar(value="All")

        self._build_ui()

    # ========================================================
    # UI
    # ========================================================

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        self._build_topbar(wrapper)

        self.page_summary_label = ttk.Label(
            wrapper,
            text="Dependency visibility and blocker analysis",
            style="Title.TLabel"
        )
        self.page_summary_label.pack(anchor="w", pady=(0, 10))

        filters = ttk.LabelFrame(wrapper, text="Filters", style="Card.TLabelframe", padding=12)
        filters.pack(fill="x", pady=(0, 10))

        ttk.Label(filters, text="Project").grid(row=0, column=0, sticky="w", pady=4)
        self.project_combo = ttk.Combobox(filters, textvariable=self.project_select_var, state="readonly")
        self.project_combo.grid(row=0, column=1, sticky="ew", padx=6)
        self.project_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_page())

        ttk.Label(filters, text="Show").grid(row=0, column=2, sticky="w", pady=4)
        self.blocker_type_combo = ttk.Combobox(
            filters,
            textvariable=self.blocker_type_var,
            values=["All", "MODULE", "TASK"],
            state="readonly"
        )
        self.blocker_type_combo.grid(row=0, column=3, sticky="ew", padx=6)
        self.blocker_type_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_page())

        ttk.Button(filters, text="Open Selected Project", command=self.open_selected_project).grid(
            row=0, column=4, sticky="ew", padx=6
        )

        filters.columnconfigure(1, weight=1)
        filters.columnconfigure(3, weight=1)

        metrics = ttk.Frame(wrapper)
        metrics.pack(fill="x", pady=(0, 10))

        self.metric_module_links = self._make_metric_card(metrics, "Module Dependencies", "0")
        self.metric_module_links.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_task_links = self._make_metric_card(metrics, "Task Dependencies", "0")
        self.metric_task_links.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_open_blockers = self._make_metric_card(metrics, "Open Blockers", "0")
        self.metric_open_blockers.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_ready_items = self._make_metric_card(metrics, "Ready Items", "0")
        self.metric_ready_items.pack(side="left", fill="x", expand=True, padx=4)

        tabs = ttk.Notebook(wrapper)
        tabs.pack(fill="both", expand=True)

        self.module_dep_tab = ttk.Frame(tabs)
        self.task_dep_tab = ttk.Frame(tabs)
        self.blockers_tab = ttk.Frame(tabs)
        self.summary_tab = ttk.Frame(tabs)

        tabs.add(self.module_dep_tab, text="Module Dependencies")
        tabs.add(self.task_dep_tab, text="Task Dependencies")
        tabs.add(self.blockers_tab, text="Blockers")
        tabs.add(self.summary_tab, text="Summary")

        self._build_module_dependency_tab(self.module_dep_tab)
        self._build_task_dependency_tab(self.task_dep_tab)
        self._build_blockers_tab(self.blockers_tab)
        self._build_summary_tab(self.summary_tab)

    def _build_topbar(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 8))

        ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)
        ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)

    def _make_metric_card(self, parent, title: str, default_value: str):
        card = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe", padding=12)
        value_var = tk.StringVar(value=default_value)
        ttk.Label(card, textvariable=value_var, style="Hero.TLabel").pack(anchor="w")
        card.value_var = value_var
        return card

    def _build_module_dependency_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("ModuleCode", "DependsOn", "ModuleStatus", "DependencyStatus", "Stage", "SourceType")
        self.module_dep_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("ModuleCode", 220),
            ("DependsOn", 220),
            ("ModuleStatus", 140),
            ("DependencyStatus", 160),
            ("Stage", 130),
            ("SourceType", 120),
        ]:
            self.module_dep_tree.heading(col, text=col)
            self.module_dep_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.module_dep_tree.yview)
        self.module_dep_tree.configure(yscrollcommand=sb.set)

        self.module_dep_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_task_dependency_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("ModuleCode", "TaskName", "DependsOnTask", "TaskStatus", "DependencyStatus", "AssignedTo")
        self.task_dep_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("ModuleCode", 180),
            ("TaskName", 240),
            ("DependsOnTask", 260),
            ("TaskStatus", 130),
            ("DependencyStatus", 150),
            ("AssignedTo", 140),
        ]:
            self.task_dep_tree.heading(col, text=col)
            self.task_dep_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.task_dep_tree.yview)
        self.task_dep_tree.configure(yscrollcommand=sb.set)

        self.task_dep_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_blockers_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("Type", "Item", "DependsOn", "CurrentStatus", "DependencyStatus", "ReadyToStart")
        self.blockers_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Type", 90),
            ("Item", 260),
            ("DependsOn", 260),
            ("CurrentStatus", 140),
            ("DependencyStatus", 160),
            ("ReadyToStart", 120),
        ]:
            self.blockers_tree.heading(col, text=col)
            self.blockers_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.blockers_tree.yview)
        self.blockers_tree.configure(yscrollcommand=sb.set)

        self.blockers_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_summary_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        self.summary_text = make_readonly_text(wrap, height=28)
        self.summary_text.pack(fill="both", expand=True)

    # ========================================================
    # Actions
    # ========================================================

    def open_selected_project(self):
        project_code = self.project_select_var.get().strip()
        if not project_code:
            show_warning("No Project", "Select a project first.")
            return
        self.app.set_selected_project(project_code)
        self.show_page("projects")

    # ========================================================
    # Refresh
    # ========================================================

    def refresh_page(self):
        if not self.require_workbook():
            return

        try:
            self._refresh_projects_combo()

            project_code = self.project_select_var.get().strip()
            if not project_code:
                selected = getattr(self.app, "selected_project_code", "").strip()
                if selected:
                    self.project_select_var.set(selected)
                    project_code = selected

            if not project_code:
                self._clear_views()
                self.page_summary_label.config(text="Dependency visibility and blocker analysis")
                return

            bundle = self.app.services.projects.get_project_bundle(project_code)
            if not bundle.project:
                self._clear_views()
                self.page_summary_label.config(text="Project not found")
                return

            module_deps = self._build_module_dependency_rows(bundle.module_links or [])
            task_deps = self._build_task_dependency_rows(bundle.project_tasks or [])
            blockers = self._build_blocker_rows(project_code, module_deps, task_deps)

            filtered_blockers = self._apply_type_filter(blockers)

            self._load_module_dependencies(module_deps)
            self._load_task_dependencies(task_deps)
            self._load_blockers(filtered_blockers)
            self._load_summary(bundle, module_deps, task_deps, blockers)

            self._update_metrics(module_deps, task_deps, blockers)

            self.page_summary_label.config(
                text=f"🔗 Dependencies | {bundle.project.project_name} | Client: {bundle.project.client_name} | Status: {bundle.project.status}"
            )
            self.set_status(f"Loaded dependency view for: {project_code}")

        except Exception as exc:
            show_error("Dependency Refresh Error", str(exc))

    def _clear_views(self):
        treeview_clear(self.module_dep_tree)
        treeview_clear(self.task_dep_tree)
        treeview_clear(self.blockers_tree)
        set_text_readonly(self.summary_text, "")

        self.metric_module_links.value_var.set("0")
        self.metric_task_links.value_var.set("0")
        self.metric_open_blockers.value_var.set("0")
        self.metric_ready_items.value_var.set("0")

    def _refresh_projects_combo(self):
        projects = self.app.services.projects.list_projects()
        values = [p.project_code for p in projects]
        set_combobox_values(self.project_combo, values, keep_current=True)

    # ========================================================
    # Data builders
    # ========================================================

    def _build_module_dependency_rows(self, module_links) -> List[Dict]:
        link_lookup = {m.module_code: m for m in module_links}
        rows: List[Dict] = []

        for m in module_links:
            dep_code = norm_text(m.dependency_module_code)
            if not dep_code:
                continue

            dep_status = ""
            if dep_code in link_lookup:
                dep_status = norm_text(link_lookup[dep_code].status)

            rows.append({
                "module_code": m.module_code,
                "depends_on": dep_code,
                "module_status": norm_text(m.status),
                "dependency_status": dep_status or "Missing",
                "stage": norm_text(m.stage),
                "source_type": norm_text(m.source_type),
                "is_blocked": dep_status not in ("Completed",),
            })

        return rows

    def _build_task_dependency_rows(self, project_tasks) -> List[Dict]:
        task_lookup = {t.project_task_id: t for t in project_tasks}
        rows: List[Dict] = []

        for t in project_tasks:
            dep_id = norm_text(t.dependency_task_id)
            if not dep_id:
                continue

            dep_task = task_lookup.get(dep_id)
            dep_name = dep_id
            dep_status = "Missing"

            if dep_task:
                dep_name = f"{dep_task.module_code} | {dep_task.task_name}"
                dep_status = norm_text(dep_task.status)

            rows.append({
                "module_code": norm_text(t.module_code),
                "task_name": norm_text(t.task_name),
                "depends_on_task": dep_name,
                "task_status": norm_text(t.status),
                "dependency_status": dep_status,
                "assigned_to": norm_text(t.assigned_to),
                "is_blocked": dep_status not in ("Completed",),
            })

        return rows

    def _build_blocker_rows(self, project_code: str, module_deps: List[Dict], task_deps: List[Dict]) -> List[Dict]:
        rows: List[Dict] = []

        for md in module_deps:
            ready = md["dependency_status"] == "Completed"
            rows.append({
                "type": "MODULE",
                "item": md["module_code"],
                "depends_on": md["depends_on"],
                "current_status": md["module_status"],
                "dependency_status": md["dependency_status"],
                "ready_to_start": "YES" if ready else "NO",
                "is_blocked": not ready,
            })

        for td in task_deps:
            ready = td["dependency_status"] == "Completed"
            rows.append({
                "type": "TASK",
                "item": f"{td['module_code']} | {td['task_name']}",
                "depends_on": td["depends_on_task"],
                "current_status": td["task_status"],
                "dependency_status": td["dependency_status"],
                "ready_to_start": "YES" if ready else "NO",
                "is_blocked": not ready,
            })

        return rows

    def _apply_type_filter(self, blockers: List[Dict]) -> List[Dict]:
        selected = self.blocker_type_var.get().strip()
        if not selected or selected == "All":
            return blockers
        return [b for b in blockers if norm_text(b["type"]) == selected]

    # ========================================================
    # Loaders
    # ========================================================

    def _load_module_dependencies(self, rows: List[Dict]):
        treeview_clear(self.module_dep_tree)
        for r in rows:
            self.module_dep_tree.insert(
                "",
                "end",
                values=(
                    r["module_code"],
                    r["depends_on"],
                    r["module_status"],
                    r["dependency_status"],
                    r["stage"],
                    r["source_type"],
                )
            )

    def _load_task_dependencies(self, rows: List[Dict]):
        treeview_clear(self.task_dep_tree)
        for r in rows:
            self.task_dep_tree.insert(
                "",
                "end",
                values=(
                    r["module_code"],
                    r["task_name"],
                    r["depends_on_task"],
                    r["task_status"],
                    r["dependency_status"],
                    r["assigned_to"],
                )
            )

    def _load_blockers(self, rows: List[Dict]):
        treeview_clear(self.blockers_tree)
        for r in rows:
            self.blockers_tree.insert(
                "",
                "end",
                values=(
                    r["type"],
                    r["item"],
                    r["depends_on"],
                    r["current_status"],
                    r["dependency_status"],
                    r["ready_to_start"],
                )
            )

    def _load_summary(self, bundle, module_deps: List[Dict], task_deps: List[Dict], blockers: List[Dict]):
        project = bundle.project

        open_blockers = [b for b in blockers if b["is_blocked"]]
        ready_items = [b for b in blockers if not b["is_blocked"]]

        lines = [
            f"Project Code: {project.project_code}",
            f"Project Name: {project.project_name}",
            f"Client: {project.client_name}",
            f"Status: {project.status}",
            f"Linked Product: {project.linked_product_code or '-'}",
            "",
            f"Module Dependency Links: {len(module_deps)}",
            f"Task Dependency Links: {len(task_deps)}",
            f"Open Blockers: {len(open_blockers)}",
            f"Ready Items: {len(ready_items)}",
            "",
            "Module Dependency Notes:",
        ]

        if not module_deps:
            lines.append(" - No module dependencies configured.")
        else:
            for md in module_deps:
                state = "BLOCKED" if md["is_blocked"] else "READY"
                lines.append(
                    f" - {md['module_code']} depends on {md['depends_on']} | dep status: {md['dependency_status']} | {state}"
                )

        lines.append("")
        lines.append("Task Dependency Notes:")

        if not task_deps:
            lines.append(" - No task dependencies configured.")
        else:
            for td in task_deps:
                state = "BLOCKED" if td["is_blocked"] else "READY"
                lines.append(
                    f" - {td['module_code']} | {td['task_name']} depends on {td['depends_on_task']} | dep status: {td['dependency_status']} | {state}"
                )

        lines.append("")
        lines.append("Interpretation:")
        lines.append(" - READY means the dependency is completed, so the dependent item can start from a dependency standpoint.")
        lines.append(" - BLOCKED means the dependency is not completed or missing.")
        lines.append(" - Missing dependency references are treated as blocked because reality is rude like that.")
        lines.append(" - Update execution status from the Live Projects page.")

        set_text_readonly(self.summary_text, "\n".join(lines))

    def _update_metrics(self, module_deps: List[Dict], task_deps: List[Dict], blockers: List[Dict]):
        open_blockers = [b for b in blockers if b["is_blocked"]]
        ready_items = [b for b in blockers if not b["is_blocked"]]

        self.metric_module_links.value_var.set(str(len(module_deps)))
        self.metric_task_links.value_var.set(str(len(task_deps)))
        self.metric_open_blockers.value_var.set(str(len(open_blockers)))
        self.metric_ready_items.value_var.set(str(len(ready_items)))