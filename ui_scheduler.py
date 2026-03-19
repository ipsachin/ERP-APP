# ============================================================
# ui_scheduler.py
# Scheduling / workload / blockers page for Liquimech ERP Desktop
# ============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Dict, List

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


class SchedulerPage(BasePage):
    PAGE_NAME = "scheduler"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        self.project_select_var = tk.StringVar()
        self.status_filter_var = tk.StringVar(value="All")
        self.department_filter_var = tk.StringVar(value="All")

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
            text="Scheduling and workload view",
            style="Title.TLabel"
        )
        self.page_summary_label.pack(anchor="w", pady=(0, 10))

        filters = ttk.LabelFrame(wrapper, text="Filters", style="Card.TLabelframe", padding=12)
        filters.pack(fill="x", pady=(0, 10))

        ttk.Label(filters, text="Project").grid(row=0, column=0, sticky="w", pady=4)
        self.project_combo = ttk.Combobox(filters, textvariable=self.project_select_var, state="readonly")
        self.project_combo.grid(row=0, column=1, sticky="ew", padx=6)
        self.project_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_page())

        ttk.Label(filters, text="Task Status").grid(row=0, column=2, sticky="w", pady=4)
        self.status_filter_combo = ttk.Combobox(
            filters,
            textvariable=self.status_filter_var,
            values=["All"] + AppConfig.TASK_STATUSES,
            state="readonly"
        )
        self.status_filter_combo.grid(row=0, column=3, sticky="ew", padx=6)
        self.status_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_page())

        ttk.Label(filters, text="Department").grid(row=0, column=4, sticky="w", pady=4)
        self.department_filter_combo = ttk.Combobox(
            filters,
            textvariable=self.department_filter_var,
            values=["All"] + AppConfig.DEPARTMENTS,
            state="readonly"
        )
        self.department_filter_combo.grid(row=0, column=5, sticky="ew", padx=6)
        self.department_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_page())

        ttk.Button(filters, text="Load Selected Project", command=self.open_selected_project).grid(row=0, column=6, sticky="ew", padx=6)

        filters.columnconfigure(1, weight=1)
        filters.columnconfigure(3, weight=1)
        filters.columnconfigure(5, weight=1)

        metrics = ttk.Frame(wrapper)
        metrics.pack(fill="x", pady=(0, 10))

        self.metric_total_tasks = self._make_metric_card(metrics, "Visible Tasks", "0")
        self.metric_total_tasks.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_total_hours = self._make_metric_card(metrics, "Visible Hours", "0.00")
        self.metric_total_hours.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_blockers = self._make_metric_card(metrics, "Open Blockers", "0")
        self.metric_blockers.pack(side="left", fill="x", expand=True, padx=4)

        self.metric_departments = self._make_metric_card(metrics, "Departments Loaded", "0")
        self.metric_departments.pack(side="left", fill="x", expand=True, padx=4)

        tabs = ttk.Notebook(wrapper)
        tabs.pack(fill="both", expand=True)

        self.workload_tab = ttk.Frame(tabs)
        self.tasks_tab = ttk.Frame(tabs)
        self.blockers_tab = ttk.Frame(tabs)
        self.summary_tab = ttk.Frame(tabs)

        tabs.add(self.workload_tab, text="Department Workload")
        tabs.add(self.tasks_tab, text="Project Tasks")
        tabs.add(self.blockers_tab, text="Open Blockers")
        tabs.add(self.summary_tab, text="Planning Summary")

        self._build_workload_tab(self.workload_tab)
        self._build_tasks_tab(self.tasks_tab)
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
        lbl = ttk.Label(card, textvariable=value_var, style="Hero.TLabel")
        lbl.pack(anchor="w")
        card.value_var = value_var
        return card

    def _build_workload_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("Department", "Hours", "TaskCount")
        self.workload_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Department", 220),
            ("Hours", 120),
            ("TaskCount", 120),
        ]:
            self.workload_tree.heading(col, text=col)
            self.workload_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.workload_tree.yview)
        self.workload_tree.configure(yscrollcommand=sb.set)

        self.workload_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_tasks_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("ModuleCode", "TaskName", "Department", "Hours", "Stage", "Status", "AssignedTo", "Dependency", "Notes")
        self.tasks_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("ModuleCode", 180),
            ("TaskName", 240),
            ("Department", 130),
            ("Hours", 80),
            ("Stage", 130),
            ("Status", 120),
            ("AssignedTo", 140),
            ("Dependency", 220),
            ("Notes", 240),
        ]:
            self.tasks_tree.heading(col, text=col)
            self.tasks_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tasks_tree.yview)
        self.tasks_tree.configure(yscrollcommand=sb.set)

        self.tasks_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_blockers_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("Type", "Item", "DependsOn", "CurrentStatus", "DependencyStatus")
        self.blockers_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("Type", 100),
            ("Item", 260),
            ("DependsOn", 260),
            ("CurrentStatus", 140),
            ("DependencyStatus", 160),
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
                # default to selected project if available
                selected = getattr(self.app, "selected_project_code", "").strip()
                if selected:
                    self.project_select_var.set(selected)
                    project_code = selected

            if not project_code:
                self._clear_views()
                self.page_summary_label.config(text="Scheduling and workload view")
                return

            bundle = self.app.services.projects.get_project_bundle(project_code)
            if not bundle.project:
                self._clear_views()
                self.page_summary_label.config(text="Project not found")
                return

            filtered_tasks = self._get_filtered_tasks(bundle.project_tasks or [])
            blockers = self._build_blockers(bundle.project_tasks or [], bundle.module_links or [], project_code)
            workload = self._build_department_workload(filtered_tasks)

            self._load_workload(workload)
            self._load_tasks(filtered_tasks)
            self._load_blockers(blockers)
            self._load_summary(bundle, filtered_tasks, workload, blockers)
            self._update_metrics(filtered_tasks, workload, blockers)

            self.page_summary_label.config(
                text=f"📅 Scheduling | {bundle.project.project_name} | Client: {bundle.project.client_name} | Status: {bundle.project.status}"
            )
            self.set_status(f"Loaded scheduling view for: {project_code}")

        except Exception as exc:
            show_error("Scheduler Refresh Error", str(exc))

    def _clear_views(self):
        treeview_clear(self.workload_tree)
        treeview_clear(self.tasks_tree)
        treeview_clear(self.blockers_tree)
        set_text_readonly(self.summary_text, "")

        self.metric_total_tasks.value_var.set("0")
        self.metric_total_hours.value_var.set("0.00")
        self.metric_blockers.value_var.set("0")
        self.metric_departments.value_var.set("0")

    def _refresh_projects_combo(self):
        projects = self.app.services.projects.list_projects()
        values = [p.project_code for p in projects]
        set_combobox_values(self.project_combo, values, keep_current=True)

    def _get_filtered_tasks(self, tasks):
        status_filter = self.status_filter_var.get().strip()
        dept_filter = self.department_filter_var.get().strip()

        filtered = []
        for t in tasks:
            if status_filter and status_filter != "All" and norm_text(t.status) != status_filter:
                continue
            if dept_filter and dept_filter != "All" and norm_text(t.department) != dept_filter:
                continue
            filtered.append(t)
        return filtered

    def _build_blockers(self, tasks, module_links, project_code: str):
        task_lookup = {norm_text(t.project_task_id): t for t in tasks}
        module_lookup = {norm_text(m.module_code): m for m in module_links}
        blockers = []

        for task in tasks:
            dep_id = norm_text(task.dependency_task_id)
            if dep_id and dep_id in task_lookup:
                dep_task = task_lookup[dep_id]
                if norm_text(dep_task.status) != "Completed":
                    blockers.append({
                        "type": "TASK",
                        "project_code": project_code,
                        "task_name": task.task_name,
                        "depends_on": dep_task.task_name,
                        "current_status": task.status,
                        "dependency_status": dep_task.status,
                    })

        for mod in module_links:
            dep_mod_code = norm_text(mod.dependency_module_code)
            if dep_mod_code and dep_mod_code in module_lookup:
                dep_mod = module_lookup[dep_mod_code]
                if norm_text(dep_mod.status) != "Completed":
                    blockers.append({
                        "type": "MODULE",
                        "project_code": project_code,
                        "module_code": mod.module_code,
                        "depends_on": dep_mod.module_code,
                        "current_status": mod.status,
                        "dependency_status": dep_mod.status,
                    })

        return blockers

    def _build_department_workload(self, tasks) -> Dict[str, Dict[str, float]]:
        data: Dict[str, Dict[str, float]] = {}

        for t in tasks:
            dept = norm_text(t.department) or "Unassigned"
            if dept not in data:
                data[dept] = {"hours": 0.0, "count": 0}
            data[dept]["hours"] += float(t.estimated_hours or 0.0)
            data[dept]["count"] += 1

        return dict(sorted(data.items(), key=lambda kv: kv[0]))

    def _load_workload(self, workload: Dict[str, Dict[str, float]]):
        treeview_clear(self.workload_tree)
        for dept, info in workload.items():
            self.workload_tree.insert(
                "",
                "end",
                values=(dept, f"{info['hours']:.2f}", info["count"])
            )

    def _load_tasks(self, tasks):
        treeview_clear(self.tasks_tree)
        all_tasks = self.app.services.projects.get_project_tasks(self.project_select_var.get().strip())
        task_lookup = {t.project_task_id: t for t in all_tasks}

        for t in tasks:
            dep_display = ""
            if norm_text(t.dependency_task_id) and t.dependency_task_id in task_lookup:
                dep = task_lookup[t.dependency_task_id]
                dep_display = f"{dep.module_code} | {dep.task_name}"
            elif norm_text(t.dependency_task_id):
                dep_display = t.dependency_task_id

            self.tasks_tree.insert(
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
                    dep_display,
                    t.notes,
                ),
                tags=(t.project_task_id,),
            )

    def _load_blockers(self, blockers: List[Dict]):
        treeview_clear(self.blockers_tree)

        for b in blockers:
            if b.get("type") == "TASK":
                item = b.get("task_name", "")
            else:
                item = b.get("module_code", "")

            self.blockers_tree.insert(
                "",
                "end",
                values=(
                    b.get("type", ""),
                    item,
                    b.get("depends_on", ""),
                    b.get("current_status", ""),
                    b.get("dependency_status", ""),
                )
            )

    def _load_summary(self, bundle, filtered_tasks, workload, blockers):
        project = bundle.project

        total_hours = sum(float(t.estimated_hours or 0.0) for t in filtered_tasks)
        completed_tasks = [t for t in filtered_tasks if norm_text(t.status) == "Completed"]
        open_tasks = [t for t in filtered_tasks if norm_text(t.status) != "Completed"]

        lines = [
            f"Project Code: {project.project_code}",
            f"Project Name: {project.project_name}",
            f"Client: {project.client_name}",
            f"Status: {project.status}",
            f"Linked Product: {project.linked_product_code or '-'}",
            f"Start Date: {project.start_date or '-'}",
            f"Due Date: {project.due_date or '-'}",
            "",
            f"Filtered Visible Tasks: {len(filtered_tasks)}",
            f"Filtered Visible Hours: {total_hours:.2f}",
            f"Completed Tasks (filtered): {len(completed_tasks)}",
            f"Open Tasks (filtered): {len(open_tasks)}",
            f"Open Blockers: {len(blockers)}",
            "",
            "Department Workload:",
        ]

        if not workload:
            lines.append(" - No workload data.")
        else:
            for dept, info in workload.items():
                lines.append(f" - {dept}: {info['hours']:.2f} hrs across {info['count']} tasks")

        lines.append("")
        lines.append("Blocker Summary:")

        if not blockers:
            lines.append(" - No current blockers.")
        else:
            for b in blockers:
                if b.get("type") == "TASK":
                    lines.append(
                        f" - TASK {b.get('task_name')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})"
                    )
                else:
                    lines.append(
                        f" - MODULE {b.get('module_code')} blocked by {b.get('depends_on')} ({b.get('dependency_status')})"
                    )

        lines.append("")
        lines.append("Planning Notes:")
        lines.append(" - This page shows filtered execution visibility, not master engineering definitions.")
        lines.append(" - Use the Live Projects page to update project module/task execution fields.")
        lines.append(" - Use the Dependencies page for a cleaner dependency-centric view.")

        set_text_readonly(self.summary_text, "\n".join(lines))

    def _update_metrics(self, filtered_tasks, workload, blockers):
        total_hours = sum(float(t.estimated_hours or 0.0) for t in filtered_tasks)

        self.metric_total_tasks.value_var.set(str(len(filtered_tasks)))
        self.metric_total_hours.value_var.set(f"{total_hours:.2f}")
        self.metric_blockers.value_var.set(str(len(blockers)))
        self.metric_departments.value_var.set(str(len(workload)))
