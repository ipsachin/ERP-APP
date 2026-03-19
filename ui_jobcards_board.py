
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from typing import Any, Dict, List

from app_config import AppConfig
from ui_common import BasePage, treeview_clear


def norm_text(value: Any) -> str:
    return str(value or "").strip()


def safe_float(value: Any) -> float:
    try:
        return float(value or 0)
    except Exception:
        return 0.0


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class JobCardsBoardPage(BasePage):
    PAGE_NAME = "jobcards"

    BOARD_COLUMNS = [
        "Not Started",
        "Procurement",
        "Fabrication",
        "Electrical",
        "Automation",
        "Assembly",
        "Testing",
        "Complete",
    ]
    ACTIVE_COLUMNS = {"Procurement", "Fabrication", "Electrical", "Automation", "Assembly", "Testing"}

    DEPARTMENT_TO_COLUMN = {
        "PROCUREMENT": "Procurement",
        "PURCHASING": "Procurement",
        "FABRICATION": "Fabrication",
        "MECHANICAL": "Fabrication",
        "ELECTRICAL": "Electrical",
        "AUTOMATION": "Automation",
        "SOFTWARE": "Automation",
        "OPERATIONS": "Assembly",
        "ASSEMBLY": "Assembly",
        "TESTING": "Testing",
        "QA": "Testing",
        "QA/QC": "Testing",
        "COMMISSIONING": "Testing",
    }
    STATUS_TO_COLUMN = {
        "NOT STARTED": "Not Started",
        "PLANNED": "Not Started",
        "OPEN": "Not Started",
        "DRAFT": "Not Started",
        "PROCUREMENT": "Procurement",
        "FABRICATION": "Fabrication",
        "ELECTRICAL": "Electrical",
        "AUTOMATION": "Automation",
        "ASSEMBLY": "Assembly",
        "TESTING": "Testing",
        "DONE": "Complete",
        "COMPLETE": "Complete",
        "COMPLETED": "Complete",
        "CLOSED": "Complete",
    }

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)
        self.order_search_var = tk.StringVar()
        self.order_filter_status_var = tk.StringVar(value="All")
        self.current_project_code = ""
        self.project_bundle = None
        self.project_cards: List[Dict[str, Any]] = []

        self._order_index_map: List[str] = []
        self._column_frames: Dict[str, tk.Widget] = {}
        self._card_widgets: Dict[str, tk.Widget] = {}
        self._expanded_cards: set[str] = set()
        self._drag_card_id: str | None = None
        self._drag_hover_column: str | None = None
        self._drag_start_xy: tuple[int, int] | None = None
        self._drag_in_progress = False

        self._build_ui()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        self._build_topbar(wrapper)

        shell = ttk.Panedwindow(wrapper, orient="horizontal")
        shell.pack(fill="both", expand=True)

        left = ttk.Frame(shell, padding=(0, 0, 8, 0))
        right = ttk.Frame(shell)
        shell.add(left, weight=1)
        shell.add(right, weight=4)

        self._build_order_browser(left)
        self._build_right_panel(right)

    def _build_topbar(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)
        ttk.Label(top, text="Job Cards Kanban Board", style="Title.TLabel").pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)

    def _build_order_browser(self, parent):
        ttk.Label(parent, text="Live Orders", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        search_row = ttk.Frame(parent)
        search_row.pack(fill="x", pady=(0, 6))
        ttk.Entry(search_row, textvariable=self.order_search_var).pack(side="left", fill="x", expand=True)
        ttk.Button(search_row, text="Search", command=self.refresh_order_list).pack(side="left", padx=(6, 0))

        filter_row = ttk.Frame(parent)
        filter_row.pack(fill="x", pady=(0, 8))
        ttk.Label(filter_row, text="Status:").pack(side="left")
        self.status_filter_combo = ttk.Combobox(
            filter_row, textvariable=self.order_filter_status_var, state="readonly",
            values=["All", "Planned", "Active", "On Hold", "Closed", "Complete"]
        )
        self.status_filter_combo.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.status_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_order_list())

        list_frame = ttk.Frame(parent)
        list_frame.pack(fill="both", expand=True)
        self.order_list = tk.Listbox(list_frame, exportselection=False)
        ysb = ttk.Scrollbar(list_frame, orient="vertical", command=self.order_list.yview)
        self.order_list.configure(yscrollcommand=ysb.set)
        self.order_list.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.order_list.bind("<<ListboxSelect>>", self._on_order_selected)

    def _build_right_panel(self, parent):
        self.header_label = ttk.Label(parent, text="Select a live order", style="Title.TLabel")
        self.header_label.pack(anchor="w", pady=(0, 8))

        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill="both", expand=True)

        self.board_tab = ttk.Frame(self.notebook)
        self.structure_tab = ttk.Frame(self.notebook)
        self.summary_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.board_tab, text="Kanban View")
        self.notebook.add(self.structure_tab, text="Order Structure")
        self.notebook.add(self.summary_tab, text="Summary")

        self._build_board_tab(self.board_tab)
        self._build_structure_tab(self.structure_tab)
        self._build_summary_tab(self.summary_tab)

    
    def _build_board_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=4, pady=4)
        self.board_hint_var = tk.StringVar(
            value="Kanban board: drag a card to another column. Expand subtasks and tick complete directly here."
        )
        ttk.Label(top, textvariable=self.board_hint_var, style="Sub.TLabel").pack(side="left")

        board_host = ttk.Frame(parent)
        board_host.pack(fill="both", expand=True)

        self.board_canvas = tk.Canvas(board_host, highlightthickness=0, bg=AppConfig.COLOR_BG)
        self.board_xscroll = ttk.Scrollbar(board_host, orient="horizontal", command=self.board_canvas.xview)
        self.board_yscroll = ttk.Scrollbar(board_host, orient="vertical", command=self.board_canvas.yview)
        self.board_canvas.configure(xscrollcommand=self.board_xscroll.set, yscrollcommand=self.board_yscroll.set)
        self.board_canvas.grid(row=0, column=0, sticky="nsew")
        self.board_yscroll.grid(row=0, column=1, sticky="ns")
        self.board_xscroll.grid(row=1, column=0, sticky="ew")
        board_host.rowconfigure(0, weight=1)
        board_host.columnconfigure(0, weight=1)

        self.board_inner = tk.Frame(self.board_canvas, bg=AppConfig.COLOR_BG)
        self.board_window = self.board_canvas.create_window((0, 0), window=self.board_inner, anchor="nw")
        self.board_inner.bind("<Configure>", lambda e: self.board_canvas.configure(scrollregion=self.board_canvas.bbox("all")))
        self.board_canvas.bind("<Configure>", self._resize_board_canvas)

    def _build_structure_tab(self, parent):
        self.structure_tree = ttk.Treeview(parent, columns=("Type", "Qty", "Department", "Status"), show="tree headings")
        self.structure_tree.heading("#0", text="Order / Assembly / Parts")
        self.structure_tree.heading("Type", text="Type")
        self.structure_tree.heading("Qty", text="Qty")
        self.structure_tree.heading("Department", text="Department")
        self.structure_tree.heading("Status", text="Status")
        self.structure_tree.column("#0", width=340, anchor="w")
        self.structure_tree.column("Type", width=120, anchor="w")
        self.structure_tree.column("Qty", width=80, anchor="center")
        self.structure_tree.column("Department", width=140, anchor="w")
        self.structure_tree.column("Status", width=120, anchor="w")
        ysb = ttk.Scrollbar(parent, orient="vertical", command=self.structure_tree.yview)
        xsb = ttk.Scrollbar(parent, orient="horizontal", command=self.structure_tree.xview)
        self.structure_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.structure_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

    def _build_summary_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=4, pady=4)
        self.summary_stats_var = tk.StringVar(value="No live order selected")
        ttk.Label(top, textvariable=self.summary_stats_var, style="Sub.TLabel").pack(anchor="w")

        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        self.summary_text = tk.Text(frame, wrap="word", height=20)
        ysb = ttk.Scrollbar(frame, orient="vertical", command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=ysb.set)
        self.summary_text.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.summary_text.configure(state="disabled")

    def refresh_page(self) -> None:
        if not self.require_workbook():
            return
        self.refresh_order_list()
        if self.current_project_code:
            self.load_project(self.current_project_code)
        elif self._order_index_map:
            self.order_list.selection_clear(0, tk.END)
            self.order_list.selection_set(0)
            self._on_order_selected()

    def refresh_order_list(self):
        self.order_list.delete(0, tk.END)
        self._order_index_map.clear()
        projects = self._list_projects()
        search_text = norm_text(self.order_search_var.get()).lower()
        wanted_status = norm_text(self.order_filter_status_var.get()).lower()
        for p in projects:
            code = norm_text(getattr(p, "project_code", ""))
            name = norm_text(getattr(p, "project_name", ""))
            client = norm_text(getattr(p, "client_name", ""))
            status = norm_text(getattr(p, "status", ""))
            hay = f"{code} {name} {client} {status}".lower()
            if search_text and search_text not in hay:
                continue
            if wanted_status and wanted_status != "all" and wanted_status != status.lower():
                continue
            label = f"{code}  |  {name or 'Unnamed Order'}"
            if client:
                label += f"  |  {client}"
            if status:
                label += f"  |  {status}"
            self.order_list.insert(tk.END, label)
            self._order_index_map.append(code)

    def _list_projects(self):
        svc = getattr(self.app.services, "projects", None) or getattr(self.app.services, "orders", None)
        if svc and hasattr(svc, "list_projects"):
            return svc.list_projects("")
        if svc and hasattr(svc, "list_orders"):
            return svc.list_orders("")
        return []

    def _get_project_bundle(self, project_code: str):
        svc = getattr(self.app.services, "projects", None) or getattr(self.app.services, "orders", None)
        if svc and hasattr(svc, "get_project_bundle"):
            return svc.get_project_bundle(project_code)
        return None

    def _get_module_components(self, module_code: str):
        svc = getattr(self.app.services, "modules", None) or getattr(self.app.services, "assemblies", None)
        if svc and hasattr(svc, "get_module_components"):
            return svc.get_module_components(module_code)
        return []


    def _get_product_bundle(self, product_code: str):
        svc = getattr(self.app.services, "products", None)
        if svc and hasattr(svc, "get_product_bundle"):
            try:
                return svc.get_product_bundle(product_code)
            except Exception:
                return None
        return None

    def _on_order_selected(self, event=None):
        if not self._order_index_map:
            return
        selection = self.order_list.curselection()
        if not selection:
            return
        idx = int(selection[0])
        if idx >= len(self._order_index_map):
            return
        self.load_project(self._order_index_map[idx])

    def load_project(self, project_code: str):
        bundle = self._get_project_bundle(project_code)
        if not bundle or not getattr(bundle, "project", None):
            self.header_label.config(text="Select a live order")
            return
        self.current_project_code = project_code
        self.project_bundle = bundle
        project = bundle.project
        name = norm_text(getattr(project, "project_name", "")) or "Unnamed Order"
        quote = norm_text(getattr(project, "quote_ref", ""))
        client = norm_text(getattr(project, "client_name", ""))
        linked_product = norm_text(getattr(project, "linked_product_code", ""))
        header = f"{project_code}  |  {name}"
        if client:
            header += f"  |  {client}"
        if quote:
            header += f"  |  Quote: {quote}"
        if linked_product:
            header += f"  |  Product: {linked_product}"
        self.header_label.config(text=header)

        self.project_cards = self._build_cards_from_bundle(bundle)
        self._render_board()
        self._render_structure(bundle)
        self._render_summary(bundle)
        self._render_kanban_summary()

    def _extract_start_finish(self, notes: str):
        start = finish = ""
        for line in norm_text(notes).splitlines():
            if line.startswith("[START]"):
                start = line.replace("[START]", "", 1).strip()
            elif line.startswith("[FINISH]"):
                finish = line.replace("[FINISH]", "", 1).strip()
        return start, finish

    def _build_cards_from_bundle(self, bundle) -> List[Dict[str, Any]]:
        cards: List[Dict[str, Any]] = []
        project = bundle.project
        project_code = norm_text(getattr(project, "project_code", ""))
        project_name = norm_text(getattr(project, "project_name", ""))

        module_task_map: Dict[str, List[Dict[str, Any]]] = {}
        for task in list(getattr(bundle, "project_tasks", []) or []):
            module_code = norm_text(getattr(task, "module_code", "")) or "GENERAL"
            notes = norm_text(getattr(task, "notes", ""))
            start_at, finish_at = self._extract_start_finish(notes)
            module_task_map.setdefault(module_code, []).append({
                "row_id": norm_text(getattr(task, "project_task_id", "")),
                "title": norm_text(getattr(task, "task_name", "")) or "Unnamed Task",
                "department": norm_text(getattr(task, "department", "")),
                "assigned_to": norm_text(getattr(task, "assigned_to", "")),
                "status": norm_text(getattr(task, "status", "")) or norm_text(getattr(task, "stage", "")),
                "hours": safe_float(getattr(task, "estimated_hours", 0)),
                "start_at": start_at,
                "finish_at": finish_at,
                "notes": notes,
            })

        project_module_links = list(getattr(bundle, "module_links", []) or [])
        module_link_by_code = {norm_text(getattr(x, "module_code", "")): x for x in project_module_links}

        for mod in list(getattr(bundle, "modules", []) or []):
            module_code = norm_text(getattr(mod, "module_code", ""))
            module_name = norm_text(getattr(mod, "module_name", "")) or module_code
            subtasks = module_task_map.get(module_code, [])
            link = module_link_by_code.get(module_code)
            link_notes = norm_text(getattr(link, "notes", "")) if link else ""
            start_at, finish_at = self._extract_start_finish(link_notes)
            base_status = norm_text(getattr(link, "status", "")) if link else ""
            base_stage = norm_text(getattr(link, "stage", "")) if link else ""
            first_status = base_status or (subtasks[0]["status"] if subtasks else "Not Started")
            first_dept = base_stage or (subtasks[0]["department"] if subtasks else "Assembly")
            column = self._pick_board_column(stage=base_stage, status=first_status, department=first_dept)
            done_count = sum(1 for x in subtasks if norm_text(x.get("status")).upper() in {"COMPLETE","COMPLETED","DONE","CLOSED"})

            cards.append({
                "kind": "MODULE_GROUP",
                "id": f"MODULEGROUP::{module_code}",
                "row_id": module_code,
                "link_id": norm_text(getattr(link, "link_id", "")) or norm_text(getattr(link, "project_module_link_id", "")),
                "column": column,
                "project_code": project_code,
                "project_name": project_name,
                "module_code": module_code,
                "title": module_name,
                "department": first_dept,
                "assigned_to": "",
                "status": first_status or "Not Started",
                "due_date": "",
                "hours": sum(x["hours"] for x in subtasks),
                "notes": link_notes,
                "subtasks": subtasks,
                "done_count": done_count,
                "start_at": start_at,
                "finish_at": finish_at,
                "color_flag": self._card_color("", first_status),
            })

        for workorder in list(getattr(bundle, "workorders", []) or []):
            stage = norm_text(getattr(workorder, "stage", ""))
            status = norm_text(getattr(workorder, "status", ""))
            owner = norm_text(getattr(workorder, "owner", ""))
            title = norm_text(getattr(workorder, "workorder_name", "")) or "Unnamed Work Order"
            notes = norm_text(getattr(workorder, "notes", ""))
            start_at, finish_at = self._extract_start_finish(notes)
            column = self._pick_board_column(stage=stage, status=status, department=stage)
            cards.append({
                "kind": "WORKORDER",
                "id": f"WORKORDER::{norm_text(getattr(workorder, 'workorder_id', ''))}",
                "row_id": norm_text(getattr(workorder, "workorder_id", "")),
                "column": column,
                "project_code": project_code,
                "project_name": project_name,
                "module_code": "",
                "title": title,
                "department": stage or "General",
                "assigned_to": owner,
                "status": status or stage or "Open",
                "due_date": norm_text(getattr(workorder, "due_date", "")),
                "hours": 0.0,
                "notes": notes,
                "subtasks": [{
                    "row_id": norm_text(getattr(workorder, "workorder_id", "")),
                    "title": title,
                    "department": stage or "General",
                    "assigned_to": owner,
                    "status": status or stage or "Open",
                    "hours": 0.0,
                    "start_at": start_at,
                    "finish_at": finish_at,
                    "notes": notes,
                }],
                "done_count": 1 if norm_text(status).upper() in {"COMPLETE","COMPLETED","DONE","CLOSED"} else 0,
                "start_at": start_at,
                "finish_at": finish_at,
                "color_flag": self._card_color(norm_text(getattr(workorder, "due_date", "")), status),
            })

        if not cards:
            for module_code, subtasks in module_task_map.items():
                first_status = subtasks[0]["status"] if subtasks else ""
                cards.append({
                    "kind": "MODULE_GROUP",
                    "id": f"MODULEGROUP::{module_code}",
                    "row_id": module_code,
                    "link_id": "",
                    "column": self._pick_board_column(status=first_status, department=subtasks[0]["department"] if subtasks else ""),
                    "project_code": project_code,
                    "project_name": project_name,
                    "module_code": "" if module_code == "GENERAL" else module_code,
                    "title": "General Tasks" if module_code == "GENERAL" else module_code,
                    "department": subtasks[0]["department"] if subtasks else "General",
                    "assigned_to": "",
                    "status": first_status or "Not Started",
                    "due_date": "",
                    "hours": sum(x["hours"] for x in subtasks),
                    "notes": "",
                    "subtasks": subtasks,
                    "done_count": sum(1 for x in subtasks if norm_text(x.get("status")).upper() in {"COMPLETE","COMPLETED","DONE","CLOSED"}),
                    "start_at": "",
                    "finish_at": "",
                    "color_flag": self._card_color("", first_status),
                })
        return cards

    def _pick_board_column(self, stage: str = "", status: str = "", department: str = "") -> str:
        for raw in [status, stage, department]:
            key = norm_text(raw).upper()
            if key in self.STATUS_TO_COLUMN:
                return self.STATUS_TO_COLUMN[key]
            if key in self.DEPARTMENT_TO_COLUMN:
                return self.DEPARTMENT_TO_COLUMN[key]
        return "Not Started"

    def _card_color(self, due_date: str, status: str) -> str:
        status_key = norm_text(status).upper()
        if status_key in {"COMPLETE", "COMPLETED", "DONE", "CLOSED"}:
            return "#D8ECFF"
        if due_date:
            due = self._parse_date(due_date)
            if due:
                today = datetime.today().date()
                if due < today:
                    return "#FFD9D9"
                if due <= today + timedelta(days=1):
                    return "#FFE9C7"
                return "#DDF4DD"
        return "#EFEFEF"

    def _parse_date(self, text: str):
        text = norm_text(text)
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(text, fmt).date()
            except Exception:
                pass
        return None

    def _render_board(self):
        for child in self.board_inner.winfo_children():
            child.destroy()
        self._column_frames.clear()
        self._card_widgets.clear()

        grouped = {name: [] for name in self.BOARD_COLUMNS}
        for card in self.project_cards:
            grouped.setdefault(card["column"], []).append(card)

        for col_idx, col_name in enumerate(self.BOARD_COLUMNS):
            column_frame = tk.Frame(self.board_inner, bg="#E6E8EB", width=310, padx=6, pady=6)
            column_frame.grid(row=0, column=col_idx, sticky="ns", padx=8, pady=6)
            column_frame.grid_propagate(False)
            self._column_frames[col_name] = column_frame

            title = tk.Label(column_frame, text=f"{col_name} ({len(grouped.get(col_name, []))})",
                             font=("Segoe UI", 11, "bold"), bg="#E6E8EB", anchor="w")
            title.pack(fill="x", pady=(0, 6))

            cards_wrap = tk.Frame(column_frame, bg="#E6E8EB")
            cards_wrap.pack(fill="both", expand=True)

            for card in grouped.get(col_name, []):
                widget = self._render_card(cards_wrap, card)
                self._card_widgets[card["id"]] = widget

        self.board_inner.update_idletasks()
        self.board_canvas.configure(scrollregion=self.board_canvas.bbox("all"))
        self._highlight_column(None)

    def _render_kanban_summary(self):
        total_cards = len(self.project_cards)
        completed_cards = sum(1 for card in self.project_cards if card.get("column") == "Complete")
        active_cards = sum(1 for card in self.project_cards if card.get("column") in self.ACTIVE_COLUMNS)
        overdue_cards = sum(1 for card in self.project_cards if card.get("color_flag") == "#FFD9D9")

        column_counts = []
        for column_name in self.BOARD_COLUMNS:
            count = sum(1 for card in self.project_cards if card.get("column") == column_name)
            if count:
                column_counts.append(f"{column_name}: {count}")

        summary = [
            f"Kanban board: {total_cards} card(s)",
            f"Active {active_cards}",
            f"Complete {completed_cards}",
        ]
        if overdue_cards:
            summary.append(f"Overdue {overdue_cards}")
        if column_counts:
            summary.append(" | ".join(column_counts))

        self.board_hint_var.set("  |  ".join(summary))

    def _render_card(self, parent, card: Dict[str, Any]):
        frame = tk.Frame(parent, bg=card["color_flag"], bd=1, relief="solid", padx=8, pady=6, width=272, cursor="hand2")
        frame.pack(fill="x", pady=4)

        header = tk.Frame(frame, bg=card["color_flag"])
        header.pack(fill="x")

        title = tk.Label(header, text=f'{card["project_code"]}\n{card["title"]}', bg=card["color_flag"],
                         justify="left", anchor="w", font=("Segoe UI", 10, "bold"))
        title.pack(side="left", fill="x", expand=True)

        subtasks = card.get("subtasks", [])
        toggle_text = f"{'▼' if card['id'] in self._expanded_cards else '▶'} {len(subtasks)} subtask(s)"
        toggle_btn = tk.Label(header, text=toggle_text, bg=card["color_flag"], fg="#4d4d4d",
                              cursor="hand2", font=("Segoe UI", 8, "underline"))
        toggle_btn.pack(side="right")

        details = []
        if card.get("module_code"):
            details.append(f"Assembly: {card['module_code']}")
        if card.get("department"):
            details.append(f"Dept: {card['department']}")
        if safe_float(card.get("hours", 0)) > 0:
            details.append(f"Hours: {safe_float(card['hours']):g}")
        details.append(f"Status: {card.get('status','')}")
        if card.get("start_at"):
            details.append(f"Start: {card['start_at']}")
        if card.get("finish_at"):
            details.append(f"Finish: {card['finish_at']}")
        if subtasks:
            details.append(f"Done: {card.get('done_count',0)}/{len(subtasks)}")
        tk.Label(frame, text="\n".join(details), bg=card["color_flag"], justify="left", anchor="w",
                 font=("Segoe UI", 9)).pack(fill="x", pady=(4, 4))

        action_row = tk.Frame(frame, bg=card["color_flag"])
        action_row.pack(fill="x")
        mod_var = tk.BooleanVar(value=norm_text(card.get("status")).upper() in {"COMPLETE","COMPLETED","DONE","CLOSED"})
        mod_chk = ttk.Checkbutton(action_row, text="Module Complete", variable=mod_var,
                                  command=lambda c=card, v=mod_var: self._toggle_module_complete(c, v.get()))
        mod_chk.pack(side="left")
        ttk.Button(action_row, text="Start Now", command=lambda c=card: self._stamp_module_time(c, "start")).pack(side="left", padx=(6, 2))
        ttk.Button(action_row, text="Finish Now", command=lambda c=card: self._stamp_module_time(c, "finish")).pack(side="left", padx=2)

        if card["id"] in self._expanded_cards and subtasks:
            holder = tk.Frame(frame, bg="#FFFFFF")
            holder.pack(fill="x", pady=(6, 0))
            for idx, st in enumerate(subtasks, start=1):
                row = tk.Frame(holder, bg="#FFFFFF")
                row.pack(fill="x", padx=4, pady=2)
                done = norm_text(st.get("status")).upper() in {"COMPLETE","COMPLETED","DONE","CLOSED"}
                st_var = tk.BooleanVar(value=done)
                ttk.Checkbutton(row, variable=st_var,
                                command=lambda s=st, v=st_var: self._toggle_subtask_complete(s, v.get())).pack(side="left")
                info = f"{idx}. {st.get('title','Task')}"
                meta = []
                if st.get("department"):
                    meta.append(st["department"])
                if st.get("assigned_to"):
                    meta.append(st["assigned_to"])
                if safe_float(st.get("hours", 0)) > 0:
                    meta.append(f"{safe_float(st['hours']):g}h")
                meta.append(st.get("status", ""))
                if st.get("start_at"):
                    meta.append(f"S:{st['start_at']}")
                if st.get("finish_at"):
                    meta.append(f"F:{st['finish_at']}")
                if meta:
                    info += "  |  " + "  |  ".join([m for m in meta if m])
                tk.Label(row, text=info, bg="#FFFFFF", anchor="w", justify="left", wraplength=180,
                         font=("Segoe UI", 8)).pack(side="left", fill="x", expand=True)
                ttk.Button(row, text="Start", width=6, command=lambda s=st: self._stamp_subtask_time(s, "start")).pack(side="right", padx=1)
                ttk.Button(row, text="Finish", width=6, command=lambda s=st: self._stamp_subtask_time(s, "finish")).pack(side="right", padx=1)

        for widget in (frame, header, title):
            widget.bind("<ButtonPress-1>", lambda e, cid=card["id"]: self._drag_start(e, cid))
            widget.bind("<B1-Motion>", lambda e, cid=card["id"]: self._drag_motion(e, cid))
            widget.bind("<ButtonRelease-1>", lambda e, cid=card["id"]: self._drag_release(e, cid))
        toggle_btn.bind("<Button-1>", lambda e, cid=card["id"]: self._toggle_expand(cid))
        return frame

    def _toggle_expand(self, card_id: str):
        if card_id in self._expanded_cards:
            self._expanded_cards.remove(card_id)
        else:
            self._expanded_cards.add(card_id)
        self._render_board()

    def _set_card_drag_state(self, card_id: str | None, is_active: bool):
        if not card_id:
            return
        widget = self._card_widgets.get(card_id)
        card = self._find_card(card_id)
        if not widget or not card:
            return

        base_bg = card.get("color_flag", "#EFEFEF")
        active_bg = "#FFF2B8"
        target_bg = active_bg if is_active else base_bg

        try:
            widget.configure(
                bg=target_bg,
                relief="raised" if is_active else "solid",
                bd=3 if is_active else 1,
                highlightthickness=3 if is_active else 0,
                highlightbackground="#D18B00" if is_active else target_bg,
            )
            self._update_card_widget_colors(widget, base_bg, target_bg, active_bg)
        except Exception:
            pass

    def _update_card_widget_colors(self, widget, base_bg: str, target_bg: str, active_bg: str):
        for child in widget.winfo_children():
            try:
                if isinstance(child, (tk.Frame, tk.Label)):
                    current_bg = child.cget("bg")
                    if current_bg in {base_bg, active_bg}:
                        child.configure(bg=target_bg)
                    if isinstance(child, tk.Label):
                        child.configure(fg="#6E4B00" if target_bg == active_bg else "#000000")
                self._update_card_widget_colors(child, base_bg, target_bg, active_bg)
            except Exception:
                pass

    def _drag_start(self, event, card_id: str):
        self._drag_card_id = card_id
        self._drag_start_xy = (event.x_root, event.y_root)
        self._drag_in_progress = False
        self._drag_hover_column = None
        self._set_card_drag_state(card_id, True)

    def _drag_motion(self, event, card_id: str):
        if self._drag_card_id != card_id:
            return
        if self._drag_start_xy:
            dx = abs(event.x_root - self._drag_start_xy[0])
            dy = abs(event.y_root - self._drag_start_xy[1])
            if not self._drag_in_progress and max(dx, dy) >= 8:
                self._drag_in_progress = True
        if not self._drag_in_progress:
            return
        target_column = self._hit_test_column(event.x_root, event.y_root)
        self._highlight_column(target_column)

    def _drag_release(self, event, card_id: str):
        self._set_card_drag_state(card_id, False)
        if self._drag_card_id != card_id:
            return
        if not self._drag_in_progress:
            self._clear_drag_state()
            return
        target_column = self._hit_test_column(event.x_root, event.y_root)
        self._clear_drag_state()
        if not target_column:
            return
        card = self._find_card(card_id)
        if not card:
            return
        if target_column == card["column"]:
            return
        if self._persist_card_move(card, target_column):
            card["column"] = target_column
            card["department"], card["status"] = self._column_to_department_status(card, target_column)
            card["color_flag"] = self._card_color(card.get("due_date", ""), card["status"])
            self._render_board()
            self._render_kanban_summary()
            self._render_summary_from_cards()

    def _highlight_column(self, target_column: str | None):
        if self._drag_hover_column == target_column and target_column is not None:
            return
        for name, frame in self._column_frames.items():
            try:
                bg = "#D8EAFB" if name == target_column else "#E6E8EB"
                frame.configure(bg=bg, highlightthickness=2 if name == target_column else 0,
                                highlightbackground="#4A90E2" if name == target_column else bg)
                for child in frame.winfo_children():
                    if isinstance(child, tk.Frame) or isinstance(child, tk.Label):
                        child.configure(bg=bg)
            except Exception:
                pass
        self._drag_hover_column = target_column
        if self._drag_in_progress:
            if target_column:
                self.board_hint_var.set(f"Drop in: {target_column}")
            else:
                self._render_kanban_summary()
        elif target_column is None:
            self._render_kanban_summary()

    def _clear_drag_state(self):
        self._drag_card_id = None
        self._drag_start_xy = None
        self._drag_in_progress = False
        self._highlight_column(None)

    def _hit_test_column(self, x_root: int, y_root: int) -> str | None:
        for name, frame in self._column_frames.items():
            try:
                left, top = frame.winfo_rootx(), frame.winfo_rooty()
                right, bottom = left + frame.winfo_width(), top + frame.winfo_height()
                if left <= x_root <= right and top <= y_root <= bottom:
                    return name
            except Exception:
                pass
        return None

    def _find_card(self, card_id: str):
        for c in self.project_cards:
            if c["id"] == card_id:
                return c
        return None

    def _column_to_department_status(self, card: Dict[str, Any], column: str):
        if column == "Not Started":
            return card.get("department", "") or "General", "Not Started"
        if column == "Complete":
            return card.get("department", "") or "General", "Complete"
        if column in self.ACTIVE_COLUMNS:
            return column, column
        return card.get("department", "") or "General", column

    def _sheet_name(self, attr_name: str, fallback: str) -> str:
        return getattr(AppConfig, attr_name, fallback)

    def _persist_notes(self, sheet: str, key_name: str, row_id: str, notes: str) -> bool:
        repo = getattr(self.app, "repo", None)
        if repo is None or not row_id:
            return False
        return bool(repo.update_row_by_key_name(sheet, key_name, row_id, {"Notes": notes, "UpdatedOn": now_stamp()}))

    def _update_note_stamp(self, existing: str, kind: str, stamp: str) -> str:
        lines = [line for line in norm_text(existing).splitlines() if not line.startswith(f"[{kind.upper()}]")]
        lines.append(f"[{kind.upper()}] {stamp}")
        return "\n".join(lines)

    def _toggle_module_complete(self, card: Dict[str, Any], is_done: bool):
        target = "Complete" if is_done else "Not Started"
        if self._persist_card_move(card, target):
            self.load_project(self.current_project_code)

    def _toggle_subtask_complete(self, subtask: Dict[str, Any], is_done: bool):
        row_id = norm_text(subtask.get("row_id"))
        if not row_id:
            return
        repo = getattr(self.app, "repo", None)
        if repo is None:
            return
        try:
            repo.update_row_by_key_name(
                self._sheet_name("SHEET_PROJECT_TASKS", "ProjectTasks"),
                "ProjectTaskID",
                row_id,
                {"Status": "Complete" if is_done else "Not Started",
                 "Stage": "Complete" if is_done else subtask.get("department", ""),
                 "UpdatedOn": now_stamp()}
            )
            self.load_project(self.current_project_code)
        except Exception as exc:
            messagebox.showerror("Subtask Update Error", str(exc))

    def _stamp_module_time(self, card: Dict[str, Any], which: str):
        sheet = self._sheet_name("SHEET_PROJECT_MODULES", "ProjectModules")
        key_name = "LinkID"
        row_id = norm_text(card.get("link_id")) or norm_text(card.get("row_id"))
        note_base = norm_text(card.get("notes"))
        notes = self._update_note_stamp(note_base, "start" if which == "start" else "finish", now_stamp())
        repo = getattr(self.app, "repo", None)
        if repo is None:
            return
        try:
            ok = False
            if norm_text(card.get("link_id")):
                ok = bool(repo.update_row_by_key_name(sheet, key_name, norm_text(card.get("link_id")), {"Notes": notes, "UpdatedOn": now_stamp()}))
            if not ok and norm_text(card.get("row_id")):
                ok = bool(repo.update_row_by_key_name(sheet, "ModuleCode", norm_text(card.get("row_id")), {"Notes": notes, "UpdatedOn": now_stamp()}))
            if ok:
                self.load_project(self.current_project_code)
        except Exception as exc:
            messagebox.showerror("Time Stamp Error", str(exc))

    def _stamp_subtask_time(self, subtask: Dict[str, Any], which: str):
        row_id = norm_text(subtask.get("row_id"))
        if not row_id:
            return
        repo = getattr(self.app, "repo", None)
        if repo is None:
            return
        try:
            notes = self._update_note_stamp(norm_text(subtask.get("notes")), which, now_stamp())
            repo.update_row_by_key_name(
                self._sheet_name("SHEET_PROJECT_TASKS", "ProjectTasks"),
                "ProjectTaskID",
                row_id,
                {"Notes": notes, "UpdatedOn": now_stamp()}
            )
            self.load_project(self.current_project_code)
        except Exception as exc:
            messagebox.showerror("Subtask Time Error", str(exc))

    def _persist_card_move(self, card: Dict[str, Any], target_column: str) -> bool:
        repo = getattr(self.app, "repo", None)
        if repo is None:
            return False
        department, status = self._column_to_department_status(card, target_column)
        try:
            if card["kind"] == "WORKORDER":
                return bool(repo.update_row_by_key_name(
                    self._sheet_name("SHEET_WORKORDERS", "WorkOrders"),
                    "WorkOrderID",
                    card["row_id"],
                    {"Stage": target_column, "Status": status, "UpdatedOn": now_stamp()}
                ))
            if card["kind"] == "MODULE_GROUP":
                ok = True
                link_id = norm_text(card.get("link_id"))
                sheet_mod = self._sheet_name("SHEET_PROJECT_MODULES", "ProjectModules")
                if link_id:
                    ok = bool(repo.update_row_by_key_name(
                        sheet_mod, "LinkID", link_id,
                        {"Stage": target_column, "Status": status, "UpdatedOn": now_stamp()}
                    ))
                task_rows = repo.read_sheet_as_dicts(self._sheet_name("SHEET_PROJECT_TASKS", "ProjectTasks"))
                for row in task_rows:
                    if norm_text(row.get("ProjectCode")) == self.current_project_code and norm_text(row.get("ModuleCode")) == card["row_id"]:
                        repo.update_row_by_key_name(
                            self._sheet_name("SHEET_PROJECT_TASKS", "ProjectTasks"),
                            "ProjectTaskID",
                            norm_text(row.get("ProjectTaskID")),
                            {"Stage": target_column, "Status": status, "Department": department, "UpdatedOn": now_stamp()}
                        )
                return ok
        except Exception as exc:
            messagebox.showerror("Drag Drop Error", str(exc))
            return False
        return False

    def _resize_board_canvas(self, event=None):
        self.board_canvas.itemconfigure(self.board_window, height=max(self.board_canvas.winfo_height(), self.board_inner.winfo_reqheight()))

    def _render_structure(self, bundle):
        treeview_clear(self.structure_tree)
        project = bundle.project
        root = self.structure_tree.insert(
            "", "end",
            text=f"{norm_text(getattr(project, 'project_code', ''))} - {norm_text(getattr(project, 'project_name', '')) or 'Unnamed Order'}",
            values=("Live Order", "", "", norm_text(getattr(project, "status", "")))
        )
        module_status_map = {}
        for link in list(getattr(bundle, "module_links", []) or []):
            module_status_map[norm_text(getattr(link, "module_code", ""))] = norm_text(getattr(link, "status", "")) or norm_text(getattr(link, "stage", ""))
        tasks_by_module: Dict[str, int] = {}
        for task in list(getattr(bundle, "project_tasks", []) or []):
            module_code = norm_text(getattr(task, "module_code", ""))
            tasks_by_module[module_code] = tasks_by_module.get(module_code, 0) + 1
        for mod in list(getattr(bundle, "modules", []) or []):
            module_code = norm_text(getattr(mod, "module_code", ""))
            mod_node = self.structure_tree.insert(root, "end", text=norm_text(getattr(mod, "module_name", "")) or module_code,
                                                  values=("Assembly", 1, "", module_status_map.get(module_code, "")), open=True)
            for comp in self._get_module_components(module_code):
                self.structure_tree.insert(
                    mod_node, "end",
                    text=norm_text(getattr(comp, "component_name", "")) or "Part",
                    values=("Part", safe_float(getattr(comp, "qty", 0)) or "", norm_text(getattr(comp, "preferred_supplier", "")), "")
                )
            if tasks_by_module.get(module_code):
                self.structure_tree.insert(mod_node, "end", text=f"{tasks_by_module[module_code]} linked task(s)",
                                           values=("Task Summary", "", "", ""))
        for wo in list(getattr(bundle, "workorders", []) or []):
            self.structure_tree.insert(root, "end", text=norm_text(getattr(wo, "workorder_name", "")) or "Work Order",
                                       values=("Job Card", "", norm_text(getattr(wo, "owner", "")), norm_text(getattr(wo, "status", ""))))
        self.structure_tree.item(root, open=True)

    
    def _render_summary(self, bundle):
        project = bundle.project
        project_code = norm_text(getattr(project, "project_code", ""))
        project_name = norm_text(getattr(project, "project_name", "")) or "Unnamed Order"
        client = norm_text(getattr(project, "client_name", ""))
        status = norm_text(getattr(project, "status", ""))
        linked_product_code = norm_text(getattr(project, "linked_product_code", ""))

        total_hours = sum(safe_float(getattr(t, "estimated_hours", 0)) for t in list(getattr(bundle, "project_tasks", []) or []))
        total_modules = len(list(getattr(bundle, "modules", []) or []))
        total_tasks = len(list(getattr(bundle, "project_tasks", []) or []))
        total_workorders = len(list(getattr(bundle, "workorders", []) or []))
        overdue_cards = sum(1 for c in self.project_cards if c["color_flag"] == "#FFD9D9")
        due_soon_cards = sum(1 for c in self.project_cards if c["color_flag"] == "#FFE9C7")

        self.summary_stats_var.set(
            f"Assemblies: {total_modules}   |   Tasks: {total_tasks}   |   Work Orders: {total_workorders}   |   Labour Hours: {total_hours:g}"
        )

        lines = [
            "LIVE ORDER SUMMARY",
            "------------------",
            f"Live Order: {project_code}",
            f"Name: {project_name}",
            f"Client: {client}",
            f"Status: {status}",
            "",
            f"Total Assemblies: {total_modules}",
            f"Total Tasks: {total_tasks}",
            f"Total Work Orders: {total_workorders}",
            f"Total Labour Hours: {total_hours:g}",
            f"Overdue Cards: {overdue_cards}",
            f"Due Soon Cards: {due_soon_cards}",
            "",
        ]

        if linked_product_code:
            pb = self._get_product_bundle(linked_product_code)
            if pb and getattr(pb, "product", None):
                prod = pb.product
                prod_name = norm_text(getattr(prod, "product_name", "")) or linked_product_code
                prod_quote = norm_text(getattr(prod, "quote_ref", "")) or norm_text(getattr(project, "quote_ref", ""))
                prod_status = norm_text(getattr(prod, "status", ""))
                prod_desc = norm_text(getattr(prod, "description", ""))
                prod_mod_links = list(getattr(pb, "module_links", []) or [])
                prod_modules = list(getattr(pb, "modules", []) or [])
                prod_hours = safe_float(getattr(pb, "total_hours", 0))

                lines += [
                    "PRODUCT / QUOTE SUMMARY",
                    "-----------------------",
                    f"Product Code: {linked_product_code}",
                    f"Product Name: {prod_name}",
                    f"Quote Ref: {prod_quote}",
                    f"Product Status: {prod_status}",
                    f"Configured Assemblies: {len(prod_mod_links)}",
                    f"Assembly Definitions Loaded: {len(prod_modules)}",
                    f"Quoted Labour Hours: {prod_hours:g}",
                ]
                if prod_desc:
                    lines += ["", "Description:", prod_desc]

                if prod_mod_links:
                    lines += ["", "Quoted Assemblies:"]
                    module_name_map = {norm_text(getattr(m, "module_code", "")): norm_text(getattr(m, "module_name", "")) for m in prod_modules}
                    for link in prod_mod_links:
                        mod_code = norm_text(getattr(link, "module_code", ""))
                        mod_name = module_name_map.get(mod_code, mod_code)
                        qty = getattr(link, "module_qty", 1)
                        lines.append(f"- {mod_code} | {mod_name} | Qty {qty}")

            else:
                lines += [
                    "PRODUCT / QUOTE SUMMARY",
                    "-----------------------",
                    f"Product Code: {linked_product_code}",
                    f"Quote Ref: {norm_text(getattr(project, 'quote_ref', ''))}",
                    "Linked product bundle could not be loaded.",
                ]
        else:
            lines += [
                "PRODUCT / QUOTE SUMMARY",
                "-----------------------",
                f"Quote Ref: {norm_text(getattr(project, 'quote_ref', ''))}",
                "No linked product on this live order.",
            ]

        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", "\n".join(lines))
        self.summary_text.configure(state="disabled")

    
    def _render_summary_from_cards(self):
        overdue_cards = sum(1 for c in self.project_cards if c["color_flag"] == "#FFD9D9")
        due_soon_cards = sum(1 for c in self.project_cards if c["color_flag"] == "#FFE9C7")
        completed = sum(1 for c in self.project_cards if c["column"] == "Complete")
        total = len(self.project_cards)
        self.summary_stats_var.set(
            f"Cards: {total}   |   Complete: {completed}   |   Overdue: {overdue_cards}   |   Due Soon: {due_soon_cards}"
        )

LiveOrderBoardPage = JobCardsBoardPage
