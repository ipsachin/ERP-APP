# ============================================================
# ui_modules.py
# Module engineering page for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk, filedialog
from typing import Dict, List, Optional

from app_config import AppConfig
from models import TaskRecord, ComponentRecord, DocumentRecord, ModuleRecord
from ui_common import (
    BasePage,
    DraggableListbox,
    attach_tooltip,
    treeview_clear,
    set_combobox_values,
    show_warning,
    show_error,
    show_info,
    open_file_with_default_app,
    make_readonly_text,
    set_text_readonly,
    get_text_value,
    Validators,
)


def norm_text(value) -> str:
    return str(value or "").strip()


class ModulePage(BasePage):
    PAGE_NAME = "modules"

    def __init__(self, master, app, **kwargs):
        super().__init__(master, app, **kwargs)

        # ----------------------------------------------------
        # Shared state
        # ----------------------------------------------------
        self.parent_task_map: Dict[str, str] = {}
        self.current_task_id: Optional[str] = None
        self.current_component_id: Optional[str] = None
        self.current_document_id: Optional[str] = None

        # ----------------------------------------------------
        # Module vars
        # ----------------------------------------------------
        self.module_code_var = tk.StringVar()
        self.quote_ref_var = tk.StringVar()
        self.module_name_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.estimated_hours_var = tk.StringVar()
        self.stock_on_hand_var = tk.StringVar()
        self.status_var_local = tk.StringVar(value="Draft")

        # ----------------------------------------------------
        # Task vars
        # ----------------------------------------------------
        self.task_name_var = tk.StringVar()
        self.task_department_var = tk.StringVar()
        self.task_hours_var = tk.StringVar()
        self.task_stage_var = tk.StringVar()
        self.task_status_var = tk.StringVar(value="Not Started")
        self.task_notes_var = tk.StringVar()
        self.task_dependency_display_var = tk.StringVar()
        self.parent_task_display_var = tk.StringVar()
        self.is_subtask_var = tk.BooleanVar(value=False)

        # ----------------------------------------------------
        # Component vars
        # ----------------------------------------------------
        self.component_name_var = tk.StringVar()
        self.component_qty_var = tk.StringVar()
        self.component_soh_var = tk.StringVar()
        self.component_supplier_var = tk.StringVar()
        self.component_lead_var = tk.StringVar()
        self.component_partno_var = tk.StringVar()
        self.component_notes_var = tk.StringVar()

        # ----------------------------------------------------
        # Document vars
        # ----------------------------------------------------
        self.doc_section_var = tk.StringVar()
        self.doc_type_var = tk.StringVar(value="Other")
        self.doc_instruction_var = tk.StringVar()

        self.module_select_var = tk.StringVar()

        # self.product_select_var = tk.StringVar()

        self._build_ui()

    # ========================================================
    # UI
    # ========================================================
    def _bind_mousewheel(self, widget):
        def _on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        widget.bind_all("<MouseWheel>", _on_mousewheel)
    # def _build_ui(self):
    #     wrapper = ttk.Frame(self, padding=14)
    #     wrapper.pack(fill="both", expand=True)

    #     self._build_topbar(wrapper)

    #     self.module_summary_label = ttk.Label(
    #         wrapper,
    #         text="No module selected",
    #         style="Title.TLabel"
    #     )
    #     self.module_summary_label.pack(anchor="w", pady=(0, 10))

    #     paned = ttk.Panedwindow(wrapper, orient="horizontal")
    #     paned.pack(fill="both", expand=True)

    #     left = ttk.Frame(paned, padding=6)
    #     right = ttk.Frame(paned, padding=6)

    #     paned.add(left, weight=2)
    #     paned.add(right, weight=3)

    #     self._build_left_panel(left)
    #     self._build_right_panel(right)

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        self._build_topbar(wrapper)

        self.module_summary_label = ttk.Label(
            wrapper,
            text="No module selected",
            style="Title.TLabel"
        )
        self.module_summary_label.pack(anchor="w", pady=(0, 10))

        paned = ttk.Panedwindow(wrapper, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ----------------------------------------------------
        # LEFT: scrollable editor area
        # ----------------------------------------------------
        left_host = ttk.Frame(paned, padding=0)
        right = ttk.Frame(paned, padding=6)

        paned.add(left_host, weight=2)
        paned.add(right, weight=3)

        self.left_canvas = tk.Canvas(left_host, bg=AppConfig.COLOR_BG, highlightthickness=0)
        self.left_scrollbar = ttk.Scrollbar(left_host, orient="vertical", command=self.left_canvas.yview)
        self.left_scrollable_frame = ttk.Frame(self.left_canvas, padding=6)

        self.left_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        )

        self.left_canvas_window = self.left_canvas.create_window(
            (0, 0),
            window=self.left_scrollable_frame,
            anchor="nw"
        )

        def _resize_left_frame(event):
            self.left_canvas.itemconfig(self.left_canvas_window, width=event.width)

        self.left_canvas.bind("<Configure>", _resize_left_frame)
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)

        self.left_canvas.pack(side="left", fill="both", expand=True)
        self.left_scrollbar.pack(side="right", fill="y")

        self._build_left_panel(self.left_scrollable_frame)
        self._build_right_panel(right)
        self._bind_mousewheel(self.left_canvas)

    # def _build_topbar(self, parent):
    #     top = ttk.Frame(parent)
    #     top.pack(fill="x", pady=(0, 8))

    #     ttk.Button(top, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)

    #     ttk.Button(top, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)
    #     ttk.Button(top, text="Email PDF", command=self.email_module_pdf).pack(side="right", padx=2)
    #     ttk.Button(top, text="Generate PDF", command=self.generate_module_pdf).pack(side="right", padx=2)

    def _build_topbar(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 8))

        left = ttk.Frame(top)
        left.pack(side="left", fill="x", expand=True)

        ttk.Button(left, text="← Back to Dashboard", command=lambda: self.show_page("home")).pack(side="left", padx=2)

        ttk.Label(left, text="Open Module").pack(side="left", padx=(12, 4))

        self.module_select_combo = ttk.Combobox(
            left,
            textvariable=self.module_select_var,
            state="readonly",
            width=50
        )
        self.module_select_combo.pack(side="left", padx=4)
        self.module_select_combo.bind("<<ComboboxSelected>>", lambda e: self.open_selected_module_from_dropdown())

        ttk.Button(left, text="Open", command=self.open_selected_module_from_dropdown).pack(side="left", padx=4)

        right = ttk.Frame(top)
        right.pack(side="right")

        ttk.Button(right, text="Refresh", command=self.refresh_page).pack(side="right", padx=2)
        ttk.Button(right, text="Email PDF", command=self.email_module_pdf).pack(side="right", padx=2)
        ttk.Button(right, text="Generate PDF", command=self.generate_module_pdf).pack(side="right", padx=2)


    def refresh_module_selector(self):
        if not self.app.workbook_manager.has_workbook():
            self.module_select_combo["values"] = []
            return

        records = self.app.services.modules.list_modules()
        values = [
            f"{m.module_code} | {m.quote_ref} | {m.module_name}"
            for m in records
        ]
        set_combobox_values(self.module_select_combo, values, keep_current=True)

        selected_code = getattr(self.app, "selected_module_code", "").strip()
        if selected_code:
            for m in records:
                if m.module_code == selected_code:
                    self.module_select_var.set(f"{m.module_code} | {m.quote_ref} | {m.module_name}")
                    break 

    def refresh_module_selector(self):
        if not self.app.workbook_manager.has_workbook():
            self.module_select_combo["values"] = []
            return

        records = self.app.services.modules.list_modules()
        values = [
            f"{m.module_code} | {m.quote_ref} | {m.module_name}"
            for m in records
        ]
        set_combobox_values(self.module_select_combo, values, keep_current=True)

        selected_code = getattr(self.app, "selected_module_code", "").strip()
        if selected_code:
            for m in records:
                if m.module_code == selected_code:
                    self.module_select_var.set(f"{m.module_code} | {m.quote_ref} | {m.module_name}")
                    break


    def _build_left_panel(self, parent):
        self._build_module_details_card(parent)
        self._build_instruction_card(parent)
        self._build_task_editor_card(parent)
        self._build_component_editor_card(parent)
        self._build_document_editor_card(parent)

    def _build_right_panel(self, parent):
        tabs = ttk.Notebook(parent)
        tabs.pack(fill="both", expand=True)

        self.tasks_tab = ttk.Frame(tabs)
        self.components_tab = ttk.Frame(tabs)
        self.documents_tab = ttk.Frame(tabs)
        self.summary_tab = ttk.Frame(tabs)

        tabs.add(self.tasks_tab, text="Tasks / Subtasks")
        tabs.add(self.components_tab, text="Components / BOM")
        tabs.add(self.documents_tab, text="Documents")
        tabs.add(self.summary_tab, text="Summary")

        self._build_tasks_tab(self.tasks_tab)
        self._build_components_tab(self.components_tab)
        self._build_documents_tab(self.documents_tab)
        self._build_summary_tab(self.summary_tab)

    # --------------------------------------------------------
    # Left cards
    # --------------------------------------------------------

    def _build_module_details_card(self, parent):
        card = ttk.LabelFrame(parent, text="Module Details", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Module Code").grid(row=0, column=0, sticky="w", pady=4)
        self.module_code_entry = ttk.Entry(card, textvariable=self.module_code_var, state="readonly")
        self.module_code_entry.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Quote Ref").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.quote_ref_var).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Module Name").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.module_name_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Description").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.description_var).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Estimated Hours").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.estimated_hours_var).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Stock On Hand").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.stock_on_hand_var).grid(row=5, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.status_var_local,
            values=AppConfig.MODULE_STATUSES,
            state="readonly"
        ).grid(row=6, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        save_btn = ttk.Button(btn_row, text="Save Module", command=self.save_module)
        save_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(save_btn, "Create a new module or update the selected module.")

        del_btn = ttk.Button(btn_row, text="Delete Module", command=self.delete_module)
        del_btn.pack(side="left", fill="x", expand=True, padx=2)
        attach_tooltip(del_btn, "Delete the selected module and its linked module-level records.")

        card.columnconfigure(1, weight=1)

    # def _build_instruction_card(self, parent):
    #     card = ttk.LabelFrame(parent, text="Module Instructions", style="Card.TLabelframe", padding=12)
    #     card.pack(fill="x", pady=6)

    #     ttk.Label(card, text="Instruction / Build Notes").pack(anchor="w", pady=(0, 4))

    #     self.instruction_text = tk.Text(
    #         card,
    #         height=7,
    #         wrap="word",
    #         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
    #         bg="#FFFFFF",
    #         relief="solid",
    #         borderwidth=1,
    #     )
    #     self.instruction_text.pack(fill="x")

    def _build_instruction_card(self, parent):
        card = ttk.LabelFrame(parent, text="Module Instructions", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        top = ttk.Frame(card)
        top.pack(fill="x", pady=(0, 4))

        ttk.Label(top, text="Instruction / Build Notes").pack(side="left")

        ttk.Button(top, text="Clear", command=self.clear_instruction_text).pack(side="right")

        self.instruction_text = tk.Text(
            card,
            height=7,
            wrap="word",
            font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
            bg="#FFFFFF",
            relief="solid",
            borderwidth=1,
        )
        self.instruction_text.pack(fill="x")

    def clear_instruction_text(self):
         self.instruction_text.delete("1.0", tk.END)

    def _build_task_editor_card(self, parent):
        card = ttk.LabelFrame(parent, text="Task / Subtask Editor", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Task Name").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.task_name_var).grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Department").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.task_department_var,
            values=AppConfig.DEPARTMENTS,
            state="readonly"
        ).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Estimated Hours").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.task_hours_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Stage").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.task_stage_var,
            values=AppConfig.MODULE_EXEC_STAGES,
            state="readonly"
        ).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Status").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(
            card,
            textvariable=self.task_status_var,
            values=AppConfig.TASK_STATUSES,
            state="readonly"
        ).grid(row=4, column=1, sticky="ew", padx=6)

        self.subtask_check = ttk.Checkbutton(
            card,
            text="This is a subtask",
            variable=self.is_subtask_var,
            command=self._toggle_subtask_mode
        )
        self.subtask_check.grid(row=5, column=0, sticky="w", pady=4)

        ttk.Label(card, text="Parent Task").grid(row=6, column=0, sticky="w", pady=4)
        self.parent_task_combo = ttk.Combobox(card, textvariable=self.parent_task_display_var, state="disabled")
        self.parent_task_combo.grid(row=6, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Dependency Task").grid(row=7, column=0, sticky="w", pady=4)
        self.task_dependency_combo = ttk.Combobox(card, textvariable=self.task_dependency_display_var, state="readonly")
        self.task_dependency_combo.grid(row=7, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Notes").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.task_notes_var).grid(row=8, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Add Task", command=self.add_task).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Update Selected", command=self.update_selected_task).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_task).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Clear", command=self.clear_task_editor).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    def _build_component_editor_card(self, parent):
        card = ttk.LabelFrame(parent, text="Component / BOM Editor", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", pady=6)

        ttk.Label(card, text="Component").grid(row=0, column=0, sticky="w", pady=4)
        self.component_name_combo = ttk.Combobox(card, textvariable=self.component_name_var, state="normal")
        self.component_name_combo.grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Qty").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.component_qty_var).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="SOH Qty").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.component_soh_var).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Supplier").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.component_supplier_var).grid(row=3, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Lead Time (days)").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.component_lead_var).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Part Number").grid(row=5, column=0, sticky="w", pady=4)
        self.component_partno_combo = ttk.Combobox(card, textvariable=self.component_partno_var, state="normal")
        self.component_partno_combo.grid(row=5, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Notes").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(card, textvariable=self.component_notes_var).grid(row=6, column=1, sticky="ew", padx=6)

        btn_row = ttk.Frame(card)
        btn_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Add Component", command=self.add_component).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Update Selected", command=self.update_selected_component).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_component).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Clear", command=self.clear_component_editor).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)
        self._bind_component_part_search()
        self._refresh_component_part_dropdowns()

    def _collect_existing_parts_for_components(self):
        parts = []
        seen = set()

        try:
            if hasattr(self.app.services, "parts") and hasattr(self.app.services.parts, "list_parts"):
                source_parts = self.app.services.parts.list_parts()
            else:
                source_parts = self.app.services.modules.list_parts()

            for part in source_parts or []:
                part_name = norm_text(getattr(part, "component_name", "")) or norm_text(getattr(part, "part_name", ""))
                part_number = norm_text(getattr(part, "part_number", "")) or norm_text(getattr(part, "part_code", ""))
                key = (part_name.lower(), part_number.lower())
                if key in seen or not (part_name or part_number):
                    continue
                seen.add(key)
                parts.append({
                    "component_id": norm_text(getattr(part, "component_id", "")),
                    "name": part_name,
                    "part_number": part_number,
                    "qty": getattr(part, "qty", ""),
                    "soh": getattr(part, "soh_qty", ""),
                    "supplier": norm_text(getattr(part, "preferred_supplier", "")) or norm_text(getattr(part, "supplier", "")),
                    "lead": getattr(part, "lead_time_days", ""),
                    "notes": norm_text(getattr(part, "notes", "")),
                })
            return sorted(parts, key=lambda p: (p["name"].lower(), p["part_number"].lower()))
        except Exception:
            pass

        try:
            repo = getattr(self.app, "repo", None)
            if repo is None:
                return []
            rows = repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
            for row in rows or []:
                if norm_text(row.get("OwnerType")) != "PART":
                    continue
                part_name = norm_text(row.get("ComponentName")) or norm_text(row.get("component_name"))
                part_number = norm_text(row.get("PartNumber")) or norm_text(row.get("part_number"))
                key = (part_name.lower(), part_number.lower())
                if key in seen or not (part_name or part_number):
                    continue
                seen.add(key)
                parts.append({
                    "component_id": norm_text(row.get("ComponentID")),
                    "name": part_name,
                    "part_number": part_number,
                    "qty": row.get("Qty", ""),
                    "soh": row.get("SOHQty", ""),
                    "supplier": norm_text(row.get("PreferredSupplier")) or norm_text(row.get("preferred_supplier")),
                    "lead": row.get("LeadTimeDays", ""),
                    "notes": norm_text(row.get("Notes")) or norm_text(row.get("notes")),
                })
        except Exception:
            return []

        return sorted(parts, key=lambda p: (p["name"].lower(), p["part_number"].lower()))

    def _refresh_component_part_dropdowns(self):
        self._existing_component_parts = self._collect_existing_parts_for_components()
        name_values = sorted({p["name"] for p in self._existing_component_parts if p["name"]})
        partno_values = sorted({p["part_number"] for p in self._existing_component_parts if p["part_number"]})

        if hasattr(self, "component_name_combo"):
            self.component_name_combo["values"] = name_values
            self.component_name_combo._all_values = name_values
        if hasattr(self, "component_partno_combo"):
            self.component_partno_combo["values"] = partno_values
            self.component_partno_combo._all_values = partno_values

    def _filter_component_part_combo(self, combo, typed_text):
        values = list(getattr(combo, "_all_values", list(combo.cget("values"))))
        typed = norm_text(typed_text).lower()
        combo["values"] = [v for v in values if typed in str(v).lower()] if typed else values

    def _reset_component_part_combo(self, combo):
        combo["values"] = list(getattr(combo, "_all_values", list(combo.cget("values"))))

    def _apply_existing_component_part(self, part):
        self.selected_existing_part_id = part.get("component_id", "")
        self.component_name_var.set(part["name"])
        self.component_partno_var.set(part["part_number"])
        if not norm_text(self.component_qty_var.get()):
            self.component_qty_var.set(str(part["qty"] or "1"))
        self.component_soh_var.set(str(part["soh"] or ""))
        self.component_supplier_var.set(part["supplier"])
        self.component_lead_var.set(str(part["lead"] or ""))
        self.component_notes_var.set(part["notes"])

    def _merge_source_part_id(self, notes):
        source_part_id = norm_text(getattr(self, "selected_existing_part_id", ""))
        notes = norm_text(notes)
        if not source_part_id:
            return notes
        tag = f"SourcePartID={source_part_id}"
        if "SourcePartID=" in notes:
            return re.sub(r"SourcePartID=[^|]+", tag, notes).strip()
        return f"{notes} | {tag}" if notes else tag

    def _extract_source_part_id(self, notes):
        match = re.search(r"SourcePartID=([^|]+)", norm_text(notes))
        return norm_text(match.group(1)) if match else ""

    def _on_component_name_selected(self, event=None):
        selected_name = norm_text(self.component_name_var.get()).lower()
        selected_partno = norm_text(self.component_partno_var.get()).lower()
        fallback = None
        for part in getattr(self, "_existing_component_parts", []):
            if part["name"].lower() != selected_name:
                continue
            if fallback is None:
                fallback = part
            if selected_partno and part["part_number"].lower() == selected_partno:
                fallback = part
                break
        if fallback is not None:
            self._apply_existing_component_part(fallback)

    def _on_component_partno_selected(self, event=None):
        selected_partno = norm_text(self.component_partno_var.get()).lower()
        selected_name = norm_text(self.component_name_var.get()).lower()
        fallback = None
        for part in getattr(self, "_existing_component_parts", []):
            if part["part_number"].lower() != selected_partno:
                continue
            if fallback is None:
                fallback = part
            if selected_name and part["name"].lower() == selected_name:
                fallback = part
                break
        if fallback is not None:
            self._apply_existing_component_part(fallback)

    def _bind_component_part_search(self):
        if not hasattr(self, "component_name_combo") or not hasattr(self, "component_partno_combo"):
            return

        self.component_name_combo.bind("<<ComboboxSelected>>", self._on_component_name_selected)
        self.component_partno_combo.bind("<<ComboboxSelected>>", self._on_component_partno_selected)
        self.component_name_combo.bind(
            "<KeyRelease>",
            lambda e: self._filter_component_part_combo(self.component_name_combo, self.component_name_var.get())
        )
        self.component_partno_combo.bind(
            "<KeyRelease>",
            lambda e: self._filter_component_part_combo(self.component_partno_combo, self.component_partno_var.get())
        )
        self.component_name_combo.configure(postcommand=lambda: self._reset_component_part_combo(self.component_name_combo))
        self.component_partno_combo.configure(postcommand=lambda: self._reset_component_part_combo(self.component_partno_combo))
        self.component_name_combo.bind("<FocusIn>", lambda e: self._reset_component_part_combo(self.component_name_combo), add="+")
        self.component_partno_combo.bind("<FocusIn>", lambda e: self._reset_component_part_combo(self.component_partno_combo), add="+")

    def _build_document_editor_card(self, parent):
        card = ttk.LabelFrame(parent, text="Document Upload", style="Card.TLabelframe", padding=12)
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

        # btn_row = ttk.Frame(card)
        # btn_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        # ttk.Button(btn_row, text="Add File", command=self.add_document).pack(side="left", fill="x", expand=True, padx=2)
        # ttk.Button(btn_row, text="Open Selected", command=self.open_selected_document).pack(side="left", fill="x", expand=True, padx=2)
        # ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_document).pack(side="left", fill="x", expand=True, padx=2)
        btn_row = ttk.Frame(card)
        btn_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Add File", command=self.add_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Open Selected", command=self.open_selected_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Delete Selected", command=self.delete_selected_document).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(btn_row, text="Clear Fields", command=self.clear_document_editor).pack(side="left", fill="x", expand=True, padx=2)

        card.columnconfigure(1, weight=1)

    # --------------------------------------------------------
    # Right tabs
    # --------------------------------------------------------
    def clear_document_editor(self):
        self.doc_section_var.set("")
        self.doc_type_var.set("Other")
        self.doc_instruction_var.set("")


    def _build_tasks_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        self.task_tree = ttk.Treeview(
            wrap,
            columns=("Hours", "Department", "Status", "Notes"),
            show="tree headings"
        )
        self.task_tree.heading("#0", text="Task Hierarchy")
        self.task_tree.heading("Hours", text="Hours")
        self.task_tree.heading("Department", text="Department")
        self.task_tree.heading("Status", text="Status")
        self.task_tree.heading("Notes", text="Notes")

        self.task_tree.column("#0", width=420, anchor="w")
        self.task_tree.column("Hours", width=90, anchor="center")
        self.task_tree.column("Department", width=140, anchor="w")
        self.task_tree.column("Status", width=120, anchor="w")
        self.task_tree.column("Notes", width=300, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=sb.set)

        self.task_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.task_tree.bind("<ButtonRelease-1>", lambda e: self.load_selected_task())
        self.task_tree.bind("<Double-1>", lambda e: self.load_selected_task())

        self.total_hours_label = ttk.Label(wrap, text="Module Task Hours: 0.00 hrs", style="Sub.TLabel")
        self.total_hours_label.pack(anchor="e", pady=(8, 0))

    def _build_components_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("component_name", "qty", "soh_qty", "supplier", "lead_time", "part_number", "notes")
        self.component_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col_id, heading_text, width in [
            ("component_name", "Component", 220),
            ("qty", "Qty", 70),
            ("soh_qty", "SOH Qty", 90),
            ("supplier", "Supplier", 160),
            ("lead_time", "Lead Time", 90),
            ("part_number", "Part Number", 140),
            ("notes", "Notes", 260),
        ]:
            self.component_tree.heading(col_id, text=heading_text)
            self.component_tree.column(col_id, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.component_tree.yview)
        self.component_tree.configure(yscrollcommand=sb.set)

        self.component_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.component_tree.bind("<ButtonRelease-1>", lambda e: self.load_selected_component())
        self.component_tree.bind("<Double-1>", lambda e: self.load_selected_component())

    def _build_documents_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        cols = ("SectionName", "DocName", "DocType", "FilePath", "AddedOn")
        self.document_tree = ttk.Treeview(wrap, columns=cols, show="headings")

        for col, width in [
            ("SectionName", 150),
            ("DocName", 220),
            ("DocType", 150),
            ("FilePath", 420),
            ("AddedOn", 160),
        ]:
            self.document_tree.heading(col, text=col)
            self.document_tree.column(col, width=width, anchor="w")

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.document_tree.yview)
        self.document_tree.configure(yscrollcommand=sb.set)

        self.document_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _build_summary_tab(self, parent):
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        self.summary_text = make_readonly_text(wrap, height=26)
        self.summary_text.pack(fill="both", expand=True)

    # ========================================================
    # Page refresh / selection
    # ========================================================

    def refresh_page(self):


        if not self.require_workbook():
            return

        module_code = getattr(self.app, "selected_module_code", "").strip()

        self.refresh_module_selector()
        self._refresh_component_part_dropdowns()

        if not module_code:
            self._clear_all_views()
            self.module_summary_label.config(text="No module selected")
            return

        try:
            bundle = self.app.services.modules.get_module_bundle(module_code)
            if not bundle.module:
                self._clear_all_views()
                self.module_summary_label.config(text="Module not found")
                return

            self._load_module_into_form(bundle.module)
            self._load_tasks(bundle.tasks or [])
            self._load_components(bundle.components or [])
            self._load_documents(bundle.documents or [])
            self._load_summary(bundle)
            self._refresh_task_dropdowns(bundle.tasks or [])

            self.module_summary_label.config(
                text=f"📦 {bundle.module.module_name} | Quote: {bundle.module.quote_ref} | Code: {bundle.module.module_code}"
            )
            self.set_status(f"Loaded module: {bundle.module.module_code}")

        except Exception as exc:
            show_error("Module Refresh Error", str(exc))

    def _clear_all_views(self):
        self.module_code_var.set("")
        self.quote_ref_var.set("")
        self.module_name_var.set("")
        self.description_var.set("")
        self.estimated_hours_var.set("")
        self.stock_on_hand_var.set("")
        self.status_var_local.set("Draft")

        self.instruction_text.delete("1.0", tk.END)

        treeview_clear(self.task_tree)
        treeview_clear(self.component_tree)
        treeview_clear(self.document_tree)
        set_text_readonly(self.summary_text, "")

        self.clear_task_editor()
        self.clear_component_editor()

    def _load_module_into_form(self, module: ModuleRecord):
        self.module_code_var.set(module.module_code)
        self.quote_ref_var.set(module.quote_ref)
        self.module_name_var.set(module.module_name)
        self.description_var.set(module.description)
        self.estimated_hours_var.set(f"{float(module.estimated_hours or 0.0):.2f}")
        self.stock_on_hand_var.set(f"{float(module.stock_on_hand or 0.0):.2f}")
        self.status_var_local.set(module.status or "Draft")

        self.instruction_text.delete("1.0", tk.END)
        self.instruction_text.insert("1.0", module.instruction_text or "")

    def _load_tasks(self, tasks: List[TaskRecord]):
        treeview_clear(self.task_tree)

        task_map = {t.task_id: t for t in tasks}
        node_map = {}
        total_hours = 0.0

        # top level first
        for task in tasks:
            if not norm_text(task.parent_task_id):
                node_id = self.task_tree.insert(
                    "",
                    "end",
                    text=task.task_name,
                    values=(f"{float(task.estimated_hours or 0.0):.2f}", task.department, task.status, task.notes),
                    tags=(task.task_id,),
                )
                node_map[task.task_id] = node_id
                total_hours += float(task.estimated_hours or 0.0)

        # subtasks
        for task in tasks:
            if norm_text(task.parent_task_id):
                parent_node = node_map.get(task.parent_task_id, "")
                self.task_tree.insert(
                    parent_node,
                    "end",
                    text=task.task_name,
                    values=(f"{float(task.estimated_hours or 0.0):.2f}", task.department, task.status, task.notes),
                    tags=(task.task_id,),
                )
                total_hours += float(task.estimated_hours or 0.0)

        self.total_hours_label.config(text=f"Module Task Hours: {total_hours:.2f} hrs")

    def _load_components(self, components: List[ComponentRecord]):
        treeview_clear(self.component_tree)
        for comp in components:
            self.component_tree.insert(
                "",
                "end",
                values=(
                    comp.component_name,
                    comp.qty,
                    comp.soh_qty,
                    comp.preferred_supplier,
                    comp.lead_time_days,
                    comp.part_number,
                    comp.notes,
                ),
                tags=(comp.component_id,),
            )

    def _load_documents(self, docs: List[DocumentRecord]):
        treeview_clear(self.document_tree)
        for doc in docs:
            self.document_tree.insert(
                "",
                "end",
                values=(
                    doc.section_name,
                    doc.doc_name,
                    doc.doc_type,
                    doc.file_path,
                    doc.added_on,
                ),
                tags=(doc.doc_id,),
            )

    def _load_summary(self, bundle):
        module = bundle.module
        tasks = bundle.tasks or []
        components = bundle.components or []
        documents = bundle.documents or []

        total_hours = sum(float(t.estimated_hours or 0.0) for t in tasks)

        lines = [
            f"Module Code: {module.module_code}",
            f"Quote Ref: {module.quote_ref}",
            f"Module Name: {module.module_name}",
            f"Description: {module.description}",
            f"Status: {module.status}",
            f"Estimated Hours (Module Header): {float(module.estimated_hours or 0.0):.2f}",
            f"Stock On Hand: {float(module.stock_on_hand or 0.0):.2f}",
            "",
            "Instruction Text:",
            module.instruction_text or "-",
            "",
            f"Task Count: {len(tasks)}",
            f"Component Count: {len(components)}",
            f"Document Count: {len(documents)}",
            f"Calculated Task Hours: {total_hours:.2f}",
            "",
            "Top Level Tasks:",
        ]

        top_tasks = [t for t in tasks if not norm_text(t.parent_task_id)]
        if not top_tasks:
            lines.append(" - None")
        else:
            for t in top_tasks:
                dep_text = f" | Depends on: {t.dependency_task_id}" if norm_text(t.dependency_task_id) else ""
                lines.append(f" - {t.task_name} ({float(t.estimated_hours or 0.0):.2f} hrs){dep_text}")

        set_text_readonly(self.summary_text, "\n".join(lines))

    def _refresh_task_dropdowns(self, tasks: List[TaskRecord]):
        self.parent_task_map = {t.task_name: t.task_id for t in tasks if not norm_text(t.parent_task_id)}
        parent_names = list(self.parent_task_map.keys())

        set_combobox_values(self.parent_task_combo, parent_names, keep_current=True)
        if self.is_subtask_var.get():
            self.parent_task_combo.config(state="readonly")
        else:
            self.parent_task_combo.config(state="disabled")

        # dependency list includes all task names
        dep_names = [t.task_name for t in tasks]
        set_combobox_values(self.task_dependency_combo, dep_names, keep_current=True)

    # ========================================================
    # Module CRUD
    # ========================================================

    def save_module(self):
        if not self.require_workbook():
            return

        try:
            quote_ref = self.quote_ref_var.get().strip()
            module_name = Validators.require_text(self.module_name_var.get(), "Module name")
            description = self.description_var.get().strip()
            estimated_hours = Validators.parse_float(self.estimated_hours_var.get(), "Estimated hours", default=0.0)
            stock_on_hand = Validators.parse_float(self.stock_on_hand_var.get(), "Stock On Hand", default=0.0)
            instruction_text = self.instruction_text.get("1.0", tk.END).strip()
            status = self.status_var_local.get().strip() or "Draft"

            existing_code = self.module_code_var.get().strip() or None

            new_code = self.app.services.modules.create_or_update_module(
                quote_ref=quote_ref,
                module_name=module_name,
                description=description,
                instruction_text=instruction_text,
                estimated_hours=estimated_hours,
                stock_on_hand=stock_on_hand,
                status=status,
                existing_module_code=existing_code,
            )

            self.app.set_selected_module(new_code)
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Module saved: {new_code}")

        except Exception as exc:
            show_error("Save Module Error", str(exc))

    def delete_module(self):
        if not self.require_workbook():
            return

        module_code = self.module_code_var.get().strip()
        if not module_code:
            show_warning("No Module", "No module selected.")
            return

        if not tk.messagebox.askyesno(
            "Delete Module",
            "Delete this module?\n\nThis will also delete linked module tasks, components, and module docs records."
        ):
            return

        try:
            self.app.services.modules.delete_module(module_code, delete_docs_files=False)
            self.app.set_selected_module("")
            self.refresh_page()
            self.app.refresh_home_page()
            self.set_status(f"Module deleted: {module_code}")
        except Exception as exc:
            show_error("Delete Module Error", str(exc))

    # ========================================================
    # Task CRUD
    # ========================================================

    def _toggle_subtask_mode(self):
        if self.is_subtask_var.get():
            self.parent_task_combo.config(state="readonly")
        else:
            self.parent_task_display_var.set("")
            self.parent_task_combo.config(state="disabled")

    def clear_task_editor(self):
        self.current_task_id = None
        self.task_name_var.set("")
        self.task_department_var.set("")
        self.task_hours_var.set("")
        self.task_stage_var.set("")
        self.task_status_var.set("Not Started")
        self.task_notes_var.set("")
        self.task_dependency_display_var.set("")
        self.parent_task_display_var.set("")
        self.is_subtask_var.set(False)
        self._toggle_subtask_mode()

    def _selected_module_code(self) -> str:
        return self.module_code_var.get().strip() or getattr(self.app, "selected_module_code", "").strip()

    def _get_task_id_by_name(self, task_name: str) -> str:
        for display_name, task_id in self.parent_task_map.items():
            if display_name == task_name:
                return task_id

        # fallback full search
        module_code = self._selected_module_code()
        tasks = self.app.services.modules.get_module_tasks(module_code)
        for t in tasks:
            if t.task_name == task_name:
                return t.task_id
        return ""

    def add_task(self):
        if not self.require_workbook():
            return

        module_code = self._selected_module_code()
        if not module_code:
            show_warning("No Module", "Select or create a module first.")
            return

        try:
            task_name = Validators.require_text(self.task_name_var.get(), "Task name")
            department = self.task_department_var.get().strip()
            estimated_hours = Validators.parse_float(self.task_hours_var.get(), "Estimated hours", default=0.0)
            stage = self.task_stage_var.get().strip()
            status = self.task_status_var.get().strip() or "Not Started"
            notes = self.task_notes_var.get().strip()

            parent_task_id = ""
            if self.is_subtask_var.get():
                parent_display = self.parent_task_display_var.get().strip()
                if not parent_display or parent_display not in self.parent_task_map:
                    raise ValueError("Select a valid parent task.")
                parent_task_id = self.parent_task_map[parent_display]

            dependency_task_id = ""
            dep_display = self.task_dependency_display_var.get().strip()
            if dep_display:
                dependency_task_id = self._get_task_id_by_name(dep_display)

            self.app.services.modules.add_module_task(
                module_code=module_code,
                task_name=task_name,
                department=department,
                estimated_hours=estimated_hours,
                parent_task_id=parent_task_id,
                dependency_task_id=dependency_task_id,
                stage=stage,
                status=status,
                notes=notes,
            )

            self.clear_task_editor()
            self.refresh_page()
            self.set_status("Task added successfully.")

        except Exception as exc:
            show_error("Add Task Error", str(exc))

    def load_selected_task(self):
        sel = self.task_tree.selection()
        if not sel:
            return

        tags = self.task_tree.item(sel[0], "tags")
        if not tags:
            return

        task_id = tags[0]
        module_code = self._selected_module_code()
        tasks = self.app.services.modules.get_module_tasks(module_code)

        selected_task = None
        for t in tasks:
            if t.task_id == task_id:
                selected_task = t
                break

        if not selected_task:
            return

        self.current_task_id = selected_task.task_id
        self.task_name_var.set(selected_task.task_name)
        self.task_department_var.set(selected_task.department)
        self.task_hours_var.set(f"{float(selected_task.estimated_hours or 0.0):.2f}")
        self.task_stage_var.set(selected_task.stage)
        self.task_status_var.set(selected_task.status or "Not Started")
        self.task_notes_var.set(selected_task.notes)

        if norm_text(selected_task.parent_task_id):
            self.is_subtask_var.set(True)
            self._toggle_subtask_mode()
            for name, tid in self.parent_task_map.items():
                if tid == selected_task.parent_task_id:
                    self.parent_task_display_var.set(name)
                    break
        else:
            self.is_subtask_var.set(False)
            self._toggle_subtask_mode()
            self.parent_task_display_var.set("")

        dep_name = ""
        if norm_text(selected_task.dependency_task_id):
            for t in tasks:
                if t.task_id == selected_task.dependency_task_id:
                    dep_name = t.task_name
                    break
        self.task_dependency_display_var.set(dep_name)

    def update_selected_task(self):
        if not self.require_workbook():
            return

        if not self.current_task_id:
            show_warning("No Selection", "Select a task first.")
            return

        try:
            task_name = Validators.require_text(self.task_name_var.get(), "Task name")
            department = self.task_department_var.get().strip()
            estimated_hours = Validators.parse_float(self.task_hours_var.get(), "Estimated hours", default=0.0)
            stage = self.task_stage_var.get().strip()
            status = self.task_status_var.get().strip() or "Not Started"
            notes = self.task_notes_var.get().strip()

            parent_task_id = ""
            if self.is_subtask_var.get():
                parent_display = self.parent_task_display_var.get().strip()
                if not parent_display or parent_display not in self.parent_task_map:
                    raise ValueError("Select a valid parent task.")
                parent_task_id = self.parent_task_map[parent_display]

            dependency_task_id = ""
            dep_display = self.task_dependency_display_var.get().strip()
            if dep_display:
                dependency_task_id = self._get_task_id_by_name(dep_display)

            # prevent self-parent / self-dependency
            if self.current_task_id == parent_task_id:
                raise ValueError("A task cannot be its own parent.")
            if self.current_task_id == dependency_task_id:
                raise ValueError("A task cannot depend on itself.")

            self.app.services.modules.update_task(
                self.current_task_id,
                {
                    "TaskName": task_name,
                    "Department": department,
                    "EstimatedHours": estimated_hours,
                    "ParentTaskID": parent_task_id,
                    "DependencyTaskID": dependency_task_id,
                    "Stage": stage,
                    "Status": status,
                    "Notes": notes,
                }
            )

            self.refresh_page()
            self.set_status("Task updated successfully.")

        except Exception as exc:
            show_error("Update Task Error", str(exc))

    def delete_selected_task(self):
        if not self.require_workbook():
            return

        if not self.current_task_id:
            show_warning("No Selection", "Select a task first.")
            return

        if not tk.messagebox.askyesno("Delete Task", "Delete the selected task?"):
            return

        try:
            self.app.services.modules.delete_task(self.current_task_id)
            self.clear_task_editor()
            self.refresh_page()
            self.set_status("Task deleted successfully.")
        except Exception as exc:
            show_error("Delete Task Error", str(exc))

    # ========================================================
    # Component CRUD
    # ========================================================

    def clear_component_editor(self):
        self.current_component_id = None
        self.selected_existing_part_id = ""
        self.component_name_var.set("")
        self.component_qty_var.set("")
        self.component_soh_var.set("")
        self.component_supplier_var.set("")
        self.component_lead_var.set("")
        self.component_partno_var.set("")
        self.component_notes_var.set("")

    def add_component(self):
        if not self.require_workbook():
            return

        module_code = self._selected_module_code()
        if not module_code:
            show_warning("No Module", "Select or create a module first.")
            return

        try:
            component_name = Validators.require_text(self.component_name_var.get(), "Component name")
            qty = Validators.parse_float(self.component_qty_var.get(), "Qty", default=0.0)
            soh_qty = Validators.parse_float(self.component_soh_var.get(), "SOH Qty", default=0.0)
            lead_time_days = Validators.parse_int(self.component_lead_var.get(), "Lead time", default=0)

            self.app.services.modules.add_module_component(
                module_code=module_code,
                component_name=component_name,
                qty=qty,
                soh_qty=soh_qty,
                preferred_supplier=self.component_supplier_var.get().strip(),
                lead_time_days=lead_time_days,
                part_number=self.component_partno_var.get().strip(),
                notes=self._merge_source_part_id(self.component_notes_var.get()),
            )

            self.clear_component_editor()
            self.refresh_page()
            self.set_status("Component added successfully.")

        except Exception as exc:
            show_error("Add Component Error", str(exc))

    def load_selected_component(self):
        sel = self.component_tree.selection()
        if not sel:
            return

        tags = self.component_tree.item(sel[0], "tags")
        if not tags:
            return

        component_id = tags[0]
        module_code = self._selected_module_code()
        components = self.app.services.modules.get_module_components(module_code)

        selected = None
        for c in components:
            if c.component_id == component_id:
                selected = c
                break

        if not selected:
            return

        self.current_component_id = selected.component_id
        self.component_name_var.set(selected.component_name)
        self.component_qty_var.set(str(selected.qty))
        self.component_soh_var.set(str(selected.soh_qty))
        self.component_supplier_var.set(selected.preferred_supplier)
        self.component_lead_var.set(str(selected.lead_time_days))
        self.component_partno_var.set(selected.part_number)
        self.component_notes_var.set(selected.notes)
        self.selected_existing_part_id = self._extract_source_part_id(selected.notes)

    def update_selected_component(self):
        if not self.require_workbook():
            return

        if not self.current_component_id:
            show_warning("No Selection", "Select a component first.")
            return

        try:
            component_name = Validators.require_text(self.component_name_var.get(), "Component name")
            qty = Validators.parse_float(self.component_qty_var.get(), "Qty", default=0.0)
            soh_qty = Validators.parse_float(self.component_soh_var.get(), "SOH Qty", default=0.0)
            lead_time_days = Validators.parse_int(self.component_lead_var.get(), "Lead time", default=0)

            self.app.services.modules.update_component(
                self.current_component_id,
                {
                    "ComponentName": component_name,
                    "Qty": qty,
                    "SOHQty": soh_qty,
                    "PreferredSupplier": self.component_supplier_var.get().strip(),
                    "LeadTimeDays": lead_time_days,
                    "PartNumber": self.component_partno_var.get().strip(),
                    "Notes": self._merge_source_part_id(self.component_notes_var.get()),
                }
            )

            self.refresh_page()
            self.set_status("Component updated successfully.")

        except Exception as exc:
            show_error("Update Component Error", str(exc))

    def delete_selected_component(self):
        if not self.require_workbook():
            return

        if not self.current_component_id:
            show_warning("No Selection", "Select a component first.")
            return

        if not tk.messagebox.askyesno("Delete Component", "Delete the selected component?"):
            return

        try:
            self.app.services.modules.delete_component(self.current_component_id)
            self.clear_component_editor()
            self.refresh_page()
            self.set_status("Component deleted successfully.")
        except Exception as exc:
            show_error("Delete Component Error", str(exc))

    # ========================================================
    # Documents CRUD
    # ========================================================

    def add_document(self):
        if not self.require_workbook():
            return

        module_code = self._selected_module_code()
        if not module_code:
            show_warning("No Module", "Select or create a module first.")
            return

        file_path = filedialog.askopenfilename(
            title="Select Module Document",
            filetypes=[
                ("Documents", "*.pdf *.dwg *.dxf *.step *.stp *.sldprt *.sldasm *.doc *.docx *.xls *.xlsx *.png *.jpg *.jpeg"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        try:
            self.app.services.modules.add_module_document(
                module_code=module_code,
                source_file_path=file_path,
                section_name=self.doc_section_var.get().strip(),
                doc_type=self.doc_type_var.get().strip() or "Other",
                instruction_text=self.doc_instruction_var.get().strip(),
                copy_file=True,
            )
            self.refresh_page()
            self.set_status("Document added successfully.")

        except Exception as exc:
            show_error("Add Document Error", str(exc))

    def _get_selected_document_info(self):
        sel = self.document_tree.selection()
        if not sel:
            return None, None

        tags = self.document_tree.item(sel[0], "tags")
        vals = self.document_tree.item(sel[0], "values")

        doc_id = tags[0] if tags else ""
        file_path = vals[3] if vals and len(vals) > 3 else ""
        return doc_id, file_path
    
    def open_selected_module_from_dropdown(self):
        selected = self.module_select_var.get().strip()
        if not selected:
            return

        try:
            module_code = selected.split("|")[0].strip()
            self.app.set_selected_module(module_code)
            self.refresh_page()
        except Exception as exc:
            show_error("Open Module Error", str(exc))

    def refresh_module_selector(self):
        if not self.app.workbook_manager.has_workbook():
            self.module_select_combo["values"] = []
            return

        records = self.app.services.modules.list_modules()

        values = [
            f"{m.module_code} | {m.quote_ref} | {m.module_name}"
            for m in records
        ]

        set_combobox_values(self.module_select_combo, values, keep_current=True)

        selected_code = getattr(self.app, "selected_module_code", "").strip()
        if selected_code:
            for m in records:
                if m.module_code == selected_code:
                    self.module_select_var.set(
                        f"{m.module_code} | {m.quote_ref} | {m.module_name}"
                    )
                    break

    def open_selected_document(self):
        doc_id, file_path = self._get_selected_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a document first.")
            return

        try:
            file_path = self.app.services.modules.resolve_document_open_path(doc_id, str(file_path))
            open_file_with_default_app(str(file_path))
        except Exception as exc:
            show_error("Open Document Error", str(exc))

    def delete_selected_document(self):
        if not self.require_workbook():
            return

        doc_id, _file_path = self._get_selected_document_info()
        if not doc_id:
            show_warning("No Selection", "Select a document first.")
            return

        if not tk.messagebox.askyesno("Delete Document", "Delete the selected document record?"):
            return

        try:
            self.app.services.modules.delete_document(doc_id, delete_file=False)
            self.refresh_page()
            self.set_status("Document deleted successfully.")
        except Exception as exc:
            show_error("Delete Document Error", str(exc))

    # ========================================================
    # PDF / email hooks
    # ========================================================

    def generate_module_pdf(self):
        module_code = self._selected_module_code()
        if not module_code:
            show_warning("No Module", "Select a module first.")
            return

        # Hook for reports.py later
        if not hasattr(self.app, "reports"):
            show_warning("Not Ready", "Report service is not wired yet. We will connect it when reports.py and main.py are added.")
            return

        try:
            self.app.reports.generate_module_report_dialog(module_code)
        except Exception as exc:
            show_error("Generate PDF Error", str(exc))

    def email_module_pdf(self):
        module_code = self._selected_module_code()
        if not module_code:
            show_warning("No Module", "Select a module first.")
            return

        if not hasattr(self.app, "mailer") or not hasattr(self.app, "reports"):
            show_warning("Not Ready", "Mailer/report services are not wired yet. We will connect them in the final files.")
            return

        try:
            self.app.reports.email_module_report_dialog(module_code)
        except Exception as exc:
            show_error("Email PDF Error", str(exc))
