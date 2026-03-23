# ============================================================
# ui_common.py
# Shared UI utilities for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from app_config import AppConfig


# ============================================================
# Generic helpers
# ============================================================

def norm_text(value: Any) -> str:
    return str(value or "").strip()


def set_combobox_values(combo: ttk.Combobox, values: Sequence[str], keep_current: bool = True) -> None:
    current = norm_text(combo.get()) if keep_current else ""
    combo["values"] = list(values)
    if keep_current and current and current in values:
        combo.set(current)
    elif values:
        combo.set(values[0])


def treeview_clear(tree: ttk.Treeview) -> None:
    for item in tree.get_children():
        tree.delete(item)


def listbox_clear(lb: tk.Listbox) -> None:
    lb.delete(0, tk.END)


def ask_yes_no(title: str, text: str) -> bool:
    return messagebox.askyesno(title, text)


def show_info(title: str, text: str) -> None:
    messagebox.showinfo(title, text)


def show_warning(title: str, text: str) -> None:
    messagebox.showwarning(title, text)


def show_error(title: str, text: str) -> None:
    messagebox.showerror(title, text)


def open_file_with_default_app(path: str) -> None:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    os.startfile(path)


# ============================================================
# Tooltip
# ============================================================

class ToolTip:
    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self.tip = None

        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip or not self.text:
            return

        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6

        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#FFF8D8",
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=5,
            font=(AppConfig.FONT_FAMILY, 9),
            wraplength=320,
        )
        label.pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


def attach_tooltip(widget, text: str) -> ToolTip:
    return ToolTip(widget, text)


# ============================================================
# Drag-drop listbox
# ============================================================

class DraggableListbox(tk.Listbox):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.cur_index: Optional[int] = None
        self.bind("<Button-1>", self._on_click)
        self.bind("<B1-Motion>", self._on_drag)

    def _on_click(self, event):
        self.cur_index = self.nearest(event.y)

    def _on_drag(self, event):
        i = self.nearest(event.y)
        if self.cur_index is None or i == self.cur_index:
            return

        txt = self.get(self.cur_index)
        self.delete(self.cur_index)
        self.insert(i, txt)

        self.selection_clear(0, tk.END)
        self.selection_set(i)
        self.cur_index = i

    def get_all(self) -> List[str]:
        return list(self.get(0, tk.END))


# ============================================================
# Styled text viewer / editor helpers
# ============================================================

def make_readonly_text(parent, height: int = 10, wrap: str = "word") -> tk.Text:
    txt = tk.Text(
        parent,
        height=height,
        wrap=wrap,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
        bg="#FFFFFF",
        fg="#000000",
        insertbackground="#000000",
        relief="solid",
        borderwidth=1,
    )
    txt.config(state="disabled")
    return txt


def set_text_value(widget: tk.Text, value: str) -> None:
    widget.config(state="normal")
    widget.delete("1.0", tk.END)
    widget.insert("1.0", value or "")
    widget.config(state="normal")


def set_text_readonly(widget: tk.Text, value: str) -> None:
    widget.config(state="normal")
    widget.delete("1.0", tk.END)
    widget.insert("1.0", value or "")
    widget.config(state="disabled")


def get_text_value(widget: tk.Text) -> str:
    return widget.get("1.0", tk.END).strip()


# ============================================================
# Form builders
# ============================================================

class FormBuilder:
    @staticmethod
    def add_label(parent, text: str, row: int, column: int = 0, padx: int = 6, pady: int = 4, sticky: str = "w"):
        lbl = ttk.Label(parent, text=text)
        lbl.grid(row=row, column=column, padx=padx, pady=pady, sticky=sticky)
        return lbl

    @staticmethod
    def add_entry(parent, variable, row: int, column: int = 1, width: int = 24, padx: int = 6, pady: int = 4, sticky: str = "ew"):
        ent = ttk.Entry(parent, textvariable=variable, width=width)
        ent.grid(row=row, column=column, padx=padx, pady=pady, sticky=sticky)
        return ent

    @staticmethod
    def add_combobox(
        parent,
        variable,
        values: Sequence[str],
        row: int,
        column: int = 1,
        width: int = 24,
        state: str = "readonly",
        padx: int = 6,
        pady: int = 4,
        sticky: str = "ew",
    ):
        combo = ttk.Combobox(parent, textvariable=variable, values=list(values), width=width, state=state)
        combo.grid(row=row, column=column, padx=padx, pady=pady, sticky=sticky)
        return combo

    @staticmethod
    def add_button(parent, text: str, command, row: int, column: int = 0, columnspan: int = 1, padx: int = 4, pady: int = 6, sticky: str = "ew"):
        btn = ttk.Button(parent, text=text, command=command)
        btn.grid(row=row, column=column, columnspan=columnspan, padx=padx, pady=pady, sticky=sticky)
        return btn


# ============================================================
# Treeview builder
# ============================================================

class TreeBuilder:
    @staticmethod
    def create_tree(parent, columns: Sequence[Tuple[str, int]], show: str = "headings") -> Tuple[ttk.Treeview, ttk.Scrollbar]:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show=show)
        for col_name, width in columns:
            tree.heading(col_name, text=col_name)
            tree.column(col_name, width=width, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)

        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        return tree, sb

    @staticmethod
    def create_hierarchical_task_tree(parent) -> Tuple[ttk.Treeview, ttk.Scrollbar]:
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=("Hours", "Department", "Status", "Notes"), show="tree headings")
        tree.heading("#0", text="Task Hierarchy")
        tree.heading("Hours", text="Hours")
        tree.heading("Department", text="Department")
        tree.heading("Status", text="Status")
        tree.heading("Notes", text="Notes")

        tree.column("#0", width=380, anchor="w")
        tree.column("Hours", width=90, anchor="center")
        tree.column("Department", width=140, anchor="w")
        tree.column("Status", width=120, anchor="w")
        tree.column("Notes", width=280, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)

        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        return tree, sb


# ============================================================
# Status bar mixin
# ============================================================

class StatusMixin:
    def set_status(self, text: str) -> None:
        if hasattr(self, "status_var"):
            self.status_var.set(text)

    def clear_status(self) -> None:
        if hasattr(self, "status_var"):
            self.status_var.set("")


# ============================================================
# Base page frame
# ============================================================

class BasePage(ttk.Frame, StatusMixin):
    """
    Common base for all pages.

    Expects app to provide:
    - app.repo
    - app.services
    - app.show_page(page_name)
    - app.require_workbook()
    - app.set_status(text)
    """

    def __init__(self, master, app, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app

    def require_workbook(self) -> bool:
        return self.app.require_workbook()

    def set_status(self, text: str) -> None:
        self.app.set_status(text)

    def show_page(self, page_name: str) -> None:
        self.app.show_page(page_name)

    def refresh_page(self) -> None:
        """
        Override in child page.
        """
        pass


# ============================================================
# Common dialog helpers
# ============================================================

class Dialogs:
    @staticmethod
    def ask_string(title: str, prompt: str, initialvalue: Optional[str] = None) -> Optional[str]:
        return simpledialog.askstring(title, prompt, initialvalue=initialvalue)

    @staticmethod
    def ask_int(title: str, prompt: str, minvalue: Optional[int] = None, maxvalue: Optional[int] = None, initialvalue: Optional[int] = None) -> Optional[int]:
        return simpledialog.askinteger(
            title,
            prompt,
            minvalue=minvalue,
            maxvalue=maxvalue,
            initialvalue=initialvalue
        )

    @staticmethod
    def ask_float(title: str, prompt: str, minvalue: Optional[float] = None, maxvalue: Optional[float] = None, initialvalue: Optional[float] = None) -> Optional[float]:
        return simpledialog.askfloat(
            title,
            prompt,
            minvalue=minvalue,
            maxvalue=maxvalue,
            initialvalue=initialvalue
        )


# ============================================================
# Validation helpers
# ============================================================

class Validators:
    @staticmethod
    def require_text(value: str, label: str) -> str:
        value = norm_text(value)
        if not value:
            raise ValueError(f"{label} is required.")
        return value

    @staticmethod
    def parse_float(value: Any, label: str, default: float = 0.0) -> float:
        txt = norm_text(value)
        if not txt:
            return default
        try:
            return float(txt)
        except Exception:
            raise ValueError(f"{label} must be numeric.")

    @staticmethod
    def parse_int(value: Any, label: str, default: int = 0) -> int:
        txt = norm_text(value)
        if not txt:
            return default
        try:
            return int(float(txt))
        except Exception:
            raise ValueError(f"{label} must be numeric.")


# ============================================================
# Theme / style setup
# ============================================================

# def setup_ttk_styles(root: tk.Tk) -> None:
#     style = ttk.Style()
#     try:
#         style.theme_use("clam")
#     except Exception:
#         pass

#     root.configure(bg=AppConfig.COLOR_BG)

#     style.configure("TFrame", background=AppConfig.COLOR_BG)
#     style.configure(
#         "TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )
#     style.configure(
#         "Title.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_PRIMARY,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_TITLE, "bold"),
#     )
#     style.configure(
#         "Sub.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_MUTED,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )
#     style.configure(
#         "Hero.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_PRIMARY,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_HERO, "bold"),
#     )

#     style.configure("Card.TLabelframe", background=AppConfig.COLOR_CARD)
#     style.configure(
#         "Card.TLabelframe.Label",
#         background=AppConfig.COLOR_CARD,
#         foreground=AppConfig.COLOR_PRIMARY,
#         font=(AppConfig.FONT_FAMILY, 11, "bold"),
#     )

#     style.configure(
#         "TButton",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=8,
#     )

#     style.configure(
#         "Treeview",
#         background="#FFFFFF",
#         fieldbackground="#FFFFFF",
#         rowheight=28,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )
#     style.configure(
#         "Treeview.Heading",
#         background="#EAF2FA",
#         foreground=AppConfig.COLOR_PRIMARY,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#     )

#     style.configure("TNotebook", background=AppConfig.COLOR_BG)
#     style.configure(
#         "TNotebook.Tab",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=(14, 8),
#     )

# def setup_ttk_styles(root: tk.Tk) -> None:
#     style = ttk.Style()
#     try:
#         style.theme_use("clam")
#     except Exception:
#         pass

#     root.configure(bg=AppConfig.COLOR_BG)

#     # Base
#     style.configure("TFrame", background=AppConfig.COLOR_BG)
#     style.configure(
#         "TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )

#     # Titles
#     style.configure(
#         "Title.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_SECONDARY,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_TITLE, "bold"),
#     )
#     style.configure(
#         "Sub.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_MUTED,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )
#     style.configure(
#         "Hero.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_SECONDARY,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_HERO, "bold"),
#     )

#     # Card frames
#     style.configure(
#         "Card.TLabelframe",
#         background=AppConfig.COLOR_CARD,
#         bordercolor=AppConfig.COLOR_BORDER,
#         relief="solid",
#         borderwidth=1,
#     )
#     style.configure(
#         "Card.TLabelframe.Label",
#         background=AppConfig.COLOR_CARD,
#         foreground=AppConfig.COLOR_SECONDARY,
#         font=(AppConfig.FONT_FAMILY, 11, "bold"),
#     )

#     # Buttons
#     style.configure(
#         "TButton",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=8,
#         background=AppConfig.COLOR_PRIMARY,
#         foreground="#FFFFFF",
#         borderwidth=1,
#         focusthickness=0,
#     )
#     style.map(
#         "TButton",
#         background=[
#             ("active", AppConfig.COLOR_SECONDARY),
#             ("pressed", AppConfig.COLOR_SECONDARY),
#         ],
#         foreground=[
#             ("active", "#FFFFFF"),
#             ("pressed", "#FFFFFF"),
#         ]
#     )

#     # Navigation / metric buttons
#     style.configure(
#         "Metric.TButton",
#         font=(AppConfig.FONT_FAMILY, 11, "bold"),
#         padding=18,
#         background="#FFFFFF",
#         foreground=AppConfig.COLOR_SECONDARY,
#         borderwidth=1,
#         focusthickness=0,
#         relief="solid",
#     )
#     style.map(
#         "Metric.TButton",
#         background=[
#             ("active", "#F8FBFE"),
#             ("pressed", "#EEF5FB"),
#         ],
#         foreground=[
#             ("active", AppConfig.COLOR_PRIMARY),
#             ("pressed", AppConfig.COLOR_PRIMARY),
#         ]
#     )

#     # Treeview
#     style.configure(
#         "Treeview",
#         background="#FFFFFF",
#         fieldbackground="#FFFFFF",
#         rowheight=28,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#         bordercolor=AppConfig.COLOR_BORDER,
#     )
#     style.configure(
#         "Treeview.Heading",
#         background=AppConfig.COLOR_SECONDARY,
#         foreground="#FFFFFF",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#     )

#     # Notebook
#     style.configure("TNotebook", background=AppConfig.COLOR_BG, borderwidth=0)
#     style.configure(
#         "TNotebook.Tab",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=(14, 8),
#         background="#EAF2F9",
#         foreground=AppConfig.COLOR_SECONDARY,
#     )
#     style.map(
#         "TNotebook.Tab",
#         background=[("selected", "#FFFFFF")],
#         foreground=[("selected", AppConfig.COLOR_PRIMARY)],
#     )

#     # Combobox / Entry
#     style.configure("TEntry", fieldbackground="#FFFFFF")
#     style.configure("TCombobox", fieldbackground="#FFFFFF")

# def setup_ttk_styles(root: tk.Tk) -> None:
#     style = ttk.Style()
#     try:
#         style.theme_use("clam")
#     except Exception:
#         pass

#     root.configure(bg=AppConfig.COLOR_BG)

#     # --------------------------------------------------------
#     # Base
#     # --------------------------------------------------------
#     style.configure("TFrame", background=AppConfig.COLOR_BG)
#     style.configure(
#         "TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )

#     # --------------------------------------------------------
#     # Titles
#     # --------------------------------------------------------
#     style.configure(
#         "Title.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_TITLE, "bold"),
#     )
#     style.configure(
#         "Sub.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_MUTED,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#     )
#     style.configure(
#         "Hero.TLabel",
#         background=AppConfig.COLOR_BG,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_HERO, "bold"),
#     )

#     # --------------------------------------------------------
#     # Card frames
#     # --------------------------------------------------------
#     style.configure(
#         "Card.TLabelframe",
#         background=AppConfig.COLOR_CARD,
#         bordercolor=AppConfig.COLOR_BORDER,
#         borderwidth=1,
#         relief="solid",
#     )
#     style.configure(
#         "Card.TLabelframe.Label",
#         background=AppConfig.COLOR_CARD,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, 11, "bold"),
#     )

#     # --------------------------------------------------------
#     # Buttons
#     # --------------------------------------------------------
#     style.configure(
#         "TButton",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=8,
#         background=AppConfig.COLOR_PRIMARY,
#         foreground="#FFFFFF",
#         borderwidth=0,
#         focusthickness=0,
#         relief="flat",
#     )
#     style.map(
#         "TButton",
#         background=[
#             ("active", "#1F8BF0"),
#             ("pressed", "#1F8BF0"),
#         ],
#         foreground=[
#             ("active", "#FFFFFF"),
#             ("pressed", "#FFFFFF"),
#         ]
#     )

#     # Dashboard metric buttons/cards
#     style.configure(
#         "Metric.TButton",
#         font=(AppConfig.FONT_FAMILY, 11, "bold"),
#         padding=20,
#         background="#FFFFFF",
#         foreground=AppConfig.COLOR_TEXT,
#         borderwidth=1,
#         focusthickness=0,
#         relief="solid",
#     )
#     style.map(
#         "Metric.TButton",
#         background=[
#             ("active", "#F7FAFD"),
#             ("pressed", "#EEF6FC"),
#         ],
#         foreground=[
#             ("active", AppConfig.COLOR_TEXT),
#             ("pressed", AppConfig.COLOR_TEXT),
#         ]
#     )

#     # --------------------------------------------------------
#     # Treeview
#     # --------------------------------------------------------
#     style.configure(
#         "Treeview",
#         background="#FFFFFF",
#         fieldbackground="#FFFFFF",
#         foreground=AppConfig.COLOR_TEXT,
#         rowheight=30,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
#         bordercolor=AppConfig.COLOR_GRID,
#         lightcolor=AppConfig.COLOR_GRID,
#         darkcolor=AppConfig.COLOR_GRID,
#     )
#     style.configure(
#         "Treeview.Heading",
#         background=AppConfig.COLOR_ACCENT,
#         foreground=AppConfig.COLOR_TEXT,
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         relief="flat",
#     )
#     style.map(
#         "Treeview",
#         background=[("selected", "#DFF1FD")],
#         foreground=[("selected", AppConfig.COLOR_TEXT)]
#     )
#     style.map(
#         "Treeview.Heading",
#         background=[("active", "#BFC3C9")]
#     )

#     # --------------------------------------------------------
#     # Notebook
#     # --------------------------------------------------------
#     style.configure("TNotebook", background=AppConfig.COLOR_BG, borderwidth=0)
#     style.configure(
#         "TNotebook.Tab",
#         font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
#         padding=(14, 8),
#         background="#EAECEF",
#         foreground=AppConfig.COLOR_TEXT,
#     )
#     style.map(
#         "TNotebook.Tab",
#         background=[("selected", "#FFFFFF")],
#         foreground=[("selected", AppConfig.COLOR_TEXT)],
#     )

#     # --------------------------------------------------------
#     # Inputs
#     # --------------------------------------------------------
#     style.configure(
#         "TEntry",
#         fieldbackground="#FFFFFF",
#         foreground=AppConfig.COLOR_TEXT,
#         bordercolor=AppConfig.COLOR_GRID,
#     )
#     style.configure(
#         "TCombobox",
#         fieldbackground="#FFFFFF",
#         foreground=AppConfig.COLOR_TEXT,
#         bordercolor=AppConfig.COLOR_GRID,
#     )

def setup_ttk_styles(root: tk.Tk) -> None:
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    root.configure(bg=AppConfig.COLOR_BG)

    # Base
    style.configure("TFrame", background=AppConfig.COLOR_BG)
    style.configure(
        "TLabel",
        background=AppConfig.COLOR_BG,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
    )

    # Titles
    style.configure(
        "Title.TLabel",
        background=AppConfig.COLOR_BG,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_TITLE, "bold"),
    )
    style.configure(
        "Sub.TLabel",
        background=AppConfig.COLOR_BG,
        foreground=AppConfig.COLOR_MUTED,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
    )
    style.configure(
        "Hero.TLabel",
        background=AppConfig.COLOR_BG,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_HERO, "bold"),
    )

    # Card frames
    style.configure(
        "Card.TLabelframe",
        background=AppConfig.COLOR_CARD,
        bordercolor=AppConfig.COLOR_BORDER,
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=AppConfig.COLOR_CARD,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, 11, "bold"),
    )
    style.configure(
        "HomeCard.TLabelframe",
        background=AppConfig.COLOR_CARD,
        bordercolor=AppConfig.COLOR_GRID,
        borderwidth=2,
        relief="solid",
    )
    style.configure(
        "HomeCard.TLabelframe.Label",
        background=AppConfig.COLOR_CARD,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, 11, "bold"),
    )

    # Standard buttons
    style.configure(
        "TButton",
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
        padding=8,
        background=AppConfig.COLOR_PRIMARY,
        foreground=AppConfig.COLOR_TEXT,
        borderwidth=1,
        focusthickness=0,
        relief="solid",
    )
    style.map(
        "TButton",
        background=[
            ("active", AppConfig.COLOR_HOVER),
            ("pressed", AppConfig.COLOR_SECONDARY),
        ],
        foreground=[
            ("active", AppConfig.COLOR_TEXT),
            ("pressed", AppConfig.COLOR_TEXT),
        ]
    )

    # Flat dashboard metric buttons
    style.configure(
        "Metric.TButton",
        font=(AppConfig.FONT_FAMILY, 12, "bold"),
        padding=8,
        background=AppConfig.COLOR_BG,
        foreground=AppConfig.COLOR_TEXT,
        borderwidth=0,
        focusthickness=0,
        relief="flat",
    )
    style.map(
        "Metric.TButton",
        background=[
            ("active", AppConfig.COLOR_BG),
            ("pressed", AppConfig.COLOR_BG),
        ],
        foreground=[
            ("active", AppConfig.COLOR_TEXT),
            ("pressed", AppConfig.COLOR_TEXT),
        ]
    )
    style.configure(
        "HomeMetric.TButton",
        font=(AppConfig.FONT_FAMILY, 12, "bold"),
        padding=14,
        background=AppConfig.COLOR_CARD,
        foreground=AppConfig.COLOR_TEXT,
        borderwidth=2,
        bordercolor=AppConfig.COLOR_GRID,
        focusthickness=0,
        relief="solid",
        anchor="center",
    )
    style.map(
        "HomeMetric.TButton",
        background=[
            ("active", AppConfig.COLOR_HOVER),
            ("pressed", AppConfig.COLOR_SECONDARY),
        ],
        foreground=[
            ("active", AppConfig.COLOR_TEXT),
            ("pressed", AppConfig.COLOR_TEXT),
        ]
    )

    # Treeview
    style.configure(
        "Treeview",
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground=AppConfig.COLOR_TEXT,
        rowheight=30,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL),
        bordercolor=AppConfig.COLOR_GRID,
        lightcolor=AppConfig.COLOR_GRID,
        darkcolor=AppConfig.COLOR_GRID,
    )
    style.configure(
        "Treeview.Heading",
        background=AppConfig.COLOR_ACCENT,
        foreground=AppConfig.COLOR_TEXT,
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
        relief="flat",
    )
    style.map(
        "Treeview",
        background=[("selected", "#EEF1F4")],
        foreground=[("selected", AppConfig.COLOR_TEXT)]
    )

    # Notebook
    style.configure("TNotebook", background=AppConfig.COLOR_BG, borderwidth=0)
    style.configure(
        "TNotebook.Tab",
        font=(AppConfig.FONT_FAMILY, AppConfig.FONT_SIZE_NORMAL, "bold"),
        padding=(14, 8),
        background="#F1F3F5",
        foreground=AppConfig.COLOR_TEXT,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", "#FFFFFF")],
        foreground=[("selected", AppConfig.COLOR_TEXT)],
    )

    # Inputs
    style.configure(
        "TEntry",
        fieldbackground="#FFFFFF",
        foreground=AppConfig.COLOR_TEXT,
        bordercolor=AppConfig.COLOR_GRID,
    )
    style.configure(
        "TCombobox",
        fieldbackground="#FFFFFF",
        foreground=AppConfig.COLOR_TEXT,
        bordercolor=AppConfig.COLOR_GRID,
    )
# ============================================================
# Reusable dashboard card
# ============================================================

# class DashboardMetricCard(ttk.Frame):
#     def __init__(self, master, title: str, value: str = "0", subtitle: str = "", **kwargs):
#         super().__init__(master, **kwargs)
#         self.configure(style="TFrame")

#         outer = ttk.LabelFrame(self, text=title, style="Card.TLabelframe", padding=12)
#         outer.pack(fill="both", expand=True)

#         self.value_var = tk.StringVar(value=value)
#         self.subtitle_var = tk.StringVar(value=subtitle)

#         ttk.Label(
#             outer,
#             textvariable=self.value_var,
#             style="Hero.TLabel"
#         ).pack(anchor="w")

#         ttk.Label(
#             outer,
#             textvariable=self.subtitle_var,
#             style="Sub.TLabel"
#         ).pack(anchor="w", pady=(4, 0))

#     def set_value(self, value: Any) -> None:
#         self.value_var.set(str(value))

#     def set_subtitle(self, subtitle: str) -> None:
#         self.subtitle_var.set(subtitle)

class DashboardMetricCard(ttk.Frame):
    def __init__(self, master, title: str, value: str = "0", subtitle: str = "", on_click=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(style="TFrame")
        self.on_click = on_click

        self.outer = ttk.LabelFrame(self, text=title, style="Card.TLabelframe", padding=12)
        self.outer.pack(fill="both", expand=True)

        self.value_var = tk.StringVar(value=value)
        self.subtitle_var = tk.StringVar(value=subtitle)

        self.value_label = ttk.Label(
            self.outer,
            textvariable=self.value_var,
            style="Hero.TLabel"
        )
        self.value_label.pack(anchor="w")

        self.subtitle_label = ttk.Label(
            self.outer,
            textvariable=self.subtitle_var,
            style="Sub.TLabel"
        )
        self.subtitle_label.pack(anchor="w", pady=(4, 0))

        if self.on_click:
            self._bind_clicks()

    def _bind_clicks(self):
        widgets = [self, self.outer, self.value_label, self.subtitle_label]
        for w in widgets:
            w.bind("<Button-1>", self._handle_click)
            w.bind("<Enter>", lambda e: self.configure(cursor="hand2"))
            w.bind("<Leave>", lambda e: self.configure(cursor=""))

    def _handle_click(self, event=None):
        if callable(self.on_click):
            self.on_click()

    def set_value(self, value: Any) -> None:
        self.value_var.set(str(value))

    def set_subtitle(self, subtitle: str) -> None:
        self.subtitle_var.set(subtitle)


# ============================================================
# Navigation button strip
# ============================================================

class NavStrip(ttk.Frame):
    def __init__(self, master, buttons: Sequence[Tuple[str, Callable]], **kwargs):
        super().__init__(master, **kwargs)
        for text, command in buttons:
            ttk.Button(self, text=text, command=command).pack(side="left", padx=4)


# ============================================================
# Table insertion helpers
# ============================================================

def insert_tree_row(tree: ttk.Treeview, values: Sequence[Any], tags: Sequence[str] = ()) -> str:
    return tree.insert("", "end", values=tuple(values), tags=tuple(tags))


def insert_task_hierarchy_row(
    tree: ttk.Treeview,
    parent: str,
    text: str,
    hours: Any = "",
    department: str = "",
    status: str = "",
    notes: str = "",
    tags: Sequence[str] = (),
) -> str:
    return tree.insert(
        parent,
        "end",
        text=text,
        values=(hours, department, status, notes),
        tags=tuple(tags),
    )


# ============================================================
# Combo refresh helpers
# ============================================================

def refresh_combo_from_records(combo: ttk.Combobox, records: Sequence[Any], attr_name: str, keep_current: bool = True) -> None:
    values = []
    for record in records:
        value = getattr(record, attr_name, "")
        if value:
            values.append(str(value))
    set_combobox_values(combo, values, keep_current=keep_current)


# ============================================================
# Parent task display helpers
# ============================================================

def build_parent_task_map(tasks: Sequence[Any]) -> Dict[str, str]:
    """
    Returns mapping {display_name: task_id} for top-level tasks only.
    Expects each item to have .task_name, .task_id, .parent_task_id
    """
    output: Dict[str, str] = {}
    for t in tasks:
        if not norm_text(getattr(t, "parent_task_id", "")):
            output[norm_text(getattr(t, "task_name", ""))] = norm_text(getattr(t, "task_id", ""))
    return output


# ============================================================
# Safe execution wrapper
# ============================================================

def guarded_action(parent, func: Callable, success_message: Optional[str] = None, on_success: Optional[Callable] = None):
    try:
        result = func()
        if success_message:
            show_info("Success", success_message)
        if on_success:
            on_success()
        return result
    except Exception as exc:
        show_error("Error", str(exc))
        return None
