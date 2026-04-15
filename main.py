from __future__ import annotations

from pathlib import Path

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

import tkinter as tk
from tkinter import ttk, messagebox

from app_config import AppConfig
from storage import WorkbookManager
from services import ERPServiceHub
from reports import ReportService
from mailer import MailerService
from updater import AppUpdater

from ui_common import compact_widget_spacing, setup_ttk_styles
from ui_home import HomePage
from ui_modules import ModulePage
from ui_products import ProductPage
from ui_projects import ProjectPage
from ui_scheduler import SchedulerPage
from ui_dependencies import DependenciesPage
from ui_jobcards_board import JobCardsBoardPage
from ui_completed_jobs import CompletedJobsPage
from ui_parts import PartsPage


class ERPDesktopApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self._save_job = None
        self._online_refresh_job = None
        self._online_refresh_in_progress = False

        AppConfig.ensure_directories()
        self.workbook_manager = WorkbookManager()
        self.repo = self.workbook_manager.repo

        if self.workbook_manager.uses_online():
            backend_suffix = "Online"
        else:
            backend_suffix = "Local"
        self.root.title(f"{AppConfig.APP_TITLE} v{AppConfig.APP_VERSION} [{backend_suffix}]")
        self.root.geometry(f"{AppConfig.WINDOW_WIDTH}x{AppConfig.WINDOW_HEIGHT}")
        self.root.minsize(AppConfig.MIN_WIDTH, AppConfig.MIN_HEIGHT)

        setup_ttk_styles(self.root)

        # ----------------------------------------------------
        # Shared app state
        # ----------------------------------------------------
        self.selected_module_code = ""
        self.selected_product_code = ""
        self.selected_project_code = ""

        # ----------------------------------------------------
        # Core services
        # ----------------------------------------------------
        self.services = ERPServiceHub(self.repo)
        self.mailer = MailerService()
        self.reports = ReportService(self.services, self.mailer)
        self.updater = AppUpdater(self)

        # ----------------------------------------------------
        # Layout shell
        # ----------------------------------------------------
        initial_status = (
            f"Connected to {self.workbook_manager.workbook_path}"
            if not self.workbook_manager.uses_local()
            else "Ready. Create or open a workbook to begin."
        )
        self.status_var = tk.StringVar(value=initial_status)
        self.pages = {}
        self.page_classes = {}

        self._build_menu()
        self._build_shell()
        self._build_pages()
        compact_widget_spacing(self.page_container)
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)

        self.show_page("home")
        self._schedule_startup_update_check()
        self._schedule_online_refresh(initial=True)

    # ========================================================
    # Shell
    # ========================================================

    def _build_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Check for Updates", command=lambda: self.check_for_updates(manual=True))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Check for Updates", command=lambda: self.check_for_updates(manual=True))
        help_menu.add_command(
            label="About",
            command=lambda: messagebox.showinfo(
                "About",
                f"{AppConfig.APP_TITLE}\nVersion {AppConfig.APP_VERSION}",
            ),
        )
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)
        self.menubar = menubar

    def _build_shell(self):
        self.topbar = ttk.Frame(self.root, padding=(8, 5), style="TFrame")
        self.topbar.pack(fill="x")

        left = ttk.Frame(self.topbar)
        left.pack(side="left", fill="x", expand=True)

        self.top_logo_label = ttk.Label(left)
        self.top_logo_label.pack(side="left", padx=(0, 6))
        self._load_top_logo()

        ttk.Label(
            left,
            text="Liquimech Project Management Suite",
            style="Title.TLabel"
        ).pack(side="left")

        self.status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            style="Sub.TLabel",
            anchor="w"
        )
        self.status_bar.pack(fill="x", padx=8, pady=(0, 3))

        self.page_container = ttk.Frame(self.root)
        self.page_container.pack(fill="both", expand=True)

    def _schedule_startup_update_check(self):
        if not AppConfig.ENABLE_STARTUP_UPDATE_CHECK:
            return
        self.root.after(2000, lambda: self.check_for_updates(manual=False))

    def _schedule_online_refresh(self, initial: bool = False):
        if not self.workbook_manager.uses_online():
            return
        if not AppConfig.ENABLE_ONLINE_AUTO_REFRESH:
            return
        if self._online_refresh_job:
            self.root.after_cancel(self._online_refresh_job)
            self._online_refresh_job = None
        delay_ms = 2000 if initial else AppConfig.ONLINE_REFRESH_INTERVAL_MS
        self._online_refresh_job = self.root.after(delay_ms, self._run_online_refresh)

    def _run_online_refresh(self):
        self._online_refresh_job = None
        if not self.workbook_manager.uses_online():
            return
        if self._online_refresh_in_progress:
            self._schedule_online_refresh()
            return

        self._online_refresh_in_progress = True
        try:
            self.repo.reload_cache()
            self.refresh_visible_pages()
        except Exception as exc:
            self.set_status(f"Online refresh warning: {exc}")
        finally:
            self._online_refresh_in_progress = False
            self._schedule_online_refresh()

    def _build_pages(self):
        page_classes = [
            ("home", HomePage),
            ("modules", ModulePage),
            ("products", ProductPage),
            ("projects", ProjectPage),
            ("jobcards", JobCardsBoardPage),
            ("scheduler", SchedulerPage),
            ("dependencies", DependenciesPage),
            ("completed_jobs", CompletedJobsPage),
            ("parts", PartsPage),
        ]

        self.page_classes = dict(page_classes)

    def _ensure_page(self, page_name: str):
        page = self.pages.get(page_name)
        if page is not None:
            return page

        page_cls = self.page_classes.get(page_name)
        if page_cls is None:
            raise ValueError(f"Unknown page: {page_name}")

        page = page_cls(self.page_container, self)
        page.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages[page_name] = page
        return page

    def _load_top_logo(self):
        if Image is None or ImageTk is None:
            self.top_logo_label.config(text="")
            return

        possible_paths = [
            Path("logo.png"),
            Path("assets/logo.png"),
            Path("assets/liquimech_logo.png"),
        ]

        logo_path = None
        for p in possible_paths:
            if p.exists():
                logo_path = p
                break

        if not logo_path:
            self.top_logo_label.config(text="")
            return

        try:
            img = Image.open(logo_path)
            img = img.resize((42, 42))
            self.top_logo_photo = ImageTk.PhotoImage(img)
            self.top_logo_label.config(image=self.top_logo_photo, background=AppConfig.COLOR_BG)
        except Exception:
            self.top_logo_label.config(text="")

    # ========================================================
    # Navigation
    # ========================================================

    def show_page(self, page_name: str):
        if page_name not in self.page_classes:
            raise ValueError(f"Unknown page: {page_name}")

        page = self._ensure_page(page_name)
        page.tkraise()

        try:
            if page_name == "home":
                page.refresh_page()
            else:
                if self.workbook_manager.has_workbook():
                    page.refresh_page()
        except Exception as exc:
            self.set_status(f"Page refresh warning: {exc}")

    # ========================================================
    # Shared helpers
    # ========================================================

    def refresh_visible_pages(self):
        for name in ["home", "parts", "modules", "products", "projects", "jobcards", "completed_jobs", "scheduler", "dependencies"]:
            if name in self.pages and self.pages[name].winfo_ismapped():
                try:
                    self.pages[name].refresh_page()
                except Exception:
                    pass

    def set_status(self, text: str):
        self.status_var.set(text)

    def check_for_updates(self, manual: bool = True):
        self.updater.check_for_updates(manual=manual)

    def require_workbook(self) -> bool:
        if not self.workbook_manager.has_workbook():
            return False
        return True

    def refresh_home_page(self):
        if "home" in self.pages:
            try:
                self.pages["home"].refresh_page()
            except Exception as exc:
                self.set_status(f"Home refresh warning: {exc}")

    def refresh_all_pages(self):
        for name, page in self.pages.items():
            try:
                page.refresh_page()
            except Exception:
                pass

    # ========================================================
    # Cross-page shared selections
    # ========================================================

    def set_selected_module(self, module_code: str):
        self.selected_module_code = module_code or ""
        self.set_status(
            f"Selected module: {self.selected_module_code}"
            if self.selected_module_code else
            "Module selection cleared."
        )

    def set_selected_product(self, product_code: str):
        self.selected_product_code = product_code or ""
        self.set_status(
            f"Selected product: {self.selected_product_code}"
            if self.selected_product_code else
            "Product selection cleared."
        )

    def set_selected_project(self, project_code: str):
        self.selected_project_code = project_code or ""
        self.set_status(
            f"Selected project: {self.selected_project_code}"
            if self.selected_project_code else
            "Project selection cleared."
        )

    # ========================================================
    # Convenience hooks for later extensions
    # ========================================================

    def open_module_page(self, module_code: str):
        self.set_selected_module(module_code)
        self.show_page("modules")

    def open_product_page(self, product_code: str):
        self.set_selected_product(product_code)
        self.show_page("products")

    def open_project_page(self, project_code: str):
        self.set_selected_project(project_code)
        self.show_page("projects")

    def open_jobcards_page(self):
        self.show_page("jobcards")

    def schedule_save(self, delay_ms=800):
        if hasattr(self, "_save_job") and self._save_job:
            self.root.after_cancel(self._save_job)

        self._save_job = self.root.after(delay_ms, self.flush_save)

    def flush_save(self):
        try:
            if getattr(self.repo, "_dirty", False):
                self.repo.save_workbook()
        except Exception as exc:
            self.set_status(f"Save warning: {exc}")

    def _handle_close(self):
        if self._save_job:
            self.root.after_cancel(self._save_job)
            self._save_job = None
        if self._online_refresh_job:
            self.root.after_cancel(self._online_refresh_job)
            self._online_refresh_job = None
        self.root.destroy()


# ============================================================
# Main launcher
# ============================================================

def main():
    root = tk.Tk()

    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = ERPDesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
