# ============================================================
# app_config.py
# Central configuration for Liquimech ERP Desktop App
# ============================================================

import sys
from pathlib import Path


def get_bundle_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_version_file_candidates() -> list[Path]:
    app_dir = get_app_dir()
    bundle_dir = get_bundle_dir()
    return [
        app_dir / "VERSION",
        bundle_dir / "VERSION",
        Path(__file__).resolve().parent / "VERSION",
    ]


def load_app_version(default: str = "0.0.0") -> str:
    for path in get_version_file_candidates():
        try:
            value = path.read_text(encoding="utf-8").strip()
        except FileNotFoundError:
            continue
        if value:
            return value
    return default


class AppConfig:
    # --------------------------------------------------------
    # Application
    # --------------------------------------------------------
    APP_TITLE = "Liquimech Project Management Suite"
    APP_VERSION = load_app_version("1.0.0")
    WINDOW_WIDTH = 1680
    WINDOW_HEIGHT = 940
    MIN_WIDTH = 1280
    MIN_HEIGHT = 760
    GITHUB_RELEASE_OWNER = "ipsachin"
    GITHUB_RELEASE_REPO = "ERP-APP"
    GITHUB_RELEASE_ASSET_NAME = "LiquimechERP-Setup.exe"
    ENABLE_STARTUP_UPDATE_CHECK = True
    GITHUB_RELEASE_CHECK_INTERVAL_SECONDS = 86400
    ENABLE_ONLINE_AUTO_REFRESH = False
    ONLINE_REFRESH_INTERVAL_MS = 5000

    # --------------------------------------------------------
    # Theme / UI
    # --------------------------------------------------------

    COLOR_BG = "#FAFAFA"           # overall app background
    COLOR_CARD = "#FFFFFF"         # white panels
    COLOR_TEXT = "#2F3A45"         # main text
    COLOR_MUTED = "#7A8591"        # softer text
    COLOR_PRIMARY = "#E9ECEF"      # light grey button
    COLOR_SECONDARY = "#F4F6F8"    # soft grey fill
    COLOR_ACCENT = "#D9DDE2"       # table header grey
    COLOR_ACCENT_SOFT = "#F7F8FA"  # light tint
    COLOR_GRID = "#E6E9ED"         # borders
    COLOR_BORDER = "#E3E6EA"
    COLOR_HOVER = "#F3F5F7"
    COLOR_SUCCESS = "#22C55E"
    COLOR_DANGER = "#EF4444"

    FONT_FAMILY = "Segoe UI"
    FONT_SIZE_NORMAL = 9
    FONT_SIZE_SMALL = 8
    FONT_SIZE_TITLE = 15
    FONT_SIZE_HERO = 20

    # --------------------------------------------------------
    # Working folders
    # --------------------------------------------------------
    ROOT_DIR = get_app_dir()
    BUNDLE_DIR = get_bundle_dir()
    ASSETS_DIR = BUNDLE_DIR / "assets"
    WORKSPACE_DIR = ROOT_DIR / "workspace"
    DATA_DIR = WORKSPACE_DIR / "data"
    EXPORTS_DIR = WORKSPACE_DIR / "exports"
    TEMP_DIR = WORKSPACE_DIR / "temp"
    MODULE_DOCS_DIR = WORKSPACE_DIR / "module_docs"
    PRODUCT_DOCS_DIR = WORKSPACE_DIR / "product_docs"
    PROJECT_DOCS_DIR = WORKSPACE_DIR / "project_docs"

    # --------------------------------------------------------
    # Excel Sheet Names
    # --------------------------------------------------------
    SHEET_PARTS = "Parts"
    SHEET_MODULES = "Modules"
    SHEET_TASKS = "Tasks"
    SHEET_COMPONENTS = "Components"
    SHEET_DOCUMENTS = "Documents"

    SHEET_PRODUCTS = "Products"
    SHEET_PRODUCT_MODULES = "ProductModules"
    SHEET_PRODUCT_DOCUMENTS = "ProductDocuments"

    SHEET_PROJECTS = "Projects"
    SHEET_PROJECT_MODULES = "ProjectModules"
    SHEET_PROJECT_TASKS = "ProjectTasks"
    SHEET_PROJECT_DOCUMENTS = "ProjectDocuments"

    SHEET_WORKORDERS = "WorkOrders"

    SHEET_COMPLETED_JOBS = "CompletedJobs"
    SHEET_COMPLETED_JOB_LINES = "CompletedJobLines"

    ALL_SHEETS = [
        SHEET_PARTS,
        SHEET_MODULES,
        SHEET_TASKS,
        SHEET_COMPONENTS,
        SHEET_DOCUMENTS,
        SHEET_PRODUCTS,
        SHEET_PRODUCT_MODULES,
        SHEET_PRODUCT_DOCUMENTS,
        SHEET_PROJECTS,
        SHEET_PROJECT_MODULES,
        SHEET_PROJECT_TASKS,
        SHEET_PROJECT_DOCUMENTS,
        SHEET_WORKORDERS,
        SHEET_COMPLETED_JOBS,
        SHEET_COMPLETED_JOB_LINES,
    ]

    # --------------------------------------------------------
    # Header definitions
    # --------------------------------------------------------
    PART_HEADERS = [
        "PartID",
        "PartNumber",
        "PartName",
        "Description",
        "UnitPrice",
        "LeadTimeDays",
        "PreferredSupplier",
        "StockOnHand",
        "Manufacturer",
        "Category",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    MODULE_HEADERS = [
        "ModuleCode",
        "QuoteRef",
        "ModuleName",
        "Description",
        "InstructionText",
        "EstimatedHours",
        "StockOnHand",
        "Status",
        "CreatedOn",
        "UpdatedOn",
    ]

    TASK_HEADERS = [
        "TaskID",
        "OwnerType",
        "OwnerCode",
        "TaskName",
        "Department",
        "EstimatedHours",
        "ParentTaskID",
        "DependencyTaskID",
        "Stage",
        "Status",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    COMPONENT_HEADERS = [
        "ComponentID",
        "OwnerType",
        "OwnerCode",
        "ComponentName",
        "Qty",
        "SOHQty",
        "PreferredSupplier",
        "LeadTimeDays",
        "UnitPrice",
        "PartNumber",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    DOCUMENT_HEADERS = [
        "DocID",
        "OwnerType",
        "OwnerCode",
        "SectionName",
        "DocName",
        "DocType",
        "FilePath",
        "InstructionText",
        "AddedOn",
        "UpdatedOn",
    ]

    PRODUCT_HEADERS = [
        "ProductCode",
        "QuoteRef",
        "ProductName",
        "Description",
        "Revision",
        "Status",
        "CreatedOn",
        "UpdatedOn",
    ]

    PRODUCT_MODULE_HEADERS = [
        "LinkID",
        "ProductCode",
        "ModuleCode",
        "ModuleOrder",
        "ModuleQty",
        "DependencyModuleCode",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    PRODUCT_DOCUMENT_HEADERS = [
        "ProdDocID",
        "ProductCode",
        "SectionName",
        "DocName",
        "DocType",
        "FilePath",
        "InstructionText",
        "AddedOn",
        "UpdatedOn",
    ]

    PROJECT_HEADERS = [
        "ProjectCode",
        "QuoteRef",
        "ProjectName",
        "ClientName",
        "Location",
        "Description",
        "LinkedProductCode",
        "Status",
        "StartDate",
        "DueDate",
        "CreatedOn",
        "UpdatedOn",
    ]

    PROJECT_MODULE_HEADERS = [
        "LinkID",
        "ProjectCode",
        "ModuleCode",
        "SourceType",
        "SourceCode",
        "ModuleOrder",
        "ModuleQty",
        "Stage",
        "Status",
        "DependencyModuleCode",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    PROJECT_TASK_HEADERS = [
        "ProjectTaskID",
        "ProjectCode",
        "ModuleCode",
        "SourceTaskID",
        "ParentProjectTaskID",
        "TaskName",
        "Department",
        "EstimatedHours",
        "Stage",
        "Status",
        "DependencyTaskID",
        "AssignedTo",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    PROJECT_DOCUMENT_HEADERS = [
        "ProjectDocID",
        "ProjectCode",
        "SectionName",
        "DocName",
        "DocType",
        "FilePath",
        "InstructionText",
        "AddedOn",
        "UpdatedOn",
    ]

    WORKORDER_HEADERS = [
        "WorkOrderID",
        "OwnerType",
        "OwnerCode",
        "WorkOrderName",
        "Stage",
        "Owner",
        "DueDate",
        "Status",
        "Notes",
        "CreatedOn",
        "UpdatedOn",
    ]

    COMPLETED_JOB_HEADERS = [
    "SnapshotID",
    "ProjectCode",
    "QuoteRef",
    "ProductCode",
    "ProductName",
    "ClientName",
    "CompletedOn",
    "LabourHours",
    "PartsTotal",
    "GrandTotal",
    "Notes",
    ]

    COMPLETED_JOB_LINE_HEADERS = [
        "SnapshotLineID",
        "SnapshotID",
        "LineType",
        "Code",
        "Description",
        "PartNumber",
        "Qty",
        "Hours",
        "UnitPrice",
        "LineTotal",
        "Source",
    ]

    SHEET_HEADERS = {
        SHEET_PARTS: PART_HEADERS,
        SHEET_MODULES: MODULE_HEADERS,
        SHEET_TASKS: TASK_HEADERS,
        SHEET_COMPONENTS: COMPONENT_HEADERS,
        SHEET_DOCUMENTS: DOCUMENT_HEADERS,
        SHEET_PRODUCTS: PRODUCT_HEADERS,
        SHEET_PRODUCT_MODULES: PRODUCT_MODULE_HEADERS,
        SHEET_PRODUCT_DOCUMENTS: PRODUCT_DOCUMENT_HEADERS,
        SHEET_PROJECTS: PROJECT_HEADERS,
        SHEET_PROJECT_MODULES: PROJECT_MODULE_HEADERS,
        SHEET_PROJECT_TASKS: PROJECT_TASK_HEADERS,
        SHEET_PROJECT_DOCUMENTS: PROJECT_DOCUMENT_HEADERS,
        SHEET_WORKORDERS: WORKORDER_HEADERS,
        SHEET_COMPLETED_JOBS: COMPLETED_JOB_HEADERS,
        SHEET_COMPLETED_JOB_LINES: COMPLETED_JOB_LINE_HEADERS,
    }

    # --------------------------------------------------------
    # Dropdown values
    # --------------------------------------------------------
    DEPARTMENTS = [
        "Electrical",
        "Mechanical",
        "Automation",
        "Software",
        "Procurement",
        "Fabrication",
        "Assembly",
        "QA/QC",
        "Operations",
        "Commissioning",
    ]

    MODULE_STATUSES = [
        "Draft",
        "Released",
        "In Build",
        "Archived",
    ]

    PRODUCT_STATUSES = [
        "Draft",
        "Quoted",
        "Released",
        "Archived",
    ]

    PROJECT_STATUSES = [
        "Planned",
        "Quoted",
        "Awarded",
        "In Progress",
        "On Hold",
        "Completed",
        "Archived",
    ]

    TASK_STATUSES = [
        "Not Started",
        "Started",
        "Blocked",
        "Fabrication",
        "Assembly",
        "Testing",
        "Completed",
    ]

    MODULE_EXEC_STAGES = [
        "Not Started",
        "Engineering",
        "Procurement",
        "Fabrication",
        "Assembly",
        "Testing",
        "QA/QC",
        "Dispatch",
        "Commissioning",
        "Completed",
    ]

    WORKORDER_STAGES = [
        "Engineering",
        "Procurement",
        "Fabrication",
        "Assembly",
        "Testing",
        "QA/QC",
        "Dispatch",
        "Commissioning",
    ]

    WORKORDER_STATUSES = [
        "Open",
        "In Progress",
        "Blocked",
        "Completed",
        "Cancelled",
    ]

    DOC_TYPES = [
        "PDF Spec Sheet",
        "Manual",
        "Photo",
        "Drawing",
        "3D Model",
        "Procedure",
        "Checklist",
        "Other",
    ]

    OWNER_TYPES = [
        "MODULE",
        "PRODUCT",
        "PROJECT",
    ]

    PROJECT_MODULE_SOURCE_TYPES = [
        "DIRECT",
        "FROM_PRODUCT",
    ]

    # --------------------------------------------------------
    # Default workbook file name
    # --------------------------------------------------------
    DEFAULT_WORKBOOK_NAME = "Liquimech_ERP_Master.xlsx"

    # --------------------------------------------------------
    # PDF / export defaults
    # --------------------------------------------------------
    PDF_MARGIN_MM = 12
    PDF_TITLE = "Liquimech ERP Report"

    # --------------------------------------------------------
    # Utility
    # --------------------------------------------------------
    @classmethod
    def ensure_directories(cls) -> None:
        cls.ASSETS_DIR.mkdir(exist_ok=True)
        cls.WORKSPACE_DIR.mkdir(exist_ok=True)
        cls.DATA_DIR.mkdir(exist_ok=True)
        cls.EXPORTS_DIR.mkdir(exist_ok=True)
        cls.TEMP_DIR.mkdir(exist_ok=True)
        cls.MODULE_DOCS_DIR.mkdir(exist_ok=True)
        cls.PRODUCT_DOCS_DIR.mkdir(exist_ok=True)
        cls.PROJECT_DOCS_DIR.mkdir(exist_ok=True)
