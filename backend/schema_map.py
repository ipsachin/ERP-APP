from __future__ import annotations

from app_config import AppConfig


SHEET_TABLE_NAMES = {
    AppConfig.SHEET_PARTS: "parts",
    AppConfig.SHEET_MODULES: "modules",
    AppConfig.SHEET_TASKS: "tasks",
    AppConfig.SHEET_COMPONENTS: "components",
    AppConfig.SHEET_DOCUMENTS: "documents",
    AppConfig.SHEET_PRODUCTS: "products",
    AppConfig.SHEET_PRODUCT_MODULES: "product_modules",
    AppConfig.SHEET_PRODUCT_DOCUMENTS: "product_documents",
    AppConfig.SHEET_PROJECTS: "projects",
    AppConfig.SHEET_PROJECT_MODULES: "project_modules",
    AppConfig.SHEET_PROJECT_TASKS: "project_tasks",
    AppConfig.SHEET_PROJECT_DOCUMENTS: "project_documents",
    AppConfig.SHEET_WORKORDERS: "workorders",
    AppConfig.SHEET_COMPLETED_JOBS: "completed_jobs",
    AppConfig.SHEET_COMPLETED_JOB_LINES: "completed_job_lines",
}

HEADER_COLUMN_OVERRIDES = {
    "WorkOrderID": "workorder_id",
    "WorkOrderName": "workorder_name",
}

DOC_SHEETS = {
    AppConfig.SHEET_DOCUMENTS: "DocID",
    AppConfig.SHEET_PRODUCT_DOCUMENTS: "ProdDocID",
    AppConfig.SHEET_PROJECT_DOCUMENTS: "ProjectDocID",
}


def header_to_column_name(header: str) -> str:
    if header in HEADER_COLUMN_OVERRIDES:
        return HEADER_COLUMN_OVERRIDES[header]
    chars: list[str] = []
    for idx, ch in enumerate(header):
        if ch.isupper() and idx > 0 and (header[idx - 1].islower() or (idx + 1 < len(header) and header[idx + 1].islower())):
            chars.append("_")
        chars.append(ch.lower())
    return "".join(chars)


def get_table_name(sheet_name: str) -> str:
    table_name = SHEET_TABLE_NAMES.get(sheet_name)
    if not table_name:
        raise ValueError(f"Unknown sheet name: {sheet_name}")
    return table_name


def get_headers(sheet_name: str) -> list[str]:
    headers = AppConfig.SHEET_HEADERS.get(sheet_name)
    if not headers:
        raise ValueError(f"Unknown sheet name: {sheet_name}")
    return headers


def get_columns(sheet_name: str) -> list[str]:
    return [header_to_column_name(header) for header in get_headers(sheet_name)]
