from __future__ import annotations

from fastapi import APIRouter, HTTPException

from app_config import AppConfig
from backend.repository import get_record_by_key, list_records


router = APIRouter(prefix="/modules", tags=["modules"])


@router.get("")
def list_modules() -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_MODULES)


@router.get("/{module_code}")
def get_module(module_code: str) -> dict[str, object]:
    record = get_record_by_key(AppConfig.SHEET_MODULES, "ModuleCode", module_code)
    if record is None:
        raise HTTPException(status_code=404, detail="Module not found")
    return record


@router.get("/{module_code}/tasks")
def get_module_tasks(module_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_TASKS, {"OwnerType": "MODULE", "OwnerCode": module_code})


@router.get("/{module_code}/components")
def get_module_components(module_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_COMPONENTS, {"OwnerType": "MODULE", "OwnerCode": module_code})


@router.get("/{module_code}/documents")
def get_module_documents(module_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_DOCUMENTS, {"OwnerType": "MODULE", "OwnerCode": module_code})
