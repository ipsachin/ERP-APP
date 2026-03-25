from __future__ import annotations

from fastapi import APIRouter, HTTPException

from app_config import AppConfig
from backend.repository import get_record_by_key, list_records


router = APIRouter(prefix="/projects", tags=["projects"])


@router.get("")
def list_projects() -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PROJECTS)


@router.get("/{project_code}")
def get_project(project_code: str) -> dict[str, object]:
    record = get_record_by_key(AppConfig.SHEET_PROJECTS, "ProjectCode", project_code)
    if record is None:
        raise HTTPException(status_code=404, detail="Project not found")
    return record


@router.get("/{project_code}/modules")
def get_project_modules(project_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PROJECT_MODULES, {"ProjectCode": project_code})


@router.get("/{project_code}/tasks")
def get_project_tasks(project_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PROJECT_TASKS, {"ProjectCode": project_code})


@router.get("/{project_code}/documents")
def get_project_documents(project_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PROJECT_DOCUMENTS, {"ProjectCode": project_code})


@router.get("/{project_code}/completed-jobs")
def get_project_completed_jobs(project_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_COMPLETED_JOBS, {"ProjectCode": project_code})
