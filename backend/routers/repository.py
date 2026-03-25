from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Body, HTTPException, Query, Request

from backend.repository import (
    append_record,
    clear_records,
    delete_record_by_key,
    get_record_by_key,
    list_records,
    rewrite_records,
    update_record_by_key,
    upsert_record,
)
from backend.schema_map import get_headers


router = APIRouter(prefix="/repository", tags=["repository"])


@router.get("/{sheet_name}/records")
def api_list_records(sheet_name: str, request: Request) -> list[dict[str, Any]]:
    filters = {key: value for key, value in request.query_params.items()}
    return list_records(sheet_name, filters)


@router.get("/{sheet_name}/record")
def api_get_record(sheet_name: str, key_name: str = Query(...), key_value: str = Query(...)) -> dict[str, Any]:
    record = get_record_by_key(sheet_name, key_name, key_value)
    if record is None:
        raise HTTPException(status_code=404, detail="Record not found")
    return record


@router.post("/{sheet_name}/append")
def api_append_record(sheet_name: str, row: dict[str, Any] = Body(...)) -> dict[str, str]:
    append_record(sheet_name, row)
    return {"status": "ok"}


@router.post("/{sheet_name}/upsert")
def api_upsert_record(sheet_name: str, key_name: str = Query(...), row: dict[str, Any] = Body(...)) -> dict[str, str]:
    status = upsert_record(sheet_name, key_name, row)
    return {"status": status}


@router.patch("/{sheet_name}/record")
def api_update_record(
    sheet_name: str,
    key_name: str = Query(...),
    key_value: str = Query(...),
    updates: dict[str, Any] = Body(...),
) -> dict[str, bool]:
    updated = update_record_by_key(sheet_name, key_name, key_value, updates)
    return {"updated": updated}


@router.delete("/{sheet_name}/record")
def api_delete_record(sheet_name: str, key_name: str = Query(...), key_value: str = Query(...)) -> dict[str, bool]:
    deleted = delete_record_by_key(sheet_name, key_name, key_value)
    return {"deleted": deleted}


@router.post("/{sheet_name}/clear")
def api_clear_records(sheet_name: str) -> dict[str, str]:
    clear_records(sheet_name)
    return {"status": "ok"}


@router.post("/{sheet_name}/rewrite")
def api_rewrite_records(sheet_name: str, rows: list[dict[str, Any]] = Body(...)) -> dict[str, str]:
    rewrite_records(sheet_name, rows)
    return {"status": "ok"}


@router.get("/{sheet_name}/headers")
def api_get_headers(sheet_name: str) -> dict[str, list[str]]:
    return {"headers": get_headers(sheet_name)}
