from __future__ import annotations

from typing import Any

from fastapi import APIRouter, HTTPException
from fastapi import Body
from fastapi.responses import Response

from backend.repository import delete_attachment, get_attachment, save_attachment
from backend.schema_map import DOC_SHEETS


router = APIRouter(prefix="/attachments", tags=["attachments"])


@router.get("/{sheet_name}/{record_id}")
def download_attachment(sheet_name: str, record_id: str) -> Response:
    canonical_sheet = next((name for name in DOC_SHEETS if name.lower() == sheet_name.lower()), None)
    if canonical_sheet is None:
        raise HTTPException(status_code=404, detail="Attachment sheet not found")

    attachment = get_attachment(canonical_sheet, record_id)
    if attachment is None:
        raise HTTPException(status_code=404, detail="Attachment not found")

    headers = {
        "Content-Disposition": f"attachment; filename={attachment['file_name']}",
    }
    return Response(
        content=attachment["file_data"],
        media_type=attachment.get("content_type") or "application/octet-stream",
        headers=headers,
    )


@router.put("/{sheet_name}/{record_id}")
def upload_attachment(sheet_name: str, record_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, str]:
    canonical_sheet = next((name for name in DOC_SHEETS if name.lower() == sheet_name.lower()), None)
    if canonical_sheet is None:
        raise HTTPException(status_code=404, detail="Attachment sheet not found")

    uri = save_attachment(
        canonical_sheet,
        record_id,
        payload["file_name"],
        payload.get("original_path", ""),
        payload.get("content_type", "application/octet-stream"),
        payload["content_base64"],
        payload.get("created_on"),
        payload.get("updated_on"),
    )
    return {"file_path": uri}


@router.delete("/{sheet_name}/{record_id}")
def remove_attachment(sheet_name: str, record_id: str) -> dict[str, str]:
    canonical_sheet = next((name for name in DOC_SHEETS if name.lower() == sheet_name.lower()), None)
    if canonical_sheet is None:
        raise HTTPException(status_code=404, detail="Attachment sheet not found")
    delete_attachment(canonical_sheet, record_id)
    return {"status": "ok"}
