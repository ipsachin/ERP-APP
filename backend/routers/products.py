from __future__ import annotations

from fastapi import APIRouter, HTTPException

from app_config import AppConfig
from backend.repository import get_record_by_key, list_records


router = APIRouter(prefix="/products", tags=["products"])


@router.get("")
def list_products() -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PRODUCTS)


@router.get("/{product_code}")
def get_product(product_code: str) -> dict[str, object]:
    record = get_record_by_key(AppConfig.SHEET_PRODUCTS, "ProductCode", product_code)
    if record is None:
        raise HTTPException(status_code=404, detail="Product not found")
    return record


@router.get("/{product_code}/modules")
def get_product_modules(product_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PRODUCT_MODULES, {"ProductCode": product_code})


@router.get("/{product_code}/documents")
def get_product_documents(product_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_PRODUCT_DOCUMENTS, {"ProductCode": product_code})


@router.get("/{product_code}/workorders")
def get_product_workorders(product_code: str) -> list[dict[str, object]]:
    return list_records(AppConfig.SHEET_WORKORDERS, {"OwnerType": "PRODUCT", "OwnerCode": product_code})
