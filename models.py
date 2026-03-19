# ============================================================
# models.py
# Data models for Liquimech ERP Desktop App
# ============================================================

from dataclasses import dataclass, asdict, fields
from typing import Any, Dict, List, Optional


# ============================================================
# Base helpers
# ============================================================

class ModelMixin:
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def field_names(cls) -> List[str]:
        return [f.name for f in fields(cls)]

    @classmethod
    def from_dict(cls, data: Dict[str, Any]):
        valid = {f.name for f in fields(cls)}
        filtered = {k: v for k, v in data.items() if k in valid}
        return cls(**filtered)



# ============================================================
# Part master
# ============================================================

@dataclass
class PartRecord(ModelMixin):
    part_id: str = ""
    part_number: str = ""
    part_name: str = ""
    description: str = ""
    unit_price: float = 0.0
    lead_time_days: int = 0
    preferred_supplier: str = ""
    stock_on_hand: float = 0.0
    manufacturer: str = ""
    category: str = ""
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Module
# ============================================================

@dataclass
class ModuleRecord(ModelMixin):
    module_code: str = ""
    quote_ref: str = ""
    module_name: str = ""
    description: str = ""
    instruction_text: str = ""
    estimated_hours: float = 0.0
    stock_on_hand: float = 0.0
    status: str = "Draft"
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Task
# ============================================================

@dataclass
class TaskRecord(ModelMixin):
    task_id: str = ""
    owner_type: str = "MODULE"     # MODULE / PRODUCT / PROJECT
    owner_code: str = ""
    task_name: str = ""
    department: str = ""
    estimated_hours: float = 0.0
    parent_task_id: str = ""
    dependency_task_id: str = ""
    stage: str = ""
    status: str = "Not Started"
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Component / BOM
# ============================================================

@dataclass
class ComponentRecord(ModelMixin):
    component_id: str = ""
    owner_type: str = "MODULE"     # MODULE / PRODUCT / PROJECT
    owner_code: str = ""
    component_name: str = ""
    qty: float = 0.0
    soh_qty: float = 0.0
    preferred_supplier: str = ""
    lead_time_days: int = 0
    unit_price: float = 0.0
    part_number: str = ""
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Documents
# ============================================================

@dataclass
class DocumentRecord(ModelMixin):
    doc_id: str = ""
    owner_type: str = "MODULE"     # MODULE / PRODUCT / PROJECT
    owner_code: str = ""
    section_name: str = ""
    doc_name: str = ""
    doc_type: str = "Other"
    file_path: str = ""
    instruction_text: str = ""
    added_on: str = ""
    updated_on: str = ""


# ============================================================
# Product
# ============================================================

@dataclass
class ProductRecord(ModelMixin):
    product_code: str = ""
    quote_ref: str = ""
    product_name: str = ""
    description: str = ""
    revision: str = "R0"
    status: str = "Draft"
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Product <-> Module link
# ============================================================

@dataclass
class ProductModuleLinkRecord(ModelMixin):
    link_id: str = ""
    product_code: str = ""
    module_code: str = ""
    module_order: int = 0
    module_qty: int = 1
    dependency_module_code: str = ""
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Product documents
# ============================================================

@dataclass
class ProductDocumentRecord(ModelMixin):
    prod_doc_id: str = ""
    product_code: str = ""
    section_name: str = ""
    doc_name: str = ""
    doc_type: str = "Other"
    file_path: str = ""
    instruction_text: str = ""
    added_on: str = ""
    updated_on: str = ""


# ============================================================
# Live Project
# ============================================================

@dataclass
class ProjectRecord(ModelMixin):
    project_code: str = ""
    quote_ref: str = ""
    project_name: str = ""
    client_name: str = ""
    location: str = ""
    description: str = ""
    linked_product_code: str = ""
    status: str = "Planned"
    start_date: str = ""
    due_date: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Project <-> Module link
# ============================================================

@dataclass
class ProjectModuleLinkRecord(ModelMixin):
    link_id: str = ""
    project_code: str = ""
    module_code: str = ""
    source_type: str = "DIRECT"        # DIRECT / FROM_PRODUCT
    source_code: str = ""
    module_order: int = 0
    module_qty: int = 1
    stage: str = "Not Started"
    status: str = "Not Started"
    dependency_module_code: str = ""
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Project task execution record
# ============================================================

@dataclass
class ProjectTaskRecord(ModelMixin):
    project_task_id: str = ""
    project_code: str = ""
    module_code: str = ""
    source_task_id: str = ""
    parent_project_task_id: str = ""
    task_name: str = ""
    department: str = ""
    estimated_hours: float = 0.0
    stage: str = ""
    status: str = "Not Started"
    dependency_task_id: str = ""
    assigned_to: str = ""
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Project documents
# ============================================================

@dataclass
class ProjectDocumentRecord(ModelMixin):
    project_doc_id: str = ""
    project_code: str = ""
    section_name: str = ""
    doc_name: str = ""
    doc_type: str = "Other"
    file_path: str = ""
    instruction_text: str = ""
    added_on: str = ""
    updated_on: str = ""


# ============================================================
# Work orders
# ============================================================

@dataclass
class WorkOrderRecord(ModelMixin):
    workorder_id: str = ""
    owner_type: str = "PRODUCT"    # PRODUCT / PROJECT
    owner_code: str = ""
    workorder_name: str = ""
    stage: str = ""
    owner: str = ""
    due_date: str = ""
    status: str = "Open"
    notes: str = ""
    created_on: str = ""
    updated_on: str = ""


# ============================================================
# Optional aggregate bundle models
# These are not direct Excel rows. They are convenience containers
# for UI/report generation.
# ============================================================

@dataclass
class ModuleBundle(ModelMixin):
    module: Optional[ModuleRecord] = None
    tasks: Optional[List[TaskRecord]] = None
    components: Optional[List[ComponentRecord]] = None
    documents: Optional[List[DocumentRecord]] = None


@dataclass
class ProductBundle(ModelMixin):
    product: Optional[ProductRecord] = None
    module_links: Optional[List[ProductModuleLinkRecord]] = None
    product_documents: Optional[List[ProductDocumentRecord]] = None
    workorders: Optional[List[WorkOrderRecord]] = None
    modules: Optional[List[ModuleRecord]] = None
    tasks_by_module: Optional[Dict[str, List[TaskRecord]]] = None
    total_hours: float = 0.0


@dataclass
class ProjectBundle(ModelMixin):
    project: Optional[ProjectRecord] = None
    module_links: Optional[List[ProjectModuleLinkRecord]] = None
    project_tasks: Optional[List[ProjectTaskRecord]] = None
    project_documents: Optional[List[ProjectDocumentRecord]] = None
    workorders: Optional[List[WorkOrderRecord]] = None
    modules: Optional[List[ModuleRecord]] = None
    total_hours: float = 0.0