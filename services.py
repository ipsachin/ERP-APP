# ============================================================
# services.py
# Business logic layer for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

import re
from types import SimpleNamespace



from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from app_config import AppConfig
from models import (
    ModuleRecord,
    TaskRecord,
    ComponentRecord,
    DocumentRecord,
    ProductRecord,
    ProductModuleLinkRecord,
    ProductDocumentRecord,
    ProjectRecord,
    ProjectModuleLinkRecord,
    ProjectTaskRecord,
    ProjectDocumentRecord,
    WorkOrderRecord,
    ModuleBundle,
    ProductBundle,
    ProjectBundle,
)
from storage import ExcelRepository, WorkbookSchema, now_str, norm_text, to_float, to_int

from datetime import timedelta
import math

# ============================================================
# Shared helpers
# ============================================================

def safe_name(text: str) -> str:
    return str(text or "").strip().replace("/", "_").replace("\\", "_")


def safe_filename(text: str, replacement: str = "_") -> str:
    text = str(text or "").strip()
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        text = text.replace(ch, replacement)
    text = "".join(c for c in text if ord(c) >= 32)
    while "__" in text:
        text = text.replace("__", "_")
    text = text.strip(" ._")
    return text or "output"


# ============================================================
# ID / Code factory
# ============================================================

class CodeFactory:
    @staticmethod
    def module_code(quote_ref: str, module_name: str) -> str:
        clean_q = norm_text(quote_ref).upper().replace(" ", "_")
        clean_m = norm_text(module_name).upper().replace(" ", "_")
        return f"{clean_q}/{clean_m}" if clean_q else clean_m

    @staticmethod
    def product_code(quote_ref: str, product_name: str) -> str:
        clean_q = norm_text(quote_ref).upper().replace(" ", "_")
        clean_p = norm_text(product_name).upper().replace(" ", "_")
        return f"P_{clean_q}_{clean_p}" if clean_q else f"P_{clean_p}"

    @staticmethod
    def project_code(quote_ref: str, project_name: str) -> str:
        clean_q = norm_text(quote_ref).upper().replace(" ", "_")
        clean_p = norm_text(project_name).upper().replace(" ", "_")
        return f"PRJ_{clean_q}_{clean_p}" if clean_q else f"PRJ_{clean_p}"

    @staticmethod
    def _stamp() -> str:
        return datetime.now().strftime("%H%M%S%f")

    @classmethod
    def task_id(cls, owner_code: str, task_name: str) -> str:
        return f"T::{safe_name(owner_code)}::{norm_text(task_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def component_id(cls, owner_code: str, component_name: str) -> str:
        return f"C::{safe_name(owner_code)}::{norm_text(component_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def document_id(cls, owner_code: str, doc_name: str) -> str:
        return f"D::{safe_name(owner_code)}::{norm_text(doc_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def product_doc_id(cls, product_code: str, doc_name: str) -> str:
        return f"PD::{safe_name(product_code)}::{norm_text(doc_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def project_doc_id(cls, project_code: str, doc_name: str) -> str:
        return f"PRJD::{safe_name(project_code)}::{norm_text(doc_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def link_id(cls, prefix: str, owner_code: str, ref_code: str) -> str:
        return f"{prefix}::{safe_name(owner_code)}::{safe_name(ref_code)}::{cls._stamp()}"

    @classmethod
    def product_module_link_id(cls, product_code: str, module_code: str) -> str:
        return cls.link_id("PM", product_code, module_code)

    @classmethod
    def project_module_link_id(cls, project_code: str, module_code: str) -> str:
        return cls.link_id("PRJM", project_code, module_code)

    @classmethod
    def workorder_id(cls, owner_code: str, workorder_name: str) -> str:
        return f"WO::{safe_name(owner_code)}::{norm_text(workorder_name).upper().replace(' ', '_')}::{cls._stamp()}"

    @classmethod
    def project_task_id(cls, project_code: str, module_code: str, task_name: str) -> str:
        return f"PT::{safe_name(project_code)}::{safe_name(module_code)}::{norm_text(task_name).upper().replace(' ', '_')}::{cls._stamp()}"


# ============================================================
# Base service
# ============================================================

class BaseService:
    def __init__(self, repo: ExcelRepository):
        self.repo = repo

    def _require_workbook(self) -> None:
        self.repo.ensure_ready()


# ============================================================
# Module service
# ============================================================

class ModuleService(BaseService):
    # --------------------------------------------------------
    # Module CRUD
    # --------------------------------------------------------
    def create_or_update_module(
        self,
        quote_ref: str,
        module_name: str,
        description: str = "",
        instruction_text: str = "",
        estimated_hours: float = 0.0,
        stock_on_hand: float = 0.0,
        status: str = "Draft",
        existing_module_code: Optional[str] = None,
    ) -> str:
        self._require_workbook()

        if not norm_text(module_name):
            raise ValueError("Module name is required.")

        module_code = CodeFactory.module_code(quote_ref, module_name)
        ts = now_str()

        record = ModuleRecord(
            module_code=module_code,
            quote_ref=norm_text(quote_ref),
            module_name=norm_text(module_name),
            description=norm_text(description),
            instruction_text=norm_text(instruction_text),
            estimated_hours=to_float(estimated_hours),
            stock_on_hand=to_float(stock_on_hand),
            status=norm_text(status) or "Draft",
            created_on=ts,
            updated_on=ts,
        )

        old_row = None
        if existing_module_code:
            old_row = self.repo.find_row(AppConfig.SHEET_MODULES, 0, existing_module_code)
        else:
            old_row = self.repo.find_row(AppConfig.SHEET_MODULES, 0, module_code)

        if old_row:
            record.created_on = norm_text(old_row[8]) or ts
            self.repo.upsert_dict(AppConfig.SHEET_MODULES, "ModuleCode", {
                "ModuleCode": record.module_code,
                "QuoteRef": record.quote_ref,
                "ModuleName": record.module_name,
                "Description": record.description,
                "InstructionText": record.instruction_text,
                "EstimatedHours": record.estimated_hours,
                "StockOnHand": record.stock_on_hand,
                "Status": record.status,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })
        else:
            self.repo.append_dict(AppConfig.SHEET_MODULES, {
                "ModuleCode": record.module_code,
                "QuoteRef": record.quote_ref,
                "ModuleName": record.module_name,
                "Description": record.description,
                "InstructionText": record.instruction_text,
                "EstimatedHours": record.estimated_hours,
                "StockOnHand": record.stock_on_hand,
                "Status": record.status,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })

        if existing_module_code and norm_text(existing_module_code) != module_code:
            self._cascade_module_code_change(existing_module_code, module_code)

        return module_code

    def _cascade_module_code_change(self, old_module_code: str, new_module_code: str) -> None:
        # Tasks owner code
        task_rows = self.repo.filter_dicts(
            AppConfig.SHEET_TASKS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(old_module_code)
        )
        for row in task_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_TASKS,
                "TaskID",
                row["TaskID"],
                {"OwnerCode": new_module_code, "UpdatedOn": now_str()}
            )

        # Components owner code
        comp_rows = self.repo.filter_dicts(
            AppConfig.SHEET_COMPONENTS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(old_module_code)
        )
        for row in comp_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_COMPONENTS,
                "ComponentID",
                row["ComponentID"],
                {"OwnerCode": new_module_code, "UpdatedOn": now_str()}
            )

        # Documents owner code
        doc_rows = self.repo.filter_dicts(
            AppConfig.SHEET_DOCUMENTS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(old_module_code)
        )
        for row in doc_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_DOCUMENTS,
                "DocID",
                row["DocID"],
                {"OwnerCode": new_module_code, "UpdatedOn": now_str()}
            )

        # Product links
        prod_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(old_module_code)
        )
        for row in prod_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PRODUCT_MODULES,
                "LinkID",
                row["LinkID"],
                {"ModuleCode": new_module_code, "UpdatedOn": now_str()}
            )

        # Project module links
        prj_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(old_module_code)
        )
        for row in prj_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                row["LinkID"],
                {"ModuleCode": new_module_code, "UpdatedOn": now_str()}
            )

        # Project execution tasks module code
        prj_task_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(old_module_code)
        )
        for row in prj_task_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                row["ProjectTaskID"],
                {"ModuleCode": new_module_code, "UpdatedOn": now_str()}
            )

    def delete_module(self, module_code: str, delete_docs_files: bool = False) -> None:
        self._require_workbook()

        docs = self.get_module_documents(module_code)

        self.repo.delete_row_by_key_name(AppConfig.SHEET_MODULES, "ModuleCode", module_code)
        self.repo.delete_rows_by_owner(AppConfig.SHEET_TASKS, "MODULE", module_code)
        self.repo.delete_rows_by_owner(AppConfig.SHEET_COMPONENTS, "MODULE", module_code)
        self.repo.delete_rows_by_owner(AppConfig.SHEET_DOCUMENTS, "MODULE", module_code)

        self.repo.delete_rows_where(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )

        if delete_docs_files:
            for d in docs:
                fp = Path(norm_text(d.file_path))
                if fp.exists():
                    try:
                        fp.unlink()
                    except Exception:
                        pass

    def get_module(self, module_code: str) -> Optional[ModuleRecord]:
        row = self.repo.find_row(AppConfig.SHEET_MODULES, 0, module_code)
        if not row:
            return None
        return ModuleRecord(
            module_code=norm_text(row[0]),
            quote_ref=norm_text(row[1]),
            module_name=norm_text(row[2]),
            description=norm_text(row[3]),
            instruction_text=norm_text(row[4]),
            estimated_hours=to_float(row[5]),
            stock_on_hand=to_float(row[6]),
            status=norm_text(row[7]),
            created_on=norm_text(row[8]),
            updated_on=norm_text(row[9]),
        )

    def list_modules(self, search_text: str = "") -> List[ModuleRecord]:
        rows = self.repo.list_rows(AppConfig.SHEET_MODULES)
        output: List[ModuleRecord] = []
        q = norm_text(search_text).lower()

        for row in rows:
            if not norm_text(row[0]):
                continue
            combined = " ".join(str(x or "") for x in row).lower()
            if q and q not in combined:
                continue

            output.append(ModuleRecord(
                module_code=norm_text(row[0]),
                quote_ref=norm_text(row[1]),
                module_name=norm_text(row[2]),
                description=norm_text(row[3]),
                instruction_text=norm_text(row[4]),
                estimated_hours=to_float(row[5]),
                stock_on_hand=to_float(row[6]),
                status=norm_text(row[7]),
                created_on=norm_text(row[8]),
                updated_on=norm_text(row[9]),
            ))

        output.sort(key=lambda x: x.updated_on, reverse=True)
        return output

    # --------------------------------------------------------
    # Module tasks
    # --------------------------------------------------------
    def add_module_task(
        self,
        module_code: str,
        task_name: str,
        department: str = "",
        estimated_hours: float = 0.0,
        parent_task_id: str = "",
        dependency_task_id: str = "",
        stage: str = "",
        status: str = "Not Started",
        notes: str = "",
    ) -> str:
        self._require_workbook()
        if not norm_text(module_code):
            raise ValueError("Module code is required.")
        if not norm_text(task_name):
            raise ValueError("Task name is required.")

        task_id = CodeFactory.task_id(module_code, task_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_TASKS, {
            "TaskID": task_id,
            "OwnerType": "MODULE",
            "OwnerCode": module_code,
            "TaskName": norm_text(task_name),
            "Department": norm_text(department),
            "EstimatedHours": to_float(estimated_hours),
            "ParentTaskID": norm_text(parent_task_id),
            "DependencyTaskID": norm_text(dependency_task_id),
            "Stage": norm_text(stage),
            "Status": norm_text(status) or "Not Started",
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        return task_id

    def update_task(self, task_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_TASKS, "TaskID", task_id, updates)

    def delete_task(self, task_id: str) -> bool:
        # Remove parent references from child tasks
        child_rows = self.repo.filter_dicts(
            AppConfig.SHEET_TASKS,
            lambda r: norm_text(r.get("ParentTaskID")) == norm_text(task_id)
        )
        for row in child_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_TASKS,
                "TaskID",
                row["TaskID"],
                {"ParentTaskID": "", "UpdatedOn": now_str()}
            )

        # Remove dependency references
        dep_rows = self.repo.filter_dicts(
            AppConfig.SHEET_TASKS,
            lambda r: norm_text(r.get("DependencyTaskID")) == norm_text(task_id)
        )
        for row in dep_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_TASKS,
                "TaskID",
                row["TaskID"],
                {"DependencyTaskID": "", "UpdatedOn": now_str()}
            )

        # Project task refs
        prj_dep_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("DependencyTaskID")) == norm_text(task_id)
        )
        for row in prj_dep_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                row["ProjectTaskID"],
                {"DependencyTaskID": "", "UpdatedOn": now_str()}
            )

        return self.repo.delete_row_by_key_name(AppConfig.SHEET_TASKS, "TaskID", task_id)

    def get_module_tasks(self, module_code: str) -> List[TaskRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_TASKS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(module_code)
        )
        output: List[TaskRecord] = []
        for r in rows:
            output.append(TaskRecord(
                task_id=norm_text(r["TaskID"]),
                owner_type=norm_text(r["OwnerType"]),
                owner_code=norm_text(r["OwnerCode"]),
                task_name=norm_text(r["TaskName"]),
                department=norm_text(r["Department"]),
                estimated_hours=to_float(r["EstimatedHours"]),
                parent_task_id=norm_text(r["ParentTaskID"]),
                dependency_task_id=norm_text(r["DependencyTaskID"]),
                stage=norm_text(r["Stage"]),
                status=norm_text(r["Status"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    # --------------------------------------------------------
    # Module components
    # --------------------------------------------------------
    def add_module_component(
        self,
        module_code: str,
        component_name: str,
        qty: float = 0.0,
        soh_qty: float = 0.0,
        preferred_supplier: str = "",
        lead_time_days: int = 0,
        part_number: str = "",
        notes: str = "",
    ) -> str:
        self._require_workbook()
        if not norm_text(component_name):
            raise ValueError("Component name is required.")

        component_id = CodeFactory.component_id(module_code, component_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_COMPONENTS, {
            "ComponentID": component_id,
            "OwnerType": "MODULE",
            "OwnerCode": module_code,
            "ComponentName": norm_text(component_name),
            "Qty": to_float(qty),
            "SOHQty": to_float(soh_qty),
            "PreferredSupplier": norm_text(preferred_supplier),
            "LeadTimeDays": to_int(lead_time_days),
            "PartNumber": norm_text(part_number),
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        return component_id

    def update_component(self, component_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", component_id, updates)

    def delete_component(self, component_id: str) -> bool:
        return self.repo.delete_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", component_id)

    def get_module_components(self, module_code: str) -> List[ComponentRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_COMPONENTS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(module_code)
        )
        output: List[ComponentRecord] = []
        for r in rows:
            output.append(ComponentRecord(
                component_id=norm_text(r["ComponentID"]),
                owner_type=norm_text(r["OwnerType"]),
                owner_code=norm_text(r["OwnerCode"]),
                component_name=norm_text(r["ComponentName"]),
                qty=to_float(r["Qty"]),
                soh_qty=to_float(r["SOHQty"]),
                preferred_supplier=norm_text(r["PreferredSupplier"]),
                lead_time_days=to_int(r["LeadTimeDays"]),
                part_number=norm_text(r["PartNumber"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    # --------------------------------------------------------
    # Module documents
    # --------------------------------------------------------
    def add_module_document(
        self,
        module_code: str,
        source_file_path: str,
        section_name: str = "",
        doc_type: str = "Other",
        instruction_text: str = "",
        copy_file: bool = True,
    ) -> str:
        self._require_workbook()

        src = Path(source_file_path)
        if not src.exists():
            raise FileNotFoundError(f"Source document not found: {source_file_path}")

        target_path = src
        if copy_file:
            folder = self.repo.get_module_docs_folder(module_code)
            candidate = folder / src.name
            if candidate.exists():
                candidate = folder / f"{candidate.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{candidate.suffix}"
            import shutil
            shutil.copy2(src, candidate)
            target_path = candidate

        doc_name = target_path.name
        doc_id = CodeFactory.document_id(module_code, doc_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_DOCUMENTS, {
            "DocID": doc_id,
            "OwnerType": "MODULE",
            "OwnerCode": module_code,
            "SectionName": norm_text(section_name),
            "DocName": doc_name,
            "DocType": norm_text(doc_type) or "Other",
            "FilePath": str(target_path),
            "InstructionText": norm_text(instruction_text),
            "AddedOn": ts,
            "UpdatedOn": ts,
        })
        return doc_id

    def update_document(self, doc_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_DOCUMENTS, "DocID", doc_id, updates)

    def delete_document(self, doc_id: str, delete_file: bool = False) -> bool:
        row = self.repo.find_row(AppConfig.SHEET_DOCUMENTS, 0, doc_id)
        if row and delete_file:
            fp = Path(norm_text(row[6]))
            if fp.exists():
                try:
                    fp.unlink()
                except Exception:
                    pass
        return self.repo.delete_row_by_key_name(AppConfig.SHEET_DOCUMENTS, "DocID", doc_id)

    def get_module_documents(self, module_code: str) -> List[DocumentRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_DOCUMENTS,
            lambda r: norm_text(r.get("OwnerType")) == "MODULE" and norm_text(r.get("OwnerCode")) == norm_text(module_code)
        )
        output: List[DocumentRecord] = []
        for r in rows:
            output.append(DocumentRecord(
                doc_id=norm_text(r["DocID"]),
                owner_type=norm_text(r["OwnerType"]),
                owner_code=norm_text(r["OwnerCode"]),
                section_name=norm_text(r["SectionName"]),
                doc_name=norm_text(r["DocName"]),
                doc_type=norm_text(r["DocType"]),
                file_path=norm_text(r["FilePath"]),
                instruction_text=norm_text(r["InstructionText"]),
                added_on=norm_text(r["AddedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    # --------------------------------------------------------
    # Bundle
    # --------------------------------------------------------
    def get_module_bundle(self, module_code: str) -> ModuleBundle:
        return ModuleBundle(
            module=self.get_module(module_code),
            tasks=self.get_module_tasks(module_code),
            components=self.get_module_components(module_code),
            documents=self.get_module_documents(module_code),
        )


# ============================================================
# Product service
# ============================================================

class ProductService(BaseService):
    def __init__(self, repo: ExcelRepository):
        super().__init__(repo)
        self.module_service = ModuleService(repo)

    # --------------------------------------------------------
    # Product CRUD
    # --------------------------------------------------------
    def create_or_update_product(
        self,
        quote_ref: str,
        product_name: str,
        description: str = "",
        revision: str = "R0",
        status: str = "Draft",
        existing_product_code: Optional[str] = None,
    ) -> str:
        self._require_workbook()
        if not norm_text(product_name):
            raise ValueError("Product name is required.")

        product_code = CodeFactory.product_code(quote_ref, product_name)
        ts = now_str()

        record = ProductRecord(
            product_code=product_code,
            quote_ref=norm_text(quote_ref),
            product_name=norm_text(product_name),
            description=norm_text(description),
            revision=norm_text(revision) or "R0",
            status=norm_text(status) or "Draft",
            created_on=ts,
            updated_on=ts,
        )

        old_row = None
        if existing_product_code:
            old_row = self.repo.find_row(AppConfig.SHEET_PRODUCTS, 0, existing_product_code)
        else:
            old_row = self.repo.find_row(AppConfig.SHEET_PRODUCTS, 0, product_code)

        if old_row:
            record.created_on = norm_text(old_row[6]) or ts
            self.repo.upsert_dict(AppConfig.SHEET_PRODUCTS, "ProductCode", {
                "ProductCode": record.product_code,
                "QuoteRef": record.quote_ref,
                "ProductName": record.product_name,
                "Description": record.description,
                "Revision": record.revision,
                "Status": record.status,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })
        else:
            self.repo.append_dict(AppConfig.SHEET_PRODUCTS, {
                "ProductCode": record.product_code,
                "QuoteRef": record.quote_ref,
                "ProductName": record.product_name,
                "Description": record.description,
                "Revision": record.revision,
                "Status": record.status,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })

        if existing_product_code and norm_text(existing_product_code) != product_code:
            self._cascade_product_code_change(existing_product_code, product_code)

        return product_code

    def _cascade_product_code_change(self, old_product_code: str, new_product_code: str) -> None:
        # Product modules
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(old_product_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PRODUCT_MODULES,
                "LinkID",
                row["LinkID"],
                {"ProductCode": new_product_code, "UpdatedOn": now_str()}
            )

        # Product docs
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_DOCUMENTS,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(old_product_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PRODUCT_DOCUMENTS,
                "ProdDocID",
                row["ProdDocID"],
                {"ProductCode": new_product_code, "UpdatedOn": now_str()}
            )

        # Projects linked product
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECTS,
            lambda r: norm_text(r.get("LinkedProductCode")) == norm_text(old_product_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECTS,
                "ProjectCode",
                row["ProjectCode"],
                {"LinkedProductCode": new_product_code, "UpdatedOn": now_str()}
            )

        # Workorders
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PRODUCT" and norm_text(r.get("OwnerCode")) == norm_text(old_product_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_WORKORDERS,
                "WorkOrderID",
                row["WorkOrderID"],
                {"OwnerCode": new_product_code, "UpdatedOn": now_str()}
            )

    def delete_product(self, product_code: str, delete_docs_files: bool = False) -> None:
        docs = self.get_product_documents(product_code)

        self.repo.delete_row_by_key_name(AppConfig.SHEET_PRODUCTS, "ProductCode", product_code)
        self.repo.delete_rows_where(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_PRODUCT_DOCUMENTS,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PRODUCT" and norm_text(r.get("OwnerCode")) == norm_text(product_code)
        )

        if delete_docs_files:
            for d in docs:
                fp = Path(norm_text(d.file_path))
                if fp.exists():
                    try:
                        fp.unlink()
                    except Exception:
                        pass

    def get_product(self, product_code: str) -> Optional[ProductRecord]:
        row = self.repo.find_row(AppConfig.SHEET_PRODUCTS, 0, product_code)
        if not row:
            return None
        return ProductRecord(
            product_code=norm_text(row[0]),
            quote_ref=norm_text(row[1]),
            product_name=norm_text(row[2]),
            description=norm_text(row[3]),
            revision=norm_text(row[4]),
            status=norm_text(row[5]),
            created_on=norm_text(row[6]),
            updated_on=norm_text(row[7]),
        )

    def list_products(self, search_text: str = "") -> List[ProductRecord]:
        rows = self.repo.list_rows(AppConfig.SHEET_PRODUCTS)
        output: List[ProductRecord] = []
        q = norm_text(search_text).lower()

        for row in rows:
            if not norm_text(row[0]):
                continue
            combined = " ".join(str(x or "") for x in row).lower()
            if q and q not in combined:
                continue

            output.append(ProductRecord(
                product_code=norm_text(row[0]),
                quote_ref=norm_text(row[1]),
                product_name=norm_text(row[2]),
                description=norm_text(row[3]),
                revision=norm_text(row[4]),
                status=norm_text(row[5]),
                created_on=norm_text(row[6]),
                updated_on=norm_text(row[7]),
            ))

        output.sort(key=lambda x: x.updated_on, reverse=True)
        return output

    # --------------------------------------------------------
    # Product module linking
    # --------------------------------------------------------
    def add_module_to_product(
        self,
        product_code: str,
        module_code: str,
        qty: int = 1,
        dependency_module_code: str = "",
        notes: str = "",
    ) -> str:
        self._require_workbook()

        existing = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
            and norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        if existing:
            raise ValueError("This module is already assigned to the product.")

        current = self.get_product_module_links(product_code)
        order_num = len(current) + 1
        link_id = CodeFactory.product_module_link_id(product_code, module_code)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_PRODUCT_MODULES, {
            "LinkID": link_id,
            "ProductCode": product_code,
            "ModuleCode": module_code,
            "ModuleOrder": order_num,
            "ModuleQty": max(1, to_int(qty, 1)),
            "DependencyModuleCode": norm_text(dependency_module_code),
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        return link_id

    def remove_module_from_product(self, product_code: str, module_code: str) -> int:
        deleted = self.repo.delete_rows_where(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
            and norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        self._normalize_product_module_order(product_code)
        return deleted

    def update_product_module_link(self, link_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_PRODUCT_MODULES, "LinkID", link_id, updates)

    def set_module_qty(self, product_code: str, module_code: str, qty: int) -> bool:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
            and norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        if not rows:
            return False
        return self.repo.update_row_by_key_name(
            AppConfig.SHEET_PRODUCT_MODULES,
            "LinkID",
            rows[0]["LinkID"],
            {"ModuleQty": max(1, to_int(qty, 1)), "UpdatedOn": now_str()}
        )

    def set_module_dependency(self, product_code: str, module_code: str, dependency_module_code: str) -> bool:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
            and norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        if not rows:
            return False
        return self.repo.update_row_by_key_name(
            AppConfig.SHEET_PRODUCT_MODULES,
            "LinkID",
            rows[0]["LinkID"],
            {"DependencyModuleCode": norm_text(dependency_module_code), "UpdatedOn": now_str()}
        )

    def save_module_order(self, product_code: str, ordered_items: List[Tuple[str, int]]) -> None:
        """
        ordered_items = [(module_code, qty), ...]
        """
        with self.repo.batch_update():
            for idx, (module_code, qty) in enumerate(ordered_items, start=1):
                rows = self.repo.filter_dicts(
                    AppConfig.SHEET_PRODUCT_MODULES,
                    lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
                    and norm_text(r.get("ModuleCode")) == norm_text(module_code)
                )
                if not rows:
                    continue
                self.repo.update_row_by_key_name(
                    AppConfig.SHEET_PRODUCT_MODULES,
                    "LinkID",
                    rows[0]["LinkID"],
                    {
                        "ModuleOrder": idx,
                        "ModuleQty": max(1, to_int(qty, 1)),
                        "UpdatedOn": now_str(),
                    }
                )
            self._normalize_product_module_order(product_code)

    def _normalize_product_module_order(self, product_code: str) -> None:
        rows = self.get_product_module_links(product_code)
        rows.sort(key=lambda x: x.module_order)
        with self.repo.batch_update():
            for idx, row in enumerate(rows, start=1):
                self.repo.update_row_by_key_name(
                    AppConfig.SHEET_PRODUCT_MODULES,
                    "LinkID",
                    row.link_id,
                    {"ModuleOrder": idx, "UpdatedOn": now_str()}
                )

    def get_product_module_links(self, product_code: str) -> List[ProductModuleLinkRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_MODULES,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
        )
        output: List[ProductModuleLinkRecord] = []
        for r in rows:
            output.append(ProductModuleLinkRecord(
                link_id=norm_text(r["LinkID"]),
                product_code=norm_text(r["ProductCode"]),
                module_code=norm_text(r["ModuleCode"]),
                module_order=to_int(r["ModuleOrder"]),
                module_qty=max(1, to_int(r["ModuleQty"], 1)),
                dependency_module_code=norm_text(r["DependencyModuleCode"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        output.sort(key=lambda x: x.module_order)
        return output

    # --------------------------------------------------------
    # Product docs
    # --------------------------------------------------------
    def add_product_document(
        self,
        product_code: str,
        source_file_path: str,
        section_name: str = "",
        doc_type: str = "Other",
        instruction_text: str = "",
        copy_file: bool = True,
    ) -> str:
        self._require_workbook()

        src = Path(source_file_path)
        if not src.exists():
            raise FileNotFoundError(f"Source document not found: {source_file_path}")

        target_path = src
        if copy_file:
            folder = self.repo.get_product_docs_folder(product_code)
            candidate = folder / src.name
            if candidate.exists():
                candidate = folder / f"{candidate.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{candidate.suffix}"
            import shutil
            shutil.copy2(src, candidate)
            target_path = candidate

        doc_name = target_path.name
        doc_id = CodeFactory.product_doc_id(product_code, doc_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_PRODUCT_DOCUMENTS, {
            "ProdDocID": doc_id,
            "ProductCode": product_code,
            "SectionName": norm_text(section_name),
            "DocName": doc_name,
            "DocType": norm_text(doc_type) or "Other",
            "FilePath": str(target_path),
            "InstructionText": norm_text(instruction_text),
            "AddedOn": ts,
            "UpdatedOn": ts,
        })
        return doc_id

    def update_product_document(self, prod_doc_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_PRODUCT_DOCUMENTS, "ProdDocID", prod_doc_id, updates)

    def delete_product_document(self, prod_doc_id: str, delete_file: bool = False) -> bool:
        row = self.repo.find_row(AppConfig.SHEET_PRODUCT_DOCUMENTS, 0, prod_doc_id)
        if row and delete_file:
            fp = Path(norm_text(row[5]))
            if fp.exists():
                try:
                    fp.unlink()
                except Exception:
                    pass
        return self.repo.delete_row_by_key_name(AppConfig.SHEET_PRODUCT_DOCUMENTS, "ProdDocID", prod_doc_id)

    def get_product_documents(self, product_code: str) -> List[ProductDocumentRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PRODUCT_DOCUMENTS,
            lambda r: norm_text(r.get("ProductCode")) == norm_text(product_code)
        )
        output: List[ProductDocumentRecord] = []
        for r in rows:
            output.append(ProductDocumentRecord(
                prod_doc_id=norm_text(r["ProdDocID"]),
                product_code=norm_text(r["ProductCode"]),
                section_name=norm_text(r["SectionName"]),
                doc_name=norm_text(r["DocName"]),
                doc_type=norm_text(r["DocType"]),
                file_path=norm_text(r["FilePath"]),
                instruction_text=norm_text(r["InstructionText"]),
                added_on=norm_text(r["AddedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    # --------------------------------------------------------
    # Product workorders
    # --------------------------------------------------------
    def add_product_workorder(
        self,
        product_code: str,
        workorder_name: str,
        stage: str = "",
        owner: str = "",
        due_date: str = "",
        status: str = "Open",
        notes: str = "",
    ) -> str:
        self._require_workbook()
        if not norm_text(workorder_name):
            raise ValueError("Work order name is required.")

        workorder_id = CodeFactory.workorder_id(product_code, workorder_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_WORKORDERS, {
            "WorkOrderID": workorder_id,
            "OwnerType": "PRODUCT",
            "OwnerCode": product_code,
            "WorkOrderName": norm_text(workorder_name),
            "Stage": norm_text(stage),
            "Owner": norm_text(owner),
            "DueDate": norm_text(due_date),
            "Status": norm_text(status) or "Open",
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        return workorder_id

    def get_product_workorders(self, product_code: str) -> List[WorkOrderRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PRODUCT"
            and norm_text(r.get("OwnerCode")) == norm_text(product_code)
        )
        output: List[WorkOrderRecord] = []
        for r in rows:
            output.append(WorkOrderRecord(
                workorder_id=norm_text(r["WorkOrderID"]),
                owner_type=norm_text(r["OwnerType"]),
                owner_code=norm_text(r["OwnerCode"]),
                workorder_name=norm_text(r["WorkOrderName"]),
                stage=norm_text(r["Stage"]),
                owner=norm_text(r["Owner"]),
                due_date=norm_text(r["DueDate"]),
                status=norm_text(r["Status"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    def update_workorder(self, workorder_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_WORKORDERS, "WorkOrderID", workorder_id, updates)

    def delete_workorder(self, workorder_id: str) -> bool:
        return self.repo.delete_row_by_key_name(AppConfig.SHEET_WORKORDERS, "WorkOrderID", workorder_id)

    # --------------------------------------------------------
    # Bundle
    # --------------------------------------------------------
    # def get_product_bundle(self, product_code: str) -> ProductBundle:
    #     product = self.get_product(product_code)
    #     links = self.get_product_module_links(product_code)
    #     docs = self.get_product_documents(product_code)
    #     workorders = self.get_product_workorders(product_code)

    #     modules: List[ModuleRecord] = []
    #     tasks_by_module: Dict[str, List[TaskRecord]] = {}
    #     total_hours = 0.0

    #     for link in links:
    #         mod = self.module_service.get_module(link.module_code)
    #         if mod:
    #             modules.append(mod)

    #         module_tasks = self.module_service.get_module_tasks(link.module_code)
    #         tasks_by_module[link.module_code] = module_tasks

    #         module_hours = sum(to_float(t.estimated_hours) for t in module_tasks)
    #         total_hours += module_hours * max(1, link.module_qty)

    #     return ProductBundle(
    #         product=product,
    #         module_links=links,
    #         product_documents=docs,
    #         workorders=workorders,
    #         modules=modules,
    #         tasks_by_module=tasks_by_module,
    #         total_hours=total_hours,
    #     )


    def get_product_bundle(self, product_code: str) -> ProductBundle:
        product = self.get_product(product_code)
        if not product:
            return ProductBundle(
                product=None,
                module_links=[],
                product_documents=[],
                workorders=[],
                modules=[],
                tasks_by_module={},
                total_hours=0.0,
            )

        links = self.get_product_module_links(product_code)
        docs = self.get_product_documents(product_code)
        workorders = self.get_product_workorders(product_code)

        modules: List[ModuleRecord] = []
        tasks_by_module: Dict[str, List[TaskRecord]] = {}
        total_hours = 0.0

        for link in links:
            mod = self.module_service.get_module(link.module_code)
            if mod:
                modules.append(mod)

            module_tasks = self.module_service.get_module_tasks(link.module_code)
            tasks_by_module[link.module_code] = module_tasks

            module_hours = sum(to_float(t.estimated_hours) for t in module_tasks)
            total_hours += module_hours * max(1, link.module_qty)

        return ProductBundle(
            product=product,
            module_links=links,
            product_documents=docs,
            workorders=workorders,
            modules=modules,
            tasks_by_module=tasks_by_module,
            total_hours=total_hours,
        )
# ============================================================
# Project service
# ============================================================

class ProjectService(BaseService):
    def __init__(self, repo: ExcelRepository):
        super().__init__(repo)
        self.module_service = ModuleService(repo)
        self.product_service = ProductService(repo)

    # --------------------------------------------------------
    # Project CRUD
    # --------------------------------------------------------
    def create_or_update_project(
        self,
        quote_ref: str,
        project_name: str,
        client_name: str = "",
        location: str = "",
        description: str = "",
        linked_product_code: str = "",
        status: str = "Planned",
        start_date: str = "",
        due_date: str = "",
        existing_project_code: Optional[str] = None,
    ) -> str:
        self._require_workbook()
        if not norm_text(project_name):
            raise ValueError("Project name is required.")

        project_code = CodeFactory.project_code(quote_ref, project_name)
        ts = now_str()

        record = ProjectRecord(
            project_code=project_code,
            quote_ref=norm_text(quote_ref),
            project_name=norm_text(project_name),
            client_name=norm_text(client_name),
            location=norm_text(location),
            description=norm_text(description),
            linked_product_code=norm_text(linked_product_code),
            status=norm_text(status) or "Planned",
            start_date=norm_text(start_date),
            due_date=norm_text(due_date),
            created_on=ts,
            updated_on=ts,
        )

        old_row = None
        old_status = ""
        if existing_project_code:
            old_row = self.repo.find_row(AppConfig.SHEET_PROJECTS, 0, existing_project_code)
        else:
            old_row = self.repo.find_row(AppConfig.SHEET_PROJECTS, 0, project_code)

        if old_row:
            record.created_on = norm_text(old_row[10]) or ts
            old_status = norm_text(old_row[7])
            self.repo.upsert_dict(AppConfig.SHEET_PROJECTS, "ProjectCode", {
                "ProjectCode": record.project_code,
                "QuoteRef": record.quote_ref,
                "ProjectName": record.project_name,
                "ClientName": record.client_name,
                "Location": record.location,
                "Description": record.description,
                "LinkedProductCode": record.linked_product_code,
                "Status": record.status,
                "StartDate": record.start_date,
                "DueDate": record.due_date,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })
        else:
            self.repo.append_dict(AppConfig.SHEET_PROJECTS, {
                "ProjectCode": record.project_code,
                "QuoteRef": record.quote_ref,
                "ProjectName": record.project_name,
                "ClientName": record.client_name,
                "Location": record.location,
                "Description": record.description,
                "LinkedProductCode": record.linked_product_code,
                "Status": record.status,
                "StartDate": record.start_date,
                "DueDate": record.due_date,
                "CreatedOn": record.created_on,
                "UpdatedOn": record.updated_on,
            })

        if existing_project_code and norm_text(existing_project_code) != project_code:
            self._cascade_project_code_change(existing_project_code, project_code)

        became_completed = norm_text(status).lower() == "completed" and old_status.lower() != "completed"
        if became_completed:
            try:
                self.snapshot_completed_project(project_code)
            except Exception as exc:
                print(f"Completed snapshot warning: {exc}")

        return project_code

    def _cascade_project_code_change(self, old_project_code: str, new_project_code: str) -> None:
        # Project modules
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(old_project_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_MODULES,
                "LinkID",
                row["LinkID"],
                {"ProjectCode": new_project_code, "UpdatedOn": now_str()}
            )

        # Project tasks
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(old_project_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                row["ProjectTaskID"],
                {"ProjectCode": new_project_code, "UpdatedOn": now_str()}
            )

        # Project docs
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_DOCUMENTS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(old_project_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_DOCUMENTS,
                "ProjectDocID",
                row["ProjectDocID"],
                {"ProjectCode": new_project_code, "UpdatedOn": now_str()}
            )

        # Workorders
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PROJECT" and norm_text(r.get("OwnerCode")) == norm_text(old_project_code)
        )
        for row in rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_WORKORDERS,
                "WorkOrderID",
                row["WorkOrderID"],
                {"OwnerCode": new_project_code, "UpdatedOn": now_str()}
            )

    def delete_project(self, project_code: str, delete_docs_files: bool = False) -> None:
        docs = self.get_project_documents(project_code)

        self.repo.delete_row_by_key_name(AppConfig.SHEET_PROJECTS, "ProjectCode", project_code)
        self.repo.delete_rows_where(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_PROJECT_DOCUMENTS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        self.repo.delete_rows_where(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PROJECT" and norm_text(r.get("OwnerCode")) == norm_text(project_code)
        )

        if delete_docs_files:
            for d in docs:
                fp = Path(norm_text(d.file_path))
                if fp.exists():
                    try:
                        fp.unlink()
                    except Exception:
                        pass

    def get_project(self, project_code: str) -> Optional[ProjectRecord]:
        row = self.repo.find_row(AppConfig.SHEET_PROJECTS, 0, project_code)
        if not row:
            return None
        return ProjectRecord(
            project_code=norm_text(row[0]),
            quote_ref=norm_text(row[1]),
            project_name=norm_text(row[2]),
            client_name=norm_text(row[3]),
            location=norm_text(row[4]),
            description=norm_text(row[5]),
            linked_product_code=norm_text(row[6]),
            status=norm_text(row[7]),
            start_date=norm_text(row[8]),
            due_date=norm_text(row[9]),
            created_on=norm_text(row[10]),
            updated_on=norm_text(row[11]),
        )

    def list_projects(self, search_text: str = "") -> List[ProjectRecord]:
        rows = self.repo.list_rows(AppConfig.SHEET_PROJECTS)
        output: List[ProjectRecord] = []
        q = norm_text(search_text).lower()

        for row in rows:
            if not norm_text(row[0]):
                continue
            combined = " ".join(str(x or "") for x in row).lower()
            if q and q not in combined:
                continue

            output.append(ProjectRecord(
                project_code=norm_text(row[0]),
                quote_ref=norm_text(row[1]),
                project_name=norm_text(row[2]),
                client_name=norm_text(row[3]),
                location=norm_text(row[4]),
                description=norm_text(row[5]),
                linked_product_code=norm_text(row[6]),
                status=norm_text(row[7]),
                start_date=norm_text(row[8]),
                due_date=norm_text(row[9]),
                created_on=norm_text(row[10]),
                updated_on=norm_text(row[11]),
            ))

        output.sort(key=lambda x: x.updated_on, reverse=True)
        return output

    # --------------------------------------------------------
    # Project module population
    # --------------------------------------------------------
    def attach_product(self, project_code: str, product_code: str, rebuild_modules: bool = True, rebuild_tasks: bool = True) -> bool:
        ok = self.repo.update_row_by_key_name(
            AppConfig.SHEET_PROJECTS,
            "ProjectCode",
            project_code,
            {"LinkedProductCode": product_code, "UpdatedOn": now_str()}
        )
        if not ok:
            return False

        if rebuild_modules:
            self.rebuild_project_modules_from_product(project_code)
        if rebuild_tasks:
            self.populate_project_tasks_from_modules(project_code)

        return True

    def add_direct_module(
        self,
        project_code: str,
        module_code: str,
        qty: int = 1,
        stage: str = "Not Started",
        status: str = "Not Started",
        dependency_module_code: str = "",
        notes: str = "",
    ) -> str:
        existing = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
            and norm_text(r.get("ModuleCode")) == norm_text(module_code)
        )
        if existing:
            raise ValueError("This module already exists in the project.")

        existing_links = self.get_project_module_links(project_code)
        order_num = len(existing_links) + 1
        link_id = CodeFactory.project_module_link_id(project_code, module_code)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_PROJECT_MODULES, {
            "LinkID": link_id,
            "ProjectCode": project_code,
            "ModuleCode": module_code,
            "SourceType": "DIRECT",
            "SourceCode": module_code,
            "ModuleOrder": order_num,
            "ModuleQty": max(1, to_int(qty, 1)),
            "Stage": norm_text(stage) or "Not Started",
            "Status": norm_text(status) or "Not Started",
            "DependencyModuleCode": norm_text(dependency_module_code),
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        self.populate_project_tasks_from_modules(project_code)
        return link_id

    def rebuild_project_modules_from_product(self, project_code: str) -> None:
        project = self.get_project(project_code)
        if not project:
            raise ValueError("Project not found.")
        if not norm_text(project.linked_product_code):
            return

        with self.repo.batch_update():
            # Clear only product-derived rows
            self.repo.delete_rows_where(
                AppConfig.SHEET_PROJECT_MODULES,
                lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
                and norm_text(r.get("SourceType")) == "FROM_PRODUCT"
            )

            product_links = self.product_service.get_product_module_links(project.linked_product_code)
            ts = now_str()

            existing = self.get_project_module_links(project_code)
            order_base = len(existing)

            for idx, link in enumerate(product_links, start=1):
                self.repo.append_dict(AppConfig.SHEET_PROJECT_MODULES, {
                    "LinkID": CodeFactory.project_module_link_id(project_code, link.module_code),
                    "ProjectCode": project_code,
                    "ModuleCode": link.module_code,
                    "SourceType": "FROM_PRODUCT",
                    "SourceCode": project.linked_product_code,
                    "ModuleOrder": order_base + idx,
                    "ModuleQty": max(1, link.module_qty),
                    "Stage": "Not Started",
                    "Status": "Not Started",
                    "DependencyModuleCode": norm_text(link.dependency_module_code),
                    "Notes": norm_text(link.notes),
                    "CreatedOn": ts,
                    "UpdatedOn": ts,
                })

            self._normalize_project_module_order(project_code)

    def _normalize_project_module_order(self, project_code: str) -> None:
        rows = self.get_project_module_links(project_code)
        rows.sort(key=lambda x: x.module_order)
        with self.repo.batch_update():
            for idx, row in enumerate(rows, start=1):
                self.repo.update_row_by_key_name(
                    AppConfig.SHEET_PROJECT_MODULES,
                    "LinkID",
                    row.link_id,
                    {"ModuleOrder": idx, "UpdatedOn": now_str()}
                )

    def get_project_module_links(self, project_code: str) -> List[ProjectModuleLinkRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_MODULES,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        output: List[ProjectModuleLinkRecord] = []
        for r in rows:
            output.append(ProjectModuleLinkRecord(
                link_id=norm_text(r["LinkID"]),
                project_code=norm_text(r["ProjectCode"]),
                module_code=norm_text(r["ModuleCode"]),
                source_type=norm_text(r["SourceType"]),
                source_code=norm_text(r["SourceCode"]),
                module_order=to_int(r["ModuleOrder"]),
                module_qty=max(1, to_int(r["ModuleQty"], 1)),
                stage=norm_text(r["Stage"]),
                status=norm_text(r["Status"]),
                dependency_module_code=norm_text(r["DependencyModuleCode"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        output.sort(key=lambda x: x.module_order)
        return output

    def update_project_module_status(self, link_id: str, stage: str, status: str, notes: str = "") -> bool:
        return self.repo.update_row_by_key_name(
            AppConfig.SHEET_PROJECT_MODULES,
            "LinkID",
            link_id,
            {
                "Stage": norm_text(stage),
                "Status": norm_text(status),
                "Notes": norm_text(notes),
                "UpdatedOn": now_str(),
            }
        )

    # --------------------------------------------------------
    # Project task reflection / execution
    # --------------------------------------------------------
    def populate_project_tasks_from_modules(self, project_code: str, clear_existing: bool = True) -> None:
        project_links = self.get_project_module_links(project_code)

        with self.repo.batch_update():
            if clear_existing:
                self.repo.delete_rows_where(
                    AppConfig.SHEET_PROJECT_TASKS,
                    lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
                )

            # For each module, copy module tasks as executable project tasks
            source_to_project_task_id: Dict[str, str] = {}

            for pm in project_links:
                module_tasks = self.module_service.get_module_tasks(pm.module_code)

                # First pass create rows without parent/dependency remap
                pending_rows = []
                for mt in module_tasks:
                    new_pt_id = CodeFactory.project_task_id(project_code, pm.module_code, mt.task_name)
                    source_to_project_task_id[mt.task_id] = new_pt_id
                    pending_rows.append((mt, new_pt_id))

                # Second pass append with remapped parent/dependency where possible
                for mt, new_pt_id in pending_rows:
                    self.repo.append_dict(AppConfig.SHEET_PROJECT_TASKS, {
                        "ProjectTaskID": new_pt_id,
                        "ProjectCode": project_code,
                        "ModuleCode": pm.module_code,
                        "SourceTaskID": mt.task_id,
                        "ParentProjectTaskID": "",
                        "TaskName": mt.task_name,
                        "Department": mt.department,
                        "EstimatedHours": mt.estimated_hours,
                        "Stage": mt.stage,
                        "Status": "Not Started",
                        "DependencyTaskID": "",
                        "AssignedTo": "",
                        "Notes": mt.notes,
                        "CreatedOn": now_str(),
                        "UpdatedOn": now_str(),
                    })

            # Remap parent/dependency relationships after creation
            all_project_tasks = self.get_project_tasks(project_code)
            source_lookup = {pt.source_task_id: pt for pt in all_project_tasks}

            for pm in project_links:
                module_tasks = self.module_service.get_module_tasks(pm.module_code)
                for mt in module_tasks:
                    pt = source_lookup.get(mt.task_id)
                    if not pt:
                        continue

                    parent_project_task_id = ""
                    dep_project_task_id = ""

                    if norm_text(mt.parent_task_id) and mt.parent_task_id in source_lookup:
                        parent_project_task_id = source_lookup[mt.parent_task_id].project_task_id

                    if norm_text(mt.dependency_task_id) and mt.dependency_task_id in source_lookup:
                        dep_project_task_id = source_lookup[mt.dependency_task_id].project_task_id

                    self.repo.update_row_by_key_name(
                        AppConfig.SHEET_PROJECT_TASKS,
                        "ProjectTaskID",
                        pt.project_task_id,
                        {
                            "ParentProjectTaskID": parent_project_task_id,
                            "DependencyTaskID": dep_project_task_id,
                            "UpdatedOn": now_str(),
                        }
                    )

    def get_project_tasks(self, project_code: str) -> List[ProjectTaskRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        output: List[ProjectTaskRecord] = []
        for r in rows:
            output.append(ProjectTaskRecord(
                project_task_id=norm_text(r["ProjectTaskID"]),
                project_code=norm_text(r["ProjectCode"]),
                module_code=norm_text(r["ModuleCode"]),
                source_task_id=norm_text(r["SourceTaskID"]),
                parent_project_task_id=norm_text(r["ParentProjectTaskID"]),
                task_name=norm_text(r["TaskName"]),
                department=norm_text(r["Department"]),
                estimated_hours=to_float(r["EstimatedHours"]),
                stage=norm_text(r["Stage"]),
                status=norm_text(r["Status"]),
                dependency_task_id=norm_text(r["DependencyTaskID"]),
                assigned_to=norm_text(r["AssignedTo"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    def update_project_task_status(
        self,
        project_task_id: str,
        stage: str = "",
        status: str = "",
        assigned_to: str = "",
        notes: str = "",
    ) -> bool:
        return self.repo.update_row_by_key_name(
            AppConfig.SHEET_PROJECT_TASKS,
            "ProjectTaskID",
            project_task_id,
            {
                "Stage": norm_text(stage),
                "Status": norm_text(status),
                "AssignedTo": norm_text(assigned_to),
                "Notes": norm_text(notes),
                "UpdatedOn": now_str(),
            }
        )

    def delete_project_task(self, project_task_id: str) -> bool:
        # Clean references
        child_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("ParentProjectTaskID")) == norm_text(project_task_id)
        )
        for row in child_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                row["ProjectTaskID"],
                {"ParentProjectTaskID": "", "UpdatedOn": now_str()}
            )

        dep_rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_TASKS,
            lambda r: norm_text(r.get("DependencyTaskID")) == norm_text(project_task_id)
        )
        for row in dep_rows:
            self.repo.update_row_by_key_name(
                AppConfig.SHEET_PROJECT_TASKS,
                "ProjectTaskID",
                row["ProjectTaskID"],
                {"DependencyTaskID": "", "UpdatedOn": now_str()}
            )

        return self.repo.delete_row_by_key_name(AppConfig.SHEET_PROJECT_TASKS, "ProjectTaskID", project_task_id)

    # --------------------------------------------------------
    # Project docs
    # --------------------------------------------------------
    def add_project_document(
        self,
        project_code: str,
        source_file_path: str,
        section_name: str = "",
        doc_type: str = "Other",
        instruction_text: str = "",
        copy_file: bool = True,
    ) -> str:
        self._require_workbook()

        src = Path(source_file_path)
        if not src.exists():
            raise FileNotFoundError(f"Source document not found: {source_file_path}")

        target_path = src
        if copy_file:
            folder = self.repo.get_project_docs_folder(project_code)
            candidate = folder / src.name
            if candidate.exists():
                candidate = folder / f"{candidate.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{candidate.suffix}"
            import shutil
            shutil.copy2(src, candidate)
            target_path = candidate

        doc_name = target_path.name
        doc_id = CodeFactory.project_doc_id(project_code, doc_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_PROJECT_DOCUMENTS, {
            "ProjectDocID": doc_id,
            "ProjectCode": project_code,
            "SectionName": norm_text(section_name),
            "DocName": doc_name,
            "DocType": norm_text(doc_type) or "Other",
            "FilePath": str(target_path),
            "InstructionText": norm_text(instruction_text),
            "AddedOn": ts,
            "UpdatedOn": ts,
        })
        return doc_id

    def get_project_documents(self, project_code: str) -> List[ProjectDocumentRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_PROJECT_DOCUMENTS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        output: List[ProjectDocumentRecord] = []
        for r in rows:
            output.append(ProjectDocumentRecord(
                project_doc_id=norm_text(r["ProjectDocID"]),
                project_code=norm_text(r["ProjectCode"]),
                section_name=norm_text(r["SectionName"]),
                doc_name=norm_text(r["DocName"]),
                doc_type=norm_text(r["DocType"]),
                file_path=norm_text(r["FilePath"]),
                instruction_text=norm_text(r["InstructionText"]),
                added_on=norm_text(r["AddedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    def update_project_document(self, project_doc_id: str, updates: Dict[str, Any]) -> bool:
        updates = dict(updates or {})
        updates["UpdatedOn"] = now_str()
        return self.repo.update_row_by_key_name(AppConfig.SHEET_PROJECT_DOCUMENTS, "ProjectDocID", project_doc_id, updates)

    def delete_project_document(self, project_doc_id: str, delete_file: bool = False) -> bool:
        row = self.repo.find_row(AppConfig.SHEET_PROJECT_DOCUMENTS, 0, project_doc_id)
        if row and delete_file:
            fp = Path(norm_text(row[5]))
            if fp.exists():
                try:
                    fp.unlink()
                except Exception:
                    pass
        return self.repo.delete_row_by_key_name(AppConfig.SHEET_PROJECT_DOCUMENTS, "ProjectDocID", project_doc_id)

    # --------------------------------------------------------
    # Project workorders
    # --------------------------------------------------------
    def add_project_workorder(
        self,
        project_code: str,
        workorder_name: str,
        stage: str = "",
        owner: str = "",
        due_date: str = "",
        status: str = "Open",
        notes: str = "",
    ) -> str:
        if not norm_text(workorder_name):
            raise ValueError("Work order name is required.")

        workorder_id = CodeFactory.workorder_id(project_code, workorder_name)
        ts = now_str()

        self.repo.append_dict(AppConfig.SHEET_WORKORDERS, {
            "WorkOrderID": workorder_id,
            "OwnerType": "PROJECT",
            "OwnerCode": project_code,
            "WorkOrderName": norm_text(workorder_name),
            "Stage": norm_text(stage),
            "Owner": norm_text(owner),
            "DueDate": norm_text(due_date),
            "Status": norm_text(status) or "Open",
            "Notes": norm_text(notes),
            "CreatedOn": ts,
            "UpdatedOn": ts,
        })
        return workorder_id

    def get_project_workorders(self, project_code: str) -> List[WorkOrderRecord]:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PROJECT"
            and norm_text(r.get("OwnerCode")) == norm_text(project_code)
        )
        output: List[WorkOrderRecord] = []
        for r in rows:
            output.append(WorkOrderRecord(
                workorder_id=norm_text(r["WorkOrderID"]),
                owner_type=norm_text(r["OwnerType"]),
                owner_code=norm_text(r["OwnerCode"]),
                workorder_name=norm_text(r["WorkOrderName"]),
                stage=norm_text(r["Stage"]),
                owner=norm_text(r["Owner"]),
                due_date=norm_text(r["DueDate"]),
                status=norm_text(r["Status"]),
                notes=norm_text(r["Notes"]),
                created_on=norm_text(r["CreatedOn"]),
                updated_on=norm_text(r["UpdatedOn"]),
            ))
        return output

    # --------------------------------------------------------
    # Bundle
    # --------------------------------------------------------
    def get_project_bundle(self, project_code: str) -> ProjectBundle:
        project = self.get_project(project_code)
        module_links = self.get_project_module_links(project_code)
        project_tasks = self.get_project_tasks(project_code)
        project_documents = self.get_project_documents(project_code)
        workorders = self.get_project_workorders(project_code)

        total_hours = 0.0
        module_rows = self.repo.list_dicts(AppConfig.SHEET_MODULES)
        module_lookup: Dict[str, ModuleRecord] = {}
        for row in module_rows:
            module_code = norm_text(row.get("ModuleCode"))
            if not module_code:
                continue
            module_lookup[module_code] = ModuleRecord(
                module_code=module_code,
                quote_ref=norm_text(row.get("QuoteRef")),
                module_name=norm_text(row.get("ModuleName")),
                description=norm_text(row.get("Description")),
                instruction_text=norm_text(row.get("InstructionText")),
                estimated_hours=to_float(row.get("EstimatedHours")),
                stock_on_hand=to_float(row.get("StockOnHand")),
                status=norm_text(row.get("Status")),
                created_on=norm_text(row.get("CreatedOn")),
                updated_on=norm_text(row.get("UpdatedOn")),
            )

        seen = set()
        modules: List[ModuleRecord] = []
        for link in module_links:
            if link.module_code in seen:
                continue
            mod = module_lookup.get(link.module_code)
            if mod:
                modules.append(mod)
                seen.add(link.module_code)

        for task in project_tasks:
            total_hours += to_float(task.estimated_hours)

        return ProjectBundle(
            project=project,
            module_links=module_links,
            project_tasks=project_tasks,
            project_documents=project_documents,
            workorders=workorders,
            modules=modules,
            total_hours=total_hours,
        )

    def get_product_bundle(self, product_code: str) -> ProductBundle:
        product = self.get_product(product_code)
        if not product:
            return ProductBundle(
                product=None,
                module_links=[],
                product_documents=[],
                workorders=[],
                modules=[],
                tasks_by_module={},
                total_hours=0.0,
            )

        all_module_rows = self.repo.list_dicts(AppConfig.SHEET_MODULES)
        all_task_rows = self.repo.list_dicts(AppConfig.SHEET_TASKS)
        all_link_rows = self.repo.list_dicts(AppConfig.SHEET_PRODUCT_MODULES)
        all_doc_rows = self.repo.list_dicts(AppConfig.SHEET_PRODUCT_DOCUMENTS)
        all_wo_rows = self.repo.list_dicts(AppConfig.SHEET_WORKORDERS)

        links = []
        for r in all_link_rows:
            if norm_text(r.get("ProductCode")) == norm_text(product_code):
                links.append(ProductModuleLinkRecord(
                    link_id=norm_text(r["LinkID"]),
                    product_code=norm_text(r["ProductCode"]),
                    module_code=norm_text(r["ModuleCode"]),
                    module_order=to_int(r["ModuleOrder"]),
                    module_qty=max(1, to_int(r["ModuleQty"], 1)),
                    dependency_module_code=norm_text(r["DependencyModuleCode"]),
                    notes=norm_text(r["Notes"]),
                    created_on=norm_text(r["CreatedOn"]),
                    updated_on=norm_text(r["UpdatedOn"]),
                ))
        links.sort(key=lambda x: x.module_order)

        module_lookup = {}
        for r in all_module_rows:
            code = norm_text(r.get("ModuleCode"))
            if code:
                module_lookup[code] = ModuleRecord(
                    module_code=code,
                    quote_ref=norm_text(r["QuoteRef"]),
                    module_name=norm_text(r["ModuleName"]),
                    description=norm_text(r["Description"]),
                    instruction_text=norm_text(r["InstructionText"]),
                    estimated_hours=to_float(r["EstimatedHours"]),
                    stock_on_hand=to_float(r["StockOnHand"]),
                    status=norm_text(r["Status"]),
                    created_on=norm_text(r["CreatedOn"]),
                    updated_on=norm_text(r["UpdatedOn"]),
                )

        task_rows_by_module: Dict[str, List[Dict[str, Any]]] = {}
        for r in all_task_rows:
            if norm_text(r.get("OwnerType")) != "MODULE":
                continue
            owner_code = norm_text(r.get("OwnerCode"))
            if not owner_code:
                continue
            task_rows_by_module.setdefault(owner_code, []).append(r)

        tasks_by_module = {}
        total_hours = 0.0
        for link in links:
            mod_tasks = []
            for r in task_rows_by_module.get(link.module_code, []):
                mod_tasks.append(TaskRecord(
                    task_id=norm_text(r["TaskID"]),
                    owner_type=norm_text(r["OwnerType"]),
                    owner_code=norm_text(r["OwnerCode"]),
                    task_name=norm_text(r["TaskName"]),
                    department=norm_text(r["Department"]),
                    estimated_hours=to_float(r["EstimatedHours"]),
                    parent_task_id=norm_text(r["ParentTaskID"]),
                    dependency_task_id=norm_text(r["DependencyTaskID"]),
                    stage=norm_text(r["Stage"]),
                    status=norm_text(r["Status"]),
                    notes=norm_text(r["Notes"]),
                    created_on=norm_text(r["CreatedOn"]),
                    updated_on=norm_text(r["UpdatedOn"]),
                ))
            tasks_by_module[link.module_code] = mod_tasks
            total_hours += sum(float(t.estimated_hours or 0.0) for t in mod_tasks) * link.module_qty

        docs = []
        for r in all_doc_rows:
            if norm_text(r.get("ProductCode")) == norm_text(product_code):
                docs.append(ProductDocumentRecord(
                    prod_doc_id=norm_text(r["ProdDocID"]),
                    product_code=norm_text(r["ProductCode"]),
                    section_name=norm_text(r["SectionName"]),
                    doc_name=norm_text(r["DocName"]),
                    doc_type=norm_text(r["DocType"]),
                    file_path=norm_text(r["FilePath"]),
                    instruction_text=norm_text(r["InstructionText"]),
                    added_on=norm_text(r["AddedOn"]),
                    updated_on=norm_text(r["UpdatedOn"]),
                ))

        workorders = []
        for r in all_wo_rows:
            if norm_text(r.get("OwnerType")) == "PRODUCT" and norm_text(r.get("OwnerCode")) == norm_text(product_code):
                workorders.append(WorkOrderRecord(
                    workorder_id=norm_text(r["WorkOrderID"]),
                    owner_type=norm_text(r["OwnerType"]),
                    owner_code=norm_text(r["OwnerCode"]),
                    workorder_name=norm_text(r["WorkOrderName"]),
                    stage=norm_text(r["Stage"]),
                    owner=norm_text(r["Owner"]),
                    due_date=norm_text(r["DueDate"]),
                    status=norm_text(r["Status"]),
                    notes=norm_text(r["Notes"]),
                    created_on=norm_text(r["CreatedOn"]),
                    updated_on=norm_text(r["UpdatedOn"]),
                ))

        modules = [module_lookup[link.module_code] for link in links if link.module_code in module_lookup]

        return ProductBundle(
            product=product,
            module_links=links,
            product_documents=docs,
            workorders=workorders,
            modules=modules,
            tasks_by_module=tasks_by_module,
            total_hours=total_hours,
        )

    def _get_direct_product_parts(self, product_code: str):
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_COMPONENTS,
            lambda r: norm_text(r.get("OwnerType")) == "PRODUCT"
            and norm_text(r.get("OwnerCode")) == norm_text(product_code)
        )
        return rows

    def _ensure_completed_job_sheets(self) -> None:
        if getattr(self.repo, "backend_name", "excel") == "excel":
            workbook_path = self.repo.require_workbook_path()
            WorkbookSchema.ensure_workbook_structure(workbook_path)
        self.repo.reload_cache()

    def _has_completed_snapshot(self, project_code: str) -> bool:
        rows = self.repo.filter_dicts(
            AppConfig.SHEET_COMPLETED_JOBS,
            lambda r: norm_text(r.get("ProjectCode")) == norm_text(project_code)
        )
        return any(norm_text(r.get("SnapshotID")) for r in rows)

    def _backfill_completed_project_snapshots(self) -> None:
        completed_projects = [
            project for project in self.list_projects("")
            if norm_text(getattr(project, "status", "")).lower() == "completed"
        ]
        for project in completed_projects:
            if not self._has_completed_snapshot(project.project_code):
                self.snapshot_completed_project(project.project_code)
    
    def list_completed_jobs(self):
        self._ensure_completed_job_sheets()
        self._backfill_completed_project_snapshots()
        try:
            rows = self.repo.list_dicts(AppConfig.SHEET_COMPLETED_JOBS)
        except ValueError as exc:
            if "Sheet not found" not in str(exc):
                raise
            self._ensure_completed_job_sheets()
            rows = self.repo.list_dicts(AppConfig.SHEET_COMPLETED_JOBS)
        out = []
        for r in rows:
            if not norm_text(r.get("SnapshotID")):
                continue
            out.append(SimpleNamespace(
                snapshot_id=norm_text(r.get("SnapshotID")),
                project_code=norm_text(r.get("ProjectCode")),
                quote_ref=norm_text(r.get("QuoteRef")),
                product_code=norm_text(r.get("ProductCode")),
                product_name=norm_text(r.get("ProductName")),
                client_name=norm_text(r.get("ClientName")),
                completed_on=norm_text(r.get("CompletedOn")),
                labour_hours=to_float(r.get("LabourHours")),
                parts_total=to_float(r.get("PartsTotal")),
                grand_total=to_float(r.get("GrandTotal")),
                notes=norm_text(r.get("Notes")),
            ))
        out.sort(key=lambda x: x.completed_on, reverse=True)
        return out 
    

    def snapshot_completed_project(self, project_code: str) -> str:
        self._ensure_completed_job_sheets()
        project = self.get_project(project_code)
        if not project:
            raise ValueError("Project not found.")

        ts = now_str()
        snapshot_id = f"SNAP::{project_code}::{datetime.now().strftime('%Y%m%d%H%M%S')}"

        bundle = self.get_project_bundle(project_code)
        product_bundle = None
        if norm_text(project.linked_product_code):
            product_bundle = self.product_service.get_product_bundle(project.linked_product_code)

        labour_hours = sum(to_float(getattr(t, "estimated_hours", 0)) for t in (bundle.project_tasks or []))
        parts_total = 0.0

        product_name = ""
        if product_bundle and getattr(product_bundle, "product", None):
            product_name = norm_text(product_bundle.product.product_name)

        self.repo.append_dict(AppConfig.SHEET_COMPLETED_JOBS, {
            "SnapshotID": snapshot_id,
            "ProjectCode": project.project_code,
            "QuoteRef": project.quote_ref,
            "ProductCode": project.linked_product_code,
            "ProductName": product_name,
            "ClientName": project.client_name,
            "CompletedOn": ts,
            "LabourHours": labour_hours,
            "PartsTotal": 0.0,
            "GrandTotal": 0.0,
            "Notes": "Frozen snapshot created on completion.",
        })

        line_no = 1

        if product_bundle:
            # Assembly breakdown
            for link in product_bundle.module_links or []:
                module_tasks = (product_bundle.tasks_by_module or {}).get(link.module_code, [])
                qty = max(1, to_int(getattr(link, "module_qty", 1), 1))
                hours = sum(to_float(getattr(t, "estimated_hours", 0)) for t in module_tasks) * qty

                self.repo.append_dict(AppConfig.SHEET_COMPLETED_JOB_LINES, {
                    "SnapshotLineID": f"{snapshot_id}-ASM-{line_no:03d}",
                    "SnapshotID": snapshot_id,
                    "LineType": "ASSEMBLY",
                    "Code": norm_text(link.module_code),
                    "Description": norm_text(link.module_code),
                    "PartNumber": "",
                    "Qty": qty,
                    "Hours": hours,
                    "UnitPrice": 0.0,
                    "LineTotal": 0.0,
                    "Source": "PRODUCT_MODULE",
                })
                line_no += 1

            # Direct product parts
            direct_parts = self._get_direct_product_parts(project.linked_product_code)
            for comp in direct_parts:
                notes = norm_text(comp.get("Notes"))
                unit_price = 0.0
                m = re.search(r"UnitPrice\\s*=\\s*([0-9]+(?:\\.[0-9]+)?)", notes, re.I)
                if m:
                    unit_price = to_float(m.group(1))

                qty = to_float(comp.get("Qty"))
                line_total = unit_price * qty
                parts_total += line_total

                self.repo.append_dict(AppConfig.SHEET_COMPLETED_JOB_LINES, {
                    "SnapshotLineID": f"{snapshot_id}-PART-{line_no:03d}",
                    "SnapshotID": snapshot_id,
                    "LineType": "PART",
                    "Code": norm_text(comp.get("ComponentName")),
                    "Description": norm_text(comp.get("ComponentName")),
                    "PartNumber": norm_text(comp.get("PartNumber")),
                    "Qty": qty,
                    "Hours": 0.0,
                    "UnitPrice": unit_price,
                    "LineTotal": line_total,
                    "Source": "PRODUCT_DIRECT_PART",
                })
                line_no += 1

        grand_total = parts_total

        # update header totals
        self.repo.update_row_by_key_name(
            AppConfig.SHEET_COMPLETED_JOBS,
            "SnapshotID",
            snapshot_id,
            {
                "PartsTotal": parts_total,
                "GrandTotal": grand_total,
            }
        )

        return snapshot_id
    
    def get_completed_job_lines(self, snapshot_id: str):
        try:
            rows = self.repo.filter_dicts(
                AppConfig.SHEET_COMPLETED_JOB_LINES,
                lambda r: norm_text(r.get("SnapshotID")) == norm_text(snapshot_id)
            )
        except ValueError as exc:
            if "Sheet not found" not in str(exc):
                raise
            self._ensure_completed_job_sheets()
            rows = self.repo.filter_dicts(
                AppConfig.SHEET_COMPLETED_JOB_LINES,
                lambda r: norm_text(r.get("SnapshotID")) == norm_text(snapshot_id)
            )
        out = []
        for r in rows:
            out.append(SimpleNamespace(
                snapshot_line_id=norm_text(r.get("SnapshotLineID")),
                snapshot_id=norm_text(r.get("SnapshotID")),
                line_type=norm_text(r.get("LineType")),
                code=norm_text(r.get("Code")),
                description=norm_text(r.get("Description")),
                part_number=norm_text(r.get("PartNumber")),
                qty=to_float(r.get("Qty")),
                hours=to_float(r.get("Hours")),
                unit_price=to_float(r.get("UnitPrice")),
                line_total=to_float(r.get("LineTotal")),
                source=norm_text(r.get("Source")),
            ))
        return out

# ============================================================
# Scheduler / dependency helper service
# ============================================================

class SchedulerService(BaseService):
    def __init__(self, repo: ExcelRepository):
        super().__init__(repo)
        self.project_service = ProjectService(repo)
        self.product_service = ProductService(repo)
        self.module_service = ModuleService(repo)

    def get_department_workload_for_project(self, project_code: str) -> Dict[str, float]:
        tasks = self.project_service.get_project_tasks(project_code)
        output: Dict[str, float] = {}
        for task in tasks:
            dept = norm_text(task.department) or "Unassigned"
            output[dept] = output.get(dept, 0.0) + to_float(task.estimated_hours)
        return dict(sorted(output.items(), key=lambda x: x[0]))

    def get_open_blockers_for_project(self, project_code: str) -> List[Dict[str, Any]]:
        project_tasks = self.project_service.get_project_tasks(project_code)
        task_lookup = {t.project_task_id: t for t in project_tasks}

        blockers: List[Dict[str, Any]] = []

        for task in project_tasks:
            dep_id = norm_text(task.dependency_task_id)
            if dep_id and dep_id in task_lookup:
                dep_task = task_lookup[dep_id]
                if norm_text(dep_task.status) != "Completed":
                    blockers.append({
                        "type": "TASK",
                        "project_code": project_code,
                        "task_name": task.task_name,
                        "depends_on": dep_task.task_name,
                        "current_status": task.status,
                        "dependency_status": dep_task.status,
                    })

        project_modules = self.project_service.get_project_module_links(project_code)
        mod_lookup = {m.module_code: m for m in project_modules}

        for mod in project_modules:
            dep_mod_code = norm_text(mod.dependency_module_code)
            if dep_mod_code and dep_mod_code in mod_lookup:
                dep_mod = mod_lookup[dep_mod_code]
                if norm_text(dep_mod.status) != "Completed":
                    blockers.append({
                        "type": "MODULE",
                        "project_code": project_code,
                        "module_code": mod.module_code,
                        "depends_on": dep_mod.module_code,
                        "current_status": mod.status,
                        "dependency_status": dep_mod.status,
                    })

        return blockers


# ============================================================
# Facade for all services
# ============================================================

class ERPServiceHub:
    """
    Single service container so UI can access:
    - hub.modules
    - hub.products
    - hub.projects
    - hub.scheduler
    """

    def __init__(self, repo: ExcelRepository):
        self.repo = repo
        self.modules = ModuleService(repo)
        self.products = ProductService(repo)
        self.projects = ProjectService(repo)
        self.completed_jobs = self.projects
        self.scheduler = SchedulerService(repo)

# ============================================================
# PATCH: product direct parts + live order auto tracker helpers
# ============================================================


_DEPARTMENT_DAY_DEFAULT = {
    "automation": 1,
    "electrical": 1,
    "assembly": 1,
    "operations": 1,
    "fabrication": 2,
    "procurement": 2,
    "software": 2,
    "mechanical": 2,
    "ga drawing": 1,
    "logistics": 1,
    "hr": 1,
    "quote/order": 1,
}


def _parse_iso_date_safe(value: str):
    value = norm_text(value)
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            pass
    return None


def _format_iso_date_safe(dt):
    if not dt:
        return ""
    return dt.strftime("%Y-%m-%d")


def _dept_default_days(department: str) -> int:
    dept = norm_text(department).lower()
    if not dept:
        return 1
    for key, val in _DEPARTMENT_DAY_DEFAULT.items():
        if key in dept:
            return val
    return 1


def _product_add_part(self, product_code: str, component_name: str, qty: float = 1.0, soh_qty: float = 0.0,
                      preferred_supplier: str = "", lead_time_days: int = 0, part_number: str = "", notes: str = "") -> str:
    self._require_workbook()
    if not norm_text(product_code):
        raise ValueError("Product code is required.")
    if not norm_text(component_name):
        raise ValueError("Part name is required.")
    component_id = CodeFactory.component_id(product_code, component_name)
    ts = now_str()
    self.repo.append_dict(AppConfig.SHEET_COMPONENTS, {
        "ComponentID": component_id,
        "OwnerType": "PRODUCT",
        "OwnerCode": product_code,
        "ComponentName": norm_text(component_name),
        "Qty": to_float(qty),
        "SOHQty": to_float(soh_qty),
        "PreferredSupplier": norm_text(preferred_supplier),
        "LeadTimeDays": to_int(lead_time_days),
        "PartNumber": norm_text(part_number),
        "Notes": norm_text(notes),
        "CreatedOn": ts,
        "UpdatedOn": ts,
    })
    return component_id


def _product_get_parts(self, product_code: str):
    rows = self.repo.filter_dicts(
        AppConfig.SHEET_COMPONENTS,
        lambda r: norm_text(r.get("OwnerType")) == "PRODUCT" and norm_text(r.get("OwnerCode")) == norm_text(product_code)
    )
    out = []
    for r in rows:
        out.append(ComponentRecord(
            component_id=norm_text(r["ComponentID"]),
            owner_type=norm_text(r["OwnerType"]),
            owner_code=norm_text(r["OwnerCode"]),
            component_name=norm_text(r["ComponentName"]),
            qty=to_float(r["Qty"]),
            soh_qty=to_float(r["SOHQty"]),
            preferred_supplier=norm_text(r["PreferredSupplier"]),
            lead_time_days=to_int(r["LeadTimeDays"]),
            part_number=norm_text(r["PartNumber"]),
            notes=norm_text(r["Notes"]),
            created_on=norm_text(r["CreatedOn"]),
            updated_on=norm_text(r["UpdatedOn"]),
        ))
    return out


def _project_get_parts(self, project_code: str):
    rows = self.repo.filter_dicts(
        AppConfig.SHEET_COMPONENTS,
        lambda r: norm_text(r.get("OwnerType")) == "PROJECT" and norm_text(r.get("OwnerCode")) == norm_text(project_code)
    )
    out = []
    for r in rows:
        out.append(ComponentRecord(
            component_id=norm_text(r["ComponentID"]),
            owner_type=norm_text(r["OwnerType"]),
            owner_code=norm_text(r["OwnerCode"]),
            component_name=norm_text(r["ComponentName"]),
            qty=to_float(r["Qty"]),
            soh_qty=to_float(r["SOHQty"]),
            preferred_supplier=norm_text(r["PreferredSupplier"]),
            lead_time_days=to_int(r["LeadTimeDays"]),
            part_number=norm_text(r["PartNumber"]),
            notes=norm_text(r["Notes"]),
            created_on=norm_text(r["CreatedOn"]),
            updated_on=norm_text(r["UpdatedOn"]),
        ))
    return out


def _project_add_direct_part(self, project_code: str, component_name: str, qty: float = 1.0, soh_qty: float = 0.0,
                             preferred_supplier: str = "", lead_time_days: int = 0, part_number: str = "", notes: str = "") -> str:
    self._require_workbook()
    if not norm_text(project_code):
        raise ValueError("Project / order code is required.")
    if not norm_text(component_name):
        raise ValueError("Part name is required.")
    component_id = CodeFactory.component_id(project_code, component_name)
    ts = now_str()
    self.repo.append_dict(AppConfig.SHEET_COMPONENTS, {
        "ComponentID": component_id,
        "OwnerType": "PROJECT",
        "OwnerCode": project_code,
        "ComponentName": norm_text(component_name),
        "Qty": to_float(qty),
        "SOHQty": to_float(soh_qty),
        "PreferredSupplier": norm_text(preferred_supplier),
        "LeadTimeDays": to_int(lead_time_days),
        "PartNumber": norm_text(part_number),
        "Notes": norm_text(notes),
        "CreatedOn": ts,
        "UpdatedOn": ts,
    })
    return component_id


def _project_sync_parts_from_product(self, project_code: str) -> None:
    project = self.get_project(project_code)
    if not project or not norm_text(project.linked_product_code):
        return
    with self.repo.batch_update():
        self.repo.delete_rows_where(
            AppConfig.SHEET_COMPONENTS,
            lambda r: norm_text(r.get("OwnerType")) == "PROJECT"
            and norm_text(r.get("OwnerCode")) == norm_text(project_code)
            and norm_text(r.get("Notes")).startswith("FROM_PRODUCT:")
        )
        for part in self.product_service.get_product_parts(project.linked_product_code):
            self.add_direct_project_part(
                project_code=project_code,
                component_name=part.component_name,
                qty=part.qty,
                soh_qty=part.soh_qty,
                preferred_supplier=part.preferred_supplier,
                lead_time_days=part.lead_time_days,
                part_number=part.part_number,
                notes=f"FROM_PRODUCT:{project.linked_product_code} | {part.notes}",
            )


def _project_autogenerate_workorders(self, project_code: str) -> None:
    project = self.get_project(project_code)
    if not project:
        return
    existing_manual = [
        w for w in self.get_project_workorders(project_code)
        if not norm_text(w.notes).startswith("AUTO_TRACKER")
    ]
    with self.repo.batch_update():
        self.repo.delete_rows_where(
            AppConfig.SHEET_WORKORDERS,
            lambda r: norm_text(r.get("OwnerType")) == "PROJECT"
            and norm_text(r.get("OwnerCode")) == norm_text(project_code)
            and norm_text(r.get("Notes")).startswith("AUTO_TRACKER")
        )
        module_links = self.get_project_module_links(project_code)
        project_tasks = self.get_project_tasks(project_code)
        start_dt = _parse_iso_date_safe(project.start_date) or _parse_iso_date_safe(project.due_date)
        current_dt = start_dt
        for link in module_links:
            module_tasks = [t for t in project_tasks if norm_text(t.module_code) == norm_text(link.module_code)]
            depts = sorted({norm_text(t.department) for t in module_tasks if norm_text(t.department)}) or ["Assembly"]
            for dept in depts:
                days = _dept_default_days(dept)
                due_dt = current_dt + timedelta(days=max(0, days - 1)) if current_dt else None
                wo_name = f"{link.module_code} | {dept}"
                self.add_project_workorder(
                    project_code=project_code,
                    workorder_name=wo_name,
                    stage=dept,
                    owner=dept,
                    due_date=_format_iso_date_safe(due_dt),
                    status="Open",
                    notes=f"AUTO_TRACKER | MODULE={link.module_code} | SOURCE={link.source_type}",
                )
                if current_dt:
                    current_dt = due_dt + timedelta(days=1)
    # preserve manual workorders, autogenerated rows already appended after delete
    return None


# _ProductService_get_product_bundle_orig = ProductService.get_product_bundle
_ProjectService_attach_product_orig = ProjectService.attach_product
_ProjectService_add_direct_module_orig = ProjectService.add_direct_module
_ProjectService_rebuild_project_modules_from_product_orig = ProjectService.rebuild_project_modules_from_product
_ProjectService_populate_project_tasks_from_modules_orig = ProjectService.populate_project_tasks_from_modules
_ProjectService_get_project_bundle_orig = ProjectService.get_project_bundle

ProductService.add_product_part = _product_add_part
ProductService.get_product_parts = _product_get_parts
ProjectService.get_project_parts = _project_get_parts
ProjectService.add_direct_project_part = _project_add_direct_part
ProjectService.sync_project_parts_from_product = _project_sync_parts_from_product
ProjectService.auto_generate_order_workorders = _project_autogenerate_workorders


def _patched_attach_product(self, project_code: str, product_code: str, rebuild_modules: bool = True, rebuild_tasks: bool = True) -> bool:
    ok = _ProjectService_attach_product_orig(self, project_code, product_code, rebuild_modules=rebuild_modules, rebuild_tasks=rebuild_tasks)
    if ok:
        try:
            self.sync_project_parts_from_product(project_code)
        except Exception:
            pass
        try:
            self.auto_generate_order_workorders(project_code)
        except Exception:
            pass
    return ok


def _patched_add_direct_module(self, *args, **kwargs):
    result = _ProjectService_add_direct_module_orig(self, *args, **kwargs)
    try:
        project_code = args[0] if args else kwargs.get('project_code', '')
        self.auto_generate_order_workorders(project_code)
    except Exception:
        pass
    return result


def _patched_rebuild_modules_from_product(self, project_code: str) -> None:
    _ProjectService_rebuild_project_modules_from_product_orig(self, project_code)
    try:
        self.sync_project_parts_from_product(project_code)
    except Exception:
        pass
    try:
        self.auto_generate_order_workorders(project_code)
    except Exception:
        pass


def _patched_populate_tasks(self, project_code: str, clear_existing: bool = True) -> None:
    _ProjectService_populate_project_tasks_from_modules_orig(self, project_code, clear_existing=clear_existing)
    try:
        self.auto_generate_order_workorders(project_code)
    except Exception:
        pass


def _patched_get_project_bundle(self, project_code: str) -> ProjectBundle:
    bundle = _ProjectService_get_project_bundle_orig(self, project_code)
    try:
        bundle.project_parts = self.get_project_parts(project_code)
    except Exception:
        bundle.project_parts = []
    return bundle

# ProductService.get_product_bundle = _ProductService_get_product_bundle_orig
ProjectService.attach_product = _patched_attach_product
ProjectService.add_direct_module = _patched_add_direct_module
ProjectService.rebuild_project_modules_from_product = _patched_rebuild_modules_from_product
ProjectService.populate_project_tasks_from_modules = _patched_populate_tasks
ProjectService.get_project_bundle = _patched_get_project_bundle


# ============================================================
# Compatibility shims for renamed UI (parts / assemblies / orders)
# ============================================================

def _module_list_parts(self, search_text: str = ""):
    self._require_workbook()
    q = norm_text(search_text).lower()
    rows = self.repo.read_sheet_as_dicts(AppConfig.SHEET_COMPONENTS)
    out = []
    for r in rows:
        item = ComponentRecord(
            component_id=norm_text(r.get("ComponentID")),
            owner_type=norm_text(r.get("OwnerType")),
            owner_code=norm_text(r.get("OwnerCode")),
            component_name=norm_text(r.get("ComponentName")),
            qty=to_float(r.get("Qty")),
            soh_qty=to_float(r.get("SOHQty")),
            preferred_supplier=norm_text(r.get("PreferredSupplier")),
            lead_time_days=to_int(r.get("LeadTimeDays")),
            part_number=norm_text(r.get("PartNumber")),
            notes=norm_text(r.get("Notes")),
            created_on=norm_text(r.get("CreatedOn")),
            updated_on=norm_text(r.get("UpdatedOn")),
        )
        hay = " ".join([
            item.component_id, item.owner_type, item.owner_code, item.component_name,
            item.preferred_supplier, item.part_number, item.notes
        ]).lower()
        if q and q not in hay:
            continue
        out.append(item)
    out.sort(key=lambda x: (x.component_name.lower(), x.owner_code.lower(), x.component_id.lower()))
    return out


def _module_get_part(self, component_id: str):
    self._require_workbook()
    row = self.repo.find_row_by_key_name(AppConfig.SHEET_COMPONENTS, "ComponentID", component_id)
    if not row:
        return None
    return ComponentRecord(
        component_id=norm_text(row.get("ComponentID")),
        owner_type=norm_text(row.get("OwnerType")),
        owner_code=norm_text(row.get("OwnerCode")),
        component_name=norm_text(row.get("ComponentName")),
        qty=to_float(row.get("Qty")),
        soh_qty=to_float(row.get("SOHQty")),
        preferred_supplier=norm_text(row.get("PreferredSupplier")),
        lead_time_days=to_int(row.get("LeadTimeDays")),
        part_number=norm_text(row.get("PartNumber")),
        notes=norm_text(row.get("Notes")),
        created_on=norm_text(row.get("CreatedOn")),
        updated_on=norm_text(row.get("UpdatedOn")),
    )


def _module_add_part(self, component_name: str, qty: float = 0.0, soh_qty: float = 0.0,
                     preferred_supplier: str = "", lead_time_days: int = 0,
                     part_number: str = "", notes: str = "", owner_code: str = "PART_MASTER"):
    self._require_workbook()
    if not norm_text(component_name):
        raise ValueError("Part name is required.")
    component_id = CodeFactory.component_id(owner_code, component_name)
    ts = now_str()
    self.repo.append_dict(AppConfig.SHEET_COMPONENTS, {
        "ComponentID": component_id,
        "OwnerType": "PART",
        "OwnerCode": norm_text(owner_code) or "PART_MASTER",
        "ComponentName": norm_text(component_name),
        "Qty": to_float(qty),
        "SOHQty": to_float(soh_qty),
        "PreferredSupplier": norm_text(preferred_supplier),
        "LeadTimeDays": to_int(lead_time_days),
        "PartNumber": norm_text(part_number),
        "Notes": norm_text(notes),
        "CreatedOn": ts,
        "UpdatedOn": ts,
    })
    return component_id


def _module_list_assemblies(self, search_text: str = ""):
    return self.list_modules(search_text)


def _project_list_orders(self, search_text: str = ""):
    return self.list_projects(search_text)


ModuleService.list_parts = _module_list_parts
ModuleService.get_part = _module_get_part
ModuleService.add_part = _module_add_part
ModuleService.list_assemblies = _module_list_assemblies
ProjectService.list_orders = _project_list_orders


class ERPServiceHub(ERPServiceHub):
    def __init__(self, repo: ExcelRepository):
        super().__init__(repo)
        self.parts = self.modules
        self.assemblies = self.modules
        self.orders = self.projects
