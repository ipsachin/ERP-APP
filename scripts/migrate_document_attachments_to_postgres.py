from __future__ import annotations

import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app_config import AppConfig
from db_config import load_project_env
from storage import PostgresRepository, norm_text


DOC_SHEETS = [
    (AppConfig.SHEET_DOCUMENTS, "DocID"),
    (AppConfig.SHEET_PRODUCT_DOCUMENTS, "ProdDocID"),
    (AppConfig.SHEET_PROJECT_DOCUMENTS, "ProjectDocID"),
]


def main() -> int:
    env_path = load_project_env()
    repo = PostgresRepository()
    print(f"Loaded environment from: {env_path}")

    uploaded = 0
    skipped_missing = 0
    skipped_existing = 0

    for sheet_name, id_field in DOC_SHEETS:
        rows = repo.list_dicts(sheet_name)
        print(f"Scanning {sheet_name}: {len(rows)} record(s)")
        for row in rows:
            record_id = norm_text(row.get(id_field))
            file_path = norm_text(row.get("FilePath"))
            doc_name = norm_text(row.get("DocName")) or Path(file_path).name or record_id

            if repo.get_document_blob(sheet_name, record_id):
                skipped_existing += 1
                continue

            src = Path(file_path)
            if file_path.startswith("db://"):
                skipped_existing += 1
                continue
            if not src.exists():
                skipped_missing += 1
                print(f"Missing file, skipped: {sheet_name} {record_id} -> {file_path}")
                continue

            created_on = norm_text(row.get("AddedOn"))
            updated_on = norm_text(row.get("UpdatedOn")) or created_on
            db_uri = repo.save_document_blob(
                sheet_name=sheet_name,
                record_id=record_id,
                file_name=doc_name,
                source_file_path=str(src),
                created_on=created_on or None,
                updated_on=updated_on or None,
            )
            repo.update_row_by_key_name(sheet_name, id_field, record_id, {"FilePath": db_uri})
            uploaded += 1

    print("\nAttachment migration complete.")
    print(f"Uploaded: {uploaded}")
    print(f"Skipped existing: {skipped_existing}")
    print(f"Skipped missing: {skipped_missing}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
