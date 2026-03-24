from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app_config import AppConfig
from db_config import load_project_env
from scripts.migrate_workbook_to_postgres import print_validation_report, validate_workbook_schema


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate workbook sheets and headers before PostgreSQL import.")
    parser.add_argument(
        "--workbook",
        default=os.getenv("ERP_WORKBOOK_PATH", str(ROOT_DIR / AppConfig.DEFAULT_WORKBOOK_NAME)),
        help="Path to the workbook to validate.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    env_path = load_project_env()
    workbook_path = Path(args.workbook).expanduser().resolve()

    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}")
        return 1

    print(f"Loaded environment from: {env_path}")
    report = validate_workbook_schema(workbook_path)
    print_validation_report(report)
    return 1 if report.has_blocking_issues else 0


if __name__ == "__main__":
    raise SystemExit(main())
