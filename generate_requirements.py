#!/usr/bin/env python3
"""
Generate requirements.txt by scanning Python files in a project.

What it does:
- Recursively scans all .py files under a project directory
- Extracts imported module names via AST
- Ignores:
  - standard library modules
  - local project modules/packages
  - common virtualenv/cache/build folders
- Maps import names to installed distribution names when possible
- Writes requirements.txt with pinned versions

Usage:
    python generate_requirements.py
    python generate_requirements.py /path/to/project
    python generate_requirements.py /path/to/project --output requirements.txt
"""

from __future__ import annotations

import argparse
import ast
import os
import sys
from pathlib import Path
from typing import Iterable, Set

try:
    import importlib.metadata as metadata
except ImportError:  # pragma: no cover
    import importlib_metadata as metadata  # type: ignore


EXCLUDED_DIRS = {
    ".git",
    ".idea",
    ".vscode",
    ".venv",
    "venv",
    "env",
    "__pycache__",
    "build",
    "dist",
    "node_modules",
    ".mypy_cache",
    ".pytest_cache",
    ".tox",
}

EXCLUDED_FILES = {
    "__init__.py",
}

# Some imports are commonly used under names that differ from the pip package
# name, so we force-map them here when needed.
IMPORT_TO_DIST_OVERRIDES = {
    "PIL": "Pillow",
    "cv2": "opencv-python",
    "yaml": "PyYAML",
    "sklearn": "scikit-learn",
    "dotenv": "python-dotenv",
    "Crypto": "pycryptodome",
    "fitz": "PyMuPDF",
    "OpenSSL": "pyOpenSSL",
    "bs4": "beautifulsoup4",
    "multipart": "python-multipart",
    "rest_framework": "djangorestframework",
}


def get_stdlib_modules() -> Set[str]:
    """Return a set of standard library module names."""
    stdlib = set(sys.builtin_module_names)

    # Python 3.10+ has sys.stdlib_module_names
    stdlib_names = getattr(sys, "stdlib_module_names", None)
    if stdlib_names:
        stdlib.update(stdlib_names)

    return stdlib


def discover_local_modules(project_root: Path) -> Set[str]:
    """
    Discover top-level local module/package names in the project.
    Example:
      project_root/foo.py       -> foo
      project_root/bar/__init__.py -> bar
    """
    local_modules: Set[str] = set()

    for item in project_root.iterdir():
        if item.name.startswith("."):
            continue
        if item.name in EXCLUDED_DIRS:
            continue

        if item.is_file() and item.suffix == ".py":
            local_modules.add(item.stem)
        elif item.is_dir():
            init_file = item / "__init__.py"
            if init_file.exists():
                local_modules.add(item.name)

    return local_modules


def find_python_files(project_root: Path) -> Iterable[Path]:
    """Yield all Python files under project_root, excluding common junk folders."""
    for root, dirs, files in os.walk(project_root):
        dirs[:] = [d for d in dirs if d not in EXCLUDED_DIRS and not d.startswith(".")]

        for filename in files:
            if not filename.endswith(".py"):
                continue
            if filename in EXCLUDED_FILES:
                continue
            yield Path(root) / filename


def extract_imports_from_file(py_file: Path) -> Set[str]:
    """Parse a Python file and return imported top-level module names."""
    imports: Set[str] = set()

    try:
        source = py_file.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        source = py_file.read_text(encoding="latin-1")

    try:
        tree = ast.parse(source, filename=str(py_file))
    except SyntaxError:
        print(f"Warning: Skipping file with syntax error: {py_file}")
        return imports

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                top_name = alias.name.split(".")[0]
                if top_name:
                    imports.add(top_name)

        elif isinstance(node, ast.ImportFrom):
            # Skip relative imports like: from . import x / from ..foo import bar
            if node.level and node.level > 0:
                continue
            if node.module:
                top_name = node.module.split(".")[0]
                if top_name:
                    imports.add(top_name)

    return imports


def build_import_to_dist_map() -> dict[str, str]:
    """
    Build a mapping from import name -> installed distribution name.
    Example: 'yaml' -> 'PyYAML'
    """
    mapping: dict[str, str] = {}

    try:
        pkg_map = metadata.packages_distributions() or {}
    except Exception:
        pkg_map = {}

    for import_name, dist_names in pkg_map.items():
        if dist_names:
            # If multiple dists provide the same import, take the first.
            mapping[import_name] = dist_names[0]

    # Apply manual overrides last so they win
    mapping.update(IMPORT_TO_DIST_OVERRIDES)
    return mapping


def resolve_distribution_name(import_name: str, import_to_dist: dict[str, str]) -> str | None:
    """
    Resolve an imported module name to an installed distribution name.
    Returns None if it cannot be resolved.
    """
    if import_name in import_to_dist:
        return import_to_dist[import_name]

    # Fallback: sometimes import name == distribution name
    try:
        metadata.version(import_name)
        return import_name
    except metadata.PackageNotFoundError:
        return None


def generate_requirements(project_root: Path) -> list[str]:
    stdlib_modules = get_stdlib_modules()
    local_modules = discover_local_modules(project_root)
    import_to_dist = build_import_to_dist_map()

    discovered_imports: Set[str] = set()

    for py_file in find_python_files(project_root):
        discovered_imports.update(extract_imports_from_file(py_file))

    filtered_imports = {
        name
        for name in discovered_imports
        if name
        and name not in stdlib_modules
        and name not in local_modules
        and not name.startswith("_")
    }

    resolved_packages: list[tuple[str, str]] = []
    unresolved_imports: list[str] = []

    for import_name in sorted(filtered_imports):
        dist_name = resolve_distribution_name(import_name, import_to_dist)
        if not dist_name:
            unresolved_imports.append(import_name)
            continue

        try:
            version = metadata.version(dist_name)
            resolved_packages.append((dist_name, version))
        except metadata.PackageNotFoundError:
            unresolved_imports.append(import_name)

    # Deduplicate by normalized dist name while preserving latest found entry
    deduped: dict[str, str] = {}
    for dist_name, version in resolved_packages:
        deduped[dist_name] = version

    requirements = [f"{dist}=={version}" for dist, version in sorted(deduped.items(), key=lambda x: x[0].lower())]

    if unresolved_imports:
        print("\nCould not resolve these imports to installed pip packages:")
        for name in unresolved_imports:
            print(f"  - {name}")
        print("You may need to add some of them manually.\n")

    return requirements


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate requirements.txt from project imports.")
    parser.add_argument(
        "project_root",
        nargs="?",
        default=".",
        help="Path to the Python project root (default: current directory)",
    )
    parser.add_argument(
        "--output",
        default="requirements.txt",
        help="Output requirements file path (default: requirements.txt)",
    )
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    output_file = Path(args.output).resolve()

    if not project_root.exists() or not project_root.is_dir():
        print(f"Error: Project root does not exist or is not a directory: {project_root}")
        return 1

    requirements = generate_requirements(project_root)

    output_file.write_text("\n".join(requirements) + ("\n" if requirements else ""), encoding="utf-8")

    print(f"Scanned project: {project_root}")
    print(f"Wrote {len(requirements)} packages to: {output_file}")

    if requirements:
        print("\nGenerated requirements:")
        for line in requirements:
            print(f"  {line}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())