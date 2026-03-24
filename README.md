# Liquimech Project Management Suite

Liquimech ERP is a desktop application built with Tkinter. The project uses Excel workbooks for stored data and includes reporting, mail, scheduling, products, projects, parts, and job card management screens.

## Project setup

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

If you need to regenerate `requirements.txt` from the current project imports, run:

```bash
python generate_requirements.py
```

This script scans the Python files in the project, resolves installed package names where possible, and rewrites `requirements.txt` with pinned versions. It is useful when dependencies have changed and you want to refresh the tracked requirements file.

Run the app:

```bash
python main.py
```

## PostgreSQL migration

This project now supports running against PostgreSQL for development.

The current migration flow has two parts:

- structured app data is migrated from the Excel workbook into PostgreSQL tables
- document attachments are migrated from local files into PostgreSQL blob storage

### 1. Prepare the environment

Copy the example environment file:

```bash
cp .env.example .env
```

Fill in the PostgreSQL values in `.env`:

```bash
ERP_DATA_BACKEND=postgres
PGHOST=your-digitalocean-host
PGPORT=25060
PGDATABASE=defaultdb
PGUSER=your-db-user
PGPASSWORD=your-db-password
PGSSLMODE=verify-full
PGSSLROOTCERT=certs/ca-certificate.crt
PGCONNECT_TIMEOUT=10
ERP_WORKBOOK_PATH=workspace/data/YourWorkbook.xlsx
```

Important:

- `PGSSLROOTCERT` should point to the CA certificate downloaded from DigitalOcean
- `ERP_DATA_BACKEND=postgres` makes the desktop app use PostgreSQL instead of Excel
- `ERP_WORKBOOK_PATH` should point to the workbook you want to import

### 2. Verify the secure database connection

Run:

```bash
python scripts/test_postgres_connection.py
```

Expected result:

- TLS verification succeeds
- the script prints the connected database, user, and server version

If this fails with a timeout:

- confirm the database is reachable from your machine
- for development, allow your public IP in DigitalOcean trusted sources
- for production, prefer a private database plus an API layer instead of direct desktop-to-database access

### 3. Apply the PostgreSQL schema

Run:

```bash
python scripts/migrate_workbook_to_postgres.py --schema-only
```

This creates the PostgreSQL tables used by the app, including:

- `parts`
- `modules`
- `tasks`
- `components`
- `documents`
- `products`
- `product_modules`
- `product_documents`
- `projects`
- `project_modules`
- `project_tasks`
- `project_documents`
- `workorders`
- `completed_jobs`
- `completed_job_lines`
- `document_blobs`

### 4. Run a workbook schema preflight

Before importing, validate the workbook structure:

```bash
python scripts/validate_workbook_schema.py
```

Or run the same validation through the importer:

```bash
python scripts/migrate_workbook_to_postgres.py --preflight-only
```

What the preflight checks:

- missing expected sheets
- extra workbook sheets
- missing expected headers
- extra headers
- reordered headers
- duplicate headers

How it behaves:

- missing or renamed required headers are treated as blocking issues
- duplicate headers are treated as blocking issues
- missing expected sheets are reported, but older workbooks can still import if those sheets are not present
- extra sheets and extra columns are reported, but ignored during import
- reordered columns are reported, but import still works because columns are matched by header name

### 5. Migrate workbook data into PostgreSQL

Run:

```bash
python scripts/migrate_workbook_to_postgres.py --truncate-first
```

This imports the workbook sheets into PostgreSQL.

Current sheet-to-table mapping:

- `Parts` -> `parts`
- `Modules` -> `modules`
- `Tasks` -> `tasks`
- `Components` -> `components`
- `Documents` -> `documents`
- `Products` -> `products`
- `ProductModules` -> `product_modules`
- `ProductDocuments` -> `product_documents`
- `Projects` -> `projects`
- `ProjectModules` -> `project_modules`
- `ProjectTasks` -> `project_tasks`
- `ProjectDocuments` -> `project_documents`
- `WorkOrders` -> `workorders`
- `CompletedJobs` -> `completed_jobs`
- `CompletedJobLines` -> `completed_job_lines`

Notes:

- the importer skips workbook sheets that do not exist yet in older files
- the importer also repairs known legacy `Components` row layouts during import
- the importer matches columns by header name, so reordered columns and extra non-app columns do not break import
- `--truncate-first` clears the PostgreSQL tables before re-importing, which is useful during development

### 6. Migrate document attachments into PostgreSQL

Workbook row data and document file contents are migrated separately.

After the workbook rows are imported, run:

```bash
python scripts/migrate_document_attachments_to_postgres.py
```

This scans these document record tables:

- `documents`
- `product_documents`
- `project_documents`

For each document row, it:

- reads the local file from the existing `FilePath`
- uploads the bytes into the `document_blobs` table
- updates the row `FilePath` to a database URI such as `db://Documents/<id>/filename`

If a local file is missing, the script skips it and reports that path.

### 7. Run the app against PostgreSQL

Once the schema, workbook rows, and attachments have been migrated, start the app:

```bash
python main.py
```

When `.env` contains `ERP_DATA_BACKEND=postgres`, the app should:

- show `[PostgreSQL]` in the window title
- show a PostgreSQL connection label on the home page
- disable the `Create Workbook` and `Open Workbook` buttons

### 8. Verify record types and attachments

Recommended checks after migration:

1. Open assemblies and confirm modules, tasks, components, and module documents load.
2. Open products and confirm product modules, product documents, and work orders load.
3. Open live orders and confirm project modules, project tasks, project documents, and completed jobs load.
4. Open a migrated document attachment and confirm it opens from PostgreSQL-backed temp storage.
5. Add a new document in PostgreSQL mode and confirm it can be opened again after refresh.

### Scripts involved in the migration

- `scripts/test_postgres_connection.py`
  Verifies the secure PostgreSQL connection using the CA certificate.
- `scripts/migrate_workbook_to_postgres.py`
  Creates the schema, validates workbook structure, and imports workbook row data into PostgreSQL.
- `scripts/validate_workbook_schema.py`
  Runs the workbook preflight report without changing the database.
- `scripts/migrate_document_attachments_to_postgres.py`
  Uploads existing local attachment files into PostgreSQL blob storage.

### Current scope

This PostgreSQL support is suitable for development and internal testing.

At the moment:

- structured records can be stored in PostgreSQL
- document attachment bytes can also be stored in PostgreSQL
- the desktop app can run directly against PostgreSQL

For a production rollout, a backend API is still recommended so the desktop app does not connect to the database directly.

## Main packaging files

- `Liquimech ERP.spec`: PyInstaller build configuration
- `installer.iss`: Inno Setup installer script for Windows

## Build the app with PyInstaller

This project is currently configured as a `--onedir` build. That means the distributable is the whole output folder, not just a single executable.

Build using the spec file:

```bash
pyinstaller "Liquimech ERP.spec"
```

If this command does not work, instead of calling pyinstaller directly, use the Python module mechanism, which is more reliable:
```bash
python -m PyInstaller '.\Liquimech ERP.spec'
```

Or build directly from the entry point.

On macOS or Linux:

```bash
pyinstaller --noconfirm --windowed --onedir --name "Liquimech ERP" \
  --add-data "assets:assets" \
  --add-data "workspace/data:workspace/data" \
  main.py
```

On Windows:

```bat
pyinstaller --noconfirm --windowed --onedir --name "Liquimech ERP" ^
  --add-data "assets;assets" ^
  --add-data "workspace/data;workspace/data" ^
  main.py
```

Build output is written to `dist/`.

## Important packaging note

PyInstaller does not cross-compile. A Windows `.exe` must be built on Windows. A macOS build produces a macOS app bundle, not a Windows executable.

That means:

- build macOS app artifacts on macOS
- build Windows app artifacts on Windows

## Windows installer

The Windows installer is defined in `installer.iss` and expects a Windows PyInstaller build output at:

```text
dist\Liquimech ERP\Liquimech ERP.exe
```

The installer is configured to install to:

```text
%LOCALAPPDATA%\Liquimech\LiquimechERP
```

This is intentional. The app creates and writes folders at runtime, so installing to a user-writable location is safer than `Program Files`.

### Build the Windows installer on Windows

1. Build the Windows app with PyInstaller.
2. Confirm this file exists:

```text
dist\Liquimech ERP\Liquimech ERP.exe
```

3. Open `installer.iss` in Inno Setup and compile it, or run:

```bat
ISCC installer.iss
```

The generated installer will be written to the folder configured by `OutputDir`, currently:

```text
installer\
```

## Windows auto-update with GitHub Releases

This project can use the built-in Python updater to check the latest GitHub Release, compare versions, download the latest installer, and launch it for the user.

Setup requirements:

- Set `GITHUB_RELEASE_OWNER` in `app_config.py`
- Set `GITHUB_RELEASE_REPO` in `app_config.py`
- Set `GITHUB_RELEASE_ASSET_NAME` in `app_config.py`
- Optionally adjust `GITHUB_RELEASE_CHECK_INTERVAL_SECONDS` in `app_config.py`
- Upload the Windows installer asset to each GitHub Release using a consistent filename

The updater checks the latest release using GitHub's latest-release API and looks for the configured installer asset.

Important:

- release tags should use a version-like format such as `v1.0.1`
- the uploaded installer asset name must match `GITHUB_RELEASE_ASSET_NAME`
- in-app installer launch is currently supported on Windows only
- automatic startup checks are rate-limited and currently default to once per day
- users can skip a specific release, and automatic checks will stay quiet until a newer version is published

## Building the Windows installer from macOS

You can compile the Inno Setup script on macOS using Docker, based on the approach described here:

https://gist.github.com/amake/3e7194e5e61d0e1850bba144797fd797

Run:

```bash
docker run --rm -i -v "$PWD:/work" amake/innosetup installer.iss
```

Important:

- this only compiles the Inno Setup installer
- it does not create a Windows `.exe` build of the app
- you still need Windows-built PyInstaller output in `dist/` before compiling the installer

In other words, macOS plus Docker can package the Windows installer, but it cannot replace the Windows PyInstaller build step.

## Portable distribution

If you do not want to use an installer, you can distribute the PyInstaller `--onedir` output folder directly.

For a Windows build, distribute the whole folder:

```text
dist\Liquimech ERP\
```

Do not send only the `.exe` from a `--onedir` build, because it depends on the other bundled files beside it.

## Release checklist

Use this checklist when preparing a new release:

1. Update the app version in the code and installer files if needed.
2. Activate the project virtual environment.
3. Run the app locally and confirm the main workflows still open correctly.
4. If dependencies changed, run `python generate_requirements.py` and review the updated `requirements.txt`.
5. Build the distributable with PyInstaller.
6. If releasing for Windows, make sure the build was created on Windows and confirm this file exists:

```text
dist\Liquimech ERP\Liquimech ERP.exe
```

7. If releasing a Windows installer, compile `installer.iss`.
8. Test the packaged build before sharing it:
   - open the app
   - confirm the logo and bundled assets load
   - open or create a workbook
   - verify exports or temporary files can be written
9. If this release should be available through the in-app updater, upload the installer to a GitHub Release and make sure the release tag and installer filename match the updater configuration in `app_config.py`.
10. For installer releases, test the installed app by launching it from the installer and from the desktop shortcut.
11. Clean up generated artifacts you do not want to commit, then commit only source and packaging definition files.
12. Distribute either:
   - the full `dist/Liquimech ERP/` folder for portable distribution
   - or the generated installer from `installer/` for Windows installer distribution

## Repository guidance

Recommended to track:

- application source files
- `requirements.txt`
- `Liquimech ERP.spec`
- `installer.iss`
- shipped assets required by the app

Recommended to ignore:

- `.venv/`
- `__pycache__/`
- `build/`
- `dist/`
- `installer/`
- `.DS_Store`
- generated runtime files under `workspace/exports/`, `workspace/temp/`, and document subfolders

## Current status

This README documents the packaging flow currently used by the project. If the app is later changed to store writable data in a dedicated user data directory, the installer strategy can be simplified further.
