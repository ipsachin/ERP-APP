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
9. For installer releases, test the installed app by launching it from the installer and from the desktop shortcut.
10. Clean up generated artifacts you do not want to commit, then commit only source and packaging definition files.
11. Distribute either:
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
