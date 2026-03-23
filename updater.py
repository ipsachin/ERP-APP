from __future__ import annotations

import json
import os
import re
import subprocess
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import tkinter as tk
from tkinter import messagebox, ttk

from app_config import AppConfig


def _normalize_version(raw: str) -> str:
    text = str(raw or "").strip()
    if text.lower().startswith("v"):
        text = text[1:]
    return text


def _version_key(raw: str) -> tuple[int, ...]:
    normalized = _normalize_version(raw)
    parts = re.findall(r"\d+", normalized)
    if not parts:
        return (0,)
    return tuple(int(part) for part in parts)


def _is_newer_version(remote_version: str, local_version: str) -> bool:
    return _version_key(remote_version) > _version_key(local_version)


@dataclass
class ReleaseAsset:
    name: str
    download_url: str


@dataclass
class ReleaseInfo:
    version: str
    tag_name: str
    name: str
    body: str
    html_url: str
    asset: ReleaseAsset


class AppUpdater:
    def __init__(self, app):
        self.app = app
        self.root = app.root
        self._check_in_progress = False
        self._download_in_progress = False
        self._state_cache: dict[str, object] | None = None

    @property
    def api_url(self) -> str:
        owner = AppConfig.GITHUB_RELEASE_OWNER.strip()
        repo = AppConfig.GITHUB_RELEASE_REPO.strip()
        return f"https://api.github.com/repos/{owner}/{repo}/releases/latest"

    @property
    def state_path(self) -> Path:
        return AppConfig.TEMP_DIR / "updater_state.json"

    def is_configured(self) -> bool:
        return bool(
            AppConfig.GITHUB_RELEASE_OWNER.strip()
            and AppConfig.GITHUB_RELEASE_REPO.strip()
            and AppConfig.GITHUB_RELEASE_ASSET_NAME.strip()
        )

    def check_for_updates(self, manual: bool = False) -> None:
        if self._check_in_progress:
            if manual:
                messagebox.showinfo("Update Check", "An update check is already running.")
            return

        if not self.is_configured():
            if manual:
                messagebox.showwarning(
                    "Updates Not Configured",
                    "Update settings are not configured.",
                )
            return

        if not manual and not self._should_run_auto_check():
            return

        self._record_check_attempt()
        self._check_in_progress = True
        self.app.set_status("Checking for updates...")
        thread = threading.Thread(
            target=self._check_for_updates_worker,
            args=(manual,),
            daemon=True,
        )
        thread.start()

    def _check_for_updates_worker(self, manual: bool) -> None:
        try:
            release = self._fetch_latest_release()
            self.root.after(0, lambda release=release, manual=manual: self._handle_release_info(release, manual))
        except Exception as exc:
            self.root.after(0, lambda exc=exc, manual=manual: self._handle_update_error(exc, manual))

    def _fetch_latest_release(self) -> ReleaseInfo:
        request = Request(
            self.api_url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"LiquimechERP/{AppConfig.APP_VERSION}",
            },
        )
        with urlopen(request, timeout=15) as response:
            payload = json.loads(response.read().decode("utf-8"))

        asset_name = AppConfig.GITHUB_RELEASE_ASSET_NAME.strip()
        for asset in payload.get("assets", []):
            if asset.get("name") == asset_name and asset.get("browser_download_url"):
                return ReleaseInfo(
                    version=_normalize_version(payload.get("tag_name") or payload.get("name") or ""),
                    tag_name=str(payload.get("tag_name") or ""),
                    name=str(payload.get("name") or ""),
                    body=str(payload.get("body") or ""),
                    html_url=str(payload.get("html_url") or ""),
                    asset=ReleaseAsset(
                        name=str(asset["name"]),
                        download_url=str(asset["browser_download_url"]),
                    ),
                )

        raise RuntimeError(
            "The latest update package could not be found."
        )

    def _handle_release_info(self, release: ReleaseInfo, manual: bool) -> None:
        self._check_in_progress = False

        remote_version = release.version or _normalize_version(release.tag_name)
        local_version = AppConfig.APP_VERSION

        if not remote_version:
            self.app.set_status("Update check failed: latest release version was empty.")
            if manual:
                messagebox.showerror(
                    "Update Check Failed",
                    "The latest available update did not expose a usable version.",
                )
            return

        if not _is_newer_version(remote_version, local_version):
            self.app.set_status("You are already on the latest version.")
            self._clear_skipped_version(remote_version)
            if manual:
                messagebox.showinfo(
                    "No Update Available",
                    f"You are already on version {local_version}.",
                )
            return

        if not manual and self._is_skipped_version(remote_version):
            self.app.set_status(f"Update {remote_version} was previously skipped.")
            return

        self.app.set_status(f"Update available: {remote_version}")
        if not self._can_install_updates():
            messagebox.showinfo(
                "Update Available",
                (
                    f"Version {remote_version} is available, but in-app installation "
                    "is currently supported on Windows only."
                ),
            )
            return

        summary = (
            f"A new version is available.\n\n"
            f"Current version: {local_version}\n"
            f"Latest version: {remote_version}\n\n"
            "Choose an action:\n"
            "- Install Now downloads and starts the installer\n"
            "- Skip This Version hides this release from automatic checks\n"
            "- Later reminds you again on a future check"
        )
        if release.body.strip():
            summary += "\n\nRelease notes:\n" + release.body.strip()[:800]

        user_choice = self._prompt_for_update_action(remote_version, summary)
        if user_choice == "install":
            self._clear_skipped_version(remote_version)
            self._download_and_install_release(release)
        elif user_choice == "skip":
            self._set_skipped_version(remote_version)
            self.app.set_status(f"Skipped update {remote_version}.")
            if manual:
                messagebox.showinfo(
                    "Update Skipped",
                    f"Version {remote_version} will be skipped until a newer version is published.",
                )

    def _handle_update_error(self, exc: Exception, manual: bool) -> None:
        self._check_in_progress = False
        message = self._friendly_error_message(exc)
        self.app.set_status(f"Update check failed: {message}")
        if manual:
            messagebox.showerror("Update Check Failed", message)

    def _friendly_error_message(self, exc: Exception) -> str:
        if isinstance(exc, HTTPError):
            return f"The update service responded with HTTP {exc.code}."
        if isinstance(exc, URLError):
            return "Unable to reach the update service. Check the internet connection and try again."
        return str(exc) or exc.__class__.__name__

    def _can_install_updates(self) -> bool:
        return os.name == "nt"

    def _download_and_install_release(self, release: ReleaseInfo) -> None:
        if self._download_in_progress:
            messagebox.showinfo("Updater Busy", "An update download is already in progress.")
            return

        self._download_in_progress = True
        self.app.set_status(f"Downloading update {release.version}...")
        thread = threading.Thread(
            target=self._download_and_install_worker,
            args=(release,),
            daemon=True,
        )
        thread.start()

    def _download_and_install_worker(self, release: ReleaseInfo) -> None:
        try:
            installer_path = self._download_installer(release.asset)
            self.root.after(
                0,
                lambda installer_path=installer_path, release=release: self._finish_install(installer_path, release),
            )
        except Exception as exc:
            self.root.after(0, lambda exc=exc: self._download_failed(exc))

    def _download_installer(self, asset: ReleaseAsset) -> Path:
        target_dir = Path(tempfile.gettempdir()) / "liquimech_erp_updates"
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / asset.name

        request = Request(
            asset.download_url,
            headers={"User-Agent": f"LiquimechERP/{AppConfig.APP_VERSION}"},
        )
        with urlopen(request, timeout=60) as response, target_path.open("wb") as handle:
            while True:
                chunk = response.read(1024 * 128)
                if not chunk:
                    break
                handle.write(chunk)

        return target_path

    def _download_failed(self, exc: Exception) -> None:
        self._download_in_progress = False
        message = self._friendly_error_message(exc)
        self.app.set_status(f"Update download failed: {message}")
        messagebox.showerror("Update Download Failed", message)

    def _finish_install(self, installer_path: Path, release: ReleaseInfo) -> None:
        self._download_in_progress = False

        if not installer_path.exists():
            self.app.set_status("Update download failed: installer file was missing.")
            messagebox.showerror(
                "Update Install Failed",
                "The installer was downloaded, but the file could not be found.",
            )
            return

        self.app.set_status(f"Launching installer for version {release.version}...")
        try:
            subprocess.Popen(
                [
                    str(installer_path),
                    "/SP-",
                    "/SILENT",
                    "/CLOSEAPPLICATIONS",
                    "/NOICONS",
                ],
                close_fds=False,
            )
        except Exception as exc:
            self.app.set_status(f"Update install failed: {exc}")
            messagebox.showerror("Update Install Failed", str(exc))
            return

        messagebox.showinfo(
            "Installer Started",
            "The update installer has been started. The application will now close.",
        )
        self.root.after(250, self.root.destroy)

    def _load_state(self) -> dict[str, object]:
        if self._state_cache is not None:
            return dict(self._state_cache)

        try:
            if self.state_path.exists():
                self._state_cache = json.loads(self.state_path.read_text(encoding="utf-8"))
            else:
                self._state_cache = {}
        except Exception:
            self._state_cache = {}
        return dict(self._state_cache)

    def _save_state(self, state: dict[str, object]) -> None:
        AppConfig.ensure_directories()
        self.state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
        self._state_cache = dict(state)

    def _record_check_attempt(self) -> None:
        state = self._load_state()
        state["last_check_ts"] = int(time.time())
        self._save_state(state)

    def _should_run_auto_check(self) -> bool:
        state = self._load_state()
        last_check_ts = int(state.get("last_check_ts") or 0)
        interval = max(3600, int(AppConfig.GITHUB_RELEASE_CHECK_INTERVAL_SECONDS))
        return (int(time.time()) - last_check_ts) >= interval

    def _is_skipped_version(self, version: str) -> bool:
        state = self._load_state()
        return _normalize_version(str(state.get("skipped_version") or "")) == _normalize_version(version)

    def _set_skipped_version(self, version: str) -> None:
        state = self._load_state()
        state["skipped_version"] = _normalize_version(version)
        state["skipped_on"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
        self._save_state(state)

    def _clear_skipped_version(self, version: str | None = None) -> None:
        state = self._load_state()
        skipped = _normalize_version(str(state.get("skipped_version") or ""))
        if not skipped:
            return
        if version is not None and skipped != _normalize_version(version):
            return
        state.pop("skipped_version", None)
        state.pop("skipped_on", None)
        self._save_state(state)

    def _prompt_for_update_action(self, version: str, message: str) -> str:
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Available")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.grab_set()

        container = ttk.Frame(dialog, padding=16)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text=f"Version {version} is available",
            style="Title.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        msg = tk.Text(container, width=68, height=16, wrap="word", relief="flat")
        msg.insert("1.0", message)
        msg.configure(state="disabled", background=dialog.cget("background"))
        msg.pack(fill="both", expand=True)

        choice = {"value": "later"}

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(12, 0))

        def close_with(value: str) -> None:
            choice["value"] = value
            dialog.destroy()

        ttk.Button(button_row, text="Install Now", command=lambda: close_with("install")).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Skip This Version", command=lambda: close_with("skip")).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="Later", command=lambda: close_with("later")).pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", lambda: close_with("later"))
        dialog.update_idletasks()
        dialog.wait_window()
        return str(choice["value"])
