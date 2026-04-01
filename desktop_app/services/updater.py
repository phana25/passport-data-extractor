import json
import os
import stat
import ssl
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path
from packaging.version import parse as parse_version

from PySide6.QtCore import QObject, QThread, Signal

# Real URL pointing to your repository's version.json file on the main branch
# Format expected: {"latest_version": "1.1.0", "windows_url": "...", "mac_url": "...", "release_notes": "..."}
# GitHub raw URL for the main branch
UPDATE_JSON_URL = "https://raw.githubusercontent.com/phana25/passport-data-extractor/main/version.json"
# Backup URL (e.g., a Public Gist) to use if the main repository is set to private
FALLBACK_UPDATE_URL = "https://gist.githubusercontent.com/phana25/14514d5de8ded8943899638598638b93/raw/version.json" 

class CheckUpdateWorker(QObject):
    update_available = Signal(str, str, str)  # version, download_url, release_notes
    error = Signal(str)
    no_update = Signal()

    def __init__(self, current_version: str):
        super().__init__()
        self.current_version = current_version

    def _fetch_data(self, url: str) -> dict | None:
        """Helper to fetch and parse JSON from a URL."""
        if not url:
            return None
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=5) as response:
                return json.loads(response.read().decode('utf-8'))
        except Exception:
            return None

    def run(self):
        # 1. Try primary URL
        data = self._fetch_data(UPDATE_JSON_URL)
        
        # 2. Try fallback if primary fails
        if data is None and FALLBACK_UPDATE_URL:
            data = self._fetch_data(FALLBACK_UPDATE_URL)
            
        if data is None:
            self.error.emit("Update check failed: Unable to reach any update server.")
            return

        try:
            latest_ver = data.get("latest_version")
            if not latest_ver:
                self.error.emit("Invalid update format in version.json")
                return

            if parse_version(latest_ver) > parse_version(self.current_version):
                # Determine appropriate OS url
                if sys.platform == "win32":
                    download_url = data.get("windows_url") or data.get("download_url")
                else:
                    download_url = data.get("mac_url") or data.get("download_url")
                    
                if download_url:
                    self.update_available.emit(latest_ver, download_url, data.get("release_notes", ""))
                else:
                    self.error.emit("No compatible download found for this OS")
            else:
                self.no_update.emit()
                    
        except Exception as e:
            self.error.emit(f"Check update processing error: {e}")



class DownloadUpdateWorker(QObject):
    progress = Signal(int)
    finished = Signal(str)  # returns the path to the downloaded file
    error = Signal(str)

    def __init__(self, download_url: str):
        super().__init__()
        self.download_url = download_url
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            temp_dir = Path(tempfile.gettempdir()) / "passport_verifier_update"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            # Determine filename from URL
            filename = self.download_url.split('/')[-1]
            if not filename or "?" in filename:
                filename = "update_package.zip" if sys.platform != "win32" else "update_package.exe"
                
            dest_path = temp_dir / filename

            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

            req = urllib.request.Request(self.download_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=10) as response:
                total_size = int(response.info().get("Content-Length", 0))
                downloaded_size = 0
                chunk_size = 8192

                with open(dest_path, "wb") as f:
                    while True:
                        if self._is_cancelled:
                            self.error.emit("Download cancelled")
                            return
                            
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        
                        if total_size > 0:
                            percent = int((downloaded_size / total_size) * 100)
                            self.progress.emit(percent)

                self.finished.emit(str(dest_path))

        except Exception as e:
            self.error.emit(str(e))


class UpdaterService:
    @staticmethod
    def install_and_restart(downloaded_file: str):
        """
        Generates and runs an OS-specific script to replace the current app and restart it.
        This must be called right before QApplication.quit()
        """
        current_executable = sys.executable
        # Note: If running from source, sys.executable is the Python interpreter.
        # A more robust detection for packaged apps:
        if getattr(sys, 'frozen', False):
            current_executable = sys.executable
            is_packaged = True
        else:
            # If we are in dev mode, we usually don't want to overwrite the dev scripts
            # like this, but we'll try to find the root folder for demonstration.
            is_packaged = False
            current_executable = os.path.abspath(sys.argv[0])
            
        temp_dir = Path(tempfile.gettempdir()) / "passport_verifier_scripts"
        temp_dir.mkdir(parents=True, exist_ok=True)

        if sys.platform == "win32":
            script_path = temp_dir / "update.bat"
            # Inno Setup flags for silent background update:
            # /SILENT /VERYSILENT: Install without showing wizard
            # /SUPPRESSMSGBOXES: Don't show any error/info messages
            # Note: the installer will auto-close the app and restart it based on AppId.
            bat_content = f"""@echo off
timeout /t 2 /nobreak > NUL
start "" "{downloaded_file}" /SILENT /VERYSILENT /SUPPRESSMSGBOXES
del "%~f0"
"""
            with open(script_path, "w") as f:
                f.write(bat_content)
                
            subprocess.Popen(
                [str(script_path)],
                creationflags=subprocess.CREATE_NEW_CONSOLE | subprocess.CREATE_NO_WINDOW
            )

            
        else:
            # macOS / Linux bash script
            # Assuming downloaded_file is a zip containing the .app bundle
            script_path = temp_dir / "update.sh"
            
            app_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_executable)))
            if not app_dir.endswith(".app"):
                # If we're not inside a .app bundle (e.g. running raw python scripts)
                app_dir = os.path.dirname(current_executable)
                
            sh_content = f"""#!/bin/bash
sleep 2

# If it's a zip we must extract it
if [[ "{downloaded_file}" == *.zip ]]; then
    unzip -o -q "{downloaded_file}" -d "{Path(downloaded_file).parent}"
    # Find the extracted .app
    EXTRACTED_APP=$(find "{Path(downloaded_file).parent}" -maxdepth 1 -name "*.app" | head -n 1)
    if [ ! -z "$EXTRACTED_APP" ]; then
        rm -rf "{app_dir}"
        mv "$EXTRACTED_APP" "{app_dir}"
        open -a "{app_dir}"
    fi
else
    # Simple replacement if it's a binary/file
    mv -f "{downloaded_file}" "{current_executable}"
    open "{current_executable}"
fi
rm "$0"
"""
            with open(script_path, "w") as f:
                f.write(sh_content)
            
            # Make the script executable
            os.chmod(script_path, os.stat(script_path).st_mode | stat.S_IEXEC)
            
            subprocess.Popen([str(script_path)], start_new_session=True)
