"""
VIPQC AI — Auto-Update Module
Checks for latest releases on GitHub, downloads updates with progress reporting,
and performs safe in-place executable replacement on Windows.
"""

import os
import sys
import json
import re
import time
import subprocess
import urllib.request
import urllib.error
import ssl

CURRENT_VERSION = "2.2.7"
GITHUB_OWNER = "levanvung"
GITHUB_REPO = "VIPQC"


def _get_ssl_context():
    """
    Returns an SSL context that bypasses corporate proxy or missing root CA errors.
    Prevents [SSL: CERTIFICATE_VERIFY_FAILED] on factory / corporate PCs.
    """
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx
    except Exception:
        pass
    try:
        return ssl._create_unverified_context()
    except Exception:
        return None

# Primary raw descriptor (No GitHub API rate limit)
GITHUB_RAW_VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}/master/version.json"
# Secondary GitHub Releases API
GITHUB_RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
# Default direct download fallback
DEFAULT_EXE_DOWNLOAD_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/raw/master/VIPQC%20AI.exe"
GITHUB_REPO_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}"


def parse_version_tuple(v_str: str) -> tuple[int, ...]:
    """
    Parses a version string like 'v2.1.0' or '2.2' into a comparable tuple of integers (2, 1, 0).
    """
    if not v_str:
        return (0, 0, 0)
    cleaned = re.sub(r'^[vV]', '', str(v_str).strip())
    parts = re.findall(r'\d+', cleaned)
    if not parts:
        return (0, 0, 0)
    return tuple(int(p) for p in parts)


def is_newer_version(current_ver: str, remote_ver: str) -> bool:
    """
    Returns True if remote_ver is strictly greater than current_ver.
    Example: (2, 2, 0) > (2, 1, 0) -> True
    """
    curr = parse_version_tuple(current_ver)
    remote = parse_version_tuple(remote_ver)
    # Pad to equal length
    max_len = max(len(curr), len(remote))
    curr_padded = curr + (0,) * (max_len - len(curr))
    remote_padded = remote + (0,) * (max_len - len(remote))
    return remote_padded > curr_padded


def check_for_updates(timeout: int = 5) -> dict:
    """
    Checks for available updates from GitHub.
    Returns a dictionary:
      {
        "has_update": bool,
        "current_version": str,
        "latest_version": str,
        "title": str,
        "changelog": str,
        "download_url": str,
        "release_url": str,
        "error": str or None
      }
    """
    result = {
        "has_update": False,
        "current_version": CURRENT_VERSION,
        "latest_version": CURRENT_VERSION,
        "title": "",
        "changelog": "",
        "download_url": DEFAULT_EXE_DOWNLOAD_URL,
        "release_url": GITHUB_REPO_URL,
        "error": None
    }

    # 1. Try raw version.json first (fastest, no rate limits)
    headers = {"User-Agent": f"VIPQC-AI-Updater/{CURRENT_VERSION}"}
    raw_data = None
    ssl_ctx = _get_ssl_context()
    try:
        req = urllib.request.Request(GITHUB_RAW_VERSION_URL, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx) as resp:
            if resp.status == 200:
                raw_data = json.loads(resp.read().decode("utf-8"))
    except Exception as ex_raw:
        # If raw fails, attempt GitHub Releases API
        pass

    if raw_data and isinstance(raw_data, dict) and "version" in raw_data:
        rem_ver = str(raw_data.get("version", "")).strip()
        result["latest_version"] = rem_ver
        result["title"] = raw_data.get("title", f"VIPQC AI v{rem_ver}")
        result["changelog"] = raw_data.get("changelog", "Bản cập nhật cải tiến hiệu năng và tính năng mới.")
        result["download_url"] = raw_data.get("download_url", DEFAULT_EXE_DOWNLOAD_URL)
        result["release_url"] = raw_data.get("release_url", GITHUB_REPO_URL)
        result["has_update"] = is_newer_version(CURRENT_VERSION, rem_ver)
        return result

    # 2. Fallback: Try GitHub Releases API
    try:
        req = urllib.request.Request(GITHUB_RELEASES_API_URL, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx) as resp:
            if resp.status == 200:
                api_data = json.loads(resp.read().decode("utf-8"))
                tag_ver = api_data.get("tag_name", "")
                result["latest_version"] = re.sub(r'^[vV]', '', tag_ver)
                result["title"] = api_data.get("name") or f"VIPQC AI v{result['latest_version']}"
                result["changelog"] = api_data.get("body", "Xem chi tiết bản phát hành trên GitHub.")
                result["release_url"] = api_data.get("html_url", GITHUB_REPO_URL)
                
                # Check assets for .exe
                assets = api_data.get("assets", [])
                for a in assets:
                    if a.get("name", "").lower().endswith(".exe"):
                        result["download_url"] = a.get("browser_download_url", DEFAULT_EXE_DOWNLOAD_URL)
                        break

                result["has_update"] = is_newer_version(CURRENT_VERSION, result["latest_version"])
                return result
    except Exception as ex_api:
        result["error"] = str(ex_api)

    return result


def download_update_file(download_url: str, target_file: str, progress_callback=None, cancel_event=None) -> bool:
    """
    Downloads the update file in chunks and calls progress_callback(pct, downloaded_bytes, total_bytes).
    Uses SSL bypass context to avoid corporate network [SSL: CERTIFICATE_VERIFY_FAILED] issues.
    """
    headers = {"User-Agent": f"VIPQC-AI-Updater/{CURRENT_VERSION}"}
    req = urllib.request.Request(download_url, headers=headers)
    ssl_ctx = _get_ssl_context()

    os.makedirs(os.path.dirname(os.path.abspath(target_file)), exist_ok=True)

    with urllib.request.urlopen(req, timeout=60, context=ssl_ctx) as resp:
        total_size = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        chunk_size = 64 * 1024  # 64 KB

        with open(target_file, "wb") as f_out:
            while True:
                if cancel_event and cancel_event.is_set():
                    return False
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f_out.write(chunk)
                downloaded += len(chunk)
                if progress_callback:
                    pct = (downloaded / total_size * 100) if total_size > 0 else 0
                    progress_callback(pct, downloaded, total_size)

    return os.path.exists(target_file) and os.path.getsize(target_file) > 1024 * 100


def get_target_exe_path() -> str:
    """
    Returns the absolute path to the current running executable (or target VIPQC AI.exe in cwd).
    """
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    # When running from Python source, target the exe in directory
    exe_candidate = os.path.abspath("VIPQC AI.exe")
    if os.path.exists(exe_candidate):
        return exe_candidate
    return os.path.abspath(sys.executable)


def apply_update_and_restart(new_exe_path: str):
    """
    Spawns a detached Windows background process to:
      1. Wait for current VIPQC process to exit and release file locks.
      2. Overwrite the target executable with the newly downloaded executable.
      3. Launch the new executable.
    Runs completely silent without opening any black CMD windows.
    Then immediately exits the current process.
    """
    target_exe = get_target_exe_path()
    updater_dir = os.path.dirname(target_exe)
    alt_exe = os.path.join(updater_dir, "BOM_Extractor.exe")

    # Clean up old temporary bat if exists
    old_bat = os.path.join(updater_dir, "_update_launcher.bat")
    if os.path.exists(old_bat):
        try:
            os.remove(old_bat)
        except Exception:
            pass

    CREATE_NO_WINDOW = 0x08000000
    DETACHED_PROCESS = 0x00000008
    flags = CREATE_NO_WINDOW | DETACHED_PROCESS

    # PowerShell in-place updater (Silent, no CMD window, retries until file unlocked)
    ps_script = (
        f"$target = '{target_exe}'; "
        f"$new = '{new_exe_path}'; "
        f"$alt = '{alt_exe}'; "
        f"Start-Sleep -Milliseconds 800; "
        f"for ($i = 0; $i -lt 30; $i++) {{ "
        f"  try {{ "
        f"    Copy-Item -LiteralPath $new -Destination $target -Force -ErrorAction Stop; "
        f"    Remove-Item -LiteralPath $new -Force -ErrorAction SilentlyContinue; "
        f"    if (Test-Path -LiteralPath $alt) {{ Copy-Item -LiteralPath $target -Destination $alt -Force -ErrorAction SilentlyContinue }}; "
        f"    Start-Process -FilePath $target; "
        f"    break; "
        f"  }} catch {{ "
        f"    Start-Sleep -Milliseconds 800; "
        f"  }} "
        f"}}"
    )

    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
            creationflags=flags
        )
    except Exception:
        # Fallback to simple silent batch file
        bat_content = f"""@echo off
timeout /t 2 /nobreak > nul
:retry_copy
copy /y "{new_exe_path}" "{target_exe}" > nul
if errorlevel 1 (
    timeout /t 1 /nobreak > nul
    goto retry_copy
)
if exist "{new_exe_path}" del /f /q "{new_exe_path}" > nul
if exist "{alt_exe}" copy /y "{target_exe}" "{alt_exe}" > nul
start "" "{target_exe}"
del /f /q "%~f0"
"""
        try:
            with open(old_bat, "w", encoding="utf-8", errors="ignore") as f:
                f.write(bat_content)
            subprocess.Popen(["cmd.exe", "/c", old_bat], creationflags=flags)
        except Exception:
            pass

    # Hard exit current process so target_exe is unlocked immediately
    os._exit(0)
