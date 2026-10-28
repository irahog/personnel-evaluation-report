"""In-app update from the public releases repository.

Only runs when the user asks (Settings → «بررسی نسخهٔ جدید»). It contacts GitHub alone, sends no
employee data (a plain request for the release list), verifies every download against the release's
SHA256SUMS.txt, and then hands over to the platform installer:

  * Windows: the Inno Setup installer runs silently for the current user and restarts the app.
  * macOS:   a small helper script waits for the app to quit, replaces the .app bundle in place
             with the one from the downloaded disk image, and reopens it.

Downloads live in the per-user app folder (`updates/`) and are deleted after installing.
"""
from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shutil
import ssl
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

from . import VERSION
from .normalize import to_fa_digits
from .paths import user_data_dir

RELEASES_REPO = "irahog/personnel-evaluation-report-releases"
API_LATEST = f"https://api.github.com/repos/{RELEASES_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{RELEASES_REPO}/releases/latest"
ALLOWED_HOSTS = ("github.com", "api.github.com")
ALLOWED_SUFFIXES = (".githubusercontent.com",)
TIMEOUT = 20
CHUNK = 256 * 1024


class UpdateError(Exception):
    """Persian, user-facing."""


def parse_version(text: str) -> tuple[int, ...]:
    m = re.match(r"^v?(\d+(?:\.\d+)*)$", text.strip())
    if not m:
        raise ValueError(text)
    return tuple(int(x) for x in m.group(1).split("."))


def is_newer(latest: str, current: str = VERSION) -> bool:
    a, b = parse_version(latest), parse_version(current)
    n = max(len(a), len(b))
    return a + (0,) * (n - len(a)) > b + (0,) * (n - len(b))


def _check_url(url: str) -> None:
    u = urllib.parse.urlparse(url)
    host = (u.hostname or "").lower()
    if u.scheme != "https" or not (host in ALLOWED_HOSTS or host.endswith(ALLOWED_SUFFIXES)):
        raise UpdateError("نشانی دانلود معتبر نیست؛ به‌روزرسانی انجام نشد.")


_ssl_context: ssl.SSLContext | None = None


def _context() -> ssl.SSLContext:
    """System trust store plus the bundled certifi list.

    Windows Python reads the Windows store (so corporate proxy CAs keep working); python.org-style
    macOS builds have no system certificates, so certifi supplies GitHub's CA there.
    """
    global _ssl_context
    if _ssl_context is None:
        ctx = ssl.create_default_context()
        try:
            import certifi
            ctx.load_verify_locations(certifi.where())
        except (ImportError, OSError):
            pass
        _ssl_context = ctx
    return _ssl_context


def _open(url: str):
    _check_url(url)
    req = urllib.request.Request(url, headers={"User-Agent": f"PersonnelEvaluationReport/{VERSION}",
                                               "Accept": "application/vnd.github+json"})
    resp = urllib.request.urlopen(req, timeout=TIMEOUT, context=_context())
    _check_url(resp.geturl())                       # also after redirects
    return resp


def asset_name(version: str) -> str | None:
    """The file this computer needs from a release (None if updates are not supported here)."""
    if sys.platform == "win32":
        return f"PersonnelEvaluationReport-Setup-{version}.exe"
    if sys.platform == "darwin":
        arch = "arm64" if platform.machine() in ("arm64", "aarch64") else "x86_64"
        return f"PersonnelEvaluationReport-{version}-macOS-{arch}.dmg"
    return None


def can_self_update() -> bool:
    return bool(getattr(sys, "frozen", False)) and asset_name(VERSION) is not None


@dataclass
class ReleaseInfo:
    version: str
    newer: bool
    page_url: str
    asset: str | None
    asset_url: str | None
    sums_url: str | None


def check() -> ReleaseInfo:
    try:
        with _open(API_LATEST) as r:
            data = json.loads(r.read().decode("utf-8"))
    except UpdateError:
        raise
    except Exception:  # noqa: BLE001 – offline, DNS, proxy, HTTP errors…
        raise UpdateError("ارتباط با گیت‌هاب برقرار نشد. اتصال اینترنت را بررسی کنید و دوباره تلاش کنید.")
    try:
        version = str(data["tag_name"]).lstrip("v")
        parse_version(version)
    except (KeyError, ValueError):
        raise UpdateError("اطلاعات نسخهٔ تازه قابل خواندن نبود.")
    assets = {a.get("name"): a.get("browser_download_url") for a in data.get("assets", [])}
    name = asset_name(version)
    return ReleaseInfo(version=version, newer=is_newer(version), page_url=data.get("html_url") or RELEASES_PAGE,
                       asset=name, asset_url=assets.get(name) if name else None, sums_url=assets.get("SHA256SUMS.txt"))


def expected_sha256(sums_text: str, name: str) -> str:
    for line in sums_text.splitlines():
        parts = line.split()
        if len(parts) >= 2 and parts[-1].lstrip("*") == name and re.fullmatch(r"[0-9a-fA-F]{64}", parts[0]):
            return parts[0].lower()
    raise UpdateError("فهرست کنترل سلامت فایل‌ها (SHA256SUMS) برای این نسخه کامل نیست.")


@dataclass
class Job:
    """Background download state, polled by the settings page."""
    state: str = "idle"          # idle | downloading | ready | installing | error
    done: int = 0
    total: int = 0
    message: str = ""
    file: Path | None = None
    version: str = ""
    lock: threading.Lock = field(default_factory=threading.Lock)

    def as_dict(self) -> dict:
        pct = round(100 * self.done / self.total) if self.total else 0
        return {"state": self.state, "percent": pct, "message": self.message, "version": self.version}


def updates_dir() -> Path:
    d = user_data_dir() / "updates"
    d.mkdir(parents=True, exist_ok=True)
    return d


def clear_downloads() -> None:
    shutil.rmtree(user_data_dir() / "updates", ignore_errors=True)


def download(info: ReleaseInfo, job: Job) -> None:
    """Download and verify the installer for this computer (run in a background thread)."""
    try:
        if not (info.asset and info.asset_url and info.sums_url):
            raise UpdateError("فایل نصب این سیستم‌عامل در نسخهٔ تازه پیدا نشد.")
        with _open(info.sums_url) as r:
            want = expected_sha256(r.read().decode("utf-8", "replace"), info.asset)
        clear_downloads()
        target = updates_dir() / info.asset
        h = hashlib.sha256()
        with _open(info.asset_url) as r, open(target, "wb") as out:
            job.total = int(r.headers.get("Content-Length") or 0)
            while True:
                chunk = r.read(CHUNK)
                if not chunk:
                    break
                out.write(chunk)
                h.update(chunk)
                job.done += len(chunk)
        if h.hexdigest() != want:
            target.unlink(missing_ok=True)
            raise UpdateError("فایل دانلودشده سالم نبود (کنترل SHA-256 ناموفق). دوباره تلاش کنید.")
        job.file, job.state = target, "ready"
        job.message = f"نسخهٔ {to_fa_digits(info.version)} دانلود و بررسی شد و آمادهٔ نصب است."
    except UpdateError as e:
        job.state, job.message = "error", str(e)
    except Exception:  # noqa: BLE001
        job.state, job.message = "error", "دانلود ناموفق بود. اتصال اینترنت را بررسی کنید و دوباره تلاش کنید."


# ── installing ───────────────────────────────────────────────────────────

def _mac_bundle() -> Path:
    """…/Name.app for the running frozen executable (…/Name.app/Contents/MacOS/exe)."""
    exe = Path(sys.executable).resolve()
    for p in exe.parents:
        if p.suffix == ".app":
            return p
    raise UpdateError("مسیر برنامه پیدا نشد؛ نسخهٔ تازه را دستی نصب کنید.")


def prepare_install(file: Path) -> list[str]:
    """Validate and return the command that finishes the update after this process exits."""
    if sys.platform == "win32":
        # /SILENT shows a progress window only; the installer restarts the app afterwards.
        return [str(file), "/SP-", "/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CURRENTUSER", "/RESTARTAPP"]
    if sys.platform == "darwin":
        bundle = _mac_bundle()
        if not os.access(bundle.parent, os.W_OK):
            raise UpdateError("برنامه در پوشه‌ای نصب شده که اجازهٔ نوشتن ندارد. فایل dmg را از صفحهٔ دانلود "
                              "بگیرید و دستی جایگزین کنید.")
        script = Path(tempfile.mkdtemp(prefix="per-update-")) / "update.sh"
        script.write_text(_MAC_SCRIPT, encoding="utf-8")
        script.chmod(0o700)
        return ["/bin/sh", str(script), str(os.getpid()), str(file), str(bundle)]
    raise UpdateError("به‌روزرسانی خودکار روی این سیستم‌عامل پشتیبانی نمی‌شود.")


def launch(cmd: list[str]) -> None:
    kw: dict = {"stdin": subprocess.DEVNULL, "stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL, "close_fds": True}
    if sys.platform == "win32":
        kw["creationflags"] = 0x00000008 | 0x00000200    # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    else:
        kw["start_new_session"] = True
    subprocess.Popen(cmd, **kw)


# Waits for the app to exit, swaps the bundle (keeping the old one until the copy succeeds),
# reopens the app and cleans up. Arguments: pid, dmg path, bundle path.
_MAC_SCRIPT = r"""#!/bin/sh
pid="$1"; dmg="$2"; app="$3"
while kill -0 "$pid" 2>/dev/null; do sleep 0.5; done
mnt=$(mktemp -d /tmp/per-update-mnt.XXXXXX)
hdiutil attach -nobrowse -readonly -mountpoint "$mnt" "$dmg" >/dev/null 2>&1 || { open "$app"; exit 1; }
new=$(ls -d "$mnt"/*.app 2>/dev/null | head -n 1)
if [ -n "$new" ]; then
  rm -rf "$app.updating"
  if ditto "$new" "$app.updating"; then
    xattr -dr com.apple.quarantine "$app.updating" 2>/dev/null
    rm -rf "$app.old"
    if mv "$app" "$app.old"; then
      if mv "$app.updating" "$app"; then rm -rf "$app.old"; else mv "$app.old" "$app"; fi
    fi
  fi
fi
hdiutil detach "$mnt" >/dev/null 2>&1; rmdir "$mnt" 2>/dev/null
rm -f "$dmg"
open "$app"
rm -f "$0"
"""
