# SPDX-License-Identifier: AGPL-3.0-or-later
"""
Update check for GrepMail (GitHub releases).

Only the public GitHub release API is used. The downloaded installer is accepted only from
GitHub hosts over HTTPS, and its SHA-256 checksum is verified when GitHub provides one.
No personal data is sent: the request carries only the program name and version.
"""
import hashlib
import json
import os
import re
import sys
import tempfile
import urllib.error
import urllib.request
from urllib.parse import urlparse

REPO = "swaidi/GrepMail"
API_URL = "https://api.github.com/repos/%s/releases/latest"
RELEASES_PAGE = "https://github.com/%s/releases"
_ALLOWED_SUFFIXES = ("github.com", "githubusercontent.com")


def parse_version(text):
    """'v2.1.0' -> (2, 1, 0). Anything after a '-' (e.g. '-beta') is ignored."""
    head = str(text or "").strip().lstrip("vV").split("-")[0].split("+")[0]
    nums = [int(n) for n in re.findall(r"\d+", head)]
    return tuple(nums) if nums else (0,)


def is_newer(latest, current):
    a, b = parse_version(latest), parse_version(current)
    n = max(len(a), len(b))
    return a + (0,) * (n - len(a)) > b + (0,) * (n - len(b))


def pick_asset(assets, windows=None):
    """The installer for this system: Setup .exe on Windows, .deb elsewhere."""
    if windows is None:
        windows = sys.platform.startswith("win")
    for a in assets or []:
        name = str(a.get("name", "")).lower()
        if windows and name.endswith(".exe") and "setup" in name:
            return a
        if not windows and name.endswith(".deb"):
            return a
    return None


def _https_github(url):
    u = urlparse(url)
    host = (u.hostname or "").lower()
    return u.scheme == "https" and any(host == s or host.endswith("." + s) for s in _ALLOWED_SUFFIXES)


def _fetch_json(url, timeout, app_version):
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json", "User-Agent": "GrepMail/%s" % app_version})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def check_latest(current_version, repo=REPO, timeout=10, fetch=None, windows=None):
    """Returns None when the program is up to date (or no release exists), otherwise a dict:
    version, notes, page, asset (name, url, size, sha256 or ''), or asset=None when this
    system has no installer in the release. Raises on network errors."""
    fetch = fetch or (lambda url: _fetch_json(url, timeout, current_version))
    try:
        data = fetch(API_URL % repo)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    tag = str(data.get("tag_name", ""))
    if not tag or not is_newer(tag, current_version):
        return None
    asset = pick_asset(data.get("assets"), windows)
    info = None
    if asset and _https_github(str(asset.get("browser_download_url", ""))):
        digest = str(asset.get("digest", "") or "")
        info = {"name": asset.get("name", ""), "url": asset["browser_download_url"],
                "size": int(asset.get("size") or 0),
                "sha256": digest.split(":", 1)[1].lower() if digest.lower().startswith("sha256:") else ""}
    return {"version": tag.lstrip("vV"), "tag": tag, "notes": str(data.get("body") or "").strip(),
            "page": data.get("html_url") or RELEASES_PAGE % repo, "asset": info}


def download(asset, dest_dir=None, progress=None, app_version="", cancel=None, opener=None):
    """Downloads the installer and verifies it. Returns the file path. Raises on any problem."""
    if not _https_github(asset["url"]):
        raise ValueError("The download address is not an HTTPS GitHub address.")
    dest_dir = dest_dir or tempfile.mkdtemp(prefix="grepmail_update_")
    os.makedirs(dest_dir, exist_ok=True)
    path = os.path.join(dest_dir, os.path.basename(asset["name"]))
    req = urllib.request.Request(asset["url"], headers={"User-Agent": "GrepMail/%s" % app_version})
    open_fn = opener or (lambda r: urllib.request.urlopen(r, timeout=30))
    h, done = hashlib.sha256(), 0
    with open_fn(req) as r:
        if hasattr(r, "geturl") and not _https_github(r.geturl()):
            raise ValueError("The download was redirected to a non-GitHub address.")
        total = int(r.headers.get("Content-Length") or asset.get("size") or 0)
        with open(path, "wb") as f:
            while True:
                if cancel is not None and cancel.is_set():
                    raise InterruptedError("Download cancelled.")
                chunk = r.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                h.update(chunk)
                done += len(chunk)
                if progress:
                    progress(done, total)
    if asset.get("size") and done != asset["size"]:
        os.remove(path)
        raise ValueError("The downloaded file has the wrong size.")
    if asset.get("sha256") and h.hexdigest().lower() != asset["sha256"]:
        os.remove(path)
        raise ValueError("The downloaded file failed the checksum test and was deleted.")
    return path
