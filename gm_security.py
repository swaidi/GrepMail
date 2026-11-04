"""
Attachment security for GrepMail.

1. Virus scan   - Microsoft Defender (Windows) or ClamAV (Linux).
2. Macro check  - looks inside macro-enabled Office files without running them
                  (uses the optional 'oletools' package when it is installed).
3. Mark of the Web - tags macro files as "from outside this computer" (NTFS),
                  so Office blocks their macros by default.

Every function is defensive: a problem in the scanner never crashes the app.
"""
import os
import re
import sys
import glob
import shutil
import zipfile
import datetime as dt
import subprocess

IS_WINDOWS = sys.platform.startswith("win")

# result statuses (shown to the user in Title Case)
CLEAN = "Clean"
THREAT = "Threat Found"
NOT_SCANNED = "Not Scanned"

MACRO_EXTENSIONS = {"xlsm", "xlsb", "docm", "xltm", "dotm", "pptm"}
LEGACY_OFFICE = {"doc", "xls", "ppt"}
OOXML = {"xlsx", "xlsm", "xlsb", "docx", "docm", "xltm", "dotm", "pptx", "pptm"}

MACRO_NONE = "No Macros"
MACRO_PRESENT = "Macros Present"
MACRO_SUSPICIOUS = "Suspicious Macros"


def _run(cmd, timeout=120):
    flags = 0x08000000 if IS_WINDOWS else 0  # CREATE_NO_WINDOW
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                          errors="replace", creationflags=flags)


# =========================================================
# VIRUS SCANNERS
# =========================================================
class Scanner:
    name = "None"

    def available(self):
        return False

    def scan(self, path):
        """Returns (status, detail). status: CLEAN / THREAT / NOT_SCANNED."""
        return NOT_SCANNED, "No scanner available"

    def status(self):
        """Returns dict: engine, available, realtime, definitions_date, age_days, note."""
        return {"engine": self.name, "available": False, "definitions_date": None,
                "age_days": None, "note": "No scanner available"}

    def update_definitions(self):
        """Returns (ok, message)."""
        return False, "No scanner available"


class DefenderScanner(Scanner):
    name = "Microsoft Defender"

    def __init__(self):
        self._exe = None

    def _find(self):
        if self._exe and os.path.exists(self._exe):
            return self._exe
        candidates = []
        pdata = os.environ.get("ProgramData", r"C:\ProgramData")
        candidates += sorted(glob.glob(os.path.join(
            pdata, "Microsoft", "Windows Defender", "Platform", "*", "MpCmdRun.exe")), reverse=True)
        for pf in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                   os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            candidates.append(os.path.join(pf, "Windows Defender", "MpCmdRun.exe"))
        for c in candidates:
            if os.path.exists(c):
                self._exe = c
                return c
        return None

    def available(self):
        return IS_WINDOWS and self._find() is not None

    def scan(self, path):
        exe = self._find()
        if not exe:
            return NOT_SCANNED, "Microsoft Defender was not found"
        if not os.path.exists(path):
            # real-time protection may already have removed the file
            return THREAT, "File removed by the antivirus on write"
        try:
            r = _run([exe, "-Scan", "-ScanType", "3", "-File", path, "-DisableRemediation"])
        except Exception as e:
            return NOT_SCANNED, "Scan could not run: %s" % e
        out = (r.stdout or "") + (r.stderr or "")
        if not os.path.exists(path):
            return THREAT, "File removed by the antivirus"
        low = out.lower()
        if "found no threats" in low or ("no threats" in low and r.returncode == 0):
            return CLEAN, ""
        m = re.search(r"Threat\s*:\s*(.+)", out)
        if m or "found 1 threat" in low or re.search(r"found \d+ threats", low):
            return THREAT, (m.group(1).strip() if m else "Threat reported by Defender")
        if r.returncode == 0:
            return CLEAN, ""
        return NOT_SCANNED, "Defender returned code %s" % r.returncode

    def status(self):
        info = {"engine": self.name, "available": self.available(), "realtime": None,
                "definitions_date": None, "age_days": None, "note": ""}
        if not info["available"]:
            info["note"] = "Microsoft Defender was not found (another antivirus may be installed)."
            return info
        try:
            ps = ("$s=Get-MpComputerStatus;"
                  "Write-Output ('{0}|{1}|{2}' -f $s.AMServiceEnabled,$s.RealTimeProtectionEnabled,"
                  "$s.AntivirusSignatureLastUpdated.ToString('yyyy-MM-dd HH:mm'))")
            r = _run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=30)
            parts = (r.stdout or "").strip().split("|")
            if len(parts) == 3:
                info["realtime"] = parts[1].strip().lower() == "true"
                when = dt.datetime.strptime(parts[2].strip(), "%Y-%m-%d %H:%M")
                info["definitions_date"] = when
                info["age_days"] = (dt.datetime.now() - when).days
                if parts[0].strip().lower() != "true":
                    info["note"] = "Defender service is off."
        except Exception as e:
            info["note"] = "Could not read Defender status: %s" % e
        return info

    def update_definitions(self):
        exe = self._find()
        if not exe:
            return False, "Microsoft Defender was not found."
        try:
            r = _run([exe, "-SignatureUpdate"], timeout=300)
            out = ((r.stdout or "") + (r.stderr or "")).strip()
            if r.returncode == 0:
                return True, "Virus definitions updated."
            return False, ("Update failed (code %s). Open Windows Security > Virus & threat "
                           "protection > Check for updates.\n%s" % (r.returncode, out[-300:]))
        except Exception as e:
            return False, "Update could not run: %s" % e


class ClamAVScanner(Scanner):
    name = "ClamAV"

    def __init__(self):
        self._scan = shutil.which("clamdscan") or shutil.which("clamscan")
        self._clamscan = shutil.which("clamscan")

    def available(self):
        return self._scan is not None

    def scan(self, path):
        if not self._scan:
            return NOT_SCANNED, "ClamAV is not installed"
        if not os.path.exists(path):
            return THREAT, "File removed by the antivirus"
        cmd = [self._scan, "--no-summary", path]
        if os.path.basename(self._scan) == "clamdscan":
            cmd.insert(1, "--fdpass")
        try:
            r = _run(cmd)
        except Exception as e:
            return NOT_SCANNED, "Scan could not run: %s" % e
        out = (r.stdout or "") + (r.stderr or "")
        if r.returncode == 0:
            return CLEAN, ""
        if r.returncode == 1:
            m = re.search(r":\s*(.+?)\s+FOUND", out)
            return THREAT, (m.group(1) if m else "Threat reported by ClamAV")
        return NOT_SCANNED, "ClamAV returned code %s: %s" % (r.returncode, out.strip()[-200:])

    def status(self):
        info = {"engine": self.name, "available": self.available(), "realtime": None,
                "definitions_date": None, "age_days": None, "note": ""}
        if not info["available"]:
            info["note"] = "ClamAV is not installed. Install it with: sudo apt install clamav clamav-daemon"
            return info
        try:
            r = _run([self._clamscan or self._scan, "--version"], timeout=30)
            m = re.search(r"/(\d+)/(.+)$", (r.stdout or "").strip())
            if m:
                when = dt.datetime.strptime(m.group(2).strip(), "%a %b %d %H:%M:%S %Y")
                info["definitions_date"] = when
                info["age_days"] = (dt.datetime.now() - when).days
        except Exception as e:
            info["note"] = "Could not read ClamAV status: %s" % e
        return info

    def update_definitions(self):
        fresh = shutil.which("freshclam")
        if not fresh:
            return False, "freshclam was not found. Install it with: sudo apt install clamav"
        try:
            r = _run([fresh], timeout=600)
            out = ((r.stdout or "") + (r.stderr or "")).strip()
            if r.returncode == 0:
                return True, "Virus definitions updated."
            return False, ("Update needs administrator rights. Run in a terminal:\n"
                           "  sudo freshclam\n%s" % out[-200:])
        except Exception as e:
            return False, "Update could not run: %s" % e


def get_scanner():
    """The right scanner for this system (Defender on Windows, ClamAV elsewhere)."""
    return DefenderScanner() if IS_WINDOWS else ClamAVScanner()


# =========================================================
# MACRO INSPECTION
# =========================================================
try:  # optional but recommended
    from oletools.olevba import VBA_Parser
    HAS_OLETOOLS = True
except Exception:  # pragma: no cover
    VBA_Parser = None
    HAS_OLETOOLS = False


def _ext(path):
    return os.path.splitext(path)[1].lstrip(".").lower()


def has_macro_container(path):
    """Cheap check (no oletools): does the file contain a VBA project?"""
    ext = _ext(path)
    try:
        if ext in OOXML:
            with zipfile.ZipFile(path) as z:
                return any(n.lower().endswith("vbaproject.bin") for n in z.namelist())
        if ext in LEGACY_OFFICE:
            with open(path, "rb") as f:
                data = f.read(8 * 1024 * 1024)
            return b"VBA" in data and (b"_VBA_PROJECT" in data or b"Macros" in data)
    except Exception:
        return False
    return False


def inspect_macros(path):
    """Returns (status, detail).
    status: MACRO_NONE / MACRO_PRESENT / MACRO_SUSPICIOUS.
    The macro code is only read, never executed."""
    ext = _ext(path)
    if ext not in MACRO_EXTENSIONS and ext not in LEGACY_OFFICE and ext not in OOXML:
        return MACRO_NONE, ""
    present = has_macro_container(path)
    if not present and ext not in MACRO_EXTENSIONS:
        return MACRO_NONE, ""
    if not HAS_OLETOOLS:
        if present or ext in MACRO_EXTENSIONS:
            return MACRO_PRESENT, "Macro code was not analyzed (oletools is not installed)"
        return MACRO_NONE, ""
    try:
        parser = VBA_Parser(path)
        try:
            if not parser.detect_vba_macros():
                return MACRO_NONE, ""
            auto, suspicious, other = [], [], []
            for kind, keyword, desc in parser.analyze_macros():
                k = (kind or "").lower()
                if k == "autoexec":
                    auto.append(keyword)
                elif k in ("suspicious", "ioc", "hex strings", "base64 strings",
                           "dridex strings", "vba obfuscated strings"):
                    suspicious.append(keyword)
                else:
                    other.append(keyword)
            if suspicious or (auto and other):
                bits = []
                if auto:
                    bits.append("runs automatically (%s)" % ", ".join(sorted(set(auto))[:3]))
                if suspicious:
                    bits.append("suspicious: %s" % ", ".join(sorted(set(suspicious))[:5]))
                return MACRO_SUSPICIOUS, "; ".join(bits)
            return MACRO_PRESENT, ("runs automatically (%s)" % ", ".join(sorted(set(auto))[:3])
                                   if auto else "")
        finally:
            try:
                parser.close()
            except Exception:
                pass
    except Exception as e:
        # a macro-capable file that cannot be read (for example encrypted) is not trusted
        if present or ext in MACRO_EXTENSIONS:
            return MACRO_SUSPICIOUS, "Could not be inspected: %s" % str(e)[:80]
        return MACRO_NONE, ""


def mark_from_internet(path):
    """Adds the 'Mark of the Web' (Windows NTFS only) so Office opens the file in
    Protected View and blocks its macros. Returns True when applied."""
    if not IS_WINDOWS:
        return False
    try:
        with open(path + ":Zone.Identifier", "w", encoding="ascii") as f:
            f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
        return True
    except Exception:
        return False


def quarantine(path, folder, reason=""):
    """Moves a file into the quarantine folder (never into the Local Path)."""
    os.makedirs(folder, exist_ok=True)
    base = os.path.basename(path)
    dest = os.path.join(folder, dt.datetime.now().strftime("%Y%m%d-%H%M%S-") + base)
    shutil.move(path, dest)
    try:
        with open(dest + ".reason.txt", "w", encoding="utf-8") as f:
            f.write(reason + "\n")
    except Exception:
        pass
    return dest


def evaluate_attachment(path, scanner, macro_mode="flag", scan_enabled=True):
    """Runs every check on a staged file.
    Returns dict(action, security, detail, macro) where
      action   : 'save' | 'threat' | 'quarantine' | 'blocked'
      security : text for the Saved Files register ('Clean', 'Clean - Macro-Enabled', ...)
    """
    detail = ""
    scan_status = NOT_SCANNED
    if scan_enabled and scanner is not None:
        scan_status, d = scanner.scan(path)
        if scan_status == THREAT:
            return {"action": "threat", "security": THREAT, "detail": d, "macro": ""}
        detail = d
    macro, mdetail = inspect_macros(path)
    base = "Clean" if scan_status == CLEAN else "Not Scanned"
    if macro == MACRO_SUSPICIOUS:
        return {"action": "quarantine", "security": base + " - Suspicious Macros",
                "detail": mdetail, "macro": macro}
    if macro == MACRO_PRESENT:
        if macro_mode == "block":
            return {"action": "blocked", "security": base + " - Macro-Enabled",
                    "detail": "Macro-enabled files are blocked in Settings", "macro": macro}
        flag = " - Macro-Enabled" if macro_mode in ("flag", "block") else ""
        return {"action": "save", "security": base + flag, "detail": mdetail or detail,
                "macro": macro}
    return {"action": "save", "security": base, "detail": detail, "macro": MACRO_NONE}
