"""
Oracle Document Organizer
=========================
Scans an _Inbox/ folder, classifies each file by Oracle module and version,
renames it using a consistent convention, and moves it to the right folder.

Folder structure created automatically:
  <root>/
    _Inbox/                   <- drop new files here
    EPM/
      FCCS/ ARCS/ PBCS/ Free_Form/ Other/
    R12_ERP/
      AP/ AR/ GL/ CE/ AHCS_FAH/ PO/ Security/
      Cost_Inventory/ Projects/ FA/ Other/
    Fusion_Cloud_ERP/
      AP/ AR/ GL/ CE/ AHCS_FAH/ PO/ Security/
      Cost_Inventory/ Projects/ FA/ Other/
    Other/
      Data_Models/            <- SQL / data model files
      Other/                  <- unclassified

File naming convention:
  MODULE_VERSION_Title_YYYY-MM-DD.ext
  e.g.  AP_R12_Invoice_Processing_Guide_2024-03-01.pdf

Usage:
  python organizer.py <root_folder> [--dry-run]

Dependencies:
  pip install pdfplumber python-docx openpyxl striprtf
"""

import argparse
import logging
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

# ── optional dependencies (graceful degradation) ──────────────────────────────

try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from striprtf.striprtf import rtf_to_text
    HAS_RTF = True
except ImportError:
    HAS_RTF = False


# ── folder / classification map ───────────────────────────────────────────────

STRUCTURE = {
    "EPM":             ["FCCS", "ARCS", "PBCS", "Free_Form", "Other"],
    "R12_ERP":         ["P2P", "O2C", "R2R", "Security",
                        "Cost_Inventory", "Projects", "FA", "Other"],
    "Fusion_Cloud_ERP":["P2P", "O2C", "R2R", "Security",
                        "Cost_Inventory", "Projects", "FA", "Other"],
    "Other":           ["Data_Models", "Other"],
}

# Keywords that map to (top_folder, sub_folder)
# Checked against filename + first ~2000 chars of content (case-insensitive)
MODULE_KEYWORDS = [
    # EPM products
    (["fccs", "financial consolidation", "close cloud"], "EPM", "FCCS"),
    (["arcs", "account reconciliation"], "EPM", "ARCS"),
    (["pbcs", "planning and budgeting", "planning cloud"], "EPM", "PBCS"),
    (["free form", "freeform", "free-form"], "EPM", "Free_Form"),

    # ERP modules — version detected separately
    # P2P — Procure to Pay: AP, PO, Procurement, AP-related Cash Management
    (["accounts payable", "payables", "invoice approval", "invoice processing",
      "supplier invoice", "supplier payment", "supplier site", "supplier bank",
      "searching supplier", "payment process", "payment file", "modify payment",
      "withholding tax", "isupplier", "spreadsheet invoice", "ap inv",
      "payables dashboard", "payables implementation", "payables essential",
      "payables payment", "invoice recognition", "ofc ap", "electronic invoic",
      "purchase order", r"\bpo\b", "iprocurement", "procurement", "requisition",
      "procure to pay", "p2p", "ofc p2p",
      "receiving", "approval setup", "document approval",
      "business unit.*scm", "scm.*procurement",
      "cash management", "bank reconciliation", "bank statement", "bank account",
      "external bank", "payments overview", "cash mgt", "automatic reconciliation",
      "roles.*po", "setups.*po", "enterprise structure.*erp", "erp.*enterprise structure"], None, "P2P"),
    # O2C — Order to Cash: AR, Order Entry, Order Management, Cash receipts
    (["accounts receivable", "order to cash", r"\bo2c\b", "order entry",
      "order management", "customer receipt", "cash receipt",
      "customer payment", "credit management", "collections",
      "customer invoice", r"\bar\b"], None, "O2C"),
    # R2R — Record to Report: GL, COA, CVR, Enterprise Structure, AHCS, FAH, SLA
    (["general ledger", r"\bgl\b", "chart of accounts", r"\bcoa\b",
      "journal entry", "journal import", "journal line",
      "cross validation", r"\bcvr\b", "segment value",
      "enterprise structure", "ledger setup", "ledger config",
      "accounting hub", "ahcs", "fah", "financial accounting hub",
      "subledger accounting", r"\bsla\b", "subledger", "record to report",
      r"\br2r\b", "period close", "intercompany", "consolidat"], None, "R2R"),
    (["user role", "responsibility", "access control", "profile option", "system admin",
      "roles and setup", "role setup", r"\brbac\b", "security setup"], None, "Security"),
    (["cost management", "costing", "inventory valuation", "cost accounting",
      "cycle count", "physical inventory", "inventory count", "inventory guide",
      "inventory", "stock", "item cost", "material cost"], None, "Cost_Inventory"),
    (["project accounting", "project costing", "project billing", "project planning",
      "project financ", "project perf", "project revenue", "ppm office",
      "projectcost", "projectrev", "projectperf", "optimize.*project"], None, "Projects"),
    (["fixed assets", "asset management", "depreciation", "asset book"], None, "FA"),
]

VERSION_KEYWORDS = {
    "R12":    ["r12", "release 12", "ebs 12", "e-business suite 12"],
    "Fusion": ["fusion", "cloud erp", "oracle cloud", "saas"],
    "R11i":   ["r11i", "11i", "release 11", "11.5", "11.0"],
    "EPM":    ["epm", "hyperion", "planning cloud", "fccs", "arcs", "pbcs"],
}

SQL_EXTENSIONS  = {".sql"}
DOC_EXTENSIONS  = {".pdf", ".docx", ".doc", ".txt", ".rtf", ".xlsx", ".xls"}
ALL_EXTENSIONS  = SQL_EXTENSIONS | DOC_EXTENSIONS


# ── text extraction ───────────────────────────────────────────────────────────

def extract_text(path: Path, max_chars: int = 3000) -> str:
    """Extract a sample of text from the file for classification."""
    ext = path.suffix.lower()

    try:
        if ext == ".pdf" and HAS_PDF:
            with pdfplumber.open(path) as pdf:
                text = ""
                for page in pdf.pages[:3]:
                    text += (page.extract_text() or "")
                    if len(text) >= max_chars:
                        break
            return text[:max_chars]

        if ext in (".docx", ".doc") and HAS_DOCX:
            doc = DocxDocument(path)
            text = " ".join(p.text for p in doc.paragraphs[:50])
            return text[:max_chars]

        if ext in (".xlsx", ".xls") and HAS_XLSX:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            text = ""
            for ws in wb.worksheets[:2]:
                for row in ws.iter_rows(max_row=20, values_only=True):
                    text += " ".join(str(c) for c in row if c) + " "
                    if len(text) >= max_chars:
                        break
            return text[:max_chars]

        if ext == ".rtf" and HAS_RTF:
            raw = path.read_bytes().decode("utf-8", errors="ignore")
            return rtf_to_text(raw)[:max_chars]

        if ext in (".txt", ".sql"):
            return path.read_text(encoding="utf-8", errors="ignore")[:max_chars]

    except Exception as e:
        logging.warning(f"Could not extract text from {path.name}: {e}")

    return ""


# ── classification ────────────────────────────────────────────────────────────

def detect_version(text: str) -> str | None:
    tl = text.lower()
    for version, keywords in VERSION_KEYWORDS.items():
        for kw in keywords:
            if re.search(kw, tl):
                return version
    return None


def classify(path: Path, content: str) -> tuple[str, str]:
    """
    Return (top_folder, sub_folder).
    SQL files always go to Other/Data_Models.
    """
    if path.suffix.lower() in SQL_EXTENSIONS:
        return "Other", "Data_Models"

    # include stem with underscores replaced by spaces so OFC_P2P matches "p2p"
    stem_spaced = path.stem.replace("_", " ")
    haystack = (path.stem + " " + stem_spaced + " " + content).lower()
    version  = detect_version(haystack)

    matched_module = None
    for keywords, top_hint, module in MODULE_KEYWORDS:
        for kw in keywords:
            if re.search(kw, haystack):
                matched_module = (top_hint, module)
                break
        if matched_module:
            break

    if not matched_module:
        return "Other", "Other"

    top_hint, module = matched_module

    # EPM modules always go under EPM
    if top_hint == "EPM":
        return "EPM", module

    # ERP modules — route by version
    if version == "R12" or version == "R11i":
        return "R12_ERP", module
    if version == "Fusion" or version == "EPM":
        return "Fusion_Cloud_ERP", module

    # version unclear — default to Fusion_Cloud_ERP rather than losing the file
    return "Fusion_Cloud_ERP", module


# ── naming ────────────────────────────────────────────────────────────────────

MAX_FILENAME = 100  # characters excluding extension

# Words to shorten in titles (case-insensitive, longest phrases first)
ABBREVIATIONS = [
    # Phrases first (before individual words)
    (r"\bBI Publisher\b",         "BIP"),
    (r"\bProcure to Pay\b",       "P2P"),
    (r"\bOrder to Cash\b",        "O2C"),
    (r"\bRecord to Report\b",     "R2R"),
    (r"\bAccounting Hub\b",       "AHCS"),
    (r"\bSubledger Accounting\b", "SLA"),
    (r"\bBest Practices\b",       "Best Prac"),
    (r"\bWhite Paper\b",          "WP"),
    # Individual words
    (r"\bPayables\b",             "AP"),
    (r"\bReceivables\b",          "AR"),
    (r"\bSubledger\b",            "SLA"),
    (r"\bConfiguration\b",        "Config"),
    (r"\bConfigure\b",            "Config"),
    (r"\bNotifications?\b",       "Notif"),
    (r"\bImplementation\b",       "Impl"),
    (r"\bProcurement\b",          "Proc"),
    (r"\bManagement\b",           "Mgmt"),
    (r"\bInformation\b",          "Info"),
    (r"\bEnterprise\b",           "Ent"),
    (r"\bStructures?\b",          "Struct"),
    (r"\bPerformance\b",          "Perf"),
    (r"\bFinancial\b",            "Fin"),
    (r"\bAccounting\b",           "Acctg"),
    (r"\bTroubleshooting\b",      "Troubleshoot"),
    (r"\bPublisher\b",            "Pub"),
    (r"\bTemplates?\b",           "Tmpl"),
    (r"\bSolutions?\b",           "Soln"),
    (r"\bOrganization\b",         "Org"),
    (r"\bAuthorization\b",        "Auth"),
    (r"\bAdministration\b",       "Admin"),
    (r"\bApplication\b",          "App"),
    (r"\bTransactions?\b",        "Txn"),
    (r"\bReconciliation\b",       "Recon"),
    (r"\bConsolidation\b",        "Consol"),
    (r"\bDocuments?\b",           "Doc"),
    (r"\bIntegration\b",          "Integ"),
    (r"\bCustomization\b",        "Custom"),
    (r"\bOverview\b",             "Overview"),
]


def _clean(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|+\-_]', " ", value)  # remove unsafe + special chars
    value = re.sub(r"\s+", " ", value.strip())          # collapse double spaces
    return value


def _shorten(title: str) -> str:
    """Apply abbreviations to reduce title length."""
    for pattern, replacement in ABBREVIATIONS:
        title = re.sub(pattern, replacement, title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title.strip())
    return title


def _r2r_module_tag(haystack: str) -> str:
    """For R2R files, detect the specific module tag to use in the filename."""
    h = haystack.lower()
    if any(re.search(k, h) for k in ["accounting hub", "ahcs", r"\bfah\b", "financial accounting hub"]):
        return "AHCS"
    if any(re.search(k, h) for k in ["subledger accounting", r"\bsla\b", "subledger"]):
        return "SLA"
    if any(re.search(k, h) for k in ["cross validation", r"\bcvr\b"]):
        return "CVR"
    if any(re.search(k, h) for k in ["chart of accounts", r"\bcoa\b"]):
        return "COA"
    if any(re.search(k, h) for k in ["enterprise structure"]):
        return "Ent Struct"
    return "GL"  # default for R2R


def build_filename(path: Path, top: str, sub: str, haystack: str = "") -> str:
    """Build new filename: MODULE VERSION Title YYYY-MM-DD.ext, max 100 chars."""
    ext = path.suffix.lower()

    # derive version tag
    version_tag = {
        "R12_ERP":          "R12",
        "Fusion_Cloud_ERP": "Fusion",
        "EPM":              "EPM",
        "Other":            "",
    }.get(top, "")

    # R2R files get a specific module tag (GL, AHCS, COA etc.) not just "R2R"
    if sub == "R2R":
        module_tag = _r2r_module_tag(haystack or path.stem)
    else:
        module_tag = sub if sub not in ("Other", "Data_Models") else ""

    # file creation/modification date
    mtime = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d")

    # strip previously applied prefixes so re-runs don't double up
    stem = path.stem
    # strip previously applied prefixes — handle both _ and space separators
    sep = r"[\s_]+"
    stem = re.sub(r"^(?:OFC" + sep + r")?(?:AP_PO|P2P|O2C|R2R|GL|FA|AHCS_FAH|AHCS|Security|Cost_Inventory|Projects)" + sep + r"(?:Fusion|R12|EPM|R11i)" + sep, "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"^(?:AP_PO|P2P|O2C|R2R|GL|FA|AHCS_FAH|AHCS|Security|Cost_Inventory|Projects)" + sep, "", stem, flags=re.IGNORECASE)
    # strip trailing duplicate dates like _2024-01-01 or space-2024-01-01
    stem = re.sub(r"[\s_]\d{4}[\s\-]\d{2}[\s\-]\d{2}$", "", stem)
    stem = re.sub(r"[\s_]\d{4}-\d{2}-\d{2}$", "", stem)
    title = _shorten(_clean(stem))

    # build full name and enforce 80-char limit on the title portion
    prefix = " ".join(p for p in [module_tag, version_tag] if p)
    suffix = mtime
    # available space for title = MAX_FILENAME - prefix - suffix - 2 spaces
    reserved = len(prefix) + len(suffix) + (2 if prefix else 0) + 1
    max_title = MAX_FILENAME - reserved
    if len(title) > max_title:
        title = title[:max_title].rsplit(" ", 1)[0]  # trim at word boundary

    parts = [p for p in [module_tag, version_tag, title, mtime] if p]
    return " ".join(parts) + ext


# ── move logic ────────────────────────────────────────────────────────────────

def _delete_empty_dirs(folder: Path):
    """Recursively delete empty directories bottom-up inside folder."""
    for dirpath in sorted(folder.rglob("*"), reverse=True):
        if dirpath.is_dir() and dirpath != folder:
            try:
                dirpath.rmdir()  # only succeeds if empty
            except OSError:
                pass  # not empty, skip


def process_inbox(root: Path, dry_run: bool):
    inbox = root / "_Inbox"
    if not inbox.exists():
        inbox.mkdir(parents=True)
        print(f"Created _Inbox at {inbox} -- add files there and re-run.")
        return

    # recursively find all supported files in inbox and subfolders
    files = [f for f in inbox.rglob("*") if f.is_file() and f.suffix.lower() in ALL_EXTENSIONS]

    if not files:
        print(f"No supported files found in {inbox}")
        return

    log_lines = []
    print(f"{'DRY RUN -- ' if dry_run else ''}Processing {len(files)} file(s) from _Inbox (incl. subfolders)\n")

    for path in sorted(files):
        # show relative path from inbox so subfolder context is visible
        rel = path.relative_to(inbox)
        print(f"  {rel}")
        content  = extract_text(path)
        top, sub = classify(path, content)
        stem_spaced = path.stem.replace("_", " ")
        haystack = (path.stem + " " + stem_spaced + " " + content).lower()
        new_name = build_filename(path, top, sub, haystack)
        dest_dir = root / top / sub
        dest     = dest_dir / new_name

        # avoid collisions
        counter = 1
        while dest.exists():
            stem = f"{Path(new_name).stem} {counter}"
            dest = dest_dir / f"{stem}{path.suffix.lower()}"
            counter += 1

        print(f"    -> {top}/{sub}/{dest.name}")
        log_lines.append(f"{rel}  ->  {top}/{sub}/{dest.name}")

        if not dry_run:
            dest_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(path), str(dest))

    # delete empty folders left behind in inbox
    if not dry_run:
        _delete_empty_dirs(inbox)
        print(f"\nEmpty inbox subfolders removed.")

    # write log
    log_path = root / f"organizer_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    summary  = "\n".join(log_lines)
    if not dry_run:
        log_path.write_text(summary, encoding="utf-8")
        print(f"\nLog saved to {log_path.name}")
    else:
        print(f"\nDry run complete -- no files moved.")

    print(f"\n{len(files)} file(s) processed.")


# ── setup ─────────────────────────────────────────────────────────────────────

def setup_folders(root: Path):
    """Create the full folder structure if it doesn't exist."""
    for top, subs in STRUCTURE.items():
        for sub in subs:
            (root / top / sub).mkdir(parents=True, exist_ok=True)
    (root / "_Inbox").mkdir(exist_ok=True)
    print(f"Folder structure ready at: {root}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Oracle document organizer.")
    parser.add_argument("root", help="Root folder where Oracle Docs structure lives")
    parser.add_argument("--dry-run", action="store_true", help="Preview moves without changing anything")
    parser.add_argument("--setup", action="store_true", help="Create folder structure and exit")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)

    if args.setup:
        setup_folders(root)
        return

    process_inbox(root, dry_run=args.dry_run)


if __name__ == "__main__":
    logging.basicConfig(level=logging.WARNING)
    main()
