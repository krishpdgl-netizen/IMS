"""
main.py - Panache Digilife Factory Inventory Management System (v2)
=====================================================================

The whole backend in one file. Storage is still ONE .xlsx file in a GitHub
repo (no SQL database). What changed in v2 is WHAT is stored and HOW:

  Items            item master (FG / sub-assembly / part / spare ...), serialized flag, HSN, warranty
  BOM              bill of materials: which components (and how many) make one finished unit
  Inwards          one row per shipment (GRN): CBU / SKD / CKD / LOCAL, supplier, invoice,
                   BL/AWB, container, Bill of Entry, kit check, QC status
  Inward Lines     one row per item on a GRN: invoice qty vs received vs damaged vs accepted
  Batches          the stock ledger. Every lot of stock sits in Warehouse + Zone + Bin.
                   Zones: QC_HOLD, RAW_STORE, WIP, FG_STORE, QUARANTINE
  Serials          one row per serialised unit (with MAC/IMEI), where it is, its status,
                   which work order built/consumed it, parent FG serial, customer, warranty
  Work Orders      production orders (SKD/CKD assembly into finished goods)
  Consumption      which component batches/serials went into which work-order output
                   (full traceability FG serial -> components -> GRN / supplier)
  Transaction Log  every movement, one row per item
  _Meta, _Journal  counters and the undo journal (do not edit by hand)

Every column is read BY HEADER NAME, never by position, so adding or
re-ordering columns in Excel can never shift data again.

Every write is a "transaction": changes are journaled so the most recent
transaction can be undone exactly, and the GitHub save uses the file's sha
so two people saving at once can never silently overwrite each other
(the request is retried on top of the newer file instead).

Environment variables (Vercel -> Project Settings):
  GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH (main), EXCEL_PATH (data/inventory_data.xlsx)
  ACCESS_CODE          shared login password
  GEMINI_API_KEY       for the assistant and the document scanner
  WAREHOUSES           optional, comma separated, default "Bhiwandi,Ghatkopar"
  AGING_THRESHOLD_DAYS optional, default 90
  OWN_COMPANY_NAME, OWN_DELIVERY_ADDRESSES  optional, help the scanner tell inward from outward
  LOCAL_XLSX           optional, path to a local .xlsx instead of GitHub (local testing)
"""

import os
import io
import re
import json
import math
import base64
import hashlib
from datetime import datetime, date, timedelta, timezone
from typing import Dict, List, Optional

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse


# ============================================================================
# SETTINGS
# ============================================================================
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN") or ""
GITHUB_REPO = os.environ.get("GITHUB_REPO") or ""
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH") or "main"
EXCEL_PATH = os.environ.get("EXCEL_PATH") or "data/inventory_data.xlsx"
ACCESS_CODE = os.environ.get("ACCESS_CODE") or "inventory2026"
LOCAL_XLSX = os.environ.get("LOCAL_XLSX") or ""

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY") or ""
GEMINI_MODEL = "gemini-3.1-flash-lite"  # fixed per requirements
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"

AGING_THRESHOLD_DAYS = int(os.environ.get("AGING_THRESHOLD_DAYS") or "90")
OWN_DELIVERY_ADDRESSES = os.environ.get("OWN_DELIVERY_ADDRESSES") or ""
OWN_COMPANY_NAME = os.environ.get("OWN_COMPANY_NAME") or "Panache Digilife"

WAREHOUSES = [w.strip() for w in (os.environ.get("WAREHOUSES") or "Bhiwandi,Ghatkopar").split(",") if w.strip()]

IST = timezone(timedelta(hours=5, minutes=30))

SCHEMA_VERSION = "2"

# ---- Fixed vocabularies ------------------------------------------------------
INWARD_TYPES = {
    "CBU": "CBU - Completely Built Up (finished goods, ready to sell)",
    "SKD": "SKD - Semi Knocked Down (needs final assembly)",
    "CKD": "CKD - Completely Knocked Down (kit of parts)",
    "LOCAL": "Local / domestic purchase",
}
ITEM_TYPES = {
    "FG": "Finished good (saleable unit)",
    "SUB_ASSEMBLY": "Sub-assembly (SKD semi-built unit)",
    "PART": "Component / part (CKD)",
    "SPARE": "Spare part",
    "PACKING": "Packing material",
    "CONSUMABLE": "Consumable",
}
ZONES = {
    "QC_HOLD": "QC Hold (awaiting inspection)",
    "RAW_STORE": "Raw / Parts Store",
    "WIP": "Production floor (WIP)",
    "FG_STORE": "Finished Goods Store",
    "QUARANTINE": "Quarantine (damaged / rejected)",
}
QC_STATUSES = ["PENDING", "PASSED", "REJECTED"]
OUTWARD_TYPES = {
    "SALE": {"label": "Sale / customer dispatch", "zones": ["FG_STORE", "RAW_STORE"], "serial_status": "Dispatched"},
    "SAMPLE": {"label": "Sample / demo unit out", "zones": ["FG_STORE"], "serial_status": "Dispatched"},
    "RTV": {"label": "Return to vendor", "zones": ["QUARANTINE", "RAW_STORE", "FG_STORE"], "serial_status": "Returned to Vendor"},
    "SCRAP": {"label": "Scrap / write-off", "zones": ["QUARANTINE", "RAW_STORE", "FG_STORE"], "serial_status": "Scrapped"},
}
MOVABLE_ZONES = ["RAW_STORE", "FG_STORE", "QUARANTINE"]   # zones a manual transfer may touch

# ---- Sheet layout -------------------------------------------------------------
ITEMS, BOM, INWARDS, LINES, BATCHES, SERIALS = "Items", "BOM", "Inwards", "Inward Lines", "Batches", "Serials"
WOS, CONS, LOG, META, JOURNAL = "Work Orders", "Consumption", "Transaction Log", "_Meta", "_Journal"

SCHEMA = {
    ITEMS: ("Item Code", ["Item Code", "Item Name", "Normalized Name", "Item Type", "Category", "Model",
                          "Variant", "Unit", "Serialized", "HSN Code", "Warranty Months", "Reorder Level",
                          "Status", "Remarks", "Created At"]),
    BOM: ("BOM Line ID", ["BOM Line ID", "FG Item Code", "FG Item Name", "Component Code", "Component Name",
                          "Qty Per Unit", "Unit", "Remarks", "Updated At"]),
    INWARDS: ("GRN No", ["GRN No", "Inward Type", "GRN Date", "Warehouse", "Supplier", "Country of Origin",
                         "PO No", "Invoice No", "Invoice Date", "Currency", "Exchange Rate", "Invoice Value",
                         "BL / AWB No", "Container No", "Bill of Entry No", "BoE Date", "Port of Entry",
                         "Vehicle No", "Gate Entry No", "Kit For Item", "Kits Count", "Kit Check",
                         "QC Status", "Total Lines", "Total Accepted", "Total Short", "Total Damaged",
                         "Received By", "Status", "Remarks", "Source Document", "Txn ID", "Created At"]),
    LINES: ("Line ID", ["Line ID", "GRN No", "Item Code", "Item Name", "Item Type", "HSN Code",
                        "Invoice Qty", "Received Qty", "Short Qty", "Excess Qty", "Damaged Qty",
                        "Accepted Qty", "Unit", "Unit Price", "Currency", "Lot No", "Mfg Date",
                        "Zone", "Bin", "QC Status", "Batch ID", "Damaged Batch ID", "Serial Count", "Remarks"]),
    BATCHES: ("Batch ID", ["Batch ID", "Item Code", "Item Name", "Item Type", "Warehouse", "Zone", "Bin",
                           "Date Received", "Qty Received", "Qty Remaining", "Unit", "Status", "QC Status",
                           "GRN No", "Inward Type", "Lot No", "Work Order", "Unit Cost", "Currency",
                           "Source", "Parent Batch ID", "Txn ID", "Created At"]),
    SERIALS: ("Serial Key", ["Serial Key", "Serial No", "Item Code", "Item Name", "MAC / IMEI", "Batch ID",
                             "GRN No", "Inward Type", "Warehouse", "Zone", "Bin", "Status", "QC Status",
                             "Work Order", "Parent Serial", "Customer", "Dispatch Doc", "Dispatch Date",
                             "Warranty Until", "Received Date", "Updated At"]),
    WOS: ("WO No", ["WO No", "FG Item Code", "FG Item Name", "Build Type", "Planned Qty", "Produced Qty",
                    "Rejected Qty", "Production Line", "Warehouse", "Status", "Start Date", "Target Date",
                    "Completed Date", "Source GRN", "Created By", "Remarks", "Txn ID", "Created At"]),
    CONS: ("Cons ID", ["Cons ID", "Txn ID", "Date", "WO No", "FG Item Code", "FG Serials", "Component Code",
                       "Component Name", "Qty", "Unit", "From Batch ID", "GRN No", "Component Serials"]),
    LOG: ("Log ID", ["Log ID", "Txn ID", "Timestamp", "Txn Type", "Ref No", "Item Code", "Item Name", "Qty",
                     "Unit", "From Warehouse", "From Zone", "To Warehouse", "To Zone", "Bin", "Serials",
                     "Party", "Doc No", "Entered By", "Status", "Message", "Remarks"]),
    META: ("Key", ["Key", "Value"]),
    JOURNAL: ("Seq", ["Seq", "Txn ID", "Op", "Sheet", "Row Key", "Column", "Old Value"]),
}
SHEET_ORDER = [ITEMS, BOM, INWARDS, LINES, BATCHES, SERIALS, WOS, CONS, LOG, META, JOURNAL]
# Rows appended by a transaction in these sheets are kept (status -> Reversed) on undo, for audit.
UNDO_KEEP = {LOG: "Status", INWARDS: "Status", WOS: "Status"}
JOURNAL_KEEP_TXNS = 100
CELL_LIMIT = 32000

_HEADER_FILL = PatternFill("solid", fgColor="0F766E")
_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)


# ============================================================================
# SMALL HELPERS
# ============================================================================
def now_iso() -> str:
    return datetime.now(IST).replace(microsecond=0).isoformat()


def today() -> str:
    return datetime.now(IST).date().isoformat()


def s(v) -> str:
    return "" if v is None else str(v).strip()


def num(v, default=0.0) -> float:
    if v is None or v == "":
        return default
    try:
        f = float(v)
    except (TypeError, ValueError):
        raise ValueError(f"'{v}' is not a number.")
    if math.isnan(f) or math.isinf(f):
        raise ValueError(f"'{v}' is not a valid number.")
    return f


def clean(v):
    """Store whole numbers as int so Excel shows 17, not 17.0."""
    if isinstance(v, float):
        v = round(v, 4)
        if v == int(v):
            return int(v)
    return v


def norm(v) -> str:
    return re.sub(r"\s+", " ", s(v)).casefold()


def valid_date(v, field="Date", required=False) -> str:
    v = s(v)
    if not v:
        if required:
            raise ValueError(f"{field} is required.")
        return ""
    v = v[:10]
    try:
        datetime.strptime(v, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"{field} must be a date (YYYY-MM-DD), got '{v}'.")
    return v


def parse_d(v) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(s(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def fit(text: str) -> str:
    text = s(text)
    if len(text) <= CELL_LIMIT:
        return text
    return text[:CELL_LIMIT - 40] + f" ... (truncated, {len(text)} chars)"


def warehouse_ok(v, field="Warehouse") -> str:
    v = s(v)
    for w in WAREHOUSES:
        if w.casefold() == v.casefold():
            return w
    raise ValueError(f"{field} must be one of: {', '.join(WAREHOUSES)}.")


def zone_ok(v, allowed=None, field="Zone") -> str:
    v = s(v).upper().replace(" ", "_")
    allowed = allowed or list(ZONES)
    if v not in allowed:
        raise ValueError(f"{field} must be one of: {', '.join(allowed)}.")
    return v


def parse_serials(raw) -> List[dict]:
    """Accepts a list or text. One unit per line (or comma separated).
    'SERIAL | MAC' or 'SERIAL<TAB>MAC' attaches a MAC/IMEI to the serial."""
    if raw is None:
        return []
    if isinstance(raw, list):
        parts = []
        for x in raw:
            if isinstance(x, dict):
                parts.append({"serial": s(x.get("serial") or x.get("serial_no")), "mac": s(x.get("mac"))})
            else:
                parts.extend(parse_serials(str(x)))
        return [p for p in parts if p["serial"]]
    out = []
    for line in str(raw).replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        if "\t" in line or "|" in line:
            a, _, b = line.replace("\t", "|").partition("|")
            if a.strip():
                out.append({"serial": a.strip(), "mac": b.strip()})
            continue
        for tok in re.split(r"[,;]", line):
            tok = tok.strip()
            if tok:
                out.append({"serial": tok, "mac": ""})
    return out


def serial_key(item_code, serial) -> str:
    return f"{s(item_code).upper()}::{s(serial).upper()}"


def add_months(d: date, months: int) -> date:
    y, m = divmod(d.month - 1 + months, 12)
    y += d.year
    m += 1
    days = [31, 29 if (y % 4 == 0 and (y % 100 or y % 400 == 0)) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return date(y, m, min(d.day, days))


def is_yes(v) -> bool:
    return s(v).upper() in ("Y", "YES", "TRUE", "1")


# ============================================================================
# STORAGE: GitHub (default) or a local file (LOCAL_XLSX, for testing)
# ============================================================================
class Conflict(Exception):
    pass


class GitHubStore:
    API = "https://api.github.com"

    def _headers(self, accept="application/vnd.github+json"):
        return {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": accept}

    def _url(self):
        return f"{self.API}/repos/{GITHUB_REPO}/contents/{EXCEL_PATH}"

    def load(self):
        if not GITHUB_TOKEN or not GITHUB_REPO:
            raise RuntimeError("GITHUB_TOKEN / GITHUB_REPO are not set on the server.")
        r = requests.get(self._url(), headers=self._headers(), params={"ref": GITHUB_BRANCH}, timeout=30)
        if r.status_code == 404:
            return None, None
        r.raise_for_status()
        meta = r.json()
        sha = meta["sha"]
        if meta.get("content") and meta.get("encoding") == "base64":
            return base64.b64decode(meta["content"]), sha
        # Files over 1 MB come back without content: fetch the raw bytes instead.
        r2 = requests.get(self._url(), headers=self._headers("application/vnd.github.raw"),
                          params={"ref": GITHUB_BRANCH}, timeout=60)
        r2.raise_for_status()
        return r2.content, sha

    def save(self, raw: bytes, sha: Optional[str], message: str):
        payload = {"message": message, "content": base64.b64encode(raw).decode(), "branch": GITHUB_BRANCH}
        if sha:
            payload["sha"] = sha
        r = requests.put(self._url(), headers=self._headers(), json=payload, timeout=60)
        if r.status_code in (409, 422) and ("sha" in r.text.lower() or r.status_code == 409):
            raise Conflict(r.text[:300])
        r.raise_for_status()

    def download_url(self):
        return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_PATH}"


class LocalStore:
    def __init__(self, path):
        self.path = path

    def load(self):
        if not os.path.exists(self.path):
            return None, None
        with open(self.path, "rb") as f:
            raw = f.read()
        return raw, hashlib.sha1(raw).hexdigest()

    def save(self, raw, sha, message):
        _, current = self.load()
        if current != sha:
            raise Conflict("file changed since it was read")
        os.makedirs(os.path.dirname(os.path.abspath(self.path)), exist_ok=True)
        with open(self.path, "wb") as f:
            f.write(raw)

    def download_url(self):
        return ""


STORE = LocalStore(LOCAL_XLSX) if LOCAL_XLSX else GitHubStore()


# ============================================================================
# TABLE: header-name based access to one sheet, with undo journaling
# ============================================================================
class Table:
    def __init__(self, book, name):
        self.book = book
        self.name = name
        self.key = SCHEMA[name][0]
        self.ws = book.wb[name]
        self._load()

    def _load(self):
        self.headers = [c.value for c in self.ws[1]]
        self.col = {h: i + 1 for i, h in enumerate(self.headers) if h}
        self.recs: List[dict] = []
        self.by_key: Dict[str, dict] = {}
        self.rownum: Dict[str, int] = {}
        kidx = self.col[self.key] - 1
        for r_i, row in enumerate(self.ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or kidx >= len(row) or row[kidx] in (None, ""):
                continue
            rec = {h: ("" if v is None else v) for h, v in zip(self.headers, row) if h}
            k = s(rec[self.key])
            self.recs.append(rec)
            self.by_key[k] = rec
            self.rownum[k] = r_i

    def rows(self) -> List[dict]:
        return self.recs

    def get(self, key) -> Optional[dict]:
        return self.by_key.get(s(key))

    def append(self, rec: dict) -> dict:
        k = s(rec.get(self.key))
        if not k:
            raise RuntimeError(f"{self.name}: missing key {self.key}")
        if k in self.by_key:
            raise ValueError(f"{self.name}: '{k}' already exists.")
        full = {h: ("" if rec.get(h) is None else clean(rec.get(h))) for h in self.headers if h}
        for h, v in full.items():
            if isinstance(v, str):
                full[h] = fit(v)
        self.ws.append([(None if full.get(h, "") == "" else full.get(h)) if h else None for h in self.headers])
        self.recs.append(full)
        self.by_key[k] = full
        self.rownum[k] = self.ws.max_row
        self.book.journal_append(self.name, k)
        return full

    def update(self, key, changes: dict):
        k = s(key)
        rec = self.by_key.get(k)
        if rec is None:
            raise RuntimeError(f"{self.name}: row '{k}' not found")
        r = self.rownum[k]
        for col, val in changes.items():
            if col not in self.col:
                continue
            val = clean(val)
            if isinstance(val, str):
                val = fit(val)
            old = rec.get(col, "")
            if old == val:
                continue
            self.book.journal_update(self.name, k, col, old)
            rec[col] = val
            self.ws.cell(row=r, column=self.col[col], value=None if val == "" else val)

    def delete_keys(self, keys):
        rows = sorted({self.rownum[k] for k in keys if k in self.rownum}, reverse=True)
        for r in rows:
            self.ws.delete_rows(r)
        self._load()


# ============================================================================
# BOOK: the whole workbook + transaction / journal handling
# ============================================================================
class Book:
    def __init__(self, wb: Workbook):
        self.wb = wb
        self.warnings: List[str] = []
        self.txn = None
        self._journal: List[tuple] = []
        self._appended = set()
        self._updates = {}
        self._ensure_schema()
        self.t = {name: Table(self, name) for name in SHEET_ORDER}

    # -- loading / schema -------------------------------------------------------
    @classmethod
    def from_bytes(cls, raw: Optional[bytes]):
        if raw is None:
            return cls(Workbook())
        return cls(load_workbook(io.BytesIO(raw)))

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        self.wb.save(buf)
        return buf.getvalue()

    def _ensure_schema(self):
        wb = self.wb
        version = None
        if META in wb.sheetnames:
            for row in wb[META].iter_rows(min_row=2, values_only=True):
                if row and row[0] == "schema_version":
                    version = str(row[1])
        legacy_products = []
        if version != SCHEMA_VERSION:
            # Old (v1) workbook: keep its sheets for reference under a "Legacy" name.
            for ws in list(wb.worksheets):
                if ws.title == "Sheet" and ws.max_row <= 1 and ws.max_column <= 1:
                    continue
                if ws.title in SCHEMA or ws.title in ("Products",):
                    if ws.title == "Products":
                        hdr = [c.value for c in ws[1]]
                        if "Product Name" in hdr:
                            i = hdr.index("Product Name")
                            legacy_products = [r[i] for r in ws.iter_rows(min_row=2, values_only=True) if r and r[i]]
                    ws.title = ("Legacy " + ws.title)[:31]
        for name in SHEET_ORDER:
            headers = SCHEMA[name][1]
            if name not in wb.sheetnames:
                ws = wb.create_sheet(name)
                for i, h in enumerate(headers, start=1):
                    c = ws.cell(row=1, column=i, value=h)
                    c.font, c.fill, c.alignment = _HEADER_FONT, _HEADER_FILL, Alignment(horizontal="center")
                ws.freeze_panes = "A2"
                for i in range(1, len(headers) + 1):
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 16
            else:
                ws = wb[name]
                existing = [c.value for c in ws[1]]
                for h in headers:
                    if h not in existing:
                        c = ws.cell(row=1, column=len(existing) + 1, value=h)
                        c.font, c.fill = _HEADER_FONT, _HEADER_FILL
                        existing.append(h)
        if "Sheet" in wb.sheetnames and wb["Sheet"].max_row <= 1 and len(wb.sheetnames) > 1:
            del wb["Sheet"]
        for name in (META, JOURNAL):
            wb[name].sheet_state = "hidden"
        if version != SCHEMA_VERSION:
            meta = wb[META]
            meta.append(["schema_version", SCHEMA_VERSION])
            self._legacy_products = legacy_products
        else:
            self._legacy_products = []

    def import_legacy_products(self):
        if not getattr(self, "_legacy_products", None):
            return
        for name in self._legacy_products:
            if not self.find_item(name):
                self.create_item({"item_name": name, "item_type": "FG", "serialized": "N", "unit": "units",
                                  "remarks": "Imported from old product list - please review type/serialized"})
        self._legacy_products = []

    # -- counters ---------------------------------------------------------------
    def next_id(self, counter: str, fmt: str) -> str:
        meta = self.t[META]
        rec = meta.get(counter)
        n = int(num(rec["Value"])) + 1 if rec else 1
        if rec:
            rec["Value"] = n
            meta.ws.cell(row=meta.rownum[counter], column=meta.col["Value"], value=n)
        else:
            meta.ws.append([counter, n])
            meta._load()
        return fmt.format(n=n, y=datetime.now(IST).year)

    def _bump(self, counter, value):
        meta = self.t[META]
        meta.by_key[counter]["Value"] = value
        meta.ws.cell(row=meta.rownum[counter], column=meta.col["Value"], value=value)

    # -- journal ----------------------------------------------------------------
    def journal_append(self, sheet, key):
        if self.txn and sheet not in (META, JOURNAL):
            self._journal.append(("append", sheet, key, "", ""))
            self._appended.add((sheet, key))

    def journal_update(self, sheet, key, col, old):
        if self.txn and sheet not in (META, JOURNAL) and (sheet, key) not in self._appended:
            olds = self._updates.get((sheet, key))
            if olds is None:
                olds = self._updates[(sheet, key)] = {}
                self._journal.append(("update", sheet, key, "*", olds))
            olds.setdefault(col, old)

    def begin(self, txn_type: str, user: str):
        self.txn = {"id": self.next_id("txn", "TXN-{n:06d}"), "type": txn_type, "user": s(user) or "unknown"}
        self._journal, self._appended, self._updates = [], set(), {}
        return self.txn["id"]

    def commit(self):
        if not self.txn:
            return
        jws = self.wb[JOURNAL]
        if self._journal:
            last = int(self.next_id("journal", "{n}"))
            for i, op in enumerate(self._journal):
                old = json.dumps(op[4], default=str) if op[0] == "update" else ""
                jws.append([last + i, self.txn["id"], op[0], op[1], op[2], op[3], fit(old)])
            self._bump("journal", last + len(self._journal) - 1)
        self._prune_journal()
        self.txn = None

    def _prune_journal(self):
        jws = self.wb[JOURNAL]
        txns = []
        for row in jws.iter_rows(min_row=2, max_col=2, values_only=True):
            if row[1] and (not txns or txns[-1] != row[1]):
                txns.append(row[1])
        if len(txns) <= JOURNAL_KEEP_TXNS:
            return
        keep = set(txns[-JOURNAL_KEEP_TXNS:])
        n_drop = 0
        for row in jws.iter_rows(min_row=2, max_col=2, values_only=True):
            if row[1] in keep:
                break
            n_drop += 1
        if n_drop:
            jws.delete_rows(2, n_drop)

    def log(self, **f):
        tx = self.txn
        self.t[LOG].append({
            "Log ID": self.next_id("log", "L{n:07d}"), "Txn ID": tx["id"], "Timestamp": now_iso(),
            "Txn Type": f.get("type") or tx["type"], "Ref No": f.get("ref", ""), "Item Code": f.get("item_code", ""),
            "Item Name": f.get("item_name", ""), "Qty": f.get("qty", ""), "Unit": f.get("unit", ""),
            "From Warehouse": f.get("from_wh", ""), "From Zone": f.get("from_zone", ""),
            "To Warehouse": f.get("to_wh", ""), "To Zone": f.get("to_zone", ""), "Bin": f.get("bin", ""),
            "Serials": ", ".join(f.get("serials") or []), "Party": f.get("party", ""), "Doc No": f.get("doc", ""),
            "Entered By": tx["user"], "Status": "Applied", "Message": f.get("message", ""),
            "Remarks": f.get("remarks", ""),
        })

    # -- undo ---------------------------------------------------------------
    def undo_last(self, user: str) -> dict:
        log = self.t[LOG]
        target = None
        for rec in reversed(log.rows()):
            if rec["Status"] == "Applied" and rec["Txn Type"] != "UNDO":
                target = rec["Txn ID"]
                break
        if not target:
            raise ValueError("There is no applied transaction to undo.")
        jws = self.wb[JOURNAL]
        ops = [r for r in jws.iter_rows(min_row=2, values_only=True) if r and r[1] == target]
        if not ops:
            raise ValueError(f"{target} is too old to undo automatically (its undo history was pruned).")
        ops.sort(key=lambda r: int(r[0]), reverse=True)
        summary = [r for r in log.rows() if r["Txn ID"] == target]
        deletes: Dict[str, List[str]] = {}
        for _, _, op, sheet, key, col, old in ops:
            tbl = self.t.get(sheet)
            if tbl is None:
                continue
            if op == "update":
                if tbl.get(key) is not None:
                    olds = json.loads(old) if old else {}
                    tbl.update(key, {c: ("" if v is None else v) for c, v in olds.items()})
            elif op == "append":
                if sheet in UNDO_KEEP:
                    if tbl.get(key) is not None:
                        tbl.update(key, {UNDO_KEEP[sheet]: "Reversed"})
                else:
                    deletes.setdefault(sheet, []).append(s(key))
        for sheet, keys in deletes.items():
            self.t[sheet].delete_keys(keys)
        # drop the journal of the undone transaction
        rows = [i for i, r in enumerate(jws.iter_rows(min_row=2, max_col=2, values_only=True), start=2) if r[1] == target]
        for r in sorted(rows, reverse=True):
            jws.delete_rows(r)
        first = summary[0] if summary else {}
        self.begin("UNDO", user)
        self.log(type="UNDO", ref=target, item_code=first.get("Item Code", ""), item_name=first.get("Item Name", ""),
                 message=f"Reversed {target} ({first.get('Txn Type', '')})")
        self._journal = []  # an undo itself is not undoable
        self.commit()
        return {"txn_id": target, "type": first.get("Txn Type", ""), "item": first.get("Item Name", ""),
                "lines": len(summary)}

    # ======================================================================
    # MASTER DATA
    # ======================================================================
    def find_item(self, ref) -> Optional[dict]:
        ref = s(ref)
        if not ref:
            return None
        rec = self.t[ITEMS].get(ref.upper()) or self.t[ITEMS].get(ref)
        if rec:
            return rec
        key = norm(ref)
        for r in self.t[ITEMS].rows():
            if r["Normalized Name"] == key or norm(r["Item Code"]) == key:
                return r
        return None

    def item(self, ref) -> dict:
        rec = self.find_item(ref)
        if not rec:
            raise ValueError(f"Item '{ref}' is not in the Item Master. Add it on the Masters page first.")
        return rec

    def create_item(self, spec: dict) -> dict:
        name = re.sub(r"\s+", " ", s(spec.get("item_name") or spec.get("name")))
        if not name:
            raise ValueError("Item name is required.")
        code = s(spec.get("item_code") or spec.get("code")).upper()
        if self.find_item(name):
            raise ValueError(f"An item named '{name}' already exists.")
        if code and self.t[ITEMS].get(code):
            raise ValueError(f"Item code '{code}' already exists.")
        if not code:
            code = self.next_id("item", "ITM-{n:05d}")
            while self.t[ITEMS].get(code):
                code = self.next_id("item", "ITM-{n:05d}")
        itype = s(spec.get("item_type") or "PART").upper()
        if itype not in ITEM_TYPES:
            raise ValueError(f"Item type must be one of: {', '.join(ITEM_TYPES)}.")
        return self.t[ITEMS].append({
            "Item Code": code, "Item Name": name, "Normalized Name": norm(name), "Item Type": itype,
            "Category": s(spec.get("category")), "Model": s(spec.get("model")), "Variant": s(spec.get("variant")),
            "Unit": s(spec.get("unit")) or ("pcs" if itype != "FG" else "units"),
            "Serialized": "Y" if is_yes(spec.get("serialized")) else "N",
            "HSN Code": s(spec.get("hsn_code")), "Warranty Months": clean(num(spec.get("warranty_months"), 0)),
            "Reorder Level": clean(num(spec.get("reorder_level"), 0)), "Status": "Active",
            "Remarks": s(spec.get("remarks")), "Created At": now_iso(),
        })

    def update_item(self, code, spec: dict) -> dict:
        rec = self.item(code)
        changes = {}
        mapping = {"item_name": "Item Name", "item_type": "Item Type", "category": "Category", "model": "Model",
                   "variant": "Variant", "unit": "Unit", "hsn_code": "HSN Code", "warranty_months": "Warranty Months",
                   "reorder_level": "Reorder Level", "status": "Status", "remarks": "Remarks"}
        for k, col in mapping.items():
            if k in spec and spec[k] is not None:
                v = spec[k]
                if k == "item_type":
                    v = s(v).upper()
                    if v not in ITEM_TYPES:
                        raise ValueError(f"Item type must be one of: {', '.join(ITEM_TYPES)}.")
                if k in ("warranty_months", "reorder_level"):
                    v = num(v, 0)
                if k == "item_name":
                    v = re.sub(r"\s+", " ", s(v))
                    other = self.find_item(v)
                    if other and other["Item Code"] != rec["Item Code"]:
                        raise ValueError(f"Another item is already called '{v}'.")
                    changes["Normalized Name"] = norm(v)
                changes[col] = v
        if "serialized" in spec and spec["serialized"] is not None:
            new = "Y" if is_yes(spec["serialized"]) else "N"
            if new != rec["Serialized"] and self.on_hand(rec["Item Code"]) > 0:
                raise ValueError("Cannot change 'serialized' while the item has stock. Move stock out first.")
            changes["Serialized"] = new
        self.t[ITEMS].update(rec["Item Code"], changes)
        return rec

    def get_or_create_item(self, line: dict) -> dict:
        ref = s(line.get("item_code")) or s(line.get("item")) or s(line.get("item_name"))
        rec = self.find_item(ref) or (self.find_item(line.get("item_name")) if line.get("item_name") else None)
        if rec:
            if rec["Status"] == "Inactive":
                raise ValueError(f"Item '{rec['Item Name']}' is marked Inactive.")
            return rec
        if not line.get("item_type"):
            raise ValueError(f"Item '{ref}' is new - choose its item type (and serialized Y/N) so it can be added to the Item Master.")
        return self.create_item({"item_name": line.get("item_name") or line.get("item"), "item_code": line.get("item_code"),
                                 "item_type": line.get("item_type"), "serialized": line.get("serialized"),
                                 "unit": line.get("unit"), "hsn_code": line.get("hsn_code"),
                                 "warranty_months": line.get("warranty_months")})

    def bom_for(self, fg_code) -> List[dict]:
        fg_code = s(fg_code)
        return [r for r in self.t[BOM].rows() if s(r["FG Item Code"]) == fg_code]

    def set_bom(self, fg_ref, components: List[dict]) -> List[dict]:
        fg = self.item(fg_ref)
        if fg["Item Type"] not in ("FG", "SUB_ASSEMBLY"):
            raise ValueError("A BOM can only be defined for a finished good or sub-assembly.")
        seen = set()
        clean_rows = []
        for c in components:
            if not s(c.get("item")) and not s(c.get("item_code")):
                continue
            comp = self.item(c.get("item_code") or c.get("item"))
            if comp["Item Code"] == fg["Item Code"]:
                raise ValueError("An item cannot be a component of itself.")
            if comp["Item Code"] in seen:
                raise ValueError(f"'{comp['Item Name']}' is listed twice in the BOM.")
            seen.add(comp["Item Code"])
            q = num(c.get("qty_per_unit") or c.get("qty"))
            if q <= 0:
                raise ValueError(f"Qty per unit for '{comp['Item Name']}' must be more than 0.")
            if comp["Serialized"] == "Y" and q != int(q):
                raise ValueError(f"'{comp['Item Name']}' is serialized, so qty per unit must be a whole number.")
            clean_rows.append((comp, q, s(c.get("remarks"))))
        old = [r["BOM Line ID"] for r in self.bom_for(fg["Item Code"])]
        self.t[BOM].delete_keys(old)
        for comp, q, rem in clean_rows:
            self.t[BOM].append({"BOM Line ID": self.next_id("bom", "BOM-{n:06d}"), "FG Item Code": fg["Item Code"],
                                "FG Item Name": fg["Item Name"], "Component Code": comp["Item Code"],
                                "Component Name": comp["Item Name"], "Qty Per Unit": q, "Unit": comp["Unit"],
                                "Remarks": rem, "Updated At": now_iso()})
        return self.bom_for(fg["Item Code"])

    # ======================================================================
    # STOCK ENGINE
    # ======================================================================
    def batches_for(self, item_code, warehouse=None, zones=None, wo=None) -> List[dict]:
        out = []
        for b in self.t[BATCHES].rows():
            if s(b["Item Code"]) != s(item_code) or num(b["Qty Remaining"]) <= 0:
                continue
            if warehouse and b["Warehouse"] != warehouse:
                continue
            if zones and b["Zone"] not in zones:
                continue
            if wo is not None and s(b["Work Order"]) != s(wo):
                continue
            out.append(b)
        out.sort(key=lambda b: (s(b["Date Received"]), s(b["Batch ID"])))
        return out

    def on_hand(self, item_code, warehouse=None, zones=None, wo=None) -> float:
        return round(sum(num(b["Qty Remaining"]) for b in self.batches_for(item_code, warehouse, zones, wo)), 4)

    def add_batch(self, item: dict, warehouse, zone, bin_, qty, date_received, source, **extra) -> dict:
        return self.t[BATCHES].append({
            "Batch ID": self.next_id("batch", "B{n:07d}"), "Item Code": item["Item Code"], "Item Name": item["Item Name"],
            "Item Type": item["Item Type"], "Warehouse": warehouse, "Zone": zone, "Bin": s(bin_),
            "Date Received": date_received, "Qty Received": qty, "Qty Remaining": qty, "Unit": item["Unit"],
            "Status": "Active", "QC Status": extra.get("qc_status", ""), "GRN No": extra.get("grn", ""),
            "Inward Type": extra.get("inward_type", ""), "Lot No": extra.get("lot", ""),
            "Work Order": extra.get("wo", ""), "Unit Cost": extra.get("unit_cost", ""),
            "Currency": extra.get("currency", ""), "Source": source, "Parent Batch ID": extra.get("parent", ""),
            "Txn ID": self.txn["id"] if self.txn else "", "Created At": now_iso(),
        })

    def _decrement(self, batch: dict, qty: float):
        left = round(num(batch["Qty Remaining"]) - qty, 4)
        if left < -1e-9:
            raise RuntimeError("batch went negative")
        self.t[BATCHES].update(batch["Batch ID"], {"Qty Remaining": max(left, 0),
                                                   "Status": "Active" if left > 0 else "Depleted"})

    def take_qty(self, item: dict, qty: float, warehouse, zones, wo=None, where="") -> List[dict]:
        """FIFO: oldest batch first. Non-serialized items only."""
        if qty <= 0:
            raise ValueError("Quantity must be more than 0.")
        avail = self.on_hand(item["Item Code"], warehouse, zones, wo)
        if avail + 1e-9 < qty:
            loc = where or f"{warehouse} / {', '.join(zones)}" + (f" / {wo}" if wo else "")
            raise ValueError(f"Not enough '{item['Item Name']}': need {clean(qty)}, only {clean(avail)} {item['Unit']} at {loc}.")
        allocs, need = [], qty
        for b in self.batches_for(item["Item Code"], warehouse, zones, wo):
            if need <= 1e-9:
                break
            t = min(num(b["Qty Remaining"]), need)
            self._decrement(b, t)
            allocs.append({"batch": dict(b), "qty": round(t, 4), "serials": []})
            need = round(need - t, 4)
        return allocs

    def take_serials(self, item: dict, serials: List[dict], warehouse, zones, wo=None) -> List[dict]:
        names = [x["serial"] for x in serials]
        dup = {x for x in names if names.count(x) > 1}
        if dup:
            raise ValueError(f"Serial(s) entered twice: {', '.join(sorted(dup)[:10])}")
        groups: Dict[str, dict] = {}
        problems = []
        for x in serials:
            rec = self.t[SERIALS].get(serial_key(item["Item Code"], x["serial"]))
            if not rec:
                problems.append(f"{x['serial']}: not found for {item['Item Name']}")
                continue
            if rec["Status"] != "In Stock":
                problems.append(f"{x['serial']}: status is {rec['Status']}")
                continue
            if warehouse and rec["Warehouse"] != warehouse:
                problems.append(f"{x['serial']}: is at {rec['Warehouse']}, not {warehouse}")
                continue
            if zones and rec["Zone"] not in zones:
                problems.append(f"{x['serial']}: is in {rec['Zone']}, not {'/'.join(zones)}")
                continue
            if wo is not None and s(rec["Work Order"]) != s(wo):
                problems.append(f"{x['serial']}: is not issued to {wo}")
                continue
            groups.setdefault(s(rec["Batch ID"]), {"serials": [], "recs": []})
            groups[s(rec["Batch ID"])]["serials"].append(rec["Serial No"])
            groups[s(rec["Batch ID"])]["recs"].append(rec)
        if problems:
            more = f" (+{len(problems) - 8} more)" if len(problems) > 8 else ""
            raise ValueError("Serial problems: " + "; ".join(problems[:8]) + more)
        allocs = []
        for bid, g in groups.items():
            b = self.t[BATCHES].get(bid)
            if not b or num(b["Qty Remaining"]) + 1e-9 < len(g["serials"]):
                raise ValueError(f"Batch {bid} does not hold the serials listed - stock and serial records disagree.")
            self._decrement(b, len(g["serials"]))
            allocs.append({"batch": dict(b), "qty": len(g["serials"]), "serials": g["serials"], "recs": g["recs"]})
        return allocs

    def pick_serials(self, item: dict, qty: int, warehouse, zones, wo=None, exclude=()) -> List[dict]:
        """Auto-pick serial numbers FIFO (oldest batch first)."""
        excl = {s(x).upper() for x in exclude}
        pool = [r for r in self.t[SERIALS].rows()
                if r["Item Code"] == item["Item Code"] and r["Status"] == "In Stock"
                and (not warehouse or r["Warehouse"] == warehouse) and (not zones or r["Zone"] in zones)
                and (wo is None or s(r["Work Order"]) == s(wo)) and s(r["Serial No"]).upper() not in excl]
        batch_date = {b["Batch ID"]: s(b["Date Received"]) for b in self.t[BATCHES].rows()}
        pool.sort(key=lambda r: (batch_date.get(r["Batch ID"], ""), s(r["Batch ID"]), s(r["Serial No"])))
        if len(pool) < qty:
            loc = f"{warehouse} / {'/'.join(zones or [])}" + (f" / {wo}" if wo else "")
            raise ValueError(f"Not enough serialised '{item['Item Name']}': need {qty}, only {len(pool)} at {loc}.")
        return [{"serial": r["Serial No"], "mac": r["MAC / IMEI"]} for r in pool[:qty]]

    def take(self, item, qty, serials, warehouse, zones, wo=None, auto_pick=False, where="") -> List[dict]:
        """Stock-out wrapper. Serialized items move by serial number."""
        if item["Serialized"] == "Y":
            if not serials:
                if not auto_pick:
                    raise ValueError(f"'{item['Item Name']}' is serialized - list the serial numbers being moved.")
                if qty != int(qty):
                    raise ValueError(f"'{item['Item Name']}' is serialized - quantity must be a whole number.")
                serials = self.pick_serials(item, int(qty), warehouse, zones, wo)
            if qty and len(serials) != int(qty):
                raise ValueError(f"'{item['Item Name']}': quantity is {clean(qty)} but {len(serials)} serial(s) were listed.")
            return self.take_serials(item, serials, warehouse, zones, wo)
        return self.take_qty(item, qty, warehouse, zones, wo, where)

    def put(self, item, allocs, warehouse, zone, bin_, source, wo="", qc_status=None, date_received=None) -> List[dict]:
        """Places taken stock somewhere else. Original receipt date, GRN, lot and cost are kept
        (so aging and traceability survive transfers)."""
        new_batches = []
        for a in allocs:
            ob = a["batch"]
            nb = self.add_batch(item, warehouse, zone, bin_ or ob["Bin"], a["qty"],
                                date_received or ob["Date Received"], source,
                                qc_status=ob["QC Status"] if qc_status is None else qc_status, grn=ob["GRN No"],
                                inward_type=ob["Inward Type"], lot=ob["Lot No"], wo=wo, unit_cost=ob["Unit Cost"],
                                currency=ob["Currency"], parent=ob["Batch ID"])
            for sn in a["serials"]:
                self.t[SERIALS].update(serial_key(item["Item Code"], sn), {
                    "Batch ID": nb["Batch ID"], "Warehouse": warehouse, "Zone": zone, "Bin": nb["Bin"],
                    "Work Order": wo if zone == "WIP" else self.t[SERIALS].get(serial_key(item["Item Code"], sn))["Work Order"],
                    "QC Status": nb["QC Status"], "Updated At": now_iso()})
            new_batches.append(nb)
        return new_batches

    def new_serials(self, item, serials: List[dict], batch: dict, **f):
        seen = set()
        for x in serials:
            k = serial_key(item["Item Code"], x["serial"])
            if k in seen:
                raise ValueError(f"Serial {x['serial']} entered twice.")
            seen.add(k)
            if self.t[SERIALS].get(k):
                ex = self.t[SERIALS].get(k)
                raise ValueError(f"Serial {x['serial']} of {item['Item Name']} already exists (status {ex['Status']}, GRN {ex['GRN No'] or '-'}).")
        for x in serials:
            self.t[SERIALS].append({
                "Serial Key": serial_key(item["Item Code"], x["serial"]), "Serial No": x["serial"],
                "Item Code": item["Item Code"], "Item Name": item["Item Name"], "MAC / IMEI": x.get("mac", ""),
                "Batch ID": batch["Batch ID"], "GRN No": batch["GRN No"], "Inward Type": batch["Inward Type"],
                "Warehouse": batch["Warehouse"], "Zone": batch["Zone"], "Bin": batch["Bin"], "Status": "In Stock",
                "QC Status": batch["QC Status"], "Work Order": f.get("wo", ""), "Parent Serial": "",
                "Received Date": batch["Date Received"], "Updated At": now_iso()})

    def check_new_serials(self, item, serials, qty, label="accepted"):
        if item["Serialized"] != "Y":
            if serials:
                self.warnings.append(f"'{item['Item Name']}' is not serialized in the Item Master; serials entered were ignored.")
            return []
        if qty != int(qty):
            raise ValueError(f"'{item['Item Name']}' is serialized - {label} quantity must be a whole number.")
        if len(serials) != int(qty):
            raise ValueError(f"'{item['Item Name']}': {label} quantity is {clean(qty)} but {len(serials)} serial number(s) were entered. "
                             f"Scan/paste one serial per unit.")
        return serials

    # ======================================================================
    # INWARD (GRN) - CBU / SKD / CKD / LOCAL
    # ======================================================================
    def kit_check(self, fg_ref, kits, received: Dict[str, float]) -> dict:
        fg = self.find_item(fg_ref)
        if not fg:
            return {"status": "Unknown kit item", "lines": []}
        bom = self.bom_for(fg["Item Code"])
        if not bom:
            return {"status": f"No BOM for {fg['Item Name']}", "lines": []}
        lines, short = [], 0
        for b in bom:
            req = round(num(b["Qty Per Unit"]) * kits, 4)
            got = round(received.get(b["Component Code"], 0), 4)
            diff = round(got - req, 4)
            if diff < 0:
                short += 1
            lines.append({"item_code": b["Component Code"], "item_name": b["Component Name"], "required": clean(req),
                          "received": clean(got), "difference": clean(diff), "unit": b["Unit"]})
        extra = [c for c in received if c not in {b["Component Code"] for b in bom}]
        for c in extra:
            it = self.find_item(c)
            lines.append({"item_code": c, "item_name": it["Item Name"] if it else c, "required": 0,
                          "received": clean(received[c]), "difference": clean(received[c]), "unit": it["Unit"] if it else "",
                          "not_in_bom": True})
        status = "Complete" if short == 0 else f"Short on {short} item(s)"
        return {"status": status, "lines": lines, "kit_item": fg["Item Name"], "kits": clean(kits)}

    def inward(self, d: dict) -> dict:
        itype = s(d.get("inward_type")).upper()
        if itype not in INWARD_TYPES:
            raise ValueError("Choose the inward type: CBU, SKD, CKD or LOCAL.")
        wh = warehouse_ok(d.get("warehouse"))
        supplier = s(d.get("supplier"))
        inv_no = s(d.get("invoice_no"))
        if not supplier:
            raise ValueError("Supplier is required.")
        if not inv_no:
            raise ValueError("Supplier invoice number is required.")
        for g in self.t[INWARDS].rows():
            if g["Status"] != "Reversed" and norm(g["Supplier"]) == norm(supplier) and norm(g["Invoice No"]) == norm(inv_no):
                raise ValueError(f"Invoice {inv_no} from {supplier} was already received on {g['GRN No']}.")
        grn_date = valid_date(d.get("grn_date") or today(), "GRN date")
        header_qc = s(d.get("qc_status") or "PENDING").upper()
        if header_qc not in ("PENDING", "PASSED"):
            raise ValueError("QC status at receipt must be PENDING or PASSED.")
        if itype in ("CBU", "SKD", "CKD") and not s(d.get("bill_of_entry_no")):
            self.warnings.append("Imported shipment saved without a Bill of Entry number - add it when available.")
        lines = [l for l in (d.get("lines") or []) if s(l.get("item")) or s(l.get("item_code")) or s(l.get("item_name"))]
        if not lines:
            raise ValueError("Add at least one item line.")
        currency = s(d.get("currency")) or ("INR" if itype == "LOCAL" else "USD")

        grn = self.next_id(f"grn{datetime.now(IST).year}", "GRN-{y}-{n:04d}")
        txn = self.txn["id"]
        received_by_code: Dict[str, float] = {}
        tot_acc = tot_short = tot_dmg = 0
        line_results = []
        for i, l in enumerate(lines, start=1):
            item = self.get_or_create_item(l)
            if itype == "CBU" and item["Item Type"] not in ("FG", "SPARE"):
                self.warnings.append(f"CBU inward but '{item['Item Name']}' is a {item['Item Type']} - check the item type.")
            if itype in ("SKD", "CKD") and item["Item Type"] == "FG":
                self.warnings.append(f"{itype} inward but '{item['Item Name']}' is a finished good - should it be a part/sub-assembly?")
            inv_q = num(l.get("invoice_qty"))
            rec_q = num(l.get("received_qty"), inv_q)
            dmg_q = num(l.get("damaged_qty"), 0)
            if inv_q < 0 or rec_q < 0 or dmg_q < 0:
                raise ValueError(f"Line {i}: quantities cannot be negative.")
            if rec_q <= 0 and inv_q <= 0:
                raise ValueError(f"Line {i} ({item['Item Name']}): enter the invoice and received quantity.")
            if dmg_q > rec_q:
                raise ValueError(f"Line {i} ({item['Item Name']}): damaged qty cannot exceed received qty.")
            acc_q = round(rec_q - dmg_q, 4)
            short_q = round(max(inv_q - rec_q, 0), 4)
            excess_q = round(max(rec_q - inv_q, 0), 4) if inv_q > 0 else 0
            bin_ = s(l.get("bin"))
            if acc_q > 0 and not bin_:
                raise ValueError(f"Line {i} ({item['Item Name']}): enter the bin / rack where it is kept.")
            qc = s(l.get("qc_status") or header_qc).upper()
            if qc not in ("PENDING", "PASSED"):
                raise ValueError(f"Line {i}: QC status must be PENDING or PASSED.")
            zone = "QC_HOLD" if qc == "PENDING" else ("FG_STORE" if item["Item Type"] == "FG" else "RAW_STORE")
            serials = self.check_new_serials(item, parse_serials(l.get("serials")), acc_q)
            dserials = parse_serials(l.get("damaged_serials"))
            if dmg_q > 0:
                dserials = self.check_new_serials(item, dserials, dmg_q, "damaged")
            price = num(l.get("unit_price"), 0)
            lot = s(l.get("lot_no"))
            mfg = valid_date(l.get("mfg_date"), f"Line {i} mfg date")
            line_id = f"{grn}-{i:02d}"
            batch = dbatch = None
            if acc_q > 0:
                batch = self.add_batch(item, wh, zone, bin_, acc_q, grn_date, "INWARD", qc_status=qc, grn=grn,
                                       inward_type=itype, lot=lot, unit_cost=price or "", currency=currency)
                self.new_serials(item, serials, batch)
                self.log(type="INWARD", ref=grn, item_code=item["Item Code"], item_name=item["Item Name"], qty=acc_q,
                         unit=item["Unit"], to_wh=wh, to_zone=zone, bin=bin_, serials=[x["serial"] for x in serials],
                         party=supplier, doc=inv_no, message=f"{itype} inward",
                         remarks=(f"Short {clean(short_q)}" if short_q else "") + (f" Excess {clean(excess_q)}" if excess_q else ""))
            if dmg_q > 0:
                dbatch = self.add_batch(item, wh, "QUARANTINE", s(l.get("damaged_bin")) or "QUARANTINE", dmg_q, grn_date,
                                        "INWARD", qc_status="DAMAGED", grn=grn, inward_type=itype, lot=lot,
                                        unit_cost=price or "", currency=currency)
                self.new_serials(item, dserials, dbatch)
                self.log(type="INWARD_DAMAGED", ref=grn, item_code=item["Item Code"], item_name=item["Item Name"],
                         qty=dmg_q, unit=item["Unit"], to_wh=wh, to_zone="QUARANTINE", bin=dbatch["Bin"],
                         serials=[x["serial"] for x in dserials], party=supplier, doc=inv_no,
                         message="Received damaged", remarks=s(l.get("remarks")))
            self.t[LINES].append({
                "Line ID": line_id, "GRN No": grn, "Item Code": item["Item Code"], "Item Name": item["Item Name"],
                "Item Type": item["Item Type"], "HSN Code": s(l.get("hsn_code")) or item["HSN Code"],
                "Invoice Qty": inv_q, "Received Qty": rec_q, "Short Qty": short_q, "Excess Qty": excess_q,
                "Damaged Qty": dmg_q, "Accepted Qty": acc_q, "Unit": item["Unit"], "Unit Price": price or "",
                "Currency": currency, "Lot No": lot, "Mfg Date": mfg, "Zone": zone if acc_q > 0 else "",
                "Bin": bin_, "QC Status": qc if acc_q > 0 else "", "Batch ID": batch["Batch ID"] if batch else "",
                "Damaged Batch ID": dbatch["Batch ID"] if dbatch else "", "Serial Count": len(serials) + len(dserials),
                "Remarks": s(l.get("remarks"))})
            received_by_code[item["Item Code"]] = received_by_code.get(item["Item Code"], 0) + acc_q
            tot_acc += acc_q
            tot_short += short_q
            tot_dmg += dmg_q
            line_results.append({"line_id": line_id, "item": item["Item Name"], "accepted": clean(acc_q),
                                 "short": clean(short_q), "damaged": clean(dmg_q), "zone": zone})

        kit = None
        kit_for = s(d.get("kit_for_item"))
        kits = num(d.get("kits_count"), 0)
        if kit_for and kits > 0:
            kit = self.kit_check(kit_for, kits, received_by_code)
        elif itype in ("SKD", "CKD"):
            self.warnings.append("No 'kit for' model / kit count entered, so kit completeness was not checked.")
        any_pending = any(r["zone"] == "QC_HOLD" for r in line_results)
        self.t[INWARDS].append({
            "GRN No": grn, "Inward Type": itype, "GRN Date": grn_date, "Warehouse": wh, "Supplier": supplier,
            "Country of Origin": s(d.get("country_of_origin")), "PO No": s(d.get("po_no")), "Invoice No": inv_no,
            "Invoice Date": valid_date(d.get("invoice_date"), "Invoice date"), "Currency": currency,
            "Exchange Rate": num(d.get("exchange_rate"), 0) or "", "Invoice Value": num(d.get("invoice_value"), 0) or "",
            "BL / AWB No": s(d.get("bl_awb_no")), "Container No": s(d.get("container_no")),
            "Bill of Entry No": s(d.get("bill_of_entry_no")), "BoE Date": valid_date(d.get("boe_date"), "BoE date"),
            "Port of Entry": s(d.get("port_of_entry")), "Vehicle No": s(d.get("vehicle_no")),
            "Gate Entry No": s(d.get("gate_entry_no")),
            "Kit For Item": (self.find_item(kit_for) or {}).get("Item Name", kit_for) if kit_for else "",
            "Kits Count": kits or "", "Kit Check": kit["status"] if kit else "",
            "QC Status": "PENDING" if any_pending else "PASSED", "Total Lines": len(line_results),
            "Total Accepted": tot_acc, "Total Short": tot_short, "Total Damaged": tot_dmg,
            "Received By": s(d.get("received_by")) or self.txn["user"], "Status": "Received",
            "Remarks": s(d.get("remarks")), "Source Document": s(d.get("source_document")), "Txn ID": txn,
            "Created At": now_iso()})
        msg = f"{grn} saved: {itype} from {supplier}, {len(line_results)} line(s), {clean(tot_acc)} accepted"
        if tot_short:
            msg += f", {clean(tot_short)} short"
        if tot_dmg:
            msg += f", {clean(tot_dmg)} damaged (to quarantine)"
        if any_pending:
            msg += ". Stock is in QC Hold until QC is recorded"
        return {"grn_no": grn, "lines": line_results, "kit_check": kit, "message": msg + "."}

    # ======================================================================
    # QC DECISION on a QC_HOLD batch
    # ======================================================================
    def qc_decision(self, d: dict) -> dict:
        b = self.t[BATCHES].get(d.get("batch_id"))
        if not b or b["Zone"] != "QC_HOLD" or num(b["Qty Remaining"]) <= 0:
            raise ValueError("That batch is not waiting in QC Hold.")
        item = self.item(b["Item Code"])
        remaining = num(b["Qty Remaining"])
        pass_q = num(d.get("pass_qty"), 0)
        rej_q = num(d.get("reject_qty"), 0)
        if pass_q < 0 or rej_q < 0 or pass_q + rej_q <= 0:
            raise ValueError("Enter how many passed and/or how many were rejected.")
        if pass_q + rej_q > remaining + 1e-9:
            raise ValueError(f"Only {clean(remaining)} {item['Unit']} are waiting for QC in this batch.")
        wh = b["Warehouse"]
        store = zone_ok(d.get("pass_zone") or ("FG_STORE" if item["Item Type"] == "FG" else "RAW_STORE"),
                        ["RAW_STORE", "FG_STORE"], "Pass zone")
        pass_bin = s(d.get("pass_bin")) or b["Bin"]
        rej_bin = s(d.get("reject_bin")) or "QUARANTINE"
        remarks = s(d.get("remarks"))
        ref = b["GRN No"] or b["Work Order"] or b["Batch ID"]
        if item["Serialized"] == "Y":
            in_batch = [r["Serial No"] for r in self.t[SERIALS].rows()
                        if r["Batch ID"] == b["Batch ID"] and r["Status"] == "In Stock"]
            rej_s = [x["serial"] for x in parse_serials(d.get("reject_serials"))]
            pass_s = [x["serial"] for x in parse_serials(d.get("pass_serials"))]
            if rej_q and len(rej_s) != int(rej_q):
                raise ValueError(f"List the {clean(rej_q)} rejected serial number(s).")
            if pass_q and not pass_s:
                rest = [x for x in in_batch if x.upper() not in {r.upper() for r in rej_s}]
                if int(pass_q) == len(rest):
                    pass_s = rest
                else:
                    raise ValueError(f"List the {clean(pass_q)} passed serial number(s) (only part of the batch is being passed).")
            if pass_q and len(pass_s) != int(pass_q):
                raise ValueError(f"Pass qty is {clean(pass_q)} but {len(pass_s)} serial(s) listed.")
        else:
            rej_s, pass_s, in_batch = [], [], []
        out = []
        in_batch_set = {x.upper() for x in in_batch} if item["Serialized"] == "Y" else set()
        for x in pass_s + rej_s:
            if x.upper() not in in_batch_set:
                raise ValueError(f"Serial {x} is not in batch {b['Batch ID']} waiting for QC.")
        if set(x.upper() for x in pass_s) & set(x.upper() for x in rej_s):
            raise ValueError("A serial cannot be both passed and rejected.")
        for q, sers, zone, bin_, kind, qcs in ((pass_q, pass_s, store, pass_bin, "QC_PASS", "PASSED"),
                                               (rej_q, rej_s, "QUARANTINE", rej_bin, "QC_REJECT", "REJECTED")):
            if not q:
                continue
            cur = self.t[BATCHES].get(b["Batch ID"])
            self._decrement(cur, q)
            allocs = [{"batch": dict(cur), "qty": q, "serials": [self.t[SERIALS].get(serial_key(item["Item Code"], x))["Serial No"] for x in sers]}]
            self.put(item, allocs, wh, zone, bin_, kind, wo=b["Work Order"], qc_status=qcs)
            self.log(type=kind, ref=ref, item_code=item["Item Code"], item_name=item["Item Name"], qty=q,
                     unit=item["Unit"], from_wh=wh, from_zone="QC_HOLD", to_wh=wh, to_zone=zone, bin=bin_,
                     serials=allocs[0]["serials"], remarks=remarks)
            out.append(f"{clean(q)} {'passed to ' + zone if kind == 'QC_PASS' else 'rejected to QUARANTINE'}")
        # update GRN line / header QC status
        for ln in self.t[LINES].rows():
            if ln["Batch ID"] == b["Batch ID"]:
                left = num(self.t[BATCHES].get(b["Batch ID"])["Qty Remaining"])
                passed = sum(num(r["Qty"]) for r in self.t[LOG].rows() if r["Txn Type"] == "QC_PASS"
                             and r["Ref No"] == ln["GRN No"] and r["Item Code"] == ln["Item Code"] and r["Status"] == "Applied")
                status = "PENDING" if left > 0 else ("PASSED" if passed >= num(ln["Accepted Qty"]) - 1e-9 else
                                                     "REJECTED" if passed == 0 else "PARTIAL")
                self.t[LINES].update(ln["Line ID"], {"QC Status": status})
                grn = ln["GRN No"]
                st = [x["QC Status"] for x in self.t[LINES].rows() if x["GRN No"] == grn and x["QC Status"]]
                hdr = "PENDING" if "PENDING" in st else ("PASSED" if all(x == "PASSED" for x in st) else "PARTIAL")
                self.t[INWARDS].update(grn, {"QC Status": hdr})
        return {"message": f"QC recorded for {item['Item Name']} ({ref}): " + ", ".join(out) + "."}

    # ======================================================================
    # OUTWARD: sale / sample / return-to-vendor / scrap
    # ======================================================================
    def dispatch(self, d: dict) -> dict:
        otype = s(d.get("outward_type") or "SALE").upper()
        if otype not in OUTWARD_TYPES:
            raise ValueError(f"Outward type must be one of: {', '.join(OUTWARD_TYPES)}.")
        cfg = OUTWARD_TYPES[otype]
        wh = warehouse_ok(d.get("warehouse"))
        party = s(d.get("party"))
        doc = s(d.get("doc_no"))
        if otype in ("SALE", "SAMPLE", "RTV") and not party:
            raise ValueError("Customer / party name is required.")
        if otype in ("SALE", "RTV") and not doc:
            raise ValueError("Invoice / delivery challan number is required.")
        ddate = valid_date(d.get("date") or today(), "Dispatch date")
        lines = [l for l in (d.get("lines") or []) if s(l.get("item")) or s(l.get("item_code"))]
        if not lines:
            raise ValueError("Add at least one item to dispatch.")
        msgs = []
        for l in lines:
            item = self.item(l.get("item_code") or l.get("item"))
            zones = cfg["zones"]
            if s(l.get("zone")):
                zones = [zone_ok(l.get("zone"), cfg["zones"], f"Zone for {otype}")]
            serials = parse_serials(l.get("serials"))
            qty = num(l.get("qty"), len(serials))
            allocs = self.take(item, qty, serials, wh, zones)
            sn = [x for a in allocs for x in a["serials"]]
            wmonths = int(num(item["Warranty Months"], 0))
            for x in sn:
                ch = {"Status": cfg["serial_status"], "Customer": party, "Dispatch Doc": doc, "Dispatch Date": ddate,
                      "Zone": "", "Bin": "", "Updated At": now_iso()}
                if otype == "SALE" and wmonths:
                    ch["Warranty Until"] = add_months(parse_d(ddate), wmonths).isoformat()
                self.t[SERIALS].update(serial_key(item["Item Code"], x), ch)
            fz = sorted({a["batch"]["Zone"] for a in allocs})
            self.log(type=otype if otype != "SALE" else "DISPATCH", ref=doc, item_code=item["Item Code"],
                     item_name=item["Item Name"], qty=qty, unit=item["Unit"], from_wh=wh, from_zone="/".join(fz),
                     serials=sn, party=party, doc=doc, remarks=s(d.get("remarks")) or s(d.get("address")),
                     message=cfg["label"])
            oldest = allocs[0]["batch"]["Date Received"] if allocs else ""
            msgs.append(f"{clean(qty)} {item['Unit']} {item['Item Name']} (oldest stock from {oldest})")
        return {"message": f"{cfg['label']} {doc or ''} to {party or '-'}: " + "; ".join(msgs) + "."}

    # ======================================================================
    # TRANSFER between warehouses / zones / bins
    # ======================================================================
    def transfer(self, d: dict) -> dict:
        item = self.item(d.get("item_code") or d.get("item"))
        fw = warehouse_ok(d.get("from_warehouse"), "From warehouse")
        tw = warehouse_ok(d.get("to_warehouse") or fw, "To warehouse")
        fz = zone_ok(d.get("from_zone") or ("FG_STORE" if item["Item Type"] == "FG" else "RAW_STORE"), MOVABLE_ZONES, "From zone")
        tz = zone_ok(d.get("to_zone") or fz, MOVABLE_ZONES, "To zone")
        tbin = s(d.get("to_bin"))
        if not tbin:
            raise ValueError("Enter the destination bin / rack.")
        serials = parse_serials(d.get("serials"))
        qty = num(d.get("qty"), len(serials))
        allocs = self.take(item, qty, serials, fw, [fz])
        if fw == tw and fz == tz and all(s(a["batch"]["Bin"]) == tbin for a in allocs):
            raise ValueError("From and to location are the same.")
        self.put(item, allocs, tw, tz, tbin, "TRANSFER")
        sn = [x for a in allocs for x in a["serials"]]
        self.log(type="TRANSFER", ref=s(d.get("doc_no")), item_code=item["Item Code"], item_name=item["Item Name"],
                 qty=qty, unit=item["Unit"], from_wh=fw, from_zone=fz, to_wh=tw, to_zone=tz, bin=tbin, serials=sn,
                 doc=s(d.get("doc_no")), remarks=s(d.get("remarks")) or s(d.get("vehicle_no")))
        return {"message": f"Moved {clean(qty)} {item['Unit']} {item['Item Name']} from {fw}/{fz} to {tw}/{tz} bin {tbin} "
                           f"(original receipt dates kept for aging)."}

    # ======================================================================
    # ADJUSTMENT (stock count correction, found, lost)
    # ======================================================================
    def adjust(self, d: dict) -> dict:
        item = self.item(d.get("item_code") or d.get("item"))
        wh = warehouse_ok(d.get("warehouse"))
        zone = zone_ok(d.get("zone") or ("FG_STORE" if item["Item Type"] == "FG" else "RAW_STORE"),
                       ["RAW_STORE", "FG_STORE", "QUARANTINE", "QC_HOLD"])
        qty = num(d.get("qty"))
        reason = s(d.get("reason"))
        if not reason:
            raise ValueError("A reason is required for every adjustment.")
        if qty == 0:
            raise ValueError("Adjustment quantity cannot be 0 (use + to add, - to remove).")
        serials = parse_serials(d.get("serials"))
        if qty > 0:
            bin_ = s(d.get("bin"))
            if not bin_:
                raise ValueError("Enter the bin / rack where the found stock is kept.")
            serials = self.check_new_serials(item, serials, qty, "adjustment")
            b = self.add_batch(item, wh, zone, bin_, qty, today(), "ADJUST",
                               qc_status="PASSED" if zone in ("RAW_STORE", "FG_STORE") else "")
            self.new_serials(item, serials, b)
            self.log(type="ADJUST_IN", item_code=item["Item Code"], item_name=item["Item Name"], qty=qty, unit=item["Unit"],
                     to_wh=wh, to_zone=zone, bin=bin_, serials=[x["serial"] for x in serials], remarks=reason)
            return {"message": f"Added {clean(qty)} {item['Unit']} {item['Item Name']} at {wh}/{zone} ({reason})."}
        allocs = self.take(item, -qty, serials, wh, [zone])
        sn = [x for a in allocs for x in a["serials"]]
        for x in sn:
            self.t[SERIALS].update(serial_key(item["Item Code"], x), {"Status": "Written Off", "Zone": "", "Bin": "",
                                                                       "Updated At": now_iso()})
        self.log(type="ADJUST_OUT", item_code=item["Item Code"], item_name=item["Item Name"], qty=qty, unit=item["Unit"],
                 from_wh=wh, from_zone=zone, serials=sn, remarks=reason)
        return {"message": f"Removed {clean(-qty)} {item['Unit']} {item['Item Name']} from {wh}/{zone} ({reason})."}

    # ======================================================================
    # SALES RETURN (customer sends units back)
    # ======================================================================
    def sales_return(self, d: dict) -> dict:
        wh = warehouse_ok(d.get("warehouse"))
        party = s(d.get("party"))
        doc = s(d.get("doc_no"))
        if not party or not doc:
            raise ValueError("Customer and return document (credit note / RMA no.) are required.")
        zone = zone_ok(d.get("zone") or "QUARANTINE", ["QUARANTINE", "QC_HOLD"])
        bin_ = s(d.get("bin")) or zone
        rdate = valid_date(d.get("date") or today(), "Return date")
        msgs = []
        for l in [l for l in (d.get("lines") or []) if s(l.get("item")) or s(l.get("item_code"))]:
            item = self.item(l.get("item_code") or l.get("item"))
            serials = parse_serials(l.get("serials"))
            qty = num(l.get("qty"), len(serials))
            if qty <= 0:
                raise ValueError(f"Enter the quantity returned for {item['Item Name']}.")
            b = self.add_batch(item, wh, zone, bin_, qty, rdate, "SALES_RETURN", qc_status="PENDING" if zone == "QC_HOLD" else "RETURNED")
            if item["Serialized"] == "Y":
                if len(serials) != int(qty):
                    raise ValueError(f"{item['Item Name']} is serialized: list the {clean(qty)} returned serial(s).")
                for x in serials:
                    k = serial_key(item["Item Code"], x["serial"])
                    rec = self.t[SERIALS].get(k)
                    if not rec:
                        raise ValueError(f"Serial {x['serial']} was never received - cannot accept it as a return.")
                    if rec["Status"] not in ("Dispatched",):
                        raise ValueError(f"Serial {x['serial']} is '{rec['Status']}', not dispatched.")
                    self.t[SERIALS].update(k, {"Status": "In Stock", "Batch ID": b["Batch ID"], "Warehouse": wh,
                                               "Zone": zone, "Bin": bin_, "QC Status": b["QC Status"],
                                               "Updated At": now_iso()})
            self.log(type="SALES_RETURN", ref=doc, item_code=item["Item Code"], item_name=item["Item Name"], qty=qty,
                     unit=item["Unit"], to_wh=wh, to_zone=zone, bin=bin_, serials=[x["serial"] for x in serials],
                     party=party, doc=doc, remarks=s(d.get("remarks")))
            msgs.append(f"{clean(qty)} {item['Item Name']}")
        if not msgs:
            raise ValueError("Add at least one returned item.")
        return {"message": f"Return {doc} from {party} received into {wh}/{zone}: " + ", ".join(msgs) + "."}

    # ======================================================================
    # PRODUCTION: work orders
    # ======================================================================
    def wo(self, wo_no) -> dict:
        w = self.t[WOS].get(s(wo_no).upper())
        if not w or w["Status"] == "Reversed":
            raise ValueError(f"Work order '{wo_no}' not found.")
        return w

    def create_wo(self, d: dict) -> dict:
        fg = self.item(d.get("fg_item"))
        if fg["Item Type"] not in ("FG", "SUB_ASSEMBLY"):
            raise ValueError("A work order must build a finished good or sub-assembly.")
        qty = num(d.get("planned_qty"))
        if qty <= 0 or qty != int(qty):
            raise ValueError("Planned quantity must be a whole number above 0.")
        wh = warehouse_ok(d.get("warehouse"))
        btype = s(d.get("build_type") or "CKD").upper()
        if btype not in ("SKD", "CKD", "OTHER"):
            raise ValueError("Build type must be SKD, CKD or OTHER.")
        if not self.bom_for(fg["Item Code"]):
            self.warnings.append(f"{fg['Item Name']} has no BOM - you will have to enter consumed materials by hand at output.")
        no = self.next_id(f"wo{datetime.now(IST).year}", "WO-{y}-{n:04d}")
        self.t[WOS].append({"WO No": no, "FG Item Code": fg["Item Code"], "FG Item Name": fg["Item Name"],
                            "Build Type": btype, "Planned Qty": qty, "Produced Qty": 0, "Rejected Qty": 0,
                            "Production Line": s(d.get("production_line")), "Warehouse": wh, "Status": "OPEN",
                            "Start Date": valid_date(d.get("start_date") or today(), "Start date"),
                            "Target Date": valid_date(d.get("target_date"), "Target date"), "Completed Date": "",
                            "Source GRN": s(d.get("source_grn")), "Created By": self.txn["user"],
                            "Remarks": s(d.get("remarks")), "Txn ID": self.txn["id"], "Created At": now_iso()})
        self.log(type="WO_CREATE", ref=no, item_code=fg["Item Code"], item_name=fg["Item Name"], qty=qty,
                 unit=fg["Unit"], to_wh=wh, message=f"Work order for {clean(qty)} {fg['Item Name']}")
        return {"wo_no": no, "message": f"{no} created to build {clean(qty)} {fg['Item Name']} at {wh}."}

    def wo_materials(self, w: dict) -> List[dict]:
        no = w["WO No"]
        planned = num(w["Planned Qty"])
        issued: Dict[str, float] = {}
        for r in self.t[LOG].rows():
            if r["Ref No"] == no and r["Status"] == "Applied":
                if r["Txn Type"] == "WO_ISSUE":
                    issued[r["Item Code"]] = issued.get(r["Item Code"], 0) + num(r["Qty"])
                elif r["Txn Type"] == "WO_RETURN":
                    issued[r["Item Code"]] = issued.get(r["Item Code"], 0) - num(r["Qty"])
        consumed: Dict[str, float] = {}
        for c in self.t[CONS].rows():
            if c["WO No"] == no:
                consumed[c["Component Code"]] = consumed.get(c["Component Code"], 0) + num(c["Qty"])
        rows, seen = [], set()
        for b in self.bom_for(w["FG Item Code"]):
            code = b["Component Code"]
            seen.add(code)
            req = num(b["Qty Per Unit"]) * planned
            iss = issued.get(code, 0)
            rows.append({"item_code": code, "item_name": b["Component Name"], "unit": b["Unit"],
                         "qty_per_unit": clean(num(b["Qty Per Unit"])), "required": clean(req), "issued": clean(round(iss, 4)),
                         "consumed": clean(round(consumed.get(code, 0), 4)),
                         "in_wip": clean(self.on_hand(code, w["Warehouse"], ["WIP"], no)),
                         "to_issue": clean(round(max(req - iss, 0), 4)),
                         "in_store": clean(self.on_hand(code, w["Warehouse"], ["RAW_STORE"]))})
        for code in set(issued) | set(consumed):
            if code in seen:
                continue
            it = self.find_item(code) or {"Item Name": code, "Unit": ""}
            rows.append({"item_code": code, "item_name": it["Item Name"], "unit": it["Unit"], "qty_per_unit": "",
                         "required": "", "issued": clean(round(issued.get(code, 0), 4)),
                         "consumed": clean(round(consumed.get(code, 0), 4)),
                         "in_wip": clean(self.on_hand(code, w["Warehouse"], ["WIP"], no)), "to_issue": 0,
                         "in_store": clean(self.on_hand(code, w["Warehouse"], ["RAW_STORE"])), "not_in_bom": True})
        return rows

    def wo_issue(self, d: dict) -> dict:
        w = self.wo(d.get("wo_no"))
        if w["Status"] not in ("OPEN", "IN_PROGRESS"):
            raise ValueError(f"{w['WO No']} is {w['Status']}.")
        wh = w["Warehouse"]
        lines = [l for l in (d.get("lines") or []) if s(l.get("item")) or s(l.get("item_code"))]
        auto = bool(d.get("issue_per_bom"))
        if auto:
            mats = self.wo_materials(w)
            lines = [{"item_code": m["item_code"], "qty": m["to_issue"]} for m in mats if num(m["to_issue"]) > 0]
            if not lines:
                raise ValueError("Everything the BOM needs for this work order has already been issued.")
            # check everything first so we don't half-issue
            short = []
            for m in mats:
                if num(m["to_issue"]) > num(m["in_store"]) + 1e-9:
                    short.append(f"{m['item_name']} (need {m['to_issue']}, store has {m['in_store']})")
            if short and not d.get("allow_partial"):
                raise ValueError("Store is short for: " + "; ".join(short[:12]) + ". Issue what is available line by line, or receive the shortage first.")
        if not lines:
            raise ValueError("Add at least one item to issue.")
        msgs = []
        for l in lines:
            item = self.item(l.get("item_code") or l.get("item"))
            serials = parse_serials(l.get("serials"))
            qty = num(l.get("qty"), len(serials))
            if auto and d.get("allow_partial"):
                qty = min(qty, self.on_hand(item["Item Code"], wh, ["RAW_STORE"]))
                if qty <= 0:
                    continue
            allocs = self.take(item, qty, serials, wh, ["RAW_STORE"], auto_pick=auto)
            self.put(item, allocs, wh, "WIP", s(l.get("line")) or w["Production Line"] or "LINE", "WO_ISSUE", wo=w["WO No"])
            sn = [x for a in allocs for x in a["serials"]]
            self.log(type="WO_ISSUE", ref=w["WO No"], item_code=item["Item Code"], item_name=item["Item Name"], qty=qty,
                     unit=item["Unit"], from_wh=wh, from_zone="RAW_STORE", to_wh=wh, to_zone="WIP",
                     bin=w["Production Line"], serials=sn)
            msgs.append(f"{clean(qty)} {item['Unit']} {item['Item Name']}" + (f" (serials {', '.join(sn[:5])}{'…' if len(sn) > 5 else ''})" if sn else ""))
        if w["Status"] == "OPEN":
            self.t[WOS].update(w["WO No"], {"Status": "IN_PROGRESS"})
        return {"message": f"Issued to {w['WO No']}: " + "; ".join(msgs) + "."}

    def wo_output(self, d: dict) -> dict:
        w = self.wo(d.get("wo_no"))
        if w["Status"] not in ("OPEN", "IN_PROGRESS"):
            raise ValueError(f"{w['WO No']} is {w['Status']}.")
        fg = self.item(w["FG Item Code"])
        wh = w["Warehouse"]
        qty = num(d.get("qty"))
        if qty <= 0 or qty != int(qty):
            raise ValueError("Output quantity must be a whole number above 0.")
        zone = zone_ok(d.get("zone") or "FG_STORE", ["FG_STORE", "QC_HOLD", "QUARANTINE"], "Output zone")
        done = num(w["Produced Qty"]) + num(w["Rejected Qty"])
        if done + qty > num(w["Planned Qty"]):
            raise ValueError(f"{w['WO No']} was planned for {clean(num(w['Planned Qty']))}; {clean(done)} already reported. "
                             f"Reporting {clean(qty)} more would exceed the plan.")
        bin_ = s(d.get("bin"))
        if not bin_:
            raise ValueError("Enter the bin / rack where the finished units are kept.")
        odate = valid_date(d.get("date") or today(), "Output date")
        fg_serials = self.check_new_serials(fg, parse_serials(d.get("fg_serials")), qty, "output")
        # component serial mapping: "FGSERIAL: comp1, comp2" per line
        cmap: Dict[str, List[str]] = {}
        raw_map = d.get("component_map") or ""
        if isinstance(raw_map, dict):
            cmap = {s(k).upper(): [s(x) for x in v] for k, v in raw_map.items()}
        else:
            for line in str(raw_map).splitlines():
                if ":" in line:
                    k, _, v = line.partition(":")
                    cmap[k.strip().upper()] = [x.strip() for x in re.split(r"[,;\s]+", v) if x.strip()]
        fg_names = {x["serial"].upper() for x in fg_serials}
        for k in cmap:
            if k not in fg_names:
                raise ValueError(f"Component map refers to FG serial {k}, which is not in this output.")
        # what to consume
        bom = self.bom_for(fg["Item Code"])
        manual = [l for l in (d.get("consumption") or []) if s(l.get("item")) or s(l.get("item_code"))]
        if manual:
            need = []
            for l in manual:
                sers = parse_serials(l.get("serials"))
                q = num(l.get("qty"), 0) or len(sers)
                if q <= 0:
                    raise ValueError(f"Enter the quantity consumed for {s(l.get('item') or l.get('item_code'))}.")
                need.append((self.item(l.get("item_code") or l.get("item")), q, sers))
        elif bom:
            need = [(self.item(b["Component Code"]), round(num(b["Qty Per Unit"]) * qty, 4), []) for b in bom]
        else:
            raise ValueError(f"{fg['Item Name']} has no BOM. Enter the materials consumed for this output.")
        # check WIP availability for all components before touching anything
        short = []
        for item, q, _ in need:
            have = self.on_hand(item["Item Code"], wh, ["WIP"], w["WO No"])
            if have + 1e-9 < q:
                short.append(f"{item['Item Name']} (need {clean(q)}, issued & unused {clean(have)})")
        if short:
            raise ValueError("Not enough material on the line for this output: " + "; ".join(short[:12]) +
                             ". Issue materials to the work order first.")
        out_batch = self.add_batch(fg, wh, zone, bin_, qty, odate, "PRODUCTION", wo=w["WO No"], inward_type=w["Build Type"],
                                   qc_status={"FG_STORE": "PASSED", "QC_HOLD": "PENDING", "QUARANTINE": "REJECTED"}[zone])
        self.new_serials(fg, fg_serials, out_batch, wo=w["WO No"])
        fg_list = [x["serial"] for x in fg_serials]
        comp_parent: Dict[str, str] = {}
        for fs, comps in cmap.items():
            for c in comps:
                comp_parent[c.upper()] = next(x for x in fg_list if x.upper() == fs)
        for item, q, sers in need:
            if item["Serialized"] == "Y":
                mapped = [c for c in comp_parent if self.t[SERIALS].get(serial_key(item["Item Code"], c))]
                chosen = [{"serial": self.t[SERIALS].get(serial_key(item["Item Code"], c))["Serial No"]} for c in mapped]
                chosen += [x for x in sers if x["serial"].upper() not in {c.upper() for c in mapped}]
                if len(chosen) > int(q):
                    raise ValueError(f"More {item['Item Name']} serials mapped ({len(chosen)}) than consumed ({clean(q)}).")
                if len(chosen) < int(q):
                    chosen += self.pick_serials(item, int(q) - len(chosen), wh, ["WIP"], w["WO No"],
                                                exclude=[c["serial"] for c in chosen])
                allocs = self.take_serials(item, chosen, wh, ["WIP"], w["WO No"])
            else:
                allocs = self.take_qty(item, q, wh, ["WIP"], w["WO No"])
            for a in allocs:
                for sn in a["serials"]:
                    self.t[SERIALS].update(serial_key(item["Item Code"], sn), {
                        "Status": "Consumed", "Parent Serial": comp_parent.get(sn.upper(), ""), "Zone": "", "Bin": "",
                        "Updated At": now_iso()})
                self.t[CONS].append({"Cons ID": self.next_id("cons", "C{n:07d}"), "Txn ID": self.txn["id"], "Date": odate,
                                     "WO No": w["WO No"], "FG Item Code": fg["Item Code"], "FG Serials": ", ".join(fg_list),
                                     "Component Code": item["Item Code"], "Component Name": item["Item Name"],
                                     "Qty": a["qty"], "Unit": item["Unit"], "From Batch ID": a["batch"]["Batch ID"],
                                     "GRN No": a["batch"]["GRN No"], "Component Serials": ", ".join(a["serials"])})
            self.log(type="WO_CONSUME", ref=w["WO No"], item_code=item["Item Code"], item_name=item["Item Name"], qty=q,
                     unit=item["Unit"], from_wh=wh, from_zone="WIP", serials=[x for a in allocs for x in a["serials"]])
        self.log(type="WO_OUTPUT", ref=w["WO No"], item_code=fg["Item Code"], item_name=fg["Item Name"], qty=qty,
                 unit=fg["Unit"], to_wh=wh, to_zone=zone, bin=bin_, serials=fg_list, remarks=s(d.get("remarks")))
        ch = {"Rejected Qty": num(w["Rejected Qty"]) + qty} if zone == "QUARANTINE" else {"Produced Qty": num(w["Produced Qty"]) + qty}
        ch["Status"] = "IN_PROGRESS"
        self.t[WOS].update(w["WO No"], ch)
        return {"message": f"{w['WO No']}: {clean(qty)} {fg['Item Name']} reported to {zone}"
                           + (f" (serials {fg_list[0]}…{fg_list[-1]})" if fg_list else "") +
                           f"; {len(need)} component(s) consumed from the line."}

    def wo_return(self, d: dict) -> dict:
        w = self.wo(d.get("wo_no"))
        wh = w["Warehouse"]
        zone = zone_ok(d.get("zone") or "RAW_STORE", ["RAW_STORE", "QUARANTINE"], "Return to zone")
        bin_ = s(d.get("bin"))
        if not bin_:
            raise ValueError("Enter the bin / rack the material goes back to.")
        lines = [l for l in (d.get("lines") or []) if s(l.get("item")) or s(l.get("item_code"))]
        if d.get("return_all"):
            codes = sorted({b["Item Code"] for b in self.t[BATCHES].rows()
                            if b["Zone"] == "WIP" and s(b["Work Order"]) == w["WO No"] and num(b["Qty Remaining"]) > 0})
            lines = [{"item_code": c, "qty": self.on_hand(c, wh, ["WIP"], w["WO No"]), "_all": True} for c in codes]
        if not lines:
            raise ValueError("Nothing to return.")
        msgs = []
        for l in lines:
            item = self.item(l.get("item_code") or l.get("item"))
            serials = parse_serials(l.get("serials"))
            qty = num(l.get("qty"), len(serials))
            allocs = self.take(item, qty, serials, wh, ["WIP"], w["WO No"], auto_pick=bool(l.get("_all")))
            self.put(item, allocs, wh, zone, bin_, "WO_RETURN", wo="", qc_status="REJECTED" if zone == "QUARANTINE" else None)
            sn = [x for a in allocs for x in a["serials"]]
            for x in sn:
                self.t[SERIALS].update(serial_key(item["Item Code"], x), {"Work Order": ""})
            self.log(type="WO_RETURN", ref=w["WO No"], item_code=item["Item Code"], item_name=item["Item Name"], qty=qty,
                     unit=item["Unit"], from_wh=wh, from_zone="WIP", to_wh=wh, to_zone=zone, bin=bin_, serials=sn,
                     remarks=s(d.get("remarks")))
            msgs.append(f"{clean(qty)} {item['Item Name']}")
        return {"message": f"Returned from {w['WO No']} to {zone}: " + ", ".join(msgs) + "."}

    def wo_close(self, d: dict) -> dict:
        w = self.wo(d.get("wo_no"))
        if w["Status"] == "CLOSED":
            raise ValueError(f"{w['WO No']} is already closed.")
        left = [b for b in self.t[BATCHES].rows() if b["Zone"] == "WIP" and s(b["Work Order"]) == w["WO No"]
                and num(b["Qty Remaining"]) > 0]
        if left:
            if not s(d.get("return_bin")):
                raise ValueError(f"{len(left)} batch(es) of unused material are still on the line for {w['WO No']}. "
                                 f"Enter a return bin to send them back to the store, then close.")
            self.wo_return({"wo_no": w["WO No"], "return_all": True, "bin": d.get("return_bin"), "zone": "RAW_STORE"})
        self.t[WOS].update(w["WO No"], {"Status": "CLOSED", "Completed Date": today()})
        self.log(type="WO_CLOSE", ref=w["WO No"], item_code=w["FG Item Code"], item_name=w["FG Item Name"],
                 qty=num(w["Produced Qty"]), message=f"Closed: {clean(num(w['Produced Qty']))} good, {clean(num(w['Rejected Qty']))} rejected")
        return {"message": f"{w['WO No']} closed ({clean(num(w['Produced Qty']))} produced, {clean(num(w['Rejected Qty']))} rejected)."}

    # ======================================================================
    # REPORTS / QUERIES
    # ======================================================================
    def stock_summary(self, item_filter=None) -> List[dict]:
        out: Dict[str, dict] = {}
        for b in self.t[BATCHES].rows():
            q = num(b["Qty Remaining"])
            if q <= 0:
                continue
            if item_filter and b["Item Code"] != item_filter:
                continue
            e = out.setdefault(b["Item Code"], {"item_code": b["Item Code"], "item_name": b["Item Name"],
                                                "item_type": b["Item Type"], "unit": b["Unit"], "total": 0,
                                                "available": 0, "by_zone": {}, "by_location": {}})
            e["total"] += q
            if b["Zone"] in ("RAW_STORE", "FG_STORE"):
                e["available"] += q
            e["by_zone"][b["Zone"]] = e["by_zone"].get(b["Zone"], 0) + q
            loc = f"{b['Warehouse']} / {b['Zone']} / {b['Bin'] or '-'}"
            e["by_location"][loc] = e["by_location"].get(loc, 0) + q
        items = {i["Item Code"]: i for i in self.t[ITEMS].rows()}
        res = []
        for code, e in sorted(out.items(), key=lambda kv: kv[1]["item_name"].lower()):
            it = items.get(code, {})
            e["serialized"] = it.get("Serialized", "N")
            e["reorder_level"] = clean(num(it.get("Reorder Level"), 0))
            e["low_stock"] = bool(e["reorder_level"]) and e["available"] < e["reorder_level"]
            e["total"], e["available"] = clean(round(e["total"], 4)), clean(round(e["available"], 4))
            e["by_zone"] = {k: clean(round(v, 4)) for k, v in e["by_zone"].items()}
            e["locations"] = [{"location": k, "quantity": clean(round(v, 4))} for k, v in sorted(e["by_location"].items())]
            del e["by_location"]
            res.append(e)
        return res

    def aging(self, threshold=AGING_THRESHOLD_DAYS) -> List[dict]:
        t = datetime.now(IST).date()
        out = []
        for b in self.t[BATCHES].rows():
            if num(b["Qty Remaining"]) <= 0 or b["Zone"] == "WIP":
                continue
            d = parse_d(b["Date Received"])
            if not d:
                continue
            age = (t - d).days
            if age >= threshold:
                out.append({**b, "Age Days": age})
        return sorted(out, key=lambda r: -r["Age Days"])

    def buildable(self) -> List[dict]:
        fgs = sorted({b["FG Item Code"] for b in self.t[BOM].rows()})
        res = []
        for code in fgs:
            fg = self.find_item(code)
            if not fg:
                continue
            bom = self.bom_for(code)
            for wh in WAREHOUSES:
                best, limiting = None, []
                for b in bom:
                    have = self.on_hand(b["Component Code"], wh, ["RAW_STORE"])
                    can = math.floor(have / num(b["Qty Per Unit"]) + 1e-9)
                    if best is None or can < best:
                        best, limiting = can, [b["Component Name"]]
                    elif can == best:
                        limiting.append(b["Component Name"])
                res.append({"fg_item_code": code, "fg_item_name": fg["Item Name"], "warehouse": wh,
                            "buildable": best or 0, "limited_by": limiting[:3], "components": len(bom)})
        return res

    def trace(self, q: str) -> List[dict]:
        q = s(q).upper()
        if not q:
            return []
        hits = [r for r in self.t[SERIALS].rows() if s(r["Serial No"]).upper() == q or s(r["MAC / IMEI"]).upper() == q]
        out = []
        for r in hits:
            grn = self.t[INWARDS].get(r["GRN No"]) if r["GRN No"] else None
            entry = {"serial": r, "grn": grn, "history": []}
            if r["Parent Serial"]:
                entry["fitted_in"] = [p for p in self.t[SERIALS].rows() if p["Serial No"] == r["Parent Serial"]]
            children = [c for c in self.t[SERIALS].rows() if s(c["Parent Serial"]).upper() == q]
            entry["components_serials"] = children
            if r["Work Order"] and r["Status"] != "Consumed":
                cons = [c for c in self.t[CONS].rows() if c["WO No"] == r["Work Order"]
                        and q in [x.strip().upper() for x in s(c["FG Serials"]).split(",")]]
                entry["built_from"] = [{"component": c["Component Name"], "qty": c["Qty"], "grn": c["GRN No"],
                                        "supplier": (self.t[INWARDS].get(c["GRN No"]) or {}).get("Supplier", ""),
                                        "serials": c["Component Serials"]} for c in cons]
                entry["work_order"] = self.t[WOS].get(r["Work Order"])
            pat = re.compile(r"(^|,\s*)" + re.escape(r["Serial No"]) + r"(\s*,|$)", re.I)
            entry["history"] = [h for h in self.t[LOG].rows() if h["Item Code"] == r["Item Code"] and pat.search(s(h["Serials"]))]
            out.append(entry)
        return out


# ============================================================================
# TRANSACTION RUNNER: load -> apply -> save, retrying on concurrent edits
# ============================================================================
_read_cache = {"sha": None, "book": None}


def read_book() -> Book:
    raw, sha = STORE.load()
    if sha and _read_cache["sha"] == sha:
        return _read_cache["book"]
    book = Book.from_bytes(raw)
    _read_cache.update(sha=sha, book=book)
    return book


def run_txn(txn_type: str, user: str, fn, commit_msg: str, dry_run=False) -> dict:
    for attempt in range(3):
        raw, sha = STORE.load()
        book = Book.from_bytes(raw)
        book.import_legacy_products()
        try:
            txn_id = book.begin(txn_type, user)
            result = fn(book) or {}
            book.commit()
        except ValueError as e:
            return {"success": False, "message": str(e), "warnings": book.warnings}
        result = {"success": True, "txn_id": txn_id, "warnings": book.warnings, **result}
        if dry_run:
            result["dry_run"] = True
            return result
        try:
            STORE.save(book.to_bytes(), sha, f"{commit_msg} [{txn_id}] by {user}"[:200])
            _read_cache.update(sha=None, book=None)
            return result
        except Conflict:
            continue
    raise HTTPException(409, "Someone else saved at the same moment and the retry also clashed. Please try again.")


def run_master(user: str, fn, commit_msg: str) -> dict:
    """Master data edits (items / BOM): saved, but not part of the undo chain."""
    for attempt in range(3):
        raw, sha = STORE.load()
        book = Book.from_bytes(raw)
        book.import_legacy_products()
        try:
            result = fn(book) or {}
        except ValueError as e:
            return {"success": False, "message": str(e)}
        try:
            STORE.save(book.to_bytes(), sha, f"{commit_msg} by {user}"[:200])
            _read_cache.update(sha=None, book=None)
            return {"success": True, "warnings": book.warnings, **result}
        except Conflict:
            continue
    raise HTTPException(409, "Someone else saved at the same moment. Please try again.")


# ============================================================================
# GEMINI
# ============================================================================
def call_gemini(system: str, parts: List[dict], temperature=0.1, timeout=60) -> dict:
    if not GEMINI_API_KEY:
        raise RuntimeError("GEMINI_API_KEY is not set on the server.")
    payload = {"systemInstruction": {"parts": [{"text": system}]},
               "contents": [{"role": "user", "parts": parts}],
               "generationConfig": {"temperature": temperature, "responseMimeType": "application/json"}}
    try:
        r = requests.post(GEMINI_URL, headers={"x-goog-api-key": GEMINI_API_KEY, "Content-Type": "application/json"},
                          json=payload, timeout=timeout)
    except requests.RequestException as e:
        raise RuntimeError(f"Could not reach Gemini: {e}") from e
    if not r.ok:
        raise RuntimeError(f"Gemini returned HTTP {r.status_code}: {r.text[:800]}")
    data = r.json()
    try:
        text = data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Unexpected Gemini response: {str(data)[:800]}") from e
    text = re.sub(r"^```json|```$", "", text.strip(), flags=re.MULTILINE).strip()
    return json.loads(text)


CHAT_PROMPT = """You are the parsing engine behind a factory inventory assistant at {company}
(electronics; goods arrive as CBU finished units, SKD semi-knocked-down kits or CKD knocked-down parts,
are assembled on production lines under work orders, and dispatched).
Return ONLY one JSON object, one of:

1) A stock action:
{{"type":"action","action":"DISPATCH"|"TRANSFER"|"ADJUST"|"ISSUE",
  "item":"<item name or code>", "qty":<number>, "serials":["..."],
  "warehouse":"<warehouse>",            // DISPATCH, ADJUST
  "party":"<customer>", "doc_no":"<invoice/challan no>",   // DISPATCH
  "from_warehouse":"..","to_warehouse":"..","to_bin":"..", // TRANSFER
  "zone":"RAW_STORE"|"FG_STORE"|"QUARANTINE",             // optional
  "bin":"..",                            // ADJUST with positive qty
  "reason":"..",                         // ADJUST (positive=found, negative=lost/damaged)
  "wo_no":"WO-...."}}                      // ISSUE (issue material to a work order)

2) A question:
{{"type":"query","query_kind":"ITEM_STOCK"|"LOCATION_STOCK"|"AGING"|"SERIAL_TRACE"|"WO_STATUS"|"BUILDABLE"|"GRN"|"PENDING_QC",
  "item":"<name or null>","warehouse":"<name or null>","zone":"<zone or null>","serial":"<serial or null>",
  "wo_no":"<wo or null>","grn_no":"<grn or null>"}}

3) {{"type":"unknown","reason":"<short reason>"}}

RULES
- Today is {today}. Warehouses allowed: {warehouses}. Normalise warehouse names to exactly one of these.
- Known items (reuse exact spelling): {items}
- New goods coming in (inward / GRN) are NOT handled in chat: return {{"type":"unknown","reason":"Please use the Inward (GRN) page to receive goods - it captures invoice, BoE, QC and serials."}}
- Production output is NOT handled in chat: return unknown with reason "Please use the Production page to report output."
- "sold", "dispatched", "shipped" -> DISPATCH. "moved", "shifted", "transferred" -> TRANSFER.
  "damaged", "lost", "found extra", "count correction" -> ADJUST. "issue to WO / line" -> ISSUE.
- "where is serial X", "trace X", "history of X" -> SERIAL_TRACE. "how many can we build/assemble" -> BUILDABLE.
  "old stock", "sitting too long" -> AGING. "waiting for QC" -> PENDING_QC.
- Never invent quantities, serials or documents. Leave unknown fields out.
"""


DOC_PROMPT = """You extract data from inventory documents for {company}, an electronics manufacturer
that imports CBU / SKD / CKD goods and buys locally. Read the attached document (commercial invoice,
packing list, bill of entry, airway bill / bill of lading, delivery challan, tax invoice or GRN).

Return ONLY JSON:
{{
 "document_type": "COMMERCIAL_INVOICE"|"PACKING_LIST"|"BILL_OF_ENTRY"|"AWB_BL"|"DELIVERY_CHALLAN"|"TAX_INVOICE"|"OTHER",
 "direction": "INWARD"|"OUTWARD"|"UNKNOWN",
 "suggested_inward_type": "CBU"|"SKD"|"CKD"|"LOCAL"|null,
 "supplier": null, "customer": null, "country_of_origin": null,
 "invoice_no": null, "invoice_date": "YYYY-MM-DD or null", "po_no": null,
 "currency": null, "invoice_value": null,
 "bl_awb_no": null, "container_no": null, "bill_of_entry_no": null, "boe_date": null, "port_of_entry": null,
 "vehicle_no": null, "delivery_address": null,
 "items": [{{"item_name": "", "item_code": "part no / model if printed", "hsn_code": null,
             "qty": 0, "unit": "pcs", "unit_price": null, "serials": ["only if printed"]}}],
 "confidence_notes": "what was unclear"
}}
RULES: INWARD = goods coming to us (our company: {company}; our addresses: {addresses}).
OUTWARD = goods going from us to a customer. suggested_inward_type: imported finished units -> CBU,
partly assembled kits / main boards + housings -> SKD, loose components for assembly -> CKD, Indian supplier -> LOCAL.
Never invent numbers, serials or dates; use null when unreadable. Keep every line item separate.
Known item names (reuse exact spelling if it is clearly the same item): {items}
"""


# ============================================================================
# API
# ============================================================================
app = FastAPI(title="Panache Digilife - Factory Inventory")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


@app.exception_handler(Exception)
async def _all_exceptions(request, exc: Exception):
    if isinstance(exc, HTTPException):
        return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})
    return JSONResponse(status_code=500, content={"detail": f"{type(exc).__name__}: {exc}"})


def _user(d: dict) -> str:
    return s(d.get("entered_by")) or "unknown"


@app.post("/api/login")
def login(d: dict):
    if s(d.get("access_code")) != ACCESS_CODE:
        return {"success": False, "message": "Incorrect access code."}
    if not s(d.get("name")):
        return {"success": False, "message": "Please enter your name."}
    return {"success": True, "name": s(d.get("name"))}


@app.get("/api/meta")
def meta():
    return {"warehouses": WAREHOUSES, "inward_types": INWARD_TYPES, "item_types": ITEM_TYPES, "zones": ZONES,
            "outward_types": {k: v["label"] for k, v in OUTWARD_TYPES.items()}, "aging_days": AGING_THRESHOLD_DAYS,
            "company": OWN_COMPANY_NAME}


# ---- masters ---------------------------------------------------------------
@app.get("/api/items")
def items():
    b = read_book()
    stock = {e["item_code"]: e for e in b.stock_summary()}
    bom_fg = {r["FG Item Code"] for r in b.t[BOM].rows()}
    out = []
    for i in b.t[ITEMS].rows():
        e = stock.get(i["Item Code"], {})
        out.append({**i, "on_hand": e.get("total", 0), "available": e.get("available", 0),
                    "has_bom": i["Item Code"] in bom_fg})
    return {"items": out}


@app.post("/api/items")
def save_item(d: dict):
    code = s(d.get("existing_code"))
    if code:
        return run_master(_user(d), lambda b: {"item": b.update_item(code, d), "message": f"Item {code} updated."},
                          f"Update item {code}")
    return run_master(_user(d), lambda b: (lambda it: {"item": it, "message": f"Item {it['Item Code']} - {it['Item Name']} added."})(b.create_item(d)),
                      f"Add item {s(d.get('item_name'))}")


@app.post("/api/items/bulk")
def bulk_items(d: dict):
    """Rows of {item_code,item_name,item_type,serialized,unit,hsn_code,warranty_months,...}. Existing names are skipped."""
    def fn(b: Book):
        added, skipped = [], []
        for i, r in enumerate(d.get("rows") or [], start=1):
            if not s(r.get("item_name")):
                continue
            if b.find_item(r.get("item_name")) or (s(r.get("item_code")) and b.find_item(r.get("item_code"))):
                skipped.append(s(r.get("item_name")))
                continue
            try:
                added.append(b.create_item(r)["Item Code"])
            except ValueError as e:
                raise ValueError(f"Row {i}: {e}")
        return {"message": f"{len(added)} item(s) added, {len(skipped)} already existed.", "skipped": skipped}
    return run_master(_user(d), fn, "Bulk add items")


@app.get("/api/bom")
def get_bom(fg: str = ""):
    b = read_book()
    if fg:
        it = b.find_item(fg)
        return {"fg": it, "lines": b.bom_for(it["Item Code"]) if it else []}
    groups: Dict[str, dict] = {}
    for r in b.t[BOM].rows():
        g = groups.setdefault(r["FG Item Code"], {"fg_item_code": r["FG Item Code"], "fg_item_name": r["FG Item Name"], "lines": []})
        g["lines"].append(r)
    return {"boms": list(groups.values())}


@app.post("/api/bom")
def save_bom(d: dict):
    def fn(b: Book):
        rows = b.set_bom(d.get("fg_item"), d.get("components") or [])
        return {"message": f"BOM saved with {len(rows)} component(s).", "lines": rows}
    return run_master(_user(d), fn, f"BOM for {s(d.get('fg_item'))}")


# ---- transactions ------------------------------------------------------------
@app.post("/api/inward")
def inward(d: dict, dry_run: bool = False):
    return run_txn("INWARD", _user(d), lambda b: b.inward(d), f"GRN {s(d.get('inward_type'))} {s(d.get('invoice_no'))}", dry_run)


@app.post("/api/qc")
def qc(d: dict):
    return run_txn("QC", _user(d), lambda b: b.qc_decision(d), f"QC {s(d.get('batch_id'))}")


@app.post("/api/dispatch")
def dispatch(d: dict):
    return run_txn("DISPATCH", _user(d), lambda b: b.dispatch(d), f"Dispatch {s(d.get('doc_no'))}")


@app.post("/api/transfer")
def transfer(d: dict):
    return run_txn("TRANSFER", _user(d), lambda b: b.transfer(d), f"Transfer {s(d.get('item'))}")


@app.post("/api/adjust")
def adjust(d: dict):
    return run_txn("ADJUST", _user(d), lambda b: b.adjust(d), f"Adjust {s(d.get('item'))}")


@app.post("/api/sales-return")
def sales_return(d: dict):
    return run_txn("SALES_RETURN", _user(d), lambda b: b.sales_return(d), f"Sales return {s(d.get('doc_no'))}")


@app.post("/api/wo")
def create_wo(d: dict):
    return run_txn("WO_CREATE", _user(d), lambda b: b.create_wo(d), f"Create WO {s(d.get('fg_item'))}")


@app.post("/api/wo/issue")
def wo_issue(d: dict):
    return run_txn("WO_ISSUE", _user(d), lambda b: b.wo_issue(d), f"Issue to {s(d.get('wo_no'))}")


@app.post("/api/wo/output")
def wo_output(d: dict):
    return run_txn("WO_OUTPUT", _user(d), lambda b: b.wo_output(d), f"Output {s(d.get('wo_no'))}")


@app.post("/api/wo/return")
def wo_return(d: dict):
    return run_txn("WO_RETURN", _user(d), lambda b: b.wo_return(d), f"Return from {s(d.get('wo_no'))}")


@app.post("/api/wo/close")
def wo_close(d: dict):
    return run_txn("WO_CLOSE", _user(d), lambda b: b.wo_close(d), f"Close {s(d.get('wo_no'))}")


@app.post("/api/undo")
def undo(d: dict = None):
    d = d or {}
    for attempt in range(3):
        raw, sha = STORE.load()
        book = Book.from_bytes(raw)
        try:
            res = book.undo_last(_user(d))
        except ValueError as e:
            return {"success": False, "message": str(e)}
        try:
            STORE.save(book.to_bytes(), sha, f"Undo {res['txn_id']} by {_user(d)}")
            _read_cache.update(sha=None, book=None)
            return {"success": True, "reversed": res,
                    "message": f"Reversed {res['txn_id']} ({res['type']}{' - ' + res['item'] if res['item'] else ''})."}
        except Conflict:
            continue
    raise HTTPException(409, "Could not save the undo, please try again.")


# ---- reads -------------------------------------------------------------------
@app.get("/api/stock")
def stock(item: str = ""):
    b = read_book()
    code = (b.find_item(item) or {}).get("Item Code") if item else None
    return {"stock": b.stock_summary(code)}


@app.get("/api/batches")
def batches(item: str = "", zone: str = "", warehouse: str = "", include_empty: bool = False):
    b = read_book()
    code = (b.find_item(item) or {}).get("Item Code") if item else None
    rows = [r for r in b.t[BATCHES].rows()
            if (include_empty or num(r["Qty Remaining"]) > 0) and (not code or r["Item Code"] == code)
            and (not zone or r["Zone"] == zone.upper()) and (not warehouse or r["Warehouse"] == warehouse)]
    return {"batches": rows}


@app.get("/api/pending-qc")
def pending_qc():
    b = read_book()
    rows = [r for r in b.t[BATCHES].rows() if r["Zone"] == "QC_HOLD" and num(r["Qty Remaining"]) > 0]
    serialized = {i["Item Code"] for i in b.t[ITEMS].rows() if i["Serialized"] == "Y"}
    return {"batches": [{**r, "serialized": r["Item Code"] in serialized} for r in rows]}


@app.get("/api/serials")
def serials(q: str = "", item: str = "", status: str = "", limit: int = 300):
    b = read_book()
    code = (b.find_item(item) or {}).get("Item Code") if item else None
    q = q.strip().upper()
    rows = [r for r in b.t[SERIALS].rows()
            if (not code or r["Item Code"] == code) and (not status or r["Status"] == status)
            and (not q or q in s(r["Serial No"]).upper() or q in s(r["MAC / IMEI"]).upper())]
    return {"serials": rows[-limit:][::-1], "total": len(rows)}


@app.get("/api/trace")
def trace(serial: str):
    return {"results": read_book().trace(serial)}


@app.get("/api/inwards")
def inwards(limit: int = 100):
    b = read_book()
    lines: Dict[str, list] = {}
    for l in b.t[LINES].rows():
        lines.setdefault(l["GRN No"], []).append(l)
    out = [{**g, "lines": lines.get(g["GRN No"], [])} for g in b.t[INWARDS].rows()]
    return {"inwards": out[::-1][:limit]}


@app.get("/api/kit-check")
def kit_check(grn: str):
    b = read_book()
    g = b.t[INWARDS].get(grn)
    if not g:
        raise HTTPException(404, "GRN not found")
    rec: Dict[str, float] = {}
    for l in b.t[LINES].rows():
        if l["GRN No"] == grn:
            rec[l["Item Code"]] = rec.get(l["Item Code"], 0) + num(l["Accepted Qty"])
    if not g["Kit For Item"]:
        return {"kit_check": None}
    return {"kit_check": b.kit_check(g["Kit For Item"], num(g["Kits Count"]), rec)}


@app.get("/api/wo")
def list_wo():
    b = read_book()
    out = []
    for w in b.t[WOS].rows():
        if w["Status"] == "Reversed":
            continue
        out.append({**w, "materials": b.wo_materials(w)})
    return {"work_orders": out[::-1]}


@app.get("/api/buildable")
def buildable():
    return {"buildable": read_book().buildable()}


@app.get("/api/aging-report")
def aging_report(threshold_days: Optional[int] = None):
    t = threshold_days or AGING_THRESHOLD_DAYS
    return {"threshold_days": t, "batches": read_book().aging(t)}


@app.get("/api/transactions")
def transactions(limit: int = 100, type: str = "", item: str = ""):
    b = read_book()
    rows = [r for r in b.t[LOG].rows() if (not type or r["Txn Type"] == type)
            and (not item or norm(item) in norm(r["Item Name"]) or norm(item) == norm(r["Item Code"]))]
    return {"transactions": rows[::-1][:limit]}


@app.get("/api/dashboard")
def dashboard():
    b = read_book()
    stock = b.stock_summary()
    zone_tot: Dict[str, float] = {z: 0 for z in ZONES}
    for e in stock:
        for z, q in e["by_zone"].items():
            zone_tot[z] = zone_tot.get(z, 0) + num(q)
    m = today()[:7]
    grn_month = [g for g in b.t[INWARDS].rows() if s(g["GRN Date"]).startswith(m) and g["Status"] != "Reversed"]
    by_type = {t: 0 for t in INWARD_TYPES}
    for g in grn_month:
        by_type[g["Inward Type"]] = by_type.get(g["Inward Type"], 0) + 1
    kit_issues = [{"grn": g["GRN No"], "kit_item": g["Kit For Item"], "status": g["Kit Check"]}
                  for g in b.t[INWARDS].rows() if g["Kit Check"] and g["Kit Check"] != "Complete" and g["Status"] != "Reversed"]
    open_wo = [w for w in b.t[WOS].rows() if w["Status"] in ("OPEN", "IN_PROGRESS")]
    pending = [r for r in b.t[BATCHES].rows() if r["Zone"] == "QC_HOLD" and num(r["Qty Remaining"]) > 0]
    serial_counts: Dict[str, int] = {}
    for r in b.t[SERIALS].rows():
        serial_counts[r["Status"]] = serial_counts.get(r["Status"], 0) + 1
    return {
        "cards": {"items_in_stock": len(stock), "fg_available": clean(sum(num(e["by_zone"].get("FG_STORE", 0)) for e in stock)),
                  "pending_qc": len(pending), "open_wo": len(open_wo), "grn_this_month": len(grn_month),
                  "aging_alerts": len(b.aging()), "low_stock": len([e for e in stock if e["low_stock"]]),
                  "kit_shortages": len(kit_issues), "quarantine_qty": clean(zone_tot.get("QUARANTINE", 0))},
        "zones": [{"zone": z, "label": ZONES[z], "qty": clean(round(zone_tot.get(z, 0), 4))} for z in ZONES],
        "grn_by_type": by_type, "kit_issues": kit_issues[-8:][::-1], "buildable": b.buildable(),
        "open_wo": [{"wo": w["WO No"], "fg": w["FG Item Name"], "planned": w["Planned Qty"], "produced": w["Produced Qty"],
                     "status": w["Status"]} for w in open_wo][::-1][:8],
        "serial_counts": serial_counts, "aging_days": AGING_THRESHOLD_DAYS,
        "top_stock": sorted(stock, key=lambda e: -num(e["total"]))[:12],
        "recent": b.t[LOG].rows()[::-1][:10],
    }


@app.get("/api/download")
def download_excel():
    raw, _ = STORE.load()
    if raw is None:
        raw = Book(Workbook()).to_bytes()
    return StreamingResponse(io.BytesIO(raw), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=panache_inventory_{today()}.xlsx"})


# ---- AI ------------------------------------------------------------------------
@app.post("/api/chat")
def chat(d: dict):
    msg = s(d.get("message"))
    if not msg:
        raise HTTPException(400, "message is empty")
    user = _user(d)
    b = read_book()
    names = [i["Item Name"] for i in b.t[ITEMS].rows()][:400]
    try:
        p = call_gemini(CHAT_PROMPT.format(company=OWN_COMPANY_NAME, today=today(), warehouses=", ".join(WAREHOUSES),
                                           items=", ".join(names) or "(none yet)"), [{"text": msg}])
    except Exception as e:
        raise HTTPException(502, f"Gemini request failed: {e}")
    return handle_chat(b, p, msg, user)


def handle_chat(b: Book, p: dict, msg: str, user: str) -> dict:
    kind = p.get("type")
    if kind == "unknown":
        return {"success": False, "message": p.get("reason") or "I couldn't understand that."}
    if kind == "action":
        a = s(p.get("action")).upper()
        sers = p.get("serials") or []
        if a == "DISPATCH":
            body = {"outward_type": "SALE", "warehouse": p.get("warehouse"), "party": p.get("party"), "doc_no": p.get("doc_no"),
                    "lines": [{"item": p.get("item"), "qty": p.get("qty"), "serials": sers, "zone": p.get("zone")}]}
            return run_txn("DISPATCH", user, lambda bk: bk.dispatch(body), f"Chat dispatch: {msg[:60]}")
        if a == "TRANSFER":
            body = {"item": p.get("item"), "qty": p.get("qty"), "serials": sers, "from_warehouse": p.get("from_warehouse"),
                    "to_warehouse": p.get("to_warehouse"), "from_zone": p.get("zone"), "to_bin": p.get("to_bin")}
            return run_txn("TRANSFER", user, lambda bk: bk.transfer(body), f"Chat transfer: {msg[:60]}")
        if a == "ADJUST":
            body = {"item": p.get("item"), "qty": p.get("qty"), "serials": sers, "warehouse": p.get("warehouse"),
                    "zone": p.get("zone"), "bin": p.get("bin"), "reason": p.get("reason") or msg}
            return run_txn("ADJUST", user, lambda bk: bk.adjust(body), f"Chat adjust: {msg[:60]}")
        if a == "ISSUE":
            body = {"wo_no": p.get("wo_no"), "lines": [{"item": p.get("item"), "qty": p.get("qty"), "serials": sers}]}
            return run_txn("WO_ISSUE", user, lambda bk: bk.wo_issue(body), f"Chat issue: {msg[:60]}")
        return {"success": False, "message": f"Action '{a}' is not supported in chat."}

    q = s(p.get("query_kind")).upper()
    if q == "SERIAL_TRACE":
        res = b.trace(p.get("serial") or "")
        if not res:
            return {"success": True, "type": "query", "message": f"No serial '{p.get('serial')}' found."}
        lines = []
        for r in res:
            sr = r["serial"]
            where = f"{sr['Warehouse']} / {sr['Zone']} / {sr['Bin']}" if sr["Status"] == "In Stock" else sr["Status"]
            line = f"{sr['Serial No']} - {sr['Item Name']}: {where}"
            if sr["GRN No"]:
                line += f"; received on {sr['GRN No']} ({sr['Inward Type']}" + (f", {r['grn']['Supplier']}" if r.get("grn") else "") + ")"
            if sr["Work Order"]:
                line += f"; work order {sr['Work Order']}"
            if sr["Parent Serial"]:
                line += f"; fitted in {sr['Parent Serial']}"
            if sr["Customer"]:
                line += f"; sent to {sr['Customer']} on {sr['Dispatch Date']} ({sr['Dispatch Doc']})"
            if sr["Warranty Until"]:
                line += f"; warranty until {sr['Warranty Until']}"
            if r.get("built_from"):
                line += "; built from: " + ", ".join(f"{c['component']} ({c['grn'] or 'no GRN'})" for c in r["built_from"][:8])
            lines.append(line)
        return {"success": True, "type": "query", "message": "\n".join(lines)}
    if q == "AGING":
        rows = b.aging()
        if not rows:
            return {"success": True, "type": "query", "message": f"Nothing has been in stock for {AGING_THRESHOLD_DAYS}+ days."}
        return {"success": True, "type": "query", "message": f"{len(rows)} batch(es) are {AGING_THRESHOLD_DAYS}+ days old:\n" +
                "\n".join(f"{r['Item Name']} - {clean(num(r['Qty Remaining']))} {r['Unit']} at {r['Warehouse']}/{r['Zone']}/{r['Bin']}, "
                          f"{r['Age Days']} days ({r['GRN No'] or r['Source']})" for r in rows[:15])}
    if q == "BUILDABLE":
        rows = b.buildable()
        if not rows:
            return {"success": True, "type": "query", "message": "No BOMs are set up yet, so I can't work out buildable units."}
        it = b.find_item(p.get("item")) if p.get("item") else None
        rows = [r for r in rows if not it or r["fg_item_code"] == it["Item Code"]]
        return {"success": True, "type": "query", "message": "\n".join(
            f"{r['fg_item_name']} at {r['warehouse']}: {r['buildable']} unit(s) from store stock"
            + (f" (limited by {', '.join(r['limited_by'])})" if r["limited_by"] else "") for r in rows)}
    if q == "WO_STATUS":
        ws = [w for w in b.t[WOS].rows() if w["Status"] != "Reversed" and (not p.get("wo_no") or w["WO No"] == s(p.get("wo_no")).upper())]
        if not ws:
            return {"success": True, "type": "query", "message": "No matching work orders."}
        return {"success": True, "type": "query", "message": "\n".join(
            f"{w['WO No']}: {w['FG Item Name']} - {clean(num(w['Produced Qty']))}/{clean(num(w['Planned Qty']))} built, "
            f"{clean(num(w['Rejected Qty']))} rejected, {w['Status']}" for w in ws[-10:])}
    if q == "PENDING_QC":
        rows = [r for r in b.t[BATCHES].rows() if r["Zone"] == "QC_HOLD" and num(r["Qty Remaining"]) > 0]
        if not rows:
            return {"success": True, "type": "query", "message": "Nothing is waiting for QC."}
        return {"success": True, "type": "query", "message": "Waiting for QC:\n" + "\n".join(
            f"{r['Item Name']} - {clean(num(r['Qty Remaining']))} {r['Unit']} ({r['GRN No'] or r['Work Order']}, {r['Warehouse']})" for r in rows[:20])}
    if q == "GRN":
        g = b.t[INWARDS].get(s(p.get("grn_no")).upper())
        if not g:
            return {"success": True, "type": "query", "message": f"GRN '{p.get('grn_no')}' not found."}
        return {"success": True, "type": "query", "message":
                f"{g['GRN No']} ({g['Inward Type']}) from {g['Supplier']}, invoice {g['Invoice No']}, BoE {g['Bill of Entry No'] or '-'}, "
                f"{clean(num(g['Total Accepted']))} accepted / {clean(num(g['Total Short']))} short / {clean(num(g['Total Damaged']))} damaged, "
                f"QC {g['QC Status']}" + (f", kit check: {g['Kit Check']}" if g["Kit Check"] else "") + "."}
    if q == "LOCATION_STOCK":
        wh = p.get("warehouse")
        zone = s(p.get("zone")).upper()
        rows = []
        for e in b.stock_summary():
            for l in e["locations"]:
                w_, z_, _ = l["location"].split(" / ", 2)
                if (not wh or w_.lower() == s(wh).lower()) and (not zone or z_ == zone):
                    rows.append(f"{e['item_name']}: {l['quantity']} {e['unit']} ({l['location']})")
        return {"success": True, "type": "query", "message": ("Stock:\n" + "\n".join(rows[:40])) if rows else "No stock there."}
    # ITEM_STOCK
    it = b.find_item(p.get("item")) if p.get("item") else None
    if not it:
        close = [i["Item Name"] for i in b.t[ITEMS].rows() if p.get("item") and norm(p.get("item")) in i["Normalized Name"]]
        return {"success": True, "type": "query", "message": f"I don't know the item '{p.get('item')}'." +
                (f" Did you mean: {', '.join(close[:5])}?" if close else "")}
    st = b.stock_summary(it["Item Code"])
    if not st:
        return {"success": True, "type": "query", "message": f"{it['Item Name']}: no stock."}
    e = st[0]
    return {"success": True, "type": "query", "message":
            f"{e['item_name']}: {e['total']} {e['unit']} in total, {e['available']} available in store.\n" +
            "\n".join(f"  {l['location']}: {l['quantity']}" for l in e["locations"])}


@app.post("/api/scan-document")
async def scan_document(file: UploadFile = File(...)):
    allowed = {"application/pdf", "image/jpeg", "image/png", "image/webp", "image/heic", "image/heif", "image/gif"}
    mime = file.content_type or "application/octet-stream"
    if mime not in allowed:
        raise HTTPException(400, "Please upload a PDF or an image (JPG, PNG, WEBP, HEIC).")
    raw = await file.read()
    if len(raw) > 3 * 1024 * 1024:
        raise HTTPException(413, "Please keep the document under 3 MB.")
    try:
        names = [i["Item Name"] for i in read_book().t[ITEMS].rows()][:300]
    except Exception:
        names = []
    try:
        parsed = call_gemini(DOC_PROMPT.format(company=OWN_COMPANY_NAME, addresses=OWN_DELIVERY_ADDRESSES or "(not set)",
                                               items=", ".join(names) or "(none)"),
                             [{"text": f"Extract this document. Filename: {file.filename}"},
                              {"inlineData": {"mimeType": mime, "data": base64.b64encode(raw).decode("ascii")}}],
                             temperature=0.0)
    except Exception as e:
        raise HTTPException(502, f"Could not read the document: {e}")
    parsed["_filename"] = file.filename
    return {"success": True, "document": parsed}


# Local development only: serve the HTML pages from the same server when LOCAL_XLSX is set.
# (On Vercel the pages are served as static files by vercel.json.)
if LOCAL_XLSX:
    from fastapi.staticfiles import StaticFiles
    app.mount("/", StaticFiles(directory=os.path.dirname(os.path.abspath(__file__)), html=True), name="static")
