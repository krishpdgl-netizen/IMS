"""
main.py
-------
The ENTIRE backend in one file, mirroring the pattern used in the Finance
Management System (GitHub-as-database, no SQL database, Gemini does the
natural-language classification, openpyxl does the bookkeeping).

WHAT THIS SYSTEM IS
  An Inventory Management System with an "AI Warehouse Assistant" chat bot.
  You type things like:
    "received 200 units of Dell Laptop in Mumbai Warehouse"
    "dispatched 50 Dell Laptop from Mumbai Warehouse to Client XYZ"
    "moved 30 Dell Laptop from Mumbai Warehouse to Delhi Warehouse"
    "lost 5 Dell Laptop in Delhi Warehouse due to damage"
    "how much Dell Laptop do we have and where is it kept?"
    "what's been sitting in stock for too long?"
  Gemini classifies the message, the backend applies it to a FIFO batch
  ledger (or, for questions, computes the answer deterministically from
  the ledger - Gemini only extracts WHAT is being asked, it never
  invents the numbers), and everything is saved back to an .xlsx file
  in a GitHub repo, downloadable at any time.

KEY DESIGN DECISION - UNLIKE THE FINANCE SYSTEM, THE CATALOG IS OPEN:
  The finance system had a fixed master template (fixed line items) that
  was never restructured, only its values changed. Here products AND
  locations are NOT fixed - the assistant can create a brand-new product
  or location the first time it's mentioned. To avoid duplicate near-
  identical products ("Dell Laptop" vs "Dell laptops"), every prompt to
  Gemini includes the current distinct list of known products/locations
  and instructs it to reuse an existing exact name whenever the message
  is clearly referring to the same thing.

FIFO MODEL:
  Stock is never stored as a single "quantity per product" number. It is
  stored as individual BATCHES (Product, Location, Date Received, Qty
  Received, Qty Remaining). Every stock-out (dispatch/adjustment-down/
  transfer-out) consumes the OLDEST active batches first for that exact
  (Product, Location) pair. This is what lets the system answer "what's
  been sitting around too long" - it's a direct query over batch ages,
  not an estimate.

Environment variables needed (set these in Vercel Project Settings):
  GITHUB_TOKEN    - GitHub Personal Access Token (Contents: Read & Write)
  GITHUB_REPO     - "your-username/your-repo"
  GITHUB_BRANCH   - usually "main"
  EXCEL_PATH      - where the live .xlsx lives in the repo, e.g. "data/inventory_data.xlsx"
  ACCESS_CODE     - the shared password used to log in
  GEMINI_API_KEY  - Google Gemini API key (from .env / platform secrets - never hardcoded)
                    (model is fixed to gemini-3.1-flash-lite, not configurable via env)
  AGING_THRESHOLD_DAYS - optional, defaults to 90 (roughly 3 months)
"""

import os
import io
import re
import json
import base64
from datetime import datetime, timezone, date
from typing import Dict, List, Optional

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel


# ============================================================================
# SETTINGS
# ============================================================================
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN") or ""
GITHUB_REPO = os.environ.get("GITHUB_REPO") or ""
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH") or "main"
EXCEL_PATH = os.environ.get("EXCEL_PATH") or "data/inventory_data.xlsx"
ACCESS_CODE = os.environ.get("ACCESS_CODE") or "inventory2026"

# Fixed warehouse master list. Do not accept free-text warehouse names;
# this prevents accidental variants such as "bhiwandi branch", "Delhi", etc.
WAREHOUSE_OPTIONS = ["Bhiwandi", "Ghatkopar"]

def validate_warehouse(value: str, field_name: str = "Warehouse") -> str:
    value = (value or "").strip()
    if value not in WAREHOUSE_OPTIONS:
        raise ValueError(f"{field_name} must be one of: {', '.join(WAREHOUSE_OPTIONS)}.")
    return value

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY") or ""
GEMINI_MODEL = "gemini-3.1-flash-lite"  # fixed per requirements - do not swap models
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"

AGING_THRESHOLD_DAYS = int(os.environ.get("AGING_THRESHOLD_DAYS") or "90")
OWN_DELIVERY_ADDRESSES = os.environ.get("OWN_DELIVERY_ADDRESSES") or ""
OWN_COMPANY_NAME = os.environ.get("OWN_COMPANY_NAME") or ""

GITHUB_API = "https://api.github.com"
GH_HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github+json",
}

# ----------------------------------------------------------------------------
# Sheet layout. No master template file is needed (unlike the finance
# system) because there are no fixed line items here - a brand-new
# workbook with just headers is the correct starting point.
# ----------------------------------------------------------------------------
BATCH_SHEET = "Batches"
BATCH_HEADERS = ["Batch ID", "Product", "Location", "Storage Location", "Date Received",
                  "Qty Received", "Qty Remaining", "Unit", "Status",
                  "Serial Number", "Reference No.", "Source Document", "Delivery Address"]

LOG_SHEET = "Transaction Log"
LOG_HEADERS = ["Timestamp", "Original Message", "Action", "Product",
               "From Location", "To Location", "Quantity", "Unit",
               "Entered By", "Status", "Reversal Data",
               "Serial Number", "Reference No.", "Document Type", "Delivery Address",
               "Storage Location", "Source Document"]

ACTIONS = {"RECEIVE", "DISPATCH", "TRANSFER", "ADJUST"}


# ============================================================================
# GITHUB READ / WRITE (the "database") - identical pattern to the finance system
# ============================================================================
def _contents_url(path):
    return f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{path}"


def _get_file_meta():
    """Returns (sha, raw_bytes). raw_bytes is None if the file doesn't exist yet."""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        raise RuntimeError("GITHUB_TOKEN / GITHUB_REPO are not set on the server.")
    resp = requests.get(_contents_url(EXCEL_PATH), headers=GH_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=30)
    if resp.status_code == 404:
        return None, None
    resp.raise_for_status()
    data = resp.json()
    return data["sha"], base64.b64decode(data["content"])


def _put_file(raw_bytes, commit_message):
    sha, _ = _get_file_meta()
    payload = {
        "message": commit_message,
        "content": base64.b64encode(raw_bytes).decode(),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    resp = requests.put(_contents_url(EXCEL_PATH), headers=GH_HEADERS, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_download_url():
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_PATH}"


# ============================================================================
# EXCEL SERVICE
# ============================================================================
_HEADER_FILL = PatternFill("solid", fgColor="0F766E")
_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)


def _style_header(ws: Worksheet, headers: List[str]):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = _HEADER_FONT
        c.fill = _HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


class ExcelService:

    @classmethod
    def new_workbook(cls) -> Workbook:
        wb = Workbook()
        ws = wb.active
        ws.title = BATCH_SHEET
        ws.append(BATCH_HEADERS)
        _style_header(ws, BATCH_HEADERS)
        log = wb.create_sheet(LOG_SHEET)
        log.append(LOG_HEADERS)
        _style_header(log, LOG_HEADERS)
        return wb

    @classmethod
    def load_live_workbook(cls) -> Workbook:
        _, raw = _get_file_meta()
        if raw is None:
            return cls.new_workbook()
        wb = load_workbook(io.BytesIO(raw), data_only=False)
        # Backward-compatible schema migration: append new columns while preserving
        # the original FIFO/log column positions.
        if BATCH_SHEET in wb.sheetnames:
            ws_existing = wb[BATCH_SHEET]
            existing = [ws_existing.cell(1, c).value for c in range(1, ws_existing.max_column + 1)]
            for header in BATCH_HEADERS:
                if header not in existing:
                    ws_existing.cell(1, ws_existing.max_column + 1, header)
                    existing.append(header)
        if LOG_SHEET in wb.sheetnames:
            log_existing = wb[LOG_SHEET]
            existing = [log_existing.cell(1, c).value for c in range(1, log_existing.max_column + 1)]
            for header in LOG_HEADERS:
                if header not in existing:
                    log_existing.cell(1, log_existing.max_column + 1, header)
                    existing.append(header)
        if BATCH_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(BATCH_SHEET)
            ws.append(BATCH_HEADERS)
            _style_header(ws, BATCH_HEADERS)
        if LOG_SHEET not in wb.sheetnames:
            log = wb.create_sheet(LOG_SHEET)
            log.append(LOG_HEADERS)
            _style_header(log, LOG_HEADERS)
        return wb

    @classmethod
    def save(cls, wb: Workbook, message: str):
        buf = io.BytesIO()
        wb.save(buf)
        return _put_file(buf.getvalue(), message)

    # -- Batches -------------------------------------------------------------
    @classmethod
    def _batch_ws(cls, wb: Workbook) -> Worksheet:
        return wb[BATCH_SHEET]

    @classmethod
    def next_batch_id(cls, wb: Workbook) -> int:
        ws = cls._batch_ws(wb)
        max_id = 0
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] is not None:
                max_id = max(max_id, int(row[0]))
        return max_id + 1

    @classmethod
    def all_batches(cls, wb: Workbook) -> List[dict]:
        ws = cls._batch_ws(wb)
        out = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            out.append(dict(zip(BATCH_HEADERS, row)))
        return out

    @classmethod
    def known_products(cls, wb: Workbook) -> List[str]:
        return sorted({b["Product"] for b in cls.all_batches(wb) if b["Product"]})

    @classmethod
    def known_locations(cls, wb: Workbook) -> List[str]:
        return sorted({b["Location"] for b in cls.all_batches(wb) if b["Location"]})

    @classmethod
    def add_batch(cls, wb: Workbook, product: str, location: str, date_received: str,
                  qty: float, unit: str, serial_number: str = "", reference_no: str = "",
                  source_document: str = "", delivery_address: str = "", storage_location: str = "") -> int:
        ws = cls._batch_ws(wb)
        batch_id = cls.next_batch_id(wb)
        ws.append([batch_id, product, location, storage_location or "", date_received, qty, qty, unit, "Active",
                   serial_number or "", reference_no or "", source_document or "", delivery_address or ""])
        return batch_id

    @classmethod
    def _batch_rows(cls, wb: Workbook):
        """Yields (row_cells) for every batch row, cells indexable like BATCH_HEADERS."""
        ws = cls._batch_ws(wb)
        for row in ws.iter_rows(min_row=2):
            if row[0].value is None:
                continue
            yield row

    @classmethod
    def active_batches_fifo(cls, wb: Workbook, product: str, location: str):
        """Active batch ROWS (cell objects) for an exact product+location,
        oldest Date Received first. Matching is case-insensitive/trimmed."""
        rows = []
        for row in cls._batch_rows(wb):
            if (str(row[1].value).strip().lower() == product.strip().lower()
                    and str(row[2].value).strip().lower() == location.strip().lower()
                    and (row[5].value or 0) > 0):
                rows.append(row)
        rows.sort(key=lambda r: str(r[3].value))
        return rows

    @classmethod
    def consume_fifo(cls, wb: Workbook, product: str, location: str, qty: float):
        """Deducts qty from the oldest active batches at (product, location).
        Returns (consumed_breakdown, shortfall) where consumed_breakdown is a
        list of {batch_id, date_received, qty_taken} and shortfall is how much
        could NOT be fulfilled (0 if fully satisfied). Caller decides whether
        a shortfall is acceptable."""
        remaining_needed = qty
        breakdown = []
        for row in cls.active_batches_fifo(wb, product, location):
            if remaining_needed <= 0:
                break
            available = row[5].value or 0
            take = min(available, remaining_needed)
            if take <= 0:
                continue
            row[5].value = round(available - take, 4)
            row[7].value = "Active" if row[5].value > 0 else "Depleted"
            breakdown.append({"batch_id": row[0].value, "date_received": str(row[3].value),
                               "qty_taken": take})
            remaining_needed -= take
        shortfall = round(max(remaining_needed, 0), 4)
        return breakdown, shortfall

    @classmethod
    def restore_fifo(cls, wb: Workbook, breakdown: List[dict]):
        """Exact reversal of consume_fifo, used by undo - adds each taken
        amount back to its exact original batch by Batch ID (never guesses)."""
        rows_by_id = {row[0].value: row for row in cls._batch_rows(wb)}
        for item in breakdown:
            row = rows_by_id.get(item["batch_id"])
            if row is None:
                continue
            row[5].value = round((row[5].value or 0) + item["qty_taken"], 4)
            row[7].value = "Active"

    @classmethod
    def stock_summary(cls, wb: Workbook, product: Optional[str] = None):
        """{'Product Name': {'Location A': qty, 'Location B': qty, ...}, ...}
        restricted to Active/remaining>0 batches. Case preserved from the
        first-seen spelling; matching for filtering is case-insensitive."""
        out: Dict[str, Dict[str, float]] = {}
        for b in cls.all_batches(wb):
            if (b["Qty Remaining"] or 0) <= 0:
                continue
            if product and product.strip().lower() != str(b["Product"]).strip().lower():
                continue
            out.setdefault(b["Product"], {})
            out[b["Product"]][b["Location"]] = round(out[b["Product"]].get(b["Location"], 0) + b["Qty Remaining"], 4)
        return out

    @classmethod
    def aging_batches(cls, wb: Workbook, threshold_days: int = AGING_THRESHOLD_DAYS):
        today = date.today()
        out = []
        for b in cls.all_batches(wb):
            if (b["Qty Remaining"] or 0) <= 0:
                continue
            try:
                d = _parse_date(b["Date Received"])
            except Exception:
                continue
            age = (today - d).days
            if age >= threshold_days:
                out.append({**b, "Age Days": age})
        return sorted(out, key=lambda r: -r["Age Days"])

    # -- Transaction Log -------------------------------------------------------
    @classmethod
    def append_log(cls, wb: Workbook, original_msg: str, action: str, product: str,
                   from_location: Optional[str], to_location: Optional[str],
                   qty: Optional[float], unit: Optional[str], entered_by: str,
                   status: str, reversal_data: Optional[list] = None,
                   serial_number: str = "", reference_no: str = "",
                   document_type: str = "", delivery_address: str = "",
                   source_document: str = "", storage_location: str = "") -> int:
        ws = wb[LOG_SHEET]
        now = datetime.now(timezone.utc).isoformat()
        ws.append([now, original_msg, action, product, from_location, to_location,
                   qty, unit, entered_by, status, json.dumps(reversal_data or []),
                   serial_number or "", reference_no or "", document_type or "",
                   delivery_address or "", storage_location or "", source_document or ""])
        return ws.max_row

    @classmethod
    def get_log_rows(cls, wb: Workbook, limit: Optional[int] = None):
        ws = wb[LOG_SHEET]
        rows = [dict(zip(LOG_HEADERS, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
        rows.reverse()  # newest first
        return rows[:limit] if limit else rows

    @classmethod
    def undo_last(cls, wb: Workbook):
        ws = wb[LOG_SHEET]
        target_row = None
        for r in range(ws.max_row, 1, -1):
            if ws.cell(row=r, column=10).value == "Applied":
                target_row = r
                break
        if target_row is None:
            raise ValueError("No applied transaction to undo.")

        action = ws.cell(row=target_row, column=3).value
        product = ws.cell(row=target_row, column=4).value
        from_loc = ws.cell(row=target_row, column=5).value
        to_loc = ws.cell(row=target_row, column=6).value
        qty = ws.cell(row=target_row, column=7).value
        unit = ws.cell(row=target_row, column=8).value
        reversal_data = json.loads(ws.cell(row=target_row, column=11).value or "[]")

        if action == "RECEIVE":
            # reversal_data holds the single batch_id that was created
            for item in reversal_data:
                for row in cls._batch_rows(wb):
                    if row[0].value == item["batch_id"]:
                        row[5].value = 0
                        row[7].value = "Depleted"
        elif action == "DISPATCH":
            cls.restore_fifo(wb, reversal_data)
        elif action == "ADJUST":
            if qty and qty > 0:
                # was a positive adjustment (new batch) -> deplete it
                for item in reversal_data:
                    for row in cls._batch_rows(wb):
                        if row[0].value == item["batch_id"]:
                            row[5].value = 0
                            row[7].value = "Depleted"
            else:
                cls.restore_fifo(wb, reversal_data)
        elif action == "TRANSFER":
            # reversal_data = {"consumed": [...], "created_batch_ids": [...]}
            cls.restore_fifo(wb, reversal_data.get("consumed", []))
            for bid in reversal_data.get("created_batch_ids", []):
                for row in cls._batch_rows(wb):
                    if row[0].value == bid:
                        row[5].value = 0
                        row[7].value = "Depleted"

        ws.cell(row=target_row, column=10, value="Reversed")
        cls.append_log(wb, f"UNDO: {ws.cell(row=target_row, column=2).value}", action, product,
                        from_loc, to_loc, qty, unit, "system", "Applied (undo)")
        return {"action": action, "product": product, "quantity": qty}


def _parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()


# ============================================================================
# GEMINI SERVICE
# Classifies a free-text message into either a STOCK MOVEMENT or a QUERY.
# Gemini never computes quantities itself - it only extracts structured
# intent; every number in the final answer comes from the Batches sheet.
# ============================================================================
class GeminiService:

    SYSTEM_TEMPLATE = """You are the parsing engine behind a warehouse inventory chat bot.
Read the user's message and return ONLY a single JSON object (no markdown
fences, no commentary) matching exactly one of these shapes:

1) Stock movement:
{{
  "type": "movement",
  "action": "RECEIVE" | "DISPATCH" | "TRANSFER" | "ADJUST",
  "product": "<name>",
  "location": "Bhiwandi" | "Ghatkopar", // required for RECEIVE, DISPATCH, ADJUST
  "storage_location": "<exact rack/shelf/bin location>", // required for RECEIVE; optional otherwise
  "from_location": "Bhiwandi" | "Ghatkopar", // required for TRANSFER only
  "to_location": "Bhiwandi" | "Ghatkopar",   // required for TRANSFER only
  "quantity": <number>,             // for ADJUST, positive = found extra stock, negative = write-off/loss/damage
  "unit": "<e.g. units, boxes, kg>",
  "date_received": "YYYY-MM-DD"     // only for RECEIVE. ONLY fill this in if the
                                     // message explicitly states a date (e.g. "received on
                                     // 3rd May", "arrived last Monday"). If no date is
                                     // mentioned, leave this field out entirely / null -
                                     // do NOT guess or default to today yourself.
}}

2) Stock question ("how much do we have", "where is it kept", "what do we have in X location"):
{{
  "type": "query",
  "query_kind": "PRODUCT_LOOKUP" | "LOCATION_LOOKUP" | "AGING_REPORT",
  "product": "<name or null>",
  "location": "<name or null>"
}}

3) Could not classify:
{{
  "type": "unknown",
  "reason": "<short reason>"
}}

RULES:
- TODAY'S ACTUAL DATE IS: {today}. If the message states a relative date
  ("yesterday", "last Monday", "3 days ago"), resolve it against this real
  date - never against your own guess of what today is.
- KNOWN PRODUCTS SO FAR: {known_products}
- APPROVED WAREHOUSES: Bhiwandi, Ghatkopar.
- Warehouse/location is NOT free text. For RECEIVE, DISPATCH, ADJUST and
  TRANSFER, normalize warehouse names to exactly "Bhiwandi" or "Ghatkopar".
  Never create or preserve variants such as "bhiwandi branch", "Bhiwandi
  Warehouse", "Delhi", etc. If the message clearly means one of the two
  approved warehouses, use that exact spelling; otherwise let the backend
  reject the invalid warehouse rather than inventing a new one.
- The known product list exists only to help reuse exact product spelling.
  New products are allowed. New warehouses are NOT allowed.
- "received", "arrived", "bought", "purchased", "added to stock", "brought in",
  "we have X of <product>" (describing stock that exists/arrived)
  -> RECEIVE.
- "dispatched", "sold", "shipped out", "sent to customer", "used", "consumed",
  "issued" -> DISPATCH.
- "moved", "transferred", "shifted" between two locations -> TRANSFER.
- "damaged", "lost", "expired", "written off", "found extra", "stock count
  correction" -> ADJUST (negative quantity for loss/damage/write-off,
  positive for found/extra).
- For RECEIVE messages, if the user mentions a rack, shelf, bin, aisle, bay, cabinet or other exact storage spot, put that in "storage_location".
- Questions about quantity/location/"how long has X been sitting"/"what's
  old stock"/"what's been here for months" -> type "query".
- Only return "unknown" when something essential and irreplaceable is
  truly missing from the message itself - e.g. no quantity at all, or no
  product mentioned, or the intent genuinely can't be told apart between
  a movement and a question. An unfamiliar product/location name is
  ALWAYS enough information on its own - never a reason for "unknown".
Return ONLY the JSON object, nothing else.
"""

    @classmethod
    def parse_message(cls, message: str, known_products: List[str], known_locations: List[str]) -> dict:
        if not GEMINI_API_KEY:
            raise RuntimeError("GEMINI_API_KEY is not set on the server.")
        system_prompt = cls.SYSTEM_TEMPLATE.format(
            today=date.today().isoformat(),
            known_products=", ".join(known_products) or "(none yet)",
            known_locations=", ".join(known_locations) or "(none yet)",
        )
        payload = {
            "system_instruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"role": "user", "parts": [{"text": message}]}],
            "generationConfig": {"temperature": 0.1, "response_mime_type": "application/json"},
        }
        resp = requests.post(f"{GEMINI_URL}?key={GEMINI_API_KEY}", json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        try:
            text = data["candidates"][0]["content"]["parts"][0]["text"]
        except (KeyError, IndexError) as e:
            raise RuntimeError(f"Unexpected Gemini response shape: {data}") from e
        text = re.sub(r"^```json|```$", "", text.strip(), flags=re.MULTILINE).strip()
        return json.loads(text)


# ============================================================================
# INVENTORY ACTION SERVICE
# Applies a parsed movement to the Batches sheet and returns a human-readable
# result. This is where FIFO logic actually gets used.
# ============================================================================
class InventoryService:

    @classmethod
    def apply_movement(cls, wb: Workbook, parsed: dict, original_msg: str, entered_by: str,
                        metadata: Optional[dict] = None) -> dict:
        action = parsed.get("action")
        product = (parsed.get("product") or "").strip()
        unit = (parsed.get("unit") or "units").strip()
        qty = parsed.get("quantity")
        metadata = metadata or {}
        serial_number = (parsed.get("serial_number") or metadata.get("serial_number") or "").strip()
        reference_no = (parsed.get("reference_no") or metadata.get("reference_no") or "").strip()
        source_document = (metadata.get("source_document") or "").strip()
        delivery_address = (parsed.get("delivery_address") or metadata.get("delivery_address") or "").strip()
        storage_location = (parsed.get("storage_location") or metadata.get("storage_location") or "").strip()
        document_type = (metadata.get("document_type") or "").strip()

        if action not in ACTIONS:
            raise ValueError(f"Unknown action '{action}'.")
        if not product:
            raise ValueError("No product specified.")

        if action == "RECEIVE":
            location = (parsed.get("location") or "").strip()
            if not location:
                raise ValueError("No warehouse specified for received stock.")
            location = validate_warehouse(location, "Warehouse")
            if not storage_location:
                raise ValueError("Please enter the exact warehouse storage location (for example: Rack A3 / Shelf 2 / Bin 04).")
            if not qty or qty <= 0:
                raise ValueError("Quantity to receive must be a positive number.")
            date_received = parsed.get("date_received") or date.today().isoformat()
            batch_id = ExcelService.add_batch(wb, product, location, date_received, qty, unit, serial_number, reference_no, source_document, delivery_address, storage_location)
            ExcelService.append_log(wb, original_msg, "RECEIVE", product, None, location,
                                     qty, unit, entered_by, "Applied",
                                     [{"batch_id": batch_id}], serial_number, reference_no,
                                     document_type, delivery_address, source_document, storage_location)
            return {"action": "RECEIVE", "product": product, "location": location,
                    "storage_location": storage_location, "quantity": qty, "unit": unit, "date_received": date_received,
                    "message": f"Recorded {qty} {unit} of {product} received at {location}, stored at {storage_location} "
                               f"(dated {date_received})."}

        if action == "DISPATCH":
            location = (parsed.get("location") or "").strip()
            if not location:
                raise ValueError("No location specified for dispatch.")
            location = validate_warehouse(location, "Warehouse")
            if not qty or qty <= 0:
                raise ValueError("Quantity to dispatch must be a positive number.")
            breakdown, shortfall = ExcelService.consume_fifo(wb, product, location, qty)
            if shortfall > 0:
                # Roll back what we just took, since a partial dispatch could
                # be wrong for the user's bookkeeping - ask instead of guessing.
                ExcelService.restore_fifo(wb, breakdown)
                available = sum(v for locs in ExcelService.stock_summary(wb, product).values()
                                 for k, v in locs.items() if k == location)
                raise ValueError(f"Only {available} {unit} of {product} available at {location}, "
                                  f"cannot dispatch {qty}.")
            ExcelService.append_log(wb, original_msg, "DISPATCH", product, location, None,
                                     qty, unit, entered_by, "Applied", breakdown,
                                     serial_number, reference_no, document_type,
                                     delivery_address, source_document)
            oldest = breakdown[0]["date_received"] if breakdown else None
            return {"action": "DISPATCH", "product": product, "location": location,
                    "quantity": qty, "unit": unit,
                    "message": f"Dispatched {qty} {unit} of {product} from {location} "
                               f"(FIFO: oldest batch used first{', from ' + oldest if oldest else ''})."}

        if action == "TRANSFER":
            from_location = (parsed.get("from_location") or "").strip()
            to_location = (parsed.get("to_location") or "").strip()
            if not from_location or not to_location:
                raise ValueError("Both from_location and to_location are required for a transfer.")
            from_location = validate_warehouse(from_location, "From warehouse")
            to_location = validate_warehouse(to_location, "To warehouse")
            if not qty or qty <= 0:
                raise ValueError("Quantity to transfer must be a positive number.")
            breakdown, shortfall = ExcelService.consume_fifo(wb, product, from_location, qty)
            if shortfall > 0:
                ExcelService.restore_fifo(wb, breakdown)
                raise ValueError(f"Not enough {product} at {from_location} to transfer {qty} {unit}.")
            created_ids = []
            for item in breakdown:
                bid = ExcelService.add_batch(wb, product, to_location, item["date_received"],
                                              item["qty_taken"], unit, serial_number, reference_no,
                                              source_document, delivery_address)
                created_ids.append(bid)
            ExcelService.append_log(wb, original_msg, "TRANSFER", product, from_location, to_location,
                                     qty, unit, entered_by, "Applied",
                                     {"consumed": breakdown, "created_batch_ids": created_ids},
                                     serial_number, reference_no, document_type,
                                     delivery_address, source_document)
            return {"action": "TRANSFER", "product": product, "from_location": from_location,
                    "to_location": to_location, "quantity": qty, "unit": unit,
                    "message": f"Transferred {qty} {unit} of {product} from {from_location} to "
                               f"{to_location} (original batch dates preserved for aging)."}

        if action == "ADJUST":
            location = (parsed.get("location") or "").strip()
            if not location:
                raise ValueError("No location specified for the adjustment.")
            location = validate_warehouse(location, "Warehouse")
            if qty is None or qty == 0:
                raise ValueError("Adjustment quantity must be a non-zero number.")
            if qty > 0:
                batch_id = ExcelService.add_batch(wb, product, location, date.today().isoformat(), qty, unit, serial_number, reference_no, source_document, delivery_address)
                ExcelService.append_log(wb, original_msg, "ADJUST", product, None, location,
                                         qty, unit, entered_by, "Applied", [{"batch_id": batch_id}],
                                         serial_number, reference_no, document_type,
                                         delivery_address, source_document)
                return {"action": "ADJUST", "product": product, "location": location,
                        "quantity": qty, "unit": unit,
                        "message": f"Added {qty} {unit} of {product} to {location} as a stock-count correction."}
            else:
                breakdown, shortfall = ExcelService.consume_fifo(wb, product, location, abs(qty))
                if shortfall > 0:
                    ExcelService.restore_fifo(wb, breakdown)
                    raise ValueError(f"Cannot write off {abs(qty)} {unit} of {product} at {location} - "
                                      f"only {abs(qty) - shortfall} {unit} on hand.")
                ExcelService.append_log(wb, original_msg, "ADJUST", product, location, None,
                                         qty, unit, entered_by, "Applied", breakdown,
                                         serial_number, reference_no, document_type,
                                         delivery_address, source_document)
                return {"action": "ADJUST", "product": product, "location": location,
                        "quantity": qty, "unit": unit,
                        "message": f"Wrote off {abs(qty)} {unit} of {product} at {location}."}

        raise ValueError("Unhandled action.")

    @classmethod
    def answer_query(cls, wb: Workbook, parsed: dict) -> dict:
        kind = parsed.get("query_kind")
        product = parsed.get("product")
        location = parsed.get("location")

        if kind == "AGING_REPORT":
            rows = ExcelService.aging_batches(wb)
            if not rows:
                return {"type": "query", "query_kind": kind,
                        "message": f"Nothing has been sitting in stock for {AGING_THRESHOLD_DAYS}+ days. All clear.",
                        "rows": []}
            lines = [f"{r['Product']} at {r['Location']}: {r['Qty Remaining']} {r['Unit']} "
                     f"(received {r['Date Received']}, {r['Age Days']} days old)" for r in rows[:15]]
            return {"type": "query", "query_kind": kind,
                    "message": f"{len(rows)} batch(es) have been in stock {AGING_THRESHOLD_DAYS}+ days:\n" + "\n".join(lines),
                    "rows": rows}

        if kind == "LOCATION_LOOKUP" and location:
            all_summary = ExcelService.stock_summary(wb)
            found = {p: locs[location] for p, locs in all_summary.items() if location in locs}
            if not found:
                return {"type": "query", "query_kind": kind,
                        "message": f"No stock currently recorded at {location}.", "rows": []}
            lines = [f"{p}: {q}" for p, q in sorted(found.items())]
            return {"type": "query", "query_kind": kind,
                    "message": f"Stock currently at {location}:\n" + "\n".join(lines),
                    "rows": found}

        # default: PRODUCT_LOOKUP
        if not product:
            return {"type": "query", "query_kind": kind,
                    "message": "I couldn't tell which product you're asking about.", "rows": []}
        summary = ExcelService.stock_summary(wb, product)
        if not summary:
            known = ExcelService.known_products(wb)
            close = [p for p in known if product.strip().lower() in p.lower() or p.lower() in product.strip().lower()]
            hint = f" Did you mean: {', '.join(close)}?" if close else ""
            return {"type": "query", "query_kind": kind,
                    "message": f"No stock currently recorded for '{product}'.{hint}", "rows": []}
        product_name = next(iter(summary))
        locs = summary[product_name]
        total = round(sum(locs.values()), 4)
        breakdown = "; ".join(f"{loc}: {q}" for loc, q in sorted(locs.items()))
        return {"type": "query", "query_kind": kind,
                "message": f"{product_name} - {total} total, kept at: {breakdown}.",
                "rows": locs}



# ============================================================================
# DOCUMENT INTELLIGENCE
# Upload a delivery challan / invoice / inward-outward bill and let Gemini
# extract the transaction. The result is ALWAYS returned for human review
# before anything is written to the ledger.
# ============================================================================
class DocumentIntelligence:
    PROMPT = """You are an inventory document extraction engine.
Read the attached delivery challan, invoice, goods receipt, or out-bill/in-bill.

Return ONLY valid JSON in this exact shape:
{
  "document_type": "DELIVERY_CHALLAN" | "INVOICE" | "GOODS_RECEIPT" | "OTHER",
  "document_number": "string or null",
  "document_date": "YYYY-MM-DD or null",
  "delivery_address": "full delivery address or null",
  "billing_address": "full billing address or null",
  "party_name": "supplier/customer/party name or null",
  "direction": "INWARD" | "OUTWARD" | "INTERNAL" | "UNKNOWN",
  "source_location": "known internal location or null",
  "destination_location": "known internal location or null",
  "items": [
    {
      "product": "clean product/item description",
      "serial_number": "serial number(s) if visible, otherwise empty",
      "quantity": number,
      "unit": "units/pcs/boxes/kg/etc.",
      "description": "short original line description"
    }
  ],
  "confidence_notes": "brief notes about ambiguous or unreadable fields"
}

DIRECTION RULE:
- INWARD means goods are coming INTO our company/warehouse. A delivery address
  matching one of our own addresses or internal warehouse locations generally
  indicates INWARD.
- OUTWARD means goods are going TO a customer/external destination. A delivery
  address that is not one of our own addresses generally indicates OUTWARD.
- INTERNAL means goods move between two of our known internal locations.
- UNKNOWN means the address evidence is insufficient.
Never invent a quantity, product, serial number, date, or address. Use null/empty
when unreadable. Preserve multiple line items separately.
Known internal locations: {known_locations}
Our company name: {company_name}
Our own delivery/warehouse addresses: {own_addresses}
"""

    @classmethod
    def parse(cls, raw: bytes, mime_type: str, filename: str, known_locations: List[str]) -> dict:
        if not GEMINI_API_KEY:
            raise RuntimeError("GEMINI_API_KEY is not set on the server.")
        prompt = cls.PROMPT.format(
            known_locations=", ".join(known_locations) or "(none yet)",
            company_name=OWN_COMPANY_NAME or "(not configured)",
            own_addresses=OWN_DELIVERY_ADDRESSES or "(not configured; use known internal locations where possible)",
        )
        # Use the canonical Gemini REST JSON field names here.  This is important
        # for Vercel/serverless deployments: the API accepts the protobuf JSON
        # representation, but the REST field names should be camelCase.
        payload = {
            "systemInstruction": {"parts": [{"text": prompt}]},
            "contents": [{
                "role": "user",
                "parts": [
                    {"text": f"Extract this inventory document. Filename: {filename}"},
                    {
                        "inlineData": {
                            "mimeType": mime_type,
                            "data": base64.b64encode(raw).decode("ascii")
                        }
                    }
                ]
            }],
            "generationConfig": {
                "temperature": 0.0,
                "responseMimeType": "application/json"
            },
        }

        # Keep the API key out of the URL and send it using Google's documented
        # authentication header. This also makes failures easier to diagnose.
        try:
            resp = requests.post(
                GEMINI_URL,
                headers={
                    "x-goog-api-key": GEMINI_API_KEY,
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=60,
            )
        except requests.RequestException as e:
            raise RuntimeError(f"Could not reach Gemini API: {e}") from e

        if not resp.ok:
            try:
                error_body = resp.json()
            except ValueError:
                error_body = resp.text[:2000]
            raise RuntimeError(
                f"Gemini API returned HTTP {resp.status_code}: {error_body}"
            )

        try:
            data = resp.json()
        except ValueError as e:
            raise RuntimeError(f"Gemini returned a non-JSON response: {resp.text[:2000]}") from e
        try:
            text = data["candidates"][0]["content"]["parts"][0]["text"]
        except (KeyError, IndexError) as e:
            raise RuntimeError(f"Unexpected Gemini document response: {data}") from e
        text = re.sub(r"^```json|```$", "", text.strip(), flags=re.MULTILINE).strip()
        parsed = json.loads(text)
        parsed["_filename"] = filename
        return parsed

# ============================================================================
# API
# ============================================================================
app = FastAPI(title="Inventory Management System - AI Warehouse Assistant")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


@app.exception_handler(Exception)
async def _all_exceptions(request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": f"{type(exc).__name__}: {exc}"})


class LoginIn(BaseModel):
    name: str
    access_code: str


class ChatIn(BaseModel):
    message: str
    entered_by: str


class TransactionIn(BaseModel):
    action: str
    product: str
    quantity: float
    unit: str = "units"
    location: Optional[str] = ""
    from_location: Optional[str] = ""
    to_location: Optional[str] = ""
    date_received: Optional[str] = None
    serial_number: Optional[str] = ""
    reference_no: Optional[str] = ""
    delivery_address: Optional[str] = ""
    storage_location: Optional[str] = ""
    document_type: Optional[str] = ""
    source_document: Optional[str] = ""
    entered_by: str


class DocumentLogIn(BaseModel):
    document_type: str = ""
    document_number: str = ""
    document_date: Optional[str] = None
    delivery_address: str = ""
    direction: str
    source_location: str = ""
    destination_location: str = ""
    party_name: str = ""
    items: list
    entered_by: str
    source_document: str = ""


@app.post("/api/login")
def login(data: LoginIn):
    if data.access_code != ACCESS_CODE:
        return {"success": False, "message": "Incorrect access code."}
    if not data.name.strip():
        return {"success": False, "message": "Please enter your name."}
    return {"success": True, "name": data.name.strip()}


@app.post("/api/log")
def log_transaction(data: TransactionIn):
    """Human-friendly form entry. Uses the exact same FIFO engine as the AI."""
    parsed = {
        "type": "movement",
        "action": data.action.upper().strip(),
        "product": data.product.strip(),
        "quantity": data.quantity,
        "unit": data.unit.strip() or "units",
        "location": (data.location or "").strip(),
        "from_location": (data.from_location or "").strip(),
        "to_location": (data.to_location or "").strip(),
        "date_received": data.date_received,
        "serial_number": (data.serial_number or "").strip(),
        "reference_no": (data.reference_no or "").strip(),
        "delivery_address": (data.delivery_address or "").strip(),
        "storage_location": (data.storage_location or "").strip(),
    }
    if parsed["action"] == "RECEIVE" and not parsed["date_received"]:
        parsed["date_received"] = date.today().isoformat()
    metadata = {
        "serial_number": data.serial_number or "",
        "reference_no": data.reference_no or "",
        "delivery_address": data.delivery_address or "",
        "storage_location": data.storage_location or "",
        "document_type": data.document_type or "MANUAL_FORM",
        "source_document": data.source_document or "",
    }
    try:
        wb = ExcelService.load_live_workbook()
        result = InventoryService.apply_movement(
            wb, parsed,
            f"Manual form entry: {parsed['action']} {parsed['quantity']} {parsed['unit']} {parsed['product']}",
            data.entered_by,
            metadata,
        )
        ExcelService.save(wb, f"Manual {result['action']}: {data.product} ({data.entered_by})")
        return {"success": True, **result}
    except ValueError as e:
        return {"success": False, "message": str(e)}
    except Exception as e:
        raise HTTPException(500, f"Could not log transaction: {e}")


@app.post("/api/scan-document")
async def scan_document(file: UploadFile = File(...)):
    """Extract a document into a reviewable transaction draft; does NOT write stock."""
    if not file.filename:
        raise HTTPException(400, "Please choose a document.")
    allowed = {
        "application/pdf", "image/jpeg", "image/png", "image/webp",
        "image/heic", "image/heif", "image/gif"
    }
    mime = file.content_type or "application/octet-stream"
    if mime not in allowed:
        raise HTTPException(400, "Please upload a PDF or image (JPG, PNG, WEBP, HEIC).")
    raw = await file.read()
    if len(raw) > 15 * 1024 * 1024:
        raise HTTPException(413, "Document is too large. Please keep it under 15 MB.")
    try:
        wb = ExcelService.load_live_workbook()
        parsed = DocumentIntelligence.parse(raw, mime, file.filename, ExcelService.known_locations(wb))
        # Convert document-level extraction into ready-to-review transaction drafts.
        direction = str(parsed.get("direction") or "UNKNOWN").upper()
        action = {"INWARD": "RECEIVE", "OUTWARD": "DISPATCH", "INTERNAL": "TRANSFER"}.get(direction, "UNKNOWN")
        drafts = []
        for item in parsed.get("items") or []:
            drafts.append({
                "action": action,
                "product": item.get("product") or "",
                "serial_number": item.get("serial_number") or "",
                "quantity": item.get("quantity") or 0,
                "unit": item.get("unit") or "units",
                "date_received": parsed.get("document_date") or date.today().isoformat(),
                "location": (
                    (parsed.get("destination_location") or "") if action == "RECEIVE"
                    else (parsed.get("source_location") or "") if action == "DISPATCH"
                    else ""
                ),
                "storage_location": "",
                "from_location": parsed.get("source_location") or "",
                "to_location": parsed.get("destination_location") or "",
                "reference_no": parsed.get("document_number") or "",
                "delivery_address": parsed.get("delivery_address") or "",
            })
        return {
            "success": True,
            "document": parsed,
            "action": action,
            "drafts": drafts,
            "needs_review": action == "UNKNOWN" or not drafts,
        }
    except Exception as e:
        raise HTTPException(502, f"Could not read the document: {e}")


@app.post("/api/log-document")
def log_document(data: DocumentLogIn):
    """Commit reviewed document drafts. Each line becomes a normal inventory transaction."""
    try:
        wb = ExcelService.load_live_workbook()
        results = []
        for item in data.items:
            action = str(item.get("action") or "").upper()
            if action == "UNKNOWN":
                raise ValueError("Direction is still unknown. Please choose Inward or Outward for every line.")
            parsed = {
                "type": "movement",
                "action": action,
                "product": str(item.get("product") or "").strip(),
                "quantity": float(item.get("quantity") or 0),
                "unit": str(item.get("unit") or "units").strip(),
                "location": str(item.get("location") or "").strip(),
                "from_location": str(item.get("from_location") or "").strip(),
                "to_location": str(item.get("to_location") or "").strip(),
                "date_received": item.get("date_received") or data.document_date or date.today().isoformat(),
                "serial_number": str(item.get("serial_number") or "").strip(),
                "reference_no": str(item.get("reference_no") or data.document_number or "").strip(),
                "delivery_address": str(item.get("delivery_address") or data.delivery_address or "").strip(),
                "storage_location": str(item.get("storage_location") or "").strip(),
            }
            metadata = {
                "serial_number": parsed["serial_number"],
                "reference_no": parsed["reference_no"],
                "delivery_address": parsed["delivery_address"],
                "storage_location": parsed["storage_location"],
                "document_type": data.document_type or "SCANNED_DOCUMENT",
                "source_document": data.source_document or "",
            }
            results.append(InventoryService.apply_movement(
                wb, parsed,
                f"Scanned {data.document_type or 'document'} {data.document_number}".strip(),
                data.entered_by, metadata
            ))
        ExcelService.save(wb, f"Logged scanned document {data.document_number} ({data.entered_by})")
        return {"success": True, "results": results}
    except ValueError as e:
        return {"success": False, "message": str(e)}
    except Exception as e:
        raise HTTPException(500, f"Could not log scanned document: {e}")


@app.post("/api/chat")
def chat(data: ChatIn):
    if not data.message.strip():
        raise HTTPException(400, "message is empty")

    wb = ExcelService.load_live_workbook()
    known_products = ExcelService.known_products(wb)
    known_locations = ExcelService.known_locations(wb)

    try:
        parsed = GeminiService.parse_message(data.message, known_products, known_locations)
    except Exception as e:
        raise HTTPException(502, f"Gemini request failed: {e}")

    if parsed.get("type") == "unknown":
        ExcelService.append_log(wb, data.message, "REJECTED", parsed.get("product", ""), None, None,
                                 None, None, data.entered_by, "Rejected")
        try:
            ExcelService.save(wb, f"Log rejected message ({data.entered_by})")
        except Exception:
            pass
        return {"success": False, "type": "unknown",
                "message": parsed.get("reason", "Could not understand that message.")}

    if parsed.get("type") == "query":
        result = InventoryService.answer_query(wb, parsed)
        return {"success": True, **result}

    # movement
    try:
        result = InventoryService.apply_movement(wb, parsed, data.message, data.entered_by)
    except ValueError as e:
        return {"success": False, "type": "movement_rejected", "message": str(e)}

    try:
        ExcelService.save(wb, f"{result['action']}: {data.message[:60]} ({data.entered_by})")
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")

    return {"success": True, "type": "movement", **result}


@app.post("/api/undo")
def undo():
    wb = ExcelService.load_live_workbook()
    try:
        result = ExcelService.undo_last(wb)
    except ValueError as e:
        raise HTTPException(400, str(e))
    try:
        ExcelService.save(wb, "Undo last transaction")
    except Exception as e:
        raise HTTPException(500, f"Could not save the workbook to GitHub: {e}")
    return {"success": True, "reversed": result}


@app.get("/api/transactions")
def transactions(limit: int = 50):
    wb = ExcelService.load_live_workbook()
    return {"transactions": ExcelService.get_log_rows(wb, limit)}


@app.get("/api/products")
def products():
    """Current stock, grouped by product then location - powers the Reports table."""
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    out = []
    batches = ExcelService.all_batches(wb)
    for product, locs in sorted(summary.items()):
        serials = []
        for b in batches:
            if str(b["Product"]).strip().lower() == product.strip().lower() and (b["Qty Remaining"] or 0) > 0:
                sn = str(b.get("Serial Number") or "").strip()
                if sn:
                    serials.append({"serial_number": sn, "location": b["Location"], "quantity": b["Qty Remaining"]})
        out.append({"product": product, "total": round(sum(locs.values()), 4),
                    "locations": [{"location": l, "quantity": q} for l, q in sorted(locs.items())],
                    "serial_numbers": serials})
    return {"products": out}


@app.get("/api/aging-report")
def aging_report(threshold_days: Optional[int] = None):
    wb = ExcelService.load_live_workbook()
    rows = ExcelService.aging_batches(wb, threshold_days or AGING_THRESHOLD_DAYS)
    return {"threshold_days": threshold_days or AGING_THRESHOLD_DAYS, "batches": rows}


@app.get("/api/stock-by-location")
def stock_by_location():
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    totals: Dict[str, float] = {}
    for product, locs in summary.items():
        for loc, q in locs.items():
            totals[loc] = round(totals.get(loc, 0) + q, 4)
    # Dashboard location chart is intentionally limited to the approved warehouse master list.
    # Legacy/free-text locations remain in the ledger for auditability but cannot create new chart categories.
    return {"locations": [{"location": l, "quantity": round(totals.get(l, 0), 4)}
                          for l in WAREHOUSE_OPTIONS if totals.get(l, 0) > 0]}


@app.get("/api/dashboard-summary")
def dashboard_summary():
    wb = ExcelService.load_live_workbook()
    summary = ExcelService.stock_summary(wb)
    all_batches = [b for b in ExcelService.all_batches(wb) if (b["Qty Remaining"] or 0) > 0]
    aging = ExcelService.aging_batches(wb)
    total_qty = round(sum(b["Qty Remaining"] for b in all_batches), 4)
    return {
        "summary": {
            "total_products": len(summary),
            "total_locations": len(ExcelService.known_locations(wb)),
            "total_quantity": total_qty,
            "active_batches": len(all_batches),
            "aging_alerts": len(aging),
            "aging_threshold_days": AGING_THRESHOLD_DAYS,
        }
    }


@app.get("/api/download")
def download_excel():
    try:
        _, raw = _get_file_meta()
        if raw is None:
            wb = ExcelService.new_workbook()
            buf = io.BytesIO()
            wb.save(buf)
            raw = buf.getvalue()
    except Exception as e:
        raise HTTPException(500, f"Could not read the workbook: {e}")
    return StreamingResponse(
        io.BytesIO(raw),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=inventory_report.xlsx"},
    )


@app.get("/api/download-url")
def download_url_route():
    return {"url": get_download_url()}
