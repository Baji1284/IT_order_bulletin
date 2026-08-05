#!/usr/bin/env python3
"""
FM Orders Bulletin - daily PDF email ? Gemini extraction ? Google Sheet.

Designed as a Cloud Run Job entrypoint (no HTTP server).
All secrets are injected via environment variables from Secret Manager.
"""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import sys
from datetime import date, datetime, timedelta
from calendar import monthrange
from email.message import EmailMessage
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import openpyxl
from google import genai
from google.auth.transport.requests import Request
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    stream=sys.stdout,
)
logger = logging.getLogger("fm-orders-job")

# ---------------------------------------------------------------------------
# Configuration (env / Secret Manager)
# ---------------------------------------------------------------------------

GMAIL_USER_EMAIL = os.environ.get(
    "GMAIL_USER_EMAIL", "itdocumentation@forbesmarshall.com"
)
EMAIL_SUBJECT = os.environ.get("EMAIL_SUBJECT", "FM Orders Bulletin")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.1-pro-preview").strip()
GEMINI_MAX_OUTPUT_TOKENS = int(os.environ.get("GEMINI_MAX_OUTPUT_TOKENS", "65535"))
# Thinking tokens come out of the same budget as the answer. Keeping each reply
# small (see GEMINI_ROW_BATCH) leaves room to think, and thinking hard matters:
# at "low" the model left roughly a third of the cells empty.
GEMINI_THINKING_LEVEL = os.environ.get("GEMINI_THINKING_LEVEL", "high").strip().lower()
# The metrics JSON is asked for in row batches so one reply never hits the output
# token ceiling and comes back truncated (= unparseable).
GEMINI_ROW_BATCH = int(os.environ.get("GEMINI_ROW_BATCH", "25"))
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "").strip()
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
GOOGLE_APPLICATION_CREDENTIALS = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS", ""
).strip()
IST = ZoneInfo("Asia/Kolkata")
MONTHLY_TAB_SUFFIX = os.environ.get("MONTHLY_TAB_SUFFIX", "_order_bulletin").strip()

# Formatted Order Bulletin template (Google Sheet matching the xlsx layout)
TEMPLATE_SHEET_ID = os.environ.get("TEMPLATE_SHEET_ID", "").strip()
TEMPLATE_WORKSHEET_NAME = os.environ.get("TEMPLATE_WORKSHEET_NAME", "Template").strip()
TEMPLATE_WORKSHEET_GID = os.environ.get("TEMPLATE_WORKSHEET_GID", "1145057086").strip()
# Dynamic current-month block (APRIL'26 section in the template = cols X-AA)
TEMPLATE_MONTH_HEADER_CELL = os.environ.get("TEMPLATE_MONTH_HEADER_CELL", "X2").strip()
TEMPLATE_HEADER_ROW = int(os.environ.get("TEMPLATE_HEADER_ROW", "3"))
TEMPLATE_DATA_START_ROW = int(os.environ.get("TEMPLATE_DATA_START_ROW", "4"))
TEMPLATE_COL_PROJ_INCL = os.environ.get("TEMPLATE_COL_PROJ_INCL", "X").strip()
TEMPLATE_COL_PROJ_DIGI = os.environ.get("TEMPLATE_COL_PROJ_DIGI", "Y").strip()
TEMPLATE_COL_ACH_INCL = os.environ.get("TEMPLATE_COL_ACH_INCL", "Z").strip()
TEMPLATE_COL_ACH_DIGI = os.environ.get("TEMPLATE_COL_ACH_DIGI", "AA").strip()
# Previous-month block (MARCH'26 section in the template = cols M-P)
TEMPLATE_PREV_MONTH_HEADER_CELL = os.environ.get(
    "TEMPLATE_PREV_MONTH_HEADER_CELL", "M2"
).strip()
TEMPLATE_COL_PREV_MONTH_PROJ_INCL = os.environ.get(
    "TEMPLATE_COL_PREV_MONTH_PROJ_INCL", "M"
).strip()
TEMPLATE_COL_PREV_MONTH_PROJ_DIGI = os.environ.get(
    "TEMPLATE_COL_PREV_MONTH_PROJ_DIGI", "N"
).strip()
TEMPLATE_COL_PREV_MONTH_ACH_INCL = os.environ.get(
    "TEMPLATE_COL_PREV_MONTH_ACH_INCL", "O"
).strip()
TEMPLATE_COL_PREV_MONTH_ACH_DIGI = os.environ.get(
    "TEMPLATE_COL_PREV_MONTH_ACH_DIGI", "P"
).strip()
# Fiscal-year section labels above the target / YTD blocks
TEMPLATE_CUR_FY_LABEL_CELL = os.environ.get("TEMPLATE_CUR_FY_LABEL_CELL", "I2").strip()
TEMPLATE_YTD_FY_LABEL_CELL = os.environ.get("TEMPLATE_YTD_FY_LABEL_CELL", "R2").strip()
TEMPLATE_COL_SEGMENT = os.environ.get("TEMPLATE_COL_SEGMENT", "A").strip()
TEMPLATE_COL_DIVISION = os.environ.get("TEMPLATE_COL_DIVISION", "B").strip()
# Previous fiscal-year block headers (the full Apr-Mar year that has closed)
TEMPLATE_PREV_FY_LABEL_CELL = os.environ.get("TEMPLATE_PREV_FY_LABEL_CELL", "C2").strip()
TEMPLATE_COL_PREV_FY_TOTAL_ACHMNT = os.environ.get(
    "TEMPLATE_COL_PREV_FY_TOTAL_ACHMNT", "C"
).strip()
TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE = os.environ.get(
    "TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE", "D"
).strip()
TEMPLATE_COL_PREV_FY_AVG = os.environ.get("TEMPLATE_COL_PREV_FY_AVG", "E").strip()
TEMPLATE_COL_PREV_FY_GROWTH = os.environ.get("TEMPLATE_COL_PREV_FY_GROWTH", "F").strip()
# Current fiscal-year block headers that name the months covered so far
TEMPLATE_COL_FY_TOTAL_ACHMNT = os.environ.get("TEMPLATE_COL_FY_TOTAL_ACHMNT", "R").strip()
TEMPLATE_COL_FY_DIGI_SUSTENANCE = os.environ.get(
    "TEMPLATE_COL_FY_DIGI_SUSTENANCE", "S"
).strip()
TEMPLATE_COL_FY_GROWTH = os.environ.get("TEMPLATE_COL_FY_GROWTH", "U").strip()

# Which PDF line feeds which template field (Vertical / Vertical Supply / tag / target)
MAPPING_FILE = os.environ.get(
    "MAPPING_FILE", "ORDERS BULLETIN_data_mapping.xlsx"
).strip()
MAPPING_SHEET = os.environ.get("MAPPING_SHEET", "").strip()
MAPPING_TOTAL_LABELS = {"FMPL TOTAL", "JV TOTAL"}

# Extra PDF totals used by derived template rows (not in the mapping workbook)
EXTRA_PDF_TOTAL_SECTIONS = [
    {
        "section": "Grand Total",
        "vertical": "Grand Total",
        "detail_lines": [],
    },
    {
        "section": "INTERCO-JV",
        "vertical": "INTERCO",
        "detail_lines": [],
    },
]
CORE_MECH_SUM_DIVISIONS = ("MECH STDS", "BOILERS", "ENERGY SERVICES")
CORE_MECH_SUM_TARGET = "MECH STDS + BOILERS + ENERGY SERVICES"

# One spreadsheet file per month, holding a dated tab for each day of that month
DRIVE_PARENT_FOLDER_ID = os.environ.get("DRIVE_PARENT_FOLDER_ID", "").strip()
# Optional: open this spreadsheet directly (skips name search / template copy).
# Useful when the monthly file already exists but Drive search cannot see it.
MONTHLY_SPREADSHEET_ID = os.environ.get("MONTHLY_SPREADSHEET_ID", "").strip()
# Service accounts own no Drive storage, so new files must be owned by a real user
# (domain-wide delegation) or live in a Shared Drive.
DRIVE_IMPERSONATE_USER = os.environ.get("DRIVE_IMPERSONATE_USER", "").strip()
MONTHLY_FILE_SHARE_EMAILS = [
    email.strip()
    for email in os.environ.get("MONTHLY_FILE_SHARE_EMAILS", "").split(",")
    if email.strip()
]

# Who gets the "today's bulletin is ready" mail with the link to the monthly file
NOTIFY_EMAILS = [
    email.strip()
    for email in re.split(
        r"[,\s]+",
        os.environ.get("NOTIFY_EMAILS", "bssali@forbesmarshall.com"),
    )
    if email.strip()
]
NOTIFY_FROM_EMAIL = os.environ.get("NOTIFY_FROM_EMAIL", GMAIL_USER_EMAIL).strip()

# Optional backfill: YYYY-MM-DD in IST. When set, fetch that day's bulletin email
# and write the matching dated tab (skips notification mail).
RUN_DATE = os.environ.get("RUN_DATE", "").strip()
# When true, only rewrite dynamic month/year headers on existing dated tabs
# (no Gmail / Gemini). Use with RUN_DATE for one tab, or alone to repair the
# whole current month file.
HEADER_REPAIR_ONLY = os.environ.get("HEADER_REPAIR_ONLY", "").strip().lower() in {
    "1",
    "true",
    "yes",
}

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
GMAIL_SEND_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_MIME = "application/vnd.google-apps.spreadsheet"

SYSTEM_INSTRUCTION = """You are an expert financial data extraction AI. I am providing a PDF containing a complex, multi-tier financial matrix (an Orders Bulletin) with hierarchical rows, vertically merged cells, and multi-level column headers.

Your task is to extract this data into a strict 2D JSON array (a list of lists) that perfectly maps the visual grid of the PDF so it can be written directly into a spreadsheet.

Follow these rules EXACTLY:

1. GRID & STRUCTURE:
- Output a 2D JSON array where each inner list represents one horizontal row in the PDF.
- The first 1 or 2 rows must be the column headers.
- Clean up PDF rotation and extraction artifacts in the headers (e.g., convert "Vercal MarketVercal SupplyDiv" to "Vertical Market | Vertical Supply | Div", and "Digi25-26" to "Digi 25-26"). However, maintain the EXACT column count and visual alignment of the original PDF.

2. HIERARCHICAL ROWS & MERGED CELLS:
- The left side contains hierarchical categories (e.g., PROCESS -> DOMESTIC-FMPL -> SSD+PAPER).
- Map these exactly to the columns they visually occupy.
- CRITICAL: If a row header is visually merged across multiple rows (e.g., "PROCESS" spans down 5 rows), place the word "PROCESS" in the first row's cell, and use empty strings "" for the subsequent merged rows. Do not repeat the word.

3. NUMERICAL PRECISION:
- Extract all numbers, decimals, and percentages EXACTLY as they appear in the cells.
- Preserve all negative signs (e.g., "-0.54", "-86.28%").
- Preserve exact decimal places (usually 2). Do not round, truncate, or alter any numbers.
- Do not recalculate or "fix" totals. Extract them exactly as printed.

4. FOOTER:
- Ignore the footer text at the very bottom (e.g., "FM Orders Bulletin (value In Crs) updated at...").

5. OUTPUT FORMAT:
- Output ONLY the raw JSON 2D array.
- DO NOT wrap it in markdown code blocks (no ```json or ```).
- DO NOT include any explanations, greetings, or conversational text.
- The very first character of your response must be "[" and the last must be "]".
"""

METRICS_SYSTEM_INSTRUCTION = """You are an expert financial data extraction AI for Forbes Marshall Orders Bulletins.

You receive:
1. The Orders Bulletin PDF.
2. TEMPLATE_COLUMNS: the spreadsheet columns to fill, each with its column letter, the
   section label above it (fiscal year or month block) and its full header text.
3. TEMPLATE_ROWS: the spreadsheet rows to fill, each with its row number, market segment
   and division label.

Your task: for EVERY template row, map the matching PDF line item and return the value for
EVERY template column, so that all numbers present in the bulletin land in the sheet.

Return ONLY a JSON object with this exact shape:
{
  "month_header": "JULY'26",
  "month_abbr": "Jul'26",
  "rows": [
    { "row": 4, "tag": "SSD+PAPER", "values": { "C": "12.34", "D": "1.23", "E": "", "I": "150.00" } }
  ]
}

RULES:
1. Identify the bulletin month from the PDF itself (title / headers / "updated at" date).
   month_header = FULL MONTH NAME uppercase + apostrophe + 2-digit year (e.g. "JULY'26").
   month_abbr = 3-letter title-case month + apostrophe + 2-digit year (e.g. "Jul'26").
2. Use the "row" numbers exactly as given in TEMPLATE_ROWS. Never invent row numbers.
3. Keys inside "values" MUST be the column letters from TEMPLATE_COLUMNS.
4. Locating the PDF line for a TEMPLATE_ROWS item:
   - When the item has "pdf_tag", read EXACTLY that line, inside the section named by
     "pdf_section". Echo the tag back in "tag". Return ONLY that line's own numbers:
     never add, merge or substitute another line, even if two items share a template row.
   - When the item has no "pdf_tag", match on the segment + division labels instead,
     including total/summary rows (e.g. "PROCESS DOMESTIC TOTAL ( FMPL) :"), and set
     "tag" to "".
   - An item appears once per PDF tag, so the same "row" may legitimately appear twice
     with different tags. Return one object per item, keeping them separate.
5. Match each template column to the correct PDF column using the section label AND header text:
   - Fiscal-year blocks (e.g. "2025- 2026", "2026 - 2027") map to the corresponding
     year-to-date / total / target / growth / % achievement columns in the PDF.
   - Month blocks map to that month's Projections and Achievement columns
     (incl. Digital Business and for Digital Sustenance).
   - "Achmnt (AVG)", "% Growth over ...", "% Achmnt to Target", "Target", "Target / Month"
     and "Pro rata Target" must come from the equivalent PDF columns.
6. NUMERICAL INTEGRITY: copy values EXACTLY as printed in the PDF. Never round, reformat,
   recalculate, or derive a value. Keep decimals, negative signs and % signs as shown
   (e.g. "-0.54", "-86.28%"). Output every value as a JSON string.
7. If the PDF has no value for a given row/column combination, use an empty string "".
8. Include an entry for every item in TEMPLATE_ROWS, even if all its values are empty.
9. Output ONLY valid JSON. No markdown fences, no commentary, no explanation.
"""

SECTION_TOTALS_SYSTEM_INSTRUCTION = """You are an expert financial data extraction AI for Forbes Marshall Orders Bulletins.

You receive:
1. The Orders Bulletin PDF.
2. TEMPLATE_COLUMNS: the spreadsheet columns to fill, each with its column letter, the
   section label above it (fiscal year or month block) and its full header text.
3. SECTIONS: the bulletin sections whose TOTAL line you must read. Each one gives its
   "section" name, the "vertical" it sits under, and "detail_lines", the PDF line items
   that belong to it. Use the vertical and the detail lines to identify the right block,
   because names like "JV" repeat across verticals.

Your task: for each section, find its total line in the PDF (e.g. "PROCESS DOMESTIC TOTAL
( FMPL) :", "INTOPS - FMPL TOTAL :", "OPC DOMESTIC TOTAL") and return that line's value for
every template column.

Special sections:
- "Grand Total": the bulletin's overall Grand Total / GROUP TOTAL line (this is the
  figure that includes INTERCO-JV when present).
- "INTERCO-JV": the Total line of the INTERCO-JV market segment.

Return ONLY a JSON object with this exact shape:
{
  "section_totals": [
    { "section": "PROCESS DOMESTIC-FMPL", "values": { "C": "12.34", "R": "5.00" } }
  ]
}

RULES:
1. Echo the "section" string exactly as given in SECTIONS.
1a. The total must be the one covering exactly that section's detail lines, in that vertical.
2. Keys inside "values" MUST be the column letters from TEMPLATE_COLUMNS.
3. Read the section's own total line only. Never add up the detail lines yourself and never
   borrow another section's total.
4. Match each template column to the correct PDF column using the section label AND header
   text, exactly as for the detail rows.
5. NUMERICAL INTEGRITY: copy values EXACTLY as printed in the PDF. Never round, reformat,
   recalculate or derive a value. Keep decimals, negative signs and % signs as shown.
   Output every value as a JSON string.
6. If the PDF has no total for a section, return it with all values as empty strings "".
7. Output ONLY valid JSON. No markdown fences, no commentary, no explanation.
"""


def _require_env(name: str, value: str) -> str:
    if not value:
        raise RuntimeError(f"Required environment variable {name} is not set")
    return value


def _load_sa_info() -> dict[str, Any]:
    if GOOGLE_SERVICE_ACCOUNT_JSON:
        return json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    if GOOGLE_APPLICATION_CREDENTIALS and os.path.isfile(
        GOOGLE_APPLICATION_CREDENTIALS
    ):
        with open(GOOGLE_APPLICATION_CREDENTIALS, encoding="utf-8") as fh:
            return json.load(fh)
    raise RuntimeError(
        "Set GOOGLE_SERVICE_ACCOUNT_JSON (Secret Manager) or "
        "GOOGLE_APPLICATION_CREDENTIALS (file path)"
    )


def load_gmail_credentials() -> service_account.Credentials:
    """SA credentials with domain-wide delegation to read the target mailbox."""
    info = _load_sa_info()
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=GMAIL_SCOPES
    )
    return creds.with_subject(GMAIL_USER_EMAIL)


def load_gmail_send_credentials() -> service_account.Credentials:
    """SA credentials delegated to send mail as the notification sender."""
    info = _load_sa_info()
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=GMAIL_SEND_SCOPES
    )
    return creds.with_subject(NOTIFY_FROM_EMAIL or GMAIL_USER_EMAIL)


def load_sheets_credentials() -> service_account.Credentials:
    """SA credentials for Sheets (share the sheet with the SA client_email)."""
    info = _load_sa_info()
    return service_account.Credentials.from_service_account_info(
        info, scopes=SHEETS_SCOPES
    )


def get_gmail_service(creds: service_account.Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def find_latest_message_id(gmail, subject: str, on_day: date | None = None) -> str:
    """
    Find the FM Orders Bulletin email.

    When on_day is set (backfill), search that calendar day in the mailbox.
    Otherwise use the most recent message from the last 2 days.
    """
    if on_day is not None:
        # Gmail after/before are date boundaries; cover the full on_day.
        day_after = on_day + timedelta(days=1)
        day_before = on_day - timedelta(days=1)
        queries = [
            (
                f'subject:"{subject}" has:attachment filename:pdf '
                f'after:{day_before.strftime("%Y/%m/%d")} '
                f'before:{day_after.strftime("%Y/%m/%d")}'
            ),
            (
                f'subject:"{subject}" '
                f'after:{day_before.strftime("%Y/%m/%d")} '
                f'before:{day_after.strftime("%Y/%m/%d")}'
            ),
        ]
    else:
        queries = [
            f'subject:"{subject}" has:attachment filename:pdf newer_than:2d',
            f'subject:"{subject}" newer_than:2d',
        ]

    messages: list[dict[str, Any]] = []
    used_query = ""
    for query in queries:
        logger.info("Searching Gmail: %s", query)
        result = (
            gmail.users()
            .messages()
            .list(userId="me", q=query, maxResults=10)
            .execute()
        )
        messages = result.get("messages") or []
        used_query = query
        if messages:
            break

    if not messages:
        target = on_day.isoformat() if on_day else "the last 2 days"
        raise RuntimeError(
            f'No email found with subject "{subject}" for {target} '
            f"(last query: {used_query})"
        )

    msg_id = messages[0]["id"]
    logger.info("Using message id=%s", msg_id)
    return msg_id


def _walk_parts(payload: dict[str, Any]) -> list[dict[str, Any]]:
    parts = payload.get("parts") or []
    if not parts:
        return [payload]
    out: list[dict[str, Any]] = []
    for part in parts:
        if part.get("parts"):
            out.extend(_walk_parts(part))
        else:
            out.append(part)
    return out


def download_pdf_attachment(gmail, message_id: str) -> bytes:
    message = (
        gmail.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )
    parts = _walk_parts(message.get("payload") or {})
    pdf_parts: list[tuple[str, str]] = []
    for part in parts:
        filename = (part.get("filename") or "").lower()
        mime = (part.get("mimeType") or "").lower()
        body = part.get("body") or {}
        att_id = body.get("attachmentId")
        if not att_id:
            continue
        if filename.endswith(".pdf") or mime == "application/pdf":
            pdf_parts.append((filename or "attachment.pdf", att_id))

    if not pdf_parts:
        raise RuntimeError(f"No PDF attachment found on message {message_id}")

    filename, att_id = pdf_parts[0]
    logger.info("Downloading PDF attachment: %s", filename)
    att = (
        gmail.users()
        .messages()
        .attachments()
        .get(userId="me", messageId=message_id, id=att_id)
        .execute()
    )
    data = att.get("data")
    if not data:
        raise RuntimeError("Attachment payload was empty")
    return base64.urlsafe_b64decode(data)


def _generation_config(system_instruction: str) -> types.GenerateContentConfig:
    kwargs: dict[str, Any] = {
        "system_instruction": system_instruction,
        "temperature": 0.0,
        "response_mime_type": "application/json",
        "max_output_tokens": GEMINI_MAX_OUTPUT_TOKENS,
    }
    if GEMINI_THINKING_LEVEL:
        kwargs["thinking_config"] = types.ThinkingConfig(
            thinking_level=GEMINI_THINKING_LEVEL
        )
    return types.GenerateContentConfig(**kwargs)


def _finish_reason(response: Any) -> str:
    candidates = getattr(response, "candidates", None) or []
    if not candidates:
        return ""
    reason = getattr(candidates[0], "finish_reason", None)
    return getattr(reason, "name", None) or str(reason or "")


def _generate_json(
    client: genai.Client,
    pdf_bytes: bytes,
    prompt: str,
    system_instruction: str,
    label: str,
) -> Any:
    """Send the PDF to Gemini and parse the JSON reply, retrying once if truncated."""
    last_error: Exception | None = None
    for attempt in (1, 2):
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_bytes(
                            data=pdf_bytes, mime_type="application/pdf"
                        ),
                        types.Part.from_text(text=prompt),
                    ],
                )
            ],
            config=_generation_config(system_instruction),
        )
        reason = _finish_reason(response)
        if reason and reason.upper() not in {"STOP", "FINISH_REASON_STOP"}:
            logger.warning("%s: Gemini stopped with reason=%s", label, reason)
        try:
            return _parse_json_response(response.text)
        except RuntimeError as exc:
            last_error = exc
            logger.warning("%s: attempt %d gave unusable JSON: %s", label, attempt, exc)

    raise RuntimeError(f"{label}: Gemini returned no valid JSON in 2 attempts") from last_error


def extract_table_with_gemini(pdf_bytes: bytes) -> list[list[Any]]:
    _require_env("GEMINI_API_KEY", GEMINI_API_KEY)
    client = genai.Client(api_key=GEMINI_API_KEY)

    logger.info(
        "Sending PDF (%d bytes) to Gemini model=%s", len(pdf_bytes), GEMINI_MODEL
    )
    data = _generate_json(
        client,
        pdf_bytes,
        "Extract the Orders Bulletin table from this PDF.",
        SYSTEM_INSTRUCTION,
        "raw table",
    )

    if not isinstance(data, list):
        raise RuntimeError("Expected a JSON array of arrays from Gemini")

    normalized: list[list[Any]] = []
    for row in data:
        if row is None:
            normalized.append([])
        elif isinstance(row, list):
            normalized.append(["" if c is None else c for c in row])
        else:
            raise RuntimeError(
                f"Expected each row to be a list, got {type(row).__name__}"
            )

    if not normalized:
        raise RuntimeError("Gemini returned an empty table")

    logger.info("Extracted %d rows from PDF", len(normalized))
    return normalized


def _now_ist(now: datetime | None = None) -> datetime:
    when = now or datetime.now(IST)
    if when.tzinfo is None:
        return when.replace(tzinfo=IST)
    return when.astimezone(IST)


def _parse_run_date() -> datetime | None:
    """Parse RUN_DATE=YYYY-MM-DD into an IST datetime at noon that day."""
    if not RUN_DATE:
        return None
    try:
        day = date.fromisoformat(RUN_DATE)
    except ValueError as exc:
        raise RuntimeError(
            f"RUN_DATE must be YYYY-MM-DD, got {RUN_DATE!r}"
        ) from exc
    return datetime(day.year, day.month, day.day, 12, 0, 0, tzinfo=IST)


def run_moment() -> datetime:
    """Effective 'today' for this execution (backfill date or current IST now)."""
    return _parse_run_date() or _now_ist()


def current_month_tab_name(now: datetime | None = None) -> str:
    """Return monthly worksheet title, e.g. Jul-2026_order_bulletin (IST)."""
    when = _now_ist(now)
    return f"{when.strftime('%b-%Y')}{MONTHLY_TAB_SUFFIX}"


def day_tab_name(day: date) -> str:
    """Daily sub-tab title, e.g. 01/07/2026."""
    return day.strftime("%d/%m/%Y")


def month_day_tab_names(year: int, month: int) -> list[str]:
    """All daily tab names for a calendar month (01/MM/YYYY - last/MM/YYYY)."""
    _, last_day = monthrange(year, month)
    return [day_tab_name(date(year, month, d)) for d in range(1, last_day + 1)]


def _worksheet_titles(spreadsheet: gspread.Spreadsheet) -> set[str]:
    return {ws.title for ws in spreadsheet.worksheets()}


def ensure_month_tabs(spreadsheet: gspread.Spreadsheet, now: datetime | None = None) -> str:
    """
    Ensure the monthly main tab and every daily date tab for the IST month exist.

    Google Sheets cannot nest tabs, so daily "sub-tabs" are sibling worksheets
    named DD/MM/YYYY under the same spreadsheet as Jul-2026_order_bulletin.
    """
    when = _now_ist(now)
    monthly_name = current_month_tab_name(when)
    day_names = month_day_tab_names(when.year, when.month)

    existing = _worksheet_titles(spreadsheet)
    to_create: list[str] = []
    if monthly_name not in existing:
        to_create.append(monthly_name)
    for name in day_names:
        if name not in existing:
            to_create.append(name)

    if to_create:
        logger.info(
            "Creating %d missing tab(s) for %s: monthly=%s, days=%d",
            len(to_create),
            when.strftime("%b-%Y"),
            monthly_name not in existing,
            sum(1 for n in to_create if n != monthly_name),
        )
        # Batch create to avoid many sequential API calls
        requests = [
            {
                "addSheet": {
                    "properties": {
                        "title": title,
                        "gridProperties": {"rowCount": 200, "columnCount": 40},
                    }
                }
            }
            for title in to_create
        ]
        spreadsheet.batch_update({"requests": requests})
    else:
        logger.info(
            "All monthly/daily tabs already exist for %s (%d day tabs + main)",
            monthly_name,
            len(day_names),
        )

    return monthly_name


def _clear_and_write(worksheet: gspread.Worksheet, values: list[list[str]]) -> None:
    logger.info(
        "Clearing worksheet %r and writing %d rows",
        worksheet.title,
        len(values),
    )
    worksheet.clear()
    worksheet.update(values, value_input_option="RAW")


def write_to_sheet(rows: list[list[Any]]) -> None:
    sheet_id = _require_env("GOOGLE_SHEET_ID", GOOGLE_SHEET_ID)
    sheets_creds = load_sheets_credentials()
    if not sheets_creds.valid:
        sheets_creds.refresh(Request())

    gc = gspread.authorize(sheets_creds)
    spreadsheet = gc.open_by_key(sheet_id)

    when = _now_ist()
    monthly_name = ensure_month_tabs(spreadsheet, when)
    today_name = day_tab_name(when.date())

    values = [["" if c is None else str(c) for c in row] for row in rows]

    # Daily sub-tab: today's bulletin snapshot for this date
    daily_ws = spreadsheet.worksheet(today_name)
    _clear_and_write(daily_ws, values)

    # Monthly main tab: latest cumulative bulletin for the month
    monthly_ws = spreadsheet.worksheet(monthly_name)
    _clear_and_write(monthly_ws, values)

    logger.info(
        "Google Sheet updated successfully on daily tab %r and monthly main %r",
        today_name,
        monthly_name,
    )


def _merge_json_documents(docs: list[Any]) -> Any:
    """Gemini sometimes answers with several JSON documents in a row."""
    if len(docs) == 1:
        return docs[0]

    logger.warning("Gemini returned %d JSON documents; merging them", len(docs))
    if all(isinstance(doc, list) for doc in docs):
        merged_list: list[Any] = []
        for doc in docs:
            merged_list.extend(doc)
        return merged_list
    if all(isinstance(doc, dict) for doc in docs):
        merged: dict[str, Any] = dict(docs[0])
        rows: list[Any] = []
        for doc in docs:
            part = doc.get("rows")
            if isinstance(part, list):
                rows.extend(part)
        if rows:
            merged["rows"] = rows
        return merged
    return docs[0]


def _parse_json_response(text: str) -> Any:
    text = (text or "").strip()
    if not text:
        raise RuntimeError("Gemini returned an empty response")
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    decoder = json.JSONDecoder()
    docs: list[Any] = []
    pos = 0
    while pos < len(text):
        while pos < len(text) and text[pos] in " \t\r\n,":
            pos += 1
        if pos >= len(text):
            break
        try:
            doc, pos = decoder.raw_decode(text, pos)
        except json.JSONDecodeError as exc:
            if docs:
                logger.warning("Ignoring unparseable tail after %d document(s): %s", len(docs), exc)
                break
            logger.error("Gemini raw response (truncated): %s", text[:2000])
            raise RuntimeError(f"Gemini did not return valid JSON: {exc}") from exc
        docs.append(doc)

    if not docs:
        raise RuntimeError("Gemini returned no JSON document")
    return _merge_json_documents(docs)


def _month_labels_from_ist(now: datetime | None = None) -> tuple[str, str]:
    when = _now_ist(now)
    full = when.strftime("%B").upper() + "'" + when.strftime("%y")
    abbr = when.strftime("%b") + "'" + when.strftime("%y")
    return full, abbr


def _previous_month(now: datetime | None = None) -> datetime:
    when = _now_ist(now)
    return when.replace(day=1) - timedelta(days=1)


def _closed_fiscal_year_end(now: datetime | None = None) -> int:
    """
    Year of the March that ended the last complete fiscal year.

    Fiscal years run Apr-Mar, so a run in Jul 2026 sits in FY 2026-27 and the
    closed one ended Mar 2026; a run in Jan 2027 is still in FY 2026-27, so the
    closed year is that same Mar 2026.
    """
    when = _now_ist(now)
    return when.year if when.month >= 4 else when.year - 1


def _col_to_index(col: str) -> int:
    """A -> 1, Z -> 26, AA -> 27."""
    idx = 0
    for ch in col.strip().upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _index_to_col(idx: int) -> str:
    col = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        col = chr(ord("A") + rem) + col
    return col


def _request_metrics_batch(
    client: genai.Client,
    pdf_bytes: bytes,
    template_columns: list[dict[str, str]],
    rows_batch: list[dict[str, Any]],
) -> dict[str, Any]:
    """One Gemini call for a slice of template rows."""
    context = json.dumps(
        {"TEMPLATE_COLUMNS": template_columns, "TEMPLATE_ROWS": rows_batch},
        ensure_ascii=False,
    )
    label = f"metrics rows {rows_batch[0]['row']}-{rows_batch[-1]['row']}"
    data = _generate_json(
        client,
        pdf_bytes,
        "Fill the Order Bulletin template. Return a value for every "
        "column of every row listed below.\n\n" + context,
        METRICS_SYSTEM_INSTRUCTION,
        label,
    )
    if not isinstance(data, dict):
        raise RuntimeError(f"{label}: expected a JSON object from Gemini")
    return data


def extract_metrics_with_gemini(
    pdf_bytes: bytes,
    template_columns: list[dict[str, str]],
    template_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """Ask Gemini to fill every template column for every template row."""
    _require_env("GEMINI_API_KEY", GEMINI_API_KEY)
    client = genai.Client(api_key=GEMINI_API_KEY)

    batches = [
        template_rows[i : i + GEMINI_ROW_BATCH]
        for i in range(0, len(template_rows), GEMINI_ROW_BATCH)
    ]
    logger.info(
        "Extracting template metrics via Gemini model=%s (%d cols x %d rows in %d batch(es))",
        GEMINI_MODEL,
        len(template_columns),
        len(template_rows),
        len(batches),
    )

    valid_rows = {int(item["row"]) for item in template_rows}
    valid_cols = {item["col"] for item in template_columns}
    fallback_full, fallback_abbr = _month_labels_from_ist()
    month_header = ""
    month_abbr = ""
    normalized_rows: list[dict[str, Any]] = []

    for index, rows_batch in enumerate(batches, start=1):
        data = _request_metrics_batch(client, pdf_bytes, template_columns, rows_batch)
        month_header = month_header or str(data.get("month_header") or "").strip()
        month_abbr = month_abbr or str(data.get("month_abbr") or "").strip()

        batch_rows = 0
        for row in data.get("rows") or []:
            if not isinstance(row, dict):
                continue
            try:
                row_num = int(row.get("row"))
            except (TypeError, ValueError):
                continue
            if row_num not in valid_rows:
                continue
            values = row.get("values")
            if not isinstance(values, dict):
                continue
            clean = {
                str(col).strip().upper(): ("" if val is None else str(val))
                for col, val in values.items()
                if str(col).strip().upper() in valid_cols
            }
            if clean:
                normalized_rows.append(
                    {
                        "row": row_num,
                        "tag": _clean_text(row.get("tag")),
                        "values": clean,
                    }
                )
                batch_rows += 1
        logger.info(
            "Batch %d/%d (rows %s-%s): %d items returned",
            index,
            len(batches),
            rows_batch[0]["row"],
            rows_batch[-1]["row"],
            batch_rows,
        )

    month_header = month_header or fallback_full
    month_abbr = month_abbr or fallback_abbr

    filled = sum(1 for r in normalized_rows for v in r["values"].values() if v != "")
    logger.info(
        "Extracted metrics for %s: %d items, %d populated cells",
        month_header,
        len(normalized_rows),
        filled,
    )
    return {
        "month_header": month_header,
        "month_abbr": month_abbr,
        "rows": normalized_rows,
    }


def describe_total_sections(mapping: dict[str, Any]) -> list[dict[str, Any]]:
    """The sections behind the company totals, plus Grand Total / INTERCO-JV."""
    names = sorted(
        {name for names in mapping["company_totals"].values() for name in names}
    )
    described: list[dict[str, Any]] = []
    for name in names:
        section = mapping["sections"].get(_norm_label(name))
        if section is None:
            logger.warning("Company total component %r is not a mapped section", name)
            described.append({"section": name, "vertical": "", "detail_lines": []})
            continue
        described.append(
            {
                "section": name,
                "vertical": section["vertical"],
                "detail_lines": [f["pdf_tag"] for f in section["fields"]],
            }
        )
    described.extend(EXTRA_PDF_TOTAL_SECTIONS)
    return described


def extract_section_totals_with_gemini(
    pdf_bytes: bytes,
    template_columns: list[dict[str, str]],
    sections: list[dict[str, Any]],
) -> dict[str, dict[str, str]]:
    """Read each section's own TOTAL line, which the company totals are built from."""
    _require_env("GEMINI_API_KEY", GEMINI_API_KEY)
    client = genai.Client(api_key=GEMINI_API_KEY)
    context = json.dumps(
        {"TEMPLATE_COLUMNS": template_columns, "SECTIONS": sections},
        ensure_ascii=False,
    )
    data = _generate_json(
        client,
        pdf_bytes,
        "Read the total line of each section listed below.\n\n" + context,
        SECTION_TOTALS_SYSTEM_INSTRUCTION,
        "section totals",
    )
    if not isinstance(data, dict):
        raise RuntimeError("section totals: expected a JSON object from Gemini")

    valid_cols = {item["col"] for item in template_columns}
    wanted = {_norm_label(s["section"]): s["section"] for s in sections}
    totals: dict[str, dict[str, str]] = {}
    for entry in data.get("section_totals") or []:
        if not isinstance(entry, dict):
            continue
        key = _norm_label(entry.get("section"))
        if key not in wanted:
            continue
        values = entry.get("values")
        if not isinstance(values, dict):
            continue
        totals[key] = {
            str(col).strip().upper(): ("" if val is None else str(val))
            for col, val in values.items()
            if str(col).strip().upper() in valid_cols
        }

    missing = [name for key, name in wanted.items() if key not in totals]
    if missing:
        logger.warning("No section total returned for: %s", ", ".join(missing))
    logger.info("Read section totals for %d of %d sections", len(totals), len(sections))
    return totals


def _parse_number(value: str) -> float | None:
    """Numbers as printed in the bulletin: 1,234.56 / -0.54 / (0.54)."""
    text = _clean_text(value).replace(",", "").replace("\u2212", "-")
    if not text or "%" in text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def _sum_values(values: list[str]) -> str:
    """
    Add up values that are plain numbers.

    Percentages cannot be added, so a column holding any of them stays empty
    rather than showing an invented figure.
    """
    present = [v for v in values if _clean_text(v)]
    if not present:
        return ""
    if any("%" in _clean_text(v) for v in present):
        return ""
    numbers = [_parse_number(v) for v in present]
    if any(n is None for n in numbers):
        return ""
    return f"{sum(n for n in numbers if n is not None):.2f}"


def _subtract_values(left: str, right: str) -> str:
    """left - right for plain bulletin numbers; blank if either side is unusable."""
    if not _clean_text(left):
        return ""
    if "%" in _clean_text(left) or ("%" in _clean_text(right) if right else False):
        return ""
    left_n = _parse_number(left)
    if left_n is None:
        return ""
    if not _clean_text(right):
        return f"{left_n:.2f}"
    right_n = _parse_number(right)
    if right_n is None:
        return ""
    return f"{left_n - right_n:.2f}"


def combine_row_items(
    items: list[dict[str, Any]], columns: list[str]
) -> list[dict[str, Any]]:
    """
    Collapse the per-tag items into one value set per template row.

    A row fed by a single PDF tag keeps its value verbatim. A row fed by several
    (INTOPS MECH STDS = SSD+PAPER + SSD+PAPER-SG) gets their sum.
    """
    grouped: dict[int, list[dict[str, Any]]] = {}
    for item in items:
        grouped.setdefault(int(item["row"]), []).append(item)

    combined: list[dict[str, Any]] = []
    for row_num, row_items in sorted(grouped.items()):
        if len(row_items) == 1:
            combined.append({"row": row_num, "values": row_items[0]["values"]})
            continue

        tags = [item.get("tag") or "?" for item in row_items]
        values = {
            col: _sum_values([item["values"].get(col, "") for item in row_items])
            for col in columns
        }
        logger.info("Row %d = sum of %s", row_num, " + ".join(tags))
        combined.append({"row": row_num, "values": values})
    return combined


def company_total_rows(
    mapping: dict[str, Any],
    company_rows: dict[str, int],
    section_totals: dict[str, dict[str, str]],
    columns: list[str],
) -> list[dict[str, Any]]:
    """FMPL TOTAL / JV TOTAL, each the sum of the section totals listed in the mapping."""
    rows: list[dict[str, Any]] = []
    for label, components in mapping["company_totals"].items():
        row_num = company_rows.get(label)
        if row_num is None:
            continue
        keys = [_norm_label(name) for name in components]
        missing = [
            name for name, key in zip(components, keys) if key not in section_totals
        ]
        if missing:
            logger.warning(
                "%s is missing section totals for: %s", label, ", ".join(missing)
            )
        values = {
            col: _sum_values(
                [section_totals.get(key, {}).get(col, "") for key in keys]
            )
            for col in columns
        }
        populated = sum(1 for v in values.values() if v)
        logger.info(
            "%s (row %d) = %s -> %d values",
            label,
            row_num,
            " + ".join(components),
            populated,
        )
        rows.append({"row": row_num, "values": values})
    return rows


def _divide_value(value: str, divisor: float) -> str:
    """Divide a bulletin number; leave blanks / percentages alone."""
    number = _parse_number(value)
    if number is None or divisor == 0:
        return ""
    return f"{number / divisor:.2f}"


def _metric_value_map(rows: list[dict[str, Any]]) -> dict[int, dict[str, str]]:
    return {int(row["row"]): dict(row.get("values") or {}) for row in rows}


def _metric_rows_from_map(value_map: dict[int, dict[str, str]]) -> list[dict[str, Any]]:
    return [{"row": row, "values": values} for row, values in sorted(value_map.items())]


def _segment_key(row: dict[str, Any]) -> str:
    return _norm_label(row.get("segment"))


def _division_key(row: dict[str, Any]) -> str:
    return _norm_label(row.get("division"))


def _find_template_row(
    template_rows: list[dict[str, Any]],
    *,
    segment_pred,
    division_pred=None,
    after_row: int = 0,
) -> dict[str, Any] | None:
    for row in template_rows:
        if row["row"] <= after_row:
            continue
        if not segment_pred(_segment_key(row)):
            continue
        if division_pred is not None and not division_pred(_division_key(row)):
            continue
        return row
    return None


def _is_process_domestic_fmpl_total(segment: str) -> bool:
    return segment == "PROCESS DOMESTIC TOTAL FMPL"


def _is_core_domestic_fmpl_total(segment: str) -> bool:
    return "CORE DOMESTIC FMPL TOTAL" in segment


def _is_digital_sustenance(segment: str) -> bool:
    return segment == "DIGITAL SUSTENANCE BUSINESS"


def _is_core_domestic_fmpl_segment(segment: str) -> bool:
    return segment == "CORE DOMESTIC FMPL"


def _apply_digital_sustenance_from_total(
    value_map: dict[int, dict[str, str]],
    *,
    source_row: int,
    target_row: int,
    label: str,
) -> None:
    """
    DIGITAL SUSTENANCE BUSINESS cells derived from the section TOTAL row:
      D  = TOTAL[D] as-is
      E  = TOTAL[D] / 12   (Achmnt AVG)
      S  = TOTAL[S] as-is  (current FY digital sustenance)
      Y  = TOTAL[Y] as-is  (current-month projections for digital sustenance)
      AA = TOTAL[AA] as-is (current-month achmnt for digital sustenance)
    """
    source = value_map.get(source_row) or {}
    digi = source.get(TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE, "")
    updates = {
        TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE: digi,
        TEMPLATE_COL_PREV_FY_AVG: _divide_value(digi, 12),
        TEMPLATE_COL_FY_DIGI_SUSTENANCE: source.get(
            TEMPLATE_COL_FY_DIGI_SUSTENANCE, ""
        ),
        TEMPLATE_COL_PROJ_DIGI: source.get(TEMPLATE_COL_PROJ_DIGI, ""),
        TEMPLATE_COL_ACH_DIGI: source.get(TEMPLATE_COL_ACH_DIGI, ""),
    }
    target = value_map.setdefault(target_row, {})
    target.update({col: val for col, val in updates.items() if val != ""})
    logger.info(
        "%s DIGITAL SUSTENANCE (row %d) <- TOTAL row %d: D=%r E=%r S=%r Y=%r AA=%r",
        label,
        target_row,
        source_row,
        updates[TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE],
        updates[TEMPLATE_COL_PREV_FY_AVG],
        updates[TEMPLATE_COL_FY_DIGI_SUSTENANCE],
        updates[TEMPLATE_COL_PROJ_DIGI],
        updates[TEMPLATE_COL_ACH_DIGI],
    )


def apply_derived_template_rows(
    template_rows: list[dict[str, Any]],
    metric_rows: list[dict[str, Any]],
    section_totals: dict[str, dict[str, str]],
    columns: list[str],
) -> list[dict[str, Any]]:
    """
    Post-Gemini / post Rule A+B calculations that the mapping sheet does not cover:

    - Process / CORE DIGITAL SUSTENANCE BUSINESS from the matching FMPL TOTAL
    - CORE MECH STDS + BOILERS + ENERGY SERVICES sum
    - GROUP TOTAL WITHOUT INTERCO = Grand Total - INTERCO-JV Total
    - GROUP TOTAL (including JV Interco) = Grand Total as printed
    """
    value_map = _metric_value_map(metric_rows)

    process_total = _find_template_row(
        template_rows, segment_pred=_is_process_domestic_fmpl_total
    )
    if process_total:
        digi = _find_template_row(
            template_rows,
            segment_pred=_is_digital_sustenance,
            after_row=process_total["row"],
        )
        if digi:
            _apply_digital_sustenance_from_total(
                value_map,
                source_row=process_total["row"],
                target_row=digi["row"],
                label="Process Domestic FMPL",
            )
        else:
            logger.warning(
                "No DIGITAL SUSTENANCE BUSINESS row after Process Domestic FMPL TOTAL"
            )
    else:
        logger.warning("PROCESS DOMESTIC TOTAL (FMPL) row not found in template")

    core_total = _find_template_row(
        template_rows, segment_pred=_is_core_domestic_fmpl_total
    )
    if core_total:
        digi = _find_template_row(
            template_rows,
            segment_pred=_is_digital_sustenance,
            after_row=core_total["row"],
        )
        if digi:
            _apply_digital_sustenance_from_total(
                value_map,
                source_row=core_total["row"],
                target_row=digi["row"],
                label="CORE DOMESTIC FMPL",
            )
        else:
            logger.warning(
                "No DIGITAL SUSTENANCE BUSINESS row after CORE DOMESTIC FMPL TOTAL"
            )
    else:
        logger.warning("CORE DOMESTIC FMPL TOTAL row not found in template")

    # CORE: Mech Stds + Boilers + Energy Services -> combined division row
    wanted = {_norm_label(name) for name in CORE_MECH_SUM_DIVISIONS}
    source_parts: dict[str, int] = {}
    target_row_num: int | None = None
    for row in template_rows:
        if not _is_core_domestic_fmpl_segment(_segment_key(row)):
            continue
        div = _division_key(row)
        if div == _norm_label(CORE_MECH_SUM_TARGET):
            target_row_num = row["row"]
        elif div in wanted:
            source_parts[div] = row["row"]

    if target_row_num is None:
        logger.warning("CORE row %r not found in template", CORE_MECH_SUM_TARGET)
    elif len(source_parts) < len(wanted):
        missing = sorted(wanted - set(source_parts))
        logger.warning(
            "CORE mech sum is missing division rows: %s", ", ".join(missing)
        )
    else:
        values = {
            col: _sum_values(
                [
                    (value_map.get(source_parts[div]) or {}).get(col, "")
                    for div in sorted(source_parts)
                ]
            )
            for col in columns
        }
        value_map[target_row_num] = values
        logger.info(
            "CORE %s (row %d) = %s -> %d values",
            CORE_MECH_SUM_TARGET,
            target_row_num,
            " + ".join(CORE_MECH_SUM_DIVISIONS),
            sum(1 for v in values.values() if v),
        )

    # GROUP TOTAL INCLUDING <- Grand Total as-is;
    # WITHOUT INTERCO <- Grand Total - INTERCO-JV
    without_row = _find_template_row(
        template_rows,
        segment_pred=lambda s: s == "GROUP TOTAL WITHOUT INTERCO",
    )
    with_row = _find_template_row(
        template_rows,
        segment_pred=lambda s: s.startswith("GROUP TOTAL") and "INCLUDING JV INTERCO" in s,
    )
    grand = section_totals.get(_norm_label("Grand Total")) or {}
    interco = section_totals.get(_norm_label("INTERCO-JV")) or {}

    if not grand:
        logger.warning("Grand Total was not returned by Gemini")

    if with_row and grand:
        merged = dict(value_map.get(with_row["row"]) or {})
        merged.update({col: val for col, val in grand.items() if val != ""})
        value_map[with_row["row"]] = merged
        logger.info(
            "GROUP TOTAL (including JV Interco) (row %d) <- Grand Total (%d values)",
            with_row["row"],
            sum(1 for v in merged.values() if v),
        )
    elif with_row is None:
        logger.warning("GROUP TOTAL (including JV Interco) row not found in template")

    if without_row is None:
        logger.warning("GROUP TOTAL WITHOUT INTERCO row not found in template")
    elif not grand:
        pass  # already logged
    else:
        if not interco:
            logger.warning(
                "INTERCO-JV total was not returned by Gemini; "
                "WITHOUT INTERCO will equal Grand Total where Interco is blank"
            )
        values = {
            col: _subtract_values(grand.get(col, ""), interco.get(col, ""))
            for col in columns
        }
        value_map[without_row["row"]] = values
        logger.info(
            "GROUP TOTAL WITHOUT INTERCO (row %d) = Grand Total - INTERCO-JV -> %d values",
            without_row["row"],
            sum(1 for v in values.values() if v),
        )

    return _metric_rows_from_map(value_map)


def _get_master_template_worksheet(spreadsheet: gspread.Spreadsheet) -> gspread.Worksheet:
    """Locate the blank formatted master tab (by name, then by gid)."""
    try:
        return spreadsheet.worksheet(TEMPLATE_WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        pass

    if TEMPLATE_WORKSHEET_GID:
        try:
            gid = int(TEMPLATE_WORKSHEET_GID)
            for ws in spreadsheet.worksheets():
                if ws.id == gid:
                    return ws
        except ValueError:
            pass

    logger.warning(
        "Template worksheet %r not found in %r; using first sheet",
        TEMPLATE_WORKSHEET_NAME,
        spreadsheet.title,
    )
    return spreadsheet.get_worksheet(0)


def load_drive_credentials() -> service_account.Credentials:
    """Drive credentials, impersonating a real user when configured."""
    info = _load_sa_info()
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=SHEETS_SCOPES
    )
    if DRIVE_IMPERSONATE_USER:
        logger.info("Using Drive impersonation for %s", DRIVE_IMPERSONATE_USER)
        return creds.with_subject(DRIVE_IMPERSONATE_USER)
    return creds


def get_drive_service(creds: service_account.Credentials):
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _service_account_email() -> str:
    return str(_load_sa_info().get("client_email") or "")


def _find_spreadsheet_by_name(drive, name: str) -> str | None:
    """Locate a spreadsheet by exact name, including Shared Drive files."""
    safe_name = name.replace("\\", "\\\\").replace("'", "\\'")

    def _search(with_parent: bool) -> str | None:
        clauses = [
            f"name = '{safe_name}'",
            f"mimeType = '{SPREADSHEET_MIME}'",
            "trashed = false",
        ]
        if with_parent and DRIVE_PARENT_FOLDER_ID:
            clauses.append(f"'{DRIVE_PARENT_FOLDER_ID}' in parents")
        response = (
            drive.files()
            .list(
                q=" and ".join(clauses),
                fields="files(id,name,parents)",
                pageSize=10,
                # Default corpora=user misses Shared Drive files even with
                # includeItemsFromAllDrives; allDrives is required here.
                corpora="allDrives",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files = response.get("files") or []
        return files[0]["id"] if files else None

    found = _search(with_parent=True)
    if found:
        return found
    if DRIVE_PARENT_FOLDER_ID:
        logger.warning(
            "Monthly file %r not found under folder %s; searching all drives",
            name,
            DRIVE_PARENT_FOLDER_ID,
        )
        found = _search(with_parent=False)
        if found:
            logger.info("Found monthly file %r outside the configured folder (%s)", name, found)
    return found


def _prune_to_master_tab(spreadsheet: gspread.Spreadsheet) -> None:
    """A fresh monthly file should start with only the blank master tab."""
    master = _get_master_template_worksheet(spreadsheet)
    for ws in spreadsheet.worksheets():
        if ws.id != master.id:
            logger.info("Removing copied tab %r from new monthly file", ws.title)
            spreadsheet.del_worksheet(ws)


def get_or_create_monthly_spreadsheet(
    drive, gc: gspread.Client, file_name: str
) -> gspread.Spreadsheet:
    """One spreadsheet file per month, copied from the master template file."""
    if MONTHLY_SPREADSHEET_ID:
        logger.info(
            "Opening monthly file via MONTHLY_SPREADSHEET_ID=%s (requested name %r)",
            MONTHLY_SPREADSHEET_ID,
            file_name,
        )
        return gc.open_by_key(MONTHLY_SPREADSHEET_ID)

    existing_id = _find_spreadsheet_by_name(drive, file_name)
    if existing_id:
        logger.info("Reusing monthly file %r (%s)", file_name, existing_id)
        return gc.open_by_key(existing_id)

    body: dict[str, Any] = {"name": file_name}
    if DRIVE_PARENT_FOLDER_ID:
        body["parents"] = [DRIVE_PARENT_FOLDER_ID]
    try:
        created = (
            drive.files()
            .copy(
                fileId=_require_env("TEMPLATE_SHEET_ID", TEMPLATE_SHEET_ID),
                body=body,
                fields="id,name",
                supportsAllDrives=True,
            )
            .execute()
        )
    except HttpError as exc:
        if "storageQuotaExceeded" in str(exc):
            raise RuntimeError(
                f"Cannot create monthly file {file_name!r}: the service account owns no "
                "Drive storage. Set DRIVE_IMPERSONATE_USER (needs the Drive scope added "
                "to domain-wide delegation), use a Shared Drive folder, or pre-create the "
                "file by copying the template and naming it exactly as above."
            ) from exc
        if getattr(exc, "resp", None) is not None and exc.resp.status in (403, 404):
            raise RuntimeError(
                f"Cannot copy template {TEMPLATE_SHEET_ID!r} to create {file_name!r}: "
                f"{exc}. Share that template spreadsheet with "
                f"{_service_account_email() or 'the runtime service account'} as Editor "
                "(and keep it in a Shared Drive the SA can access)."
            ) from exc
        raise

    file_id = created["id"]
    logger.info("Created monthly file %r (%s)", file_name, file_id)

    # The job writes cell data as the plain service account, so it needs access too
    grantees = list(MONTHLY_FILE_SHARE_EMAILS)
    if DRIVE_IMPERSONATE_USER:
        sa_email = _service_account_email()
        if sa_email:
            grantees.append(sa_email)
    for email in grantees:
        drive.permissions().create(
            fileId=file_id,
            body={"type": "user", "role": "writer", "emailAddress": email},
            sendNotificationEmail=False,
            supportsAllDrives=True,
        ).execute()
        logger.info("Shared monthly file with %s", email)

    spreadsheet = gc.open_by_key(file_id)
    _prune_to_master_tab(spreadsheet)
    return spreadsheet


def _get_or_create_dated_template_tab(
    spreadsheet: gspread.Spreadsheet,
    master_ws: gspread.Worksheet,
    tab_name: str,
) -> gspread.Worksheet:
    """Duplicate the master template into a date-named tab (formatting preserved)."""
    try:
        worksheet = spreadsheet.worksheet(tab_name)
        logger.info("Reusing existing dated template tab %r", tab_name)
        return worksheet
    except gspread.WorksheetNotFound:
        logger.info(
            "Duplicating master template %r into new tab %r",
            master_ws.title,
            tab_name,
        )
        return spreadsheet.duplicate_sheet(
            source_sheet_id=master_ws.id,
            new_sheet_name=tab_name,
            insert_sheet_index=1,
        )


def _read_template_columns(worksheet: gspread.Worksheet) -> list[dict[str, str]]:
    """Columns to fill: every column with a header, excluding the label columns."""
    last_col = _index_to_col(min(worksheet.col_count, 40))
    section_row, header_row = worksheet.batch_get(
        [
            f"A{TEMPLATE_HEADER_ROW - 1}:{last_col}{TEMPLATE_HEADER_ROW - 1}",
            f"A{TEMPLATE_HEADER_ROW}:{last_col}{TEMPLATE_HEADER_ROW}",
        ]
    )
    sections = section_row[0] if section_row else []
    headers = header_row[0] if header_row else []

    skip = {
        _col_to_index(TEMPLATE_COL_SEGMENT),
        _col_to_index(TEMPLATE_COL_DIVISION),
    }
    columns: list[dict[str, str]] = []
    current_section = ""
    for offset, header in enumerate(headers):
        col_index = offset + 1
        section = sections[offset] if offset < len(sections) else ""
        if str(section).strip():
            current_section = _clean_text(section)
        if col_index in skip:
            continue
        header_text = _clean_text(header)
        if not header_text:
            continue
        columns.append(
            {
                "col": _index_to_col(col_index),
                "section": current_section,
                "header": header_text,
            }
        )
    return columns


def _clean_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def _norm_label(value: Any) -> str:
    """Punctuation-free upper-case form, so labels match across the two files."""
    text = _clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return text.strip()


def load_field_mapping(path: str = "") -> dict[str, Any]:
    """
    Read ORDERS BULLETIN_data_mapping.xlsx.

    Columns: Vertical | Vertical Supply | tags_from_the_order_bulletin_pdf | MAPPING.
    Returns the PDF tag -> template field mapping per section, plus the component
    sections behind the FMPL TOTAL / JV TOTAL company rows.
    """
    source = path or MAPPING_FILE
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    sheet = workbook[MAPPING_SHEET] if MAPPING_SHEET else workbook.worksheets[0]

    sections: dict[str, dict[str, Any]] = {}
    company_totals: dict[str, list[str]] = {}
    vertical = ""
    supply = ""

    for row in sheet.iter_rows(min_row=2, max_col=4, values_only=True):
        cells = list(row) + [None] * (4 - len(row))
        vertical = _clean_text(cells[0]) or vertical
        supply = _clean_text(cells[1]) or supply
        pdf_tag = _clean_text(cells[2])
        target = _clean_text(cells[3])

        if _norm_label(supply) in MAPPING_TOTAL_LABELS:
            if target:
                company_totals.setdefault(_norm_label(supply), []).append(target)
            continue
        if not pdf_tag or not target or _norm_label(pdf_tag) == "TOTAL":
            continue

        key = _norm_label(supply)
        section = sections.setdefault(
            key, {"vertical": vertical, "supply": supply, "fields": []}
        )
        section["fields"].append({"pdf_tag": pdf_tag, "target": target})

    workbook.close()
    if not sections:
        raise RuntimeError(f"No field mappings found in {source!r}")

    logger.info(
        "Loaded mapping from %r: %d sections, %d fields, company totals for %s",
        source,
        len(sections),
        sum(len(s["fields"]) for s in sections.values()),
        ", ".join(sorted(company_totals)) or "nothing",
    )
    return {"sections": sections, "company_totals": company_totals}


def _segment_candidates(vertical: str, supply: str) -> list[str]:
    """
    Mapping labels and template labels do not always agree.

    "CORE DOMESTIC" is "CORE DOMESTIC FMPL" in the template, and a bare "JV"
    only makes sense with its vertical attached ("CORE DOMESTIC JV").
    """
    return [
        _norm_label(supply),
        _norm_label(f"{vertical} {supply}"),
        _norm_label(f"{supply} FMPL"),
        _norm_label(f"{vertical} {supply} FMPL"),
    ]


def resolve_pdf_tag_rows(
    template_rows: list[dict[str, Any]], mapping: dict[str, Any]
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """
    Turn the mapping into one work item per (template row, PDF tag).

    A row fed by two PDF tags yields two items, which is what makes the INTOPS
    MECH STDS sum possible. Rows the mapping says nothing about (section totals
    and similar) keep their label-matching item with no tag.
    """
    by_segment: dict[str, list[dict[str, Any]]] = {}
    for row in template_rows:
        by_segment.setdefault(_norm_label(row["segment"]), []).append(row)

    tagged: dict[int, list[dict[str, str]]] = {}
    for key, section in mapping["sections"].items():
        segment_rows: list[dict[str, Any]] = []
        for candidate in _segment_candidates(section["vertical"], section["supply"]):
            if candidate in by_segment:
                segment_rows = by_segment[candidate]
                break
        if not segment_rows:
            logger.warning(
                "Mapping section %r (%s) matches no template segment",
                section["supply"],
                section["vertical"],
            )
            continue

        by_division = {_norm_label(r["division"]): r for r in segment_rows}
        for field in section["fields"]:
            target = _norm_label(field["target"])
            row = by_division.get(target)
            if row is None and len(segment_rows) == 1 and not segment_rows[0]["division"]:
                # Sections like "OPC - International" are a single unnamed row
                row = segment_rows[0]
            if row is None:
                logger.warning(
                    "Mapping %r -> %r has no row in template segment %r",
                    field["pdf_tag"],
                    field["target"],
                    section["supply"],
                )
                continue
            tagged.setdefault(row["row"], []).append(
                {"pdf_section": section["supply"], "pdf_tag": field["pdf_tag"]}
            )

    company_rows: dict[str, int] = {}
    for label in mapping["company_totals"]:
        for row in template_rows:
            if _norm_label(row["segment"]) == label:
                company_rows[label] = row["row"]
                break
        if label not in company_rows:
            logger.warning("Company total %r matches no template row", label)

    items: list[dict[str, Any]] = []
    for row in template_rows:
        if row["row"] in company_rows.values():
            continue  # filled by the FMPL/JV TOTAL aggregation instead
        for tag in tagged.get(row["row"], [{}]):
            items.append({**row, **tag})

    logger.info(
        "Mapped %d template rows to PDF tags (%d work items, %d company total rows)",
        len(tagged),
        len(items),
        len(company_rows),
    )
    return items, company_rows


def _update_month_headers(
    worksheet: gspread.Worksheet,
    month_header: str,
    month_abbr: str,
    now: datetime | None = None,
) -> None:
    """Update every date-dependent section / column title from the run date."""
    when = _now_ist(now)
    prev = _previous_month(when)
    prev_header = prev.strftime("%B").upper() + "'" + prev.strftime("%y")
    prev_abbr = prev.strftime("%b") + "'" + prev.strftime("%y")
    # The fiscal year runs Apr-Mar, so these cover April up to last month. In
    # January the previous month belongs to the year before, hence prev's own year.
    prev_span = f"Apr - {prev.strftime('%b')}'{prev.strftime('%y')}"
    growth_span = f"{when.year - 1} - {when.year}"
    # The closed Apr-Mar year that the C-G block reports on, plus the year
    # before it that its growth column compares against.
    closed_end = _closed_fiscal_year_end(when)
    closed_yy = f"{closed_end % 100:02d}"
    closed_span = f"{closed_end - 1} - {closed_end}"
    closed_growth_span = f"{closed_end - 2} - {closed_end - 1}"
    # Open fiscal year label (Apr-Mar), e.g. Aug 2026 -> 2026 - 2027
    open_fy_start = when.year if when.month >= 4 else when.year - 1
    open_fy_span = f"{open_fy_start} - {open_fy_start + 1}"

    header_updates = [
        {"range": TEMPLATE_PREV_FY_LABEL_CELL, "values": [[closed_span]]},
        {
            "range": f"{TEMPLATE_COL_PREV_FY_TOTAL_ACHMNT}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Total Achmnt \n Apr - Mar'{closed_yy}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_PREV_FY_DIGI_SUSTENANCE}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt\n  Apr-Mar'{closed_yy}\n  for Digital Sustenance"]],
        },
        {
            "range": f"{TEMPLATE_COL_PREV_FY_GROWTH}{TEMPLATE_HEADER_ROW}",
            "values": [[f"% Growth over \n {closed_growth_span}\n (AVG)"]],
        },
        {"range": TEMPLATE_CUR_FY_LABEL_CELL, "values": [[open_fy_span]]},
        {"range": TEMPLATE_YTD_FY_LABEL_CELL, "values": [[open_fy_span]]},
        # Previous calendar month block (template's MARCH'26 section)
        {"range": TEMPLATE_PREV_MONTH_HEADER_CELL, "values": [[prev_header]]},
        {
            "range": f"{TEMPLATE_COL_PREV_MONTH_PROJ_INCL}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Projections\n{prev_abbr}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_PREV_MONTH_PROJ_DIGI}{TEMPLATE_HEADER_ROW}",
            "values": [["Projections\n for Digital Sustenance"]],
        },
        {
            "range": f"{TEMPLATE_COL_PREV_MONTH_ACH_INCL}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt \n {prev_abbr}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_PREV_MONTH_ACH_DIGI}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt \n {prev_abbr}\n  for Digital Sustenance"]],
        },
        # Current calendar month block
        {"range": TEMPLATE_MONTH_HEADER_CELL, "values": [[month_header]]},
        {
            "range": f"{TEMPLATE_COL_PROJ_INCL}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Projections\n{month_abbr}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_PROJ_DIGI}{TEMPLATE_HEADER_ROW}",
            "values": [["Projections\n for Digital Sustenance"]],
        },
        {
            "range": f"{TEMPLATE_COL_ACH_INCL}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt \n {month_abbr}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_ACH_DIGI}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt \n {month_abbr}\n  for Digital Sustenance"]],
        },
        {
            "range": f"{TEMPLATE_COL_FY_TOTAL_ACHMNT}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Total Achmnt \n {prev_span}\n (incl. Digital Business)"]],
        },
        {
            "range": f"{TEMPLATE_COL_FY_DIGI_SUSTENANCE}{TEMPLATE_HEADER_ROW}",
            "values": [[f"Achmnt\n {prev_span}\n  for Digital Sustenance"]],
        },
        {
            "range": f"{TEMPLATE_COL_FY_GROWTH}{TEMPLATE_HEADER_ROW}",
            "values": [[f"% Growth over \n {growth_span}\n (AVG)"]],
        },
    ]
    worksheet.batch_update(header_updates, value_input_option="RAW")
    logger.info(
        "Updated headers: prev-month %s=%r / %r; current %s=%r / %r; "
        "%s/%s = %r; open FY %r; closed FY %r",
        TEMPLATE_PREV_MONTH_HEADER_CELL,
        prev_header,
        prev_abbr,
        TEMPLATE_MONTH_HEADER_CELL,
        month_header,
        month_abbr,
        TEMPLATE_COL_FY_TOTAL_ACHMNT,
        TEMPLATE_COL_FY_DIGI_SUSTENANCE,
        prev_span,
        open_fy_span,
        closed_span,
    )


def _read_template_rows(worksheet: gspread.Worksheet) -> list[dict[str, Any]]:
    """Rows to fill, with their segment (col A) and division (col B) labels."""
    last_row = min(max(worksheet.row_count, TEMPLATE_DATA_START_ROW), 120)
    segments, divisions = worksheet.batch_get(
        [
            f"{TEMPLATE_COL_SEGMENT}{TEMPLATE_DATA_START_ROW}:{TEMPLATE_COL_SEGMENT}{last_row}",
            f"{TEMPLATE_COL_DIVISION}{TEMPLATE_DATA_START_ROW}:{TEMPLATE_COL_DIVISION}{last_row}",
        ]
    )

    rows: list[dict[str, Any]] = []
    current_segment = ""
    for offset in range(last_row - TEMPLATE_DATA_START_ROW + 1):
        seg = _clean_text(
            segments[offset][0] if offset < len(segments) and segments[offset] else ""
        )
        div = _clean_text(
            divisions[offset][0] if offset < len(divisions) and divisions[offset] else ""
        )
        if seg:
            current_segment = seg
        if not seg and not div:
            continue
        rows.append(
            {
                "row": TEMPLATE_DATA_START_ROW + offset,
                "segment": seg or current_segment,
                "division": div,
            }
        )
    return rows


def _column_runs(cols: list[str]) -> list[tuple[str, str, list[str]]]:
    """Group column letters into contiguous runs to minimise API ranges."""
    indexed = sorted({_col_to_index(c) for c in cols})
    runs: list[tuple[str, str, list[str]]] = []
    start = prev = None
    current: list[str] = []
    for idx in indexed:
        if start is None:
            start = prev = idx
            current = [_index_to_col(idx)]
            continue
        if idx == prev + 1:
            current.append(_index_to_col(idx))
            prev = idx
            continue
        runs.append((_index_to_col(start), _index_to_col(prev), current))
        start = prev = idx
        current = [_index_to_col(idx)]
    if start is not None:
        runs.append((_index_to_col(start), _index_to_col(prev), current))
    return runs


def _write_metric_values(
    worksheet: gspread.Worksheet,
    metrics: dict[str, Any],
    template_columns: list[dict[str, str]],
) -> tuple[int, int]:
    """Write all extracted values, one batched range per contiguous column run."""
    all_cols = [c["col"] for c in template_columns]
    runs = _column_runs(all_cols)

    updates: list[dict[str, Any]] = []
    populated = 0
    for row in metrics.get("rows") or []:
        row_num = row["row"]
        values = row["values"]
        for first, last, cols in runs:
            block = [values.get(col, "") for col in cols]
            if not any(v != "" for v in block):
                continue
            updates.append(
                {
                    "range": f"{first}{row_num}:{last}{row_num}",
                    "values": [block],
                }
            )
            populated += sum(1 for v in block if v != "")

    if updates:
        worksheet.batch_update(updates, value_input_option="RAW")
    return len(updates), populated


def update_formatted_template(
    pdf_bytes: bytes, when: datetime | None = None
) -> dict[str, Any] | None:
    """Fill a date-named tab inside this month's own spreadsheet file."""
    if not TEMPLATE_SHEET_ID:
        logger.warning("TEMPLATE_SHEET_ID not set; skipping formatted template update")
        return None

    sheets_creds = load_sheets_credentials()
    if not sheets_creds.valid:
        sheets_creds.refresh(Request())
    gc = gspread.authorize(sheets_creds)
    drive = get_drive_service(load_drive_credentials())

    when = _now_ist(when or run_moment())
    monthly_file_name = current_month_tab_name(when)
    spreadsheet = get_or_create_monthly_spreadsheet(drive, gc, monthly_file_name)

    master_ws = _get_master_template_worksheet(spreadsheet)
    template_columns = _read_template_columns(master_ws)
    template_rows = _read_template_rows(master_ws)
    if not template_columns or not template_rows:
        raise RuntimeError(
            "Could not read template structure "
            f"(cols={len(template_columns)}, rows={len(template_rows)})"
        )

    column_letters = [c["col"] for c in template_columns]
    mapping = load_field_mapping()
    work_items, company_rows = resolve_pdf_tag_rows(template_rows, mapping)

    metrics = extract_metrics_with_gemini(pdf_bytes, template_columns, work_items)
    metrics["rows"] = combine_row_items(metrics["rows"], column_letters)

    section_totals = extract_section_totals_with_gemini(
        pdf_bytes, template_columns, describe_total_sections(mapping)
    )
    if company_rows:
        metrics["rows"].extend(
            company_total_rows(mapping, company_rows, section_totals, column_letters)
        )
    metrics["rows"] = apply_derived_template_rows(
        template_rows, metrics["rows"], section_totals, column_letters
    )

    tab_name = day_tab_name(when.date())
    worksheet = _get_or_create_dated_template_tab(spreadsheet, master_ws, tab_name)

    # The master tab stays an empty formatted shell; data lives only in dated tabs
    data_ranges = [
        f"{first}{TEMPLATE_DATA_START_ROW}:{last}{min(master_ws.row_count, 120)}"
        for first, last, _ in _column_runs([c["col"] for c in template_columns])
    ]
    master_ws.batch_clear(data_ranges)

    month_header, month_abbr = _month_labels_from_ist(when)
    pdf_month = str(metrics.get("month_header") or "").strip()
    if pdf_month and pdf_month != month_header:
        logger.info(
            "PDF month_header=%r differs from run-date %r; using run-date for column titles",
            pdf_month,
            month_header,
        )
    _update_month_headers(worksheet, month_header, month_abbr, now=when)

    # Fresh snapshot each run: clear only mapped columns, never the merged spacers
    clear_end = min(worksheet.row_count, 120)
    worksheet.batch_clear(
        [
            f"{first}{TEMPLATE_DATA_START_ROW}:{last}{clear_end}"
            for first, last, _ in _column_runs([c["col"] for c in template_columns])
        ]
    )

    ranges, populated = _write_metric_values(worksheet, metrics, template_columns)
    logger.info(
        "Monthly file %r tab %r filled: %d cells across %d ranges (%d cols x %d rows)",
        monthly_file_name,
        tab_name,
        populated,
        ranges,
        len(template_columns),
        len(template_rows),
    )
    return {
        "file_name": monthly_file_name,
        "file_id": spreadsheet.id,
        "tab_name": tab_name,
        "tab_gid": worksheet.id,
        "month_header": month_header,
        "populated": populated,
    }


def _notification_message(result: dict[str, Any], recipients: list[str]) -> EmailMessage:
    file_url = (
        f"https://docs.google.com/spreadsheets/d/{result['file_id']}"
        f"/edit#gid={result['tab_gid']}"
    )
    subject = f"FM Orders Bulletin - {result['tab_name']} ({result['file_name']})"
    body = (
        f"The Orders Bulletin for {result['tab_name']} has been extracted and written "
        f"to the {result['month_header']} sheet.\n\n"
        f"File: {result['file_name']}\n"
        f"Tab: {result['tab_name']}\n"
        f"Values filled: {result['populated']}\n\n"
        f"Open it here: {file_url}\n\n"
        "-- Automated message from the FM Orders Bulletin job."
    )

    message = EmailMessage()
    message["To"] = ", ".join(recipients)
    message["From"] = NOTIFY_FROM_EMAIL or GMAIL_USER_EMAIL
    message["Subject"] = subject
    message.set_content(body)
    message.add_alternative(
        f"""<html><body>
<p>The Orders Bulletin for <b>{result['tab_name']}</b> has been extracted and written to
the <b>{result['month_header']}</b> sheet.</p>
<ul>
  <li>File: <b>{result['file_name']}</b></li>
  <li>Tab: <b>{result['tab_name']}</b></li>
  <li>Values filled: {result['populated']}</li>
</ul>
<p><a href="{file_url}">Open the bulletin</a></p>
<p style="color:#888;font-size:12px">Automated message from the FM Orders Bulletin job.</p>
</body></html>""",
        subtype="html",
    )
    return message


def send_bulletin_link_email(result: dict[str, Any]) -> None:
    """Mail the link to the freshly written monthly file."""
    if not NOTIFY_EMAILS:
        logger.info("NOTIFY_EMAILS is empty; skipping notification mail")
        return

    message = _notification_message(result, NOTIFY_EMAILS)
    gmail = get_gmail_service(load_gmail_send_credentials())
    gmail.users().messages().send(
        userId="me",
        body={"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()},
    ).execute()
    logger.info(
        "Sent bulletin link for tab %r to %s",
        result["tab_name"],
        ", ".join(NOTIFY_EMAILS),
    )


def repair_month_headers(when: datetime | None = None) -> dict[str, Any]:
    """
    Rewrite dynamic month/year headers on existing dated tabs only.

    No Gmail or Gemini calls. With RUN_DATE, repairs that one tab; otherwise
    repairs every DD/MM/YYYY tab inside the current month's spreadsheet.
    """
    if not TEMPLATE_SHEET_ID:
        raise RuntimeError("TEMPLATE_SHEET_ID is required for header repair")

    when = _now_ist(when or run_moment())
    sheets_creds = load_sheets_credentials()
    if not sheets_creds.valid:
        sheets_creds.refresh(Request())
    gc = gspread.authorize(sheets_creds)
    drive = get_drive_service(load_drive_credentials())

    monthly_file_name = current_month_tab_name(when)
    spreadsheet = get_or_create_monthly_spreadsheet(drive, gc, monthly_file_name)

    if RUN_DATE:
        targets = [(day_tab_name(when.date()), when)]
    else:
        targets = []
        for ws in spreadsheet.worksheets():
            try:
                tab_day = datetime.strptime(ws.title, "%d/%m/%Y").replace(tzinfo=IST)
            except ValueError:
                continue
            targets.append((ws.title, tab_day))

    if not targets:
        raise RuntimeError(
            f"No dated tabs found to repair in {monthly_file_name!r}"
        )

    repaired = 0
    for tab_name, tab_when in targets:
        try:
            worksheet = spreadsheet.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            logger.warning("Tab %r not found; skipping", tab_name)
            continue
        month_header, month_abbr = _month_labels_from_ist(tab_when)
        _update_month_headers(worksheet, month_header, month_abbr, now=tab_when)
        repaired += 1
        logger.info("Repaired headers on tab %r for %s", tab_name, tab_when.date())

    return {
        "file_name": monthly_file_name,
        "file_id": spreadsheet.id,
        "repaired_tabs": repaired,
    }


def repair_month_headers(when: datetime | None = None) -> dict[str, Any]:
    """
    Rewrite dynamic month/year headers on existing dated tabs only.

    No Gmail or Gemini calls. With RUN_DATE, repairs that one tab; otherwise
    repairs every DD/MM/YYYY tab inside the current month's spreadsheet.
    """
    if not TEMPLATE_SHEET_ID:
        raise RuntimeError("TEMPLATE_SHEET_ID is required for header repair")

    when = _now_ist(when or run_moment())
    sheets_creds = load_sheets_credentials()
    if not sheets_creds.valid:
        sheets_creds.refresh(Request())
    gc = gspread.authorize(sheets_creds)
    drive = get_drive_service(load_drive_credentials())

    monthly_file_name = current_month_tab_name(when)
    spreadsheet = get_or_create_monthly_spreadsheet(drive, gc, monthly_file_name)

    if RUN_DATE:
        targets = [(day_tab_name(when.date()), when)]
    else:
        targets = []
        for ws in spreadsheet.worksheets():
            try:
                tab_day = datetime.strptime(ws.title, "%d/%m/%Y").replace(tzinfo=IST)
            except ValueError:
                continue
            targets.append((ws.title, tab_day))

    if not targets:
        raise RuntimeError(
            f"No dated tabs found to repair in {monthly_file_name!r}"
        )

    repaired = 0
    for tab_name, tab_when in targets:
        try:
            worksheet = spreadsheet.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            logger.warning("Tab %r not found; skipping", tab_name)
            continue
        month_header, month_abbr = _month_labels_from_ist(tab_when)
        _update_month_headers(worksheet, month_header, month_abbr, now=tab_when)
        repaired += 1
        logger.info("Repaired headers on tab %r for %s", tab_name, tab_when.date())

    return {
        "file_name": monthly_file_name,
        "file_id": spreadsheet.id,
        "repaired_tabs": repaired,
    }


def run() -> None:
    when = run_moment()
    if HEADER_REPAIR_ONLY:
        logger.info(
            "Header-repair mode (effective date=%s%s)",
            when.date().isoformat(),
            f", RUN_DATE={RUN_DATE}" if RUN_DATE else ", all dated tabs",
        )
        result = repair_month_headers(when)
        logger.info(
            "Header repair done on %r (%d tab(s))",
            result["file_name"],
            result["repaired_tabs"],
        )
        logger.info("Job completed successfully")
        return

    logger.info(
        "Starting FM Orders Bulletin job (effective date=%s%s)",
        when.date().isoformat(),
        f", RUN_DATE={RUN_DATE}" if RUN_DATE else "",
    )
    gmail_creds = load_gmail_credentials()
    gmail = get_gmail_service(gmail_creds)
    message_id = find_latest_message_id(
        gmail, EMAIL_SUBJECT, on_day=when.date() if RUN_DATE else None
    )
    pdf_bytes = download_pdf_attachment(gmail, message_id)

    # Production path: dated tabs inside the monthly Shared Drive file.
    # The legacy raw dump to GOOGLE_SHEET_ID is only used when no template is
    # configured - and must never abort the formatted write (a 403 on that
    # old sheet was failing the whole daily job after Gemini had already run).
    if TEMPLATE_SHEET_ID:
        result = update_formatted_template(pdf_bytes, when=when)
        if result and not RUN_DATE:
            try:
                send_bulletin_link_email(result)
            except Exception:
                logger.exception("Could not send the bulletin link mail")
        elif result and RUN_DATE:
            logger.info(
                "Skipping notification mail for backfill RUN_DATE=%s (tab %r)",
                RUN_DATE,
                result.get("tab_name"),
            )
    else:
        rows = extract_table_with_gemini(pdf_bytes)
        write_to_sheet(rows)

    logger.info("Job completed successfully")


if __name__ == "__main__":
    try:
        run()
    except Exception:
        logger.exception("Job failed")
        sys.exit(1)
