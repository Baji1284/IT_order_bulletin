#!/usr/bin/env python3
"""
FM Orders Bulletin — daily PDF email → Gemini extraction → Google Sheet.

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
from typing import Any

import gspread
from google import genai
from google.auth.transport.requests import Request
from google.genai import types
from google.oauth2 import service_account
from googleapiclient.discovery import build

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
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
SHEET_RANGE = os.environ.get("SHEET_RANGE", "Sheet1")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_APPLICATION_CREDENTIALS = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

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


def load_sheets_credentials() -> service_account.Credentials:
    """SA credentials for Sheets (share the sheet with the SA client_email)."""
    info = _load_sa_info()
    return service_account.Credentials.from_service_account_info(
        info, scopes=SHEETS_SCOPES
    )


def get_gmail_service(creds: service_account.Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def find_latest_message_id(gmail, subject: str) -> str:
    query = f'subject:"{subject}" has:attachment filename:pdf newer_than:2d'
    logger.info("Searching Gmail: %s", query)
    result = (
        gmail.users()
        .messages()
        .list(userId="me", q=query, maxResults=5)
        .execute()
    )
    messages = result.get("messages") or []
    if not messages:
        query = f'subject:"{subject}" newer_than:2d'
        result = (
            gmail.users()
            .messages()
            .list(userId="me", q=query, maxResults=5)
            .execute()
        )
        messages = result.get("messages") or []

    if not messages:
        raise RuntimeError(f'No recent email found with subject "{subject}"')

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


def extract_table_with_gemini(pdf_bytes: bytes) -> list[list[Any]]:
    _require_env("GEMINI_API_KEY", GEMINI_API_KEY)
    client = genai.Client(api_key=GEMINI_API_KEY)

    logger.info(
        "Sending PDF (%d bytes) to Gemini model=%s", len(pdf_bytes), GEMINI_MODEL
    )
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            types.Content(
                role="user",
                parts=[
                    types.Part.from_bytes(
                        data=pdf_bytes, mime_type="application/pdf"
                    ),
                    types.Part.from_text(
                        text="Extract the Orders Bulletin table from this PDF."
                    ),
                ],
            )
        ],
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            temperature=0.0,
            response_mime_type="application/json",
        ),
    )

    text = (response.text or "").strip()
    if not text:
        raise RuntimeError("Gemini returned an empty response")

    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        logger.error("Gemini raw response (truncated): %s", text[:2000])
        raise RuntimeError(f"Gemini did not return valid JSON: {exc}") from exc

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


def write_to_sheet(rows: list[list[Any]]) -> None:
    sheet_id = _require_env("GOOGLE_SHEET_ID", GOOGLE_SHEET_ID)
    sheets_creds = load_sheets_credentials()
    if not sheets_creds.valid:
        sheets_creds.refresh(Request())

    gc = gspread.authorize(sheets_creds)
    spreadsheet = gc.open_by_key(sheet_id)

    worksheet_name = SHEET_RANGE.split("!")[0].strip() or "Sheet1"
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.get_worksheet(0)
        logger.warning(
            "Worksheet %r not found; using first sheet %r",
            worksheet_name,
            worksheet.title,
        )

    # Strings preserve exact numeric text from the PDF / Gemini output
    values = [["" if c is None else str(c) for c in row] for row in rows]

    logger.info(
        "Clearing worksheet %r and writing %d rows", worksheet.title, len(values)
    )
    worksheet.clear()
    worksheet.update(values, value_input_option="RAW")
    logger.info("Google Sheet updated successfully")


def run() -> None:
    logger.info("Starting FM Orders Bulletin job")
    gmail_creds = load_gmail_credentials()
    gmail = get_gmail_service(gmail_creds)
    message_id = find_latest_message_id(gmail, EMAIL_SUBJECT)
    pdf_bytes = download_pdf_attachment(gmail, message_id)
    rows = extract_table_with_gemini(pdf_bytes)
    write_to_sheet(rows)
    logger.info("Job completed successfully")


if __name__ == "__main__":
    try:
        run()
    except Exception:
        logger.exception("Job failed")
        sys.exit(1)
