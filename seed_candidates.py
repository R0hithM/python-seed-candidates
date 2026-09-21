#!/usr/bin/env python3
"""Seed ATS candidates from ProfilesData.xlsx via the Google Drive resume API."""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_API_URL = (
    "https://ats-api-agbpc2drfqccg4gg.southindia-01.azurewebsites.net/api/google-drive"
)
REQUEST_TIMEOUT_SECONDS = 90
ERROR_BODY_MAX_LEN = 500

RESUME_LINK_HEADER = "Resume Link"
CANDIDATE_NAME_HEADER = "Candidate Name"

COL_DOCUMENT_ID = "documentId"
COL_CANDIDATE_ID = "candidateId"
COL_IMPORT_ID = "importId"
COL_SEED_STATUS = "seed_status"
COL_SEED_HTTP_STATUS = "seed_http_status"
COL_SEED_ERROR = "seed_error"

RESULT_COLUMNS = (
    COL_DOCUMENT_ID,
    COL_CANDIDATE_ID,
    COL_IMPORT_ID,
    COL_SEED_STATUS,
    COL_SEED_HTTP_STATUS,
    COL_SEED_ERROR,
)

STATUS_SUCCESS = "SUCCESS"
STATUS_FAILED = "FAILED"
STATUS_SKIPPED = "SKIPPED"
TERMINAL_STATUSES = {STATUS_SUCCESS, STATUS_FAILED, STATUS_SKIPPED}


def load_config() -> tuple[str, str]:
    load_dotenv()
    api_url = os.getenv("ATS_API_URL", DEFAULT_API_URL).strip()
    token = os.getenv("ATS_BEARER_TOKEN", "").strip()
    if not token or token == "your_jwt_here":
        print(
            "ATS_BEARER_TOKEN is missing. Copy .env.example to .env and set the token.",
            file=sys.stderr,
        )
        sys.exit(1)
    return api_url, token


def header_map(ws: Worksheet) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        mapping[str(cell.value).strip()] = cell.column
    return mapping


def ensure_result_columns(ws: Worksheet, headers: dict[str, int]) -> dict[str, int]:
    next_col = ws.max_column
    for name in RESULT_COLUMNS:
        if name in headers:
            continue
        next_col += 1
        ws.cell(row=1, column=next_col, value=name)
        headers[name] = next_col
    return headers


def cell_str(ws: Worksheet, row: int, col: int | None) -> str:
    if col is None:
        return ""
    value = ws.cell(row=row, column=col).value
    if value is None:
        return ""
    return str(value).strip()


def is_google_drive_link(url: str) -> bool:
    if not url:
        return False
    try:
        parsed = urlparse(url)
    except ValueError:
        return False
    if parsed.scheme not in {"http", "https"}:
        return False
    host = (parsed.netloc or "").lower()
    return host == "drive.google.com" or host.endswith(".google.com") and "drive" in host


def write_result(
    ws: Worksheet,
    row: int,
    headers: dict[str, int],
    *,
    document_id: str = "",
    candidate_id: str = "",
    import_id: str = "",
    status: str,
    http_status: int | str = "",
    error: str = "",
) -> None:
    ws.cell(row=row, column=headers[COL_DOCUMENT_ID], value=document_id)
    ws.cell(row=row, column=headers[COL_CANDIDATE_ID], value=candidate_id)
    ws.cell(row=row, column=headers[COL_IMPORT_ID], value=import_id)
    ws.cell(row=row, column=headers[COL_SEED_STATUS], value=status)
    ws.cell(row=row, column=headers[COL_SEED_HTTP_STATUS], value=http_status)
    ws.cell(row=row, column=headers[COL_SEED_ERROR], value=error)


def truncate_error(text: str) -> str:
    text = " ".join(text.split())
    if len(text) <= ERROR_BODY_MAX_LEN:
        return text
    return text[: ERROR_BODY_MAX_LEN - 3] + "..."


def extract_ids(payload: Any) -> tuple[str, str, str] | None:
    if not isinstance(payload, dict):
        return None
    document_id = str(payload.get("documentId") or "").strip()
    candidate_id = str(payload.get("candidateId") or "").strip()
    import_id = str(payload.get("importId") or "").strip()
    if document_id and candidate_id and import_id:
        return document_id, candidate_id, import_id
    return None


def call_api(api_url: str, token: str, drive_link: str) -> tuple[int | str, Any, str | None]:
    """Return (http_status, json_or_none, error_text_or_none)."""
    try:
        response = requests.post(
            api_url,
            headers={
                "accept": "application/json",
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"driveLink": drive_link},
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
    except requests.Timeout:
        return "", None, "timeout after 90s"
    except requests.RequestException as exc:
        return "", None, truncate_error(f"request error: {exc}")

    body_text = response.text or ""
    payload: Any = None
    try:
        payload = response.json()
    except ValueError:
        payload = None

    return response.status_code, payload, body_text


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Seed ATS candidates from an Excel file of Google Drive resume links."
    )
    parser.add_argument(
        "--file",
        default="ProfilesData.xlsx",
        help="Path to the Excel workbook (default: ProfilesData.xlsx)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Process at most N pending rows this run (for smoke tests)",
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="Re-attempt rows previously marked FAILED (SUCCESS and SKIPPED still skipped)",
    )
    return parser.parse_args()


def should_skip_row(status: str, retry_failed: bool) -> bool:
    if not status:
        return False
    if retry_failed and status == STATUS_FAILED:
        return False
    return status in TERMINAL_STATUSES


def main() -> int:
    args = parse_args()
    api_url, token = load_config()

    path = Path(args.file)
    if not path.exists():
        print(f"Excel file not found: {path}", file=sys.stderr)
        return 1

    workbook = load_workbook(path)
    ws = workbook.active
    headers = ensure_result_columns(ws, header_map(ws))

    resume_col = headers.get(RESUME_LINK_HEADER)
    if resume_col is None:
        print(f"Column '{RESUME_LINK_HEADER}' not found in {path}", file=sys.stderr)
        return 1
    name_col = headers.get(CANDIDATE_NAME_HEADER)

    processed = 0
    success = 0
    failed = 0
    skipped = 0
    already_done = 0
    dirty = False

    def save() -> None:
        nonlocal dirty
        if not dirty:
            return
        workbook.save(path)
        dirty = False

    max_row = ws.max_row
    print(f"Loaded {path} sheet={ws.title!r} rows={max_row - 1} api={api_url}")

    try:
        for row in range(2, max_row + 1):
            if args.limit is not None and processed >= args.limit:
                break

            status = cell_str(ws, row, headers[COL_SEED_STATUS])
            name = cell_str(ws, row, name_col) or f"row {row}"

            if should_skip_row(status, args.retry_failed):
                already_done += 1
                continue

            drive_link = cell_str(ws, row, resume_col)
            if not is_google_drive_link(drive_link):
                reason = "missing resume link" if not drive_link else "not a Google Drive URL"
                write_result(
                    ws,
                    row,
                    headers,
                    status=STATUS_SKIPPED,
                    error=reason,
                )
                dirty = True
                save()
                skipped += 1
                processed += 1
                print(f"[{row}/{max_row}] SKIPPED {name}: {reason}")
                continue

            started = time.perf_counter()
            print(f"[{row}/{max_row}] POST {name} ...", flush=True)
            http_status, payload, body_or_error = call_api(api_url, token, drive_link)
            elapsed = time.perf_counter() - started

            ids = extract_ids(payload) if http_status == 200 else None
            if http_status == 201 and ids:
                document_id, candidate_id, import_id = ids
                write_result(
                    ws,
                    row,
                    headers,
                    document_id=document_id,
                    candidate_id=candidate_id,
                    import_id=import_id,
                    status=STATUS_SUCCESS,
                    http_status=http_status,
                )
                dirty = True
                save()
                success += 1
                processed += 1
                print(
                    f"[{row}/{max_row}] SUCCESS {name} "
                    f"candidateId={candidate_id} ({elapsed:.1f}s)"
                )
                continue

            if payload is not None:
                error_text = truncate_error(str(payload))
            else:
                error_text = truncate_error(body_or_error or "empty response")

            if http_status == 201 and ids is None:
                error_text = truncate_error(
                    f"200 response missing documentId/candidateId/importId: {error_text}"
                )

            write_result(
                ws,
                row,
                headers,
                status=STATUS_FAILED,
                http_status=http_status,
                error=error_text,
            )
            dirty = True
            save()
            failed += 1
            processed += 1
            print(
                f"[{row}/{max_row}] FAILED {name} http={http_status or '-'} "
                f"({elapsed:.1f}s) {error_text}"
            )
    except KeyboardInterrupt:
        print("\nInterrupted; saving workbook...", file=sys.stderr)
        save()
        print(
            f"Stopped. processed={processed} success={success} failed={failed} "
            f"skipped={skipped} already_done={already_done}"
        )
        return 130
    finally:
        save()

    print(
        f"Done. processed={processed} success={success} failed={failed} "
        f"skipped={skipped} already_done={already_done}"
    )
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
