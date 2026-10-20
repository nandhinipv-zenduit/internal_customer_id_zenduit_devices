"""
Append new Zenduit devices from Zoho Analytics to the billing-mapping Google Sheet.

APPEND-ONLY: rows already in the sheet are never read for update, changed or deleted.
Once a device is in the sheet, only the Apps Script edit screen changes it.

Flow (one run):
  1. Export Device_Id, serial_number, AccountId, Company_Name from the Zoho Analytics
     "Zenduit Devices" table (CSV).
  2. Read the Device ID column of the sheet's "Mapping" tab.
  3. Append every Device ID that is not in the sheet yet (values stored as plain text).
  4. Write one line to the "Sync Log" tab and one line per new device to "New Devices Log".

Configuration comes from environment variables (GitHub Actions secrets / variables):
  required secrets : ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, ZOHO_REFRESH_TOKEN, ZOHO_ORG_ID,
                     GOOGLE_SERVICE_ACCOUNT_JSON
  optional         : see DEFAULTS below; DRY_RUN=1 to report without writing;
                     ALLOW_LARGE_APPEND=1 to bypass the safety limit once.
"""
from __future__ import annotations

import csv
import io
import json
import logging
import os
import sys
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo

DEFAULTS = {
    "ZOHO_ACCOUNTS_HOST": "https://accounts.zoho.com",
    "ZOHO_API_HOST": "https://analyticsapi.zoho.com",
    "ZOHO_WORKSPACE_ID": "953790000013364003",
    "ZOHO_SOURCE_VIEW_ID": "953790000054827175",          # Zenduit Devices
    "SHEET_ID": "1knoy-tqSgtHl_v0uWIkphCObSspyRcD2B0nIDur79eo",  # Zenduit Billing Mapping
    "MAPPING_TAB": "Mapping",
    "SYNC_LOG_TAB": "Sync Log",
    "NEW_DEVICES_LOG_TAB": "New Devices Log",
    "TIMEZONE": "Asia/Kolkata",
    # Safety: refuse to append more than max(MIN, FRACTION * existing rows) in one run.
    "MAX_NEW_ROWS_MIN": "1000",
    "MAX_NEW_ROWS_FRACTION": "0.10",
}

# Zoho column -> meaning
SOURCE_COLS = {
    "device_id": "Device_Id",
    "serial": "serial_number",
    "account_id": "AccountId",
    "account_name": "Company_Name",
}
# Sheet header (row 1 of the Mapping tab), in order
SHEET_HEADERS = ["Device ID", "Serial", "Account ID", "Customer Name"]

SYNC_LOG_HEADERS = ["Started", "Finished", "Status", "New devices added", "Message"]
NEW_DEVICES_LOG_HEADERS = ["Added at", "Device ID", "Serial", "Account ID", "Customer Name", "Source"]

APPEND_CHUNK = 5000
log = logging.getLogger("billing-sync")


class SyncError(Exception):
    """Expected, explainable failure – message goes to the Sync Log."""


def cfg(name: str) -> str:
    value = os.environ.get(name, DEFAULTS.get(name))
    if value is None or value == "":
        raise SyncError(f"Missing configuration: {name}")
    return value


def now_str() -> str:
    return datetime.now(ZoneInfo(cfg("TIMEZONE"))).strftime("%Y-%m-%d %H:%M:%S")


# ----------------------------------------------------------------------------- Zoho

@dataclass
class Device:
    device_id: str
    serial: str
    account_id: str
    account_name: str

    def as_sheet_row(self) -> list[str]:
        return [self.device_id, self.serial, self.account_id, self.account_name]


def zoho_access_token(http) -> str:
    resp = http.post(
        f"{cfg('ZOHO_ACCOUNTS_HOST')}/oauth/v2/token",
        params={
            "refresh_token": cfg("ZOHO_REFRESH_TOKEN"),
            "client_id": cfg("ZOHO_CLIENT_ID"),
            "client_secret": cfg("ZOHO_CLIENT_SECRET"),
            "grant_type": "refresh_token",
        },
        timeout=60,
    )
    body = resp.json() if resp.content else {}
    token = body.get("access_token")
    if not token:
        raise SyncError(f"Could not get Zoho access token (HTTP {resp.status_code}): {str(body)[:300]}")
    return token


def export_devices(http) -> list[Device]:
    token = zoho_access_token(http)
    config = {"responseFormat": "csv", "selectedColumns": list(SOURCE_COLS.values())}
    url = (f"{cfg('ZOHO_API_HOST')}/restapi/v2/workspaces/{cfg('ZOHO_WORKSPACE_ID')}"
           f"/views/{cfg('ZOHO_SOURCE_VIEW_ID')}/data")
    resp = http.get(
        url,
        params={"CONFIG": json.dumps(config)},
        headers={"Authorization": f"Zoho-oauthtoken {token}", "ZANALYTICS-ORGID": cfg("ZOHO_ORG_ID")},
        timeout=600,
    )
    if resp.status_code != 200:
        raise SyncError(f"Zoho export failed (HTTP {resp.status_code}): {resp.text[:500]}")
    return parse_export(resp.content.decode("utf-8-sig"))


def parse_export(text: str) -> list[Device]:
    reader = csv.DictReader(io.StringIO(text))
    header = [h.strip() for h in (reader.fieldnames or [])]
    missing = [c for c in SOURCE_COLS.values() if c not in header]
    if missing:
        raise SyncError(f"Zoho export is missing column(s) {missing}. Got: {header}")
    reader.fieldnames = header

    devices: list[Device] = []
    seen: set[str] = set()
    for row in reader:
        dev = Device(**{k: (row.get(col) or "").strip() for k, col in SOURCE_COLS.items()})
        if not dev.device_id or dev.device_id in seen:
            continue
        seen.add(dev.device_id)
        devices.append(dev)
    if not devices:
        raise SyncError("Zoho export returned no devices – nothing compared, nothing appended.")
    return devices


# ---------------------------------------------------------------------------- Sheet

def open_spreadsheet():
    import gspread  # imported here so tests don't need it
    from google.oauth2.service_account import Credentials

    info = json.loads(cfg("GOOGLE_SERVICE_ACCOUNT_JSON"))
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds).open_by_key(cfg("SHEET_ID"))


def existing_device_ids(ws) -> set[str]:
    header = [h.strip() for h in ws.row_values(1)]
    if header[: len(SHEET_HEADERS)] != SHEET_HEADERS:
        raise SyncError(f"Mapping tab header changed. Expected {SHEET_HEADERS}, found {header}. "
                        "Nothing appended.")
    ids = [v.strip() for v in ws.col_values(1)[1:]]
    return {v for v in ids if v}


def get_or_create_tab(spreadsheet, title: str, headers: list[str]):
    try:
        return spreadsheet.worksheet(title)
    except Exception:  # gspread.WorksheetNotFound
        ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=len(headers))
        ws.append_row(headers, value_input_option="RAW")
        return ws


def append_rows(ws, rows: list[list[str]]) -> None:
    for i in range(0, len(rows), APPEND_CHUNK):
        ws.append_rows(rows[i:i + APPEND_CHUNK], value_input_option="RAW",
                       insert_data_option="INSERT_ROWS", table_range="A1")


# ----------------------------------------------------------------------------- Run

def find_new(devices: list[Device], existing: set[str]) -> list[Device]:
    return [d for d in devices if d.device_id not in existing]


def check_limit(new_count: int, existing_count: int) -> None:
    if os.environ.get("ALLOW_LARGE_APPEND") == "1":
        return
    limit = max(int(cfg("MAX_NEW_ROWS_MIN")), int(float(cfg("MAX_NEW_ROWS_FRACTION")) * existing_count))
    if new_count > limit:
        raise SyncError(
            f"{new_count} new devices found, above the safety limit of {limit}. "
            "This usually means the sheet or the export was misread. Nothing appended. "
            "If it is genuine, re-run once with ALLOW_LARGE_APPEND=1.")


def run(http, spreadsheet, dry_run: bool = False) -> int:
    """Returns the number of devices appended. Raises SyncError on any problem."""
    mapping = spreadsheet.worksheet(cfg("MAPPING_TAB"))
    existing = existing_device_ids(mapping)
    if not existing:
        raise SyncError("Mapping tab has no Device IDs – refusing to run (sheet misread?).")

    devices = export_devices(http)
    new = find_new(devices, existing)
    log.info("Zoho devices: %d | in sheet: %d | new: %d", len(devices), len(existing), len(new))
    check_limit(len(new), len(existing))

    if not new or dry_run:
        return len(new)          # dry run: number that WOULD be appended

    append_rows(mapping, [d.as_sheet_row() for d in new])
    stamp = now_str()
    append_rows(get_or_create_tab(spreadsheet, cfg("NEW_DEVICES_LOG_TAB"), NEW_DEVICES_LOG_HEADERS),
                [[stamp, *d.as_sheet_row(), "Zoho Analytics – Zenduit Devices"] for d in new])
    return len(new)


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    import requests

    dry_run = os.environ.get("DRY_RUN") == "1"
    started = now_str()
    spreadsheet = None
    try:
        spreadsheet = open_spreadsheet()
        with requests.Session() as http:
            added = run(http, spreadsheet, dry_run=dry_run)
        status, code = ("DRY RUN" if dry_run else "OK"), 0
        message = f"Would append {added} device(s); nothing written." if dry_run else ""
        log.info("Done. %s %d device(s).", "Would append" if dry_run else "Appended", added)
    except SyncError as e:
        added, status, message, code = 0, "FAILED", str(e), 1
        log.error(message)
    except Exception as e:  # unexpected – still log it, then fail the job
        added, status, message, code = 0, "FAILED", f"{type(e).__name__}: {e}", 1
        log.exception("Unexpected error")

    if spreadsheet is not None and not dry_run:
        try:
            append_rows(get_or_create_tab(spreadsheet, cfg("SYNC_LOG_TAB"), SYNC_LOG_HEADERS),
                        [[started, now_str(), status, str(added), message[:500]]])
        except Exception:
            log.exception("Could not write to Sync Log")
    return code


if __name__ == "__main__":
    sys.exit(main())
