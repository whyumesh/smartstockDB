"""
Stockist Mail Dispatcher  v3
==============================
• Sends FROM a shared Outlook mailbox  (stockandsales@abbott.com)
• Uses win32com.client  →  requires Outlook desktop installed & profile open
• ONE email per unique customer email ID
• All TBM and ABM addresses go into BCC only — To: field is intentionally blank
• All attachment files for a customer bundled in that single email
• Number of emails sent == number of unique customer email IDs
• Dual date-verification on attachments  →  filename timestamp AND file
  system modified-time must BOTH fall within the window
• Full dispatch log written to Excel after every run

REQUIREMENTS
    pip install pywin32 pandas openpyxl

USAGE
    1. Fill in the USER CONFIG section below
    2. Run:  python stockist_mail_dispatcher_v3.py
    3. Review dispatch_log.xlsx  (all emails saved to Drafts in dry-run mode)
    4. Set SEND_EMAILS = True  and run again to actually send
"""

import os
import re
import logging
import pandas as pd

from datetime import datetime
from pathlib import Path

# win32com is Windows/Outlook only; import guarded so the file can be
# syntax-checked on any OS without crashing
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ═══════════════════════════════════════════════════════════════════════════════
#  ── USER CONFIG ──  (edit ONLY this section)
# ═══════════════════════════════════════════════════════════════════════════════

# Date-time window (inclusive on both ends)
# Format: "YYYY-MM-DD HH:MM:SS"
START_DATETIME = "2026-03-13 00:00:00"
END_DATETIME   = "2026-03-13 23:59:59"

# File paths
STOCKIST_LOG_PATH  = r"C:\StockistMailer\Stockist_Log.xlsx"
BASE_FILE_PATH     = r"C:\StockistMailer\Umesh_Mail_Reminder_Base_File_March_2026_dummy.xlsx"
ATTACHMENTS_FOLDER = r"C:\StockistMailer\Attachments"
OUTPUT_LOG_PATH    = r"C:\StockistMailer\dispatch_log.xlsx"

# Shared mailbox to send FROM  (must be added to your Outlook profile)
SHARED_MAILBOX  = "stockandsales@abbott.com"
SENDER_DISPLAY  = "Stock & Sales Operations"

# Set False for a safe dry-run (emails saved to Outlook Drafts, NOT sent)
# Set True only after reviewing the dry-run dispatch_log.xlsx
SEND_EMAILS = False

# ── Email template ─────────────────────────────────────────────────────────────
# Placeholders filled automatically — do NOT rename them
# {stockist_code}   → CustomerNo from Stockist Log
# {received_date}   → date extracted from attachment filename (YYYYMMDD portion)
# {customer_name}   → c_cust_name column from Base File
# ──────────────────────────────────────────────────────────────────────────────

EMAIL_SUBJECT = "Stock Update - {stockist_code}"

EMAIL_BODY = """\
Dear Team,

Please find attached the stock document recieved from the stockist.

Customer Code : {stockist_code}
Received Date : {received_date}
Customer Name : {customer_name}

Regards,
SmartStock Team"""

# ═══════════════════════════════════════════════════════════════════════════════
#  ── END OF USER CONFIG ──
# ═══════════════════════════════════════════════════════════════════════════════


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 1 – Load and clean Stockist Log
# ─────────────────────────────────────────────────────────────────────────────

def load_log(path: str) -> pd.DataFrame:
    log.info("Loading Stockist Log: %s", path)
    df = pd.read_excel(path, sheet_name="Sheet1", dtype={"CustomerNo": str})

    # Status column arrives as  "Pass'\n"  or  "Fail'"  – strip all junk
    df["Status"] = (
        df["Status"]
        .astype(str)
        .str.replace(r"['\n\r]", "", regex=True)
        .str.strip()
    )

    # CustomerNo – some failed rows store literal "'" ; keep digits only
    df["CustomerNo"] = (
        df["CustomerNo"]
        .astype(str)
        .str.replace(r"[^0-9]", "", regex=True)
        .str.strip()
    )

    # Timestamp
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")

    log.info("Log loaded: %d total rows", len(df))
    return df


def filter_log(df: pd.DataFrame, start: datetime, end: datetime) -> pd.DataFrame:
    """Return one row per stockist code (latest PASS) within [start, end]."""
    mask = (
        (df["Status"].str.upper() == "PASS")
        & (df["Timestamp"] >= start)
        & (df["Timestamp"] <= end)
        & (df["CustomerNo"].str.match(r"^\d{5,}$", na=False))
    )
    filtered = df[mask].copy()
    log.info("PASS rows in window (before dedup): %d", len(filtered))

    if filtered.empty:
        return filtered

    # Keep only the single LATEST submission per stockist code.
    filtered.sort_values("Timestamp", ascending=False, inplace=True)
    filtered.drop_duplicates(subset=["CustomerNo"], keep="first", inplace=True)
    filtered.reset_index(drop=True, inplace=True)

    log.info("Unique stockists after dedup (one per code): %d", len(filtered))
    return filtered


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2 – Load Base File and build lookup
# ─────────────────────────────────────────────────────────────────────────────

def load_base(path: str) -> pd.DataFrame:
    log.info("Loading Base File: %s", path)
    df = pd.read_excel(path, dtype={"c_cust_code": str})
    df["c_cust_code"] = df["c_cust_code"].astype(str).str.strip()

    for col in ["Email", "ABM Email", "TBM Employ. Name", "ABM_Name",
                "Division Name", "c_cust_name"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
                .replace({"nan": "", "None": "", "NOT ASSIGNED": "", "0": ""})
            )
    log.info("Base File loaded: %d rows", len(df))
    return df


def build_lookup(base_df: pd.DataFrame) -> dict:
    """
    Collapse ALL divisions for a stockist into ONE record.
    Collects unique TBM and ABM email sets so no address is duplicated.

    Returns:
        {
          stockist_code: {
            "tbm_emails"   : set of lowercase email strings,
            "abm_emails"   : set of lowercase email strings,
            "customer_name": first non-empty c_cust_name found,
            "divisions"    : list of division name strings,
          }
        }
    """
    lookup: dict = {}

    for _, row in base_df.iterrows():
        code = str(row.get("c_cust_code", "")).strip()
        if not code or code == "nan":
            continue

        tbm_email     = str(row.get("Email", "")).strip()
        abm_email     = str(row.get("ABM Email", "")).strip()
        customer_name = str(row.get("c_cust_name", "")).strip()
        division      = str(row.get("Division Name", "")).strip()

        if code not in lookup:
            lookup[code] = {
                "tbm_emails":    set(),
                "abm_emails":    set(),
                "customer_name": "",
                "divisions":     [],
            }

        if tbm_email and tbm_email not in ("nan", ""):
            lookup[code]["tbm_emails"].add(tbm_email.lower())

        if abm_email and abm_email not in ("nan", ""):
            lookup[code]["abm_emails"].add(abm_email.lower())

        if customer_name and customer_name not in ("nan", ""):
            if not lookup[code]["customer_name"]:
                lookup[code]["customer_name"] = customer_name

        if division and division not in ("nan", ""):
            lookup[code]["divisions"].append(division)

    log.info("Lookup built: %d unique stockist codes", len(lookup))
    return lookup


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3 – Find attachments with DUAL date verification
# ─────────────────────────────────────────────────────────────────────────────

# Filename convention:  {code}_{YYYYMMDD}_{HHMMSS}_{originalname}.ext
FILE_PATTERN = re.compile(
    r"^(?P<code>\d+)_(?P<date>\d{8})_(?P<time>\d{6})_.+\.\w+$",
    re.IGNORECASE,
)


def _parse_filename_dt(filename: str):
    """Extract (stockist_code, datetime) embedded in filename, or (None, None)."""
    m = FILE_PATTERN.match(filename)
    if not m:
        return None, None
    try:
        dt = datetime.strptime(m.group("date") + m.group("time"), "%Y%m%d%H%M%S")
        return m.group("code"), dt
    except ValueError:
        return None, None


def _received_date_from_files(file_paths: list) -> str:
    """
    Return the date string (DD-Mon-YYYY) extracted from the FIRST matched
    filename's embedded timestamp.  Falls back to today if none found.
    """
    for fp in file_paths:
        _, fn_dt = _parse_filename_dt(fp.name)
        if fn_dt:
            return fn_dt.strftime("%d-%b-%Y")
    return datetime.today().strftime("%d-%b-%Y")


def find_attachments(
    folder: str,
    stockist_codes: set,
    start: datetime,
    end: datetime,
) -> dict:
    """
    Returns { stockist_code: [Path, ...] }

    DUAL verification per file:
      1. Filename timestamp must fall within [start, end]
      2. File OS modified-time must ALSO fall within [start, end]
    Both must pass — prevents mislabelled or copied files sneaking in.
    """
    folder_path = Path(folder)
    if not folder_path.exists():
        log.warning("Attachments folder does not exist: %s", folder)
        return {}

    result: dict = {}
    counts = {
        "matched": 0, "bad_name": 0, "wrong_code": 0,
        "fn_out_range": 0, "mtime_out_range": 0,
    }

    for fp in folder_path.rglob("*"):
        if not fp.is_file():
            continue

        code, fn_dt = _parse_filename_dt(fp.name)

        if code is None:
            counts["bad_name"] += 1
            continue

        if code not in stockist_codes:
            counts["wrong_code"] += 1
            continue

        # Check 1: filename timestamp
        if fn_dt is None or not (start <= fn_dt <= end):
            counts["fn_out_range"] += 1
            log.debug("SKIP (filename dt out of range): %s", fp.name)
            continue

        # Check 2: OS modified time
        mtime = datetime.fromtimestamp(fp.stat().st_mtime)
        if not (start <= mtime <= end):
            counts["mtime_out_range"] += 1
            log.warning(
                "SKIP (mtime %s outside window) → %s",
                mtime.strftime("%Y-%m-%d %H:%M:%S"), fp.name,
            )
            continue

        result.setdefault(code, []).append(fp)
        counts["matched"] += 1

    log.info(
        "Attachment scan → matched=%d | bad_name=%d | wrong_code=%d | "
        "fn_out_range=%d | mtime_out_range=%d",
        counts["matched"], counts["bad_name"], counts["wrong_code"],
        counts["fn_out_range"], counts["mtime_out_range"],
    )
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 4 – Send via Outlook shared mailbox (win32com)
# ─────────────────────────────────────────────────────────────────────────────

def get_outlook_app():
    """Connect to the running Outlook instance."""
    if not WIN32_AVAILABLE:
        raise EnvironmentError(
            "pywin32 is not installed.\n"
            "Run:  pip install pywin32\n"
            "Then: python -m win32com.client.makepy"
        )
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as exc:
        raise EnvironmentError(f"Could not connect to Outlook: {exc}") from exc


def send_via_outlook(
    outlook,
    from_address: str,
    bcc_addresses: list,
    subject: str,
    body: str,
    attachment_paths: list,
    dry_run: bool = True,
) -> str:
    """
    Create and optionally send one email FROM the shared mailbox.

    To: is intentionally left blank.
    All TBM and ABM addresses go into BCC only.

    dry_run=True  → saves to Drafts (safe to inspect)
    dry_run=False → calls mail.Send()
    """
    mail = outlook.CreateItem(0)   # 0 = olMailItem

    # Send FROM the shared mailbox
    mail.SentOnBehalfOfName = from_address

    mail.Subject = subject
    mail.Body    = body

    # To: intentionally left blank — BCC only
    mail.To = ""

    if bcc_addresses:
        mail.BCC = "; ".join(bcc_addresses)

    for fp in attachment_paths:
        mail.Attachments.Add(str(fp.resolve()))

    if dry_run:
        mail.Save()     # goes to Drafts folder of the shared mailbox
        return "DRAFT SAVED"
    else:
        mail.Send()
        return "SENT"


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5 – Build dispatch plan
#           ONE email per unique customer email ID
#           All TBM + ABM in BCC, To: blank
#           All attachments for that customer in the same email
# ─────────────────────────────────────────────────────────────────────────────

def build_dispatch_plan(
    filtered_df: pd.DataFrame,
    lookup: dict,
    attachments_map: dict,
) -> list:
    """
    Returns one dispatch record per stockist code.

    Rules enforced:
      - To:  always empty
      - BCC: union of all TBM emails + all ABM emails for that code (deduped)
      - All attachments belonging to the code in one email
      - Subject/Body follow the configured template
    """
    records = []

    for _, row in filtered_df.iterrows():
        code           = row["CustomerNo"]
        stockist_email = row["EmailID"]
        timestamp      = row["Timestamp"]

        entry = lookup.get(code)
        if not entry:
            log.warning("Code %s not in Base File – skipping", code)
            records.append(_record(
                code, stockist_email, timestamp,
                status="SKIPPED – code not in Base File",
            ))
            continue

        # Union of TBM and ABM into one BCC set — no duplicates
        bcc_emails = entry["tbm_emails"] | entry["abm_emails"]
        bcc_emails.discard("")   # remove any blank that slipped in

        if not bcc_emails:
            log.warning("No valid TBM or ABM email for code %s – skipping", code)
            records.append(_record(
                code, stockist_email, timestamp,
                status="SKIPPED – no TBM/ABM email found",
            ))
            continue

        files = attachments_map.get(code, [])
        if not files:
            log.warning("No attachments found for code %s in window", code)
            # Email will still be sent without attachments.
            # To SKIP when no files found, uncomment the two lines below:
            # records.append(_record(code, stockist_email, timestamp,
            #                        status="SKIPPED – no attachments"))
            # continue

        received_date = _received_date_from_files(files)
        customer_name = entry["customer_name"] or "N/A"

        subject = EMAIL_SUBJECT.format(
            stockist_code=code,
        )
        body = EMAIL_BODY.format(
            stockist_code=code,
            received_date=received_date,
            customer_name=customer_name,
        )

        records.append(_record(
            code, stockist_email, timestamp,
            bcc       = "; ".join(sorted(bcc_emails)),
            divisions = ", ".join(entry["divisions"]),
            num_files = len(files),
            file_names= "; ".join(f.name for f in files),
            subject   = subject,
            body      = body,
            status    = "PENDING",
        ))

    pending = sum(1 for r in records if r["Status"] == "PENDING")
    log.info("Dispatch plan: %d records total | %d PENDING", len(records), pending)
    return records


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 6 – Save dispatch log to Excel
# ─────────────────────────────────────────────────────────────────────────────

def save_dispatch_log(records: list, output_path: str):
    df = pd.DataFrame(records)
    # Body column is for sending only – exclude from the Excel log
    df.drop(columns=["Body"], errors="ignore", inplace=True)
    col_order = [
        "Stockist Code", "Stockist Email", "Submission Datetime",
        "Divisions", "BCC Emails",
        "Num Attachments", "Attachment Files",
        "Subject", "Status",
    ]
    col_order = [c for c in col_order if c in df.columns]
    df[col_order].to_excel(output_path, index=False)
    log.info("Dispatch log saved → %s  (%d rows)", output_path, len(df))


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    start_dt = datetime.strptime(START_DATETIME, "%Y-%m-%d %H:%M:%S")
    end_dt   = datetime.strptime(END_DATETIME,   "%Y-%m-%d %H:%M:%S")

    log.info("=" * 70)
    log.info("Stockist Mail Dispatcher  v3")
    log.info("From   : %s", SHARED_MAILBOX)
    log.info("Window : %s  →  %s", start_dt, end_dt)
    log.info("Mode   : %s", "DRY-RUN (Drafts only)" if not SEND_EMAILS else "LIVE – emails will be sent")
    log.info("=" * 70)

    # 1. Load & filter log
    log_df      = load_log(STOCKIST_LOG_PATH)
    filtered_df = filter_log(log_df, start_dt, end_dt)

    if filtered_df.empty:
        log.warning("No PASS entries in window. Nothing to do.")
        return

    # 2. Base file lookup
    base_df = load_base(BASE_FILE_PATH)
    lookup  = build_lookup(base_df)

    # 3. Find attachments (dual date check)
    stockist_codes  = set(filtered_df["CustomerNo"].tolist())
    attachments_map = find_attachments(ATTACHMENTS_FOLDER, stockist_codes, start_dt, end_dt)

    # 4. Build dispatch plan
    records = build_dispatch_plan(filtered_df, lookup, attachments_map)

    if not any(r["Status"] == "PENDING" for r in records):
        log.warning("No PENDING emails after plan build. Check warnings above.")
        save_dispatch_log(records, OUTPUT_LOG_PATH)
        return

    # 5. Connect to Outlook
    try:
        outlook = get_outlook_app()
        log.info("Connected to Outlook.")
    except EnvironmentError as exc:
        log.error("%s", exc)
        for r in records:
            if r["Status"] == "PENDING":
                r["Status"] = f"ERROR – Outlook not available: {exc}"
        save_dispatch_log(records, OUTPUT_LOG_PATH)
        return

    # 6. Send / save drafts
    #    One email per record == one email per unique customer email ID
    for rec in records:
        if rec["Status"] != "PENDING":
            continue

        code      = rec["Stockist Code"]
        bcc_list  = [e.strip() for e in rec["BCC Emails"].split(";") if e.strip()]
        files     = attachments_map.get(code, [])

        try:
            status = send_via_outlook(
                outlook          = outlook,
                from_address     = SHARED_MAILBOX,
                bcc_addresses    = bcc_list,
                subject          = rec["Subject"],
                body             = rec["Body"],
                attachment_paths = files,
                dry_run          = not SEND_EMAILS,
            )
            rec["Status"] = status
            log.info(
                "  [%s] BCC=%s | Files=%d → %s",
                code, rec["BCC Emails"],
                rec["Num Attachments"], status,
            )
        except Exception as exc:
            rec["Status"] = f"FAILED – {exc}"
            log.error("  [%s] FAILED: %s", code, exc)

    # 7. Save log
    save_dispatch_log(records, OUTPUT_LOG_PATH)

    sent    = sum(1 for r in records if r["Status"] == "SENT")
    drafted = sum(1 for r in records if r["Status"] == "DRAFT SAVED")
    failed  = sum(1 for r in records if r["Status"] not in ("SENT", "DRAFT SAVED"))

    log.info("=" * 70)
    log.info("DONE  |  Sent=%d  Drafts=%d  Skipped/Failed=%d", sent, drafted, failed)
    log.info("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
#  Helper – single dispatch record dict
# ─────────────────────────────────────────────────────────────────────────────

def _record(
    code, stockist_email, timestamp,
    bcc="", divisions="", num_files=0, file_names="",
    subject="", body="", status="",
):
    return {
        "Stockist Code":       code,
        "Stockist Email":      stockist_email,
        "Submission Datetime": (
            timestamp.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(timestamp, "strftime") else str(timestamp)
        ),
        "Divisions":           divisions,
        "BCC Emails":          bcc,
        "Num Attachments":     num_files,
        "Attachment Files":    file_names,
        "Subject":             subject,
        "Body":                body,   # used for sending, stripped before Excel write
        "Status":              status,
    }


if __name__ == "__main__":
    main()
