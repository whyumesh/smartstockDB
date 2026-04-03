"""
Stockist Mail Dispatcher
========================
Reads the Stockist Log, filters PASS entries within a date-time window,
looks up TBM & ABM emails from the Base File, picks matching attachment
files from a folder, drafts emails, and produces a dispatch log.

Usage:
    python stockist_mail_dispatcher.py

Configure the section marked  ── USER CONFIG ──  below.
"""

import os
import re
import smtplib
import logging
import pandas as pd

from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════════════════
#  ── USER CONFIG ──  (edit only this section)
# ═══════════════════════════════════════════════════════════════════════════════

# Date-time window to filter the log (inclusive on both ends)
# Format: "YYYY-MM-DD HH:MM:SS"
START_DATETIME = "2026-03-13 00:00:00"
END_DATETIME   = "2026-03-13 23:59:59"

# Paths
STOCKIST_LOG_PATH = r"Stockist_Log.xlsx"                                         # Log Excel file
BASE_FILE_PATH    = r"Umesh_Mail_Reminder_Base_File_March_2026_dummy.xlsx"       # Base/master Excel file
ATTACHMENTS_FOLDER = r"C:\Stockist_Attachments"                                  # Folder containing saved attachments
OUTPUT_LOG_PATH    = r"dispatch_log.xlsx"                                         # Where the dispatch report is saved

# SMTP settings  (set SEND_EMAILS=False to do a dry-run without actually sending)
SEND_EMAILS   = False          # ← flip to True when ready to go live
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_USER     = "your_email@gmail.com"
SMTP_PASSWORD = "your_app_password"
SENDER_NAME   = "Stockist Mailer"

# Email content
EMAIL_SUBJECT = "Stockist Data Submission – {stockist_code}"
EMAIL_BODY    = """\
Dear {tbm_name},

Please find attached the data submission received from Stockist {stockist_code}
on {submission_datetime}.

ABM ({abm_email}) has also been marked in CC.

Regards,
Stockist Operations Team
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  ── END OF USER CONFIG ──
# ═══════════════════════════════════════════════════════════════════════════════


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────────────────────────
#  STEP 1 – Load and clean the Stockist Log
# ──────────────────────────────────────────────────────────────────────────────

def load_log(path: str) -> pd.DataFrame:
    """Load the Stockist Log and return a cleaned DataFrame."""
    log.info("Loading Stockist Log: %s", path)
    df = pd.read_excel(path, sheet_name="Sheet1", dtype={"CustomerNo": str})

    # ── clean Status column  (raw values look like  "Pass'\n"  or  "Fail'")
    df["Status"] = (
        df["Status"]
        .astype(str)
        .str.strip()
        .str.strip("'")
        .str.strip("\\n")
        .str.strip()
    )

    # ── clean CustomerNo  (some rows contain  "'"  for failed entries)
    df["CustomerNo"] = df["CustomerNo"].str.strip().str.strip("'").str.strip()

    # ── ensure Timestamp is datetime
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")

    log.info("Log loaded: %d rows total", len(df))
    return df


def filter_log(df: pd.DataFrame, start: datetime, end: datetime) -> pd.DataFrame:
    """Keep only PASS rows within [start, end]."""
    mask = (
        (df["Status"].str.upper() == "PASS") &
        (df["Timestamp"] >= start) &
        (df["Timestamp"] <= end) &
        (df["CustomerNo"].str.match(r"^\d+$", na=False))   # valid numeric codes only
    )
    filtered = df[mask].copy()
    log.info("Filtered rows (PASS, in window): %d", len(filtered))

    # De-duplicate: keep the latest entry per (CustomerNo, EmailID) pair
    filtered.sort_values("Timestamp", ascending=False, inplace=True)
    filtered.drop_duplicates(subset=["CustomerNo", "EmailID"], inplace=True)
    log.info("After dedup (latest per code+email): %d rows", len(filtered))
    return filtered


# ──────────────────────────────────────────────────────────────────────────────
#  STEP 2 – Load the Base File and build a lookup dict
# ──────────────────────────────────────────────────────────────────────────────

def load_base(path: str) -> pd.DataFrame:
    """Load the Base File and return a DataFrame keyed on c_cust_code."""
    log.info("Loading Base File: %s", path)
    df = pd.read_excel(path, dtype={"c_cust_code": str})
    df["c_cust_code"] = df["c_cust_code"].astype(str).str.strip()

    # Normalise email columns
    for col in ["Email", "ABM Email", "TBM Employ. Name", "ABM_Name"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "None": "", "NOT ASSIGNED": ""})

    log.info("Base File loaded: %d rows", len(df))
    return df


def build_lookup(base_df: pd.DataFrame) -> dict:
    """
    Returns a dict  {stockist_code: [list of dicts with tbm_email, abm_email, ...]}
    One stockist can appear in multiple divisions, so we return a list.
    """
    lookup: dict = {}
    for _, row in base_df.iterrows():
        code = str(row["c_cust_code"]).strip()
        if not code or code == "nan":
            continue
        entry = {
            "tbm_email":  row.get("Email", ""),
            "tbm_name":   row.get("TBM Employ. Name", ""),
            "abm_email":  row.get("ABM Email", ""),
            "abm_name":   row.get("ABM_Name", ""),
            "division":   row.get("Division Name", ""),
        }
        lookup.setdefault(code, []).append(entry)
    log.info("Lookup built for %d unique stockist codes", len(lookup))
    return lookup


# ──────────────────────────────────────────────────────────────────────────────
#  STEP 3 – Match attachment files from the folder
# ──────────────────────────────────────────────────────────────────────────────

# Expected filename convention:
#   {stockistcode}_{YYYYMMDD}_{HHMMSS}_{originalfilename}.ext
# e.g.  10009847_20260313_103759_sales_report.xlsx

FILE_PATTERN = re.compile(
    r"^(?P<code>\d+)_(?P<date>\d{8})_(?P<time>\d{6})_.+$",
    re.IGNORECASE,
)


def parse_file_datetime(filename: str) -> datetime | None:
    """Extract datetime from filename convention; return None if unparseable."""
    m = FILE_PATTERN.match(filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group("date") + m.group("time"), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def find_attachments(
    folder: str,
    stockist_codes: set,
    start: datetime,
    end: datetime,
) -> dict:
    """
    Scan *folder* recursively and return a dict:
        { stockist_code: [Path, Path, ...] }
    Only files whose embedded datetime falls within [start, end] and whose
    code is in *stockist_codes* are included.
    """
    folder_path = Path(folder)
    if not folder_path.exists():
        log.warning("Attachments folder not found: %s", folder)
        return {}

    result: dict = {}
    skipped = 0

    for fp in folder_path.rglob("*"):
        if not fp.is_file():
            continue
        m = FILE_PATTERN.match(fp.name)
        if not m:
            skipped += 1
            continue

        code      = m.group("code")
        file_dt   = parse_file_datetime(fp.name)

        if code not in stockist_codes:
            continue
        if file_dt is None or not (start <= file_dt <= end):
            continue

        result.setdefault(code, []).append(fp)

    log.info(
        "Attachment scan complete: %d codes matched, %d files skipped (name mismatch)",
        len(result), skipped,
    )
    return result


# ──────────────────────────────────────────────────────────────────────────────
#  STEP 4 – Build and send emails
# ──────────────────────────────────────────────────────────────────────────────

def build_email(
    sender_addr: str,
    sender_name: str,
    to_addr: str,
    cc_addr: str,
    subject: str,
    body: str,
    attachments: list[Path],
) -> MIMEMultipart:
    """Construct a MIME email with optional attachments."""
    msg = MIMEMultipart()
    msg["From"]    = f"{sender_name} <{sender_addr}>"
    msg["To"]      = to_addr
    msg["Cc"]      = cc_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    for fp in attachments:
        with open(fp, "rb") as fh:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{fp.name}"',
        )
        msg.attach(part)

    return msg


def send_email(smtp_conn: smtplib.SMTP, msg: MIMEMultipart, all_recipients: list[str]):
    try:
        smtp_conn.sendmail(msg["From"], all_recipients, msg.as_string())
        return True, "Sent"
    except smtplib.SMTPException as exc:
        return False, str(exc)


# ──────────────────────────────────────────────────────────────────────────────
#  STEP 5 – Dispatch log
# ──────────────────────────────────────────────────────────────────────────────

def save_dispatch_log(records: list[dict], output_path: str):
    """Write the dispatch results to an Excel file."""
    df = pd.DataFrame(records)
    df.to_excel(output_path, index=False)
    log.info("Dispatch log saved → %s  (%d rows)", output_path, len(df))


# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    start_dt = datetime.strptime(START_DATETIME, "%Y-%m-%d %H:%M:%S")
    end_dt   = datetime.strptime(END_DATETIME,   "%Y-%m-%d %H:%M:%S")

    log.info("=" * 70)
    log.info("Stockist Mail Dispatcher")
    log.info("Window: %s  →  %s", start_dt, end_dt)
    log.info("=" * 70)

    # ── 1. Load & filter log
    log_df      = load_log(STOCKIST_LOG_PATH)
    filtered_df = filter_log(log_df, start_dt, end_dt)

    if filtered_df.empty:
        log.warning("No PASS entries found in the given window. Nothing to send.")
        return

    # ── 2. Load base file & build lookup
    base_df = load_base(BASE_FILE_PATH)
    lookup  = build_lookup(base_df)

    # ── 3. Find attachments
    stockist_codes_in_window = set(filtered_df["CustomerNo"].tolist())
    attachments_map = find_attachments(
        ATTACHMENTS_FOLDER,
        stockist_codes_in_window,
        start_dt,
        end_dt,
    )

    # ── 4. Build dispatch plan
    dispatch_records: list[dict] = []

    for _, row in filtered_df.iterrows():
        code       = row["CustomerNo"]
        stockist_email = row["EmailID"]
        timestamp  = row["Timestamp"]

        base_entries = lookup.get(code)
        if not base_entries:
            log.warning("Code %s not found in Base File — skipping.", code)
            dispatch_records.append(_record(
                code, stockist_email, timestamp,
                status="SKIPPED – code not in Base File",
            ))
            continue

        files = attachments_map.get(code, [])
        if not files:
            log.warning("No attachment files found for code %s in window.", code)
            # Still send the email (without attachments) – adjust logic here if needed
            # Uncomment next two lines to SKIP instead:
            # dispatch_records.append(_record(code, stockist_email, timestamp, status="SKIPPED – no files found"))
            # continue

        # One code can map to multiple divisions (multiple TBM/ABM pairs)
        for entry in base_entries:
            tbm_email = entry["tbm_email"]
            abm_email = entry["abm_email"]
            tbm_name  = entry["tbm_name"]
            division  = entry["division"]

            if not tbm_email and not abm_email:
                log.warning(
                    "No TBM or ABM email for code %s / division %s — skipping entry.",
                    code, division,
                )
                dispatch_records.append(_record(
                    code, stockist_email, timestamp,
                    tbm=tbm_email, abm=abm_email, division=division,
                    status="SKIPPED – no TBM/ABM email",
                ))
                continue

            subject = EMAIL_SUBJECT.format(stockist_code=code)
            body    = EMAIL_BODY.format(
                tbm_name           = tbm_name or tbm_email,
                stockist_code      = code,
                submission_datetime= timestamp.strftime("%d-%b-%Y %H:%M:%S"),
                abm_email          = abm_email,
            )

            dispatch_records.append(_record(
                code, stockist_email, timestamp,
                tbm=tbm_email, abm=abm_email, division=division,
                num_files=len(files),
                file_names="; ".join(f.name for f in files),
                subject=subject,
                status="PENDING",
            ))

    log.info("Dispatch plan built: %d email actions", len(dispatch_records))

    # ── 5. Send
    if not SEND_EMAILS:
        log.info("SEND_EMAILS=False — dry-run mode. No emails will be sent.")
        for rec in dispatch_records:
            if rec["Status"] == "PENDING":
                rec["Status"] = "DRY-RUN (not sent)"
    else:
        log.info("Connecting to SMTP: %s:%d", SMTP_HOST, SMTP_PORT)
        try:
            smtp = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
            smtp.ehlo()
            smtp.starttls()
            smtp.login(SMTP_USER, SMTP_PASSWORD)
        except Exception as exc:
            log.error("SMTP connection failed: %s", exc)
            for rec in dispatch_records:
                if rec["Status"] == "PENDING":
                    rec["Status"] = f"SMTP ERROR – {exc}"
            save_dispatch_log(dispatch_records, OUTPUT_LOG_PATH)
            return

        for rec in dispatch_records:
            if rec["Status"] != "PENDING":
                continue

            code      = rec["Stockist Code"]
            tbm_email = rec["TBM Email"]
            abm_email = rec["ABM Email"]
            subject   = rec["Subject"]
            files     = attachments_map.get(code, [])

            body = EMAIL_BODY.format(
                tbm_name           = rec.get("TBM Name", tbm_email),
                stockist_code      = code,
                submission_datetime= rec["Submission Datetime"],
                abm_email          = abm_email,
            )

            recipients = [r for r in [tbm_email, abm_email] if r]
            msg = build_email(
                sender_addr = SMTP_USER,
                sender_name = SENDER_NAME,
                to_addr     = tbm_email,
                cc_addr     = abm_email,
                subject     = subject,
                body        = body,
                attachments = files,
            )

            success, status_msg = send_email(smtp, msg, recipients)
            rec["Status"] = "SENT" if success else f"FAILED – {status_msg}"
            log.info("  [%s] → TBM:%s CC:%s  | %s", code, tbm_email, abm_email, rec["Status"])

        smtp.quit()
        log.info("SMTP connection closed.")

    # ── 6. Save dispatch log
    save_dispatch_log(dispatch_records, OUTPUT_LOG_PATH)
    log.info("Done.")


# ──────────────────────────────────────────────────────────────────────────────
#  Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _record(
    code, stockist_email, timestamp,
    tbm="", abm="", division="",
    num_files=0, file_names="",
    subject="", status="",
):
    return {
        "Stockist Code":       code,
        "Stockist Email":      stockist_email,
        "Submission Datetime": timestamp.strftime("%Y-%m-%d %H:%M:%S") if hasattr(timestamp, "strftime") else str(timestamp),
        "Division":            division,
        "TBM Email":           tbm,
        "ABM Email":           abm,
        "Num Attachments":     num_files,
        "Attachment Files":    file_names,
        "Subject":             subject,
        "Status":              status,
    }


if __name__ == "__main__":
    main()
