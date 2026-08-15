"""Excel reading, template rendering, and sequential email sending."""

from __future__ import annotations

import csv
import json
import re
import smtplib
import ssl
from datetime import datetime
from email.header import Header
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path
from typing import Any

import pandas as pd

LOG_DIR = Path("logs")
STATE_PATH = LOG_DIR / "send_state.json"
LOG_PATH = LOG_DIR / "send_log.csv"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}|\{([^{}]+)\}")
HIDDEN_CHARS_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069]")


def excel_row_to_index(excel_row: int, header_row: int = 1) -> int:
    """Convert a 1-based Excel row number to a pandas 0-based index."""
    return excel_row - header_row - 1


def index_to_excel_row(index: int, header_row: int = 1) -> int:
    return index + header_row + 1


def load_excel(file_or_path: Any) -> pd.DataFrame:
    df = pd.read_excel(file_or_path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df.reset_index(drop=True)


def detect_email_column(columns: list[str]) -> str | None:
    keys = ("email", "e-mail", "mail", "ايميل", "إيميل", "بريد")
    for col in columns:
        lowered = col.lower()
        if any(k in lowered for k in keys):
            return col
    return columns[0] if columns else None


def slice_rows(
    df: pd.DataFrame,
    start_excel_row: int,
    end_excel_row: int,
    header_row: int = 1,
) -> pd.DataFrame:
    start = max(excel_row_to_index(start_excel_row, header_row), 0)
    end = excel_row_to_index(end_excel_row, header_row)
    if start > end:
        start, end = end, start
    end = min(end, len(df) - 1)
    start = min(max(start, 0), max(len(df) - 1, 0))
    if df.empty:
        return df
    return df.iloc[start : end + 1].copy()


def sanitize_email(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return HIDDEN_CHARS_RE.sub("", str(value)).strip()


def is_valid_email(value: Any) -> bool:
    return bool(EMAIL_RE.match(sanitize_email(value)))


def cell_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def render_template(template: str, row: dict[str, Any]) -> str:
    lookup = {str(k).strip(): cell_text(v) for k, v in row.items()}
    lookup_ci = {k.lower(): v for k, v in lookup.items()}

    def replace(match: re.Match[str]) -> str:
        key = (match.group(1) or match.group(2) or "").strip()
        if key in lookup:
            return lookup[key]
        return lookup_ci.get(key.lower(), match.group(0))

    return PLACEHOLDER_RE.sub(replace, template or "")


def build_message(
    sender_email: str,
    sender_name: str,
    to_email: str,
    subject: str,
    body: str,
    html: bool,
) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = subject or "(بدون عنوان)"
    if sender_name:
        msg["From"] = formataddr((str(Header(sender_name, "utf-8")), sender_email))
    else:
        msg["From"] = sender_email
    msg["To"] = to_email
    msg["Date"] = formatdate(localtime=True)
    domain = sender_email.split("@")[-1] if "@" in sender_email else "gmail.com"
    msg["Message-ID"] = make_msgid(domain=domain)
    msg["Reply-To"] = sender_email
    msg["List-Unsubscribe"] = f"<mailto:{sender_email}?subject=unsubscribe>"
    if html:
        msg.set_content("فتح الرسالة يتطلب عرض HTML.")
        msg.add_alternative(body, subtype="html")
    else:
        msg.set_content(body or " ")
    return msg


def _deliver(server: smtplib.SMTP, msg: EmailMessage, sender_email: str, recipients: list[str]) -> None:
    refused = server.send_message(msg, from_addr=sender_email, to_addrs=recipients)
    if refused:
        raise RuntimeError(f"السيرفر رفض المستقبل: {refused}")


def _send_with_ssl(
    host: str,
    port: int,
    username: str,
    password: str,
    msg: EmailMessage,
    sender_email: str,
    recipients: list[str],
) -> None:
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
        server.login(username, password)
        _deliver(server, msg, sender_email, recipients)


def _send_with_starttls(
    host: str,
    port: int,
    username: str,
    password: str,
    msg: EmailMessage,
    sender_email: str,
    recipients: list[str],
) -> None:
    context = ssl.create_default_context()
    with smtplib.SMTP(host, port, timeout=30) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(username, password)
        _deliver(server, msg, sender_email, recipients)


def send_one(
    *,
    host: str,
    port: int,
    username: str,
    password: str,
    sender_email: str,
    sender_name: str,
    to_email: str,
    subject: str,
    body: str,
    html: bool,
    use_ssl: bool = False,
    copy_to_sender: bool = False,
) -> None:
    port = int(port)
    host = host.strip()
    username = username.strip()
    sender_email = sanitize_email(sender_email)
    to_email = sanitize_email(to_email)
    msg = build_message(sender_email, sender_name, to_email, subject, body, html)
    recipients = [to_email]
    if copy_to_sender and sender_email.lower() not in {to_email.lower()}:
        recipients.append(sender_email)

    if port == 465:
        implicit_ssl = True
    elif port in (587, 25):
        implicit_ssl = False
    else:
        implicit_ssl = use_ssl

    first = _send_with_ssl if implicit_ssl else _send_with_starttls
    second = _send_with_starttls if implicit_ssl else _send_with_ssl
    kwargs = dict(
        host=host,
        port=port,
        username=username,
        password=password,
        msg=msg,
        sender_email=sender_email,
        recipients=recipients,
    )
    try:
        first(**kwargs)
    except ssl.SSLError as exc:
        if "WRONG_VERSION_NUMBER" not in str(exc).upper() and "WRONG_VERSION" not in str(exc).upper():
            raise
        second(**kwargs)


def clear_sent_state() -> None:
    save_state({"sent_keys": []})


def ensure_log_dir() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def load_state() -> dict[str, Any]:
    ensure_log_dir()
    if not STATE_PATH.exists():
        return {"sent_keys": []}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"sent_keys": []}


def save_state(state: dict[str, Any]) -> None:
    ensure_log_dir()
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def row_key(excel_row: int, email: str) -> str:
    return f"{excel_row}|{email.strip().lower()}"


def already_sent(state: dict[str, Any], excel_row: int, email: str) -> bool:
    return row_key(excel_row, email) in set(state.get("sent_keys", []))


def mark_sent(state: dict[str, Any], excel_row: int, email: str) -> None:
    keys = list(state.get("sent_keys", []))
    key = row_key(excel_row, email)
    if key not in keys:
        keys.append(key)
    state["sent_keys"] = keys
    save_state(state)


def append_log(row: dict[str, Any]) -> None:
    ensure_log_dir()
    exists = LOG_PATH.exists()
    with LOG_PATH.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["time", "excel_row", "email", "status", "subject", "error"],
        )
        if not exists:
            writer.writeheader()
        writer.writerow(row)


def log_result(
    excel_row: int,
    email: str,
    status: str,
    subject: str,
    error: str = "",
) -> None:
    append_log(
        {
            "time": datetime.now().isoformat(timespec="seconds"),
            "excel_row": excel_row,
            "email": email,
            "status": status,
            "subject": subject,
            "error": error,
        }
    )
