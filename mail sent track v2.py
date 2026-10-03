import socket
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import threading
import smtplib
import imaplib
import email as email_lib
from email.header import decode_header
import pandas as pd
import time
import logging
import copy
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, date
import os
import json
import re
import ssl

# ── Gmail Signature (OAuth) ──────────────────────────────────────
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.settings.basic"]

# Always resolve these next to the script itself, regardless of the
# folder the script was launched from (fixes "No such file or directory").
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.pickle")

def get_gmail_signature(gmail_address):
    """Fetch the user's default Gmail signature via OAuth. Returns HTML string."""
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                raise FileNotFoundError(
                    f"credentials.json not found at {CREDENTIALS_PATH}. "
                    f"Place it in the same folder as this script."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)

    service = build("gmail", "v1", credentials=creds)
    result = service.users().settings().sendAs().get(
        userId="me", sendAsEmail=gmail_address
    ).execute()
    raw_sig = result.get("signature", "")
    return _sanitize_signature_html(raw_sig)


def _sanitize_signature_html(sig_html):
    """
    Gmail's API returns the signature as a full HTML document (often with its
    own <!DOCTYPE>, <html>, <head>, and <body> tags). Nesting that directly
    inside another <html><body> block is invalid and causes clients to mangle
    the main message. Strip those wrapper tags so only the inner content
    survives, then wrap it in a plain <div> as a normal inline block.
    """
    if not sig_html.strip():
        return ""
    cleaned = sig_html
    # Drop full <head>...</head> blocks entirely (styles/meta not needed inline)
    cleaned = re.sub(r'(?is)<head\b.*?</head>', '', cleaned)
    # Strip the outer html/body wrapper tags, keep their inner content
    cleaned = re.sub(r'(?is)</?html[^>]*>', '', cleaned)
    cleaned = re.sub(r'(?is)</?body[^>]*>', '', cleaned)
    cleaned = re.sub(r'(?is)<!DOCTYPE[^>]*>', '', cleaned)
    # Gmail marks its own signatures with class="gmail_signature" and
    # data-smartmail="gmail_signature". Recipient Gmail clients auto-collapse
    # anything carrying that marker into the "quoted text ..." section, so
    # strip it out or the signature disappears into the trimmed part.
    cleaned = re.sub(r'\s*data-smartmail="gmail_signature"', '', cleaned)
    cleaned = re.sub(r'(class="[^"]*)\bgmail_signature\b\s*', r'\1', cleaned)
    cleaned = re.sub(r'class="\s*"', '', cleaned)
    return f'<div style="font-family:Arial,sans-serif;font-size:13px;color:#333;">{cleaned.strip()}</div>'

# ── Config ─────────────────────────────────────────────────────
LOG_FILE      = "email_log.txt"
SAVE_FILE     = "draft.json"
PROGRESS_FILE = "manual_fu_progress.json"
BATCH_SIZE    = 200
# Pause between recipients that only trigger an IMAP lookup (no email sent) —
# i.e. a skip because no original thread was found in the label, or because
# the recipient already replied at least once. This is a fixed, short wait
# (per spec: 1 second) — it only needs to avoid hammering Gmail's IMAP server
# with rapid-fire commands. It is intentionally decoupled from the user's
# configured send delay, which paces actual SENDS against Gmail's outbound
# sending rate limit — a very different constraint.
IMAP_SKIP_PACE = 1

logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s | %(levelname)s | %(message)s")

# ── Save / Load draft ───────────────────────────────────────────
def save_draft(subject, body, sheet_url, pdf_path, fu_body="", threads=0,
               mfu_sheet_url="", mfu_label="", mfu_delay=30,
               mfu_active_index=0, mfu_followups=None):
    """Save both main draft and all follow-up sub-tabs drafts to JSON."""
    data = {
        "subject": subject,
        "body": body,
        "sheet_url": sheet_url,
        "pdf_path": pdf_path,
        "fu_body": fu_body,
        "threads": threads,
        "mfu_sheet_url": mfu_sheet_url,
        "mfu_label": mfu_label,
        "mfu_delay": mfu_delay,
        "mfu_active_index": mfu_active_index,
        "mfu_followups": mfu_followups if mfu_followups is not None else []
    }
    with open(SAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_draft():
    """Load saved drafts from JSON if available, with backwards compatibility."""
    if os.path.exists(SAVE_FILE):
        try:
            with open(SAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.warning(f"Failed to load draft.json: {e}")
    return {
        "subject": "", "body": "", "sheet_url": "", "pdf_path": "",
        "fu_body": "", "threads": 0, "mfu_sheet_url": "", "mfu_label": "Internship",
        "mfu_delay": 30, "mfu_active_index": 0, "mfu_followups": []
    }

# ── Progress tracker ────────────────────────────────────────────
def save_progress(total, sent_emails):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({"total": total, "sent": sent_emails,
                   "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}, f)

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def clear_progress():
    if os.path.exists(PROGRESS_FILE):
        os.remove(PROGRESS_FILE)

# ── Text → HTML ─────────────────────────────────────────────────
def text_to_html(text):
    lines = text.split('\n')
    html_lines = []
    for line in lines:
        if line.strip() == "":
            html_lines.append("<br>")
            continue
        stripped = line.lstrip(" ")
        leading  = len(line) - len(stripped)
        prefix   = "&nbsp;" * leading
        processed = stripped
        processed = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', processed)
        processed = re.sub(r'\*(.+?)\*',     r'<em>\1</em>',          processed)
        processed = re.sub(r'__(.+?)__',     r'<u>\1</u>',            processed)
        processed = re.sub(r'(?<!["\'=>])(https?://[^\s<>"]+)',
                           r'<a href="\1">\1</a>', processed)
        html_lines.append(prefix + processed)
    html = "<br>\n".join(html_lines)
    html = re.sub(r'(<br>\s*){3,}', '<br><br>', html)
    return html

# ── Google Sheet reader ─────────────────────────────────────────
def read_google_sheet(sheet_url):
    sheet_id = sheet_url.split("/d/")[1].split("/")[0]
    if "gid=" in sheet_url:
        gid = sheet_url.split("gid=")[1].split("&")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    else:
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    df = pd.read_csv(export_url)
    df = df.dropna(subset=["Email"])
    df = df.drop_duplicates(subset=["Email"])
    return df


# ── Sent-status checker (read-only Gmail IMAP; outputs CSV) ───────────────
def read_google_sheet_for_checker(sheet_url):
    """Read the public CSV export while preserving every source row."""
    sheet_id = sheet_url.split("/d/")[1].split("/")[0]
    if "gid=" in sheet_url:
        gid = sheet_url.split("gid=")[1].split("&")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    else:
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(export_url, dtype=str).fillna("")


def _imap_select_sent(imap):
    """Select Gmail's Sent Mail folder in read-only mode."""
    for mailbox in ('"[Gmail]/Sent Mail"', '"Sent Mail"', '"[Gmail]/Sent"', '"Sent"'):
        status, _ = imap.select(mailbox, readonly=True)
        if status == "OK":
            return mailbox
    raise RuntimeError("Could not open Gmail's Sent Mail folder. Make sure IMAP is enabled.")


def _imap_search_sent_for_recipient(imap, to_email):
    """Return True if Gmail Sent Mail contains a message addressed to to_email."""
    safe_email = str(to_email).strip().replace("\\", "\\\\").replace('"', '\\"')
    status, data = imap.search(None, f'(TO "{safe_email}")')
    if status != "OK":
        raise RuntimeError(f"Gmail search failed for {to_email}")
    return bool(data and data[0].split())


def check_sent_status(sheet_url, gmail, password, log_callback, done_callback, stop_event):
    """Check each recipient against Gmail Sent Mail and export a CSV."""
    imap = None
    try:
        log_callback("📡 Reading contacts from Google Sheet...\n")
        df = read_google_sheet_for_checker(sheet_url)
        required = {"Name", "Company", "Email"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError("Missing required column(s): " + ", ".join(sorted(missing)))

        total = len(df)
        log_callback(f"📋 {total} rows loaded (duplicates preserved)\n")
        log_callback("🔌 Connecting to Gmail IMAP...\n")
        imap = _imap_connect(gmail, password)
        mailbox = _imap_select_sent(imap)
        log_callback(f"✅ Opened Sent Mail: {mailbox}\n\n")

        statuses = []
        sent_count = not_sent_count = invalid_count = 0

        for pos, (_, row) in enumerate(df.iterrows(), start=1):
            if stop_event.is_set():
                log_callback("\n⛔ Sent-status check stopped. No partial CSV was exported.\n")
                done_callback(False, None)
                return

            to_email = str(row.get("Email", "")).strip()
            name = str(row.get("Name", "")).strip()

            if not to_email or "@" not in to_email:
                status = "Invalid Email"
                invalid_count += 1
            else:
                try:
                    sent = _imap_search_sent_for_recipient(imap, to_email)
                except (imaplib.IMAP4.abort, OSError, ssl.SSLError, EOFError):
                    log_callback("  ⚠️ IMAP connection dropped — reconnecting...\n")
                    try:
                        imap.logout()
                    except Exception:
                        pass
                    imap = _imap_connect(gmail, password)
                    _imap_select_sent(imap)
                    sent = _imap_search_sent_for_recipient(imap, to_email)

                if sent:
                    status = "Sent"
                    sent_count += 1
                else:
                    status = "Not Sent"
                    not_sent_count += 1

            statuses.append(status)
            icon = "✅" if status == "Sent" else "❌" if status == "Not Sent" else "⚠️"
            log_callback(f"[{pos}/{total}] {icon} {name} <{to_email}> — {status}\n")

        df["Status"] = statuses
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(SCRIPT_DIR, f"sent_status_{timestamp}.csv")
        df.to_csv(output_path, index=False, encoding="utf-8-sig")

        log_callback("\n" + "─" * 60 + "\n")
        log_callback("🎉 Check complete!\n")
        log_callback(f"✅ Sent: {sent_count}\n")
        log_callback(f"❌ Not Sent: {not_sent_count}\n")
        log_callback(f"⚠️ Invalid Email: {invalid_count}\n")
        log_callback(f"📄 CSV saved to:\n{output_path}\n")
        done_callback(True, output_path)
    except Exception as e:
        log_callback(f"\n❌ Sent-status checker error: {e}\n")
        done_callback(False, None)
    finally:
        if imap is not None:
            try:
                imap.logout()
            except Exception:
                pass

# ── IMAP helpers ────────────────────────────────────────────────
def decode_str(s):
    if s is None: return ""
    decoded, enc = decode_header(s)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(enc or "utf-8", errors="ignore")
    return decoded

def _imap_connect(gmail, password):
    """Open a fresh IMAP connection and log in. Used for the initial
    connection and to reconnect after a dropped socket."""
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(gmail, password)
    return imap

def _select_label(imap, label):
    """Selects a Gmail label, trying the raw name first, then the
    '[Gmail]/<label>' form some accounts require. Returns True/False."""
    status, _ = imap.select(f'"{label}"', readonly=True)
    if status == "OK":
        return True
    status, _ = imap.select(f'"[Gmail]/{label}"', readonly=True)
    return status == "OK"

# ── Manual follow-up decision logic ──────────────────────────────
# This is the exact two-step check the manual follow-up flow relies on,
# per recipient, in this order:
#
#   STEP 1 — Does an original email TO this recipient exist in the label?
#            If NO  -> SKIP ("no thread found"). Don't send a fresh,
#                      un-threaded email. Wait 1s, move to next recipient.
#
#   STEP 2 — Has this recipient EVER sent at least one reply that landed
#            in this label (i.e. anywhere in that thread — not just to
#            the latest message)?
#            If YES -> SKIP ("already replied"), permanently, regardless
#                      of who sent the most recent message in the thread.
#                      Once a human has replied even once, the thread is
#                      considered human-handled and the follow-up system
#                      never touches it again.
#            If NO  -> proceed to send the follow-up as a Reply-All into
#                      that thread, using the full user-set delay.
#
# Both lookups happen against the SAME selected label so we only pay for
# one IMAP SELECT (the expensive round-trip) per recipient.
def check_followup_status(imap, to_email, label):
    """
    Returns a dict:
      {"status": "no_thread"} -> nothing to do, recipient skipped
      {"status": "replied"}   -> recipient already replied at least once, skipped
      {"status": "ready", "message_id":..., "subject":..., "to":..., "cc":...}
                               -> safe to send the follow-up
    Raises on connection-level IMAP failures so the caller can reconnect.
    """
    try:
        if not _select_label(imap, label):
            return {"status": "no_thread"}

        # STEP 1 — does an original sent-to-them message exist in this label?
        _, res = imap.search(None, f'(TO "{to_email}")')
        to_ids = res[0].split()
        if not to_ids:
            return {"status": "no_thread"}

        # STEP 2 — has the recipient EVER replied into this label/thread?
        _, res2 = imap.search(None, f'(FROM "{to_email}")')
        if res2[0].split():
            return {"status": "replied"}

        # Original exists, no reply ever — fetch headers for Reply-All.
        _, data = imap.fetch(to_ids[-1],
            "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT TO CC)])")
        msg = email_lib.message_from_bytes(data[0][1])
        return {
            "status": "ready",
            "message_id": decode_str(msg.get("Message-ID", "")),
            "subject":    decode_str(msg.get("Subject", "")),
            "to":         decode_str(msg.get("To", "")),
            "cc":         decode_str(msg.get("CC", "")),
        }
    except (imaplib.IMAP4.abort, OSError, ssl.SSLError, EOFError):
        # Connection-level failure (dropped socket, closed by server, etc).
        # Propagate so the caller can reconnect and retry, instead of
        # silently mis-reporting every subsequent recipient.
        raise
    except Exception:
        return {"status": "no_thread"}

# ── Split a DataFrame into N near-equal chunks (always returns DataFrames) ──
def _split_dataframe(df, n):
    """
    np.array_split() is supposed to preserve DataFrame chunks, but on some
    numpy/pandas version combinations it silently returns raw numpy ndarrays
    instead — which have no .iterrows(). Slicing with .iloc guarantees a
    real DataFrame back every time, regardless of numpy/pandas versions.
    """
    n = max(1, min(n, len(df))) if len(df) else 1
    total = len(df)
    base, extra = divmod(total, n)
    chunks = []
    start = 0
    for i in range(n):
        size = base + (1 if i < extra else 0)
        if size == 0:
            continue
        chunks.append(df.iloc[start:start + size])
        start += size
    return chunks


# ── Build a single MIME message (shared by sequential + parallel paths) ──
def _build_message(email_from, to_email, subject_tpl, body_tpl, row_data,
                    pdf_template_part, signature_html):
    subject   = subject_tpl.format_map(row_data)
    body_html = (f'<html><body style="font-family:Arial,sans-serif;'
                 f'font-size:14px;color:#000;">\n'
                 f'{body_tpl.format_map(row_data)}\n<br><br>{signature_html}\n</body></html>')
    msg = MIMEMultipart()
    msg["From"] = email_from; msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html"))
    if pdf_template_part is not None:
        # MIMEBase parts become stateful once attached (headers get set on them),
        # so every message — and every thread — needs its OWN copy, never the
        # same object reused, or concurrent sends can corrupt each other's part.
        msg.attach(copy.deepcopy(pdf_template_part))
    return msg

# ── Core sending logic (sequential OR parallel via num_threads) ─────────
def send_emails(email, password, sheet_url, subject_tpl, body_tpl, pdf_path,
                delay, log_callback, done_callback, stop_event, pause_event,
                signature_html="", num_threads=0):
    """
    num_threads = 0 or 1  -> original single-connection sequential behaviour.
    num_threads >= 2      -> splits the recipient list into num_threads chunks,
                              each with its own SMTP connection, running in
                              parallel. Each thread still paces itself with the
                              same per-email `delay`, so total throughput scales
                              ~linearly with thread count.
    """
    try:
        log_callback("📡 Fetching contacts from Google Sheet...\n")
        df = read_google_sheet(sheet_url)
        total = len(df)
        log_callback(f"📋 {total} valid contacts loaded\n")

        pdf_template_part = None
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            pdf_name = os.path.basename(pdf_path)
            size_mb = len(pdf_bytes) / (1024 * 1024)
            log_callback(f"📎 PDF: {pdf_name} ({size_mb:.2f} MB) — Pre-encoding attachment...\n")

            part = MIMEBase("application", "octet-stream")
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{pdf_name}"')
            pdf_template_part = part
        else:
            log_callback("📭 No PDF attachment\n")

        logging.info(f"Session started — {total} contacts | threads={num_threads}")

        if num_threads and num_threads >= 2:
            log_callback(f"🧵 Parallel mode: {num_threads} threads\n\n")
            _send_emails_parallel(email, password, df, subject_tpl, body_tpl,
                                   pdf_template_part, delay, log_callback,
                                   done_callback, stop_event, pause_event,
                                   signature_html, num_threads, total)
        else:
            log_callback("\n")
            _send_emails_sequential(email, password, df, subject_tpl, body_tpl,
                                     pdf_template_part, delay, log_callback,
                                     done_callback, stop_event, pause_event,
                                     signature_html, total)
    except Exception as e:
        log_callback(f"\n❌ Error: {e}\n"); done_callback(False)


def _send_emails_sequential(email, password, df, subject_tpl, body_tpl,
                             pdf_template_part, delay, log_callback,
                             done_callback, stop_event, pause_event,
                             signature_html, total):
    sent = failed = 0
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(email, password)
        try:
            if hasattr(server, "sock") and server.sock:
                server.sock.setsockopt(socket.SOL_SOCKET, socket.SO_SNDBUF, 1024 * 1024)
                server.sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
        except Exception:
            pass
        log_callback("✅ Logged in to Gmail successfully\n\n")

        for i, row in df.iterrows():
            if stop_event.is_set():
                log_callback(f"\n⛔ Stopped. Sent:{sent} Failed:{failed}\n")
                done_callback(False); return
            while pause_event.is_set():
                log_callback("⏸  Paused...\n"); time.sleep(1)
                if stop_event.is_set(): done_callback(False); return
                break

            row_data = {col: str(val) for col, val in row.items()}
            to_email = str(row["Email"]).strip()
            name     = row_data.get("Name", to_email)

            try:
                msg = _build_message(email, to_email, subject_tpl, body_tpl,
                                      row_data, pdf_template_part, signature_html)
                server.sendmail(email, to_email, msg.as_string())
                sent += 1
                log_callback(f"[{sent+failed}/{total}] ✅ {name} <{to_email}>\n")
                logging.info(f"SUCCESS | {name} | {to_email}")
            except Exception as e:
                failed += 1
                log_callback(f"[{sent+failed}/{total}] ❌ {name}: {e}\n")

            if i < total - 1:
                for _ in range(delay):
                    if stop_event.is_set(): break
                    if pause_event.is_set():
                        while pause_event.is_set() and not stop_event.is_set():
                            time.sleep(0.5)
                        if not stop_event.is_set():
                            log_callback("▶️  Resumed!\n")
                    time.sleep(1)
                log_callback(f"    ⏳ Waited {delay}s\n")

    if not stop_event.is_set():
        log_callback(f"\n🎉 Done! Sent:{sent} Failed:{failed}\n")
        logging.info(f"Session complete. Sent:{sent}, Failed:{failed}")
        done_callback(True)


def _send_chunk(chunk_df, email, password, subject_tpl, body_tpl,
                 pdf_template_part, delay, log_callback, stop_event,
                 pause_event, signature_html, counters, counters_lock,
                 thread_tag):
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(email, password)
            try:
                if hasattr(server, "sock") and server.sock:
                    server.sock.setsockopt(socket.SOL_SOCKET, socket.SO_SNDBUF, 1024 * 1024)
                    server.sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            except Exception:
                pass

            rows = list(chunk_df.iterrows())
            for idx, (i, row) in enumerate(rows):
                if stop_event.is_set():
                    return
                while pause_event.is_set():
                    time.sleep(0.5)
                    if stop_event.is_set(): return

                row_data = {col: str(val) for col, val in row.items()}
                to_email = str(row["Email"]).strip()
                name     = row_data.get("Name", to_email)

                try:
                    msg = _build_message(email, to_email, subject_tpl, body_tpl,
                                          row_data, pdf_template_part, signature_html)
                    server.sendmail(email, to_email, msg.as_string())
                    with counters_lock:
                        counters["sent"] += 1
                        s, f, t = counters["sent"], counters["failed"], counters["total"]
                    log_callback(f"[{s+f}/{t}] ✅ {thread_tag} {name} <{to_email}>\n")
                    logging.info(f"SUCCESS | {name} | {to_email}")
                except Exception as e:
                    with counters_lock:
                        counters["failed"] += 1
                        s, f, t = counters["sent"], counters["failed"], counters["total"]
                    log_callback(f"[{s+f}/{t}] ❌ {thread_tag} {name}: {e}\n")

                if idx < len(rows) - 1:
                    for _ in range(delay):
                        if stop_event.is_set(): break
                        if pause_event.is_set():
                            while pause_event.is_set() and not stop_event.is_set():
                                time.sleep(0.5)
                        time.sleep(1)
    except Exception as e:
        log_callback(f"❌ {thread_tag} thread error: {e}\n")


def _send_emails_parallel(email, password, df, subject_tpl, body_tpl,
                           pdf_template_part, delay, log_callback,
                           done_callback, stop_event, pause_event,
                           signature_html, num_threads, total):
    # Never spin up more threads than rows to send.
    num_threads = max(1, min(num_threads, total)) if total else 1
    chunks = _split_dataframe(df, num_threads)

    counters = {"sent": 0, "failed": 0, "total": total}
    counters_lock = threading.Lock()
    log_lock = threading.Lock()

    def safe_log(msg):
        with log_lock:
            log_callback(msg)

    threads = []
    for idx, chunk in enumerate(chunks):
        if len(chunk) == 0:
            continue
        t = threading.Thread(
            target=_send_chunk,
            args=(chunk, email, password, subject_tpl, body_tpl,
                  pdf_template_part, delay, safe_log, stop_event, pause_event,
                  signature_html, counters, counters_lock, f"[T{idx+1}]"),
            daemon=True)
        threads.append(t)
        t.start()

    for t in threads:
        t.join()

    if stop_event.is_set():
        log_callback(f"\n⛔ Stopped. Sent:{counters['sent']} Failed:{counters['failed']}\n")
        done_callback(False)
    else:
        log_callback(f"\n🎉 Done! Sent:{counters['sent']} Failed:{counters['failed']}\n")
        logging.info(f"Session complete. Sent:{counters['sent']}, Failed:{counters['failed']}")
        done_callback(True)


# ── Manual follow-up sender (label-aware, reply-all) ────────────
# Kept sequential — IMAP label lookups happen on a single connection per
# recipient and threading IMAP reliably needs a connection per thread plus
# more careful state handling, so this tab is left as-is on purpose.
def send_manual_followups(gmail, password, sheet_url, label, subject_tpl, body_tpl,
                          delay, log_cb, done_cb, stop_event, pause_event,
                          resume_sent=None, signature_html="", draft_name="Follow-up"):
    try:
        log_cb("📡 Reading contacts from Google Sheet...\n")
        df = read_google_sheet(sheet_url)
        total = len(df)
        log_cb(f"📋 {total} contacts loaded\n")
        log_cb(f"📨 Using draft: {draft_name}\n\n")

        already_sent = resume_sent or []
        save_progress(total, already_sent)
        sent = failed = 0
        replied_skip = 0
        no_thread_skip = 0
        skipped = len(already_sent)
        if skipped > 0: log_cb(f"⏭  Skipping {skipped} already sent\n\n")

        log_cb("🔌 Connecting to Gmail IMAP...\n")
        imap = _imap_connect(gmail, password)
        log_cb(f"✅ IMAP connected\n")
        log_cb(f"🏷  Searching label: '{label}'\n\n")

        # Run an IMAP-dependent call against the current connection. Gmail
        # kills IMAP connections that get hit with too many rapid-fire
        # commands, which used to crash the whole session with a raw socket
        # error. Now: on a connection-level failure we reconnect once and
        # retry — a single recipient's lookup may fail, but the run
        # continues instead of aborting for everyone after it.
        def imap_call(fn, *args):
            nonlocal imap
            try:
                return fn(imap, *args)
            except (imaplib.IMAP4.abort, OSError, ssl.SSLError, EOFError) as e:
                log_cb(f"  ⚠️  IMAP connection dropped ({e}) — reconnecting...\n")
                imap = _imap_connect(gmail, password)
                return fn(imap, *args)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail, password)
            log_cb("✅ SMTP connected\n\n")

            for i, row in df.iterrows():
                if stop_event.is_set():
                    log_cb(f"\n⛔ Stopped. Sent:{sent} Failed:{failed} Replied(skipped):{replied_skip} NoThread(skipped):{no_thread_skip}\n")
                    done_cb(False); return
                while pause_event.is_set():
                    time.sleep(0.5)
                    if stop_event.is_set(): done_cb(False); return

                row_data = {col: str(val) for col, val in row.items()}
                to_email = str(row["Email"]).strip()
                name     = row_data.get("Name", to_email)
                company  = row_data.get("Company", "")

                if to_email in already_sent:
                    log_cb(f"  ⏭ {name} <{to_email}>\n"); continue

                # 'sent'/'failed' pace with the full user-set delay (that's
                # what protects against Gmail's SENDING rate limits). A pure
                # IMAP lookup skip (no thread found, or already replied) only
                # waits IMAP_SKIP_PACE (1s) before moving to the next email —
                # pacing it at the full send-delay made a mostly-skipped list
                # crawl for hours looking "stuck".
                outcome = None
                try:
                    result = imap_call(check_followup_status, to_email, label)
                    status = result["status"]

                    if status == "no_thread":
                        no_thread_skip += 1
                        outcome = "skip"
                        log_cb(f"  ⏭️  {name} <{to_email}> — not found in '{label}', skipping\n")

                    elif status == "replied":
                        replied_skip += 1
                        outcome = "skip"
                        log_cb(f"  💬 {name} <{to_email}> — already replied at least once, skipping (human-handled)\n")

                    else:  # status == "ready"
                        orig_subject = result["subject"]
                        orig_to      = result["to"]
                        orig_cc      = result["cc"]
                        msg_id       = result["message_id"]

                        subject = (orig_subject
                                   if orig_subject.lower().startswith("re:")
                                   else f"Re: {orig_subject}")

                        # Build Reply-All recipient list
                        all_to_addrs = [to_email]
                        if orig_to:
                            for addr in orig_to.split(","):
                                m = re.search(r"[\w.\-+]+@[\w.\-]+", addr.strip())
                                if m:
                                    extracted = m.group(0).lower()
                                    if extracted != gmail.lower() and extracted not in [a.lower() for a in all_to_addrs]:
                                        all_to_addrs.append(m.group(0))

                        cc_addrs = []
                        if orig_cc:
                            for addr in orig_cc.split(","):
                                m = re.search(r"[\w.\-+]+@[\w.\-]+", addr.strip())
                                if m:
                                    extracted = m.group(0).lower()
                                    if extracted != gmail.lower() and extracted not in [a.lower() for a in all_to_addrs]:
                                        cc_addrs.append(m.group(0))

                        row_data["Company"] = company
                        body_html = (f'<html><body style="font-family:Arial,sans-serif;'
                                     f'font-size:14px;color:#000;">\n'
                                     f'{body_tpl.format_map(row_data)}\n<br><br>{signature_html}\n</body></html>')

                        msg = MIMEMultipart()
                        msg["From"]    = gmail
                        msg["To"]      = ", ".join(all_to_addrs)
                        msg["Subject"] = subject
                        if cc_addrs:
                            msg["CC"] = ", ".join(cc_addrs)
                        if msg_id:
                            msg["In-Reply-To"] = msg_id
                            msg["References"]  = msg_id
                        msg.attach(MIMEText(body_html, "html"))

                        all_recipients = all_to_addrs + cc_addrs
                        smtp.sendmail(gmail, all_recipients, msg.as_string())

                        sent += 1
                        outcome = "sent"
                        already_sent.append(to_email)
                        save_progress(total, already_sent)
                        cc_note = f" (CC: {', '.join(cc_addrs)})" if cc_addrs else ""
                        log_cb(f"[{sent+failed}/{total-skipped}] ✅ {name} <{to_email}>{cc_note}\n")
                except Exception as e:
                    failed += 1
                    outcome = "failed"
                    log_cb(f"[{sent+failed}/{total-skipped}] ❌ {name}: {e}\n")

                # Sends and failures pace with the full user-set delay
                # (protects against Gmail's sending rate limit). Skips
                # (no-thread or already-replied) only wait IMAP_SKIP_PACE
                # (1s), just enough to avoid hammering the IMAP connection.
                pace = delay if outcome in ("sent", "failed") else min(delay, IMAP_SKIP_PACE)
                if i < total - 1:
                    for _ in range(pace):
                        if stop_event.is_set(): break
                        if pause_event.is_set():
                            while pause_event.is_set() and not stop_event.is_set():
                                time.sleep(0.5)
                            if not stop_event.is_set(): log_cb("▶️  Resumed!\n")
                        time.sleep(1)
                    if pace >= 1:
                        log_cb(f"    ⏳ Waited {pace}s\n")

        imap.logout()
        if not stop_event.is_set():
            clear_progress()
            log_cb(f"\n🎉 Done! Sent:{sent} Failed:{failed} Replied(skipped):{replied_skip} NoThread(skipped):{no_thread_skip}\n")
            done_cb(True)
    except Exception as e:
        log_cb(f"\n❌ Error: {e}\n"); done_cb(False)


# ── GUI ─────────────────────────────────────────────────────────
BG = "#0f0f0f"; CARD = "#1a1a1a"; ACCENT = "#4f9cf9"
SUCCESS = "#22c55e"; TEXT = "#e5e5e5"; MUTED = "#666"
BORDER = "#2a2a2a"; FONT = ("Consolas", 11)

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Email Automation — Tanish (with Multi Follow-up Drafts)")
        self.root.geometry("1920x1080")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        self.email_var    = tk.StringVar()
        self.password_var = tk.StringVar()
        self.sheet_url_var = tk.StringVar()
        self.subject_var  = tk.StringVar()
        self.pdf_path_var = tk.StringVar(value="No PDF selected (optional)")
        self.delay_var    = tk.IntVar(value=5)
        self.threads_var  = tk.IntVar(value=0)   # 0 = sequential (no parallel threading)
        self.sending      = False
        self.stop_event   = threading.Event()
        self.pause_event  = threading.Event()

        # Manual Follow-up general settings
        self.mfu_sheet_url_var = tk.StringVar()
        self.mfu_label_var     = tk.StringVar(value="Internship")
        self.mfu_delay_var     = tk.IntVar(value=30)
        self.mfu_sending       = False
        self.mfu_stop_event    = threading.Event()
        self.mfu_pause_event   = threading.Event()

        # Follow-up sub-tabs tracking list of dicts:
        # [{"title": str, "subject_var": StringVar, "body_box": ScrolledText, "frame": Frame}, ...]
        self.followup_tabs = []

        # Sent Status Checker
        self.check_sheet_url_var = tk.StringVar()
        self.check_sending = False
        self.check_stop_event = threading.Event()
        self.last_status_csv = None

        self._build_ui()
        self._load_saved_draft()

        # Auto-save all drafts on window close so nothing is ever lost
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        try:
            self._save_all_drafts(silent=True)
        except Exception as e:
            logging.warning(f"Error saving drafts on close: {e}")
        self.root.destroy()

    # ── Scrollable frame ────────────────────────────────────────
    def _scrollable(self, parent):
        canvas = tk.Canvas(parent, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        SF = tk.Frame(canvas, bg=BG)
        SF.bind("<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=SF, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        def _mw(e): canvas.yview_scroll(-1*(e.delta//120), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return SF, canvas

    # ── Build UI ────────────────────────────────────────────────
    def _build_ui(self):
        header = tk.Frame(self.root, bg="#111", pady=18)
        header.pack(fill="x")
        tk.Label(header, text="✉  EMAIL AUTOMATION",
                 font=("Consolas", 16, "bold"), fg=ACCENT, bg="#111").pack()
        tk.Label(header, text="IIT Kharagpur  ·  Mail Merge Tool with Multi Follow-up Drafts",
                 font=("Consolas", 9), fg=MUTED, bg="#111").pack()

        sty = ttk.Style(); sty.theme_use("default")
        # Main notebook style
        sty.configure("TNotebook", background=BG, borderwidth=0)
        sty.configure("TNotebook.Tab", background=CARD, foreground="#888",
                      font=("Consolas", 10, "bold"), padding=[20, 8])
        sty.map("TNotebook.Tab", background=[("selected", ACCENT)],
                foreground=[("selected", "white")])

        # Sub notebook style for Follow-up tabs
        sty.configure("Sub.TNotebook", background=CARD, borderwidth=0)
        sty.configure("Sub.TNotebook.Tab", background="#222222", foreground="#999999",
                      font=("Consolas", 9, "bold"), padding=[16, 7])
        sty.map("Sub.TNotebook.Tab",
                background=[("selected", "#f59e0b")],
                foreground=[("selected", "#000000")])

        nb = ttk.Notebook(self.root); nb.pack(fill="both", expand=True)
        self.notebook = nb

        t1 = tk.Frame(nb, bg=BG); nb.add(t1, text="  🚀 Send Emails  ")
        t2 = tk.Frame(nb, bg=BG); nb.add(t2, text="  🔧 Text → HTML  ")
        t3 = tk.Frame(nb, bg=BG); nb.add(t3, text="  📤 Manual Follow-up  ")
        t4 = tk.Frame(nb, bg=BG); nb.add(t4, text="  🤖 Automatic Follow-up  ")
        t5 = tk.Frame(nb, bg=BG); nb.add(t5, text="  📋 Check Sent Status  ")

        self._build_send_tab(t1)
        self._build_converter_tab(t2)
        self._build_manual_followup_tab(t3)
        self._build_auto_followup_tab(t4)
        self._build_check_sent_tab(t5)

    # ── Shared helpers ──────────────────────────────────────────
    def _card_label(self, parent, text):
        tk.Label(parent, text=text, font=("Consolas", 8, "bold"),
                 fg="#555", bg=BG).pack(anchor="w", padx=22, pady=(12, 2))

    def _field(self, parent, label, var, show=None):
        tk.Label(parent, text=label, font=("Consolas", 9),
                 fg=MUTED, bg=CARD).pack(anchor="w", pady=(4, 0))
        kw = dict(textvariable=var, font=FONT, bg="#111", fg=TEXT,
                  insertbackground=TEXT, relief="flat",
                  highlightthickness=1, highlightbackground=BORDER,
                  highlightcolor=ACCENT)
        if show: kw["show"] = show
        tk.Entry(parent, **kw).pack(fill="x", pady=(2, 8), ipady=6)

    def _log_append(self, box, msg):
        box.config(state="normal")
        box.insert("end", msg)
        box.see("end")
        box.config(state="disabled")

    def _make_log(self, parent, height=10):
        box = scrolledtext.ScrolledText(
            parent, font=("Consolas", 9), bg="#0a0a0a", fg="#a3e635",
            insertbackground="white", relief="flat",
            height=height, state="disabled")
        box.pack(fill="x")
        return box

    # ── TAB 1: Send Emails ──────────────────────────────────────
    def _build_send_tab(self, parent):
        SF, _ = self._scrollable(parent)

        self._card_label(SF, "GMAIL CREDENTIALS")
        c = tk.Frame(SF, bg=CARD, padx=20, pady=14); c.pack(fill="x", padx=20)
        self._field(c, "Gmail Address", self.email_var)
        self._field(c, "App Password",  self.password_var, show="●")

        self._card_label(SF, "GOOGLE SHEET URL")
        sf = tk.Frame(SF, bg=CARD, padx=20, pady=14); sf.pack(fill="x", padx=20)
        self._field(sf, "Paste Google Sheet tab URL", self.sheet_url_var)
        tk.Label(sf, text="⚠  Sheet must be 'Anyone with link can view'",
                 font=("Consolas", 8), fg="#f59e0b", bg=CARD).pack(anchor="w")

        self._card_label(SF, "PDF ATTACHMENT  (optional)")
        pf = tk.Frame(SF, bg=CARD, padx=20, pady=14); pf.pack(fill="x", padx=20)
        pr = tk.Frame(pf, bg=CARD); pr.pack(fill="x")
        self.pdf_label = tk.Label(pr, textvariable=self.pdf_path_var,
                                  font=("Consolas", 9), fg=MUTED, bg=CARD,
                                  anchor="w", width=55)
        self.pdf_label.pack(side="left")
        bc = tk.Frame(pr, bg=CARD); bc.pack(side="right")
        tk.Button(bc, text="Browse PDF", font=("Consolas", 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat", padx=12, pady=4,
                  cursor="hand2", command=self._browse_pdf).pack(side="left", padx=(0,6))
        tk.Button(bc, text="✕ Clear", font=("Consolas", 9),
                  bg="#3a1a1a", fg="#ef4444", relief="flat", padx=8, pady=4,
                  cursor="hand2", command=self._clear_pdf).pack(side="left")

        self._card_label(SF, "EMAIL COMPOSER")
        co = tk.Frame(SF, bg=CARD, padx=20, pady=14); co.pack(fill="x", padx=20)
        self._field(co, "Subject", self.subject_var)
        tk.Label(co, text="Email Body  (HTML — use Text→HTML tab to convert)",
                 font=("Consolas", 9), fg=MUTED, bg=CARD).pack(anchor="w", pady=(4,0))
        self.body_box = scrolledtext.ScrolledText(
            co, font=("Consolas", 10), bg="#111", fg=TEXT,
            insertbackground=TEXT, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT, height=12, wrap="word")
        self.body_box.pack(fill="x", pady=(2,8))
        tk.Button(co, text="💾  Save Draft", font=("Consolas", 9),
                  bg=BORDER, fg=MUTED, relief="flat", padx=10, pady=4,
                  cursor="hand2", command=lambda: self._save_all_drafts(silent=False)).pack(anchor="e")

        self._card_label(SF, "SEND DELAY (seconds, per-thread pacing)")
        df = tk.Frame(SF, bg=CARD, padx=20, pady=14); df.pack(fill="x", padx=20)
        dr = tk.Frame(df, bg=CARD); dr.pack(fill="x")
        tk.Scale(dr, from_=3, to=120, orient="horizontal", variable=self.delay_var,
                 bg=CARD, fg=TEXT, troughcolor=BORDER, highlightthickness=0,
                 font=("Consolas", 9), length=580,
                 activebackground=ACCENT).pack(side="left")
        tk.Label(dr, textvariable=self.delay_var,
                 font=("Consolas", 12, "bold"), fg=ACCENT, bg=CARD,
                 width=4).pack(side="left")
        tk.Label(dr, text="s", font=("Consolas", 10),
                 fg=MUTED, bg=CARD).pack(side="left")

        # ── Parallel threading slider ────────────────────────────
        self._card_label(SF, "PARALLEL THREADS  (0 = sequential, no parallel sending)")
        tf = tk.Frame(SF, bg=CARD, padx=20, pady=14); tf.pack(fill="x", padx=20)
        th_hint = tk.Frame(tf, bg="#1a1a2a", padx=10, pady=8)
        th_hint.pack(fill="x", pady=(0,10))
        tk.Label(th_hint,
                 text="💡 Splits recipients across N SMTP connections sent in parallel.",
                 font=("Consolas",8,"bold"), fg=ACCENT, bg="#1a1a2a").pack(anchor="w")
        tk.Label(th_hint,
                 text="   Each thread still waits the delay above between its own sends —\n"
                      "   more threads = higher total throughput, same per-email pacing.\n"
                      "   0 = original single-connection behavior.",
                 font=("Consolas",8), fg=MUTED, bg="#1a1a2a", justify="left").pack(anchor="w")
        tr = tk.Frame(tf, bg=CARD); tr.pack(fill="x")
        tk.Scale(tr, from_=0, to=10, orient="horizontal", variable=self.threads_var,
                 bg=CARD, fg=TEXT, troughcolor=BORDER, highlightthickness=0,
                 font=("Consolas", 9), length=580,
                 activebackground=ACCENT, command=self._on_threads_change).pack(side="left")
        self.threads_val_label = tk.Label(tr, text="0 (sequential)",
                 font=("Consolas", 12, "bold"), fg=ACCENT, bg=CARD, width=16, anchor="w")
        self.threads_val_label.pack(side="left")

        self._card_label(SF, "CONTROLS")
        cf = tk.Frame(SF, bg=BG, pady=16); cf.pack(fill="x", padx=20)
        self.send_btn = tk.Button(cf, text="🚀  START SENDING",
                                  font=("Consolas", 13, "bold"), bg=SUCCESS, fg="white",
                                  relief="flat", padx=20, pady=12, cursor="hand2",
                                  command=self._start_sending)
        self.send_btn.pack(fill="x", pady=(0,8))
        psr = tk.Frame(cf, bg=BG); psr.pack(fill="x")
        self.pause_btn = tk.Button(psr, text="⏸  PAUSE",
                                   font=("Consolas", 11, "bold"), bg="#f59e0b", fg="white",
                                   relief="flat", padx=20, pady=10, cursor="hand2",
                                   state="disabled", command=self._toggle_pause)
        self.pause_btn.pack(side="left", fill="x", expand=True, padx=(0,6))
        self.stop_btn = tk.Button(psr, text="⛔  STOP",
                                  font=("Consolas", 11, "bold"), bg="#ef4444", fg="white",
                                  relief="flat", padx=20, pady=10, cursor="hand2",
                                  state="disabled", command=self._stop_sending)
        self.stop_btn.pack(side="left", fill="x", expand=True)

        self._card_label(SF, "LIVE LOG")
        lf = tk.Frame(SF, bg=CARD, padx=20, pady=10)
        lf.pack(fill="x", padx=20, pady=(0,20))
        self.log_box = self._make_log(lf, height=10)

    def _on_threads_change(self, val):
        n = int(float(val))
        self.threads_val_label.config(
            text="0 (sequential)" if n == 0 else f"{n} threads (~{n}x throughput)")

    # ── TAB 2: Text → HTML ──────────────────────────────────────
    def _build_converter_tab(self, parent):
        SF, _ = self._scrollable(parent)

        info = tk.Frame(SF, bg="#1a2a1a", padx=16, pady=12)
        info.pack(fill="x", padx=20, pady=(16,0))
        tk.Label(info, text="🔧 Text → HTML Converter",
                 font=("Consolas", 12, "bold"), fg="#a3e635",
                 bg="#1a2a1a").pack(anchor="w")
        tk.Label(info, text="Type your email naturally — app converts to HTML.",
                 font=("Consolas", 8), fg=MUTED, bg="#1a2a1a").pack(anchor="w", pady=(4,0))

        self._card_label(SF, "FORMATTING GUIDE")
        guide = tk.Frame(SF, bg=CARD, padx=20, pady=14); guide.pack(fill="x", padx=20)
        for left, right in [
            ("Blank line","→ Paragraph break"), ("New line","→ Line break <br>"),
            ("**word**","→ Bold"),              ("*word*","→ Italic"),
            ("__word__","→ Underline"),         ("https://...","→ Clickable link"),
            ("{Name}","→ Placeholder (kept)")
        ]:
            r = tk.Frame(guide, bg=CARD); r.pack(fill="x", pady=1)
            tk.Label(r, text=left, font=("Consolas",9,"bold"), fg="#a3e635",
                     bg=CARD, width=20, anchor="w").pack(side="left")
            tk.Label(r, text=right, font=("Consolas",9), fg=MUTED,
                     bg=CARD, anchor="w").pack(side="left")

        pane = tk.Frame(SF, bg=BG); pane.pack(fill="x", padx=20, pady=(8,0))

        lf = tk.Frame(pane, bg=CARD, padx=12, pady=12)
        lf.pack(side="left", fill="both", expand=True, padx=(0,8))
        tk.Label(lf, text="✏️  PLAIN TEXT INPUT",
                 font=("Consolas",9,"bold"), fg=ACCENT,
                 bg=CARD).pack(anchor="w", pady=(0,6))
        self.plain_box = scrolledtext.ScrolledText(
            lf, font=("Consolas",10), bg="#111", fg=TEXT,
            insertbackground=TEXT, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT, height=22, wrap="word")
        self.plain_box.pack(fill="both", expand=True)

        rf = tk.Frame(pane, bg=CARD, padx=12, pady=12)
        rf.pack(side="left", fill="both", expand=True)
        tk.Label(rf, text="⚙️  HTML OUTPUT",
                 font=("Consolas",9,"bold"), fg="#f59e0b",
                 bg=CARD).pack(anchor="w", pady=(0,6))
        self.html_box = scrolledtext.ScrolledText(
            rf, font=("Consolas",10), bg="#111", fg="#f59e0b",
            insertbackground=TEXT, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor="#f59e0b", height=22, wrap="word", state="disabled")
        self.html_box.pack(fill="both", expand=True)

        bf = tk.Frame(SF, bg=BG, pady=12); bf.pack(fill="x", padx=20)
        for txt, col, cmd in [
            ("⚡  CONVERT",                  "#a3e635", self._convert_to_html),
            ("📋  COPY HTML",                ACCENT,    self._copy_html),
            ("📤  SEND TO COMPOSER",         SUCCESS,   self._send_to_composer),
            ("📤  SEND TO ACTIVE FOLLOW-UP", "#f59e0b", self._send_to_active_followup),
        ]:
            tk.Button(bf, text=txt, font=("Consolas",10,"bold"), bg=col,
                      fg="#000" if col in ("#a3e635", "#f59e0b") else "white",
                      relief="flat", padx=14, pady=10, cursor="hand2",
                      command=cmd).pack(side="left", padx=(0,8))
        tk.Button(bf, text="🗑  CLEAR", font=("Consolas",10),
                  bg="#3a1a1a", fg="#ef4444", relief="flat", padx=14, pady=10,
                  cursor="hand2", command=self._clear_converter).pack(side="right")

        self.conv_status = tk.Label(SF, text="", font=("Consolas",9),
                                    fg=SUCCESS, bg=BG)
        self.conv_status.pack(anchor="w", padx=22, pady=(4,16))

    # ── TAB 3: Manual Follow-up ─────────────────────────────────
    def _build_manual_followup_tab(self, parent):
        SF, _ = self._scrollable(parent)

        info = tk.Frame(SF, bg="#2a1a0a", padx=16, pady=12)
        info.pack(fill="x", padx=20, pady=(16,0))
        tk.Label(info, text="📤 Manual Follow-up Sender",
                 font=("Consolas",12,"bold"), fg="#f59e0b",
                 bg="#2a1a0a").pack(anchor="w")
        tk.Label(info, text="Per recipient: check the label for an original thread → check if they've ever replied → send Reply-All",
                 font=("Consolas",8), fg=MUTED, bg="#2a1a0a").pack(anchor="w", pady=(4,0))
        tk.Label(info, text="✅ Progress auto-saved — resume if connection drops",
                 font=("Consolas",8), fg=SUCCESS, bg="#2a1a0a").pack(anchor="w", pady=(2,0))
        tk.Label(info, text="⏭️  Not found in the label → skipped, 1s wait, next recipient (never sent as new mail)",
                 font=("Consolas",8), fg=ACCENT, bg="#2a1a0a").pack(anchor="w", pady=(2,0))
        tk.Label(info, text="💬 Replied even once, ever → skipped permanently, 1s wait, next recipient (thread is human-handled)",
                 font=("Consolas",8), fg=ACCENT, bg="#2a1a0a").pack(anchor="w", pady=(2,0))

        self._card_label(SF, "USES CREDENTIALS FROM SEND TAB")
        n = tk.Frame(SF, bg=CARD, padx=20, pady=10); n.pack(fill="x", padx=20)
        tk.Label(n, text="⚠  Enable IMAP: Gmail → Settings → Forwarding and POP/IMAP",
                 font=("Consolas",8), fg="#f59e0b", bg=CARD).pack(anchor="w")

        # ── Gmail Label ─────────────────────────────────────────
        self._card_label(SF, "GMAIL LABEL TO SEARCH")
        la = tk.Frame(SF, bg=CARD, padx=20, pady=14); la.pack(fill="x", padx=20)

        label_hint = tk.Frame(la, bg="#1a1a2a", padx=10, pady=8)
        label_hint.pack(fill="x", pady=(0,10))
        tk.Label(label_hint,
                 text="💡 Enter the Gmail label where the original emails were sent from",
                 font=("Consolas",8,"bold"), fg=ACCENT, bg="#1a1a2a").pack(anchor="w")
        tk.Label(label_hint,
                 text='   e.g. "Internship" or "Job Apps"  —  must match Gmail label exactly (case-sensitive)',
                 font=("Consolas",8), fg=MUTED, bg="#1a1a2a").pack(anchor="w")

        self._field(la, "Gmail Label Name", self.mfu_label_var)

        # ── Google Sheet ────────────────────────────────────────
        self._card_label(SF, "GOOGLE SHEET URL  (Name, Email, Company columns required)")
        su = tk.Frame(SF, bg=CARD, padx=20, pady=14); su.pack(fill="x", padx=20)
        self._field(su, "Paste Google Sheet tab URL", self.mfu_sheet_url_var)
        tk.Label(su, text="⚠  Sheet must be 'Anyone with link can view'",
                 font=("Consolas",8), fg="#f59e0b", bg=CARD).pack(anchor="w")

        # ── SUB-TABS COMPOSER: Follow-up 1, Follow-up 2, ... ────
        self._card_label(SF, "FOLLOW-UP DRAFTS & SUB-TABS (Follow-up 1 to 5+)")
        co_outer = tk.Frame(SF, bg=CARD, padx=16, pady=14); co_outer.pack(fill="x", padx=20)

        hint = tk.Frame(co_outer, bg="#1e2a1e", padx=10, pady=8)
        hint.pack(fill="x", pady=(0,10))
        tk.Label(hint,
                 text="💡 Each sub-tab holds its own draft. The currently selected sub-tab is sent when you click 'Send Follow-ups'.",
                 font=("Consolas",8,"bold"), fg="#a3e635",
                 bg="#1e2a1e").pack(anchor="w")
        tk.Label(hint,
                 text="   Use {Name}, {Company} placeholders. Subject is auto-set as 'Re: <original>' (Fallback Subject used only if not found).",
                 font=("Consolas",8), fg=MUTED, bg="#1e2a1e").pack(anchor="w")

        # Toolbar above sub-tabs notebook
        subtab_tb = tk.Frame(co_outer, bg=CARD)
        subtab_tb.pack(fill="x", pady=(0, 8))

        self.active_fu_label = tk.Label(
            subtab_tb, text="📌 Active Draft: Follow-up 1",
            font=("Consolas", 10, "bold"), fg="#f59e0b", bg=CARD
        )
        self.active_fu_label.pack(side="left", anchor="w")

        btn_box = tk.Frame(subtab_tb, bg=CARD)
        btn_box.pack(side="right")

        tk.Button(btn_box, text="➕ Add Follow-up", font=("Consolas", 8, "bold"),
                  bg="#2563eb", fg="white", relief="flat", padx=10, pady=3,
                  cursor="hand2", command=self._add_followup_tab).pack(side="left", padx=(0, 6))

        tk.Button(btn_box, text="🗑 Remove Active Tab", font=("Consolas", 8),
                  bg="#3a1a1a", fg="#ef4444", relief="flat", padx=10, pady=3,
                  cursor="hand2", command=self._remove_active_followup_tab).pack(side="left", padx=(0, 6))

        tk.Button(btn_box, text="💾 Save All Drafts", font=("Consolas", 8, "bold"),
                  bg="#15803d", fg="white", relief="flat", padx=10, pady=3,
                  cursor="hand2", command=lambda: self._save_all_drafts(silent=False)).pack(side="left")

        # Nested Notebook for Follow-up 1, Follow-up 2, ...
        self.fu_notebook = ttk.Notebook(co_outer, style="Sub.TNotebook")
        self.fu_notebook.pack(fill="x", expand=True, pady=(2, 6))
        self.fu_notebook.bind("<<NotebookTabChanged>>", self._on_fu_tab_changed)

        # ── Delay ───────────────────────────────────────────────
        self._card_label(SF, "SEND DELAY (seconds — applies to actual sends only; skips wait 1s)")
        df = tk.Frame(SF, bg=CARD, padx=20, pady=14); df.pack(fill="x", padx=20)
        dr = tk.Frame(df, bg=CARD); dr.pack(fill="x")
        tk.Scale(dr, from_=3, to=120, orient="horizontal",
                 variable=self.mfu_delay_var,
                 bg=CARD, fg=TEXT, troughcolor=BORDER, highlightthickness=0,
                 font=("Consolas",9), length=580,
                 activebackground=ACCENT).pack(side="left")
        tk.Label(dr, textvariable=self.mfu_delay_var,
                 font=("Consolas",12,"bold"), fg=ACCENT, bg=CARD,
                 width=4).pack(side="left")
        tk.Label(dr, text="s", font=("Consolas",10),
                 fg=MUTED, bg=CARD).pack(side="left")

        # ── Controls ────────────────────────────────────────────
        self._card_label(SF, "CONTROLS")
        cf = tk.Frame(SF, bg=BG, pady=16); cf.pack(fill="x", padx=20)

        self.mfu_active_indicator = tk.Label(
            cf, text="🎯 Draft Selected to Send: Follow-up 1",
            font=("Consolas", 11, "bold"), fg="#f59e0b", bg=BG
        )
        self.mfu_active_indicator.pack(anchor="w", pady=(0, 8))

        self.mfu_send_btn = tk.Button(cf, text="📤  SEND FOLLOW-UPS",
                                      font=("Consolas",13,"bold"), bg="#f59e0b", fg="white",
                                      relief="flat", padx=20, pady=12, cursor="hand2",
                                      command=self._start_manual_followup)
        self.mfu_send_btn.pack(fill="x", pady=(0,8))
        self.mfu_resume_btn = tk.Button(cf, text="🔄  RESUME INTERRUPTED SESSION",
                                        font=("Consolas",11,"bold"), bg=ACCENT, fg="white",
                                        relief="flat", padx=20, pady=10, cursor="hand2",
                                        command=self._resume_manual_followup)
        self.mfu_resume_btn.pack(fill="x", pady=(0,8))
        psr = tk.Frame(cf, bg=BG); psr.pack(fill="x")
        self.mfu_pause_btn = tk.Button(psr, text="⏸  PAUSE",
                                       font=("Consolas",11,"bold"), bg="#f59e0b", fg="white",
                                       relief="flat", padx=20, pady=10, cursor="hand2",
                                       state="disabled", command=self._toggle_mfu_pause)
        self.mfu_pause_btn.pack(side="left", fill="x", expand=True, padx=(0,6))
        self.mfu_stop_btn = tk.Button(psr, text="⛔  STOP",
                                      font=("Consolas",11,"bold"), bg="#ef4444", fg="white",
                                      relief="flat", padx=20, pady=10, cursor="hand2",
                                      state="disabled", command=self._stop_manual_followup)
        self.mfu_stop_btn.pack(side="left", fill="x", expand=True)
        self.mfu_progress_label = tk.Label(cf, text="", font=("Consolas",9),
                                           fg=SUCCESS, bg=BG)
        self.mfu_progress_label.pack(anchor="w", pady=(8,0))

        self._card_label(SF, "LIVE LOG")
        lf = tk.Frame(SF, bg=CARD, padx=20, pady=10)
        lf.pack(fill="x", padx=20, pady=(0,20))
        self.mfu_log_box = self._make_log(lf, height=12)

    # ── Follow-up Sub-Tabs Management ───────────────────────────
    def _create_followup_subtab(self, title, subject="", body=""):
        """Create a single sub-tab inside the Follow-up notebook."""
        tab_frame = tk.Frame(self.fu_notebook, bg=CARD, padx=12, pady=10)

        # Fallback Subject
        tk.Label(tab_frame, text=f"Fallback Subject for '{title}' (used only if original email not found in label):",
                 font=("Consolas", 8, "bold"), fg=MUTED, bg=CARD).pack(anchor="w", pady=(2, 2))

        subj_var = tk.StringVar(value=subject or f"Following up on my previous email")
        subj_entry = tk.Entry(tab_frame, textvariable=subj_var, font=FONT,
                              bg="#111", fg=TEXT, insertbackground=TEXT, relief="flat",
                              highlightthickness=1, highlightbackground=BORDER,
                              highlightcolor="#f59e0b")
        subj_entry.pack(fill="x", pady=(0, 8), ipady=5)

        # Body Box
        tk.Label(tab_frame, text=f"Email Body (HTML supported) — {title}:",
                 font=("Consolas", 8, "bold"), fg=MUTED, bg=CARD).pack(anchor="w", pady=(2, 2))

        body_box = scrolledtext.ScrolledText(
            tab_frame, font=("Consolas", 10), bg="#111", fg=TEXT,
            insertbackground=TEXT, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor="#f59e0b", height=13, wrap="word")
        body_box.pack(fill="x", pady=(0, 8))
        if body:
            body_box.insert("1.0", body)

        # Snippet buttons & Action buttons
        br = tk.Frame(tab_frame, bg=CARD); br.pack(fill="x", pady=(0, 4))
        tk.Label(br, text="Quick Insert:", font=("Consolas", 8),
                 fg=MUTED, bg=CARD).pack(side="left", padx=(0, 6))

        for lbl, snip in [("Bold", "<strong></strong>"),
                          ("Link", '<a href="URL">Text</a>'),
                          ("BR", "<br>"), ("P", "<p></p>")]:
            tk.Button(br, text=lbl, font=("Consolas", 8), bg=BORDER, fg=TEXT,
                      relief="flat", padx=8, pady=2, cursor="hand2",
                      command=lambda s=snip, b=body_box: b.insert(tk.INSERT, s)
                      ).pack(side="left", padx=2)

        tk.Button(br, text="💾 Save This Draft", font=("Consolas", 8, "bold"),
                  bg="#15803d", fg="white", relief="flat", padx=10, pady=2,
                  cursor="hand2", command=lambda: self._save_all_drafts(silent=False)
                  ).pack(side="right", padx=(4, 0))

        tk.Button(br, text="🗑 Clear Draft", font=("Consolas", 8),
                  bg="#3a1a1a", fg="#ef4444", relief="flat", padx=8, pady=2,
                  cursor="hand2", command=lambda b=body_box: b.delete("1.0", "end")
                  ).pack(side="right", padx=2)

        self.fu_notebook.add(tab_frame, text=f"  {title}  ")

        tab_data = {
            "title": title,
            "subject_var": subj_var,
            "body_box": body_box,
            "frame": tab_frame
        }
        self.followup_tabs.append(tab_data)
        return tab_data

    def _on_fu_tab_changed(self, event=None):
        """Update indicators when user switches sub-tabs."""
        try:
            current_idx = self._get_active_followup_index()
            if 0 <= current_idx < len(self.followup_tabs):
                active_tab = self.followup_tabs[current_idx]
                title = active_tab["title"]
                self.active_fu_label.config(text=f"📌 Active Draft: {title}")
                self.mfu_active_indicator.config(text=f"🎯 Draft Selected to Send: {title}")
        except Exception:
            pass

    def _get_active_followup_index(self):
        """Return the currently selected sub-tab index in fu_notebook."""
        try:
            sel = self.fu_notebook.select()
            if sel:
                return self.fu_notebook.index(sel)
        except Exception:
            pass
        return 0

    def _get_active_followup_tab(self):
        """Return dict for the currently selected sub-tab."""
        idx = self._get_active_followup_index()
        if 0 <= idx < len(self.followup_tabs):
            return self.followup_tabs[idx]
        if self.followup_tabs:
            return self.followup_tabs[0]
        return None

    # Properties for backwards compatibility
    @property
    def mfu_body_box(self):
        t = self._get_active_followup_tab()
        return t["body_box"] if t else None

    @property
    def mfu_subject_var(self):
        t = self._get_active_followup_tab()
        return t["subject_var"] if t else None

    def _add_followup_tab(self):
        """Add a new follow-up sub-tab (Follow-up N)."""
        next_num = len(self.followup_tabs) + 1
        new_title = f"Follow-up {next_num}"
        self._create_followup_subtab(
            new_title,
            subject=f"Following up on our previous conversation ({new_title})",
            body=""
        )
        # Select newly created tab
        self.fu_notebook.select(len(self.followup_tabs) - 1)
        self._on_fu_tab_changed()
        self._save_all_drafts(silent=True)

    def _remove_active_followup_tab(self):
        """Remove the active follow-up sub-tab, leaving at least 1."""
        if len(self.followup_tabs) <= 1:
            messagebox.showwarning("Cannot Remove", "You must have at least one Follow-up tab.")
            return

        active_idx = self._get_active_followup_index()
        tab_to_remove = self.followup_tabs[active_idx]
        title = tab_to_remove["title"]

        if not messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove '{title}'?"):
            return

        self.fu_notebook.forget(tab_to_remove["frame"])
        self.followup_tabs.pop(active_idx)

        # Select previous or 0
        new_idx = max(0, min(active_idx, len(self.followup_tabs) - 1))
        self.fu_notebook.select(new_idx)
        self._on_fu_tab_changed()
        self._save_all_drafts(silent=True)

    # ── TAB 4: Automatic Follow-up (placeholder — not implemented yet) ──
    def _build_auto_followup_tab(self, parent):
        SF, _ = self._scrollable(parent)

        info = tk.Frame(SF, bg="#1a1a2a", padx=16, pady=12)
        info.pack(fill="x", padx=20, pady=(16,0))
        tk.Label(info, text="🤖 Automatic Follow-up",
                 font=("Consolas",12,"bold"), fg=ACCENT,
                 bg="#1a1a2a").pack(anchor="w")
        tk.Label(info, text="Coming soon — this tab is a placeholder for now.",
                 font=("Consolas",9), fg=MUTED, bg="#1a1a2a").pack(anchor="w", pady=(4,0))
        tk.Label(info,
                 text="The idea: run the same skip logic as Manual Follow-up (no thread → skip,\n"
                      "already replied → skip) automatically on a schedule (e.g. every N days after\n"
                      "the original send), instead of you clicking 'Send Follow-ups' by hand.",
                 font=("Consolas",8), fg=MUTED, bg="#1a1a2a", justify="left").pack(anchor="w", pady=(6,0))

        placeholder = tk.Frame(SF, bg=CARD, padx=20, pady=40); placeholder.pack(fill="x", padx=20, pady=20)
        tk.Label(placeholder, text="🚧  Not built yet.",
                 font=("Consolas",14,"bold"), fg="#555", bg=CARD).pack()
        tk.Label(placeholder, text="Use the Manual Follow-up tab for now.",
                 font=("Consolas",9), fg=MUTED, bg=CARD).pack(pady=(6,0))

    # ── TAB 5: Check Sent Status ─────────────────────────────────
    def _build_check_sent_tab(self, parent):
        SF, _ = self._scrollable(parent)

        info = tk.Frame(SF, bg="#1a2a1a", padx=16, pady=12)
        info.pack(fill="x", padx=20, pady=(16,0))
        tk.Label(info, text="📋 Check Sent Status",
                 font=("Consolas",12,"bold"), fg=SUCCESS, bg="#1a2a1a").pack(anchor="w")
        tk.Label(info,
                 text="Checks Gmail Sent Mail for each recipient and exports a CSV — your Google Sheet is never modified.",
                 font=("Consolas",8), fg=MUTED, bg="#1a2a1a").pack(anchor="w", pady=(4,0))
        tk.Label(info,
                 text="🔎 Matching is based on the recipient email address in Gmail Sent Mail.",
                 font=("Consolas",8), fg=MUTED, bg="#1a2a1a").pack(anchor="w", pady=(2,0))
        tk.Label(info,
                 text="📄 All original rows are preserved and a Status column is added to the CSV.",
                 font=("Consolas",8), fg=ACCENT, bg="#1a2a1a").pack(anchor="w", pady=(2,0))

        self._card_label(SF, "USES CREDENTIALS FROM SEND TAB")
        cr = tk.Frame(SF, bg=CARD, padx=20, pady=12); cr.pack(fill="x", padx=20)
        tk.Label(cr,
                 text="Gmail Address and App Password are taken automatically from the Send Emails tab.",
                 font=("Consolas",9), fg=TEXT, bg=CARD).pack(anchor="w")

        self._card_label(SF, "GOOGLE SHEET URL")
        su = tk.Frame(SF, bg=CARD, padx=20, pady=14); su.pack(fill="x", padx=20)
        self._field(su, "Paste Google Sheet tab URL", self.check_sheet_url_var)
        tk.Label(su,
                 text="⚠ Sheet must be 'Anyone with link can view' and contain Name, Company, Email columns.",
                 font=("Consolas",8), fg="#f59e0b", bg=CARD).pack(anchor="w")

        self._card_label(SF, "CONTROLS")
        cf = tk.Frame(SF, bg=BG, pady=16); cf.pack(fill="x", padx=20)
        self.check_btn = tk.Button(cf, text="🔎  CHECK SENT STATUS",
                                   font=("Consolas",13,"bold"), bg=SUCCESS, fg="white",
                                   relief="flat", padx=20, pady=12, cursor="hand2",
                                   command=self._start_check_sent)
        self.check_btn.pack(fill="x", pady=(0,8))
        self.check_stop_btn = tk.Button(cf, text="⛔  STOP",
                                        font=("Consolas",11,"bold"), bg="#ef4444", fg="white",
                                        relief="flat", padx=20, pady=10, cursor="hand2",
                                        state="disabled", command=self._stop_check_sent)
        self.check_stop_btn.pack(fill="x", pady=(0,8))
        self.check_progress_label = tk.Label(cf, text="", font=("Consolas",9),
                                             fg=SUCCESS, bg=BG, justify="left")
        self.check_progress_label.pack(anchor="w", pady=(4,0))

        self._card_label(SF, "LIVE LOG")
        lf = tk.Frame(SF, bg=CARD, padx=20, pady=10)
        lf.pack(fill="x", padx=20, pady=(0,20))
        self.check_log_box = self._make_log(lf, height=18)

    def _start_check_sent(self):
        if self.check_sending:
            return

        email = self.email_var.get().strip()
        pw = self.password_var.get().strip()
        url = self.check_sheet_url_var.get().strip()

        if not email or not pw:
            messagebox.showerror(
                "Missing",
                "Enter Gmail Address and App Password in the Send Emails tab."
            )
            return
        if not url or "docs.google.com/spreadsheets" not in url:
            messagebox.showerror("Invalid", "Paste a valid Google Sheets URL.")
            return

        self.check_stop_event.clear()
        self.check_sending = True
        self.last_status_csv = None
        self.check_btn.config(text="⏳  CHECKING...", state="disabled", bg="#555")
        self.check_stop_btn.config(state="normal")
        self.check_progress_label.config(text="")
        self._log_append(
            self.check_log_box,
            f"\n{'─'*60}\n▶ {datetime.now().strftime('%H:%M:%S')}\n"
        )

        threading.Thread(
            target=check_sent_status,
            args=(url, email, pw,
                  lambda m: self._log_append(self.check_log_box, m),
                  self._check_sent_done,
                  self.check_stop_event),
            daemon=True
        ).start()

    def _stop_check_sent(self):
        self.check_stop_event.set()
        self.check_stop_btn.config(state="disabled")
        self._log_append(
            self.check_log_box,
            "⛔ Stop requested — finishing current Gmail lookup...\n"
        )

    def _check_sent_done(self, ok, output_path):
        def update_ui():
            self.check_sending = False
            self.check_btn.config(
                text="🔎  CHECK SENT STATUS",
                state="normal",
                bg=SUCCESS if ok else "#ef4444"
            )
            self.check_stop_btn.config(state="disabled")
            if ok and output_path:
                self.last_status_csv = output_path
                self.check_progress_label.config(text=f"✅ CSV ready:\n{output_path}")
                messagebox.showinfo(
                    "Complete",
                    f"Sent-status check completed.\n\nCSV saved to:\n{output_path}"
                )
            else:
                self.check_progress_label.config(
                    text="⚠️ Check stopped or failed. See the log above."
                )
            self.root.after(3000, lambda: self.check_btn.config(bg=SUCCESS))
        self.root.after(0, update_ui)

    # ── Send tab actions ────────────────────────────────────────
    def _browse_pdf(self):
        p = filedialog.askopenfilename(title="Select PDF",
                                       filetypes=[("PDF","*.pdf")])
        if p: self.pdf_path_var.set(p); self.pdf_label.config(fg=TEXT)

    def _clear_pdf(self):
        self.pdf_path_var.set("No PDF selected (optional)")
        self.pdf_label.config(fg=MUTED)

    def _load_saved_draft(self):
        """Load draft.json, restoring Tab 1 and all Follow-up sub-tabs (1 to 5+)."""
        d = load_draft()

        # Send Tab draft
        self.subject_var.set(d.get("subject", ""))
        self.sheet_url_var.set(d.get("sheet_url", ""))
        pdf = d.get("pdf_path", "")
        if pdf and os.path.exists(pdf):
            self.pdf_path_var.set(pdf)
            self.pdf_label.config(fg=TEXT)
        if d.get("body", ""):
            self.body_box.delete("1.0", "end")
            self.body_box.insert("1.0", d["body"])
        saved_threads = int(d.get("threads", 0) or 0)
        self.threads_var.set(saved_threads)
        self._on_threads_change(saved_threads)

        # Manual follow-up general settings
        if d.get("mfu_sheet_url"):
            self.mfu_sheet_url_var.set(d["mfu_sheet_url"])
        if d.get("mfu_label"):
            self.mfu_label_var.set(d["mfu_label"])
        if d.get("mfu_delay"):
            try:
                self.mfu_delay_var.set(int(d["mfu_delay"]))
            except Exception:
                pass

        # Load Follow-up sub-tabs (guaranteeing at least 5 tabs)
        saved_fus = d.get("mfu_followups", [])
        legacy_fu_body = d.get("fu_body", "")

        # If empty or not a list, create default Follow-up 1 to 5
        if not saved_fus or not isinstance(saved_fus, list):
            saved_fus = [
                {"title": "Follow-up 1", "subject": "Following up on my previous email", "body": legacy_fu_body or ""},
                {"title": "Follow-up 2", "subject": "Second follow-up regarding our proposal", "body": ""},
                {"title": "Follow-up 3", "subject": "Third follow-up regarding partnership", "body": ""},
                {"title": "Follow-up 4", "subject": "Checking in regarding our collaboration", "body": ""},
                {"title": "Follow-up 5", "subject": "Final follow-up regarding Kshitij IIT Kharagpur", "body": ""},
            ]
        else:
            # Ensure at least 5 follow-up tabs exist
            existing_count = len(saved_fus)
            if existing_count < 5:
                for i in range(existing_count + 1, 6):
                    saved_fus.append({
                        "title": f"Follow-up {i}",
                        "subject": f"Follow-up {i} regarding our previous email",
                        "body": ""
                    })

        # Clear existing tabs from notebook if any
        for tab in self.followup_tabs:
            try:
                self.fu_notebook.forget(tab["frame"])
            except Exception:
                pass
        self.followup_tabs.clear()

        # Build each tab
        for item in saved_fus:
            self._create_followup_subtab(
                title=item.get("title", f"Follow-up {len(self.followup_tabs)+1}"),
                subject=item.get("subject", ""),
                body=item.get("body", "")
            )

        # Restore active sub-tab index
        active_idx = int(d.get("mfu_active_index", 0) or 0)
        if 0 <= active_idx < len(self.followup_tabs):
            self.fu_notebook.select(active_idx)
        elif self.followup_tabs:
            self.fu_notebook.select(0)
        self._on_fu_tab_changed()

    def _save_all_drafts(self, silent=True):
        """Save Send Tab draft and all Follow-up sub-tabs drafts to draft.json."""
        fus_data = []
        for tab in self.followup_tabs:
            fus_data.append({
                "title": tab["title"],
                "subject": tab["subject_var"].get(),
                "body": tab["body_box"].get("1.0", "end-1c")
            })

        active_idx = self._get_active_followup_index()
        active_body = ""
        if 0 <= active_idx < len(fus_data):
            active_body = fus_data[active_idx]["body"]

        save_draft(
            subject=self.subject_var.get(),
            body=self.body_box.get("1.0", "end-1c"),
            sheet_url=self.sheet_url_var.get(),
            pdf_path=self.pdf_path_var.get(),
            fu_body=active_body,
            threads=self.threads_var.get(),
            mfu_sheet_url=self.mfu_sheet_url_var.get(),
            mfu_label=self.mfu_label_var.get(),
            mfu_delay=self.mfu_delay_var.get(),
            mfu_active_index=active_idx,
            mfu_followups=fus_data
        )
        if not silent:
            messagebox.showinfo("Saved", "✅ All drafts (Send tab & all Follow-up sub-tabs) saved successfully!")

    def _save_draft_manual(self):
        self._save_all_drafts(silent=False)

    def _toggle_pause(self):
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.pause_btn.config(text="⏸  PAUSE", bg="#f59e0b")
            self._log_append(self.log_box, "▶️  Resumed!\n")
        else:
            self.pause_event.set()
            self.pause_btn.config(text="▶️  RESUME", bg=SUCCESS)
            self._log_append(self.log_box, "⏸  Paused\n")

    def _stop_sending(self):
        self.stop_event.set(); self.pause_event.clear()
        self.pause_btn.config(state="disabled", text="⏸  PAUSE", bg="#f59e0b")
        self.stop_btn.config(state="disabled")

    def _start_sending(self):
        if self.sending: return
        email = self.email_var.get().strip()
        pw    = self.password_var.get().strip()
        url   = self.sheet_url_var.get().strip()
        subj  = self.subject_var.get().strip()
        body  = self.body_box.get("1.0","end-1c").strip()
        pdf   = self.pdf_path_var.get().strip()
        if pdf == "No PDF selected (optional)": pdf = ""
        if not all([email,pw,url,subj,body]):
            messagebox.showerror("Missing","Fill all required fields."); return
        num_threads = self.threads_var.get()
        self._save_all_drafts(silent=True)
        self.stop_event.clear(); self.pause_event.clear()
        self.sending = True
        mode_txt = "⏳  SENDING..." if num_threads < 2 else f"⏳  SENDING ({num_threads} threads)..."
        self.send_btn.config(text=mode_txt, state="disabled", bg="#555")
        self.pause_btn.config(state="normal"); self.stop_btn.config(state="normal")
        self._log_append(self.log_box,
            f"\n{'─'*50}\n▶ {datetime.now().strftime('%H:%M:%S')}\n")

        # Fetch the user's real Gmail signature via OAuth before sending.
        try:
            self._log_append(self.log_box, "🔏 Fetching Gmail signature...\n")
            sig = get_gmail_signature(email)
        except Exception as e:
            self._log_append(self.log_box, f"⚠️  Could not fetch signature: {e}\n")
            sig = ""

        threading.Thread(target=send_emails,
            args=(email, pw, url, subj, body, pdf, self.delay_var.get(),
                  lambda m: self._log_append(self.log_box, m),
                  self._done, self.stop_event, self.pause_event, sig, num_threads),
            daemon=True).start()

    def _done(self, ok):
        self.sending = False
        self.send_btn.config(text="🚀  START SENDING", state="normal",
                             bg=SUCCESS if ok else "#ef4444")
        self.pause_btn.config(state="disabled", text="⏸  PAUSE", bg="#f59e0b")
        self.stop_btn.config(state="disabled")
        self.root.after(3000, lambda: self.send_btn.config(bg=SUCCESS))

    # ── Converter actions ───────────────────────────────────────
    def _convert_to_html(self):
        plain = self.plain_box.get("1.0","end-1c")
        if not plain.strip():
            messagebox.showwarning("Empty","Enter text first."); return
        html = text_to_html(plain)
        self.html_box.config(state="normal")
        self.html_box.delete("1.0","end")
        self.html_box.insert("1.0", html)
        self.html_box.config(state="disabled")
        self.conv_status.config(text="✅ Converted!")

    def _copy_html(self):
        html = self.html_box.get("1.0","end-1c")
        if not html.strip():
            messagebox.showwarning("Empty","Convert first."); return
        self.root.clipboard_clear(); self.root.clipboard_append(html)
        self.conv_status.config(text="📋 Copied!")

    def _send_to_composer(self):
        html = self.html_box.get("1.0","end-1c")
        if not html.strip():
            messagebox.showwarning("Empty","Convert first."); return
        self.body_box.delete("1.0","end")
        self.body_box.insert("1.0", html)
        self.notebook.select(0)
        self.conv_status.config(text="✅ Sent to Composer!")
        self._save_all_drafts(silent=True)

    def _send_to_active_followup(self):
        """Send the converted HTML directly into the currently selected Follow-up sub-tab."""
        html = self.html_box.get("1.0","end-1c")
        if not html.strip():
            messagebox.showwarning("Empty","Convert first."); return
        active_tab = self._get_active_followup_tab()
        if active_tab:
            active_tab["body_box"].delete("1.0","end")
            active_tab["body_box"].insert("1.0", html)
            self.notebook.select(2) # Switch to Manual Follow-up tab
            self.conv_status.config(text=f"✅ Sent to {active_tab['title']}!")
            self._save_all_drafts(silent=True)

    def _clear_converter(self):
        self.plain_box.delete("1.0","end")
        self.html_box.config(state="normal")
        self.html_box.delete("1.0","end")
        self.html_box.config(state="disabled")
        self.conv_status.config(text="")

    # ── Manual FU actions ───────────────────────────────────────
    def _mfu_log(self, msg): self._log_append(self.mfu_log_box, msg)

    def _toggle_mfu_pause(self):
        if self.mfu_pause_event.is_set():
            self.mfu_pause_event.clear()
            self.mfu_pause_btn.config(text="⏸  PAUSE", bg="#f59e0b")
            self._mfu_log("▶️  Resumed!\n")
        else:
            self.mfu_pause_event.set()
            self.mfu_pause_btn.config(text="▶️  RESUME", bg=SUCCESS)
            self._mfu_log("⏸  Paused\n")

    def _stop_manual_followup(self):
        self.mfu_stop_event.set(); self.mfu_pause_event.clear()
        self.mfu_pause_btn.config(state="disabled",
                                  text="⏸  PAUSE", bg="#f59e0b")
        self.mfu_stop_btn.config(state="disabled")

    def _start_manual_followup(self, resume_sent=None):
        if self.mfu_sending: return
        email = self.email_var.get().strip()
        pw    = self.password_var.get().strip()
        url   = self.mfu_sheet_url_var.get().strip()
        label = self.mfu_label_var.get().strip()

        # Read draft from the currently selected Follow-up sub-tab
        active_tab = self._get_active_followup_tab()
        if not active_tab:
            messagebox.showerror("Error", "No follow-up tab found."); return

        tab_title = active_tab["title"]
        subj = active_tab["subject_var"].get().strip() or "Following up on my previous email"
        body = active_tab["body_box"].get("1.0","end-1c").strip()

        if not email or not pw:
            messagebox.showerror("Missing","Enter credentials in Send tab."); return
        if not label:
            messagebox.showerror("Missing","Enter the Gmail label name."); return
        if not url or "docs.google.com" not in url:
            messagebox.showerror("Invalid","Paste a valid Google Sheets URL."); return
        if not body:
            messagebox.showerror("Missing", f"Enter follow-up body for '{tab_title}'."); return

        # Auto-save all drafts so user never loses their changes
        self._save_all_drafts(silent=True)

        delay = self.mfu_delay_var.get()
        self.mfu_stop_event.clear(); self.mfu_pause_event.clear()
        self.mfu_sending = True
        self.mfu_send_btn.config(text=f"⏳  SENDING ({tab_title})...", state="disabled", bg="#555")
        self.mfu_resume_btn.config(state="disabled")
        self.mfu_pause_btn.config(state="normal")
        self.mfu_stop_btn.config(state="normal")
        self.mfu_progress_label.config(text="")
        self._mfu_log(f"\n{'─'*50}\n▶ {datetime.now().strftime('%H:%M:%S')}\n")
        self._mfu_log(f"🎯 Selected Draft: {tab_title}\n")
        self._mfu_log(f"📝 Fallback Subject: {subj}\n\n")

        if resume_sent:
            self._mfu_log(
                f"🔄 Resuming — skipping {len(resume_sent)} already sent\n\n")

        # Fetch the user's real Gmail signature via OAuth before sending.
        try:
            self._mfu_log("🔏 Fetching Gmail signature...\n")
            sig = get_gmail_signature(email)
        except Exception as e:
            self._mfu_log(f"⚠️  Could not fetch signature: {e}\n")
            sig = ""

        threading.Thread(target=send_manual_followups,
            args=(email, pw, url, label, subj, body, delay,
                  self._mfu_log, self._mfu_done,
                  self.mfu_stop_event, self.mfu_pause_event, resume_sent, sig, tab_title),
            daemon=True).start()

    def _resume_manual_followup(self):
        progress = load_progress()
        if not progress:
            messagebox.showinfo("No Session","No interrupted session found."); return
        sent_list = progress.get("sent",[])
        total     = progress.get("total",0)
        timestamp = progress.get("timestamp","unknown")
        remaining = total - len(sent_list)

        active_tab = self._get_active_followup_tab()
        active_name = active_tab["title"] if active_tab else "Follow-up"

        if messagebox.askyesno("Resume",
            f"Session from {timestamp}\n"
            f"Already sent: {len(sent_list)}\n"
            f"Remaining: {remaining}\n"
            f"Draft to use: {active_name}\n\nResume?"):
            self._start_manual_followup(resume_sent=sent_list)

    def _mfu_done(self, ok):
        self.mfu_sending = False
        self.mfu_send_btn.config(
            text="📤  SEND FOLLOW-UPS", state="normal",
            bg="#f59e0b" if ok else "#ef4444")
        self.mfu_resume_btn.config(state="normal")
        self.mfu_pause_btn.config(state="disabled",
                                  text="⏸  PAUSE", bg="#f59e0b")
        self.mfu_stop_btn.config(state="disabled")
        self.mfu_progress_label.config(
            text="✅ Done!" if ok
            else "⚠️  Stopped — click Resume to continue")
        self.root.after(3000, lambda: self.mfu_send_btn.config(bg="#f59e0b"))


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
