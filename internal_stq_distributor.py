# -*- coding: utf-8 -*-
"""
Proyapi STQ Incoming -> Internal Distribution Automation (ULTRA PREMIUM GOLD v2.4.1)

Fixes:
- ✅ Prevent resend loop: PRE-MARK as READ (lock) before sending + SAVE, then POST-MARK again.
- ✅ Strong dedupe: store InternetMessageID (stable) + EntryID fallback in state.
- ✅ Best performance kept: Restrict(UnRead + ReceivedTime) + MAX_SCAN
- ✅ Same body wording (1 vs 2+)
"""

import json
import os
import re
import time
import tempfile
import traceback
import shutil
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any

import win32com.client  # pip install pywin32


# =========================================================
# CONFIG
# =========================================================

BASE_DIR = Path(r"C:\Users\husey\OneDrive\Desktop\Development\Work_Automation")

STATE_PATH = BASE_DIR / "state_proyapi_stq_processed.json"
LOG_PATH = BASE_DIR / "logs" / "proyapi_stq_internal.log"

MAILBOX_NAME_HINT = "spp2dcc@kolin.com.tr"
WATCH_FOLDER_PATH = r"Inbox\From Proyapi"  # TEST üçün lazım olsa: r"Inbox"

# TEMP TEST: sender filter
SENDER_EMAIL = "chuseyn@kolin.com.tr"

# OFT template (signature/layout stable)
OFT_TEMPLATE_PATH = Path(
    r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.2. STQ - Internal Sharing.oft"
)

SIGNATURE_MODE = "oft"  # "oft" or "file"
SIGNATURE_NAME = None

POLL_SECONDS = 5  # test üçün 5; realda 60 tövsiyə
LOOKBACK_DAYS = 7
MAX_SCAN = 200

# Performance & Reliability settings
MAX_RETRY_ATTEMPTS = 3
RETRY_DELAY = 5  # seconds
COM_RECONNECT_INTERVAL = 3600

# Send behavior
SEND_MODE = "display"  # "send" | "draft" | "display"
DISPLAY_MODAL = False
AUTO_SEND_AFTER_DISPLAY = True

DEBUG_PREVIEW = True
DEBUG_PREVIEW_LIMIT = 5

# ✅ STQ attachments you want to distribute
ALLOWED_ATTACH_EXTS = {".xlsx", ".xls", ".xlsm"}

# TEMP TEST recipients
RAW_TO_RECIPIENTS = [
    "Cavid Huseyn <chuseyn@kolin.com.tr>",
]

CC_RECIPIENTS: List[str] = []

# State limits
MAX_STATE_IDS = 5000
MAX_DOC_HISTORY = 5000

# ✅ STQ subject extractor (suffix-safe)
STQ_ID_RE = re.compile(
    r"\bKLN-SPP2-STQ-[A-Z]{2}-GN00-\d{3,4}_R\d{2}(?=_|\b|$)",
    re.IGNORECASE,
)

OUTLOOK_MAILITEM_CLASS = 43

# ✅ Best stable id for emails across sessions (Exchange-safe)
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"


# =========================================================
# COM CONNECTION MANAGER
# =========================================================
class OutlookConnectionManager:
    """Manages Outlook COM connection with auto-reconnect on errors."""

    def __init__(self):
        self.app = None
        self.ns = None
        self.last_connect_time = None
        self.reconnect()

    def reconnect(self):
        log("🔄 Connecting to Outlook...")
        self.app, self.ns = get_outlook_with_retry()
        self.last_connect_time = time.time()

    def should_reconnect(self):
        if self.last_connect_time is None:
            return True
        return (time.time() - self.last_connect_time) >= COM_RECONNECT_INTERVAL

    def execute_with_retry(self, func, *args, **kwargs):
        for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_str = str(e).lower()
                is_com_error = any(
                    k in error_str
                    for k in [
                        "rpc server",
                        "rpc_e_",
                        "disconnected",
                        "not available",
                        "invalid",
                        "automation error",
                    ]
                )
                if is_com_error and attempt < MAX_RETRY_ATTEMPTS:
                    log(f"⚠️ COM error (attempt {attempt}/{MAX_RETRY_ATTEMPTS}): {e}")
                    log("🔄 Reconnecting to Outlook...")
                    time.sleep(RETRY_DELAY)
                    self.reconnect()
                else:
                    raise
        raise RuntimeError("Failed after all retry attempts")

    def get_folder(self, root, folder_path: str):
        def _get_folder():
            parts = folder_path.split("\\")
            f = root
            for p in parts:
                f = f.Folders.Item(p)
            return f

        return self.execute_with_retry(_get_folder)


# =========================================================
# DATA MODEL
# =========================================================
@dataclass
class IncomingPayload:
    doc_id: str
    incoming_entry_id: str
    incoming_uid: str  # InternetMessageID (preferred) else EntryID
    attach_files: List[Path]
    mail_item: object
    sent_dt: Optional[datetime]
    subject: str
    sig: str


# =========================================================
# HELPERS
# =========================================================
def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    ensure_parent(LOG_PATH)
    line = f"[{now_ts()}] {msg}\n"
    print(line, end="")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line)


def safe_iso(dt_obj: Optional[datetime]) -> str:
    if not dt_obj:
        return ""
    try:
        return dt_obj.isoformat()
    except Exception:
        return ""


def load_state() -> Dict[str, Any]:
    try:
        if STATE_PATH.exists():
            with open(STATE_PATH, "r", encoding="utf-8") as f:
                s = json.load(f)
                s.setdefault("processed_entry_ids", [])
                s.setdefault("processed_uids", [])  # ✅ NEW
                s.setdefault("doc_history", {})
                return s
    except Exception as e:
        log(f"STATE read error (will reset safely): {e}")
    return {"processed_entry_ids": [], "processed_uids": [], "doc_history": {}}


def save_state(state: Dict[str, Any]) -> None:
    ensure_parent(STATE_PATH)
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def prune_list(lst: List[str], max_len: int) -> List[str]:
    return lst if len(lst) <= max_len else lst[-max_len:]


def prune_history(hist: Dict[str, Any]) -> Dict[str, Any]:
    if len(hist) <= MAX_DOC_HISTORY:
        return hist
    items = [(k, v.get("last_seen_ts", "")) for k, v in hist.items()]
    items.sort(key=lambda x: x[1])
    drop_count = max(0, len(items) - MAX_DOC_HISTORY)
    for k, _ in items[:drop_count]:
        hist.pop(k, None)
    return hist


def dedupe_recipients(raw_list: List[str]) -> Tuple[List[str], List[str]]:
    seen = set()
    unique = []
    removed = []
    email_re = re.compile(r"<([^>]+)>")

    for item in raw_list:
        item = (item or "").strip()
        if not item:
            continue
        m = email_re.search(item)
        key = m.group(1).strip().lower() if m else item.lower()
        if key in seen:
            removed.append(item)
            continue
        seen.add(key)
        unique.append(item)

    return unique, removed


def normalize_outlook_dt(dt_obj):
    try:
        if dt_obj is None:
            return None
        tz = getattr(dt_obj, "tzinfo", None)
        return dt_obj.replace(tzinfo=None) if tz is not None else dt_obj
    except Exception:
        return dt_obj


def format_en_date(dt_obj: Optional[datetime]) -> str:
    dt_obj = dt_obj or datetime.now()
    return dt_obj.strftime("%d %b %Y")


def extract_first_stq_id(subject: str) -> Optional[str]:
    if not subject:
        return None
    m = STQ_ID_RE.search(subject)
    return m.group(0).upper() if m else None


def attachments_signature(files: List[Path]) -> str:
    parts = []
    for p in files:
        try:
            parts.append(f"{p.name.lower()}:{p.stat().st_size}")
        except Exception:
            parts.append(f"{p.name.lower()}:?")
    return "|".join(sorted(parts))


def get_outlook_with_retry(max_attempts=MAX_RETRY_ATTEMPTS):
    for attempt in range(1, max_attempts + 1):
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            ns = app.GetNamespace("MAPI")
            _ = ns.Folders.Count
            log(f"✅ Outlook connection established (attempt {attempt}/{max_attempts})")
            return app, ns
        except Exception as e:
            log(f"⚠️ Outlook connection attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                log(f"Retrying in {RETRY_DELAY} seconds...")
                time.sleep(RETRY_DELAY)
            else:
                log("❌ Failed to connect to Outlook after all retry attempts")
                raise
    raise RuntimeError("Failed to establish Outlook connection")


def find_mailbox_root(ns, mailbox_hint: str):
    hint = (mailbox_hint or "").strip().lower()
    best = None
    names = []

    for i in range(1, ns.Folders.Count + 1):
        f = ns.Folders.Item(i)
        name = str(getattr(f, "Name", "") or "")
        names.append(name)

        if name.strip().lower() == hint:
            return f
        if hint and hint in name.strip().lower():
            best = f

    if best:
        log(f"✅ Mailbox matched by contains: '{best.Name}'")
        return best

    log("⚠️ Mailbox not matched. Falling back to first mailbox.")
    log("Available mailboxes: " + " | ".join(names))
    return ns.Folders.Item(1)


def get_sender_smtp_address(mail_item) -> str:
    try:
        sender = (getattr(mail_item, "SenderEmailAddress", "") or "").strip()
        if sender and "@" in sender:
            return sender.lower()

        sender_obj = getattr(mail_item, "Sender", None)
        if sender_obj is not None:
            try:
                exch_user = sender_obj.GetExchangeUser()
                if exch_user is not None:
                    primary = (
                        getattr(exch_user, "PrimarySmtpAddress", "") or ""
                    ).strip()
                    if primary and "@" in primary:
                        return primary.lower()
            except Exception:
                pass
        return sender.lower()
    except Exception:
        return ""


def get_internet_message_id(mail_item) -> str:
    try:
        pa = getattr(mail_item, "PropertyAccessor", None)
        if pa is None:
            return ""
        val = pa.GetProperty(PR_INTERNET_MESSAGE_ID)
        return str(val or "").strip().lower()
    except Exception:
        return ""


# =========================================================
# SIGNATURE (fallback)
# =========================================================
def get_signature_html(signature_name: Optional[str] = None) -> str:
    appdata = os.environ.get("APPDATA", "")
    sig_dir = Path(appdata) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""

    if signature_name:
        target = sig_dir / f"{signature_name}.htm"
        if target.exists():
            return target.read_text(encoding="utf-8", errors="ignore")

    htm_files = list(sig_dir.glob("*.htm"))
    if not htm_files:
        return ""

    newest = max(htm_files, key=lambda p: p.stat().st_mtime)
    return newest.read_text(encoding="utf-8", errors="ignore")


# =========================================================
# ATTACHMENTS
# =========================================================
def save_allowed_attachments(mail_item) -> List[Path]:
    temp_dir = Path(tempfile.mkdtemp(prefix="proyapi_stq_"))
    saved: List[Path] = []

    atts = mail_item.Attachments
    for i in range(1, atts.Count + 1):
        att = atts.Item(i)
        filename = str(att.FileName or "")
        ext = Path(filename).suffix.lower()

        if ext in ALLOWED_ATTACH_EXTS:
            safe_name = re.sub(r'[<>:"/\\|?*]', "_", filename)
            out_path = temp_dir / safe_name
            att.SaveAsFile(str(out_path))
            saved.append(out_path)

    return saved


def cleanup_temp(files: List[Path]) -> None:
    try:
        if not files:
            return
        parent = files[0].parent
        if parent.exists() and parent.name.startswith("proyapi_stq_"):
            shutil.rmtree(parent, ignore_errors=True)
    except Exception as e:
        log(f"⚠️ Could not delete temp dir: {e}")


# =========================================================
# MAIL BUILD
# =========================================================
def build_internal_body_html_stq(sent_dt: Optional[datetime], attach_count: int) -> str:
    date_str = format_en_date(sent_dt)
    font_style = "font-family:Bahnschrift, Calibri, Arial, sans-serif; font-size:11pt;"

    if attach_count == 1:
        tail = "paylaşılan STQ dosyası ekte sunulmuştur."
    else:
        tail = "paylaşılan STQ dosyaları ekte sunulmuştur."

    return (
        f"<div style='{font_style}'>"
        "<p>Sayın İlgililer,</p>"
        f"<p>Müşavir tarafından <b>{date_str}</b> tarihinde Sitalçay 2 Üretim Tesisi kapsamında {tail}</p>"
        "<br>"
        "</div>"
    )


def prepend_intro_keep_existing_html(msg, intro_html: str) -> None:
    existing_html = msg.HTMLBody or ""
    msg.HTMLBody = intro_html + existing_html


def create_internal_mail(
    outlook_app,
    incoming_subject: str,
    sent_dt: Optional[datetime],
    attach_files: List[Path],
) -> None:
    to_unique, removed = dedupe_recipients(RAW_TO_RECIPIENTS)
    if removed:
        log(f"Recipient duplicates removed ({len(removed)}): " + " | ".join(removed))

    if SIGNATURE_MODE == "oft" and OFT_TEMPLATE_PATH.exists():
        try:
            msg = outlook_app.CreateItemFromTemplate(str(OFT_TEMPLATE_PATH))
        except Exception as e:
            log(f"⚠️ OFT failed ({e}). Falling back to blank mail item.")
            msg = outlook_app.CreateItem(0)
    else:
        msg = outlook_app.CreateItem(0)
        sig_html = (
            get_signature_html(SIGNATURE_NAME) if SIGNATURE_MODE == "file" else ""
        )
        if sig_html:
            msg.HTMLBody = sig_html

    msg.To = "; ".join(to_unique)
    msg.CC = "; ".join(CC_RECIPIENTS) if CC_RECIPIENTS else ""
    msg.Subject = incoming_subject.strip()

    intro_html = build_internal_body_html_stq(sent_dt, len(attach_files))
    prepend_intro_keep_existing_html(msg, intro_html)

    for p in attach_files:
        if p.exists():
            msg.Attachments.Add(str(p))

    if SEND_MODE == "send":
        msg.Send()
        log(f"✅ Sent internal mail: {msg.Subject} | Attachments={len(attach_files)}")
        return

    if SEND_MODE == "draft":
        msg.Save()
        log(f"📝 Saved draft: {msg.Subject} | Attachments={len(attach_files)}")
        return

    msg.Display(DISPLAY_MODAL)
    log(
        f"👀 Displayed: {msg.Subject} | Attachments={len(attach_files)} (modal={DISPLAY_MODAL})"
    )

    if AUTO_SEND_AFTER_DISPLAY and not DISPLAY_MODAL:
        msg.Send()
        log(f"✅ Sent after display: {msg.Subject} | Attachments={len(attach_files)}")


def mark_mail_as_read(mail_item) -> None:
    """Hard mark as read + save (Exchange sometimes needs explicit save)."""
    try:
        mail_item.UnRead = False
        mail_item.Save()
    except Exception as e:
        log(f"⚠️ Could not mark mail as read: {e}")


def lock_mail_as_read_before_send(mail_item) -> None:
    """
    ULTRA IMPORTANT:
    Mark read BEFORE send to prevent duplicate processing if Outlook/Exchange lags.
    """
    mark_mail_as_read(mail_item)


# =========================================================
# SCAN (FAST)
# =========================================================
def try_restrict_items(items, cutoff_dt: datetime):
    try:
        cutoff_str = cutoff_dt.strftime("%m/%d/%Y %I:%M %p")
        restriction = f"[UnRead] = True AND [ReceivedTime] >= '{cutoff_str}'"
        return items.Restrict(restriction)
    except Exception as e:
        log(f"⚠️ Restrict(UnRead+ReceivedTime) failed, fallback: {e}")
        try:
            return items.Restrict("[UnRead] = True")
        except Exception as e2:
            log(f"⚠️ Restrict(UnRead) failed too, fallback manual: {e2}")
            return items


def scan_unread_stq_mails(
    folder, processed_entry_ids: set, processed_uids: set
) -> List[IncomingPayload]:
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    items = try_restrict_items(items, cutoff)

    results: List[IncomingPayload] = []
    checked = matched = 0
    skipped_read = skipped_sender = skipped_no_stq = skipped_processed = 0
    preview_left = DEBUG_PREVIEW_LIMIT

    try:
        count = min(items.Count, MAX_SCAN)
    except Exception:
        count = MAX_SCAN

    for idx in range(1, count + 1):
        try:
            mail = items.Item(idx)
            checked += 1

            if getattr(mail, "Class", None) != OUTLOOK_MAILITEM_CLASS:
                continue
            if not bool(getattr(mail, "UnRead", False)):
                skipped_read += 1
                continue

            entry_id = str(getattr(mail, "EntryID", "") or "")
            uid = get_internet_message_id(mail) or entry_id

            if entry_id and entry_id in processed_entry_ids:
                skipped_processed += 1
                continue
            if uid and uid in processed_uids:
                skipped_processed += 1
                continue

            subject = str(getattr(mail, "Subject", "") or "")
            received = normalize_outlook_dt(getattr(mail, "ReceivedTime", None))
            if received and received < cutoff:
                break

            stq_id = extract_first_stq_id(subject)
            if not stq_id:
                skipped_no_stq += 1
                continue

            sender = get_sender_smtp_address(mail)
            sender_l = (sender or "").lower()

            if DEBUG_PREVIEW and preview_left > 0:
                log(
                    f"DEBUG => UnRead={mail.UnRead} | UID={uid[:60]} | Sender={sender_l} | Subject={subject}"
                )
                preview_left -= 1

            if SENDER_EMAIL.lower() not in sender_l:
                skipped_sender += 1
                continue

            sent_dt = normalize_outlook_dt(getattr(mail, "SentOn", None)) or received

            files = save_allowed_attachments(mail)
            if not files:
                log(f"⚠️ STQ matched but no allowed attachment found: {subject}")
                continue

            sig = attachments_signature(files)

            results.append(
                IncomingPayload(
                    doc_id=stq_id,
                    incoming_entry_id=entry_id,
                    incoming_uid=uid,
                    attach_files=files,
                    mail_item=mail,
                    sent_dt=sent_dt,
                    subject=subject,
                    sig=sig,
                )
            )
            matched += 1

        except Exception as e:
            log(f"Error while scanning item #{idx}: {e}")
            log(traceback.format_exc())

    log(
        f"Scan stats => checked:{checked} matched:{matched} | "
        f"skipped: read={skipped_read} sender={skipped_sender} no_stq={skipped_no_stq} processed={skipped_processed}"
    )
    return results


# =========================================================
# MAIN LOOP
# =========================================================
def main():
    log("=== Proyapi STQ Internal Distributor started (v2.4.1) ===")

    state = load_state()
    processed_entry_ids = set(state.get("processed_entry_ids", []))
    processed_uids = set(state.get("processed_uids", []))
    doc_history = state.get("doc_history", {}) or {}

    conn_mgr = OutlookConnectionManager()
    outlook_app = conn_mgr.app
    ns = conn_mgr.ns

    root = find_mailbox_root(ns, MAILBOX_NAME_HINT)
    folder = conn_mgr.get_folder(root, WATCH_FOLDER_PATH)

    log(f"Watching: {root.Name} / {WATCH_FOLDER_PATH}")
    log(f"Sender filter: {SENDER_EMAIL}")
    log(
        f"Mode: {SEND_MODE} | Poll: {POLL_SECONDS}s | Lookback: {LOOKBACK_DAYS}d | MaxScan:{MAX_SCAN}"
    )
    log(f"Allowed attachment exts: {sorted(list(ALLOWED_ATTACH_EXTS))}")
    log(
        "Filter: ONLY UNREAD mails. Lock mail as READ before send to prevent duplicates."
    )

    loop_count = 0
    while True:
        try:
            loop_count += 1

            if conn_mgr.should_reconnect():
                log("⏰ Proactive reconnect scheduled (hourly refresh)")
                conn_mgr.reconnect()
                outlook_app = conn_mgr.app
                ns = conn_mgr.ns
                root = find_mailbox_root(ns, MAILBOX_NAME_HINT)
                folder = conn_mgr.get_folder(root, WATCH_FOLDER_PATH)

            log(f"\n[Loop #{loop_count}] Scanning for unread STQ mails...")
            payloads = scan_unread_stq_mails(
                folder, processed_entry_ids, processed_uids
            )

            if not payloads:
                log(f"No matching unread STQ found. Sleeping {POLL_SECONDS}s...")
            else:
                for p in payloads:
                    prev = (
                        doc_history.get(p.doc_id, {})
                        if isinstance(doc_history, dict)
                        else {}
                    )
                    prev_sig = prev.get("last_sig", "")

                    log(
                        f"FOUND => {p.doc_id} | Files={len(p.attach_files)} | Date={format_en_date(p.sent_dt)}"
                    )

                    # ✅ lock & state-update EARLY (anti-loop)
                    try:
                        lock_mail_as_read_before_send(p.mail_item)

                        if p.incoming_entry_id:
                            processed_entry_ids.add(p.incoming_entry_id)
                        if p.incoming_uid:
                            processed_uids.add(p.incoming_uid)

                        state["processed_entry_ids"] = prune_list(
                            list(processed_entry_ids), MAX_STATE_IDS
                        )
                        state["processed_uids"] = prune_list(
                            list(processed_uids), MAX_STATE_IDS
                        )
                        save_state(state)

                    except Exception as lock_err:
                        log(f"⚠️ Could not lock mail as read before send: {lock_err}")

                    try:
                        # anti-spam: same STQ + same attachments => SKIP
                        if prev and prev_sig and p.sig == prev_sig:
                            log(
                                f"⏭️ SKIP duplicate content: {p.doc_id} (same attachment signature)"
                            )

                            doc_history[p.doc_id]["last_seen_ts"] = now_ts()
                            state["doc_history"] = prune_history(doc_history)
                            save_state(state)
                            continue

                        create_internal_mail(
                            outlook_app=outlook_app,
                            incoming_subject=p.subject,
                            sent_dt=p.sent_dt,
                            attach_files=p.attach_files,
                        )

                        # ✅ finalize: mark read AGAIN (Exchange lag safe)
                        mark_mail_as_read(p.mail_item)

                        doc_history[p.doc_id] = {
                            "last_entry_id": p.incoming_entry_id,
                            "last_uid": p.incoming_uid,
                            "last_sent_dt": safe_iso(p.sent_dt),
                            "last_sig": p.sig,
                            "last_seen_ts": now_ts(),
                            "last_subject": p.subject,
                        }
                        state["doc_history"] = prune_history(doc_history)
                        save_state(state)

                        log(
                            f"✅ DONE => {p.doc_id} (incoming marked READ + stored in state)"
                        )

                    finally:
                        cleanup_temp(p.attach_files)

        except Exception as e:
            error_str = str(e).lower()
            is_com_error = any(
                k in error_str
                for k in [
                    "rpc server",
                    "rpc_e_",
                    "disconnected",
                    "not available",
                    "invalid",
                    "automation error",
                ]
            )

            if is_com_error:
                log(f"❌ COM/RPC error in main loop: {e}")
                log(traceback.format_exc())
                log("🔄 Attempting to reconnect Outlook...")
                try:
                    conn_mgr.reconnect()
                    outlook_app = conn_mgr.app
                    ns = conn_mgr.ns
                    root = find_mailbox_root(ns, MAILBOX_NAME_HINT)
                    folder = conn_mgr.get_folder(root, WATCH_FOLDER_PATH)
                    log("✅ Reconnection successful, continuing...")
                except Exception as reconnect_error:
                    log(f"❌ Reconnection failed: {reconnect_error}")
                    log(f"⏸️ Waiting {POLL_SECONDS}s before retry...")
            else:
                log(f"❌ Loop error: {e}")
                log(traceback.format_exc())

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
