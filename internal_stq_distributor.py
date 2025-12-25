# -*- coding: utf-8 -*-
"""
proyapi_stq_internal_distributor.py (ULTRA PREMIUM ULTIMATE v3.1 - SINGLE RUN + STQ HARD MATCH + SQLITE WAL + OVERLAP LOCK)

✅ Load-bearing guarantees:
- Single-run (NO internal polling). Use Windows Task Scheduler every 3 minutes.
- Overlap-safe: prevents concurrent runs via lock file.
- Snapshot ALL unread once => prevents Outlook Items drift
- STQ trigger: prefix anywhere (subject OR attachment filename)
- Multi-STQ subjects separated by '/' supported (extracts all IDs)
- Attachments policy: ANY file type, ANY count (saves everything)
- PRE-LOCK (mark read + DB processed) before sending => prevents resend loops
- Shared mailbox safe: GetItemFromID(entry_id, store_id)
- SendUsingAccount best-effort
- RPC-proof reconnect + single retry per item
- SQLite WAL for ultimate fast state
"""

from __future__ import annotations

import gc
import hashlib
import os
import re
import shutil
import sqlite3
import tempfile
import traceback
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple

import pythoncom
import pywintypes
import win32com.client  # pip install pywin32


# =========================
# CONFIG
# =========================
BASE_DIR = Path(r"C:\Users\husey\OneDrive\Desktop\Development\Work_Automation")
DB_PATH = BASE_DIR / "state_proyapi_stq_internal.db"
LOG_PATH = BASE_DIR / "logs" / "proyapi_stq_internal_ultimate.log"
LOCK_PATH = BASE_DIR / "locks" / "proyapi_stq_internal.lock"

MAILBOX_HINT = "spp2dcc@kolin.com.tr"
WATCH_FOLDER = r"Inbox\From Proyapi"  # test üçün r"Inbox" edə bilərsən

# Sender filter (contains). Set "" to disable.
SENDER_FILTER_CONTAINS = "chuseyn@kolin.com.tr"

# Receiver(s)
TO_RECIPIENTS = ["Huseyn Cavid <huseyn.cavid.dev@outlook.com>"]
CC_RECIPIENTS: List[str] = []

# Template
OFT_TEMPLATE_PATH = Path(
    r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.2. STQ - Internal Sharing.oft"
)

# Scan
LOOKBACK_DAYS = 14
MAX_SCAN = 800  # unread snapshot only

# Send behavior
SEND_MODE = "send"  # "send" | "draft" | "display"
DISPLAY_MODAL = False
AUTO_SEND_AFTER_DISPLAY = False

# Attachments
SAVE_ANY_ATTACHMENT = True
SKIP_DUPLICATE_CONTENT = False  # if True: skip when same doc_key signature seen before


# =========================
# CONSTANTS
# =========================
OUTLOOK_MAILITEM_CLASS = 43
RPC_ERR = -2147023174  # "The RPC server is unavailable."

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"

STQ_PREFIX = "KLN-SPP2-STQ"

# Match until _Rxx, regardless what comes after (_ / space / / / end)
# Examples:
# - KLN-SPP2-STQ-MC-EW13-349_R00_Prokon_Reply
# - KLN-SPP2-STQ-WE-GN00-362_R00_Prokon_Reply/ KLN-SPP2-STQ-EL-ES01-363_R00_...
STQ_ID_RE = re.compile(
    r"(" + re.escape(STQ_PREFIX) + r"[A-Z0-9\-_]*?_R\d{2})(?=($|[_/\s]))",
    re.IGNORECASE,
)


# =========================
# UTIL / LOGGING
# =========================
def ensure_parent(p: Path) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    ensure_parent(LOG_PATH)
    line = f"[{now_str()}] {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def is_rpc_error(e: Exception) -> bool:
    try:
        return isinstance(e, pywintypes.com_error) and e.args and e.args[0] == RPC_ERR
    except Exception:
        return False


def short_hash(s: str) -> str:
    b = (s or "").encode("utf-8", errors="ignore")
    return hashlib.sha1(b).hexdigest()[:10].upper()


def safe_dt(x) -> Optional[datetime]:
    try:
        if not x:
            return None
        return x.replace(tzinfo=None)
    except Exception:
        return None


@contextmanager
def single_instance_lock(lock_path: Path):
    """
    Prevent concurrent runs (Task Scheduler can overlap).
    Uses exclusive create. If file exists -> another run is active.
    """
    ensure_parent(lock_path)
    try:
        fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise RuntimeError(f"LOCKED: another instance is already running: {lock_path}")
    try:
        os.write(
            fd, f"pid={os.getpid()} ts={datetime.now().isoformat()}".encode("utf-8")
        )
        os.close(fd)
        yield
    finally:
        try:
            lock_path.unlink(missing_ok=True)  # py3.8+ supports missing_ok; ok on 3.11
        except Exception:
            pass


# =========================
# SQLITE STATE (FAST + SAFE)
# =========================
def open_db() -> sqlite3.Connection:
    ensure_parent(DB_PATH)
    con = sqlite3.connect(DB_PATH, timeout=30.0)
    con.execute("PRAGMA journal_mode=WAL;")
    con.execute("PRAGMA synchronous=NORMAL;")
    con.execute("PRAGMA temp_store=MEMORY;")
    con.execute("PRAGMA cache_size=-64000;")  # ~64MB
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS processed (
            uid TEXT PRIMARY KEY,
            ts  TEXT NOT NULL
        );
        """
    )
    con.execute(
        """
        CREATE TABLE IF NOT EXISTS stq_history (
            doc_key TEXT PRIMARY KEY,
            last_sig TEXT,
            last_dt  TEXT,
            last_ts  TEXT NOT NULL
        );
        """
    )
    con.execute("CREATE INDEX IF NOT EXISTS idx_processed_uid ON processed(uid);")
    return con


def already_processed(con: sqlite3.Connection, uid: str) -> bool:
    if not uid:
        return False
    row = con.execute("SELECT 1 FROM processed WHERE uid=? LIMIT 1", (uid,)).fetchone()
    return row is not None


def mark_processed(con: sqlite3.Connection, uid: str) -> None:
    if not uid:
        return
    con.execute(
        "INSERT OR IGNORE INTO processed(uid, ts) VALUES(?, ?)",
        (uid, datetime.now().isoformat()),
    )


def get_hist(con: sqlite3.Connection, doc_key: str) -> Tuple[str, str]:
    row = con.execute(
        "SELECT last_sig, last_dt FROM stq_history WHERE doc_key=? LIMIT 1", (doc_key,)
    ).fetchone()
    return (row[0], row[1]) if row else ("", "")


def upsert_hist(con: sqlite3.Connection, doc_key: str, sig: str, dt_iso: str) -> None:
    con.execute(
        """
        INSERT INTO stq_history(doc_key, last_sig, last_dt, last_ts)
        VALUES(?, ?, ?, ?)
        ON CONFLICT(doc_key) DO UPDATE SET
            last_sig=excluded.last_sig,
            last_dt=excluded.last_dt,
            last_ts=excluded.last_ts
        """,
        (doc_key, sig, dt_iso, datetime.now().isoformat()),
    )


# =========================
# OUTLOOK CORE
# =========================
@contextmanager
def outlook_session():
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            app = win32com.client.DispatchEx("Outlook.Application")

        ns = app.GetNamespace("MAPI")
        try:
            ns.Logon("", "", False, False)
        except Exception:
            pass

        _ = ns.Folders.Count  # touch
        yield app, ns
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def find_mailbox_root(ns, hint: str):
    hint = (hint or "").strip().lower()
    best = None
    for i in range(1, ns.Folders.Count + 1):
        f = ns.Folders.Item(i)
        name = str(getattr(f, "Name", "") or "").strip()
        low = name.lower()
        if low == hint:
            return f
        if hint and hint in low:
            best = f
    return best or ns.Folders.Item(1)


def get_folder(root, path: str):
    f = root
    for p in path.split("\\"):
        f = f.Folders.Item(p)
    return f


def resolve_watch_folder(ns):
    root = find_mailbox_root(ns, MAILBOX_HINT)
    folder = get_folder(root, WATCH_FOLDER)
    return root, folder


def sender_smtp(mail) -> str:
    """
    Robust sender SMTP extraction (Exchange-safe).
    """
    try:
        s = (getattr(mail, "SenderEmailAddress", "") or "").strip()
        if s and "@" in s:
            return s.lower()

        sender_obj = getattr(mail, "Sender", None)
        if sender_obj:
            try:
                ex = sender_obj.GetExchangeUser()
                if ex:
                    p = (getattr(ex, "PrimarySmtpAddress", "") or "").strip()
                    if p and "@" in p:
                        return p.lower()
            except Exception:
                pass

        pa = getattr(mail, "PropertyAccessor", None)
        if pa:
            try:
                v = str(pa.GetProperty(PR_SENDER_SMTP_ADDRESS) or "").strip()
                if v and "@" in v:
                    return v.lower()
            except Exception:
                pass

        return (s or "").lower()
    except Exception:
        return ""


def internet_message_id(mail) -> str:
    try:
        pa = getattr(mail, "PropertyAccessor", None)
        if not pa:
            return ""
        v = pa.GetProperty(PR_INTERNET_MESSAGE_ID)
        return str(v or "").strip().lower()
    except Exception:
        return ""


def mark_read(mail) -> None:
    try:
        mail.UnRead = False
        mail.Save()
    except Exception as e:
        log(f"⚠️ mark_read failed: {e}")


def try_set_sending_account(ns, msg, mailbox_hint: str) -> None:
    try:
        hint = (mailbox_hint or "").lower().strip()
        accts = ns.Session.Accounts
        for i in range(1, accts.Count + 1):
            a = accts.Item(i)
            smtp = (getattr(a, "SmtpAddress", "") or "").lower().strip()
            disp = (getattr(a, "DisplayName", "") or "").lower().strip()
            if hint and (hint in smtp or hint in disp):
                msg.SendUsingAccount = a
                log(f"✅ SendUsingAccount set => {smtp or disp}")
                return
    except Exception as e:
        log(f"⚠️ SendUsingAccount not set: {e}")


# =========================
# SCAN (SNAPSHOT)
# =========================
def restrict_unread(items, cutoff: datetime):
    try:
        cut = cutoff.strftime("%m/%d/%Y %I:%M %p")
        return items.Restrict(f"[UnRead] = True AND [ReceivedTime] >= '{cut}'")
    except Exception:
        try:
            return items.Restrict("[UnRead] = True")
        except Exception:
            return items


def snapshot_unread_entryids(items, max_n: int) -> List[str]:
    out: List[str] = []
    try:
        n = min(items.Count, max_n)
    except Exception:
        n = max_n

    for i in range(1, n + 1):
        try:
            m = items.Item(i)
            if getattr(m, "Class", None) != OUTLOOK_MAILITEM_CLASS:
                del m
                continue
            if not bool(getattr(m, "UnRead", False)):
                del m
                continue
            eid = str(getattr(m, "EntryID", "") or "")
            if eid:
                out.append(eid)
            del m
        except Exception:
            continue
    return out


# =========================
# STQ DETECTION + EXTRACTION
# =========================
def extract_stq_ids(text: str) -> List[str]:
    if not text:
        return []
    ids = [m.group(1).upper() for m in STQ_ID_RE.finditer(text)]
    seen = set()
    out: List[str] = []
    for x in ids:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def mail_has_stq_prefix_anywhere(mail, subject: str) -> bool:
    if STQ_PREFIX.lower() in (subject or "").lower():
        return True
    try:
        atts = getattr(mail, "Attachments", None)
        if not atts:
            return False
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = str(getattr(att, "FileName", "") or "")
            del att
            if STQ_PREFIX.lower() in fn.lower():
                return True
        del atts
    except Exception:
        pass
    return False


def build_doc_key(subject: str, mail) -> str:
    # 1) subject IDs
    ids = extract_stq_ids(subject or "")
    # 2) attachment filename IDs if subject is weird
    if not ids:
        try:
            atts = getattr(mail, "Attachments", None)
            if atts:
                for i in range(1, atts.Count + 1):
                    att = atts.Item(i)
                    fn = str(getattr(att, "FileName", "") or "")
                    del att
                    ids += extract_stq_ids(fn)
        except Exception:
            pass

    # Dedupe keep order
    seen = set()
    uniq: List[str] = []
    for x in ids:
        if x not in seen:
            seen.add(x)
            uniq.append(x)

    if uniq:
        return " | ".join(uniq)

    # HARD fallback: prefix exists but no parseable id
    return f"{STQ_PREFIX}_AUTO_{short_hash(subject or '')}"


# =========================
# ATTACHMENTS (ANY TYPE)
# =========================
def save_attachments_any(mail, prefix: str) -> List[Path]:
    temp_dir = Path(tempfile.mkdtemp(prefix=prefix))
    out: List[Path] = []

    atts = mail.Attachments
    cnt = atts.Count
    for i in range(1, cnt + 1):
        att = atts.Item(i)
        fn = str(getattr(att, "FileName", "") or "")
        if not fn:
            fn = f"attachment_{i}"
        safe = re.sub(r'[<>:"/\\|?*]', "_", fn)
        p = temp_dir / safe
        try:
            att.SaveAsFile(str(p))
            out.append(p)
        except Exception as e:
            log(f"⚠️ Attachment save failed ({fn}): {e}")
        del att
    del atts
    return out


def attachments_signature(files: List[Path]) -> str:
    parts: List[str] = []
    for p in files:
        try:
            parts.append(f"{p.name.lower()}:{p.stat().st_size}")
        except Exception:
            parts.append(f"{p.name.lower()}:?")
    return "|".join(sorted(parts))


def cleanup_temp(files: List[Path], prefix: str) -> None:
    if not files:
        return
    try:
        d = files[0].parent
        if d.exists() and d.name.startswith(prefix):
            shutil.rmtree(d, ignore_errors=True)
    except Exception as e:
        log(f"⚠️ cleanup failed: {e}")


# =========================
# BODY + SEND
# =========================
def en_date(dt: Optional[datetime]) -> str:
    return (dt or datetime.now()).strftime("%d %b %Y")


def build_intro_html(sent_dt: Optional[datetime], attach_count: int) -> str:
    tail = "STQ dosyası" if attach_count == 1 else "STQ dosyaları"
    return (
        "<div style='font-family:Bahnschrift,Calibri,Arial,sans-serif;font-size:11pt;'>"
        "<p>Sayın İlgililer,</p>"
        f"<p>Müşavir tarafından <b>{en_date(sent_dt)}</b> tarihinde Sitalçay 2 Üretim Tesisi kapsamında paylaşılan {tail} ekte sunulmuştur.</p>"
        "<br></div>"
    )


def create_internal_mail(
    app, ns, subject: str, sent_dt: Optional[datetime], files: List[Path]
) -> None:
    # Template first
    try:
        msg = (
            app.CreateItemFromTemplate(str(OFT_TEMPLATE_PATH))
            if OFT_TEMPLATE_PATH.exists()
            else app.CreateItem(0)
        )
    except Exception:
        msg = app.CreateItem(0)

    try_set_sending_account(ns, msg, MAILBOX_HINT)

    msg.To = "; ".join(TO_RECIPIENTS)
    msg.CC = "; ".join(CC_RECIPIENTS) if CC_RECIPIENTS else ""
    msg.Subject = (subject or "").strip()

    intro = build_intro_html(sent_dt, len(files))
    msg.HTMLBody = intro + (msg.HTMLBody or "")

    for p in files:
        if p.exists():
            msg.Attachments.Add(str(p))

    try:
        msg.Save()
    except Exception:
        pass

    if SEND_MODE == "send":
        msg.Send()
        return
    if SEND_MODE == "draft":
        msg.Save()
        return

    msg.Display(DISPLAY_MODAL)
    if AUTO_SEND_AFTER_DISPLAY and not DISPLAY_MODAL:
        msg.Send()


# =========================
# MAIN (SINGLE RUN)
# =========================
def run_once() -> int:
    """
    Returns: number of processed STQ mails in this run.
    """
    log("=== STQ Internal Distributor ULTIMATE v3.1 (SINGLE RUN) started ===")
    log(
        f"WATCH='{WATCH_FOLDER}' | MAX_SCAN={MAX_SCAN} | LOOKBACK_DAYS={LOOKBACK_DAYS} | SEND_MODE={SEND_MODE}"
    )
    log(
        f"SENDER_FILTER_CONTAINS='{SENDER_FILTER_CONTAINS}' | TO='{'; '.join(TO_RECIPIENTS)}'"
    )

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    processed_count = 0
    debug_left = 10

    with open_db() as con:
        with outlook_session() as (app, ns):
            root, folder = resolve_watch_folder(ns)
            store_id = getattr(folder, "StoreID", None)

            items = folder.Items
            items = restrict_unread(items, cutoff)
            try:
                items.Sort("[ReceivedTime]", True)  # newest first
            except Exception:
                pass

            entry_ids = snapshot_unread_entryids(items, MAX_SCAN)
            log(
                f"Snapshot unread: {len(entry_ids)} | mailbox='{root.Name}' store_id={'OK' if store_id else 'NONE'}"
            )

            for eid in entry_ids:
                mail = None
                tried_retry = False

                while True:
                    try:
                        mail = (
                            ns.GetItemFromID(eid, store_id)
                            if store_id
                            else ns.GetItemFromID(eid)
                        )

                        if getattr(mail, "Class", None) != OUTLOOK_MAILITEM_CLASS:
                            break
                        if not bool(getattr(mail, "UnRead", False)):
                            break

                        received = safe_dt(getattr(mail, "ReceivedTime", None))
                        if received and received < cutoff:
                            break

                        sender = sender_smtp(mail)
                        if SENDER_FILTER_CONTAINS and (
                            SENDER_FILTER_CONTAINS.lower() not in (sender or "").lower()
                        ):
                            if debug_left > 0:
                                log(
                                    f"DEBUG_SKIP(sender) => sender={sender} subject='{getattr(mail,'Subject','')}'"
                                )
                                debug_left -= 1
                            break

                        subject = str(getattr(mail, "Subject", "") or "")

                        # HARD STQ trigger: prefix anywhere
                        if not mail_has_stq_prefix_anywhere(mail, subject):
                            if debug_left > 0:
                                log(f"DEBUG_SKIP(no STQ prefix) => subject='{subject}'")
                                debug_left -= 1
                            break

                        uid = internet_message_id(mail) or str(
                            getattr(mail, "EntryID", "") or ""
                        )
                        if uid and already_processed(con, uid):
                            break

                        doc_key = build_doc_key(subject, mail)
                        if debug_left > 0:
                            log(
                                f"DEBUG_MATCH(STQ) => doc_key='{doc_key}' sender='{sender}' subject='{subject}'"
                            )
                            debug_left -= 1

                        # PRE-LOCK to avoid resend loops
                        mark_read(mail)
                        mark_processed(con, uid)
                        con.commit()

                        sent_dt = safe_dt(getattr(mail, "SentOn", None)) or received

                        prefix = "proyapi_stq_"
                        files = (
                            save_attachments_any(mail, prefix)
                            if SAVE_ANY_ATTACHMENT
                            else []
                        )

                        if not files:
                            log(
                                f"⚠️ STQ matched but has NO attachments: subject='{subject}'"
                            )
                            cleanup_temp(files, prefix)
                            break

                        sig = attachments_signature(files)

                        if SKIP_DUPLICATE_CONTENT:
                            prev_sig, _ = get_hist(con, doc_key)
                            if prev_sig and prev_sig == sig:
                                log(
                                    f"⏭️ SKIP duplicate content: doc_key='{doc_key}' (same sig)"
                                )
                                cleanup_temp(files, prefix)
                                break

                        create_internal_mail(app, ns, subject, sent_dt, files)

                        # finalize
                        mark_read(mail)
                        upsert_hist(
                            con, doc_key, sig, (sent_dt.isoformat() if sent_dt else "")
                        )
                        con.commit()

                        cleanup_temp(files, prefix)
                        processed_count += 1
                        log(f"✅ DONE [STQ] doc_key='{doc_key}' files={len(files)}")
                        break

                    except Exception as e:
                        if is_rpc_error(e):
                            log(f"⚠️ RPC dropped. Reconnecting... ({e})")
                            gc.collect()

                            if tried_retry:
                                log(
                                    "❌ RPC retry already used for this item. Skipping."
                                )
                                break

                            # reconnect session once for this item
                            try:
                                with outlook_session() as (app2, ns2):
                                    app, ns = app2, ns2
                                    tried_retry = True
                                    log("✅ Outlook re-attached. Retrying item once.")
                                    continue
                            except Exception as e2:
                                log(f"❌ Reconnect failed: {e2}")
                                break

                        log(f"⚠️ Item error: {e}")
                        log(traceback.format_exc())
                        break

                    finally:
                        try:
                            if mail is not None:
                                del mail
                        except Exception:
                            pass

    if processed_count == 0:
        log("No matching unread STQ.")
    return processed_count


def main() -> None:
    try:
        with single_instance_lock(LOCK_PATH):
            run_once()
    except RuntimeError as e:
        # Another instance is running
        log(f"⏭️ {e}")
    except Exception as e:
        log(f"❌ Fatal error: {e}")
        log(traceback.format_exc())


if __name__ == "__main__":
    main()
