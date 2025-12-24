# -*- coding: utf-8 -*-
"""
proyapi_unified_distributor.py (ULTRA PREMIUM v5.2 - INSTANT SEND + STQ PREFIX MATCH + ACCOUNT SAFE + RPC RETRY)

Load-bearing fixes:
- ✅ Scans ALL unread mails once (snapshot) -> prevents Outlook Items drift
- ✅ STQ detection: if "KLN-SPP2-STQ" exists anywhere (subject/attachment) -> treat as STQ (no building-code strictness)
- ✅ Extracts STQ/TRN doc_id robustly (cuts at _Rxx) and ignores Reply suffixes
- ✅ Subject first, attachment filename fallback
- ✅ Shared mailbox safe: ns.GetItemFromID(entry_id, store_id)
- ✅ PRE-LOCK (mark read + DB processed) before sending -> prevents resend loops
- ✅ TRN UPDATED + signature SKIP (anti-spam)
- ✅ SendUsingAccount best-effort (shared mailbox safe)
- ✅ RPC-proof reconnect + single retry per item
"""

import gc
import re
import time
import sqlite3
import tempfile
import shutil
import traceback
import hashlib
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Set, Tuple

import pythoncom
import pywintypes
import win32com.client  # pip install pywin32


# =========================
# CONFIG
# =========================
BASE_DIR = Path(r"C:\Users\husey\OneDrive\Desktop\Development\Work_Automation")
DB_PATH = BASE_DIR / "state_proyapi_unified.db"
LOG_PATH = BASE_DIR / "logs" / "proyapi_unified.log"

MAILBOX_HINT = "spp2dcc@kolin.com.tr"
WATCH_FOLDER = r"Inbox\From Proyapi"

POLL_SECONDS = 300
LOOKBACK_DAYS = 10
MAX_SCAN = 500  # unread snapshot only

# ✅ Instant send by default
SEND_MODE = "send"  # "send" | "draft" | "display"
DISPLAY_MODAL = False  # used only when SEND_MODE="display"
AUTO_SEND_AFTER_DISPLAY = False  # keep False for stability

TO_RECIPIENTS = [
    "Hakan Teke <hteke@kolin.com.tr>",
    "Saadet Gülbin Kalaycı <sgkalayci@kolin.com.tr>",
    "Ertuğ Kuban <ekuban@kolin.com.tr>",
    "Ali Orhan Barç <obarc@kolin.com.tr>",
    "Cenk Erdoğan <cerdogan@kolin.com.tr>",
    "Gülşah Der <gder@kolin.com.tr>",
    "Davud Kerimov <dkerimov@kolin.com.tr>",
    "Perviz Memmedov <pmemmedov@kolin.com.tr>",
    "Azer Bayramlı <abayramli@kolin.com.tr>",
    "Gökçe Çolakoğlu <gcolakoglu@kolin.com.tr>",
    "Zafer Altay <zaltay@kolin.com.tr>",
    "Yusif Esedov <yesedov@kolin.com.tr>",
    "Mehemmedeli Hesenli <mhesenli@kolin.com.tr>",
    "Vüqar Agaverdiyev <vagaverdiyev@kolin.com.tr>",
    "Furkan Gökhan Karakaya <fgkarakaya@kolin.com.tr>",
    "Tugay Altuntaş <taltuntas@kolin.com.tr>",
    "Turgay Bal <tbal@kolin.com.tr>",
    "Ali Doğan Karakuş <adkarakus@kolin.com.tr>",
    "Aziz Yaşar İşidoğru <ayisidogru@kolin.com.tr>",
    "Bersis Kök <bkok@kolin.com.tr>",
    "Damla Yüceer <dyuceer@kolin.com.tr>",
    "Erdinç Bey <ebey@kolin.com.tr>",
    "Mehmet Özgün DEDEKARGINOĞLU <modedekarginoglu@kolin.com.tr>",
    "Mehmet Tevfik Çelikkol <mtcelikkol@kolin.com.tr>",
    "Mustafa Can ÜNVER <mcunver@kolin.com.tr>",
    "Serdar Osman Boz <sboz@kolin.com.tr>",
    "Anıl Uzun <auzun@kolin.com.tr>",
    "Nurettin Biçer <nbicer@kolin.com.tr>",
    "Yiğit Yücel <yyucel@kolin.com.tr>",
    "Atilla Gündüz <atilla.gunduz@kolin.com.tr>",
    "Ayşe KARA <akara@kolin.com.tr>",
    "Burak PAPİLA <bpapila@kolin.com.tr>",
    "Göker İnceoğlu <ginceoglu@kolin.com.tr>",
    "Ali İsazade <aisazade@kolin.com.tr>",
    "Fidan Quliyeva <fquliyeva@kolin.com.tr>",
    "Orhan Doğan <odogan@kolin.com.tr>",
    "Orxan Dursunov <odursunov@kolin.com.tr>",
]
CC_RECIPIENTS: List[str] = []

SENDER_FILTER_CONTAINS = "dccspp2@proyapimusavirlik.com"  # same for TRN+STQ

OUTLOOK_MAILITEM_CLASS = 43

# MAPI properties
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"

# RPC error code
RPC_ERR = -2147023174  # "The RPC server is unavailable."

UPDATED_RE = re.compile(
    r"\b(updated|update|corrected|correction|revised|revision|rev\.)\b|"
    r"\b(güncellendi|güncellenmiş|güncellenmistir|düzəldildi|duzeldildi|duzeltme)\b",
    re.IGNORECASE,
)

# STQ prefix trigger (as you requested)
STQ_PREFIX = "KLN-SPP2-STQ"

# Robust extraction patterns
# - STQ: capture from KLN-SPP2-STQ ... up to _Rxx (ignore trailing _Reply etc.)
STQ_ID_RE = re.compile(
    r"\b(" + re.escape(STQ_PREFIX) + r"[A-Z0-9\-_]*?_R\d{2})\b",
    re.IGNORECASE,
)

TRN_ID_RE = re.compile(r"\b(SPP2-PRO-KLN-TRN-\d{4})\b", re.IGNORECASE)


# =========================
# LOGGING
# =========================
def log(msg: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def is_rpc_error(e: Exception) -> bool:
    try:
        return isinstance(e, pywintypes.com_error) and e.args and e.args[0] == RPC_ERR
    except Exception:
        return False


def short_hash(s: str) -> str:
    s = (s or "").encode("utf-8", errors="ignore")
    return hashlib.sha1(s).hexdigest()[:10].upper()


# =========================
# PROFILES
# =========================
@dataclass(frozen=True)
class Profile:
    name: str
    exts: Set[str]
    oft: Path
    keep_incoming_subject: bool
    enable_updated: bool


PROFILES_BY_NAME = {
    "STQ": Profile(
        name="STQ",
        exts={".xlsx", ".xls", ".xlsm"},
        oft=Path(
            r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.2. STQ - Internal Sharing.oft"
        ),
        keep_incoming_subject=True,
        enable_updated=False,
    ),
    "TRN": Profile(
        name="TRN",
        exts={".pdf"},
        oft=Path(
            r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.1. SPP2-TRN_internal.oft"
        ),
        keep_incoming_subject=False,
        enable_updated=True,
    ),
}


# =========================
# SQLITE STATE
# =========================
def db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=30.0)
    con.execute("PRAGMA journal_mode=WAL;")
    con.execute("PRAGMA synchronous=NORMAL;")
    con.execute("PRAGMA temp_store=MEMORY;")
    con.execute("PRAGMA cache_size=-64000;")  # 64MB

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
        CREATE TABLE IF NOT EXISTS history (
            profile TEXT NOT NULL,
            doc_id   TEXT NOT NULL,
            last_sig TEXT,
            last_dt  TEXT,
            last_ts  TEXT NOT NULL,
            PRIMARY KEY(profile, doc_id)
        );
        """
    )
    con.execute("CREATE INDEX IF NOT EXISTS idx_processed_uid ON processed(uid);")
    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_history_lookup ON history(profile, doc_id);"
    )
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


def get_hist(con: sqlite3.Connection, profile: str, doc_id: str) -> Tuple[str, str]:
    row = con.execute(
        "SELECT last_sig, last_dt FROM history WHERE profile=? AND doc_id=? LIMIT 1",
        (profile, doc_id),
    ).fetchone()
    return (row[0], row[1]) if row else ("", "")


def upsert_hist(
    con: sqlite3.Connection, profile: str, doc_id: str, sig: str, dt_iso: str
) -> None:
    con.execute(
        """
        INSERT INTO history(profile, doc_id, last_sig, last_dt, last_ts)
        VALUES(?, ?, ?, ?, ?)
        ON CONFLICT(profile, doc_id) DO UPDATE SET
            last_sig=excluded.last_sig,
            last_dt=excluded.last_dt,
            last_ts=excluded.last_ts
        """,
        (profile, doc_id, sig, dt_iso, datetime.now().isoformat()),
    )


# =========================
# OUTLOOK CORE
# =========================
def get_outlook():
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        app = win32com.client.DispatchEx("Outlook.Application")

    ns = app.GetNamespace("MAPI")
    try:
        ns.Logon("", "", False, False)
    except Exception:
        pass

    _ = ns.Folders.Count
    return app, ns


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


def safe_dt(x) -> Optional[datetime]:
    try:
        if not x:
            return None
        return x.replace(tzinfo=None)
    except Exception:
        return None


def sender_smtp(mail) -> str:
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
    """
    Best-effort: set SendUsingAccount to the account matching MAILBOX_HINT.
    In many corp environments this avoids silent send failures on shared mailboxes.
    """
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
# FILTER + SNAPSHOT
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
# MATCHING (SUBJECT FIRST, ATTACHMENT SECOND)
# =========================
def extract_stq_id(text: str) -> str:
    """
    Extract best STQ doc_id:
    - If full pattern found -> return up to _Rxx
    - Else if prefix present -> fallback doc_id based on subject hash (still idempotent)
    """
    if not text:
        return ""
    m = STQ_ID_RE.search(text)
    if m:
        return m.group(1).upper()
    if STQ_PREFIX.lower() in text.lower():
        return f"{STQ_PREFIX}_AUTO_{short_hash(text)}"
    return ""


def extract_trn_id(text: str) -> str:
    if not text:
        return ""
    m = TRN_ID_RE.search(text)
    return m.group(1).upper() if m else ""


def find_doc_in_mail(mail, subject: str) -> Tuple[Optional[Profile], str]:
    subj = subject or ""

    # 1) TRN subject match
    trn = extract_trn_id(subj)
    if trn:
        return PROFILES_BY_NAME["TRN"], trn

    # 2) STQ subject match by prefix (your requirement)
    stq = extract_stq_id(subj)
    if stq:
        return PROFILES_BY_NAME["STQ"], stq

    # 3) Attachment filename fallback
    try:
        atts = getattr(mail, "Attachments", None)
        if not atts:
            return None, ""
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fn = str(getattr(att, "FileName", "") or "")
            del att

            trn2 = extract_trn_id(fn)
            if trn2:
                return PROFILES_BY_NAME["TRN"], trn2

            stq2 = extract_stq_id(fn)
            if stq2:
                return PROFILES_BY_NAME["STQ"], stq2

        del atts
    except Exception:
        pass

    return None, ""


# =========================
# ATTACHMENTS
# =========================
def save_attachments(mail, exts: Set[str], prefix: str) -> List[Path]:
    td = Path(tempfile.mkdtemp(prefix=prefix))
    out: List[Path] = []

    atts = mail.Attachments
    cnt = atts.Count
    for i in range(1, cnt + 1):
        att = atts.Item(i)
        fn = str(att.FileName or "")
        ext = Path(fn).suffix.lower()
        if ext in exts:
            safe = re.sub(r'[<>:"/\\|?*]', "_", fn)
            p = td / safe
            att.SaveAsFile(str(p))
            out.append(p)
        del att
    del atts
    return out


def sig_of(files: List[Path]) -> str:
    parts = []
    for p in files:
        try:
            parts.append(f"{p.name.lower()}:{p.stat().st_size}")
        except Exception:
            parts.append(f"{p.name.lower()}:?")
    return "|".join(sorted(parts))


def cleanup(files: List[Path], prefix: str) -> None:
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


_BODY_STQ = (
    "<div style='font-family:Bahnschrift,Calibri,Arial,sans-serif;font-size:11pt;'>"
    "<p>Sayın İlgililer,</p>"
    "<p>Müşavir tarafından <b>{date}</b> tarihinde Sitalçay 2 Üretim Tesisi kapsamında {tail}</p><br>"
    "</div>"
)

_BODY_TRN = (
    "<div style='font-family:Bahnschrift,Calibri,Arial,sans-serif;font-size:11pt;'>"
    "<p>Sayın İlgililer,</p>"
    "<p>Müşavir tarafından <b>{date}</b> tarihinde Sitalçay 2 Üretim Tesisi için "
    "paylaşılan transmittal № <b>{trn}</b> ekte sunulmuştur.</p><br></div>"
)

_BODY_TRN_UPDATED = (
    "<div style='font-family:Bahnschrift,Calibri,Arial,sans-serif;font-size:11pt;'>"
    "<p>Sayın İlgililer,</p>"
    "<p>Müşavir tarafından <b>{date}</b> tarihinde Sitalçay 2 Üretim Tesisi için "
    "paylaşılan <b>güncellenmiş</b> transmittal № <b>{trn}</b> ekte sunulmuştur.</p>"
    "{note}<br></div>"
)


def body_stq(dt: Optional[datetime], n: int) -> str:
    tail = (
        "paylaşılan STQ dosyası ekte sunulmuştur."
        if n == 1
        else "paylaşılan STQ dosyaları ekte sunulmuştur."
    )
    return _BODY_STQ.format(date=en_date(dt), tail=tail)


def body_trn(
    trn: str, dt: Optional[datetime], is_update: bool, prev_dt_iso: str
) -> str:
    if not is_update:
        return _BODY_TRN.format(date=en_date(dt), trn=trn)

    prev_str = en_date(datetime.fromisoformat(prev_dt_iso)) if prev_dt_iso else ""
    note = "<p><b>Not:</b> Bu e-posta, daha önce paylaşılan transmittalın <b>güncellenmiş</b> versiyonudur"
    if prev_str:
        note += f" (önceki paylaşım tarihi: <b>{prev_str}</b>)"
    note += ". Lütfen önceki versiyonu dikkate almayınız.</p>"

    return _BODY_TRN_UPDATED.format(date=en_date(dt), trn=trn, note=note)


def send_internal(
    app,
    ns,
    prof: Profile,
    incoming_subject: str,
    doc_id: str,
    sent_dt: Optional[datetime],
    files: List[Path],
    is_update: bool,
    prev_dt_iso: str,
):
    try:
        msg = (
            app.CreateItemFromTemplate(str(prof.oft))
            if prof.oft.exists()
            else app.CreateItem(0)
        )
    except Exception:
        msg = app.CreateItem(0)

    # best-effort account selection
    try_set_sending_account(ns, msg, MAILBOX_HINT)

    msg.To = "; ".join(TO_RECIPIENTS)
    msg.CC = "; ".join(CC_RECIPIENTS) if CC_RECIPIENTS else ""

    # Subject policy
    if prof.keep_incoming_subject:
        msg.Subject = (incoming_subject or "").strip()
    else:
        msg.Subject = (
            f"{doc_id} (UPDATED)" if (prof.enable_updated and is_update) else doc_id
        )

    intro = (
        body_stq(sent_dt, len(files))
        if prof.name == "STQ"
        else body_trn(doc_id, sent_dt, is_update, prev_dt_iso)
    )
    msg.HTMLBody = intro + (msg.HTMLBody or "")

    for p in files:
        if p.exists():
            msg.Attachments.Add(str(p))

    # save before send improves reliability on some setups
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
# MAIN LOOP
# =========================
def main():
    log("=== Proyapi Unified Distributor v5.2 started ===")
    log(
        f"Profiles => {list(PROFILES_BY_NAME.keys())} | WATCH='{WATCH_FOLDER}' | MAX_SCAN={MAX_SCAN} | LOOKBACK_DAYS={LOOKBACK_DAYS}"
    )
    log(f"SENDER_FILTER_CONTAINS => {SENDER_FILTER_CONTAINS}")
    log(f"SEND_MODE => {SEND_MODE}")

    con = db()
    app = ns = None
    loop = 0

    while True:
        loop += 1
        try:
            if app is None or ns is None:
                app, ns = get_outlook()
                log("✅ Outlook session attached.")

            root, folder = resolve_watch_folder(ns)
            store_id = getattr(folder, "StoreID", None)
            log(
                f"\n[Loop #{loop}] scanning... mailbox='{root.Name}' folder='{WATCH_FOLDER}' store_id={'OK' if store_id else 'NONE'}"
            )

            items = folder.Items
            cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
            items = restrict_unread(items, cutoff)

            try:
                items.Sort("[ReceivedTime]", True)  # newest first
            except Exception:
                pass

            entry_ids = snapshot_unread_entryids(items, MAX_SCAN)
            log(f"Snapshot unread: {len(entry_ids)}")

            processed_count = 0
            preview_left = 6

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
                        if (
                            SENDER_FILTER_CONTAINS
                            and SENDER_FILTER_CONTAINS.lower() not in (sender or "")
                        ):
                            break

                        subject = str(getattr(mail, "Subject", "") or "")

                        prof, doc_id = find_doc_in_mail(mail, subject)
                        if not prof or not doc_id:
                            if preview_left > 0:
                                log(
                                    f"DEBUG_SKIP => sender={sender} subject='{subject}' (no STQ/TRN match)"
                                )
                                preview_left -= 1
                            break

                        uid = internet_message_id(mail) or str(
                            getattr(mail, "EntryID", "") or ""
                        )
                        if uid and already_processed(con, uid):
                            break

                        if preview_left > 0:
                            log(
                                f"DEBUG_MATCH => [{prof.name}] doc_id='{doc_id}' sender={sender} subject='{subject}'"
                            )
                            preview_left -= 1

                        # PRE-LOCK (no resend loops)
                        mark_read(mail)
                        mark_processed(con, uid)
                        con.commit()

                        sent_dt = safe_dt(getattr(mail, "SentOn", None)) or received

                        prefix = f"proyapi_{prof.name.lower()}_"
                        files = save_attachments(mail, prof.exts, prefix)
                        if not files:
                            log(
                                f"⚠️ {prof.name} matched but no allowed attachments: subject='{subject}'"
                            )
                            cleanup(files, prefix)
                            break

                        sig = sig_of(files)
                        prev_sig, prev_dt_iso = get_hist(con, prof.name, doc_id)

                        is_update = False
                        if prof.enable_updated:
                            if prev_sig and prev_sig == sig:
                                log(f"⏭️ SKIP[{prof.name}] {doc_id} same signature")
                                cleanup(files, prefix)
                                break
                            is_update = bool(UPDATED_RE.search(subject or "")) or bool(
                                prev_sig
                            )

                        send_internal(
                            app=app,
                            ns=ns,
                            prof=prof,
                            incoming_subject=subject,
                            doc_id=doc_id,
                            sent_dt=sent_dt,
                            files=files,
                            is_update=is_update,
                            prev_dt_iso=prev_dt_iso,
                        )

                        # finalize
                        mark_read(mail)
                        upsert_hist(
                            con,
                            prof.name,
                            doc_id,
                            sig,
                            (sent_dt.isoformat() if sent_dt else ""),
                        )
                        con.commit()

                        cleanup(files, prefix)
                        processed_count += 1
                        log(
                            f"✅ DONE [{prof.name}] {doc_id} files={len(files)} update={is_update}"
                        )
                        break

                    except Exception as e:
                        if is_rpc_error(e):
                            log(f"⚠️ RPC dropped. Reconnecting... ({e})")
                            app = ns = None
                            gc.collect()

                            if tried_retry:
                                log(
                                    "❌ RPC retry already used for this item. Skipping."
                                )
                                break

                            # reconnect and retry once for this same eid
                            try:
                                app, ns = get_outlook()
                                tried_retry = True
                                log(
                                    "✅ Outlook session re-attached. Retrying item once."
                                )
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

                # end while True retry loop

            if processed_count == 0:
                log("No matching unread docs.")

            try:
                del items
                del folder
                del root
            except Exception:
                pass
            gc.collect()

        except Exception as e:
            if is_rpc_error(e):
                log(f"❌ RPC loop error. Reconnecting... ({e})")
                app = ns = None
                gc.collect()
            else:
                log(f"❌ Loop error: {e}")
                log(traceback.format_exc())

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
