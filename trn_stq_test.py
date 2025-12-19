# -*- coding: utf-8 -*-
"""
proyapi_unified_distributor.py (ULTRA PREMIUM v4.0 - SHORT + FAST)

- One scan loop, profile-driven (TRN+STQ)
- SQLite state (dedupe + history) => crash-safe, no resend loops
- PRE-LOCK read + POST mark read (Exchange lag safe)
- TRN UPDATED + signature skip, STQ plural/singular body
"""

import os, re, time, sqlite3, tempfile, shutil, traceback
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Set

import win32com.client  # pywin32


# =========================
# CONFIG
# =========================
BASE_DIR = Path(r"C:\Users\husey\OneDrive\Desktop\Development\Work_Automation")
DB_PATH = BASE_DIR / "state_proyapi_unified.db"
LOG_PATH = BASE_DIR / "logs" / "proyapi_unified.log"

MAILBOX_HINT = "spp2dcc@kolin.com.tr"
WATCH_FOLDER = r"Inbox\From Proyapi"

POLL_SECONDS = 10  # test: 10, prod: 60-300
LOOKBACK_DAYS = 7
MAX_SCAN = 200

SEND_MODE = "display"  # "send" | "draft" | "display"
DISPLAY_MODAL = False
AUTO_SEND_AFTER_DISPLAY = True

# TEST recipients (only Cavid). Prod-da dəyişərsən.
TO_RECIPIENTS = ["Cavid Huseyn <chuseyn@kolin.com.tr>"]
CC_RECIPIENTS: List[str] = []

OUTLOOK_MAILITEM_CLASS = 43
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
UPDATED_RE = re.compile(
    r"\b(updated|update|corrected|correction|revised|revision|rev\.)\b|"
    r"\b(güncellendi|güncellenmiş|güncellenmistir|düzəldildi|duzeldildi|duzeltme)\b",
    re.IGNORECASE,
)


def log(msg: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# =========================
# PROFILES
# =========================
@dataclass(frozen=True)
class Profile:
    name: str
    sender_contains: str
    id_re: re.Pattern
    exts: Set[str]
    oft: Path
    keep_incoming_subject: bool
    enable_updated: bool


PROFILES = [
    Profile(
        name="STQ",
        sender_contains="chuseyn@kolin.com.tr",  # TEST. Prod: proyapi sender
        id_re=re.compile(
            r"\bKLN-SPP2-STQ-[A-Z]{2}-GN00-\d{3,4}_R\d{2}(?=_|\b|$)", re.I
        ),
        exts={".xlsx", ".xls", ".xlsm"},
        oft=Path(
            r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.2. STQ - Internal Sharing.oft"
        ),
        keep_incoming_subject=True,
        enable_updated=False,
    ),
    Profile(
        name="TRN",
        sender_contains="chuseyn@kolin.com.tr",
        id_re=re.compile(r"\bSPP2-PRO-KLN-TRN-\d{4}\b", re.I),
        exts={".pdf"},
        oft=Path(
            r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.1. SPP2-TRN_internal.oft"
        ),
        keep_incoming_subject=False,
        enable_updated=True,
    ),
]


# =========================
# SQLITE STATE (FAST + SAFE)
# =========================
def db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=30.0)
    con.execute("PRAGMA journal_mode=WAL;")
    con.execute("PRAGMA synchronous=NORMAL;")
    con.execute("PRAGMA cache_size=-64000;")  # 64MB cache
    con.execute("PRAGMA temp_store=MEMORY;")
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


def already_processed(con, uid: str) -> bool:
    if not uid:
        return False
    cur = con.execute("SELECT 1 FROM processed WHERE uid=? LIMIT 1", (uid,))
    return cur.fetchone() is not None


def mark_processed(con, uid: str):
    if not uid:
        return
    con.execute(
        "INSERT OR IGNORE INTO processed(uid, ts) VALUES(?, ?)",
        (uid, datetime.now().isoformat()),
    )


def get_hist(con, profile: str, doc_id: str):
    cur = con.execute(
        "SELECT last_sig, last_dt FROM history WHERE profile=? AND doc_id=? LIMIT 1",
        (profile, doc_id),
    )
    row = cur.fetchone()
    return (row[0], row[1]) if row else ("", "")


def upsert_hist(con, profile: str, doc_id: str, sig: str, dt_iso: str):
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
# OUTLOOK HELPERS
# =========================
def get_outlook():
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    _ = ns.Folders.Count
    return app, ns


def find_mailbox_root(ns, hint: str):
    hint = (hint or "").strip().lower()
    best = None
    for i in range(1, ns.Folders.Count + 1):
        f = ns.Folders.Item(i)
        name = str(getattr(f, "Name", "") or "")
        if name.strip().lower() == hint:
            return f
        if hint and hint in name.strip().lower():
            best = f
    return best or ns.Folders.Item(1)


def get_folder(root, path: str):
    f = root
    for p in path.split("\\"):
        f = f.Folders.Item(p)
    return f


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
            finally:
                del ex
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


def mark_read(mail):
    try:
        mail.UnRead = False
        mail.Save()
    except Exception as e:
        log(f"⚠️ mark_read failed: {e}")


def save_attachments(mail, exts: Set[str], prefix: str) -> List[Path]:
    td = Path(tempfile.mkdtemp(prefix=prefix))
    out = []
    atts = mail.Attachments
    count = atts.Count
    for i in range(1, count + 1):
        att = atts.Item(i)
        fn = str(att.FileName or "")
        ext = Path(fn).suffix.lower()
        if ext in exts:
            safe = re.sub(r'[<>:"/\\|?*]', "_", fn)
            p = td / safe
            att.SaveAsFile(str(p))
            out.append(p)
    return out


def sig_of(files: List[Path]) -> str:
    parts = [
        f"{p.name.lower()}:{p.stat().st_size if p.exists() else '?'}" for p in files
    ]
    return "|".join(sorted(parts))


def cleanup(files: List[Path], prefix: str):
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


_BODY_STQ_TEMPLATE = (
    "<div style='font-family:Bahnschrift,Calibri,Arial,sans-serif;font-size:11pt;'>"
    "<p>Sayın İlgililer,</p>"
    "<p>Müşavir tarafından <b>{date}</b> tarihinde Sitalçay 2 Üretim Tesisi kapsamında {tail}</p><br>"
    "</div>"
)


def body_stq(dt: Optional[datetime], n: int) -> str:
    tail = (
        "paylaşılan STQ dosyası ekte sunulmuştur."
        if n == 1
        else "paylaşılan STQ dosyaları ekte sunulmuştur."
    )
    return _BODY_STQ_TEMPLATE.format(date=en_date(dt), tail=tail)


_BODY_TRN_BASE = (
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


def body_trn(
    trn: str, dt: Optional[datetime], is_update: bool, prev_dt_iso: str
) -> str:
    date_str = en_date(dt)
    if is_update:
        prev_str = en_date(datetime.fromisoformat(prev_dt_iso)) if prev_dt_iso else ""
        note = "<p><b>Not:</b> Bu e-posta, daha önce paylaşılan transmittalın <b>güncellenmiş</b> versiyonudur"
        if prev_str:
            note += f" (önceki paylaşım tarihi: <b>{prev_str}</b>)"
        note += ". Lütfen önceki versiyonu dikkate almayınız.</p>"
        return _BODY_TRN_UPDATED.format(date=date_str, trn=trn, note=note)
    return _BODY_TRN_BASE.format(date=date_str, trn=trn)


def send_internal(
    app,
    profile: Profile,
    subject_in: str,
    doc_id: str,
    dt: Optional[datetime],
    files: List[Path],
    is_update: bool,
    prev_dt_iso: str,
):
    # create msg (OFT best)
    try:
        msg = (
            app.CreateItemFromTemplate(str(profile.oft))
            if profile.oft.exists()
            else app.CreateItem(0)
        )
    except Exception:
        msg = app.CreateItem(0)

    msg.To = "; ".join(TO_RECIPIENTS)
    msg.CC = "; ".join(CC_RECIPIENTS) if CC_RECIPIENTS else ""

    if profile.keep_incoming_subject:
        msg.Subject = (subject_in or "").strip()
    else:
        msg.Subject = (
            f"{doc_id} (UPDATED)" if (profile.enable_updated and is_update) else doc_id
        )

    intro = (
        body_stq(dt, len(files))
        if profile.name == "STQ"
        else body_trn(doc_id, dt, is_update, prev_dt_iso)
    )
    msg.HTMLBody = intro + (msg.HTMLBody or "")

    for p in files:
        if p.exists():
            msg.Attachments.Add(str(p))

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
# ONE-SCAN ENGINE
# =========================
def match_profile(
    subject: str, sender: str
) -> tuple[Optional[Profile], Optional[re.Match]]:
    s = sender.lower()
    for p in PROFILES:
        if p.sender_contains.lower() in s:
            m = p.id_re.search(subject or "")
            if m:
                return p, m
    return None, None


def extract_id(match: Optional[re.Match]) -> str:
    return (match.group(0) if match else "").upper()


def restrict_unread(items, cutoff: datetime):
    try:
        cut = cutoff.strftime("%m/%d/%Y %I:%M %p")
        return items.Restrict(f"[UnRead] = True AND [ReceivedTime] >= '{cut}'")
    except Exception:
        try:
            return items.Restrict("[UnRead] = True")
        except Exception:
            return items


def main():
    log("=== Proyapi Unified Distributor v4.0 started ===")
    con = db()
    app, ns = get_outlook()
    root = find_mailbox_root(ns, MAILBOX_HINT)
    folder = get_folder(root, WATCH_FOLDER)

    log(f"Watching: {root.Name} / {WATCH_FOLDER}")
    log(
        f"Mode={SEND_MODE} Poll={POLL_SECONDS}s Lookback={LOOKBACK_DAYS}d MaxScan={MAX_SCAN}"
    )
    log("Profiles: " + ", ".join([p.name for p in PROFILES]))

    loop = 0
    while True:
        loop += 1
        try:
            log(f"\n[Loop #{loop}] scanning unread...")
            items = folder.Items
            items.Sort("[ReceivedTime]", True)

            cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
            items = restrict_unread(items, cutoff)

            try:
                n = min(items.Count, MAX_SCAN)
            except Exception:
                n = MAX_SCAN

            found = 0
            preview_left = 3

            for i in range(1, n + 1):
                try:
                    mail = items.Item(i)
                except Exception:
                    continue

                if getattr(mail, "Class", None) != OUTLOOK_MAILITEM_CLASS:
                    continue
                if not bool(getattr(mail, "UnRead", False)):
                    continue

                received = getattr(mail, "ReceivedTime", None)
                received = received.replace(tzinfo=None) if received else None
                if received and received < cutoff:
                    break

                subject = str(getattr(mail, "Subject", "") or "")
                sender = sender_smtp(mail)

                prof, id_match = match_profile(subject, sender)
                if not prof:
                    continue

                doc_id = extract_id(id_match)
                if not doc_id:
                    continue

                uid = internet_message_id(mail) or str(
                    getattr(mail, "EntryID", "") or ""
                )
                if uid and already_processed(con, uid):
                    continue

                if preview_left > 0:
                    log(
                        f"DEBUG => [{prof.name}] uid={uid[:60]} sender={sender} subject={subject}"
                    )
                    preview_left -= 1

                # PRE-LOCK + DB COMMIT => no resend loop
                mark_read(mail)
                mark_processed(con, uid)
                con.commit()

                sent_dt = getattr(mail, "SentOn", None) or received
                sent_dt = sent_dt.replace(tzinfo=None) if sent_dt else None

                prefix = f"proyapi_{prof.name.lower()}_"
                files = save_attachments(mail, prof.exts, prefix)
                if not files:
                    log(f"⚠️ {prof.name} matched but no allowed attachments: {subject}")
                    continue

                sig = sig_of(files)
                prev_sig, prev_dt_iso = get_hist(con, prof.name, doc_id)

                is_update = False
                if prof.enable_updated:
                    # TRN anti-spam: same doc + same signature => skip EARLY
                    if prev_sig and prev_sig == sig:
                        log(f"⏭️ SKIP[{prof.name}] {doc_id} same signature")
                        cleanup(files, prefix)
                        continue
                    is_update = bool(UPDATED_RE.search(subject or "")) or bool(prev_sig)

                send_internal(
                    app, prof, subject, doc_id, sent_dt, files, is_update, prev_dt_iso
                )

                # POST mark read again (Exchange lag safe)
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
                found += 1
                log(
                    f"✅ DONE [{prof.name}] {doc_id} files={len(files)} update={is_update}"
                )

            if found == 0:
                log("No matching unread docs.")

        except Exception as e:
            log(f"❌ Loop error: {e}")
            log(traceback.format_exc())

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
