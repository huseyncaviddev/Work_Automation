# -*- coding: utf-8 -*-
"""
Proyapi TRN Incoming -> Internal Distribution Automation (ULTRA PREMIUM GOLD)

Fixes:
- Body date is dynamic from incoming mail (SentOn preferred, fallback ReceivedTime)
- Date format is EN: "22 Sep 2025"
- TRN number is BOLD (HTML)
- Keeps fast performance: Restrict(UnRead + ReceivedTime) + MAX_SCAN
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
from typing import List, Optional, Tuple

import win32com.client  # pip install pywin32


# =========================================================
# CONFIG
# =========================================================

BASE_DIR = Path(r"C:\Users\husey\OneDrive\Desktop\Development\Work_Automation")

STATE_PATH = BASE_DIR / "state_proyapi_trn_processed.json"
LOG_PATH = BASE_DIR / "logs" / "proyapi_trn_internal.log"

MAILBOX_NAME_HINT = "spp2dcc@kolin.com.tr"
WATCH_FOLDER_PATH = r"Inbox\From Proyapi"

# ✅ Proyapi sender filter
SENDER_EMAIL = "chuseyn@kolin.com.tr"
# SENDER_EMAIL = "dccspp2@proyapimusavirlik.com"

# ✅ OFT template (signature/layout stable)
OFT_TEMPLATE_PATH = Path(
    r"C:\Users\husey\OneDrive\Desktop\SPP2-OFT\4.1. SPP2-TRN_internal.oft"
)

# "oft" recommended
SIGNATURE_MODE = "oft"  # "oft" or "file"
SIGNATURE_NAME = None

POLL_SECONDS = 300  # 5 minutes
LOOKBACK_DAYS = 7
MAX_SCAN = 200

SEND_MODE = "display"
DISPLAY_MODAL = False

DEBUG_PREVIEW = True
DEBUG_PREVIEW_LIMIT = 5

RAW_TO_RECIPIENTS = [
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
    "Azizakhanim Maharramli <amaharramli@kolin.com.tr>",
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

MAX_STATE_IDS = 5000

TRN_SUBJECT_RE = re.compile(r"\bSPP2-PRO-KLN-TRN-\d{4}\b", re.IGNORECASE)
OUTLOOK_MAILITEM_CLASS = 43


# =========================================================
# DATA MODEL
# =========================================================
@dataclass
class TrnMailPayload:
    trn_no: str
    incoming_entry_id: str
    pdf_files: List[Path]
    mail_item: object
    sent_dt: Optional[datetime]  # ✅ dynamic date source


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


def load_state() -> dict:
    try:
        if STATE_PATH.exists():
            with open(STATE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        log(f"STATE read error: {e}")
    return {"processed_entry_ids": []}


def save_state(state: dict) -> None:
    ensure_parent(STATE_PATH)
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def prune_state_ids(ids_list: List[str]) -> List[str]:
    if len(ids_list) <= MAX_STATE_IDS:
        return ids_list
    return ids_list[-MAX_STATE_IDS:]


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


def extract_trn_no(subject: str) -> Optional[str]:
    if not subject:
        return None
    m = TRN_SUBJECT_RE.search(subject)
    return m.group(0).upper() if m else None


def format_en_date(dt_obj: Optional[datetime]) -> str:
    # EN date like: 22 Sep 2025
    dt_obj = dt_obj or datetime.now()
    return dt_obj.strftime("%d %b %Y")


def get_outlook():
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    return app, ns


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


def get_folder(root, folder_path: str):
    parts = folder_path.split("\\")
    f = root
    for p in parts:
        f = f.Folders.Item(p)
    return f


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
def save_pdf_attachments(mail_item) -> List[Path]:
    temp_dir = Path(tempfile.mkdtemp(prefix="proyapi_trn_"))
    pdf_paths: List[Path] = []

    atts = mail_item.Attachments
    for i in range(1, atts.Count + 1):
        att = atts.Item(i)
        filename = str(att.FileName or "")
        if filename.lower().endswith(".pdf"):
            safe_name = re.sub(r'[<>:"/\\|?*]', "_", filename)
            out_path = temp_dir / safe_name
            att.SaveAsFile(str(out_path))
            pdf_paths.append(out_path)

    return pdf_paths


def cleanup_pdf_temp(pdf_files: List[Path]) -> None:
    try:
        if not pdf_files:
            return
        parent = pdf_files[0].parent
        if parent.exists() and parent.name.startswith("proyapi_trn_"):
            shutil.rmtree(parent, ignore_errors=True)
    except Exception as e:
        log(f"⚠️ Could not delete temp PDF dir: {e}")


# =========================================================
# MAIL BUILD (✅ Dynamic date + EN format + BOLD TRN)
# =========================================================
def build_internal_body_html(trn_no: str, sent_dt: Optional[datetime]) -> str:
    date_str = format_en_date(sent_dt)

    # Bahnschrift yoxdursa fallback-lar işə düşəcək
    font_style = "font-family:Bahnschrift, Calibri, Arial, sans-serif; font-size:11pt;"

    return (
        f"<div style='{font_style}'>"
        "<p>Sayın İlgililer,</p>"
        f"<p>Müşavir tarafından <b>{date_str}</b> tarihinde Sitalçay 2 Üretim Tesisi için "
        f"paylaşılan transmittal № <b>{trn_no}</b> ekte bulunan dosyadaki gibidir.</p>"
        "<br>"
        "</div>"
    )


def prepend_intro_keep_existing_html(msg, intro_html: str) -> None:
    existing_html = msg.HTMLBody or ""
    msg.HTMLBody = intro_html + existing_html


def create_internal_mail(
    outlook_app, trn_no: str, sent_dt: Optional[datetime], pdf_files: List[Path]
) -> None:
    to_unique, removed = dedupe_recipients(RAW_TO_RECIPIENTS)
    if removed:
        log(f"Recipient duplicates removed ({len(removed)}): " + " | ".join(removed))

    if SIGNATURE_MODE == "oft" and OFT_TEMPLATE_PATH.exists():
        msg = outlook_app.CreateItemFromTemplate(str(OFT_TEMPLATE_PATH))
    else:
        msg = outlook_app.CreateItem(0)
        sig_html = (
            get_signature_html(SIGNATURE_NAME) if SIGNATURE_MODE == "file" else ""
        )
        if sig_html:
            msg.HTMLBody = sig_html

    msg.To = "; ".join(to_unique)
    msg.CC = "; ".join(CC_RECIPIENTS) if CC_RECIPIENTS else ""
    msg.Subject = trn_no

    intro_html = build_internal_body_html(trn_no, sent_dt)
    prepend_intro_keep_existing_html(msg, intro_html)

    for p in pdf_files:
        if p.exists():
            msg.Attachments.Add(str(p))

    if SEND_MODE == "send":
        msg.Send()
        log(f"✅ Sent internal mail: {trn_no} | PDFs={len(pdf_files)}")
    elif SEND_MODE == "draft":
        msg.Save()
        log(f"📝 Saved draft: {trn_no} | PDFs={len(pdf_files)}")
    else:
        msg.Display(DISPLAY_MODAL)
    log(f"👀 Displayed: {trn_no} | PDFs={len(pdf_files)} (modal={DISPLAY_MODAL})")

    # ✅ Display-dən sonra auto-send
    if not DISPLAY_MODAL:
        msg.Send()
        log(f"✅ Sent after display: {trn_no} | PDFs={len(pdf_files)}")


def mark_mail_as_read(mail_item) -> None:
    try:
        mail_item.UnRead = False
        mail_item.Save()
    except Exception as e:
        log(f"⚠️ Could not mark mail as read: {e}")


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


def scan_unread_trn_mails(folder, processed_ids: set) -> List[TrnMailPayload]:
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    items = try_restrict_items(items, cutoff)

    results: List[TrnMailPayload] = []
    checked = 0
    matched = 0
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
                continue

            entry_id = str(mail.EntryID)
            if entry_id in processed_ids:
                continue

            sender = get_sender_smtp_address(mail)
            subject = str(getattr(mail, "Subject", "") or "")
            received = normalize_outlook_dt(getattr(mail, "ReceivedTime", None))

            if DEBUG_PREVIEW and preview_left > 0:
                log(
                    f"DEBUG => UnRead={mail.UnRead} | Sender={sender} | Subject={subject}"
                )
                preview_left -= 1

            if received and received < cutoff:
                break

            if SENDER_EMAIL.lower() not in sender:
                continue

            trn_no = extract_trn_no(subject)
            if not trn_no:
                continue

            # ✅ dynamic date source: SentOn preferred, fallback ReceivedTime
            sent_dt = normalize_outlook_dt(getattr(mail, "SentOn", None)) or received

            pdfs = save_pdf_attachments(mail)
            if not pdfs:
                log(f"⚠️ TRN matched but no PDF found: {subject}")
                continue

            results.append(
                TrnMailPayload(
                    trn_no=trn_no,
                    incoming_entry_id=entry_id,
                    pdf_files=pdfs,
                    mail_item=mail,
                    sent_dt=sent_dt,
                )
            )
            matched += 1

        except Exception as e:
            log(f"Error while scanning item #{idx}: {e}")
            log(traceback.format_exc())

    log(f"Scan stats => checked:{checked} matched:{matched} (Unread-only)")
    return results


# =========================================================
# MAIN LOOP
# =========================================================
def main():
    log("=== Proyapi TRN Internal Distributor started ===")

    state = load_state()
    processed_ids = set(state.get("processed_entry_ids", []))

    outlook_app, ns = get_outlook()
    root = find_mailbox_root(ns, MAILBOX_NAME_HINT)
    folder = get_folder(root, WATCH_FOLDER_PATH)

    log(f"Watching: {root.Name} / {WATCH_FOLDER_PATH}")
    log(f"Sender filter: {SENDER_EMAIL}")
    log(
        f"Mode: {SEND_MODE} | Poll: {POLL_SECONDS}s | Lookback: {LOOKBACK_DAYS}d | MaxScan:{MAX_SCAN}"
    )
    log(
        f"Signature mode: {SIGNATURE_MODE} | OFT exists: {OFT_TEMPLATE_PATH.exists()} | Display modal: {DISPLAY_MODAL}"
    )
    log("Filter: ONLY UNREAD mails. Processed mail will be marked READ.")

    while True:
        try:
            payloads = scan_unread_trn_mails(folder, processed_ids)

            if not payloads:
                log(f"No matching unread TRN found. Sleeping {POLL_SECONDS}s...")
            else:
                for p in payloads:
                    log(
                        f"FOUND => {p.trn_no} | PDFs={len(p.pdf_files)} | Date={format_en_date(p.sent_dt)}"
                    )

                    try:
                        create_internal_mail(
                            outlook_app, p.trn_no, p.sent_dt, p.pdf_files
                        )
                        mark_mail_as_read(p.mail_item)

                        processed_ids.add(p.incoming_entry_id)
                        state["processed_entry_ids"] = prune_state_ids(
                            list(processed_ids)
                        )
                        save_state(state)

                        log(f"✅ DONE => {p.trn_no} (incoming marked READ)")

                    finally:
                        cleanup_pdf_temp(p.pdf_files)

        except Exception as e:
            log(f"Loop error: {e}")
            log(traceback.format_exc())

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
