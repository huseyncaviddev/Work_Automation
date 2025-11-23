# outlook_trn_stq_auto_log_final.py

import os
import re
from pathlib import Path
from datetime import datetime, timedelta
import win32com.client as win32

# ---------------------------------
# KONFİQ
# ---------------------------------
MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\TO PROYAPI\TRN"

STQ_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ")
TRN_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN")

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}

# Neçə gün geriyə qədər mail yoxlanılsın
LOOKBACK_DAYS = 3


# ---------------------------------
# TRN – NÖVBƏTİ QOVLUQ CREATOR
# ---------------------------------
def get_next_trn_folder(root: Path) -> Path:
    """
    Yalnız SPP2-KLN-PRO-TRN-XXXX formatındakı qovluqları sayır.
    '001', 'test' və s. kimi yad qovluqları ignore edir.
    Hər run-da növbəti TRN qovluğunu yaradır.
    """
    if not root.exists():
        root.mkdir(parents=True, exist_ok=True)

    max_num = 0

    for d in root.iterdir():
        if not d.is_dir():
            continue

        m = re.match(r"^SPP2-KLN-PRO-TRN-(\d{4})$", d.name)
        if not m:
            continue  # 001 və s. SKIP

        num = int(m.group(1))
        if num > max_num:
            max_num = num

    next_num = max_num + 1
    folder_name = f"SPP2-KLN-PRO-TRN-{next_num:04d}"

    new_folder = root / folder_name
    new_folder.mkdir(parents=True, exist_ok=True)

    # Subfolder-lər
    for sub in ["1. main", "2. attachments", "3. docs"]:
        (new_folder / sub).mkdir(exist_ok=True)

    print(f"*** NEW TRN FOLDER CREATED → {new_folder} ***")
    return new_folder


# ---------------------------------
# STQ – növbəti index
# ---------------------------------
def get_next_stq_index(root: Path) -> int:
    """
    STQ qovluqlarının əvvəlindəki rəqəmə baxır:
    '336. KLN-SPP2-STQ-...' → 337
    """
    if not root.exists():
        root.mkdir(parents=True, exist_ok=True)

    max_num = 0
    for d in root.iterdir():
        if not d.is_dir():
            continue

        m = re.match(r"^(\d+)", d.name)
        if not m:
            continue

        num = int(m.group(1))
        if num > max_num:
            max_num = num

    return max_num + 1 if max_num > 0 else 1


# ---------------------------------
# Fayl helper-ləri
# ---------------------------------
def is_code_file(filename: str) -> bool:
    name, _ = os.path.splitext(filename)
    return name.upper().startswith("KLN-")


def clean_filename_keep_code_only(filename: str) -> str:
    """
    _R00/_R01-ə qədər saxlayır, sonrası silinir.
    """
    name, ext = os.path.splitext(filename)
    m = re.search(r"_R\d{2}", name, flags=re.IGNORECASE)
    if m:
        code = name[: m.end()]
    else:
        code = name.split(" ")[0]

    code = re.sub(r'[\\/:*?"<>|]', "_", code)
    return code + ext


def extract_stq_prefix(filename: str) -> str:
    """
    KLN-SPP2-STQ-EL-GN00-336_R00 → KLN-SPP2-STQ-EL-GN00
    """
    name, _ = os.path.splitext(filename)
    upper = name.upper()

    m = re.search(r"(KLN-SPP2-STQ-[A-Z0-9]+-[A-Z0-9]+)", upper)
    if m:
        base = m.group(1)
    else:
        base = re.split(r"_R\d{2}", upper)[0]

    base = re.sub(r"-\d{1,4}$", "", base)
    return base


# ---------------------------------
# Outlook helper-i
# ---------------------------------
def get_target_folder(ns, mailbox, subpath):
    folder = ns.Folders[mailbox]
    for part in subpath.split("\\"):
        if part:
            folder = folder.Folders[part]
    return folder


# ---------------------------------
# MAIN
# ---------------------------------
def main():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    next_stq_no = get_next_stq_index(STQ_ROOT)

    # Bu run üçün növbəti TRN qovluğunu 1 DƏFƏ yarat
    trn_folder = get_next_trn_folder(TRN_ROOT)
    # TRN faylları üçün artıq 3. docs istifadə edirik
    trn_docs = trn_folder / "3. docs"

    saved_stq = 0
    saved_trn = 0

    for item in items:
        # yalnız MailItem
        if getattr(item, "Class", None) != 43:
            continue

        recv_time = getattr(item, "ReceivedTime", None)

        # Outlook datetime timezone-li ola bilər → naive-a çeviririk
        if isinstance(recv_time, datetime):
            recv_time_naive = recv_time.replace(tzinfo=None)
        else:
            recv_time_naive = None

        if recv_time_naive and recv_time_naive < cutoff:
            print("Reached cutoff date. Stopping scan.")
            break

        subject = getattr(item, "Subject", "") or ""
        subject_upper = subject.upper()

        # Body oxumuruq – performans üçün
        is_stq_mail = ("STQ" in subject_upper)

        if not item.Attachments or item.Attachments.Count == 0:
            continue

        for att in item.Attachments:
            raw = att.FileName
            ext = os.path.splitext(raw)[1].lower()

            # 1) şəkil → skip
            if ext in IMAGE_EXTS:
                continue

            # 2) KLN- prefiksi yoxdursa → skip
            if not is_code_file(raw):
                continue

            # ---------- STQ CASE ----------
            if ext in EXCEL_EXTS and is_stq_mail:
                stq_prefix = extract_stq_prefix(raw)
                stq_code = f"{stq_prefix}-{next_stq_no}"
                folder_name = f"{next_stq_no}. {stq_code}"

                stq_folder = STQ_ROOT / folder_name
                stq_folder.mkdir(parents=True, exist_ok=True)

                new_filename = f"{stq_code}_R00{ext}"
                target = stq_folder / new_filename

                if not target.exists():
                    att.SaveAsFile(str(target))
                    saved_stq += 1

                next_stq_no += 1
                continue

            # ---------- TRN CASE ----------
            clean_name = clean_filename_keep_code_only(raw)
            target = trn_docs / clean_name   # ⬅ BURANI DƏYİŞDİRDİK

            if not target.exists():
                att.SaveAsFile(str(target))
                saved_trn += 1

    print("\nDone.")
    print("Saved STQ :", saved_stq)
    print("Saved TRN :", saved_trn)


if __name__ == "__main__":
    main()
