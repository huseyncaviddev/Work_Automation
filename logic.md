# outlook_trn_stq_auto_log_v2.py

import os
import re
from pathlib import Path
import win32com.client as win32

MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\TO PROYAPI\TRN"

# LOG kök path-lər

STQ_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ")
TRN_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN")

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}

def is*code_file(filename: str) -> bool:
name, * = os.path.splitext(filename)
return name.upper().startswith("KLN-")

def clean_filename_keep_code_only(filename: str) -> str:
"""
\_R00 / \_R01-ə qədər saxlayır, sonrası silinir.
"""
name, ext = os.path.splitext(filename)

    m = re.search(r"_R\d{2}", name, flags=re.IGNORECASE)
    if m:
        code = name[: m.end()]
    else:
        code = name.split(" ")[0]

    code = re.sub(r'[\\/:*?"<>|]', "_", code)
    return code + ext

def extract*stq_prefix(filename: str) -> str:
"""
KLN-SPP2-STQ-EL-GN00-336_R00 -> KLN-SPP2-STQ-EL-GN00
"""
name, * = os.path.splitext(filename)
upper = name.upper()

    m = re.search(r"(KLN-SPP2-STQ-[A-Z0-9]+-[A-Z0-9]+)", upper)
    if m:
        base = m.group(1)
    else:
        base = re.split(r"_R\d{2}", upper)[0]

    base = re.sub(r"-\d{1,4}$", "", base)
    return base

def get_next_index(root: Path) -> int:
if not root.exists():
root.mkdir(parents=True, exist_ok=True)

    nums = []
    for d in root.iterdir():
        if not d.is_dir():
            continue
        m = re.match(r"^(\d+)", d.name)
        if m:
            nums.append(int(m.group(1)))

    return (max(nums) + 1) if nums else 1

def get_target_folder(ns, mailbox, subpath):
folder = ns.Folders[mailbox]
for part in subpath.split("\\"):
if part:
folder = folder.Folders[part]
return folder

def main():
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = get_target_folder(outlook, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # STQ üçün növbəti index
    next_stq_no = get_next_index(STQ_ROOT)

    # TRN üçün BU RUN-da yalnız *bir* yeni qovluq yaradılır
    next_trn_no = get_next_index(TRN_ROOT)
    trn_folder_name = f"{next_trn_no:03d}"
    trn_folder = TRN_ROOT / trn_folder_name
    trn_folder.mkdir(parents=True, exist_ok=True)
    print(f"TRN folder created for this run: {trn_folder}")

    saved_stq = 0
    saved_trn = 0
    skipped_image = 0
    skipped_no_code = 0

    for item in items:
        if getattr(item, "Class", None) != 43:
            continue

        subject = getattr(item, "Subject", "") or ""
        body = getattr(item, "Body", "") or ""
        mail_text_upper = (subject + " " + body).upper()

        is_stq_mail = ("STQ" in mail_text_upper)

        for att in item.Attachments:
            raw = att.FileName
            ext = os.path.splitext(raw)[1].lower()

            # 1) şəkil → skip
            if ext in IMAGE_EXTS:
                skipped_image += 1
                continue

            # 2) KLN- prefiksi yoxdursa → skip
            if not is_code_file(raw):
                print(f"SKIP (no KLN- prefix): {raw}")
                skipped_no_code += 1
                continue

            # 3) STQ CASE: excel + mail STQ-dursa → hər fayla ayrıca folder
            if ext in EXCEL_EXTS and is_stq_mail:
                stq_prefix = extract_stq_prefix(raw)   # KLN-SPP2-STQ-EL-GN00
                stq_code = f"{stq_prefix}-{next_stq_no}"  # KLN-SPP2-STQ-EL-GN00-336
                folder_name = f"{next_stq_no}. {stq_code}"

                stq_folder = STQ_ROOT / folder_name
                stq_folder.mkdir(parents=True, exist_ok=True)

                new_filename = f"{stq_code}_R00{ext}"
                target = stq_folder / new_filename

                if target.exists():
                    print(f"SKIP STQ (already exists): {target}")
                else:
                    att.SaveAsFile(str(target))
                    saved_stq += 1
                    print(f"Saved STQ #{next_stq_no}: {target}")

                next_stq_no += 1
                continue

            # 4) TRN CASE: qalan bütün kodlu fayllar → BU RUN-un TRN qovluğuna
            clean_name = clean_filename_keep_code_only(raw)
            target = trn_folder / clean_name

            if target.exists():
                print(f"SKIP TRN (already exists): {target}")
            else:
                att.SaveAsFile(str(target))
                saved_trn += 1
                print(f"Saved TRN: {target}")

    print("\n----------------------------")
    print(f"Saved STQ files : {saved_stq}")
    print(f"Saved TRN files : {saved_trn}")
    print(f"Skipped images  : {skipped_image}")
    print(f"Skipped no-code : {skipped_no_code}")

if **name** == "**main**":
main()
