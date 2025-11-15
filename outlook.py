# outlook_trn_save_attachments_codes_only_v2.py

import os
import re
from pathlib import Path
import win32com.client as win32

MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\TO PROYAPI\TRN"
MAX_ROWS = 200
SAVE_DIR = r"C:\Users\X\Desktop\new"

# şəkillər – skip
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}


def is_code_file(filename: str) -> bool:
    """
    Yalnız kodlu faylları saxla:
    - KLN- ilə başlayanlar (KLN-SPP2-..., KLN-SPP2-MAR-... və s.)
    """
    name, _ = os.path.splitext(filename)
    return name.upper().startswith("KLN-")


def clean_filename_keep_code_only(filename: str) -> str:
    """
    Fayl adından yalnız kod hissəsini saxlayır.

    Nümunələr:
    KLN-SPP2-MAR-WE-GN00-045_R00 Fire Alarm System Part-2 (MOXA).pdf
        -> KLN-SPP2-MAR-WE-GN00-045_R00.pdf

    KLN-SPP2-MES-CV-GN00-103_R01_METHOD.pdf
        -> KLN-SPP2-MES-CV-GN00-103_R01.pdf

    KLN-SPP2-STQ-AR-GN00-326_R00_Prokon_Proyapi_Reply.xlsx
        -> KLN-SPP2-STQ-AR-GN00-326_R00.xlsx

    KLN-PRO-SPP2-MOM-PM-037_20251105_engineer comments.docx
        -> KLN-PRO-SPP2-MOM-PM-037_20251105.docx
    """
    name, ext = os.path.splitext(filename)

    # 1) _R00, _R01 və s.
    m = re.search(r"_R\d{2}", name, flags=re.IGNORECASE)
    if m:
        code = name[: m.end()]
    else:
        # 2) -R00 tipli kodlar
        m = re.search(r"-R\d{2}", name, flags=re.IGNORECASE)
        if m:
            code = name[: m.end()]
        else:
            # 3) _YYYYMMDD (tarix)
            m = re.search(r"_\d{8}", name)
            if m:
                code = name[: m.end()]
            else:
                # fallback: ilk boşluğa qədər
                code = name.split(" ")[0]

    # Windows üçün təhlükəli simvolları təmizlə
    code = re.sub(r'[\\/:*?"<>|]', "_", code)

    return code + ext


def unique_path(base_dir: Path, filename: str) -> Path:
    """
    Eyni adlı fayl varsa, sonuna _1, _2 və s. əlavə edir.
    """
    path = base_dir / filename
    if not path.exists():
        return path

    stem, ext = os.path.splitext(filename)
    i = 1
    while True:
        candidate = base_dir / f"{stem}_{i}{ext}"
        if not candidate.exists():
            return candidate
        i += 1


def get_target_folder(ns, mailbox, subpath):
    folder = ns.Folders[mailbox]
    for part in subpath.split("\\"):
        if part:
            folder = folder.Folders[part]
    return folder


def main():
    save_path = Path(SAVE_DIR)
    save_path.mkdir(parents=True, exist_ok=True)

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    saved_total = 0
    skipped_no_code = 0
    skipped_image = 0

    for item in items:
        if getattr(item, "Class", None) != 43:
            continue

        for att in item.Attachments:
            raw = att.FileName
            ext = os.path.splitext(raw)[1].lower()

            # 1) şəkilləri at
            if ext in IMAGE_EXTS:
                skipped_image += 1
                continue

            # 2) KLN- ilə başlamırsa, ümumiyyətlə götürmə
            if not is_code_file(raw):
                print(f"SKIP (no code prefix): {raw}")
                skipped_no_code += 1
                continue

            # 3) yalnız kod hissəsini saxla
            clean_name = clean_filename_keep_code_only(raw)

            target = unique_path(save_path, clean_name)
            att.SaveAsFile(str(target))
            saved_total += 1
            print(f"Saved: {target.name}")

    print("\n-------------------------------")
    print(f"Saved total code files      : {saved_total}")
    print(f"Skipped (no code in name)   : {skipped_no_code}")
    print(f"Skipped (image attachments) : {skipped_image}")


if __name__ == "__main__":
    main()
