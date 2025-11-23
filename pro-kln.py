# outlook_proyapi_incoming_trn_auto.py
# Proyapi'den gelen incoming TRN'leri avtomatik PROYAPI QA-QC'den
# Log\2. Incoming\1. TRN altına kopyalayıb, içini strukturlaşdırır.

import os
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta

import win32com.client as win32

# ---------------------------------
# KONFİQ
# ---------------------------------
MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\From Proyapi"  # Proyapi mails folder

LOOKBACK_DAYS = 3  # neçə gün geriyə baxaq

# QA-QC-də PROYAPI TRN source root
PROYAPI_TRN_SOURCE_ROOT = Path(
    r"\\10.10.8.253\DataServer\QA-QC\QA-QC Proyapi\SPP2\99_Temporary\DCC\PRO-KLN-TRN"
)

# Incoming TRN root (Log)
INCOMING_TRN_ROOT = Path(
    r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN"
)


# ---------------------------------
# Outlook helper
# ---------------------------------
def get_target_folder(ns, mailbox, subpath):
    folder = ns.Folders[mailbox]
    for part in subpath.split("\\"):
        if part:
            folder = folder.Folders[part]
    return folder


# ---------------------------------
# TRN processing helper-lər
# ---------------------------------
def notify_missing_trn(trn_code: str):
    """
    Burada əslində sənin WP application üçün hook olacaq.
    Hal-hazırda sadəcə log yazır.
    """
    msg = (
        f"[MISSING TRN] {trn_code} QA-QC Proyapi folderində tapılmadı. "
        f"Səidə xanıma WP ilə xəbər verilməlidir."
    )
    print(msg)
    # TODO: Buraya WP application üçün subprocess və ya API call əlavə edə bilərsən.


def ensure_subfolders(base_folder: Path):
    """
    1. main / 2. attachments / 3. docs qovluqlarını yaradır.
    """
    for name in ["1. main", "2. attachments", "3. docs"]:
        p = base_folder / name
        p.mkdir(exist_ok=True)


def move_pdf_and_zip(base_folder: Path, trn_code: str):
    """
    PDF → 1. main
    ZIP extract → 3. docs
    ZIP → 2. attachments
    """
    main_folder = base_folder / "1. main"
    attachments_folder = base_folder / "2. attachments"
    docs_folder = base_folder / "3. docs"

    pdf_path = base_folder / f"{trn_code}.pdf"
    zip_path = base_folder / f"{trn_code}.zip"
    extracted_folder = base_folder / trn_code  # zip çıxanda yaranan qovluq

    # 1) PDF → 1. main (move)
    if pdf_path.exists():
        target_pdf = main_folder / pdf_path.name
        shutil.move(str(pdf_path), str(target_pdf))
        print(f"[TRN] PDF moved → {target_pdf}")
    else:
        print(f"[WARN] PDF not found: {pdf_path}")

    # 2) ZIP-i extract et
    if zip_path.exists():
        try:
            shutil.unpack_archive(str(zip_path), str(base_folder))
            print(f"[TRN] ZIP extracted → {base_folder}")

            # 3) Extracted folderi 3. docs-a move et
            if extracted_folder.exists():
                target_extracted = docs_folder / extracted_folder.name
                if target_extracted.exists():
                    shutil.rmtree(target_extracted)
                shutil.move(str(extracted_folder), str(target_extracted))
                print(f"[TRN] Extracted folder moved → {target_extracted}")
            else:
                print(f"[WARN] Extracted folder not found: {extracted_folder}")

            # 4) ZIP-i 2. attachments-a move et
            target_zip = attachments_folder / zip_path.name
            shutil.move(str(zip_path), str(target_zip))
            print(f"[TRN] ZIP moved → {target_zip}")

        except Exception as e:
            print(f"[ERROR] Error while extracting ZIP {zip_path}: {e}")
    else:
        print(f"[WARN] ZIP not found: {zip_path}")


def copy_all_pdfs_to_docs_root(docs_folder: Path):
    """
    3. docs altındakı bütün pdf-ləri rekursiv tapır,
    3. docs root-a kopyalayır. Eyni ad varsa → _copy əlavə edir.
    """
    if not docs_folder.exists():
        return

    for pdf in docs_folder.rglob("*.pdf"):
        # Artıq root 3. docs-dadırsa → keç
        if pdf.parent == docs_folder:
            continue

        target = docs_folder / pdf.name
        if target.exists():
            stem = target.stem
            ext = target.suffix
            new_name = f"{stem}_copy{ext}"
            target = docs_folder / new_name

        shutil.copy2(str(pdf), str(target))
        print(f"[TRN] PDF copied to docs root → {target}")


def rename_r_dash_to_r_underscore(docs_folder: Path):
    """
    3. docs içindəki bütün pdf adlarında '-R' → '_R'.
    """
    if not docs_folder.exists():
        return

    for pdf in docs_folder.glob("*.pdf"):
        old_name = pdf.name
        new_name = old_name.replace("-R", "_R")
        if new_name != old_name:
            target = docs_folder / new_name
            if not target.exists():
                pdf.rename(target)
                print(f"[RENAME] {old_name} → {new_name}")
            else:
                print(f"[RENAME-SKIP] Target exists: {target}")


def title_cleanup_pattern(docs_folder: Path):
    """
    Fayl adlarını KLN-SPP2-XXX-..._R00 formatına salır:
    Pattern: (_Rdd)_.* → yalnız _Rdd saxlanır.
    Məs:
      KLN-SPP2-FRM-MC-GN00-137_R00_Prokon_Reply → KLN-SPP2-FRM-MC-GN00-137_R00.pdf
    """
    if not docs_folder.exists():
        return

    for pdf in docs_folder.glob("*.pdf"):
        old_name = pdf.name
        base = pdf.stem  # uzantısız ad
        # _Rdd-dən sonrakı hər şeyi sil
        new_base = re.sub(r"(_R\d{2})_.*", r"\1", base)
        new_name = new_base + ".pdf"
        target = docs_folder / new_name

        if new_name == old_name:
            continue

        if target.exists() and target != pdf:
            print(f"[TITLE-SKIP] {new_name} already exists, skipping {old_name}")
            continue

        pdf.rename(target)
        print(f"[TITLE] {old_name} → {new_name}")


def convert_main_pdf_to_docx(base_folder: Path, trn_code: str):
    """
    1. main içindəki SPP2-PRO-KLN-TRN-XXXX.pdf → eyni addan .docx
    """
    main_folder = base_folder / "1. main"
    pdf_path = main_folder / f"{trn_code}.pdf"
    if not pdf_path.exists():
        print(f"[DOCX] Main PDF not found, skipping Word conversion: {pdf_path}")
        return

    docx_path = pdf_path.with_suffix(".docx")

    print(f"[DOCX] Converting to Word: {pdf_path} → {docx_path}")

    word = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(str(pdf_path))
        wdFormatXMLDocument = 16  # .docx
        doc.SaveAs(str(docx_path), FileFormat=wdFormatXMLDocument)
        doc.Close()
        print(f"[DOCX] Saved: {docx_path}")
    except Exception as e:
        print(f"[ERROR] Word conversion failed: {e}")
    finally:
        if word is not None:
            word.Quit()


def process_single_trn(trn_code: str):
    """
    Bir konkret TRN üçün:
      1) QA-QC Proyapi-də folderi tap
      2) Incoming 1. TRN altına copy
      3) İçində 1.main / 2.attachments / 3.docs yarat
      4) PDF və ZIP ilə əməliyyatları icra et
      5) 3.docs içində flatten + rename
      6) Main PDF-i DOCX-ə çevir
    """
    source_folder = PROYAPI_TRN_SOURCE_ROOT / trn_code
    if not source_folder.exists():
        print(f"[TRN] Source folder NOT FOUND: {source_folder}")
        notify_missing_trn(trn_code)
        return

    # Incoming target folder
    if not INCOMING_TRN_ROOT.exists():
        INCOMING_TRN_ROOT.mkdir(parents=True, exist_ok=True)

    target_folder = INCOMING_TRN_ROOT / trn_code

    if not target_folder.exists():
        shutil.copytree(source_folder, target_folder)
        print(f"[TRN] Copied source → {target_folder}")
    else:
        print(f"[TRN] Incoming folder already exists, will reuse: {target_folder}")

    # 1.main / 2.attachments / 3.docs
    ensure_subfolders(target_folder)

    # PDF / ZIP operation
    move_pdf_and_zip(target_folder, trn_code)

    docs_folder = target_folder / "3. docs"

    # Bütün PDF-ləri 3.docs root-a copy
    copy_all_pdfs_to_docs_root(docs_folder)

    # -R → _R
    rename_r_dash_to_r_underscore(docs_folder)

    # Title cleanup (pattern)
    title_cleanup_pattern(docs_folder)

    # Main PDF → DOCX
    convert_main_pdf_to_docx(target_folder, trn_code)


# ---------------------------------
# MAIN
# ---------------------------------
def main():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    processed_trns = set()

    for item in items:
        if getattr(item, "Class", None) != 43:  # yalnız MailItem
            continue

        recv_time = getattr(item, "ReceivedTime", None)
        if isinstance(recv_time, datetime):
            recv_time_naive = recv_time.replace(tzinfo=None)
        else:
            recv_time_naive = None

        if recv_time_naive and recv_time_naive < cutoff:
            print("Reached cutoff date. Stopping scan.")
            break

        # SenderEmailAddress-a toxunmuruq → popup yoxdur
        subject = getattr(item, "Subject", "") or ""
        subject = subject.strip()

        # Subject: SPP2-PRO-KLN-TRN-0489
        m = re.match(r"^(SPP2-PRO-KLN-TRN-\d{4})$", subject)
        if not m:
            continue

        trn_code = m.group(1)
        if trn_code in processed_trns:
            continue

        print(f"\n[MAIL] Found PROYAPI TRN mail: {trn_code}")
        process_single_trn(trn_code)
        processed_trns.add(trn_code)

    print("\nDONE.")
    print("Processed TRNs:", ", ".join(sorted(processed_trns)) if processed_trns else "none")


if __name__ == "__main__":
    main()
