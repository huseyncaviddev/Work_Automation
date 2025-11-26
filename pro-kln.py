# outlook_proyapi_incoming_auto.py
# Proyapi'dən gələn TRN, STQ, LET maillərini avtomatik emal edir.

import os
import re
import shutil
import subprocess
from pathlib import Path

import win32com.client as win32

# ---------------------------------------------------------
# KONFİQ
# ---------------------------------------------------------
MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\From Proyapi"

# QA-QC-də PROYAPI TRN source root
PROYAPI_TRN_SOURCE_ROOT = Path(
    r"\\10.10.8.253\DataServer\QA-QC\QA-QC Proyapi\SPP2\99_Temporary\DCC\PRO-KLN-TRN"
)

# Incoming TRN root (Log)
INCOMING_TRN_ROOT = Path(
    r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\2. Incoming\1. TRN"
)

# Outgoing STQ root (Log\1. Outgoing\3. STQ)
STQ_OUTGOING_ROOT = Path(
    r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ"
)

# Incoming LET root (Google Drive)
LET_INCOMING_ROOT = Path(
    r"G:\My Drive\4-S1 ve S2 Ortak Dökümanlar\03-SPP LETTERS\SPP2-LET\1. KLN-PRO\02-Incoming"
)

IMG_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


# ---------------------------------------------------------
# OUTLOOK FOLDER
# ---------------------------------------------------------
def get_target_folder(ns, mailbox, subpath):
    folder = ns.Folders[mailbox]
    for p in subpath.split("\\"):
        if p:
            folder = folder.Folders[p]
    return folder


# ---------------------------------------------------------
# SIMPLE PROGRESS BAR
# ---------------------------------------------------------
def print_progress(current: int, total: int, label: str = ""):
    """
    Konsolda sadə progress bar:
    [_ _ _ _ _ _ _ _ _ _ _ _ _ _     ] 60% Processing mailbox...
    """
    if total <= 0:
        return

    percent = int(current * 100 / total)
    bar_len = 30
    filled = int(bar_len * percent / 100)

    bar = "_" * filled + " " * (bar_len - filled)
    suffix = f" {label}" if label else ""
    print(f"\r[{bar}] {percent:3d}%{suffix}", end="", flush=True)


# ---------------------------------------------------------
# WINRAR EXTRACT
# ---------------------------------------------------------
def extract_with_winrar(zip_path: Path, out_dir: Path):
    """
    ZIP faylını WinRAR ilə extract edir.
    WinRAR tapılmazsa → shutil.unpack_archive fallback.
    """
    winrar_paths = [
        r"C:\Program Files\WinRAR\WinRAR.exe",
        r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
    ]

    winrar_exe = None
    for p in winrar_paths:
        if Path(p).exists():
            winrar_exe = p
            break

    if not winrar_exe:
        print("\n[WARN] WinRAR tapılmadı, shutil.unpack_archive istifadə olunur.")
        shutil.unpack_archive(str(zip_path), str(out_dir))
        return

    cmd = [
        winrar_exe,
        "x",          # extract
        "-o+",        # overwrite all without prompt
        str(zip_path),
        str(out_dir) + "\\",
    ]

    print(f"\n[TRN] Using WinRAR → {zip_path}")
    try:
        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        print(f"\n[WARN] WinRAR extract alınmadı ({e}), shutil.unpack_archive istifadə olunur.")
        shutil.unpack_archive(str(zip_path), str(out_dir))


# ---------------------------------------------------------
# TRN FUNKSİYALARI
# ---------------------------------------------------------
def notify_missing_trn(trn_code):
    msg = (
        f"[MISSING TRN] {trn_code} QA-QC Proyapi folderində tapılmadı. "
        f"Səidə xanıma WP ilə xəbər verilməlidir."
    )
    print(msg)


def ensure_trn_subfolders(base_folder: Path):
    """1. main / 2. attachments / 3. docs qovluqlarını yaradır."""
    for name in ["1. main", "2. attachments", "3. docs"]:
        (base_folder / name).mkdir(exist_ok=True)


def is_trn_already_processed(base_folder: Path, trn_code: str) -> bool:
    """
    TRN artıq full işlənibsə → True (docx + 3. docs var).
    Bu halda sonrakı TRN əməliyyatlarını SKIP edirik (sürət üçün).
    """
    main_folder = base_folder / "1. main"
    docs_folder = base_folder / "3. docs"

    docx_path = main_folder / f"{trn_code}.docx"
    if docx_path.exists() and docs_folder.exists():
        print(f"[TRN] Already processed → {trn_code}, heavy ops skipped.")
        return True
    return False


def cleanup_trn_base_folder(base_folder: Path, trn_code: str):
    """
    Base folder-də qalan artıqları təmizləyir:
      - SPP2-PRO-KLN-TRN-0493 qovluğu varsa → 3. docs altına daşı
      - SPP2-PRO-KLN-TRN-0493.zip varsa → 2. attachments altına daşı
    """
    docs_folder = base_folder / "3. docs"
    attachments_folder = base_folder / "2. attachments"
    ensure_trn_subfolders(base_folder)

    raw_folder = base_folder / trn_code
    if raw_folder.exists() and raw_folder.is_dir():
        target_extracted = docs_folder / raw_folder.name
        if target_extracted.exists():
            # artıq orada varsa, qalıq raw_folder-i silək
            shutil.rmtree(raw_folder)
            print(f"[TRN-CLEAN] Extra raw folder removed → {raw_folder}")
        else:
            shutil.move(str(raw_folder), str(target_extracted))
            print(f"[TRN-CLEAN] Raw folder moved → {target_extracted}")

    raw_zip = base_folder / f"{trn_code}.zip"
    if raw_zip.exists() and raw_zip.is_file():
        target_zip = attachments_folder / raw_zip.name
        if target_zip.exists():
            raw_zip.unlink()
            print(f"[TRN-CLEAN] Extra ZIP removed → {raw_zip}")
        else:
            shutil.move(str(raw_zip), str(target_zip))
            print(f"[TRN-CLEAN] ZIP moved → {target_zip}")


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

    # PDF → 1. main (move)
    if pdf_path.exists():
        target_pdf = main_folder / pdf_path.name
        if not target_pdf.exists():
            shutil.move(str(pdf_path), str(target_pdf))
            print(f"[TRN] PDF moved → {target_pdf}")
        else:
            print(f"[TRN] PDF already in main → {target_pdf}")
    else:
        print(f"[WARN] PDF not found: {pdf_path}")

    # ZIP → extract
    if zip_path.exists():
        try:
            extract_with_winrar(zip_path, base_folder)
            print(f"[TRN] ZIP extracted → {base_folder}")

            # Extracted folderi 3. docs-a move et
            if extracted_folder.exists():
                target_extracted = docs_folder / extracted_folder.name
                if target_extracted.exists():
                    shutil.rmtree(target_extracted)
                shutil.move(str(extracted_folder), str(target_extracted))
                print(f"[TRN] Extracted folder moved → {target_extracted}")
            else:
                print(f"[WARN] Extracted folder not found: {extracted_folder}")

            # ZIP → 2. attachments
            target_zip = attachments_folder / zip_path.name
            if not target_zip.exists():
                shutil.move(str(zip_path), str(target_zip))
                print(f"[TRN] ZIP moved → {target_zip}")
            else:
                print(f"[TRN] ZIP already in attachments → {target_zip}")

        except Exception as e:
            print(f"[ERROR] Error while extracting ZIP {zip_path}: {e}")
    else:
        print(f"[WARN] ZIP not found: {zip_path}")

    # Əlavə safety cleanup – screenshotdakı problemi öldürmək üçün
    cleanup_trn_base_folder(base_folder, trn_code)


def copy_all_pdfs_to_docs_root(docs_folder: Path):
    """
    3. docs altındakı bütün pdf-ləri rekursiv tapır,
    3. docs root-a kopyalayır. Eyni ad varsa → SKIP.
    """
    if not docs_folder.exists():
        return

    for pdf in docs_folder.rglob("*.pdf"):
        if pdf.parent == docs_folder:
            continue

        target = docs_folder / pdf.name
        if target.exists():
            print(f"[TRN] PDF already in docs root, skip → {target}")
            continue

        shutil.copy2(str(pdf), str(target))
        print(f"[TRN] PDF copied to docs root → {target}")


def rename_r_dash_to_r_underscore(docs_folder: Path):
    """3. docs içindəki bütün pdf adlarında '-R' → '_R'."""
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
        base = pdf.stem
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
    if docx_path.exists():
        print(f"[DOCX] Already exists, skip → {docx_path}")
        return

    print(f"[DOCX] Converting to Word: {pdf_path} → {docx_path}")

    word = None    # Word açmaq ağır əməliyyatdır – yalnız ehtiyac olanda
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
    1) QA-QC Proyapi-də folderi tap
    2) Incoming 1. TRN altına copy
    3) Already processed? → skip heavy ops
    4) 1.main / 2.attachments / 3.docs yarat
    5) PDF & ZIP əməliyyatları
    6) 3.docs flatten + rename
    7) Main PDF → DOCX
    """
    source_folder = PROYAPI_TRN_SOURCE_ROOT / trn_code
    if not source_folder.exists():
        print(f"[TRN] Source folder NOT FOUND: {source_folder}")
        notify_missing_trn(trn_code)
        return False

    if not INCOMING_TRN_ROOT.exists():
        INCOMING_TRN_ROOT.mkdir(parents=True, exist_ok=True)

    target_folder = INCOMING_TRN_ROOT / trn_code

    if not target_folder.exists():
        shutil.copytree(source_folder, target_folder)
        print(f"[TRN] Copied source → {target_folder}")
    else:
        print(f"[TRN] Incoming folder already exists, will reuse: {target_folder}")

    ensure_trn_subfolders(target_folder)

    # FULL işlənmişdirsə → sonik skip
    if is_trn_already_processed(target_folder, trn_code):
        # Yenə də əmin olmaq üçün cleanup (əgər köhnə run artıqları qalıbsa)
        cleanup_trn_base_folder(target_folder, trn_code)
        return True

    move_pdf_and_zip(target_folder, trn_code)

    docs_folder = target_folder / "3. docs"
    copy_all_pdfs_to_docs_root(docs_folder)
    rename_r_dash_to_r_underscore(docs_folder)
    title_cleanup_pattern(docs_folder)
    convert_main_pdf_to_docx(target_folder, trn_code)

    return True


# ---------------------------------------------------------
# STQ FUNKSİYALARI  (multi-STQ, idempotent)
# ---------------------------------------------------------
def find_stq_target_folder(base_code: str):
    """
    STQ qovluqları:
       309. KLN-SPP2-STQ-WE-GN00-309
    base_code = KLN-SPP2-STQ-WE-GN00-309
    """
    if not STQ_OUTGOING_ROOT.exists():
        print(f"[STQ] ROOT YOXDUR → {STQ_OUTGOING_ROOT}")
        return None

    for child in STQ_OUTGOING_ROOT.iterdir():
        if child.is_dir() and child.name.endswith(base_code):
            return child

    for p in STQ_OUTGOING_ROOT.rglob("*"):
        if p.is_dir() and p.name.endswith(base_code):
            return p

    print(f"[STQ] QOVLUQ TAPILMADI → {base_code}")
    return None


def process_single_stq_mail(mail_item, subject: str) -> int:
    """
    Hər attachment üçün ayrıca işləyir.
    STQ kodu və rev-i əsasən attachment adından çıxardır:

      KLN-SPP2-STQ-AR-GN00-282_R00_Proyapi_Reply.xlsx
      KLN-SPP2-STQ-CV-GN00-211_R00_Prokon_Reply.xlsx

    Eyni mail içində birdən çox STQ ola bilər – hamısını götürür.
    Eyni fayl adı artıq mövcuddursa → SKIP (yenidən kopyalamır).
    """
    atts = mail_item.Attachments
    counters = {}  # "base_Rxx" -> idx
    saved_count = 0

    for att in atts:
        fname = att.FileName or ""
        ext = Path(fname).suffix.lower()

        if ext in IMG_EXTS or fname.lower().startswith("image00"):
            print(f"[STQ] Skip image → {fname}")
            continue

        base_code = None
        rev = None

        # 1) Əsas qaynaq: fayl adı
        m_fname = re.match(
            r"^(KLN-SPP2-STQ-[A-Za-z0-9-]+)_R(\d{2})[_.-].*$", fname
        )
        if m_fname:
            base_code = m_fname.group(1)
            rev = m_fname.group(2)
        else:
            # 2) Fallback: subject-dən ilk STQ kodu
            m_subj = re.search(
                r"(KLN-SPP2-STQ-[A-Za-z0-9-]+)(?:_R(\d{2}))?", subject
            )
            if m_subj:
                base_code = m_subj.group(1)
                rev = m_subj.group(2) or "00"

        if not base_code:
            print(f"[STQ] Attachment skipped, no STQ code found: {fname}")
            continue

        key = f"{base_code}_R{rev}"
        idx = counters.get(key, 0) + 1
        counters[key] = idx

        target = find_stq_target_folder(base_code)
        if target is None:
            continue

        if idx == 1:
            new_name = f"{base_code}_R{rev} Reply{ext}"
        else:
            new_name = f"{base_code}_R{rev} Reply_{idx}{ext}"

        save_path = target / new_name
        if save_path.exists():
            print(f"[STQ] Exists, skipping: {save_path}")
            continue

        att.SaveAsFile(str(save_path))
        print(f"[STQ] Saved: {save_path}")
        saved_count += 1

    return saved_count


# ---------------------------------------------------------
# LET FUNKSİYALARI  (DƏYİŞƏN HİSSƏ)
# ---------------------------------------------------------
def get_let_folder_by_code(root: Path, let_code: str) -> Path:
    """
    Məs: let_code = 'SPP2-PRO-KLN-LET-0022'
    Əgər bu qovluq varsa → onu istifadə edir,
    yoxdursa → eyni adla yeni qovluq yaradır.
    """
    base = root / let_code
    base.mkdir(parents=True, exist_ok=True)
    return base


def process_single_let_mail(mail_item, subject: str, let_code: str) -> bool:
    """
    LET mail:
      - Qovluq adı birbaşa subject-dən çıxan LET kodu olur
        (məs: SPP2-PRO-KLN-LET-0022)
      - Qovluq artıq varsa, yenisi yaradılmır, reuse olunur.
    """
    base = get_let_folder_by_code(LET_INCOMING_ROOT, let_code)
    letter_f = base / "1. letter"
    docs_f = base / "2. docs"

    letter_f.mkdir(parents=True, exist_ok=True)
    docs_f.mkdir(exist_ok=True)

    print(f"[LET] {let_code}")

    atts = mail_item.Attachments
    subject_pdf = subject + ".pdf"
    letter_saved = False
    any_saved = False

    for att in atts:
        fname = att.FileName or ""
        ext = Path(fname).suffix.lower()

        if ext in IMG_EXTS or fname.lower().startswith("image00"):
            print(f"[LET] Skip image → {fname}")
            continue

        if not letter_saved and fname.lower() == subject_pdf.lower():
            save_path = letter_f / fname
            letter_saved = True
        else:
            save_path = docs_f / fname

        if save_path.exists():
            print(f"[LET] Exists, skipping: {save_path}")
            continue

        att.SaveAsFile(str(save_path))
        print(f"[LET] Saved: {save_path}")
        any_saved = True

    if not letter_saved:
        print("[LET] WARNING: main letter PDF (subject code) not found.")

    return any_saved


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = get_target_folder(outlook, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    total_items = items.Count
    processed_items = 0

    done_trn = set()   # TRN kodları
    done_stq = set()   # Mail EntryID-ləri
    done_let = set()   # Mail EntryID-ləri

    stq_file_count = 0

    print_progress(0, total_items, "Processing mailbox...")

    for mail in items:
        processed_items += 1
        print_progress(processed_items, total_items, "Processing mailbox...")

        if getattr(mail, "Class", None) != 43:
            continue

        subject = (mail.Subject or "").strip()
        if not subject:
            continue

        # ---------- TRN ----------
        m_trn = re.search(r"(SPP2-PRO-KLN-TRN-\d{4})", subject)
        if m_trn:
            code = m_trn.group(1)
            if code not in done_trn:
                print(f"\n--- TRN FOUND: {code}")
                if process_single_trn(code):
                    done_trn.add(code)
            continue

        # ---------- STQ ----------
        if re.search(r"KLN-SPP2-STQ-[A-Za-z0-9-]+", subject):
            entry_id = mail.EntryID
            if entry_id not in done_stq:
                print(f"\n--- STQ MAIL FOUND: {subject}")
                saved = process_single_stq_mail(mail, subject)
                if saved > 0:
                    done_stq.add(entry_id)
                    stq_file_count += saved
            continue

        # ---------- LET ----------
        m_let = re.search(r"(SPP2-PRO-KLN-LET-\d{4})", subject)
        if m_let:
            entry_id = mail.EntryID
            if entry_id not in done_let:
                let_code = m_let.group(1)  # məsələn SPP2-PRO-KLN-LET-0022
                print(f"\n--- LET FOUND: {subject}")
                if process_single_let_mail(mail, subject, let_code):
                    done_let.add(entry_id)
            continue

    print()  # progress line-i qırmaq üçün

    print("\nDONE.")
    print(f"Processed TRNs: {sorted(done_trn) if done_trn else 'none'}")
    print(f"Processed STQs (files): {stq_file_count}")
    print(f"Processed LETs (mails): {len(done_let)}")


if __name__ == "__main__":
    main()
