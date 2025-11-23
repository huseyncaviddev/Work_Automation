# outlook_proyapi_incoming_auto.py
# Proyapi'dən gələn TRN, STQ, LET maillərini avtomatik emal edir.

import os
import re
import shutil
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
# TRN FUNKSİYALARI  (v1 lojiqası)
# ---------------------------------------------------------
def notify_missing_trn(trn_code):
    msg = (
        f"[MISSING TRN] {trn_code} QA-QC Proyapi folderində tapılmadı. "
        f"Səidə xanıma WP ilə xəbər verilməlidir."
    )
    print(msg)


def ensure_trn_subfolders(base_folder):
    """1. main / 2. attachments / 3. docs qovluqlarını yaradır."""
    for name in ["1. main", "2. attachments", "3. docs"]:
        (base_folder / name).mkdir(exist_ok=True)


def move_pdf_and_zip(base_folder, trn_code):
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
        shutil.move(str(pdf_path), str(target_pdf))
        print(f"[TRN] PDF moved → {target_pdf}")
    else:
        print(f"[WARN] PDF not found: {pdf_path}")

    # ZIP → extract
    if zip_path.exists():
        try:
            shutil.unpack_archive(str(zip_path), str(base_folder))
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
            shutil.move(str(zip_path), str(target_zip))
            print(f"[TRN] ZIP moved → {target_zip}")

        except Exception as e:
            print(f"[ERROR] Error while extracting ZIP {zip_path}: {e}")
    else:
        print(f"[WARN] ZIP not found: {zip_path}")


def copy_all_pdfs_to_docs_root(docs_folder):
    """
    3. docs altındakı bütün pdf-ləri rekursiv tapır,
    3. docs root-a kopyalayır. Eyni ad varsa → _copy əlavə edir.
    """
    if not docs_folder.exists():
        return

    for pdf in docs_folder.rglob("*.pdf"):
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


def rename_r_dash_to_r_underscore(docs_folder):
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


def title_cleanup_pattern(docs_folder):
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


def convert_main_pdf_to_docx(base_folder, trn_code):
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


def process_single_trn(trn_code):
    """
    1) QA-QC Proyapi-də folderi tap
    2) Incoming 1. TRN altına copy
    3) 1.main / 2.attachments / 3.docs yarat
    4) PDF & ZIP əməliyyatları
    5) 3.docs flatten + rename
    6) Main PDF → DOCX
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
    move_pdf_and_zip(target_folder, trn_code)

    docs_folder = target_folder / "3. docs"
    copy_all_pdfs_to_docs_root(docs_folder)
    rename_r_dash_to_r_underscore(docs_folder)
    title_cleanup_pattern(docs_folder)
    convert_main_pdf_to_docx(target_folder, trn_code)

    return True


# ---------------------------------------------------------
# STQ FUNKSİYALARI  (rev-lisiz subject üçün də)
# ---------------------------------------------------------
def find_stq_target_folder(base_code):
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


def process_single_stq_mail(mail_item, subject):
    """
    STQ subject variantları:
      1) KLN-SPP2-STQ-WE-GN00-309_R00_Prokon_Reply
      2) KLN-SPP2-STQ-MC-GF04-277_Prokon_Reply
    """
    base_code = None
    rev = None

    # Variant 1: subject-də rev var
    m = re.match(r"^(KLN-SPP2-STQ-[A-Za-z0-9-]+)_R(\d{2})_.*$", subject)
    if m:
        base_code = m.group(1)
        rev = m.group(2)
    else:
        # Variant 2: subject-də rev YOXDUR → base_code götür
        m2 = re.match(r"^(KLN-SPP2-STQ-[A-Za-z0-9-]+)_.*$", subject)
        if not m2:
            print(f"[STQ] Subject STQ pattern-ə düşmədi → {subject}")
            return False

        base_code = m2.group(1)

        # Rev-i attachment adlarından tapmağa çalış
        atts = mail_item.Attachments
        for att in atts:
            fname = att.FileName or ""
            mr = re.search(r"_R(\d{2})[_\-.]", fname)
            if mr:
                rev = mr.group(1)
                break

        # Hələ də tapılmadısa → 00
        if rev is None:
            rev = "00"
            print(f"[STQ] Rev subject və attach-də tapılmadı, default R{rev} istifadə olunur.")

    target = find_stq_target_folder(base_code)
    if target is None:
        return False

    print(f"[STQ] → {target} (rev R{rev})")

    atts = mail_item.Attachments
    idx = 0
    saved = False

    for att in atts:
        fname = att.FileName or ""
        ext = Path(fname).suffix.lower()

        if ext in IMG_EXTS or fname.lower().startswith("image00"):
            print(f"[STQ] Skip image → {fname}")
            continue

        idx += 1
        if idx == 1:
            new = f"{base_code}_R{rev} Reply{ext}"
        else:
            new = f"{base_code}_R{rev} Reply_{idx}{ext}"

        save_path = target / new
        if save_path.exists():
            save_path = target / (save_path.stem + "_copy" + save_path.suffix)

        att.SaveAsFile(str(save_path))
        print(f"[STQ] Saved: {save_path}")
        saved = True

    return saved


# ---------------------------------------------------------
# LET FUNKSİYALARI
# ---------------------------------------------------------
def get_next_let_folder_name(root):
    max_n = 0
    if root.exists():
        for child in root.iterdir():
            if not child.is_dir():
                continue
            m = re.match(r"SPP2-PRO-KLN-LET-(\d{4})", child.name)
            if m:
                n = int(m.group(1))
                max_n = max(max_n, n)
    return f"SPP2-PRO-KLN-LET-{max_n + 1:04d}"


def process_single_let_mail(mail_item, subject):
    folder_name = get_next_let_folder_name(LET_INCOMING_ROOT)
    base = LET_INCOMING_ROOT / folder_name
    letter_f = base / "1. letter"
    docs_f = base / "2. docs"

    letter_f.mkdir(parents=True, exist_ok=True)
    docs_f.mkdir(exist_ok=True)

    print(f"[LET] {folder_name}")

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
            save_path = save_path.with_name(save_path.stem + "_copy" + save_path.suffix)

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

    done_trn = set()
    done_stq = set()
    done_let = set()

    for mail in items:
        if getattr(mail, "Class", None) != 43:
            continue

        subject = (mail.Subject or "").strip()
        if not subject:
            continue

        # ---------- TRN ----------
        m_trn = re.match(r"^(SPP2-PRO-KLN-TRN-\d{4})$", subject)
        if m_trn:
            code = m_trn.group(1)
            if code not in done_trn:
                print(f"\n--- TRN FOUND: {code}")
                if process_single_trn(code):
                    done_trn.add(code)
            continue

        # ---------- STQ ----------
        # Rev olsa da, olmasa da STQ kimi tut
        m_stq = re.match(r"^(KLN-SPP2-STQ-[A-Za-z0-9-]+)(?:_R\d{2})?_.*$", subject)
        if m_stq and subject not in done_stq:
            print(f"\n--- STQ FOUND: {subject}")
            if process_single_stq_mail(mail, subject):
                done_stq.add(subject)
            continue

        # ---------- LET ----------
        m_let = re.match(r"^(SPP2-PRO-KLN-LET-\d{4})$", subject)
        if m_let and subject not in done_let:
            print(f"\n--- LET FOUND: {subject}")
            if process_single_let_mail(mail, subject):
                done_let.add(subject)
            continue

    print("\nDONE.")
    print(f"Processed TRNs: {sorted(done_trn) if done_trn else 'none'}")
    print(f"Processed STQs: {len(done_stq)}")
    print(f"Processed LETs: {len(done_let)}")


if __name__ == "__main__":
    main()
