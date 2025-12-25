# outlook_proyapi_incoming_auto.py
# Proyapi'dən gələn TRN, STQ, LET maillərini avtomatik emal edir.

import os
import re
import shutil
import subprocess
import time
import stat
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
# IO HARDENING (NO LOGIC CHANGE) ✅
# ---------------------------------------------------------
def _sleep_backoff(attempt: int, base: float = 0.35, cap: float = 2.5):
    # 0.35, 0.7, 1.4, 2.5, 2.5...
    t = min(cap, base * (2 ** max(0, attempt - 1)))
    time.sleep(t)


def robust_move(src: Path, dst: Path, attempts: int = 6) -> bool:
    """
    Network share-də AV/Index lock səbəbi ilə move bəzən WinError 5 verir.
    Retry + backoff.
    """
    src = Path(src)
    dst = Path(dst)

    for i in range(1, attempts + 1):
        try:
            shutil.move(str(src), str(dst))
            return True
        except PermissionError as e:
            print(f"[IO-RETRY] MOVE denied (try {i}/{attempts}) → {src} -> {dst} | {e}")
            _sleep_backoff(i)
        except OSError as e:
            # transient network / file busy
            print(f"[IO-RETRY] MOVE error (try {i}/{attempts}) → {src} -> {dst} | {e}")
            _sleep_backoff(i)

    print(f"[IO-FAIL] MOVE failed after {attempts} tries → {src} -> {dst}")
    return False


def _rmtree_onerror(func, path, exc_info):
    """
    shutil.rmtree onerror handler:
    - read-only flag -> chmod -> retry
    """
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as e:
        print(f"[IO-RM-ONERROR] Could not remove: {path} | {e}")


def robust_rmtree(p: Path, attempts: int = 6) -> bool:
    """
    Locked/read-only file tree olduqda rmtree partlayır.
    Retry + chmod fix.
    """
    p = Path(p)
    if not p.exists():
        return True

    for i in range(1, attempts + 1):
        try:
            shutil.rmtree(str(p), onerror=_rmtree_onerror)
            return True
        except PermissionError as e:
            print(f"[IO-RETRY] RMTREE denied (try {i}/{attempts}) → {p} | {e}")
            _sleep_backoff(i)
        except OSError as e:
            print(f"[IO-RETRY] RMTREE error (try {i}/{attempts}) → {p} | {e}")
            _sleep_backoff(i)

    print(f"[IO-FAIL] RMTREE failed after {attempts} tries → {p}")
    return False


def robust_unlink(p: Path, attempts: int = 6) -> bool:
    p = Path(p)
    if not p.exists():
        return True

    for i in range(1, attempts + 1):
        try:
            p.chmod(stat.S_IWRITE)
            p.unlink()
            return True
        except PermissionError as e:
            print(f"[IO-RETRY] UNLINK denied (try {i}/{attempts}) → {p} | {e}")
            _sleep_backoff(i)
        except OSError as e:
            print(f"[IO-RETRY] UNLINK error (try {i}/{attempts}) → {p} | {e}")
            _sleep_backoff(i)

    print(f"[IO-FAIL] UNLINK failed after {attempts} tries → {p}")
    return False


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

    HARDENED:
    - Retry + backoff (network lock/AV)
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

    out_dir.mkdir(parents=True, exist_ok=True)

    if not winrar_exe:
        print("\n[WARN] WinRAR tapılmadı, shutil.unpack_archive istifadə olunur.")
        # unpack_archive bəzən də lock yeyir -> retry
        for i in range(1, 5):
            try:
                shutil.unpack_archive(str(zip_path), str(out_dir))
                return
            except Exception as e:
                print(f"[IO-RETRY] unpack_archive error (try {i}/4) → {e}")
                _sleep_backoff(i)
        raise

    cmd = [
        winrar_exe,
        "x",  # extract
        "-o+",  # overwrite all without prompt
        "-idq",  # quiet
        str(zip_path),
        str(out_dir) + "\\",
    ]

    print(f"\n[TRN] Using WinRAR → {zip_path}")

    # Retry WinRAR: bəzən AV lock zamanı WinRAR exit code verə bilir
    last_err = None
    for i in range(1, 6):
        try:
            res = subprocess.run(
                cmd,
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return
        except Exception as e:
            last_err = e
            print(f"[IO-RETRY] WinRAR extract failed (try {i}/5) → {e}")
            _sleep_backoff(i)

    print(
        f"\n[WARN] WinRAR extract alınmadı ({last_err}), shutil.unpack_archive fallback."
    )
    for i in range(1, 5):
        try:
            shutil.unpack_archive(str(zip_path), str(out_dir))
            return
        except Exception as e:
            print(f"[IO-RETRY] unpack_archive error (try {i}/4) → {e}")
            _sleep_backoff(i)
    raise last_err


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

    HARDENED:
    - rmtree/move/unlink crash etməsin (warn + continue)
    """
    docs_folder = base_folder / "3. docs"
    attachments_folder = base_folder / "2. attachments"
    ensure_trn_subfolders(base_folder)

    raw_folder = base_folder / trn_code
    if raw_folder.exists() and raw_folder.is_dir():
        target_extracted = docs_folder / raw_folder.name
        if target_extracted.exists():
            ok = robust_rmtree(raw_folder)
            if ok:
                print(f"[TRN-CLEAN] Extra raw folder removed → {raw_folder}")
            else:
                print(
                    f"[TRN-CLEAN-WARN] Could not remove locked raw folder → {raw_folder}"
                )
        else:
            ok = robust_move(raw_folder, target_extracted)
            if ok:
                print(f"[TRN-CLEAN] Raw folder moved → {target_extracted}")
            else:
                print(
                    f"[TRN-CLEAN-WARN] Could not move raw folder (locked?) → {raw_folder}"
                )

    raw_zip = base_folder / f"{trn_code}.zip"
    if raw_zip.exists() and raw_zip.is_file():
        target_zip = attachments_folder / raw_zip.name
        if target_zip.exists():
            ok = robust_unlink(raw_zip)
            if ok:
                print(f"[TRN-CLEAN] Extra ZIP removed → {raw_zip}")
            else:
                print(f"[TRN-CLEAN-WARN] Could not delete ZIP (locked?) → {raw_zip}")
        else:
            ok = robust_move(raw_zip, target_zip)
            if ok:
                print(f"[TRN-CLEAN] ZIP moved → {target_zip}")
            else:
                print(f"[TRN-CLEAN-WARN] Could not move ZIP (locked?) → {raw_zip}")


def move_pdf_and_zip(base_folder: Path, trn_code: str):
    """
    PDF → 1. main
    ZIP extract → 3. docs
    ZIP → 2. attachments

    HARDENED:
    - extract/move/cleanup retry
    - exception olsa belə script ölməsin
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
            ok = robust_move(pdf_path, target_pdf)
            if ok:
                print(f"[TRN] PDF moved → {target_pdf}")
            else:
                print(f"[TRN-WARN] PDF move failed (locked?) → {pdf_path}")
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
                    robust_rmtree(target_extracted)

                ok = robust_move(extracted_folder, target_extracted)
                if ok:
                    print(f"[TRN] Extracted folder moved → {target_extracted}")
                else:
                    print(
                        f"[TRN-WARN] Extracted folder move failed (locked?) → {extracted_folder}"
                    )
            else:
                print(f"[WARN] Extracted folder not found: {extracted_folder}")

            # ZIP → 2. attachments
            target_zip = attachments_folder / zip_path.name
            if not target_zip.exists():
                ok = robust_move(zip_path, target_zip)
                if ok:
                    print(f"[TRN] ZIP moved → {target_zip}")
                else:
                    print(f"[TRN-WARN] ZIP move failed (locked?) → {zip_path}")
            else:
                print(f"[TRN] ZIP already in attachments → {target_zip}")

        except Exception as e:
            print(f"[ERROR] Error while extracting ZIP {zip_path}: {e}")

        # Əlavə safety cleanup – screenshotdakı problemi öldürmək üçün
        try:
            cleanup_trn_base_folder(base_folder, trn_code)
        except Exception as e:
            print(f"[TRN-CLEAN-WARN] cleanup failed but continuing → {e}")

    else:
        print(f"[WARN] ZIP not found: {zip_path}")
        try:
            cleanup_trn_base_folder(base_folder, trn_code)
        except Exception as e:
            print(f"[TRN-CLEAN-WARN] cleanup failed but continuing → {e}")


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

        try:
            shutil.copy2(str(pdf), str(target))
            print(f"[TRN] PDF copied to docs root → {target}")
        except PermissionError as e:
            print(f"[IO-WARN] copy2 denied → {pdf} -> {target} | {e}")


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
                try:
                    pdf.rename(target)
                    print(f"[RENAME] {old_name} → {new_name}")
                except PermissionError as e:
                    print(f"[IO-WARN] rename denied → {pdf} | {e}")
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

        try:
            pdf.rename(target)
            print(f"[TITLE] {old_name} → {new_name}")
        except PermissionError as e:
            print(f"[IO-WARN] title rename denied → {pdf} | {e}")


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

    word = None  # Word açmaq ağır əməliyyatdır – yalnız ehtiyac olanda
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
    STQ kodu və rev-i əsasən attachment adından çıxardır.

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
        m_fname = re.match(r"^(KLN-SPP2-STQ-[A-Za-z0-9-]+)_R(\d{2})[_.-].*$", fname)
        if m_fname:
            base_code = m_fname.group(1)
            rev = m_fname.group(2)
        else:
            # 2) Fallback: subject-dən ilk STQ kodu
            m_subj = re.search(r"(KLN-SPP2-STQ-[A-Za-z0-9-]+)(?:_R(\d{2}))?", subject)
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

        try:
            att.SaveAsFile(str(save_path))
            print(f"[STQ] Saved: {save_path}")
            saved_count += 1
        except Exception as e:
            print(f"[STQ-WARN] SaveAsFile failed → {save_path} | {e}")

    return saved_count


# ---------------------------------------------------------
# LET FUNKSİYALARI
# ---------------------------------------------------------
def get_let_folder_by_code(root: Path, let_code: str) -> Path:
    base = root / let_code
    base.mkdir(parents=True, exist_ok=True)
    return base


def process_single_let_mail(mail_item, subject: str, let_code: str) -> bool:
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

        try:
            att.SaveAsFile(str(save_path))
            print(f"[LET] Saved: {save_path}")
            any_saved = True
        except Exception as e:
            print(f"[LET-WARN] SaveAsFile failed → {save_path} | {e}")

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

    done_trn = set()  # TRN kodları
    done_stq = set()  # Mail EntryID-ləri
    done_let = set()  # Mail EntryID-ləri

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
                try:
                    if process_single_trn(code):
                        done_trn.add(code)
                except Exception as e:
                    print(f"[TRN-FAIL] {code} failed but continuing → {e}")
            continue

        # ---------- STQ ----------
        if re.search(r"KLN-SPP2-STQ-[A-Za-z0-9-]+", subject):
            entry_id = mail.EntryID
            if entry_id not in done_stq:
                print(f"\n--- STQ MAIL FOUND: {subject}")
                try:
                    saved = process_single_stq_mail(mail, subject)
                    if saved > 0:
                        done_stq.add(entry_id)
                        stq_file_count += saved
                except Exception as e:
                    print(f"[STQ-FAIL] STQ mail failed but continuing → {e}")
            continue

        # ---------- LET ----------
        m_let = re.search(r"(SPP2-PRO-KLN-LET-\d{4})", subject)
        if m_let:
            entry_id = mail.EntryID
            if entry_id not in done_let:
                let_code = m_let.group(1)
                print(f"\n--- LET FOUND: {subject}")
                try:
                    if process_single_let_mail(mail, subject, let_code):
                        done_let.add(entry_id)
                except Exception as e:
                    print(f"[LET-FAIL] LET mail failed but continuing → {e}")
            continue

    print()  # progress line-i qırmaq üçün

    print("\nDONE.")
    print(f"Processed TRNs: {sorted(done_trn) if done_trn else 'none'}")
    print(f"Processed STQs (files): {stq_file_count}")
    print(f"Processed LETs (mails): {len(done_let)}")


if __name__ == "__main__":
    main()
