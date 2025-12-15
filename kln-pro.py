# outlook_kolin_sunulacaklar_auto.py

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
SUBPATH = r"Inbox\Sunulacaklar"  # KOLIN dept mail-ləri

TRN_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN")
STQ_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ")
LET_ROOT = Path(
    r"G:\My Drive\4-S1 ve S2 Ortak Dökümanlar\03-SPP LETTERS\SPP2-LET\1. KLN-PRO\01-Outgoing"
)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}

LOOKBACK_DAYS = None

# TRN ilə göndərilən sənəd növləri
TRN_DOC_TYPES = {
    "CLC",
    "DWG",
    "FRM",
    "ITP",
    "JSA",
    "LOG",
    "LST",
    "MAR",
    "MES",
    "NCR",
    "ORG",
    "REP",
    "SPE",
    "SAR",
    "LPL",
    "SRF",
}


# ---------------------------------
# TRN – növbəti qovluq
# ---------------------------------
def get_next_trn_folder(root: Path) -> Path:
    if not root.exists():
        root.mkdir(parents=True, exist_ok=True)

    max_num = 0
    for d in root.iterdir():
        if not d.is_dir():
            continue
        m = re.match(r"^SPP2-KLN-PRO-TRN-(\d{4})$", d.name)
        if not m:
            continue
        num = int(m.group(1))
        if num > max_num:
            max_num = num

    next_num = max_num + 1
    folder_name = f"SPP2-KLN-PRO-TRN-{next_num:04d}"
    new_folder = root / folder_name
    new_folder.mkdir(parents=True, exist_ok=True)

    for sub in ["1. main", "2. attachments", "3. docs"]:
        (new_folder / sub).mkdir(exist_ok=True)

    print(f"[TRN] Created → {new_folder}")
    return new_folder


# ---------------------------------
# STQ – növbəti index
# ---------------------------------
def get_next_stq_index(root: Path) -> int:
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


def create_stq_folder_and_save(
    next_index: int, stq_attachment, base_folder: Path
) -> tuple[int, Path]:
    """
    Verilən STQ attachment üçün:
      - STQ kodunu çıxarır
      - Növbəti index ilə folder yaradır
      - STQ faylını rename edib həmin folderə saxlayır
    Geri:
      (next_index_after, stq_folder_path)
    """
    raw = stq_attachment.FileName
    name, ext = os.path.splitext(raw)
    upper = name.upper()

    m = re.search(r"(KLN-SPP2-STQ-[A-Z0-9]+-[A-Z0-9]+)", upper)
    if m:
        prefix = m.group(1)
    else:
        prefix = re.split(r"_R\d{2}", upper)[0]

    stq_code = f"{prefix}-{next_index}"
    folder_name = f"{next_index}. {stq_code}"
    stq_folder = base_folder / folder_name
    stq_folder.mkdir(parents=True, exist_ok=True)

    m_rev = re.search(r"_R\d{2}", upper)
    rev_part = m_rev.group(0) if m_rev else "_R00"
    new_filename = f"{stq_code}{rev_part}{ext}"

    target = stq_folder / new_filename
    stq_attachment.SaveAsFile(str(target))

    print(f"[STQ] Saved main STQ → {target}")
    return next_index + 1, stq_folder


# ---------------------------------
# LET – növbəti folder
# ---------------------------------
def get_next_let_folder(root: Path) -> Path:
    if not root.exists():
        root.mkdir(parents=True, exist_ok=True)

    max_num = 0
    for d in root.iterdir():
        if not d.is_dir():
            continue
        m = re.match(r"^SPP2-KLN-PRO-LET-(\d{4})$", d.name)
        if not m:
            continue
        num = int(m.group(1))
        if num > max_num:
            max_num = num

    next_num = max_num + 1
    folder_name = f"SPP2-KLN-PRO-LET-{next_num:04d}"
    new_folder = root / folder_name
    new_folder.mkdir(parents=True, exist_ok=True)

    for sub in ["1. letter", "2. docs"]:
        (new_folder / sub).mkdir(exist_ok=True)

    print(f"[LET] Created → {new_folder}")
    return new_folder


# ---------------------------------
# Helper-lər
# ---------------------------------
def get_target_folder(ns, mailbox, subpath):
    folder = ns.Folders[mailbox]
    for part in subpath.split("\\"):
        if part:
            folder = folder.Folders[part]
    return folder


def is_kln_code_file(filename: str) -> bool:
    upper = filename.upper()
    return upper.startswith("KLN-SPP2-") or upper.startswith("PRO-SPP2-")


def get_doc_type_from_filename(filename: str) -> str | None:
    # COMPANY-SPP2-XXX-... formatını tutur (KLN, PRO fərq etmir)
    m = re.match(r"(?i)[A-Z0-9]+-SPP2-([A-Z0-9]{3})-", filename)
    return m.group(1).upper() if m else None


def clean_filename_keep_code_only(filename: str) -> str:
    name, ext = os.path.splitext(filename)
    m = re.search(r"_R\d{2}", name, flags=re.IGNORECASE)
    if m:
        code = name[: m.end()]
    else:
        code = name.split(" ")[0]
    code = re.sub(r'[\\/:*?"<>|]', "_", code)
    return code + ext


def extract_shd_paths_from_mail(item) -> list[str]:
    """
    Body + HTMLBody-dən UNC path-ləri çıxarır:
      \\DATA\DataServer\Elektrik\11- SHOPDRAWING\PROYAPI SUNUM\...\ES03
    """
    body = getattr(item, "Body", "") or ""
    html = getattr(item, "HTMLBody", "") or ""
    text = body + "\n" + html

    raw_matches = re.findall(r"(\\\\[^\r\n<>]+)", text)

    cleaned: list[str] = []
    for m in raw_matches:
        p = m.strip()
        p = p.rstrip(" .;,)")
        if p not in cleaned:
            cleaned.append(p)

    return cleaned


def copy_shd_folder_to_trn(shd_path: str, trn_docs_folder: Path) -> bool:
    """
    SHD sunum path-lərini TRN docs qovluğuna kopyalayır.

    Məsələn:
      \\DATA\...\SOCKET SYSTEM INSTALLATION\GF19
      \\DATA\...\LIGHTING INSTALLATION\GF19

    TRN strukturu belə olsun:
      TRN-XXXX\3. docs\GF19\SOCKET SYSTEM INSTALLATION\...
      TRN-XXXX\3. docs\GF19\LIGHTING INSTALLATION\...
    """
    src = Path(shd_path)

    if not src.exists():
        print(f"[SHD] SKIP — path does NOT exist: {src}")
        return False

    # Son qovluq adı (məs: GF19)
    zone_folder = src.name

    # Onun parent qovluğu (məs: SOCKET SYSTEM INSTALLATION / LIGHTING INSTALLATION)
    parent = src.parent
    parent_name = parent.name if parent and parent != src else None

    # Əsas GF19 root-u (3. docs içində)
    zone_root = trn_docs_folder / zone_folder
    zone_root.mkdir(parents=True, exist_ok=True)

    # Əgər parent varsa → GF19 altında parent qovluğu açıb ora kopyalayırıq
    if parent_name:
        dst = zone_root / parent_name
    else:
        # Fallback: köhnə lojiqa kimi birbaşa 3. docs altına
        dst = zone_root

    if dst.exists():
        print(f"[SHD] SKIP — already copied: {dst}")
        return False

    try:
        shutil.copytree(src, dst)
        print(f"[SHD] COPIED → {dst}")
        return True
    except Exception as e:
        print(f"[SHD] ERROR copying {src} → {e}")
        return False


# --------- SHD üçün PowerShell lojiqasının Python versiyası ---------
def add_r00_to_subfiles_without_rev(docs_root: Path):
    """
    Subfolder faylları üçün rev normalizasiya:
      - Hər hansı rev varsa (_Rdd və ya -Rdd) → rev var sayılır
      - Əgər -Rdd varsa → _Rdd-ə çevirilir
      - Əgər ümumiyyətlə rev yoxdursa → sona _R00 əlavə edilir
    """
    if not docs_root.exists():
        return

    rev_pattern = re.compile(r"([_-]R(\d{2}))", re.IGNORECASE)

    for f in docs_root.rglob("*"):
        if not f.is_file():
            continue
        if f.parent == docs_root:
            continue  # Root-a toxunmuruq

        stem = f.stem
        suffix = f.suffix

        match = rev_pattern.search(stem)

        if match:
            full_rev = match.group(1)  # -R00 və ya _R00
            rev_number = match.group(2)  # 00, 01, 02...

            if full_rev.startswith("-"):
                fixed_rev = f"_R{rev_number}"
                new_stem = rev_pattern.sub(fixed_rev, stem)

                new_name = new_stem + suffix
                target = f.with_name(new_name)

                if f.name != new_name:
                    try:
                        if target.exists():
                            target.unlink()
                        f.rename(target)
                        print(f"[SHD-REV-FIX] {f.name} → {new_name}")
                    except Exception as e:
                        print(f"[SHD-ERR] {f.name} → {e}")
            continue
        else:
            new_name = f"{stem}_R00{suffix}"
            target = f.with_name(new_name)

            try:
                if target.exists():
                    target.unlink()
                f.rename(target)
                print(f"[SHD-R00-ADD] {f.name} → {new_name}")
            except Exception as e:
                print(f"[SHD-R00-ERR] {f.name} → {e}")


def copy_pdfs_from_subfolders_to_root(docs_root: Path):
    """
    PS: PDF Copier
    - Subfolder-lərdəki bütün PDF-ləri 3. docs root-a kopyalayır
    """
    if not docs_root.exists():
        return

    for pdf in docs_root.rglob("*.pdf"):
        if pdf.parent == docs_root:
            continue
        dest = docs_root / pdf.name
        try:
            if dest.exists():
                dest.unlink()
            shutil.copy2(pdf, dest)
            print(f"[SHD-PDF] {pdf} → {dest}")
        except Exception as e:
            print(f"[SHD-PDF-ERR] {pdf} → {e}")


def move_xlsx_docx_from_subfolders_to_root(docs_root: Path):
    """
    PS: XLSX & DOCX Mover
    - Subfolder-lərdəki .xlsx və .docx fayllarını 3. docs root-a move edir
    """
    if not docs_root.exists():
        return

    exts = {".xlsx", ".docx"}

    for f in docs_root.rglob("*"):
        if not f.is_file():
            continue
        if f.parent == docs_root:
            continue
        if f.suffix.lower() not in exts:
            continue

        dest = docs_root / f.name
        try:
            if dest.exists():
                dest.unlink()
            shutil.move(str(f), str(dest))
            print(f"[SHD-MOVE] {f} → {dest}")
        except Exception as e:
            print(f"[SHD-MOVE-ERR] {f} → {e}")


# ---------------------------------
# MAIN
# ---------------------------------
def main():
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.Session
    folder = get_target_folder(session, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # LOOKBACK_DAYS None olarsa → cutoff yoxdur
    if LOOKBACK_DAYS is not None:
        cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    else:
        cutoff = None

    trn_attachments = []  # a) KLN-SPP2-* fayllar
    shd_paths = set()  # b) SHD sunum UNC path-ləri
    stq_jobs = []  # c) STQ paketləri
    let_attachments = []  # d) LET faylları

    for item in items:
        if getattr(item, "Class", None) != 43:
            continue

        recv_time = getattr(item, "ReceivedTime", None)
        if isinstance(recv_time, datetime):
            recv_time_naive = recv_time.replace(tzinfo=None)
        else:
            recv_time_naive = None

        # cutoff yalnız təyin olunubsa işləsin
        if cutoff is not None and recv_time_naive and recv_time_naive < cutoff:
            print("Reached cutoff date. Stopping scan.")
            break

        # SHD sunum linkləri
        for p in extract_shd_paths_from_mail(item):
            shd_paths.add(p)

        if not item.Attachments:
            continue

        # Attachmentləri listə yığırıq ki, eyni maili 2 dəfə dolaşa bilək
        attachments = [att for att in item.Attachments]

        # 1) Bu maildə STQ varmı? (multi-STQ dəstəyi)
        stq_atts = []
        non_stq_atts = []

        for att in attachments:
            fname = att.FileName
            ext = os.path.splitext(fname)[1].lower()

            if ext in IMAGE_EXTS:
                continue

            if is_kln_code_file(fname):
                doc_type = get_doc_type_from_filename(fname)
                if doc_type == "STQ":
                    # Hər STQ faylı ayrıca işlənəcək
                    stq_atts.append(att)
                    continue

            # STQ olmayan (amma image də olmayan) bütün attachmentlər
            non_stq_atts.append(att)

        if stq_atts:
            # Bu mail STQ mailidir → hər STQ üçün ayrı job, eyni extra-lar paylaşılır
            for stq_att in stq_atts:
                stq_jobs.append((stq_att, non_stq_atts))
            # Bu mail üçün TRN/LET lojiqasına girmirik
            continue

        # 2) STQ yoxdursa, əvvəlki kimi TRN/LET lojiqası
        for att in attachments:
            fname = att.FileName
            ext = os.path.splitext(fname)[1].lower()

            if ext in IMAGE_EXTS:
                continue

            # LET (d) → SPP2-KLN-PRO-LET-XXXX.docx
            if re.match(r"(?i)^SPP2-KLN-PRO-LET-\d{4}\.docx$", fname):
                let_attachments.append(att)
                continue

            # KLN-SPP2-* sənədlər
            if not is_kln_code_file(fname):
                continue

            doc_type = get_doc_type_from_filename(fname)

            if doc_type in TRN_DOC_TYPES:
                trn_attachments.append(att)

    # 1) a-bəndi: TRN docs üçün AYRI transmittal
    if trn_attachments:
        trn_folder_docs = get_next_trn_folder(TRN_ROOT)
        docs_dir_docs = trn_folder_docs / "3. docs"

        for att in trn_attachments:
            clean_name = clean_filename_keep_code_only(att.FileName)
            target = docs_dir_docs / clean_name
            if not target.exists():
                att.SaveAsFile(str(target))
                print(f"[TRN-doc] Saved → {target}")

    # 2) b-bəndi: SHD sunumlar üçün AYRI transmittal (+ PS lojiqası)
    if shd_paths:
        trn_folder_shd = get_next_trn_folder(TRN_ROOT)
        docs_dir_shd = trn_folder_shd / "3. docs"

        for shd in shd_paths:
            copy_shd_folder_to_trn(shd, docs_dir_shd)

        add_r00_to_subfiles_without_rev(docs_dir_shd)
        copy_pdfs_from_subfolders_to_root(docs_dir_shd)
        move_xlsx_docx_from_subfolders_to_root(docs_dir_shd)

    # 3) c-bəndi: STQ-lər (mail + bütün attachmentləri ilə)
    if stq_jobs:
        next_stq_no = get_next_stq_index(STQ_ROOT)
        for stq_att, extra_atts in stq_jobs:
            next_stq_no, stq_folder = create_stq_folder_and_save(
                next_stq_no, stq_att, STQ_ROOT
            )

            # Eyni maildəki digər attachmentləri də STQ qovluğuna saxlayırıq
            for att in extra_atts:
                target = stq_folder / att.FileName
                if target.exists():
                    # overwrite etmə; sadəcə xəbər ver və keç
                    print(f"[STQ-extra] SKIP (exists) → {target}")
                    continue
                att.SaveAsFile(str(target))
                print(f"[STQ-extra] Saved → {target}")

    # 4) d-bəndi: LET-lər
    for att in let_attachments:
        let_folder = get_next_let_folder(LET_ROOT)
        docs_dir = let_folder / "2. docs"
        target = docs_dir / att.FileName
        if not target.exists():
            att.SaveAsFile(str(target))
            print(f"[LET-doc] Saved → {target}")

    print("\nSUMMARY:")
    print(f"  TRN docs  : {len(trn_attachments)}")
    print(f"  SHD paths : {len(shd_paths)}")
    print(f"  STQ mails : {len(stq_jobs)}")
    print(f"  LET files : {len(let_attachments)}")


if __name__ == "__main__":
    main()
