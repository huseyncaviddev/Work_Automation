# -*- coding: utf-8 -*-
# outlook_kolin_sunulacaklar_auto.py

import os
import re
import html
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

import win32com.client as win32


# ---------------------------------
# KONFİQ (ULTRA FAST)
# ---------------------------------
MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\Sunulacaklar"  # KOLIN dept mail-ləri

TRN_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN")
STQ_ROOT = Path(r"\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\3. STQ")
LET_ROOT = Path(
    r"G:\My Drive\4-S1 ve S2 Ortak Dökümanlar\03-SPP LETTERS\SPP2-LET\1. KLN-PRO\01-Outgoing"
)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"}

# Outlook scan performansı üçün:
MAX_SCAN = 250  # son 250 mail (istəsən None elə → limitsiz)
LOOKBACK_DAYS = 5  # son 5 gün (istəsən None elə → cutoff yoxdur)

# Robocopy performansı
ROBOCOPY_MT = 16  # 8/16/32 yaxşıdır
ROBO_RETRY = 1
ROBO_WAIT = 1

# SHD copy mode:
# True  -> yalnız lazım olan fayllar (FAST): pdf/xlsx/xlsm/docx
# False -> hər şeyi kopyala (slow, amma full)
SHD_COPY_FILTERED = True

SHD_INCLUDE_GLOBS = ["*.pdf", "*.xlsx", "*.xlsm", "*.docx"]  # filtered mode üçün

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


# ---------------------------------
# SHD path extraction (fix: (R1), HTML, zibil)
# ---------------------------------
UNC_RE = re.compile(r"(\\\\[^\r\n<>]+)", re.IGNORECASE)


def normalize_unc_path(p: str) -> str | None:
    if not p:
        return None

    p = html.unescape(p).strip()
    p = p.strip(" \t\r\n\"'")
    p = p.rstrip(" .;,)>]\"'")

    # maildə path sonuna əlavə olunan "(R1)" kimi hissələri çıxar
    p = re.sub(r"\s*\(R\d+\)\s*$", "", p, flags=re.IGNORECASE)

    # slash normalize
    p = p.replace("/", "\\").rstrip("\\")

    if not p.startswith("\\\\"):
        return None
    return p


def extract_shd_paths_from_mail(item) -> list[str]:
    body = getattr(item, "Body", "") or ""
    html_body = getattr(item, "HTMLBody", "") or ""
    text = body + "\n" + html_body

    raw = UNC_RE.findall(text)

    out: list[str] = []
    seen = set()
    for m in raw:
        p = normalize_unc_path(m)
        if not p:
            continue
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


# ---------------------------------
# ROBOCOPY (FAST)
# ---------------------------------
def robocopy_dir(
    src: Path, dst: Path, include_globs: list[str] | None, mt: int
) -> bool:
    """
    Robocopy exit codes:
      0-7 success (0: nothing copied)
      >=8 failure
    """
    dst.mkdir(parents=True, exist_ok=True)

    cmd = [
        "robocopy",
        str(src),
        str(dst),
        "/E",
        f"/MT:{mt}",
        f"/R:{ROBO_RETRY}",
        f"/W:{ROBO_WAIT}",
        "/NP",
        "/NFL",
        "/NDL",
        "/NJH",
        "/NJS",
        "/XO",
    ]

    if include_globs:
        cmd += include_globs

    r = subprocess.run(cmd, capture_output=True, text=True)

    if r.returncode >= 8:
        print("[ROBOCOPY-FAIL] returncode:", r.returncode)
        print("[ROBOCOPY-OUT]", (r.stdout or "")[-1200:])
        print("[ROBOCOPY-ERR]", (r.stderr or "")[-1200:])
        return False

    return True


def copy_shd_folder_to_trn(shd_path: str, trn_docs_folder: Path) -> bool:
    """
    Məsələn:
      \\DATA\...\SOCKET SYSTEM INSTALLATION\GF19
    TRN-də:
      TRN-XXXX\3. docs\GF19\SOCKET SYSTEM INSTALLATION\...
    """
    src = Path(shd_path)

    if not src.exists():
        print(f"[SHD] SKIP — path does NOT exist: {src}")
        return False

    zone_folder = src.name  # GF19
    parent_name = src.parent.name  # SOCKET SYSTEM INSTALLATION

    zone_root = trn_docs_folder / zone_folder
    zone_root.mkdir(parents=True, exist_ok=True)

    dst = zone_root / parent_name

    marker = dst / ".copy_complete"
    if marker.exists():
        print(f"[SHD] SKIP — already completed: {dst}")
        return False

    include = SHD_INCLUDE_GLOBS if SHD_COPY_FILTERED else None

    print(f"[SHD] ROBOCOPY → {src} -> {dst}")
    ok = robocopy_dir(src, dst, include_globs=include, mt=ROBOCOPY_MT)

    if ok:
        marker.parent.mkdir(parents=True, exist_ok=True)
        marker.write_text("ok", encoding="utf-8")
        print(f"[SHD] DONE → {dst}")
        return True

    return False


# --------- SHD üçün post-process ---------
def add_r00_to_subfiles_without_rev(docs_root: Path):
    if not docs_root.exists():
        return

    rev_pattern = re.compile(r"([_-]R(\d{2}))", re.IGNORECASE)

    for f in docs_root.rglob("*"):
        if not f.is_file():
            continue
        if f.parent == docs_root:
            continue

        stem = f.stem
        suffix = f.suffix

        match = rev_pattern.search(stem)

        if match:
            full_rev = match.group(1)
            rev_number = match.group(2)

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
            # print(f"[SHD-PDF] {pdf} → {dest}")
        except Exception as e:
            print(f"[SHD-PDF-ERR] {pdf} → {e}")


def move_xlsx_docx_from_subfolders_to_root(docs_root: Path):
    if not docs_root.exists():
        return

    exts = {".xlsx", ".docx", ".xlsm"}

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
            # print(f"[SHD-MOVE] {f} → {dest}")
        except Exception as e:
            print(f"[SHD-MOVE-ERR] {f} → {e}")


# ---------------------------------
# MAIN
# ---------------------------------
def iter_attachments(item):
    """Outlook COM attachment iterator (stable)."""
    atts = getattr(item, "Attachments", None)
    if not atts:
        return []
    out = []
    try:
        cnt = atts.Count
        for i in range(1, cnt + 1):
            out.append(atts.Item(i))
    except Exception:
        return []
    return out


def main():
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.Session
    folder = get_target_folder(session, MAILBOX_NAME, SUBPATH)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = None
    if LOOKBACK_DAYS is not None:
        cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)

    # Lazy TRN docs folder (yalnız lazım olanda yaradılır)
    trn_folder_docs = None
    docs_dir_docs = None
    trn_docs_count = 0

    # SHD path-lər toplanır, sonra 1 TRN açılır
    shd_paths = set()

    # STQ index scan əvvəlcədən (fast)
    next_stq_no = get_next_stq_index(STQ_ROOT)
    stq_count = 0

    # LET count
    let_count = 0

    total = items.Count
    scan_count = total if MAX_SCAN is None else min(total, MAX_SCAN)

    print(
        f"[SCAN] Total items: {total} | Scanning: {scan_count} | LOOKBACK_DAYS={LOOKBACK_DAYS}"
    )

    for idx in range(1, scan_count + 1):
        item = items.Item(idx)

        if getattr(item, "Class", None) != 43:
            continue

        recv_time = getattr(item, "ReceivedTime", None)
        if isinstance(recv_time, datetime):
            recv_time_naive = recv_time.replace(tzinfo=None)
        else:
            recv_time_naive = None

        if cutoff is not None and recv_time_naive and recv_time_naive < cutoff:
            print("[SCAN] Reached cutoff date. Stopping scan.")
            break

        # SHD linkləri
        for p in extract_shd_paths_from_mail(item):
            shd_paths.add(p)

        attachments = iter_attachments(item)
        if not attachments:
            continue

        # STQ detection (multi STQ)
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
                    stq_atts.append(att)
                    continue

            non_stq_atts.append(att)

        # STQ maili → dərhal işləyirik
        if stq_atts:
            for stq_att in stq_atts:
                next_stq_no, stq_folder = create_stq_folder_and_save(
                    next_stq_no, stq_att, STQ_ROOT
                )
                stq_count += 1

                for att in non_stq_atts:
                    target = stq_folder / att.FileName
                    if target.exists():
                        print(f"[STQ-extra] SKIP (exists) → {target}")
                        continue
                    try:
                        att.SaveAsFile(str(target))
                        print(f"[STQ-extra] Saved → {target}")
                    except Exception as e:
                        print(f"[STQ-extra] ERROR {att.FileName} → {e}")
            continue

        # STQ yoxdursa: TRN/LET
        for att in attachments:
            fname = att.FileName
            ext = os.path.splitext(fname)[1].lower()

            if ext in IMAGE_EXTS:
                continue

            # LET
            if re.match(r"(?i)^SPP2-KLN-PRO-LET-\d{4}\.docx$", fname):
                let_folder = get_next_let_folder(LET_ROOT)
                docs_dir = let_folder / "2. docs"
                target = docs_dir / fname
                if not target.exists():
                    try:
                        att.SaveAsFile(str(target))
                        let_count += 1
                        print(f"[LET-doc] Saved → {target}")
                    except Exception as e:
                        print(f"[LET-doc] ERROR {fname} → {e}")
                continue

            # TRN docs
            if not is_kln_code_file(fname):
                continue

            doc_type = get_doc_type_from_filename(fname)
            if doc_type in TRN_DOC_TYPES:
                if trn_folder_docs is None:
                    trn_folder_docs = get_next_trn_folder(TRN_ROOT)
                    docs_dir_docs = trn_folder_docs / "3. docs"

                clean_name = clean_filename_keep_code_only(fname)
                target = docs_dir_docs / clean_name
                if not target.exists():
                    try:
                        att.SaveAsFile(str(target))
                        trn_docs_count += 1
                        print(f"[TRN-doc] Saved → {target}")
                    except Exception as e:
                        print(f"[TRN-doc] ERROR {fname} → {e}")

    # SHD paths üçün ayrı TRN (+ post-process)
    shd_copied = 0
    if shd_paths:
        trn_folder_shd = get_next_trn_folder(TRN_ROOT)
        docs_dir_shd = trn_folder_shd / "3. docs"

        for shd in sorted(shd_paths):
            try:
                if copy_shd_folder_to_trn(shd, docs_dir_shd):
                    shd_copied += 1
            except Exception as e:
                print(f"[SHD] ERROR {shd} → {e}")

        # post-process (yalnız copied data üzərində)
        add_r00_to_subfiles_without_rev(docs_dir_shd)
        copy_pdfs_from_subfolders_to_root(docs_dir_shd)
        move_xlsx_docx_from_subfolders_to_root(docs_dir_shd)

    print("\nSUMMARY:")
    print(f"  TRN docs saved : {trn_docs_count}")
    print(f"  SHD paths found: {len(shd_paths)} | copied groups: {shd_copied}")
    print(f"  STQ processed  : {stq_count}")
    print(f"  LET saved      : {let_count}")


if __name__ == "__main__":
    main()
