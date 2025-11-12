import os
import re
import sys
import win32com.client as win32

# ===== KONFİQURASİYA =====
MAILBOX_NAME = "spp2dcc@kolin.com.tr"
SUBPATH = r"Inbox\TO PROYAPI\TRN"

# Saxlanılacaq qovluq (sənin verdiyin ünvan)
SAVE_DIR = r"\\10.10.8.253\DataServer\Teknik Ofis\Huseyn Cavid\Software Cavid"

# Boş set() = bütün uzantılar; istəsən məhdudlaşdır: {".pdf", ".xlsx", ".xls", ".docx", ".dwg"}
ONLY_EXTENSIONS = set()

# True etsən, hər maili ayrıca alt qovluğa (YYYY-MM-DD\Subject) saxlayar
ORGANIZE_BY_DATE_AND_SUBJECT = False
# =========================

SAFE_CHARS = re.compile(r'[^A-Za-z0-9._ -]')

def sanitize(s: str, maxlen=120) -> str:
    """Fayl/klasör adını təhlükəsizləşdir."""
    if not s:
        return "No Name"
    s = s.replace("\r", " ").replace("\n", " ").strip()
    s = SAFE_CHARS.sub("_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:maxlen] if len(s) > maxlen else s

def unique_path(path):
    """Eyni ad varsa name (1).ext kimi unikallaşdır."""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 1
    while True:
        cand = f"{root} ({i}){ext}"
        if not os.path.exists(cand):
            return cand
        i += 1

def get_mailbox(ns, mailbox_name: str):
    """Mailbox-ı adı ilə tap (case-insensitive)."""
    target = mailbox_name.strip().lower()
    for i in range(1, ns.Folders.Count + 1):
        store = ns.Folders.Item(i)
        if store.Name.strip().lower() == target:
            return store
    return None

def child_by_name_ci(parent, name):
    """Alt qovluğu ada görə tap (case-insensitive)."""
    want = name.strip().lower()
    for i in range(1, parent.Folders.Count + 1):
        f = parent.Folders.Item(i)
        if f.Name.strip().lower() == want:
            return f
    return None

def get_folder_by_path(root, path_str: str):
    """Root-dan başlayıb 'A\\B\\C' yolunu travers edir."""
    parts = [p for p in path_str.split("\\") if p.strip()]
    folder = root
    if parts and parts[0].strip().lower() == "inbox":
        folder = root.Folders["Inbox"]
        parts = parts[1:]
    for part in parts:
        nxt = child_by_name_ci(folder, part)
        if not nxt:
            available = [folder.Folders.Item(i+1).Name for i in range(folder.Folders.Count)]
            raise RuntimeError(f"'{part}' tapılmadı. Bu səviyyədə olanlar: {available}")
        folder = nxt
    return folder

def save_attachments_from_folder(folder):
    """Verilən Outlook qovluğundakı bütün maillərin attachmentlərini SAVE_DIR-ə saxla."""
    os.makedirs(SAVE_DIR, exist_ok=True)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # ən yenilər üstdə
    total = items.Count
    print(f"📬 Tapıldı: {total} mail. Saxlama qovluğu: {SAVE_DIR}\n")

    saved_files = 0
    processed_mails = 0

    for m in items:
        if getattr(m, "Class", None) != 43:  # 43 = MailItem
            continue

        subject = getattr(m, "Subject", "") or "(No Subject)"
        received = getattr(m, "ReceivedTime", None)
        atts = getattr(m, "Attachments", None)
        if not atts or atts.Count == 0:
            continue

        # Hədəf qovluq: düz (flat) və ya tarix/subject ilə
        target_dir = SAVE_DIR
        if ORGANIZE_BY_DATE_AND_SUBJECT and received:
            day = f"{received.year:04d}-{received.month:02d}-{received.day:02d}"
            target_dir = os.path.join(SAVE_DIR, day, sanitize(subject, 100))
        os.makedirs(target_dir, exist_ok=True)

        wrote_any = False
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            name = att.FileName or "attachment"
            ext = os.path.splitext(name)[1].lower()

            if ONLY_EXTENSIONS and ext not in ONLY_EXTENSIONS:
                continue

            safe_name = sanitize(name, 180)
            dst = unique_path(os.path.join(target_dir, safe_name))

            try:
                att.SaveAsFile(dst)
                saved_files += 1
                wrote_any = True
                print(f"✅ {dst}")
            except Exception as e:
                print(f"⚠️ Saxlanmadı: {name} → {e}")

        if wrote_any:
            processed_mails += 1

    print(f"\n🏁 Bitdi. Mail işlənib: {processed_mails}, fayl saxlanıb: {saved_files}")

def main():
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        store = get_mailbox(outlook, MAILBOX_NAME)
        if not store:
            print(f"❌ Mailbox tapılmadı: {MAILBOX_NAME}")
            print("➡️ Mövcud mailbox-lar:")
            for i in range(1, outlook.Folders.Count + 1):
                print(" -", outlook.Folders.Item(i).Name)
            sys.exit(1)

        trn = get_folder_by_path(store, SUBPATH)
        save_attachments_from_folder(trn)
        sys.exit(0)

    except PermissionError:
        print("❌ İcazə xətası: Şəbəkə qovluğuna yazma icazən yoxdur.")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Xəta: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
