# folder_creator.py
from pathlib import Path
import re
import sys

# 1) Baza yol
BASE_PATH = Path(r'\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN')

# 2) TRN klasör adı şablonu
TRN_PATTERN = re.compile(r'^SPP2-KLN-PRO-TRN-(\d{4})$')

# 3) Oluşturulacak alt klasörler
SUBFOLDERS = ["1. main", "2. attachments", "3. docs"]


def find_next_folder_name(base: Path) -> str:
    if not base.exists():
        raise FileNotFoundError(f"Path not found: {base}")

    max_num = -1
    for entry in base.iterdir():
        if entry.is_dir():
            m = TRN_PATTERN.match(entry.name)
            if m:
                num = int(m.group(1))
                if num > max_num:
                    max_num = num

    next_num = (max_num + 1) if max_num >= 0 else 0
    return f"SPP2-KLN-PRO-TRN-{next_num:04d}"


def ensure_subfolders(root: Path):
    # root içinde SUBFOLDERS listesindeki alt klasörleri oluşturur (eksikse).
    created = []
    for name in SUBFOLDERS:
        p = root / name
        print(p, 'p')
        if not p.exists():
            p.mkdir(parents=False, exist_ok=False)
            created.append(str(p))
    return created


def main():
    try:
        next_name = find_next_folder_name(BASE_PATH)
        target = BASE_PATH / next_name
        print(next_name, 'next_name')

        if not target.exists():
            # Yeni TRN klasörünü oluştur
            target.mkdir(parents=False, exist_ok=False)
            print(f'✅ Yaradıldı: {target}')
        else:
            print(f'ℹ️ Qovluq artıq mövcuddur: {target}')

        # Alt klasörleri oluştur (eksik olanları tamamlar)
        created = ensure_subfolders(target)
        if created:
            print("✅ Aşağıdakı alt qovluqlar yaradıldı:")
            for c in created:
                print("  -", c)
        else:
            print("ℹ️ Bütün alt qovluqlar artıq mövcuddur.")

        # İş bitti
        sys.exit(0)

    except PermissionError:
        print('❌ İcazə xətası: Şəbəkə qovluğunda yazmaq icazəniz olduğundan əmin olun.')
    except FileNotFoundError as e:
        print(f'❌ Yol tapılmadı: {e}')
    except OSError as e:
        print(f'❌ OS xətası: {e}')
    except Exception as e:
        print(f'❌ Gözlənilməz xəta: {e}')


if __name__ == '__main__':
    main()
