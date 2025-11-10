# folder_creator.py
from pathlib import Path
import re
import sys

# 1) Define the base path
BASE_PATH = Path(r'\\10.10.8.253\DataServer\STP-S2-Projeler\Log\1. Outgoing\1. TRN')


# 2) TRN folder name pattern
TRN_PATTERN = re.compile(r'^SPP2-KLN-PRO-TRN-(\d{4})$')


def find_next_folder_name(base: Path) -> str:
    if not base.exists():
        raise FileNotFoundError(f"Path not found: {base}")

    max_num = -1

    # Bütün qovluqlara baxıb şablona uyğun olanları götürmək
    for entry in base.iterdir():
        if entry.is_dir():
            m = TRN_PATTERN.match(entry.name)
            if m: 
                num = int(m.group(1))
                if num > max_num:
                    max_num = num
    
    # Heç uyğun qovluq yoxdursa, 0000 -dan başla
    next_num = (max_num + 1) if max_num >= 0 else 0
    return f"SPP2-KLN-PRO-TRN-{next_num:04d}" 


def main():
    try:
        next_name = find_next_folder_name(BASE_PATH)
        print(next_name, 'next_name')
        target = BASE_PATH / next_name
        if target.exists():
            print(f'Qovluq artıq mövcuddur: {target}')
            sys.exit(0)

        target.mkdir(parents=False, exist_ok=False)    
        print(f'✅ Yaradıldı: {target}')

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
