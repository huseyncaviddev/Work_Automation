from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image


THIN = Side(border_style="thin", color="000000")


def apply_border(ws, cell_range: str):
    """Apply thin border to all cells in an A1-style range, e.g. 'A1:Z10'."""
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def add_logos(ws, left_logo_path: Path, right_logo_path: Path):
    """
    Logoları header-in sol və sağ bloklarına əlavə edir.
    Şəkillərin ölçüsünü ehtiyac olsa aşağıda tweak edə bilərsən.
    """
    if left_logo_path.is_file():
        left_img = Image(str(left_logo_path))
        left_img.width = 140
        left_img.height = 50
        # Sol böyük boş blok: A2:F6
        left_img.anchor = "A2"
        ws.add_image(left_img)

    if right_logo_path.is_file():
        right_img = Image(str(right_logo_path))
        right_img.width = 140
        right_img.height = 60
        # Sağ böyük blok: U2:Y6
        right_img.anchor = "U2"
        ws.add_image(right_img)


def safe_save_workbook(wb: Workbook, output_path: Path) -> Path:
    """
    Faylı təhlükəsiz saxlayır:
      - Əgər eyni adda fayl varsa, yanına '_NEW' əlavə edir.
    PermissionError / overwrite problemlərinin qabağını alır.
    """
    output_path = Path(output_path)

    # Eyni adda artıq fayl varsa, yeni adla yaz
    if output_path.exists():
        output_path = output_path.with_name(
            output_path.stem + "_NEW" + output_path.suffix
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def create_trn_excel(
    output_path: Path = Path("SPP2-KLN-PRO-TRN-0164_AUTO.xlsx"),
    trn_no: str = "SPP2-KLN-PRO-TRN-0164",
    date_str: str = "29-Jul-2025",
    left_logo: str = "vektords.png",
    right_logo: str = "proyapi_prokon.png",
):
    wb = Workbook()
    ws = wb.active
    ws.title = "TRN Maker"

    # --- Base defaults ---
    base_font = Font(name="Times New Roman", size=9)
    for row in range(1, 60):
        for col in range(1, 27):
            c = ws.cell(row=row, column=col)
            c.font = base_font
            c.alignment = Alignment(vertical="center", wrap_text=True)

    # Column widths – sənin templatedən təxmini götürülüb
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 4
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 4
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 4
    ws.column_dimensions["K"].width = 4
    ws.column_dimensions["L"].width = 4
    ws.column_dimensions["M"].width = 4
    ws.column_dimensions["N"].width = 4
    ws.column_dimensions["O"].width = 4
    ws.column_dimensions["P"].width = 4
    ws.column_dimensions["Q"].width = 4
    ws.column_dimensions["R"].width = 4
    ws.column_dimensions["S"].width = 4
    ws.column_dimensions["T"].width = 4
    ws.column_dimensions["U"].width = 4
    ws.column_dimensions["V"].width = 4
    ws.column_dimensions["W"].width = 4
    ws.column_dimensions["X"].width = 4
    ws.column_dimensions["Y"].width = 4
    ws.column_dimensions["Z"].width = 4

    # Row heights – səninkinə yaxınlaşdırılıb
    for r in range(11, 16):
        ws.row_dimensions[r].height = 16
    for r in range(19, 25):
        ws.row_dimensions[r].height = 17.5
    for r in range(21, 39):
        ws.row_dimensions[r].height = 18.5
    ws.row_dimensions[49].height = 13.5
    ws.row_dimensions[50].height = 13.5

    # === ÜST BOŞ QUTU ===
    ws.merge_cells("A1:Z1")
    apply_border(ws, "A1:Z1")
    fill_green = PatternFill(start_color="91D050", end_color="91D050", fill_type="solid")
    ws["A1"].fill = fill_green

    # === HEADER / TITLE AREA ===
    # Sol blok (logo üçün)
    ws.merge_cells("A2:E6")
    apply_border(ws, "A2:E6")

    # Başlıq ortada
    ws.merge_cells("F2:U4")
    title = ws["F2"]
    title.value = "SITALCHAY 2 PRODUCTION PLANT\nDOCUMENTATION TRANSMITTAL"
    title.font = Font(name="Times New Roman", size=10, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    apply_border(ws, "F2:U4")
    # Sağ blok (logo + page/rev)
    ws.merge_cells("V2:Z6")
    apply_border(ws, "V2:Z6")

    # TRANSMITTAL NUMBER sahəsi
    ws.merge_cells("F5:M6")
    ws["F5"].value = "TRANSMITTAL  NUMBER:"
    ws["F5"].font = Font(name="Times New Roman", size=10, bold=True)
    ws["F5"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("N5:U6")
    ws["N5"].value = trn_no
    ws["N5"].alignment = Alignment(horizontal="left", vertical="center")
    apply_border(ws, "F5:U6")

    # DATE / PROJECT / LOCATION / PAGE / REV sətiri
    # Sol – Date
    ws.merge_cells("A7:E8")
    ws["A7"].value = f"DATE: {date_str}"
    ws["A7"].alignment = Alignment(horizontal="center", vertical="center")

    # PROJECT
    ws.merge_cells("F7:M8")
    ws["F7"].value = "PROJECT: SPP2 \nSITALCHAY 2 PRODUCTION PLANT "
    ws["F7"].alignment = Alignment(horizontal="left", vertical="center")

    # LOCATION
    ws.merge_cells("N7:U8")
    ws["N7"].value = "LOCATION: \nSUMGAIT AZERBAIJAN "
    ws["N7"].alignment = Alignment(horizontal="left", vertical="center")

   
    # Page & Rev sağda
    ws.merge_cells("V7:X8")
    ws["V7"].value = "Page 1 of 1"
    ws["V7"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("Y7:Z8")
    ws["Y7"].value = "Rev.03"
    ws["Y7"].alignment = Alignment(horizontal="center", vertical="center")

    apply_border(ws, "A2:Z8")

    # === FROM / TO BLOCK ===
    ws.merge_cells("A11:M11")
    ws.merge_cells("N11:Y11")
    ws["A11"].value = "From:"
    ws["N11"].value = "To:"

    ws.merge_cells("A12:M12")
    ws.merge_cells("N12:Y12")
    ws["A12"].value = '   “KOLIN”  İNŞAAT VE TICARET A.Ş'
    ws["N12"].value = '“PROYAPI/PROKON” JV'

    ws.merge_cells("A13:M13")
    ws.merge_cells("N13:Y13")
    ws["A13"].value = "Teoman Uludag"
    ws["N13"].value = "Mesut Sorgec"

    ws.merge_cells("A14:M14")
    ws.merge_cells("N14:Y14")
    ws["A14"].value = "Project Manager"
    ws["N14"].value = "Project Manager"

    ws.merge_cells("A15:M15")
    ws.merge_cells("N15:Y15")
    ws["A15"].value = "tuludag@kolin.com.tr"
    ws["N15"].value = "mesutsorgec@proyapimusavirlik.com"

    apply_border(ws, "A11:Y15")

    # === DOCUMENT LIST TITLE ===
    ws.merge_cells("J17:O17")
    ws["J17"].value = "DOCUMENT LIST"
    ws["J17"].font = Font(name="Calibri", size=10, bold=True)
    ws["J17"].alignment = Alignment(horizontal="center", vertical="center")

    # === DOCUMENT LIST TABLE HEADER ===
    header_fill = PatternFill("solid", fgColor="FFE7E6E6")

    ws.merge_cells("A19:A20")
    ws.merge_cells("B19:G20")
    ws.merge_cells("H19:J20")
    ws.merge_cells("K19:L20")
    ws.merge_cells("M19:N20")
    ws.merge_cells("O19:Y20")

    headers = {
        "A19": "#",
        "B19": "Document Number",
        "H19": "Format",
        "K19": "Rev.",
        "M19": "Issue\nCode",
        "O19": "Document Title",
    }

    for cell_ref, text in headers.items():
        c = ws[cell_ref]
        c.value = text
        c.font = Font(name="Calibri", size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = header_fill

    apply_border(ws, "A19:Y20")

    # === DATA ROWLAR (21–38) ===
    for row in range(21, 39):
        ws.merge_cells(f"B{row}:G{row}")
        ws.merge_cells(f"H{row}:J{row}")
        ws.merge_cells(f"K{row}:L{row}")
        ws.merge_cells(f"M{row}:N{row}")
        ws.merge_cells(f"O{row}:Y{row}")

    # Row 21 – nümunə
    ws["A21"].value = 1
    ws["A21"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B21"].value = "KLN-SPP2-ITP-CV-GN00-201"
    ws["H21"].value = "PDF"
    ws["K21"].value = "00"
    ws["M21"].value = "IFA"
    ws["O21"].value = "Inspection And Test Plan For Concrete And Insulation Works"

    for col in ("B", "H", "K", "M"):
        ws[f"{col}21"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["O21"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Növbəti sətirlər üçün auto nömrələmə formula
    for row in range(22, 39):
        ws[f"A{row}"].value = f"=A{row-1}+1"
        ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")

    # END sətiri
    ws["A39"].value = "=A38+1"
    ws["A39"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("B39:G39")
    ws.merge_cells("H39:J39")
    ws.merge_cells("K39:L39")
    ws.merge_cells("M39:N39")
    ws.merge_cells("O39:Y39")
    ws["B39"].value = "*END*"
    ws["B39"].alignment = Alignment(horizontal="left", vertical="center")

    apply_border(ws, "A21:Y39")

    # === FOOTER ===
    ws.merge_cells("A41:Y41")
    ws["A41"].value = "Attachment: ITP, MAR"
    ws["A41"].alignment = Alignment(horizontal="left", vertical="center")
    apply_border(ws, "A41:Y41")

    ws.merge_cells("A45:Y48")
    ws["A45"].value = (
        "Status Code: A = Accepted, AC = Accepted with Comments, CR = Commented-Resubmit, NA = Not Accepted\n"
        "ADV = Advanced Copy, IFD = Issued For Design, IFI = Issued For Information, IFR = Issued For Review, IFA = Issued For Approval\n"
        "IFC = Issued For Construction"
    )
    ws["A45"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    apply_border(ws, "A45:Y48")

    ws.merge_cells("A51:Y52")
    ws["A51"].value = (
        "VektorDS LLC | U.Hajibeyli str., 62, Baku, Azerbaijan. info@vektords.az\n"
        "This Document is VEKTORDS LLC property and cannot be used by others for any purpose without prior written consent."
    )
    ws["A51"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    apply_border(ws, "A51:Y52")

    # Logoları əlavə et
    script_dir = Path(__file__).resolve().parent
    add_logos(ws, script_dir / left_logo, script_dir / right_logo)

    # Faylı təhlükəsiz saxla
    saved = safe_save_workbook(wb, output_path)
    print(f"TRN Excel yaradıldı: {saved}")


if __name__ == "__main__":
    create_trn_excel()
