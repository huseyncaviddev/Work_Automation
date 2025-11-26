import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

THIN = Side(border_style="thin", color="000000")

# Faylın yaranacağı yer
OUTPUT_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = OUTPUT_DIR / "SPP2-KLN-PRO-TRN-0164.xlsx"


def create_trn_0164_excel(output_path: Path = OUTPUT_PATH):

    wb = Workbook()
    ws = wb.active
    ws.title = "Transmittal"

    # ================== COLUMN WIDTHS ==================
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 10

    # ================== HEADER ==================
    # Row heights
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 4

    # ✔ GREEN BAR — A1:F1 MERGED
    ws.merge_cells("A1:F1")
    ws["A1"] = ""
    ws["A1"].fill = PatternFill(start_color="91D050", end_color="91D050", fill_type="solid")
    ws["A1"].border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)

    # ======== Header Blocks (Logos + Title) ========
    # LEFT BLOCK — VEKTORDS
    ws.merge_cells("A2:B3")
    ws["A2"] = "VEKTORDS\n(LOGO)"
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A2"].font = Font(bold=True)

    # CENTER BLOCK — TITLE
    ws.merge_cells("C2:D3")
    ws["C2"] = "SITALCHAY 2 PRODUCTION PLANT\nDOCUMENTATION TRANSMITTAL"
    ws["C2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["C2"].font = Font(size=12, bold=True)

    # RIGHT BLOCK — PROYAPI/PROKON
    ws.merge_cells("E2:F3")
    ws["E2"] = "PROYAPI / PROKON\n(LOGOS)"
    ws["E2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["E2"].font = Font(bold=True)

    # Border around header blocks
    for r in range(2, 4):
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)

    # ======== INFO ROW 1 (DATE / TRN / PAGE / REV) ========
    ws.merge_cells("A4:B4")
    ws["A4"] = "DATE: 29-Jul-2025"
    ws["A4"].font = Font(bold=True)

    ws.merge_cells("C4:D4")
    ws["C4"] = "TRANSMITTAL NUMBER: SPP2-KLN-PRO-TRN-0164"
    ws["C4"].font = Font(bold=True)

    ws["E4"] = "Page 1 of 1"
    ws["E4"].alignment = Alignment(horizontal="center")

    ws["F4"] = "Rev.03"
    ws["F4"].alignment = Alignment(horizontal="center")

    # ======== INFO ROW 2 (PROJECT / LOCATION) ========
    ws.merge_cells("A5:C5")
    ws["A5"] = "PROJECT: SPP2 - SITALCHAY 2 PRODUCTION PLANT"
    ws["A5"].font = Font(bold=True)

    ws.merge_cells("D5:F5")
    ws["D5"] = "LOCATION: SUMGAIT AZERBAIJAN"
    ws["D5"].font = Font(bold=True)

    # Borders for info rows
    for r in range(4, 6):
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)

    # ================== FROM / TO BLOCKS ==================
    ws.merge_cells("A7:C7")
    ws["A7"] = "From:"
    ws["A7"].font = Font(bold=True)

    ws.merge_cells("D7:F7")
    ws["D7"] = "To:"
    ws["D7"].font = Font(bold=True)

    from_block = (
        '"KOLIN" İNŞAAT SANAYİ VE TİCARET A.Ş\n'
        "Teoman Uludag\n"
        "Project Manager\n"
        "tuludag@kolin.com.tr"
    )

    to_block = (
        '"PROYAPI/PROKON" JV\n'
        "Mesut Sorgec\n"
        "Project Manager\n"
        "mesutsorgec@proyapimusavirlik.com"
    )

    ws.merge_cells("A8:C11")
    ws["A8"] = from_block
    ws["A8"].alignment = Alignment(wrap_text=True, vertical="top")

    ws.merge_cells("D8:F11")
    ws["D8"] = to_block
    ws["D8"].alignment = Alignment(wrap_text=True, vertical="top")

    # Borders
    for r in range(7, 12):
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)

    # ================== DOCUMENT LIST HEADER ==================
    ws.merge_cells("A13:F13")
    ws["A13"] = "DOCUMENT LIST"
    ws["A13"].font = Font(bold=True)
    ws["A13"].alignment = Alignment(horizontal="center")

    header_row = 15
    headers = ["#", "Document Number", "Format", "Rev.", "Issue Code", "Document Title"]

    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)
        cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

    # ================== DOCUMENTS ==================
    docs = [
        (1, "KLN-SPP2-ITP-CV-GN00-201", "PDF", "05", "IFA",
         "Inspection And Test Plan For Concrete And Insulation Works"),
        (2, "KLN-SPP2-MAR-AR-GN00-037", "PDF", "00", "IFA",
         "VesnaMetal Jacketing Starting U Profile"),
        (3, "KLN-SPP2-MAR-AR-GN00-038", "PDF", "00", "IFA",
         "Aluminium Verticale Profil 140/120/100/80"),
        (4, "KLN-SPP2-MAR-AR-GN00-039", "PDF", "00", "IFA",
         "Knauf Corner Profile"),
        (5, "KLN-SPP2-MAR-CV-GN00-065", "PDF", "00", "IFA",
         "Razor Wire"),
        (6, "KLN-SPP2-MAR-MC-GN00-072", "PDF", "00", "IFA",
         "Pipe Grooved Couplings"),
        (7, "KLN-SPP2-MAR-MC-GN00-073", "PDF", "00", "IFA",
         "Flexible Air Ducts"),
        (8, "*END*", "", "", "", ""),
    ]

    row = header_row + 1
    for no, doc_no, fmt, rev, issue, title in docs:
        ws.cell(row=row, column=1, value=no)
        ws.cell(row=row, column=2, value=doc_no)
        ws.cell(row=row, column=3, value=fmt)
        ws.cell(row=row, column=4, value=rev)
        ws.cell(row=row, column=5, value=issue)
        ws.cell(row=row, column=6, value=title)

        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)
            if col == 6:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = Alignment(horizontal="center")

        row += 1

    # ================== FOOTER ==================
    attach_row = row + 2
    ws.merge_cells(f"A{attach_row}:F{attach_row}")
    ws[f"A{attach_row}"] = "Attachment : ITP, MAR"

    footer_row = attach_row + 2
    ws.merge_cells(f"A{footer_row}:F{footer_row}")
    ws[f"A{footer_row}"] = "VektorDS LLC | U.Hajibeyli str., 62, Baku, Azerbaijan. info@vektords.az"

    footer_row2 = footer_row + 2
    ws.merge_cells(f"A{footer_row2}:F{footer_row2}")
    ws[f"A{footer_row2}"] = (
        "Status Code: A=Accepted, AC=Accepted with Comments, CR=Commented-Resubmit, "
        "NA=Not Accepted, ADV=Advanced Copy, IFD=Issued For Design, IFI=Issued For Information, "
        "IFR=Issued For Review, IFA=Issued For Approval, IFC=Issued For Construction"
    )
    ws[f"A{footer_row2}"].alignment = Alignment(wrap_text=True)


    wb.save(output_path)
    print(f"UGURLA YARADILDI → {output_path}")


def main():
    print("Working dir:", Path.cwd())
    print("Excel faylı:", OUTPUT_PATH)
    create_trn_0164_excel(OUTPUT_PATH)


if __name__ == "__main__":
    main()
