"""Script to create the sample data/zonas.xlsx file.

Run once: py create_zonas_sample.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

data = [
    # (Nombre, Zona, Región)
    ("VENDEDOR DEMO 01", "ZC-01", "CAPITAL"),
    ("VENDEDOR DEMO 02", "ZC-02", "CAPITAL"),
    ("VENDEDOR DEMO 03", "ZCE-01", "CENTRO"),
    ("VENDEDOR DEMO 04", "ZCE-02", "CENTRO"),
    ("VENDEDOR DEMO 05", "ZO-01", "OCCIDENTE"),
    ("VENDEDOR DEMO 06", "ZA-01", "ANDES"),
    ("VENDEDOR DEMO 07", "ZON-01", "ORIENTE NORTE"),
    ("VENDEDOR DEMO 08", "ZOS-01", "ORIENTE SUR"),
    ("VENDEDOR DEMO 09", "ZL-01", "LLANOS"),
    ("VENDEDOR DEMO 10", "ZCC-01", "CENTRO COSTA"),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Zonas"

# Header
headers = ["Nombre", "Zona", "Región"]
header_fill = PatternFill(fill_type="solid", fgColor="1F3864")
header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
thin = Side(border_style="thin", color="CCCCCC")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = border

for row_idx, (nombre, zona, region) in enumerate(data, 2):
    for col, val in enumerate([nombre, zona, region], 1):
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.border = border
        cell.alignment = Alignment(horizontal="left" if col == 1 else "center")

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 12
ws.column_dimensions["C"].width = 18

out = Path(__file__).parent / "data" / "zonas.xlsx"
out.parent.mkdir(exist_ok=True)
wb.save(out)
print(f"Created: {out}")
