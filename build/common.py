"""Stili condivisi dai modelli."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=16, bold=True, color="1B1D21")
SUB = Font(name="Arial", size=10, color="545A64")
MUTED = Font(name="Arial", size=9, color="6B7280")
HEAD_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")

INPUT_FILL = PatternFill("solid", fgColor="FFF6D5")
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
BAND_FILL = PatternFill("solid", fgColor="F3F5F8")

_side = Side(style="thin", color="E3E5E9")
THIN = Border(left=_side, right=_side, top=_side, bottom=_side)

EURO = '#,##0.00 "€";-#,##0.00 "€";-'
PCT = "0.0%"


def header_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 22


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def input_cell(c):
    c.fill = INPUT_FILL
    c.border = THIN
    c.font = FONT
