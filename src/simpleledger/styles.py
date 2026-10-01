"""Visual language shared by every sheet.

One palette, one font, five cell roles (input / auto / report / warning / nav).
Keeping these in one place is what makes the workbook look like a product
instead of a spreadsheet.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONT_NAME = "Arial"

# Palette
NAVY = "1F3A5F"
TEAL = "2A9D8F"
SLATE = "475569"
MUTED = "64748B"
LIGHT = "F1F5F9"
LINE = "CBD5E1"
WHITE = "FFFFFF"
INPUT_FILL = "FFF8DC"   # soft yellow: "type here"
INPUT_FONT = "1D4ED8"   # blue: hard-coded user input
AUTO_FILL = "EEF2F6"    # grey-blue: formula, do not type
OK_GREEN = "15803D"
WARN_AMBER = "B45309"
ERR_RED = "B91C1C"
ERR_FILL = "FDECEA"
OK_FILL = "E8F5EC"
WARN_FILL = "FFF4E5"

# Number formats
NUM = '#,##0.00;(#,##0.00);"-"'
NUM0 = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "dd-mmm-yyyy"
INT = "0"


def font(size=10, bold=False, color=SLATE, italic=False, underline=None):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)


thin = Side(style="thin", color=LINE)
medium = Side(style="medium", color=NAVY)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TOP_RULE = Border(top=thin)
TOTAL_RULE = Border(top=thin, bottom=Side(style="double", color=NAVY))
UNDERLINE = Border(bottom=medium)

LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
INDENT1 = Alignment(horizontal="left", vertical="center", indent=1)
INDENT2 = Alignment(horizontal="left", vertical="center", indent=2)


def style(cell, *, f=None, bg=None, fmt=None, align=None, border=None):
    if f is not None:
        cell.font = f
    if bg is not None:
        cell.fill = fill(bg)
    if fmt is not None:
        cell.number_format = fmt
    if align is not None:
        cell.alignment = align
    if border is not None:
        cell.border = border
    return cell


def as_input(cell, fmt=None):
    from openpyxl.styles import Protection

    style(cell, f=font(color=INPUT_FONT), bg=INPUT_FILL, fmt=fmt, border=BOX)
    cell.protection = Protection(locked=False)
    return cell


def as_auto(cell, fmt=None, bold=False):
    style(cell, f=font(bold=bold, color=SLATE), bg=AUTO_FILL, fmt=fmt)
    return cell
