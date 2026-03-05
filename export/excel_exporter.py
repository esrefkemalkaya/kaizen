"""
Generates a monthly drilling invoice Excel workbook.

Layout per contractor sheet:
  - Header block (project, contractor, period, invoice #)
  - Drilling Details table (per hole)
  - Drilling Subtotals
  - Deductions table (PPE or Diesel depending on contractor type)
  - Deduction Subtotal
  - Net Payable
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from datetime import date

import db.models as m

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

# ── Color palette ─────────────────────────────────────────────────────────────
C_HEADER_BG   = "1E2A38"   # dark navy — invoice header
C_SECTION_BG  = "37474F"   # dark grey — section titles
C_COL_HDR_BG  = "1565C0"   # blue — column headers
C_DEDUCT_BG   = "B71C1C"   # dark red — deduction header
C_SUBTOTAL_BG = "E3F2FD"   # light blue — subtotal rows
C_NET_BG      = "E8F5E9"   # light green — net payable row
C_WHITE       = "FFFFFF"
C_BLACK       = "000000"

# ── Border helpers ────────────────────────────────────────────────────────────
_thin  = Side(style="thin",   color=C_BLACK)
_thick = Side(style="medium", color=C_BLACK)

def thin_border():
    return Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def thick_bottom():
    return Border(left=_thin, right=_thin, top=_thin, bottom=_thick)


def _cell(ws, row, col, value="", bold=False, italic=False,
          fg=C_BLACK, bg=None, align="left", fmt=None, size=11):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, italic=italic, color=fg, size=size)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(
        horizontal=align, vertical="center", wrap_text=True
    )
    cell.border = thin_border()
    if fmt:
        cell.number_format = fmt
    return cell


def _section_header(ws, row, col_start, col_end, label, bg=C_SECTION_BG):
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row,   end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=label)
    cell.font = Font(bold=True, color=C_WHITE, size=12)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border()
    return row + 1


def generate_invoice(
    project_id: int,
    contractor_id: int,
    month: int,
    year: int,
    output_path: str
):
    project_rows = [p for p in m.get_projects() if p["id"] == project_id]
    if not project_rows:
        raise ValueError("Project not found.")
    project = project_rows[0]

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    contractors = m.get_contractors(project_id)
    summary_data = []

    for contractor in contractors:
        if contractor_id != -1 and contractor["id"] != contractor_id:
            continue
        _build_sheet(wb, project, contractor, month, year, summary_data)

    if len(summary_data) > 1:
        _build_summary_sheet(wb, summary_data, project, month, year)

    if not wb.worksheets:
        raise ValueError("No data to export.")

    wb.save(output_path)


def _build_sheet(wb, project, contractor, month, year, summary_data):
    cname = contractor["name"]
    ctype = contractor["type"]
    rate_m = contractor["rate_per_meter"]
    rate_s = contractor["standby_hour_rate"]

    entries  = m.get_drilling_entries(contractor["id"], month, year)
    ppe_rows = m.get_ppe_charges(contractor["id"], month, year)    if ctype == "underground" else []
    dsl_rows = m.get_diesel_charges(contractor["id"], month, year) if ctype == "surface"     else []

    sheet_name = cname[:31]  # Excel sheet name limit
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_view.showGridLines = False

    # Column widths
    col_widths = [6, 30, 14, 14, 14, 14, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    NCOLS = 7
    row = 1

    # ── Invoice header block ──────────────────────────────────────────────────
    ws.row_dimensions[row].height = 40
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
    c = ws.cell(row=row, column=1, value="DRILLING INVOICE")
    c.font  = Font(bold=True, size=18, color=C_WHITE)
    c.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    meta = [
        ("Project",    project["name"]),
        ("Location",   project["location"] or ""),
        ("Contractor", cname),
        ("Type",       ctype.capitalize()),
        ("Period",     f"{MONTHS[month - 1]} {year}"),
        ("Generated",  date.today().strftime("%d %b %Y")),
    ]
    for label, value in meta:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        _cell(ws, row, 1, label, bold=True, bg="ECEFF1", align="right")
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=NCOLS)
        _cell(ws, row, 3, value, align="left")
        row += 1

    row += 1  # spacer

    # ── Drilling details section ──────────────────────────────────────────────
    row = _section_header(ws, row, 1, NCOLS, "  DRILLING DETAILS")

    # Column headers
    hdrs = ["#", "Hole ID", "Meters Drilled", "Standby Hours",
            "Rate/m", "Standby Rate/hr", "Total Amount"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_COL_HDR_BG, align="center")
    row += 1

    total_meters   = 0.0
    total_standby  = 0.0
    total_drill    = 0.0

    for idx, e in enumerate(entries, 1):
        meters  = e["meters_drilled"]
        standby = e["standby_hours"]
        amount  = meters * rate_m + standby * rate_s
        total_meters  += meters
        total_standby += standby
        total_drill   += amount

        _cell(ws, row, 1, idx,       align="center")
        _cell(ws, row, 2, e["hole_id"])
        _cell(ws, row, 3, meters,    align="right", fmt="#,##0.00")
        _cell(ws, row, 4, standby,   align="right", fmt="#,##0.00")
        _cell(ws, row, 5, rate_m,    align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 6, rate_s,    align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 7, amount,    align="right", fmt='"$"#,##0.00', bold=True)
        row += 1

    # Drilling subtotal
    bg_sub = "E3F2FD"
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    _cell(ws, row, 1, "DRILLING SUBTOTAL", bold=True, bg=bg_sub, align="right")
    _cell(ws, row, 6, total_meters,  align="right", fmt="#,##0.00",   bold=True, bg=bg_sub)
    _cell(ws, row, 7, total_drill,   align="right", fmt='"$"#,##0.00', bold=True, bg=bg_sub)
    row += 2

    # ── Deductions section ────────────────────────────────────────────────────
    total_deductions = 0.0
    deduct_label = "PPE CHARGES (DEDUCTION)" if ctype == "underground" else "DIESEL CHARGES (DEDUCTION)"
    row = _section_header(ws, row, 1, NCOLS, f"  {deduct_label}", bg=C_DEDUCT_BG)

    deduct_hdrs = ["#", "Description" if ctype == "surface" else "Item", "", "Quantity", "Unit Price", "", "Total"]
    for col, h in enumerate(deduct_hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_DEDUCT_BG, align="center")
    row += 1

    charge_rows = ppe_rows if ctype == "underground" else dsl_rows
    for idx, r in enumerate(charge_rows, 1):
        name  = r["item_name"]  if ctype == "underground" else r["description"]
        qty   = r["quantity"]
        price = r["unit_price"]
        total = qty * price
        total_deductions += total

        _cell(ws, row, 1, idx,   align="center")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        _cell(ws, row, 2, name)
        _cell(ws, row, 4, qty,   align="right", fmt="#,##0.00")
        _cell(ws, row, 5, price, align="right", fmt='"$"#,##0.00')
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=6)
        _cell(ws, row, 6, "")
        _cell(ws, row, 7, total, align="right", fmt='"$"#,##0.00', bold=True)
        row += 1

    if not charge_rows:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        _cell(ws, row, 1, "No charges for this period.", italic=True, align="center")
        row += 1

    # Deduction subtotal
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    _cell(ws, row, 1, "DEDUCTION SUBTOTAL", bold=True, bg="FFEBEE", align="right")
    _cell(ws, row, 7, total_deductions, align="right", fmt='"$"#,##0.00',
          bold=True, bg="FFEBEE", fg=C_DEDUCT_BG)
    row += 2

    # ── Net payable ───────────────────────────────────────────────────────────
    net = total_drill - total_deductions
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row=row, column=1, value="NET PAYABLE TO CONTRACTOR")
    c.font = Font(bold=True, size=13, color=C_WHITE)
    c.fill = PatternFill("solid", fgColor="2E7D32")
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = thick_bottom()

    nc = ws.cell(row=row, column=7, value=net)
    nc.font = Font(bold=True, size=13, color="1B5E20")
    nc.fill = PatternFill("solid", fgColor=C_NET_BG)
    nc.alignment = Alignment(horizontal="right", vertical="center")
    nc.border = thick_bottom()
    nc.number_format = '"$"#,##0.00'
    ws.row_dimensions[row].height = 28

    summary_data.append({
        "name":        cname,
        "type":        ctype,
        "drilling":    total_drill,
        "deductions":  total_deductions,
        "net":         net,
    })


def _build_summary_sheet(wb, summary_data, project, month, year):
    ws = wb.create_sheet(title="Summary", index=0)
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([30, 14, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    NCOLS = 5
    row = 1

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    c = ws.cell(row=1, column=1, value=f"MONTHLY SUMMARY — {MONTHS[month-1].upper()} {year}")
    c.font = Font(bold=True, size=16, color=C_WHITE)
    c.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35
    row = 2

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
    _cell(ws, row, 1, f"Project: {project['name']}", bold=True, align="center", bg="ECEFF1")
    row += 2

    hdrs = ["Contractor", "Type", "Drilling Total", "Deductions", "Net Payable"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_COL_HDR_BG, align="center")
    row += 1

    grand_drill = grand_deduct = grand_net = 0.0
    for s in summary_data:
        _cell(ws, row, 1, s["name"])
        _cell(ws, row, 2, s["type"].capitalize(), align="center")
        _cell(ws, row, 3, s["drilling"],   align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 4, s["deductions"], align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 5, s["net"],        align="right", fmt='"$"#,##0.00', bold=True)
        grand_drill  += s["drilling"]
        grand_deduct += s["deductions"]
        grand_net    += s["net"]
        row += 1

    # Grand total row
    _cell(ws, row, 1, "TOTAL", bold=True, bg=C_NET_BG, align="right")
    _cell(ws, row, 2, "", bg=C_NET_BG)
    _cell(ws, row, 3, grand_drill,  align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG)
    _cell(ws, row, 4, grand_deduct, align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG)
    _cell(ws, row, 5, grand_net,    align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG,
          fg="1B5E20")
