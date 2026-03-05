"""
Generates a monthly drilling invoice Excel workbook.

Layout per contractor sheet:
  - Header block (project, contractor, period)
  - Borehole Drilling table (per hole: hole id, meters, rate, amount)
  - Borehole subtotal
  - Standby Hours section (grouped by rig: rig sub-header, entries, rig subtotal)
  - Standby subtotal
  - Deductions table (PPE or Diesel)
  - Deduction subtotal
  - Net Payable
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from itertools import groupby

import db.models as m

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

# ── Color palette ──────────────────────────────────────────────────────────────
C_HEADER_BG   = "1E2A38"
C_SECTION_BG  = "37474F"
C_COL_HDR_BG  = "1565C0"
C_STANDBY_BG  = "E65100"   # orange — standby section
C_RIG_BG      = "FBE9E7"   # light orange — rig sub-header
C_DEDUCT_BG   = "B71C1C"
C_SUBTOTAL_BG = "E3F2FD"
C_NET_BG      = "E8F5E9"
C_WHITE       = "FFFFFF"
C_BLACK       = "000000"

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
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = thin_border()
    if fmt:
        cell.number_format = fmt
    return cell


def _section_header(ws, row, col_start, col_end, label, bg=C_SECTION_BG):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row,   end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=label)
    cell.font = Font(bold=True, color=C_WHITE, size=12)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = thin_border()
    return row + 1


def generate_invoice(project_id, contractor_id, month, year, output_path):
    project_rows = [p for p in m.get_projects() if p["id"] == project_id]
    if not project_rows:
        raise ValueError("Project not found.")
    project = project_rows[0]

    wb = Workbook()
    wb.remove(wb.active)

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
    cname  = contractor["name"]
    ctype  = contractor["type"]
    rate_m = contractor["rate_per_meter"]
    rate_s = contractor["standby_hour_rate"]

    borehole_entries = m.get_drilling_entries(contractor["id"], month, year)
    standby_entries  = m.get_standby_entries(contractor["id"], month, year)
    ppe_rows = m.get_ppe_charges(contractor["id"], month, year)   if ctype == "underground" else []
    dsl_rows = m.get_diesel_charges(contractor["id"], month, year) if ctype == "surface"     else []

    ws = wb.create_sheet(title=cname[:31])
    ws.sheet_view.showGridLines = False

    # Col layout: # | Hole/Rig/Item | Description | Qty | Rate | (blank) | Amount
    NCOLS = 7
    col_widths = [5, 22, 22, 12, 13, 5, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    # ── Invoice header ─────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 40
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
    c = ws.cell(row=row, column=1, value="DRILLING INVOICE")
    c.font = Font(bold=True, size=18, color=C_WHITE)
    c.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    for label, value in [
        ("Project",    project["name"]),
        ("Location",   project["location"] or ""),
        ("Contractor", cname),
        ("Type",       ctype.capitalize()),
        ("Period",     f"{MONTHS[month - 1]} {year}"),
        ("Generated",  date.today().strftime("%d %b %Y")),
    ]:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        _cell(ws, row, 1, label, bold=True, bg="ECEFF1", align="right")
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=NCOLS)
        _cell(ws, row, 3, value)
        row += 1

    row += 1  # spacer

    # ── 1. BOREHOLE DRILLING ───────────────────────────────────────────────────
    row = _section_header(ws, row, 1, NCOLS, "  BOREHOLE DRILLING")

    hdrs = ["#", "Hole ID", "", "Meters Drilled", "Rate / m", "", "Amount"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_COL_HDR_BG, align="center")
    row += 1

    total_meters = total_borehole = 0.0
    for idx, e in enumerate(borehole_entries, 1):
        meters = e["meters_drilled"]
        amount = meters * rate_m
        total_meters    += meters
        total_borehole  += amount

        _cell(ws, row, 1, idx, align="center")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        _cell(ws, row, 2, e["hole_id"])
        _cell(ws, row, 4, meters, align="right", fmt="#,##0.00")
        _cell(ws, row, 5, rate_m, align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 6, "")
        _cell(ws, row, 7, amount, align="right", fmt='"$"#,##0.00', bold=True)
        row += 1

    if not borehole_entries:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        _cell(ws, row, 1, "No borehole entries for this period.", italic=True, align="center")
        row += 1

    # Borehole subtotal
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    _cell(ws, row, 1, "BOREHOLE SUBTOTAL", bold=True, bg=C_SUBTOTAL_BG, align="right")
    _cell(ws, row, 4, total_meters, align="right", fmt="#,##0.00", bold=True, bg=C_SUBTOTAL_BG)
    _cell(ws, row, 5, "", bg=C_SUBTOTAL_BG)
    _cell(ws, row, 6, "", bg=C_SUBTOTAL_BG)
    _cell(ws, row, 7, total_borehole, align="right", fmt='"$"#,##0.00', bold=True, bg=C_SUBTOTAL_BG)
    row += 2

    # ── 2. STANDBY HOURS ──────────────────────────────────────────────────────
    row = _section_header(ws, row, 1, NCOLS, "  STANDBY HOURS", bg=C_STANDBY_BG)

    hdrs = ["#", "Rig Name", "Description / Reason", "Hours", "Rate / hr", "", "Amount"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_STANDBY_BG, align="center")
    row += 1

    total_standby = 0.0

    # Group by rig_name
    sorted_standby = sorted(standby_entries, key=lambda e: e["rig_name"])
    for rig_name, rig_entries in groupby(sorted_standby, key=lambda e: e["rig_name"]):
        rig_entries = list(rig_entries)

        # Rig sub-header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        c = ws.cell(row=row, column=1, value=f"  Rig: {rig_name}")
        c.font = Font(bold=True, color="BF360C", size=10)
        c.fill = PatternFill("solid", fgColor=C_RIG_BG)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border()
        row += 1

        rig_hours = rig_amount = 0.0
        for idx, e in enumerate(rig_entries, 1):
            hours  = e["hours"]
            amount = hours * rate_s
            rig_hours  += hours
            rig_amount += amount
            total_standby += amount

            _cell(ws, row, 1, idx, align="center")
            _cell(ws, row, 2, e["rig_name"])
            _cell(ws, row, 3, e["description"])
            _cell(ws, row, 4, hours,  align="right", fmt="#,##0.00")
            _cell(ws, row, 5, rate_s, align="right", fmt='"$"#,##0.00')
            _cell(ws, row, 6, "")
            _cell(ws, row, 7, amount, align="right", fmt='"$"#,##0.00', bold=True)
            row += 1

        # Rig subtotal
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        _cell(ws, row, 1, f"  Subtotal — {rig_name}", bold=True, bg="FFF3E0", align="right",
              fg="BF360C")
        _cell(ws, row, 4, rig_hours,  align="right", fmt="#,##0.00", bold=True, bg="FFF3E0")
        _cell(ws, row, 5, "",  bg="FFF3E0")
        _cell(ws, row, 6, "",  bg="FFF3E0")
        _cell(ws, row, 7, rig_amount, align="right", fmt='"$"#,##0.00', bold=True, bg="FFF3E0",
              fg="BF360C")
        row += 1

    if not sorted_standby:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        _cell(ws, row, 1, "No standby entries for this period.", italic=True, align="center")
        row += 1

    # Standby grand subtotal
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    _cell(ws, row, 1, "STANDBY SUBTOTAL", bold=True, bg="FFE0B2", align="right", fg=C_STANDBY_BG)
    _cell(ws, row, 7, total_standby, align="right", fmt='"$"#,##0.00', bold=True,
          bg="FFE0B2", fg=C_STANDBY_BG)
    row += 2

    # ── 3. DEDUCTIONS ─────────────────────────────────────────────────────────
    total_deductions = 0.0
    deduct_label = "PPE CHARGES (DEDUCTION)" if ctype == "underground" else "DIESEL CHARGES (DEDUCTION)"
    row = _section_header(ws, row, 1, NCOLS, f"  {deduct_label}", bg=C_DEDUCT_BG)

    hdrs = ["#", "Item / Description", "", "Quantity", "Unit Price", "", "Total"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_DEDUCT_BG, align="center")
    row += 1

    charge_rows = ppe_rows if ctype == "underground" else dsl_rows
    for idx, r in enumerate(charge_rows, 1):
        name  = r["item_name"] if ctype == "underground" else r["description"]
        qty   = r["quantity"]
        price = r["unit_price"]
        total = qty * price
        total_deductions += total

        _cell(ws, row, 1, idx, align="center")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        _cell(ws, row, 2, name)
        _cell(ws, row, 4, qty,   align="right", fmt="#,##0.00")
        _cell(ws, row, 5, price, align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 6, "")
        _cell(ws, row, 7, total, align="right", fmt='"$"#,##0.00', bold=True)
        row += 1

    if not charge_rows:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
        _cell(ws, row, 1, "No charges for this period.", italic=True, align="center")
        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    _cell(ws, row, 1, "DEDUCTION SUBTOTAL", bold=True, bg="FFEBEE", align="right")
    _cell(ws, row, 7, total_deductions, align="right", fmt='"$"#,##0.00',
          bold=True, bg="FFEBEE", fg=C_DEDUCT_BG)
    row += 2

    # ── Net Payable ────────────────────────────────────────────────────────────
    total_work = total_borehole + total_standby
    net = total_work - total_deductions

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
        "name":       cname,
        "type":       ctype,
        "boreholes":  total_borehole,
        "standby":    total_standby,
        "drilling":   total_work,
        "deductions": total_deductions,
        "net":        net,
    })


def _build_summary_sheet(wb, summary_data, project, month, year):
    ws = wb.create_sheet(title="Summary", index=0)
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([28, 13, 14, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    NCOLS = 6
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

    hdrs = ["Contractor", "Type", "Boreholes", "Standby", "Deductions", "Net Payable"]
    for col, h in enumerate(hdrs, 1):
        _cell(ws, row, col, h, bold=True, fg=C_WHITE, bg=C_COL_HDR_BG, align="center")
    row += 1

    grand_bh = grand_sb = grand_ded = grand_net = 0.0
    for s in summary_data:
        _cell(ws, row, 1, s["name"])
        _cell(ws, row, 2, s["type"].capitalize(), align="center")
        _cell(ws, row, 3, s["boreholes"],  align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 4, s["standby"],    align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 5, s["deductions"], align="right", fmt='"$"#,##0.00')
        _cell(ws, row, 6, s["net"],        align="right", fmt='"$"#,##0.00', bold=True)
        grand_bh  += s["boreholes"]
        grand_sb  += s["standby"]
        grand_ded += s["deductions"]
        grand_net += s["net"]
        row += 1

    _cell(ws, row, 1, "TOTAL", bold=True, bg=C_NET_BG, align="right")
    _cell(ws, row, 2, "", bg=C_NET_BG)
    _cell(ws, row, 3, grand_bh,  align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG)
    _cell(ws, row, 4, grand_sb,  align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG)
    _cell(ws, row, 5, grand_ded, align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG)
    _cell(ws, row, 6, grand_net, align="right", fmt='"$"#,##0.00', bold=True, bg=C_NET_BG,
          fg="1B5E20")
