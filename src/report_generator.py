import os

from openpyxl import Workbook

from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment,
    numbers
)

from openpyxl.chart import PieChart, Reference
from openpyxl.utils import get_column_letter
from datetime import datetime
from dateutil.relativedelta import relativedelta


# ═══════════════════════════════════════════
# Color Palette — Professional Corporate
# ═══════════════════════════════════════════

# Primary header — deep navy
HEADER_BG = "1B2A4A"
HEADER_FONT_COLOR = "FFFFFF"

# Sub-header / section title
SECTION_TITLE_BG = "2C3E6B"
SECTION_TITLE_FONT = "FFFFFF"

# Alternating row stripes
ROW_EVEN = "F5F7FA"
ROW_ODD = "FFFFFF"

# Total / summary row
TOTAL_ROW_BG = "1B2A4A"
TOTAL_ROW_FONT = "FFFFFF"

# Status indicators
STATUS_INCREASED_BG = "FDEDED"
STATUS_INCREASED_FONT = "C0392B"
STATUS_DECREASED_BG = "E8F8F5"
STATUS_DECREASED_FONT = "1E8449"
STATUS_STABLE_BG = "FEF9E7"
STATUS_STABLE_FONT = "B7950B"

# Account title in Sheet 3
ACCOUNT_TITLE_BG = "2C3E6B"
ACCOUNT_TITLE_FONT = "FFFFFF"

# Sheet tab colors
TAB_COLOR_SUMMARY = "1B2A4A"
TAB_COLOR_OBSERVATION = "2980B9"
TAB_COLOR_HISTORY = "27AE60"

# Report subtitle / metadata
META_FONT_COLOR = "7F8C8D"

# Border color
BORDER_COLOR = "D5D8DC"

# Tax column highlight
TAX_HEADER_BG = "6C3483"
TAX_COL_EVEN = "F5EEF8"
TAX_COL_ODD = "FFFFFF"


# ═══════════════════════════════════════════
# Reusable Style Factories
# ═══════════════════════════════════════════

def _header_font(size=10):
    return Font(
        name="Calibri",
        bold=True,
        size=size,
        color=HEADER_FONT_COLOR
    )


def _header_fill():
    return PatternFill(
        start_color=HEADER_BG,
        end_color=HEADER_BG,
        fill_type="solid"
    )


def _body_font(bold=False, color="2C3E50", size=10):
    return Font(
        name="Calibri",
        bold=bold,
        size=size,
        color=color
    )


def _fill(color):
    return PatternFill(
        start_color=color,
        end_color=color,
        fill_type="solid"
    )


def _thin_border():
    side = Side(style="thin", color=BORDER_COLOR)
    return Border(
        left=side,
        right=side,
        top=side,
        bottom=side
    )


def _bottom_border():
    return Border(
        bottom=Side(style="medium", color=HEADER_BG)
    )


CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(
    horizontal="left",
    vertical="top",
    wrap_text=True
)


# ═══════════════════════════════════════════
# Helper — Apply Alternating Row Stripes
# ═══════════════════════════════════════════

def _apply_stripes(ws, data_start_row, data_end_row, col_count):
    """Apply alternating row colors for readability."""

    for row_idx in range(data_start_row, data_end_row + 1):

        bg = ROW_EVEN if (row_idx - data_start_row) % 2 == 0 else ROW_ODD

        for col_idx in range(1, col_count + 1):

            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = _fill(bg)
            cell.border = _thin_border()
            cell.font = _body_font()


# ═══════════════════════════════════════════
# Helper — Style Header Row
# ═══════════════════════════════════════════

def _style_header_row(ws, row_num, col_count):
    """Apply professional dark header styling."""

    for col_idx in range(1, col_count + 1):

        cell = ws.cell(row=row_num, column=col_idx)
        cell.fill = _header_fill()
        cell.font = _header_font()
        cell.alignment = CENTER
        cell.border = _thin_border()


# ═══════════════════════════════════════════
# Helper — Write Cell Values
# ═══════════════════════════════════════════

def _write_cells(ws, row_num, values):
    """Write a list of values to a row."""

    for col_idx, value in enumerate(values, 1):
        ws.cell(row=row_num, column=col_idx, value=value)


# ═══════════════════════════════════════════
# Helper — Style Total Row
# ═══════════════════════════════════════════

def _style_total_row(ws, row_num, col_count):
    """Apply total/summary row styling."""

    for col_idx in range(1, col_count + 1):

        cell = ws.cell(row=row_num, column=col_idx)
        cell.fill = _fill(TOTAL_ROW_BG)
        cell.font = Font(
            name="Calibri",
            bold=True,
            size=11,
            color=TOTAL_ROW_FONT
        )
        cell.border = _thin_border()


# ═══════════════════════════════════════════
# Helper — Highlight Tax Column Header
# ═══════════════════════════════════════════

def _style_tax_header(ws, row_num, col_idx):
    """Give the Tax header a distinct purple color."""

    cell = ws.cell(row=row_num, column=col_idx)
    cell.fill = _fill(TAX_HEADER_BG)
    cell.font = Font(
        name="Calibri",
        bold=True,
        size=10,
        color="FFFFFF"
    )
    cell.alignment = CENTER
    cell.border = _thin_border()


# ═══════════════════════════════════════════
# Main Report Generator
# ═══════════════════════════════════════════

def create_multi_account_report(results):

    # ---------- Compute Month Names ----------

    today = datetime.today()
    current_month_start = today.replace(day=1)
    previous_month_start = (
        current_month_start - relativedelta(months=1)
    )
    two_months_ago_start = (
        current_month_start - relativedelta(months=2)
    )

    # Current billing month (the last full month)
    curr_month_name = previous_month_start.strftime("%B %Y")
    prev_month_name = two_months_ago_start.strftime("%B %Y")

    report_date = datetime.now().strftime("%B %d, %Y")
    num_accounts = len(results)

    # Use /tmp in Lambda, output/ locally
    if os.environ.get("AWS_LAMBDA_FUNCTION_NAME"):
        output_dir = "/tmp"
    else:
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)

    file_name = (
        f"{output_dir}/aws_report_"
        f"{previous_month_start.strftime('%B_%Y').lower()}"
        f".xlsx"
    )

    # ---------- Build Combined Data ----------

    summary_rows = []
    observation_rows = []
    total_excl = 0
    total_tax = 0
    total_incl = 0

    for i, row in enumerate(results):

        curr_excl = round(
            row["Current Month (Without Tax)"], 2
        )
        prev_excl = round(
            row["Previous Month (Without Tax)"], 2
        )
        curr_incl = round(
            row["Current Month (With Tax)"], 2
        )

        tax = round(curr_incl - curr_excl, 2)
        change_value = round(curr_excl - prev_excl, 2)

        if change_value > 0:
            status = "Increased"
        elif change_value < 0:
            status = "Decreased"
        else:
            status = "Stable"

        if change_value < 0:
            change_text = f"-${abs(change_value):,.2f}"
        else:
            change_text = f"${change_value:,.2f}"

        # Use without-tax reasons (service-level changes)
        reason_parts = []

        for r in row["Reasons (Without Tax)"]:

            service_name = r[0]
            change_amt = r[1]

            if change_amt >= 0:
                reason_parts.append(
                    f"{service_name} (+${change_amt:,.2f})"
                )
            else:
                reason_parts.append(
                    f"{service_name} (-${abs(change_amt):,.2f})"
                )

        reason_text = ", ".join(reason_parts)

        summary_rows.append([
            i + 1,
            row["Account Name"],
            f"${curr_excl:,.2f}",
            f"${tax:,.2f}",
            f"${curr_incl:,.2f}"
        ])

        observation_rows.append([
            row["Account Name"],
            row["Account ID"],
            f"${prev_excl:,.2f}",
            f"${curr_excl:,.2f}",
            change_text,
            status,
            reason_text
        ])

        total_excl += curr_excl
        total_tax += tax
        total_incl += curr_incl

    summary_rows.append([
        "",
        "Total",
        f"${total_excl:,.2f}",
        f"${total_tax:,.2f}",
        f"${total_incl:,.2f}"
    ])

    # ---------- Create Workbook ----------

    wb = Workbook()

    # =================================================================
    # Sheet 1 — Account Cost Summary
    # =================================================================

    ws1 = wb.active
    ws1.title = "Account Cost Summary"
    ws1.sheet_properties.tabColor = TAB_COLOR_SUMMARY

    # Title
    ws1.merge_cells("A1:E1")
    title_cell = ws1["A1"]
    title_cell.value = (
        f"Account Cost Summary — {curr_month_name}"
    )
    title_cell.font = Font(
        name="Calibri",
        bold=True,
        size=16,
        color=HEADER_BG
    )
    title_cell.alignment = LEFT
    title_cell.border = Border()

    # Subtitle
    ws1.merge_cells("A2:E2")
    sub_cell = ws1["A2"]
    sub_cell.value = f"Generated on {report_date}"
    sub_cell.font = Font(
        name="Calibri",
        italic=True,
        size=9,
        color=META_FONT_COLOR
    )
    sub_cell.alignment = LEFT

    # Separator line under subtitle
    for col in range(1, 6):
        ws1.cell(row=3, column=col).border = _bottom_border()

    # Column widths
    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 32
    ws1.column_dimensions["C"].width = 22
    ws1.column_dimensions["D"].width = 16
    ws1.column_dimensions["E"].width = 22

    summary_headers = [
        "SN", "Account Name",
        "Cost (Excl. Tax)", "Tax",
        "Total (Incl. Tax)"
    ]

    # Header row (row 4)
    _write_cells(ws1, 4, summary_headers)
    _style_header_row(ws1, 4, 5)
    _style_tax_header(ws1, 4, 4)  # Tax column purple

    # Data rows
    row = 5
    data_start = row

    for data_row in summary_rows:
        _write_cells(ws1, row, data_row)
        row += 1

    data_end = row - 1
    total_row = data_end

    # Stripes (exclude total row)
    if num_accounts > 0:
        _apply_stripes(ws1, data_start, total_row - 1, 5)

    # Alignment
    for row_idx in range(data_start, total_row + 1):

        ws1.cell(row=row_idx, column=1).alignment = CENTER
        ws1.cell(row=row_idx, column=2).alignment = LEFT
        ws1.cell(row=row_idx, column=3).alignment = RIGHT
        ws1.cell(row=row_idx, column=4).alignment = RIGHT
        ws1.cell(row=row_idx, column=5).alignment = RIGHT

    # Total row styling
    _style_total_row(ws1, total_row, 5)
    ws1.cell(row=total_row, column=2).alignment = LEFT
    ws1.cell(row=total_row, column=3).alignment = RIGHT
    ws1.cell(row=total_row, column=4).alignment = RIGHT
    ws1.cell(row=total_row, column=5).alignment = RIGHT

    ws1.freeze_panes = "A5"

    # =================================================================
    # Sheet 2 — Monthly Cost Observation
    # =================================================================

    ws2 = wb.create_sheet("Monthly Cost Observation")
    ws2.sheet_properties.tabColor = TAB_COLOR_OBSERVATION

    # Title
    ws2.merge_cells("A1:G1")
    title_cell = ws2["A1"]
    title_cell.value = "Monthly Cost Observation"
    title_cell.font = Font(
        name="Calibri",
        bold=True,
        size=16,
        color=HEADER_BG
    )
    title_cell.alignment = LEFT

    # Subtitle
    ws2.merge_cells("A2:G2")
    sub_cell = ws2["A2"]
    sub_cell.value = f"Generated on {report_date}"
    sub_cell.font = Font(
        name="Calibri",
        italic=True,
        size=9,
        color=META_FONT_COLOR
    )
    sub_cell.alignment = LEFT

    # Separator
    for col in range(1, 8):
        ws2.cell(row=3, column=col).border = _bottom_border()

    # Column widths
    ws2.column_dimensions["A"].width = 30   # Account Name
    ws2.column_dimensions["B"].width = 18   # Account ID
    ws2.column_dimensions["C"].width = 20   # Prev Month
    ws2.column_dimensions["D"].width = 20   # Curr Month
    ws2.column_dimensions["E"].width = 16   # Change
    ws2.column_dimensions["F"].width = 14   # Status
    ws2.column_dimensions["G"].width = 100  # Primary Reason

    obs_headers = [
        "Account Name", "Account ID",
        prev_month_name, curr_month_name,
        "Change", "Status", "Primary Reason"
    ]

    # Header row (row 4)
    _write_cells(ws2, 4, obs_headers)
    _style_header_row(ws2, 4, 7)

    # Data rows
    row = 5
    obs_data_start = row

    for data_row in observation_rows:
        _write_cells(ws2, row, data_row)
        row += 1

    obs_data_end = row - 1

    # Stripes
    _apply_stripes(ws2, obs_data_start, obs_data_end, 7)

    # Alignment & conditional formatting
    for row_idx in range(obs_data_start, obs_data_end + 1):

        ws2.cell(row=row_idx, column=1).alignment = LEFT
        ws2.cell(row=row_idx, column=2).alignment = CENTER
        ws2.cell(row=row_idx, column=3).alignment = RIGHT
        ws2.cell(row=row_idx, column=4).alignment = RIGHT
        ws2.cell(row=row_idx, column=5).alignment = RIGHT
        ws2.cell(row=row_idx, column=6).alignment = CENTER
        reason_cell = ws2.cell(row=row_idx, column=7)
        reason_cell.alignment = WRAP

        # Auto-calculate row height based on content length
        reason_text = reason_cell.value or ""
        col_g_width = 100  # matches column G width
        # Approximate chars per line (width * ~1.1 chars)
        chars_per_line = int(col_g_width * 1.1)
        if chars_per_line > 0 and len(reason_text) > 0:
            num_lines = max(
                1,
                -(-len(reason_text) // chars_per_line)  # ceiling div
            )
            # 15 points per line, minimum 20
            ws2.row_dimensions[row_idx].height = max(
                20, num_lines * 18
            )
        else:
            ws2.row_dimensions[row_idx].height = 20

        # Status cell conditional formatting
        status_cell = ws2.cell(row=row_idx, column=6)
        status_val = status_cell.value

        if status_val == "Increased":

            status_cell.fill = _fill(STATUS_INCREASED_BG)
            status_cell.font = _body_font(
                bold=True,
                color=STATUS_INCREASED_FONT
            )

        elif status_val == "Decreased":

            status_cell.fill = _fill(STATUS_DECREASED_BG)
            status_cell.font = _body_font(
                bold=True,
                color=STATUS_DECREASED_FONT
            )

        elif status_val == "Stable":

            status_cell.fill = _fill(STATUS_STABLE_BG)
            status_cell.font = _body_font(
                bold=True,
                color=STATUS_STABLE_FONT
            )

        # Change column coloring (mirrors status)
        change_cell = ws2.cell(row=row_idx, column=5)

        if status_val == "Increased":
            change_cell.font = _body_font(
                bold=True,
                color=STATUS_INCREASED_FONT
            )
        elif status_val == "Decreased":
            change_cell.font = _body_font(
                bold=True,
                color=STATUS_DECREASED_FONT
            )

    ws2.freeze_panes = "A5"

    # =================================================================
    # Sheet 3 — Last 5 Months Billing History
    # =================================================================

    ws3 = wb.create_sheet("Last 5 Months Billing History")
    ws3.sheet_properties.tabColor = TAB_COLOR_HISTORY

    # Sheet heading
    ws3.merge_cells("A1:D1")
    sheet_title = ws3["A1"]
    sheet_title.value = "Last 5 Months Billing History"
    sheet_title.font = Font(
        name="Calibri",
        bold=True,
        size=16,
        color=HEADER_BG
    )
    sheet_title.alignment = LEFT

    # Subtitle
    ws3.merge_cells("A2:D2")
    sub_cell = ws3["A2"]
    sub_cell.value = f"Generated on {report_date}"
    sub_cell.font = Font(
        name="Calibri",
        italic=True,
        size=9,
        color=META_FONT_COLOR
    )
    sub_cell.alignment = LEFT

    # Separator
    for col in range(1, 5):
        ws3.cell(row=3, column=col).border = _bottom_border()

    ws3.column_dimensions["A"].width = 16
    ws3.column_dimensions["B"].width = 22
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 22

    history_headers = [
        "Month", "Cost (Excl. Tax)",
        "Tax", "Total (Incl. Tax)"
    ]

    current_row = 5

    for result in results:

        account_name = result["Account Name"]

        history_with_tax = result["History (With Tax)"]
        history_without_tax = result["History (Without Tax)"]

        # ── Account Title Row ──

        ws3.merge_cells(
            start_row=current_row, start_column=1,
            end_row=current_row, end_column=4
        )

        title_cell = ws3.cell(row=current_row, column=1)
        title_cell.value = account_name
        title_cell.font = Font(
            name="Calibri",
            bold=True,
            size=13,
            color=ACCOUNT_TITLE_FONT
        )
        title_cell.fill = _fill(ACCOUNT_TITLE_BG)
        title_cell.alignment = LEFT

        for c in range(2, 5):
            ws3.cell(
                row=current_row, column=c
            ).fill = _fill(ACCOUNT_TITLE_BG)

        current_row += 1

        # ── Header Row ──

        _write_cells(ws3, current_row, history_headers)
        _style_header_row(ws3, current_row, 4)
        _style_tax_header(ws3, current_row, 3)  # Tax purple

        current_row += 1

        # ── Data Rows ──

        h_start = current_row

        for j in range(len(history_without_tax)):

            cost_excl = round(
                history_without_tax[j]["Cost"], 2
            )
            cost_incl = round(
                history_with_tax[j]["Cost"], 2
            )
            tax = round(cost_incl - cost_excl, 2)

            _write_cells(ws3, current_row, [
                history_without_tax[j]["Month"],
                f"${cost_excl:,.2f}",
                f"${tax:,.2f}",
                f"${cost_incl:,.2f}"
            ])

            current_row += 1

        h_end = current_row - 1

        _apply_stripes(ws3, h_start, h_end, 4)

        for row_idx in range(h_start, h_end + 1):
            ws3.cell(row=row_idx, column=1).alignment = CENTER
            ws3.cell(row=row_idx, column=2).alignment = RIGHT
            ws3.cell(row=row_idx, column=3).alignment = RIGHT
            ws3.cell(row=row_idx, column=4).alignment = RIGHT

        # Gap before next account
        current_row += 3

    # Remove default freeze on sheet 3 (multiple tables)
    ws3.freeze_panes = None

    # =================================================================
    # Sheet 4 — Service Level Cost Breakdown
    # =================================================================

    ws4 = wb.create_sheet("Service Level Cost")
    ws4.sheet_properties.tabColor = "8E44AD"

    # Sheet heading
    ws4.merge_cells("A1:D1")
    sheet_title = ws4["A1"]
    sheet_title.value = (
        f"Service Level Cost Breakdown — {curr_month_name}"
    )
    sheet_title.font = Font(
        name="Calibri",
        bold=True,
        size=16,
        color=HEADER_BG
    )
    sheet_title.alignment = LEFT

    # Subtitle
    ws4.merge_cells("A2:D2")
    sub_cell = ws4["A2"]
    sub_cell.value = f"Generated on {report_date}"
    sub_cell.font = Font(
        name="Calibri",
        italic=True,
        size=9,
        color=META_FONT_COLOR
    )
    sub_cell.alignment = LEFT

    # Separator
    for col in range(1, 5):
        ws4.cell(row=3, column=col).border = _bottom_border()

    ws4.column_dimensions["A"].width = 40
    ws4.column_dimensions["B"].width = 20
    ws4.column_dimensions["C"].width = 14
    ws4.column_dimensions["D"].width = 5   # Spacer

    service_headers = ["Service", "Cost (USD)", "%"]
    current_row = 5

    TOP_N = 10  # Show top 10 services, group rest as "Other"

    for result in results:

        account_name = result["Account Name"]
        services = result.get("Services (Without Tax)", {})

        # Sort services by cost descending
        sorted_services = sorted(
            services.items(),
            key=lambda x: x[1],
            reverse=True
        )

        # Filter out zero/negative cost services
        sorted_services = [
            (name, cost) for name, cost
            in sorted_services if cost > 0.01
        ]

        total_service_cost = sum(
            cost for _, cost in sorted_services
        )

        # Group into top N + Other
        if len(sorted_services) > TOP_N:

            top_services = sorted_services[:TOP_N]

            other_cost = sum(
                cost for _, cost
                in sorted_services[TOP_N:]
            )

            top_services.append(("Other", other_cost))

        else:
            top_services = sorted_services

        # ── Account Title Row ──

        ws4.merge_cells(
            start_row=current_row, start_column=1,
            end_row=current_row, end_column=3
        )

        title_cell = ws4.cell(row=current_row, column=1)
        title_cell.value = account_name
        title_cell.font = Font(
            name="Calibri",
            bold=True,
            size=13,
            color=ACCOUNT_TITLE_FONT
        )
        title_cell.fill = _fill(ACCOUNT_TITLE_BG)
        title_cell.alignment = LEFT

        for c in range(2, 4):
            ws4.cell(
                row=current_row, column=c
            ).fill = _fill(ACCOUNT_TITLE_BG)

        current_row += 1

        # ── Header Row ──

        _write_cells(ws4, current_row, service_headers)
        _style_header_row(ws4, current_row, 3)

        header_row = current_row
        current_row += 1

        # ── Data Rows (numeric values for chart) ──

        data_start = current_row

        for service_name, cost in top_services:

            cost_rounded = round(cost, 2)

            pct = (
                round((cost / total_service_cost) * 100, 1)
                if total_service_cost > 0 else 0
            )

            ws4.cell(
                row=current_row, column=1,
                value=service_name
            )

            # Numeric value (for chart reference)
            cost_cell = ws4.cell(
                row=current_row, column=2,
                value=cost_rounded
            )
            cost_cell.number_format = '$#,##0.00'

            pct_cell = ws4.cell(
                row=current_row, column=3,
                value=pct / 100
            )
            pct_cell.number_format = '0.0%'

            current_row += 1

        data_end = current_row - 1

        # Stripes & alignment
        _apply_stripes(ws4, data_start, data_end, 3)

        for row_idx in range(data_start, data_end + 1):
            ws4.cell(row=row_idx, column=1).alignment = LEFT
            ws4.cell(row=row_idx, column=2).alignment = RIGHT
            ws4.cell(row=row_idx, column=3).alignment = CENTER

        # ── Total Row ──

        ws4.cell(
            row=current_row, column=1, value="Total"
        )

        total_cell = ws4.cell(
            row=current_row, column=2,
            value=round(total_service_cost, 2)
        )
        total_cell.number_format = '$#,##0.00'

        ws4.cell(
            row=current_row, column=3, value=1.0
        ).number_format = '0.0%'

        _style_total_row(ws4, current_row, 3)
        ws4.cell(row=current_row, column=1).alignment = LEFT
        ws4.cell(row=current_row, column=2).alignment = RIGHT
        ws4.cell(row=current_row, column=3).alignment = CENTER

        # ── Pie Chart ──

        if len(top_services) > 0:

            from openpyxl.chart.label import DataLabelList

            chart = PieChart()
            chart.title = account_name
            chart.style = 26

            # Scale chart to match table height
            # Table rows: title + header + data + total
            num_rows = len(top_services) + 3
            chart.width = 16
            chart.height = max(10, num_rows * 0.9)

            # Data reference — start from data rows
            # (skip header to avoid "Cost (USD)" label)
            data_ref = Reference(
                ws4,
                min_col=2,
                min_row=data_start,
                max_row=data_end
            )

            # Category reference (Service names)
            cats_ref = Reference(
                ws4,
                min_col=1,
                min_row=data_start,
                max_row=data_end
            )

            chart.add_data(data_ref, titles_from_data=False)
            chart.set_categories(cats_ref)

            # Labels: percentage only, no series name
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showPercent = True
            chart.dataLabels.showCatName = False
            chart.dataLabels.showVal = False
            chart.dataLabels.showSerName = False
            chart.dataLabels.showLeaderLines = True

            # Legend on the right — shows service names
            chart.legend.position = "r"

            # Place chart next to the table
            chart_anchor = f"E{header_row - 1}"
            ws4.add_chart(chart, chart_anchor)

        # Gap so next account doesn't overlap the chart
        chart_rows = max(20, int(num_rows * 1.8) + 4)
        min_next_row = header_row - 1 + chart_rows
        current_row = max(current_row + 3, min_next_row)

    ws4.freeze_panes = None

    # ---------- Save Workbook ----------

    wb.save(file_name)

    return file_name