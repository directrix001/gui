from asyncio import log
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from datetime import datetime, timedelta


def select_excel(prompt):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title=prompt,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
# ------------------------------------------------
# SELECT THE MAIN WORKBOOK (contains ALL sheets)
# ------------------------------------------------

file1 = select_excel("Select the Excel file that contains Calc Template, Excelpayslip, and NI")
print("Workbook selected:", file1)

# Load workbook
wb = load_workbook(file1, data_only=True)

# ------------------------------------------------
# FIXED SHEET ASSIGNMENTS (NO USER INPUT)
# ------------------------------------------------

ws  = wb["Calc Template"]   # destination sheet
ws2 = wb["Excelpayslip"]    # source sheet 1
ws3 = wb["NI"]     # source sheet 2
ws4 = wb["TA-BH"]         # source sheet 3

print(wb.sheetnames)
# ------------------------------------------------
# COPY VALUES FROM ws2 (Excelpayslip)
# ------------------------------------------------

for row in range(1, ws2.max_row + 1):
    if ws2.cell(row=row, column=1).value == "BASIC":
        ws["C5"] = ws2.cell(row=row, column=3).value
        break

for row in range(1, ws2.max_row + 1):

    code = str(ws2.cell(row=row, column=1).value).strip()

    if code == "BASIC":
        ws["C5"] = ws2.cell(row=row, column=3).value

    elif code == "UKSIPREF":
        ws["C38"] = ws2.cell(row=row, column=3).value

    elif code == "OT @ 1.5":
        ws["C16"] = ws2.cell(row=row, column=3).value

    elif code.startswith("SHIFT"):
        ws["C6"] = ws2.cell(row=row, column=3).value

    elif code == "RGOLD":

        value = str(ws2.cell(row=row, column=3).value).strip()

        if value.endswith("C"):
            value = value[:-1]

        ws["C7"] = -abs(float(value))

for row in range(1, ws2.max_row + 1):

    code = str(ws2.cell(row=row, column=4).value or "").strip()

    value = str(ws2.cell(row=row, column=6).value or "").strip()

    if value.endswith("C"):
        value = value[:-1]

    try:
        value = float(value)
    except:
        continue

    if code == "Income Tax":
        ws["C44"] = value

    elif code == "Nat.Ins.":
        ws["C45"] = value

    elif code == "STD LOAN":
        ws["C49"] = value

    elif code == "RECSF":
        ws["C62"] = -abs(value)

    elif code == "SPTS&SOC":
        ws["C64"] = value

    elif code == "FIT CENT":
        ws["C65"] = value

    elif code == "NUFCDRAW":
        ws["C68"] = value

    elif code == "HOL DRAW":
        ws["C69"] = value

    elif code == "CAR DRAW":
        ws["C70"] = value

    elif code == "UNION":
        ws["C71"] = value

# Copy C64 value to B64
ws = wb["Calc Template"]

ws["B64"] = ws["C64"].value
# ------------------------------------------------
# COPY VALUES FROM ws3 (NI)
# ------------------------------------------------

ws["C85"] = float(str(ws3["B4"].value).rstrip("C").strip() or 0)
ws["C86"] = float(str(ws3["B5"].value).rstrip("C").strip() or 0)
ws["C80"] = float(str(ws3["B6"].value).rstrip("C").strip() or 0)
ws["C81"] = float(str(ws3["B7"].value).rstrip("C").strip() or 0)
#ws["C82"] = float(str(ws3["B8"].value).rstrip("C").strip() or 0)
# ==========================================================
# D12 = B12 - C12
# ==========================================================

ws["D7"] = round(
    float(ws["B7"].value or 0)
    -
    float(ws["C7"].value or 0),
    2
)
# ==========================================================
# OT EARLY -> C12
# Column A = Code
# Column C = Value
# ==========================================================

ot_early_value = 0

for row_num in range(2, ws2.max_row + 1):

    code = str(ws2[f"A{row_num}"].value or "").strip().upper()

    if "OT EARLY" in code:

        try:
            ot_early_value += float(ws2[f"C{row_num}"].value or 0)
        except:
            pass

ws["C12"] = round(ot_early_value, 2)

print("OT EARLY =", ot_early_value)




# ==========================================================
# OT NIGHT -> C14
# Column A = Code
# Column C = Value
# ==========================================================

ot_night_value = 0

for row_num in range(2, ws2.max_row + 1):

    code = str(
        ws2[f"A{row_num}"].value or ""
    ).strip().upper()

    if code == "OT NIGHT":

        ot_night_value += float(
            ws2[f"C{row_num}"].value or 0
        )

ws["C14"] = round(ot_night_value, 2)
# ==========================================================
# HOURLY RATE
# ==========================================================

hourly_rate = 0

for row_num in range(1, ws2.max_row):

    text_value = str(
        ws2[f"A{row_num}"].value or ""
    ).strip()

    if "CNTRCTHR" in text_value.upper():

        next_value = str(
            ws2[f"A{row_num + 1}"].value or ""
        ).strip()

        next_value = next_value.replace("(", "")
        next_value = next_value.replace(")", "")

        if "X" in next_value:

            parts = next_value.split("X")

            if len(parts) == 2:

                hourly_rate = float(
                    parts[1].strip()
                )

        break


# ==========================================================
# OT EARLY HOURS
# ==========================================================

ot_early_hours = 0

time_entry_col = 0
calc_tag_col = 0

# Find columns from row 2
for col_num in range(1, ws4.max_column + 1):

    header = str(
        ws4.cell(row=2, column=col_num).value or ""
    ).strip().upper()

    if header == "TIME ENTRY CODE":
        time_entry_col = col_num

    elif header == "CALCULATION TAGS":
        calc_tag_col = col_num

# Sum matching Reported Quantity values from Column E
for row_num in range(3, ws4.max_row + 1):

    time_entry = str(
        ws4.cell(
            row=row_num,
            column=time_entry_col
        ).value or ""
    ).strip().upper()

    calc_tag = str(
        ws4.cell(
            row=row_num,
            column=calc_tag_col
        ).value or ""
    ).strip().upper()

    qty = float(
        ws4[f"E{row_num}"].value or 0
    )

    if (
    time_entry in [
        "UK - OVERTIME X1.5",
        "UK - OVERTIME X2.0",
        "UK - SHORT NOTICE OVERTIME (SNOT) X1.5"
    ]
    and
    "EARLY SHIFT" in calc_tag
):


        ot_early_hours += qty

    print(
        "Matched Row:",
        row_num,
        "|",
        time_entry,
        "|",
        calc_tag,
        "| Qty =",
        qty
    )

# ==========================================================
# B12 = (Hourly Rate * OT Early Hours * 13.33%) + C12
# ==========================================================

premium_value = (
    hourly_rate *
    ot_early_hours *
    0.1333
)

ws["B12"] = round(
    premium_value +
    float(ws["C12"].value or 0),
    2
)

# ==========================================================
# D12 = B12 - C12
# ==========================================================

ws["D12"] = round(
    float(ws["B12"].value or 0)
    -
    float(ws["C12"].value or 0),
    2
)
# ==========================================================
# OT LATE -> C13
# Column A = Code
# Column C = Value
# ==========================================================

ot_late_value = 0

for row_num in range(2, ws2.max_row + 1):

    code = str(ws2[f"A{row_num}"].value or "").strip().upper()

    if "OT LATE" in code:
        try:
            ot_late_value += float(ws2[f"C{row_num}"].value or 0)
        except:
            pass

ws["C13"] = round(ot_late_value, 2)

print("OT LATE Total =", ws["C13"].value)
# ==========================================================
# OT LATE HOURS
# ==========================================================

ot_late_hours = 0

for row_num in range(3, ws4.max_row + 1):

    time_entry = str(
        ws4.cell(row=row_num, column=time_entry_col).value or ""
    ).strip()

    calc_tag = str(
        ws4.cell(row=row_num, column=calc_tag_col).value or ""
    ).replace("\n", "").replace("\r", "").strip()

    if (
        time_entry in [
            "UK - Short Notice Overtime (SNOT) x1.5",
            "UK - Overtime x1.5",
            "UK - Overtime x2.0"
        ]
        and
        "LATE SHIFT" in calc_tag.upper()
        and
        "OVERTIME" in calc_tag.upper()
    ):

        qty = float(ws4.cell(row=row_num, column=5).value or 0)

        ot_late_hours += qty

        print(
            f"MATCH FOUND - Row {row_num} | "
            f"Time Entry: {time_entry} | "
            f"Calc Tag: {calc_tag} | "
            f"Qty: {qty}"
        )

print("OT Late Hours =", ot_late_hours)



# ==========================================================
# B13 = (Hourly Rate * OT Late Hours * 20%) + C13
# ==========================================================

premium_value = (
    hourly_rate *
    ot_late_hours *
    0.20
)

ws["B13"] = round(
    premium_value +
    float(ws["C13"].value or 0),
    2
)

# ==========================================================
# D13 = B13 - C13
# ==========================================================

ws["D13"] = round(
    float(ws["B13"].value or 0)
    -
    float(ws["C13"].value or 0),
    2
)

# ==========================================================
# OT NIGHT HOURS
# ==========================================================

ot_night_hours = 0

time_entry_col = 0
calc_tag_col = 0

# Find columns from Row 2
for col_num in range(1, ws4.max_column + 1):

    header = str(
        ws4.cell(row=2, column=col_num).value or ""
    ).strip().upper()

    if header == "TIME ENTRY CODE":
        time_entry_col = col_num

    elif header == "CALCULATION TAGS":
        calc_tag_col = col_num

# Sum Reported Quantity (Column E)
for row_num in range(3, ws4.max_row + 1):

    time_entry = str(
        ws4.cell(
            row=row_num,
            column=time_entry_col
        ).value or ""
    ).strip().upper()

    calc_tag = str(
        ws4.cell(
            row=row_num,
            column=calc_tag_col
        ).value or ""
    ).strip().upper()

    if (
        (
            time_entry == "UK - OVERTIME X1.5"
            and
            "NIGHT SHIFT" in calc_tag
            and
            "1.5X" in calc_tag
        )
        or
        (
            time_entry == "UK - OVERTIME X2.0"
            and
            "NIGHT SHIFT" in calc_tag
            and
            "2.0X" in calc_tag
        )
    ):

        ot_night_hours += float(
            ws4[f"E{row_num}"].value or 0
        )

# ==========================================================
# B14 = (Hourly Rate * OT Night Hours * 33.33%) + C14
# ==========================================================

premium_value = (
    hourly_rate *
    ot_night_hours *
    0.3333
)

ws["B14"] = round(
    premium_value +
    float(ws["C14"].value or 0),
    2
)

# ==========================================================
# D14 = B14 - C14
# ==========================================================

ws["D14"] = round(
    float(ws["B14"].value or 0)
    -
    float(ws["C14"].value or 0),
    2
)


# ==========================================================
# OCCASIONAL SHIFT CHANGE DAY TO EARLY
# ==========================================================

occasional_early_hours = 0

time_entry_col = 0
calc_tag_col = 0

# Find columns from Row 2
for col_num in range(1, ws4.max_column + 1):

    header = str(
        ws4.cell(
            row=2,
            column=col_num
        ).value or ""
    ).strip().upper()

    if header == "TIME ENTRY CODE":
        time_entry_col = col_num

    elif header == "CALCULATION TAGS":
        calc_tag_col = col_num

# Sum Reported Quantity (Column E)
for row_num in range(3, ws4.max_row + 1):

    time_entry = ws4.cell(
        row=row_num,
        column=time_entry_col
    ).value

    calc_tag = str(
        ws4.cell(
            row=row_num,
            column=calc_tag_col
        ).value or ""
    ).strip().upper()

    if (
        (time_entry is None or str(time_entry).strip() == "")
        and
        "OCCASIONAL SHIFT CHANGE DAY TO EARLY" in calc_tag
    ):

        occasional_early_hours += float(
            ws4[f"E{row_num}"].value or 0
        )
# --------------------------------------------
# SEARCH COLUMN A FOR "D TO E"
# GET VALUE FROM COLUMN C
# PASTE TO Calc Template C18
# --------------------------------------------

d_to_e_value = 0

for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row):

    col_a = str(row[0].value or "").upper().strip()

    if "D TO E" in col_a:

        try:
            d_to_e_value = float(ws2[f"C{row[0].row}"].value or 0)
        except:
            d_to_e_value = 0

        print(f"Found D TO E in row {row[0].row}")
        print(f"Column C Value = {d_to_e_value}")

        break

ws["C18"] = round(d_to_e_value, 2)

print("Calc Template C18 =", ws["C18"].value)
# ==========================================================
# B18 = (Hourly Rate * Hours * 13.33%) + C18
# ==========================================================

premium_value = (
    hourly_rate *
    occasional_early_hours *
    0.1333
)

ws["B18"] = round(
    premium_value +
    float(ws["C18"].value or 0),
    2
)

# ==========================================================
# D18 = B18 - C18
# ==========================================================

ws["D18"] = round(
    float(ws["B18"].value or 0)
    -
    float(ws["C18"].value or 0),
    2
)


# ==========================================================
# OCCASIONAL SHIFT CHANGE DAY TO NIGHT
#
# Time Entry Code = Blank
#
# Include:
# Day ShiftOccasional Shift Change Day to Night
# Day Shift Occasional Shift Change Day to Night
#
# Exclude:
# A Non-Scheduled Day (UK) Day Shift Occasional Shift Change Day to Night
#
# Reported Quantity = Column E
# ==========================================================

occasional_night_hours = 0

time_entry_col = 0
calc_tag_col = 0

# Find columns from Row 2
for col_num in range(1, ws4.max_column + 1):

    header = str(
        ws4.cell(
            row=2,
            column=col_num
        ).value or ""
    ).strip().upper()

    if header == "TIME ENTRY CODE":
        time_entry_col = col_num

    elif header == "CALCULATION TAGS":
        calc_tag_col = col_num

# Sum matching Reported Quantity values
for row_num in range(3, ws4.max_row + 1):

    time_entry = ws4.cell(
        row=row_num,
        column=time_entry_col
    ).value

    calc_tag = str(
        ws4.cell(
            row=row_num,
            column=calc_tag_col
        ).value or ""
    ).strip().upper()

    if (
        (time_entry is None or str(time_entry).strip() == "")
        and
        "OCCASIONAL SHIFT CHANGE DAY TO NIGHT" in calc_tag
        and
        "NON-SCHEDULED DAY" not in calc_tag
    ):

        occasional_night_hours += float(
            ws4[f"E{row_num}"].value or 0
        )

# ==========================================================
# B19 = (Hourly Rate * Hours * 33.33%) + C19
# ==========================================================

premium_value = (
    hourly_rate *
    occasional_night_hours *
    0.3333
)

ws["B19"] = round(
    premium_value +
    float(ws["C19"].value or 0),
    2
)

# ==========================================================
# D19 = B19 - C19
# ==========================================================

ws["D19"] = round(
    float(ws["B19"].value or 0)
    -
    float(ws["C19"].value or 0),
    2
)

# ==========================================================
# OCCASIONAL SHIFT CHANGE DAY TO LATE
#
# Time Entry Code = Blank
#
# Include:
# Day ShiftOccasional Shift Change Day to Late
# Day Shift Occasional Shift Change Day to Late
#
# Exclude:
# A Non-Scheduled Day (UK) Day Shift Occasional Shift Change Day to Late
#
# Reported Quantity = Column E
# ==========================================================

occasional_late_hours = 0

time_entry_col = 0
calc_tag_col = 0

# Find columns from Row 2
for col_num in range(1, ws4.max_column + 1):

    header = str(
        ws4.cell(
            row=2,
            column=col_num
        ).value or ""
    ).strip().upper()

    if header == "TIME ENTRY CODE":
        time_entry_col = col_num

    elif header == "CALCULATION TAGS":
        calc_tag_col = col_num

# Sum matching Reported Quantity values
for row_num in range(3, ws4.max_row + 1):

    time_entry = ws4.cell(
        row=row_num,
        column=time_entry_col
    ).value

    calc_tag = str(
        ws4.cell(
            row=row_num,
            column=calc_tag_col
        ).value or ""
    ).strip().upper()

    if (
        (time_entry is None or str(time_entry).strip() == "")
        and
        "OCCASIONAL SHIFT CHANGE DAY TO LATE" in calc_tag
        and
        "NON-SCHEDULED DAY" not in calc_tag
    ):

        occasional_late_hours += float(
            ws4[f"E{row_num}"].value or 0
        )

# ==========================================================
# B20 = (Hourly Rate * Hours * 20%) + C20
# ==========================================================

premium_value = (
    hourly_rate *
    occasional_late_hours *
    0.20
)

ws["B20"] = round(
    premium_value +
    float(ws["C20"].value or 0),
    2
)

# ==========================================================
# D20 = B20 - C20
# ==========================================================

ws["D20"] = round(
    float(ws["B20"].value or 0)
    -
    float(ws["C20"].value or 0),
    2
)


# Default value for C22
c22_value = 0

# Scan ws2 column A for "LUMP SUM"
for row in ws2.iter_rows(min_col=1, max_col=1):
    cell_value = str(row[0].value).upper().strip() if row[0].value else ""
    
    if "LUMP SUM" in cell_value:
        # Value is in column C of the same row
        c22_value = float(ws2[f"C{row[0].row}"].value or 0)
        break

# Write result to ws C22
ws["C22"] = c22_value


def count_weekdays(start_date, end_date):
    day_count = 0
    current = start_date

    while current <= end_date:
        if current.weekday() < 5:  # Monday=0 ... Friday=4
            day_count += 1
        current += timedelta(days=1)

    return day_count

# --- Read date from B2 ---
excel_date = ws["B2"].value

# Convert string to date if needed
if isinstance(excel_date, str):
    excel_date = datetime.strptime(excel_date, "%d/%m/%Y")

# First day of the same month
first_day = excel_date.replace(day=1).date()

# Target date
target_day = excel_date.date()

# Count weekdays within the month
weekday_count = count_weekdays(first_day, target_day)

# Write result into B5
ws["B5"] = weekday_count


import re

# ------------------------------------------------
# EXTRACT b_value FROM && @ CNTRCTHR
# ------------------------------------------------

b_value = 0.0

for row in range(1, ws2.max_row + 1):

    cell_text = str(ws2.cell(row=row, column=1).value or "")

    if "&& @ CNTRCTHR" in cell_text.upper():

        next_row_text = str(
            ws2.cell(row=row + 1, column=1).value or ""
        )

        numbers = re.findall(
            r"[-+]?\d*\.\d+|\d+",
            next_row_text
        )

        print("CNTRCTHR Row =", cell_text)
        print("Next Row =", next_row_text)
        print("Numbers =", numbers)

        if len(numbers) >= 2:
            b_value = float(numbers[1])   # 23.6057
        elif len(numbers) == 1:
            b_value = float(numbers[0])

        break

print("b_value =", b_value)

# ------------------------------------------------
# 2) Read weekday_count
# ------------------------------------------------
weekday_count = float(ws["B5"].value or 0)

import re


# ------------------------------------------------
# FIND PAY GROUP IN COLUMN A
# ------------------------------------------------

pay_group = None

for row in range(1, ws2.max_row + 1):

    text = str(ws2.cell(row=row, column=1).value or "").upper().strip()

    if "PAY GROUP" in text:

        for code in ["NMUK", "NTCE", "NITCO", "NMPC", "NMGB", "NDE"]:
            if code in text:
                pay_group = code
                break

        if pay_group:
            break

# ------------------------------------------------
# DETERMINE RATE
# ------------------------------------------------

rate = None

if pay_group in ["NMUK", "NTCE", "NITCO", "NMPC"]:
    rate = 7.8
elif pay_group in ["NMGB", "NDE"]:
    rate = 7.5

# ------------------------------------------------
# CALCULATE B5
# ------------------------------------------------

if rate is None:
    print("WARNING: No valid pay group found.")
    ws["B5"] = 0
else:
    ws["B5"] = b_value * weekday_count * rate

# ------------------------------------------------
# CALCULATE D5 = B5 - C5
# ------------------------------------------------

b5_value = float(ws["B5"].value or 0)
c5_value = float(ws["C5"].value or 0)

ws["D5"] = b5_value - c5_value

import re

shift_text = ""

# --- Scan ws2 column A for SHIFT or SH ---
for row in ws2.iter_rows(min_col=1, max_col=1):
    cell_value = str(row[0].value).upper().strip()
    if cell_value.startswith("SHIFT") or cell_value.startswith("SH "):
        shift_text = cell_value
        break

# Clean text
shift_text = re.sub(r"[^A-Z0-9 /]", " ", shift_text)
shift_text = " ".join(shift_text.split())

# Split into words
parts = shift_text.split()

# No SHIFT found
if len(parts) < 2:
    print(f"WARNING: No shift code found in '{shift_text}'")
    ws["D6"] = 0
else:
    # Extract the shift code after SHIFT/SH
    shift_part = " ".join(parts[1:])  # everything after SHIFT

    # Handle MN/SP etc.
    if "/" in shift_part:
        shift_candidates = shift_part.split("/")
    else:
        shift_candidates = [shift_part]

    # SHIFT RATE TABLE
    shift_rates = {
        "DL": 0.12,
        "DF": 0.22,
        "DT": 0.22,
        "MF": 0.38,
        "RT": 0.38,
        "WE A": 0.30,
        "WE B": 0.10,
        "WE": 0.225,
        "ES": 0.1333,
        "MN": 0.185,
        "SP": 0.185,
        "WEC": 0.20,
        "TS": 0.27
    }

    rate_percent = None

    # Try two‑word codes first
    for c in shift_candidates:
        if c in ["WE A", "WE B"]:
            rate_percent = shift_rates[c]
            break

    # Try one‑word codes
    if rate_percent is None:
        for c in shift_candidates:
            # Only take the first token (DL, DF, DT, etc.)
            token = c.split()[0]
            if token in shift_rates:
                rate_percent = shift_rates[token]
                break

    # Apply or warn
    if rate_percent is None:
        print(f"WARNING: Unknown shift code in ws2 column A → '{shift_text}'")
        ws["D6"] = 0
    else:
        b5_value = float(ws["B5"].value or 0)
        ws["D6"] = b5_value * rate_percent


# --- Calculate B6 = C6 + D6 ---
c6_value = float(ws["C6"].value or 0)
d6_value = float(ws["D6"].value or 0)

ws["B6"] = c6_value + d6_value

import re

# --- Scan ws2 column A for rating code ---
rating_text = ""

for row in ws2.iter_rows(min_col=1, max_col=1):
    cell_value = str(row[0].value).upper().strip()

    if "RS AE 3" in cell_value:
        rating_text = "RS AE 3"
        break
    elif "RGOLD" in cell_value:
        rating_text = "RGOLD"
        break
    elif "RPLATINUM" in cell_value:
        rating_text = "RPLATINUM"
        break
    elif "RSILVER" in cell_value:
        rating_text = "RSILVER"
        break
    elif "RBRONZE" in cell_value:
        rating_text = "RBRONZE"
        break

# Rating percentage table
rating_rates = {
    "RS AE 3": 0.03,
    "RGOLD": 0.05,
    "RPLATINUM": 0.07,
    "RSILVER": 0.04,
    "RBRONZE": 0.03
}

rate_percent = rating_rates.get(rating_text)

# Read values
b5_value = float(ws["B5"].value or 0)
b6_value = float(ws["B6"].value or 0)
c22_value = float(ws["C22"].value or 0)

# Apply formula
if rate_percent is None:
    print(f"WARNING: Unknown rating code '{rating_text}' — B7 set to 0")
    ws["B7"] = 0
else:
    ws["B7"] = - (b5_value + b6_value + c22_value) * rate_percent

# ------------------------------------------------
# Read RS AE 3 / RGOLD / RPLATINUM / RSILVER / RBRONZE
# from Excelpayslip Column A and copy Column C value
# to Calc Template C7
# ------------------------------------------------

ws["C7"] = 0

target_codes = {
    "RS AE 3",
    "RGOLD",
    "RPLATINUM",
    "RSILVER",
    "RBRONZE"
}

for row in range(1, ws2.max_row + 1):

    code = str(ws2[f"A{row}"].value or "").strip().upper()

    if code in target_codes:

        value = ws2[f"C{row}"].value

        try:
            raw_value = str(value).strip().upper()

            if raw_value.endswith("C"):
                amount = float(raw_value[:-1].replace(",", ""))
                ws["C7"] = -abs(amount)
            else:
                ws["C7"] = float(str(value).replace(",", ""))

        except:
            ws["C7"] = value or 0

        break

print("C7 =", ws["C7"].value)
# --- Calculate base amount ---
b5_value = float(ws["B5"].value or 0)
b6_value = float(ws["B6"].value or 0)
c22_value = float(ws["C22"].value or 0)

base_amount = b5_value + b6_value + c22_value
# --- B88 percentage table Employer---
b88_rates = {
    "RGOLD": 0.10,
    "RPLATINUM": 0.12,
    "RSILVER": 0.08,
    "RBRONZE": 0.06,
    "RS AE 3": 0.06,
    "RS AE 2": 0.04
}

# Get the percentage for the detected rating
rate_b88 = b88_rates.get(rating_text, 0)

# Write B88
ws["B88"] = round(float(base_amount) * float(rate_b88), 2)

#employee B87
base_amount = b5_value + b6_value + c22_value
b87_rates = {
    "RGOLD": 0.05,
    "RPLATINUM": 0.07,
    "RSILVER": 0.04,
    "RBRONZE": 0.03,
    "RS AE 3": 0.03,
    "RS AE 2": 0.02
}

rate_b87 = b87_rates.get(rating_text, 0)

ws["B87"] = base_amount * rate_b87

# --------------------------------------------
# C87 = C5 × Rate
# --------------------------------------------

rates = {
    "RGOLD": 0.05,
    "RPLATINUM": 0.07,
    "RSILVER": 0.04,
    "RBRONZE": 0.03,
    "RS AE 3": 0.03,
    "RS AE 2": 0.02
}

c5_value = float(ws["C5"].value or 0)

rate_c87 = rates.get(rating_text, 0)

ws["C87"] = round(c5_value * rate_c87, 2)


# --------------------------------------------
# C88 = C5 × Rate
# --------------------------------------------

rates = {
    "RGOLD": 0.10,
    "RPLATINUM": 0.12,
    "RSILVER": 0.08,
    "RBRONZE": 0.06,
    "RS AE 3": 0.06,
    "RS AE 2": 0.04
}

c5_value = float(ws["C5"].value or 0)

rate_c88 = rates.get(rating_text, 0)

ws["C88"] = round(c5_value * rate_c88, 2)

# --------------------------------------------
# D87 = B87 - C87
# --------------------------------------------
b87_value = float(ws["B87"].value or 0)
c87_value = float(ws["C87"].value or 0)

ws["D87"] = round(b87_value - c87_value, 2)

print("D87 =", ws["D87"].value)

# --------------------------------------------
# D88 = B88 - C88
# --------------------------------------------
b88_value = float(ws["B88"].value or 0)
c88_value = float(ws["C88"].value or 0)

ws["D88"] = round(b88_value - c88_value, 2)

print("D88 =", ws["D88"].value)
# --- Calculate negative B8 = - (B5 + B6 + B22) * 5% ---

b5_value = float(ws["B5"].value or 0)
b6_value = float(ws["B6"].value or 0)
b22_value = float(ws["B22"].value or 0)

base_amount = b5_value + b6_value + b22_value

#ws["B8"] = - base_amount * 0.05
# --- Calculate B89 = (B5 + B6 + B22) * 5% ---

b5_value = float(ws["B5"].value or 0)
b6_value = float(ws["B6"].value or 0)
b22_value = float(ws["B22"].value or 0)

base_amount = b5_value + b6_value + b22_value

# Default value
b9_value = 0

# Scan rows
for row in ws2.iter_rows():
    a = str(row[0].value).upper().strip() if row[0].value else ""
    d = str(row[3].value).upper().strip() if row[3].value else ""
    g = str(row[6].value).upper().strip() if row[6].value else ""

    # Check column A
    if "RSMAVCFIX" in a:
        b9_value = float(row[2].value or 0)   # column C
        break

    # Check column D
    if "RSMAVCFIX" in d:
        b9_value = float(row[5].value or 0)   # column F
        break

    # Check column G
    if "RSMAVCFIX" in g:
        b9_value = float(row[8].value or 0)   # column I
        break

# Write result
ws["B9"] = b9_value


# Determine the last used column in row 2
max_col = ws4.max_column
last_col_letter = ws4.cell(row=2, column=max_col).column_letter

# Apply filter to row 2 across all columns
ws4.auto_filter.ref = f"A2:{last_col_letter}2"

ws4 = wb["TA-BH"]

# --- Find the "Time Entry Code" column ---
time_entry_col = None
for cell in ws4[2]:
    if str(cell.value).strip() == "Time Entry Code":
        time_entry_col = cell.col_idx
        break

if time_entry_col is None:
    raise ValueError("Column 'Time Entry Code' not found.")

# --- Loop through all rows and HIDE non-matching rows ---
for row in range(3, ws4.max_row + 1):
    cell_value = ws4.cell(row=row, column=time_entry_col).value
    text = str(cell_value).strip() if cell_value else ""

    if text == "UK - Overtime x1.5":
        ws4.row_dimensions[row].hidden = False   # show row
    else:
        ws4.row_dimensions[row].hidden = True    # hide row

ws4 = wb["TA-BH"]
ws  = wb["Calc Template"]  # ← replace with your actual main sheet name

# --- Find the "Reported Quantity" column in ws4 ---
reported_qty_col = None
for cell in ws4[2]:
    if str(cell.value).strip() == "Reported Quantity":
        reported_qty_col = cell.col_idx
        break

if reported_qty_col is None:
    raise ValueError("Column 'Reported Quantity' not found.")

# --- Sum values from visible rows only in ws4 ---
total_qty = 0
for row in range(3, ws4.max_row + 1):
    if ws4.row_dimensions[row].hidden:   # skip hidden rows
        continue

    val = ws4.cell(row=row, column=reported_qty_col).value
    if isinstance(val, (int, float)):
        total_qty += val

# --- Write the sum to ws (NOT ws4) ---
ws["B16"] = total_qty

# --- Extract b_value from ws2["A11"] safely ---
raw_value = str(ws2["A11"].value)
numbers = re.findall(r"[-+]?\d*\.\d+|\d+", raw_value)

if len(numbers) >= 2:
    b_value = float(numbers[1])
elif len(numbers) == 1:
    b_value = float(numbers[0])


# --- Read weekday count from ws["B5"] ---
weekday_count = float(ws["B5"].value or 0)

# --- Read existing value in C16 ---
c16_value = float(ws["C16"].value or 0)

# --- Final B16 calculation ---
ws["B16"] = total_qty * b_value * 1.5 + c16_value


# Default value for C11
c11_value = 0

# Scan ws2 column A for the word "LONDON"
for row in ws2.iter_rows(min_col=1, max_col=1):
    cell_value = str(row[0].value).upper().strip() if row[0].value else ""

    if "LONDON" in cell_value:
        # Value is in column C of the same row
        c11_value = float(ws2[f"C{row[0].row}"].value or 0)
        break

# Write result to ws C11
ws["C11"] = c11_value

import calendar
from datetime import date

# --- ALWAYS read B2 first ---
date_value = ws["B2"].value

if date_value is None:
    raise ValueError("B2 does not contain a valid date")

year = date_value.year
month = date_value.month
day_limit = date_value.day   # up to the employee's last working day

# --- Total days in the month ---
days_in_month = calendar.monthrange(year, month)[1]

# --- Weekdays from day 1 → B2 date (your 17‑day logic) ---
weekday_count = 0
for day in range(1, day_limit + 1):
    if date(year, month, day).weekday() < 5:   # Mon–Fri
        weekday_count += 1

# --- Read C11 ---
c11_value = float(ws["C11"].value or 0)

# --- Final B11 calculation ---
b11_value = (c11_value / days_in_month) * weekday_count
ws["B11"] = b11_value

import calendar
from datetime import date

# --- Read date from B2 ---
date_value = ws["B2"].value
year = date_value.year
month = date_value.month
day_limit = date_value.day   # count weekdays up to this day

# --- Total days in the month ---
days_in_month = calendar.monthrange(year, month)[1]

# --- Weekdays from 1st → B2 date ---
weekday_count = 0
for day in range(1, day_limit + 1):
    if date(year, month, day).weekday() < 5:  # Monday–Friday
        weekday_count += 1

# --- Read C22 (Senior Lumpsum value) ---
c22_value = float(ws["C22"].value or 0)

# --- Final B22 calculation ---
b11_value = (c22_value / days_in_month) * weekday_count
ws["B22"] = b11_value


# ------------------------------------------------
# OT @ 2.0 (Excelpayslip -> Calc Template!C17)
# ------------------------------------------------

excel_ws = wb["Excelpayslip"]
calc_ws = wb["Calc Template"]

ot_value = 0

for row in excel_ws.iter_rows(min_row=1, max_row=excel_ws.max_row):
    code = str(row[0].value or "").strip().upper()  # Column A

    if code == "OT @ 2.0":
        try:
            ot_value = float(row[2].value or 0)  # Column C
        except (ValueError, TypeError):
            ot_value = 0
        break

calc_ws["C17"] = ot_value

print(f'OT @ 2.0 value copied to C17: {ot_value}')
ws4 = wb["TA-BH"]
ws  = wb["Calc Template"]

# --- Find Time Entry Code column ---
time_entry_col = None
for cell in ws4[2]:
    if str(cell.value).strip() == "Time Entry Code":
        time_entry_col = cell.col_idx
        break

reported_qty_col = 5   # fixed column

# --- FILTER rows for UK - Overtime x2.0 ---
overtime_found = False

for row in range(3, ws4.max_row + 1):
    tec = ws4.cell(row=row, column=time_entry_col).value
    tec_text = str(tec).strip() if tec else ""

    if tec_text == "UK - Overtime x2.0":
        ws4.row_dimensions[row].hidden = False
        overtime_found = True
    else:
        ws4.row_dimensions[row].hidden = True

# --- SUM visible Reported Quantity ---
total_qty_b17 = 0
for row in range(3, ws4.max_row + 1):
    if ws4.row_dimensions[row].hidden:
        continue

    val = ws4.cell(row=row, column=reported_qty_col).value
    if isinstance(val, (int, float)):
        total_qty_b17 += val

# --- Extract b_value from ws2["A11"] ---
raw_value = str(ws2["A11"].value)
numbers = re.findall(r"[-+]?\d*\.\d+|\d+", raw_value)

if len(numbers) >= 2:
    b_value = float(numbers[1])
elif len(numbers) == 1:
    b_value = float(numbers[0])
else:
    b_value = 0.0   # or whatever default value you want
# --- Read C17 ---
c17_value = float(ws["C17"].value or 0)


# --- Final B17 calculation ---
if overtime_found:
    b17_value = total_qty_b17 * b_value * 2.0 + c17_value
else:
    b17_value = c17_value

ws["B17"] = round(b17_value, 2)

print("Overtime Found =", overtime_found)
print("B17 =", ws["B17"].value)




ws4 = wb["TA-BH"]
ws  = wb["Calc Template"]

# --- Find Time Entry Code column ---
time_entry_col = None
for cell in ws4[2]:
    if str(cell.value).strip() == "Time Entry Code":
        time_entry_col = cell.col_idx
        break

reported_qty_col = 5   # fixed column

# --- FILTER rows for UK - Overtime ---
for row in range(3, ws4.max_row + 1):
    tec = ws4.cell(row=row, column=time_entry_col).value
    tec_text = str(tec).strip() if tec else ""

    if tec_text == "UK - Overtime":
        ws4.row_dimensions[row].hidden = False
    else:
        ws4.row_dimensions[row].hidden = True

# --- SUM visible Reported Quantity ---
total_qty_b15 = 0
for row in range(3, ws4.max_row + 1):
    if ws4.row_dimensions[row].hidden:
        continue

    val = ws4.cell(row=row, column=reported_qty_col).value
    if isinstance(val, (int, float)):
        total_qty_b15 += val

# --- Extract b_value from ws2["A11"] ---
raw_value = str(ws2["A11"].value)
numbers = re.findall(r"[-+]?\d*\.\d+|\d+", raw_value)

if len(numbers) >= 2:
    b_value = float(numbers[1])
elif len(numbers) == 1:
    b_value = float(numbers[0])
else:
    b_value = 0.0   # or whatever default value you want
# --- Read C15 ---
c15_value = float(ws["C15"].value or 0)

# --- Final B15 calculation ---
b15_value = total_qty_b15 * b_value * 1.0 + c15_value
ws["B15"] = b15_value

# ------------------------------------------------
# FIND SPTS&SOC or SPTSSOC IN ws2 COLUMN D
# AND WRITE VALUE FROM COLUMN F INTO ws["C64"]
# ------------------------------------------------

target_codes = ["SPTS&SOC", "SPTSSOC"]
total_value = 0

for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""

    if any(code in cell_text for code in target_codes):
        row_index = row[0].row
        f_value = ws2[f"F{row_index}"].value

        try:
            total_value += float(f_value or 0)
        except:
            total_value += 0

# Write result to Calc Template C64
ws["C64"] = total_value

# ------------------------------------------------
# SCAN ws2 COLUMN D FOR SPECIFIC CODES
# READ VALUE FROM COLUMN F
# WRITE INTO MULTIPLE TARGET CELLS IN ws (Calc Template)
# ------------------------------------------------

code_map = {
    "FIT CENT": ["C65", "B65"],
    "NSSC GYM": ["C66", "B66"],
    "DENPLAN": ["C74", "B74"],
    "UNION": ["C71", "B71"],
    "GAYE": ["C72", "B72"],
    "FLY DAYS": ["C23", "B23"]
}

# Initialize all target cells to 0
for cells in code_map.values():
    for cell in cells:
        ws[cell] = 0

# Scan column D
for row in ws2.iter_rows(min_col=4, max_col=4):
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    for code, target_cells in code_map.items():
        if code in cell_text:
            f_value = ws2[f"F{row_index}"].value

            try:
                numeric_value = float(f_value or 0)
            except:
                numeric_value = 0

            # Write into all assigned cells
            for cell in target_cells:
                ws[cell] = numeric_value


# ------------------------------------------------
# SCAN ws2 COLUMN D FOR CAR DRAW & HOL DRAW
# WRITE COLUMN F VALUES INTO MULTIPLE TARGET CELLS
# ------------------------------------------------

multi_code_map = {
    "CAR DRAW": ["C70", "B70"],
    "HOL DRAW": ["C69", "B69"]
}

# Initialize all target cells to 0
for cells in multi_code_map.values():
    for cell in cells:
        ws[cell] = 0

# Scan column D
for row in ws2.iter_rows(min_col=4, max_col=4):
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    for code, target_cells in multi_code_map.items():
        if code in cell_text:
            f_value = ws2[f"F{row_index}"].value

            try:
                numeric_value = float(f_value or 0)
            except:
                numeric_value = 0

            # Write into all assigned cells
            for cell in target_cells:
                ws[cell] = numeric_value

# ------------------------------------------------
# SCAN ws2 COLUMN D FOR NUFCDRAW & SAFCDRAW
# WRITE COLUMN F VALUES INTO MULTIPLE TARGET CELLS
# ------------------------------------------------

draw_code_map = {
    "NUFCDRAW": ["C68", "B68"],
    "SAFCDRAW": ["C67", "B67"]
}

# Initialize all target cells to 0
for cells in draw_code_map.values():
    for cell in cells:
        ws[cell] = 0

# Scan column D
for row in ws2.iter_rows(min_col=4, max_col=4):
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    for code, target_cells in draw_code_map.items():
        if code in cell_text:
            f_value = ws2[f"F{row_index}"].value

            try:
                numeric_value = float(f_value or 0)
            except:
                numeric_value = 0

            # Write into all assigned cells
            for cell in target_cells:
                ws[cell] = numeric_value

# ------------------------------------------------
# DEDUCTION LOGIC FOR DRAW CODES (COLUMN D)
# VALUES MUST BE NEGATIVE
# ------------------------------------------------

# Read leaving date from Calc Template B2
leaving_date = ws["B2"].value

# Convert to datetime if needed
if isinstance(leaving_date, str):
    leaving_date = datetime.strptime(leaving_date, "%d/%m/%Y")

leave_day = leaving_date.day

# Deduction allowed only if leaving on or after 25th
deduction_allowed = leave_day >= 25

deduction_map = {
    "SAFCDRAW": "D67",
    "NUFCDRAW": "D68",
    "HOL DRAW": "D69",
    "CAR DRAW": "D70"
}

# Initialize all deduction cells to 0
for cell in deduction_map.values():
    ws[cell] = 0

# Scan column D of Excelpayslip
for row in ws2.iter_rows(min_col=4, max_col=4):
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    for code, target_cell in deduction_map.items():
        if code in cell_text:

            # If deduction is NOT allowed → leave as 0
            if not deduction_allowed:
                ws[target_cell] = 0
                continue

            # Deduction allowed → read column F
            f_value = ws2[f"F{row_index}"].value

            try:
                numeric_value = float(f_value or 0)
            except:
                numeric_value = 0

            # Deduction must be NEGATIVE
            ws[target_cell] = -abs(numeric_value)
# ------------------------------------------------
# IF LEAVING DATE IS BEFORE 25TH
# CLEAR B67, B68, B69, B70
# ------------------------------------------------

if leave_day < 25:

    ws["B67"] = 0
    ws["B68"] = 0
    ws["B69"] = 0
    ws["B70"] = 0

    print("Leave day is before 25th.")
    print("B67, B68, B69 and B70 set to 0.")

# --------------------------------------------
# D67 = B67 - C67
# --------------------------------------------
ws["D67"] = round(
    float(ws["B67"].value or 0) -
    float(ws["C67"].value or 0),
    2
)

# --------------------------------------------
# D68 = B68 - C68
# --------------------------------------------
ws["D68"] = round(
    float(ws["B68"].value or 0) -
    float(ws["C68"].value or 0),
    2
)

# --------------------------------------------
# D69 = B69 - C69
# --------------------------------------------
ws["D69"] = round(
    float(ws["B69"].value or 0) -
    float(ws["C69"].value or 0),
    2
)

# --------------------------------------------
# D70 = B70 - C70
# --------------------------------------------
ws["D70"] = round(
    float(ws["B70"].value or 0) -
    float(ws["C70"].value or 0),
    2
)

# ------------------------------------------------
# SCAN ws2 COLUMN D FOR JDEBT J1
# WRITE COLUMN F VALUE INTO C46
# ------------------------------------------------

import re

ws["C46"] = 0

for row in ws2.iter_rows(min_col=4, max_col=4):
    raw = row[0].value
    cell_text = str(raw).upper()

    # Normalize all whitespace
    cell_text_clean = re.sub(r"\s+", " ", cell_text).strip()

    if cell_text_clean == "JDEBT J1":
        row_index = row[0].row
        f_value = ws2[f"F{row_index}"].value

        try:
            ws["C46"] = float(f_value or 0)
        except:
            ws["C46"] = 0

        break

from datetime import datetime
import calendar

# ------------------------------------------------
# PPP → C50 (direct value)
# PPP → B50 (calculated value)
# ------------------------------------------------

ws["C50"] = 0
ws["B50"] = 0

# Read date from B2
date_in_b2 = ws["B2"].value
if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

year = date_in_b2.year
month = date_in_b2.month
leave_day = date_in_b2.day

# Total days in month
total_days_in_month = calendar.monthrange(year, month)[1]

# ⭐ Working days = from 1st → leaving date (includes Sat + Sun)
working_days = leave_day

# Scan column D for PPP
for row in ws2.iter_rows(min_col=4, max_col=4):
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    if "PPP" in cell_text:
        f_value = ws2[f"F{row_index}"].value

        try:
            c50_value = float(f_value or 0)
        except:
            c50_value = 0

        ws["C50"] = c50_value

        # ⭐ Correct formula:
        ws["B50"] = (c50_value / total_days_in_month) * working_days

        break
# ------------------------------------------------
# ECSF → C62 (direct value only)
# ------------------------------------------------

ws["C62"] = 0   # initialize

# Scan Column D of Excelpayslip for ECSF
for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    # Must match ECSF EXACTLY, not RECSF
    if cell_text == "ECSF":
        f_value = ws2[f"F{row_index}"].value

        try:
            ecsf_amount = float(f_value or 0)
        except:
            ecsf_amount = 0

        ws["C62"] = ecsf_amount
        break
# ------------------------------------------------
# B62 = C62 × number of previous months (Feb → previous month)
# ALWAYS NEGATIVE
# D62 = B62 - C62
# ------------------------------------------------

c62_value = ws["C62"].value
if c62_value is None:
    c62_value = 0

# Read date from B2
date_in_b2 = ws["B2"].value
if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

month = date_in_b2.month

# Count previous months from Feb → previous month
# Feb = 0, Mar = 1, ..., Nov = 9
if month < 2:
    months_count = 0
elif 2 <= month <= 11:
    months_count = month - 2
else:
    months_count = 9   # December or later → full previous period

# Calculate B62
b62_value = c62_value * months_count

# ALWAYS NEGATIVE
b62_value = -abs(b62_value)

ws["B62"] = b62_value

# D62 = B62 - C62
ws["D62"] = b62_value - c62_value


# ------------------------------------------------
# RECSF → C59 and B59 (direct value only, C = negative)
# ------------------------------------------------

ws["C59"] = 0
ws["B59"] = 0

def parse_recsf(value):
    if value is None:
        return 0

    s = str(value).strip().upper()

    # Check if last character is C (negative)
    is_negative = s.endswith("C")

    # Remove all non-numeric characters except dot and minus
    s_clean = re.sub(r"[^0-9.\-]", "", s)

    try:
        num = float(s_clean)
    except:
        num = 0

    # Apply negative rule
    if is_negative:
        num = -abs(num)
    else:
        num = abs(num)

    return num

# Scan Column D of Excelpayslip for RECSF
for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    if cell_text == "RECSF":
        f_value = ws2[f"F{row_index}"].value
        recsf_amount = parse_recsf(f_value)

        # Write directly into C59 and B59
        ws["C59"] = recsf_amount
        ws["B59"] = recsf_amount

        break
# ------------------------------------------------
# ESSF → C63 (direct value only)
# ------------------------------------------------

ws["C63"] = 0   # initialize

def clean_number(value):
    if value is None:
        return 0
    s = str(value).strip().upper()
    s = re.sub(r"[^0-9.\-]", "", s)  # remove letters
    try:
        return float(s)
    except:
        return 0

# Scan Column D of Excelpayslip for ESSF
for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    if cell_text == "ESSF":
        f_value = ws2[f"F{row_index}"].value
        essf_amount = clean_number(f_value)

        ws["C63"] = essf_amount
        break
# ------------------------------------------------
# B63 = C63 × number of previous months (Sep → previous month)
# Negative until May, positive in June
# ------------------------------------------------

c63_value = ws["C63"].value or 0

# Read date from B2
date_in_b2 = ws["B2"].value
if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

month = date_in_b2.month

# Determine previous months count (Sep → previous month)
if month >= 9:            # Sep–Dec
    months_count = month - 9
elif 1 <= month <= 6:     # Jan–Jun
    months_count = month + 3
else:
    months_count = 0      # July–August → outside saving period

# Calculate B63
b63_value = c63_value * months_count

# Apply sign rule
if month == 6:            # June → positive
    b63_value = abs(b63_value)
else:                     # Sep → May → negative
    b63_value = -abs(b63_value)

ws["B63"] = b63_value

# ------------------------------------------------
# NEW LOGIC: D62 = B62 - C62
# ------------------------------------------------
ws["D63"] = b63_value - c63_value
# ------------------------------------------------
# EXP DED → C52 and B52 (direct value only, C = negative)
# ------------------------------------------------

ws["C52"] = 0
ws["B52"] = 0

def parse_expded(value):
    if value is None:
        return 0

    s = str(value).strip().upper()

    # If last character is C → negative
    is_negative = s.endswith("C")

    # Remove all non-numeric characters except dot and minus
    s_clean = re.sub(r"[^0-9.\-]", "", s)

    try:
        num = float(s_clean)
    except:
        num = 0

    # Apply negative rule
    if is_negative:
        num = -abs(num)
    else:
        num = abs(num)

    return num

# Scan Column D of Excelpayslip for EXP DED
for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    if cell_text == "EXP DED":
        f_value = ws2[f"F{row_index}"].value
        expded_amount = parse_expded(f_value)

        ws["C52"] = expded_amount
        ws["B52"] = expded_amount

        break

# ------------------------------------------------
# SAL.ADV. → B51 and C51 (direct value only, C = negative)
# ------------------------------------------------

ws["B51"] = 0   # initialize
ws["C51"] = 0   # initialize

def parse_saladv(value):
    if value is None:
        return 0

    s = str(value).strip().upper()

    # If last character is C → negative
    is_negative = s.endswith("C")

    # Remove all non-numeric characters except dot and minus
    s_clean = re.sub(r"[^0-9.\-]", "", s)

    try:
        num = float(s_clean)
    except:
        num = 0

    # Apply negative rule
    if is_negative:
        num = -abs(num)
    else:
        num = abs(num)

    return num

# Scan Column D of Excelpayslip for SAL.ADV.
for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).upper().strip() if row[0].value else ""
    row_index = row[0].row

    if cell_text == "SAL.ADV.":
        f_value = ws2[f"F{row_index}"].value
        saladv_amount = parse_saladv(f_value)

        ws["B51"] = saladv_amount
        ws["C51"] = saladv_amount   # ← COPY VALUE TO C51

        break

# ------------------------------------------------
# RCYCLES1 / RCYCLES2 → C40 (Negative Value)
# B40 = 0
# ------------------------------------------------

ws["B40"] = 0
ws["C40"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):
    cell_a = row[0].value

    if cell_a:
        cell_text = str(cell_a).strip().upper()

        if "RCYCLES1" in cell_text or "RCYCLES2" in cell_text:
            row_num = row[0].row
            value = ws2[f"C{row_num}"].value

            try:
                ws["C40"] = -abs(float(value))
            except (TypeError, ValueError):
                ws["C40"] = 0

            ws["B40"] = 0
            break
        # D40 = B40 - C40
ws["D40"] = (ws["B40"].value or 0) - (ws["C40"].value or 0)
# ------------------------------------------------
# B73 = D40 × Remaining Months Count
# Cycle: May → April
# Current month excluded
# ------------------------------------------------

from datetime import datetime

date_in_b2 = ws["B2"].value

if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

month = date_in_b2.month

# Completed months before current month
if month >= 5:      # May-Dec
    months_count = month - 5
else:               # Jan-Apr
    months_count = month + 7

# Remaining months in cycle
remaining_months = 11 - months_count

# B73 = D40 × remaining months
ws["B73"] = (ws["D40"].value or 0) * remaining_months

# D73 = B73 - C73
ws["D73"] = (ws["B73"].value or 0) - (ws["C73"].value or 0)
ws["D73"] = (ws["B73"].value or 0) - (ws["C73"].value or 0)

# ------------------------------------------------
# RGAYE NEW → B29, C29
# D29 = B29 - C29
# Search Column A in Excelpayslip and get value from Column C
# ------------------------------------------------

ws["B29"] = 0
ws["C29"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_a = row[0].value

    if cell_a:
        cell_text = str(cell_a).strip().upper()

        if "RGAYE NEW" in cell_text:
            row_num = row[0].row
            value = ws2[f"C{row_num}"].value

            try:
                value = float(value)
            except (TypeError, ValueError):
                value = 0

            ws["B29"] = value
            ws["C29"] = value
            break

# D29 = B29 - C29
ws["D29"] = (ws["B29"].value or 0) - (ws["C29"].value or 0)
ws["C5"] = float(str(ws["C5"].value or 0).replace(",", ""))
ws["C16"] = float(str(ws["C16"].value or 0).replace(",", ""))
# ------------------------------------------------
# WORKWEAR → B53, C53
# Search Column D in Excelpayslip and get value from Column F
# D53 = B53 - C53
# ------------------------------------------------

ws["B53"] = 0
ws["C53"] = 0

for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "WORKWEAR":
        value = ws2[f"F{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except (TypeError, ValueError):
            value = 0

        ws["B53"] = value
        ws["C53"] = value
        break

# D53 = B53 - C53
ws["D53"] = (ws["B53"].value or 0) - (ws["C53"].value or 0)

# ------------------------------------------------
# RENT → B54, C54
# Search Column D in Excelpayslip and get value from Column F
# D54 = B54 - C54
# ------------------------------------------------

ws["B54"] = 0
ws["C54"] = 0

for row in ws2.iter_rows(min_col=4, max_col=4):  # Column D
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "RENT":
        value = ws2[f"F{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except (TypeError, ValueError):
            value = 0

        ws["B54"] = value
        ws["C54"] = value
        break

# D54 = B54 - C54
ws["D54"] = (ws["B54"].value or 0) - (ws["C54"].value or 0)

import re

# ------------------------------------------------
# LeaverForm - Extract PILON / Holiday Pay / Ex Gratia
# ------------------------------------------------

leaver_ws = wb["LeaverForm"]

pilon_value = 0
holiday_value = 0
exgratia_value = 0

for row in leaver_ws.iter_rows():
    for cell in row:

        if cell.value is None:
            continue

        text = str(cell.value).strip()

        # Debug - Uncomment if required
        # print(cell.coordinate, "=", text)

        # PILON
        pilon_match = re.search(
            r'([\d,]+(?:\.\d+)?)\s*PILON',
            text,
            re.IGNORECASE
        )

        if pilon_match:
            pilon_value = float(pilon_match.group(1).replace(",", ""))

        # Holiday Pay
        holiday_match = re.search(
            r'([\d,]+(?:\.\d+)?)\s*HOLIDAY\s*PAY',
            text,
            re.IGNORECASE
        )

        if holiday_match:
            holiday_value = float(holiday_match.group(1).replace(",", ""))

        # Ex Gratia
        exgratia_match = re.search(
            r'([\d,]+(?:\.\d+)?)\s*EX\s*GRATIA|([\d,]+(?:\.\d+)?)\s*EXGRATIA',
            text,
            re.IGNORECASE
        )

        if exgratia_match:
            exgratia_amount = (
                exgratia_match.group(1)
                if exgratia_match.group(1)
                else exgratia_match.group(2)
            )

            exgratia_value = float(exgratia_amount.replace(",", ""))

# ------------------------------------------------
# PILON -> B37
# ------------------------------------------------

ws["B37"] = pilon_value

# D37 = B37 - C37
ws["D37"] = (ws["B37"].value or 0) - (ws["C37"].value or 0)

# ------------------------------------------------
# Holiday Pay -> B21
# ------------------------------------------------
ws["B21"] = holiday_value

# D21 = B21 - C21
ws["D21"] = (ws["B21"].value or 0) - (ws["C21"].value or 0)
# ------------------------------------------------
# Ex Gratia -> B76 / B77
# ------------------------------------------------

if exgratia_value <= 30000:
    ws["B76"] = exgratia_value
    ws["B77"] = 0
else:
    ws["B76"] = 30000
    ws["B77"] = exgratia_value - 30000

# D76 = B76 - C76
ws["D76"] = (ws["B76"].value or 0) - (ws["C76"].value or 0)

# D77 = B77 - C77
ws["D77"] = (ws["B77"].value or 0) - (ws["C77"].value or 0)
# D7 = B7 - C7
ws["D7"] = (float(ws["B7"].value or 0)) - (float(ws["C7"].value or 0))

# D16 = B16 - C16
ws["D16"] = (float(ws["B16"].value or 0)) - (float(ws["C16"].value or 0))

# D22 = B22 - C22
ws["D22"] = (float(ws["B22"].value or 0)) - (float(ws["C22"].value or 0))

# D23 = B23 - C23
ws["D23"] = (float(ws["B23"].value or 0)) - (float(ws["C23"].value or 0))

# D15 = B15 - C15
ws["D15"] = (float(ws["B15"].value or 0)) - (float(ws["C15"].value or 0))

# D17 = B17 - C17
ws["D17"] = (float(ws["B17"].value or 0)) - (float(ws["C17"].value or 0))
# D9 = B9 - C9
ws["D9"] = (float(ws["B9"].value or 0)) - (float(ws["C9"].value or 0))

# D11 = B11 - C11
ws["D11"] = (float(ws["B11"].value or 0)) - (float(ws["C11"].value or 0))

from datetime import datetime
import calendar

# ------------------------------------------------
# HOUSING -> C10
# Search Column A in Excelpayslip
# Value from Column C
# ------------------------------------------------

ws["C10"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "HOUSING":
        value = ws2[f"C{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except:
            value = 0

        ws["C10"] = value
        break

from datetime import datetime, date
import calendar

# ------------------------------------------------
# HOUSING -> C10
# Search Column A in Excelpayslip
# Value from Column C
# ------------------------------------------------

ws["B10"] = 0
ws["C10"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "HOUSING":
        value = ws2[f"C{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except:
            value = 0

        ws["C10"] = value
        break

# ------------------------------------------------
# B10 Calculation
# Mon-Fri Only
# B10 = C10 / Total Working Days * Worked Days
# ------------------------------------------------

date_in_b2 = ws["B2"].value

if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

year = date_in_b2.year
month = date_in_b2.month
day = date_in_b2.day

# Total working days in month (Mon-Fri only)
days_in_month = calendar.monthrange(year, month)[1]

total_working_days = 0

for d in range(1, days_in_month + 1):
    current_date = date(year, month, d)

    if current_date.weekday() < 5:   # Mon=0 ... Fri=4
        total_working_days += 1

# Working days up to B2 date (inclusive)
worked_days = 0

for d in range(1, day + 1):
    current_date = date(year, month, d)

    if current_date.weekday() < 5:
        worked_days += 1

c10_value = ws["C10"].value or 0

if total_working_days > 0:
    b10_value = (c10_value / total_working_days) * worked_days
else:
    b10_value = 0

ws["B10"] = round(b10_value, 2)

# ------------------------------------------------
# D10 = B10 - C10
# ------------------------------------------------

ws["D10"] = (ws["B10"].value or 0) - (ws["C10"].value or 0)

# ------------------------------------------------
# UKSIPREF -> B38
# Search Column A in Excelpayslip
# Value from Column C
# D38 = B38 - C38
# ------------------------------------------------

ws["B38"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "UKSIPREF":
        value = ws2[f"C{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except:
            value = 0

        ws["B38"] = value
        break

# D38 = B38 - C38
ws["D38"] = (ws["B38"].value or 0) - (ws["C38"].value or 0)

# ------------------------------------------------
# UKSIP -> B36
# Search Column A in Excelpayslip
# Value from Column C
# D36 = B36 - C36
# ------------------------------------------------

ws["B36"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "UKSIP":
        value = ws2[f"C{row_num}"].value

        try:
            value = float(str(value).replace(",", ""))
        except:
            value = 0

        ws["B36"] = value
        break

# D36 = B36 - C36
ws["D36"] = (ws["B36"].value or 0) - (ws["C36"].value or 0)

from datetime import datetime, date
import calendar

# ------------------------------------------------
# DEP PAY -> B39
# Search Column A, value from Column C
# ------------------------------------------------

dep_pay = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "DEP PAY":
        value = ws2[f"C{row_num}"].value

        try:
            dep_pay = float(str(value).replace(",", ""))
        except:
            dep_pay = 0

        break

# ------------------------------------------------
# Get date from B2
# ------------------------------------------------

date_in_b2 = ws["B2"].value

if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

year = date_in_b2.year
month = date_in_b2.month
day = date_in_b2.day

# ------------------------------------------------
# Count working days worked in current month
# Mon-Fri only
# ------------------------------------------------

worked_days = 0

for d in range(1, day + 1):
    current_date = date(year, month, d)

    if current_date.weekday() < 5:  # Mon-Fri
        worked_days += 1

# ------------------------------------------------
# B39 Calculation
# DEP PAY × 12 / 260 × worked days
# ------------------------------------------------

b39_value = (dep_pay * 12 / 260) * worked_days

ws["B39"] = round(b39_value, 2)

# D39 = B39 - C39
ws["D39"] = (ws["B39"].value or 0) - (ws["C39"].value or 0)
from datetime import datetime, date
import calendar

# ------------------------------------------------
# FLEXIBAL -> B41
# Search Column A in Excelpayslip
# Value from Column C
# ------------------------------------------------

flexibal_value = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "FLEXIBAL":
        value = ws2[f"C{row_num}"].value

        try:
            flexibal_value = float(str(value).replace(",", ""))
        except:
            flexibal_value = 0

        break

# ------------------------------------------------
# Get date from B2
# ------------------------------------------------

date_in_b2 = ws["B2"].value

if isinstance(date_in_b2, str):
    date_in_b2 = datetime.strptime(date_in_b2, "%d/%m/%Y")

year = date_in_b2.year
month = date_in_b2.month
day = date_in_b2.day

# ------------------------------------------------
# Total Working Days in Month (Mon-Fri)
# ------------------------------------------------

days_in_month = calendar.monthrange(year, month)[1]

total_working_days = 0

for d in range(1, days_in_month + 1):
    if date(year, month, d).weekday() < 5:
        total_working_days += 1

# ------------------------------------------------
# Worked Days up to B2 Date (Mon-Fri)
# ------------------------------------------------

worked_days = 0

for d in range(1, day + 1):
    if date(year, month, d).weekday() < 5:
        worked_days += 1

# ------------------------------------------------
# B41 Calculation
# B41 = FLEXIBAL / Total Working Days * Worked Days
# ------------------------------------------------

if total_working_days > 0:
    b41_value = (flexibal_value / total_working_days) * worked_days
else:
    b41_value = 0

ws["B41"] = round(b41_value, 2)

# D41 = B41 - C41
ws["D41"] = (ws["B41"].value or 0) - (ws["C41"].value or 0)
# ------------------------------------------------
# RC VOUCH -> B42
# Search Column A in Excelpayslip
# Value from Column C
# If value ends with 'C' => Negative
# D42 = B42 - C42
# ------------------------------------------------

ws["B42"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "RC VOUCH":

        raw_value = ws2[f"C{row_num}"].value

        if raw_value is not None:
            value_str = str(raw_value).strip().upper()

            # Check if value ends with C
            is_negative = value_str.endswith("C")

            # Remove non-numeric characters except . and -
            value_str = re.sub(r"[^0-9.\-]", "", value_str)

            try:
                amount = float(value_str)

                if is_negative:
                    amount = -abs(amount)
                else:
                    amount = abs(amount)

            except:
                amount = 0
        else:
            amount = 0

        ws["B42"] = amount
        break

# D42 = B42 - C42
ws["D42"] = (ws["B42"].value or 0) - (ws["C42"].value or 0)

# ------------------------------------------------
# RABS A/B -> C27
# Search Column A in Excelpayslip
# Value from Column C
# Trailing C = Negative Value
# ------------------------------------------------

ws["C27"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):  # Column A
    cell_text = str(row[0].value).strip().upper() if row[0].value else ""
    row_num = row[0].row

    if cell_text == "RABS A/B":

        raw_value = ws2[f"C{row_num}"].value

        if raw_value is not None:
            value_str = str(raw_value).strip().upper()

            # Check for trailing C
            is_negative = value_str.endswith("C")

            # Keep only numeric characters
            value_str = re.sub(r"[^0-9.\-]", "", value_str)

            try:
                amount = float(value_str)

                if is_negative:
                    amount = -abs(amount)
                else:
                    amount = abs(amount)

            except:
                amount = 0
        else:
            amount = 0

        ws["C27"] = amount
        break

# D27 = B27 - C27
ws["D27"] = (ws["B27"].value or 0) - (ws["C27"].value or 0)

# ------------------------------------------------
# Absence details
# Column J = Time Off
# Column F = Unit
# Set Unit to 0 when Time Off matches
# ------------------------------------------------

absence_ws = wb["Absence details"]

for row in range(3, absence_ws.max_row + 1):  # Data starts after header row 2

    time_off = absence_ws[f"J{row}"].value

    if time_off is not None:

        time_off_text = str(time_off).strip().upper()

        if time_off_text == "NMUK - PAID ABSENCE - AUTHORISED TIME OFF".upper():

            print(f"Match found at row {row}")

            absence_ws[f"F{row}"] = 0

# ------------------------------------------------
# Get Shift Code (DF / DL / MF)
# ------------------------------------------------

shift_code = ""

for row in ws2.iter_rows(min_col=1, max_col=1):

    text = str(row[0].value or "").strip().upper()

    if "SHIFT DF" in text:
        shift_code = "DF"
        break

    elif "SHIFT DL" in text:
        shift_code = "DL"
        break

    elif "SHIFT MF" in text:
        shift_code = "MF"
        break

print("Shift Code =", shift_code)

# ------------------------------------------------
# Run only if Shift Code exists
# ------------------------------------------------

if not shift_code:
    print("No shift code found. Skipping absence unit update.")

else:

    # ------------------------------------------------
    # Update Unit Column F
    # ------------------------------------------------

    for r in range(3, absence_ws.max_row + 1):

        time_off = str(
            absence_ws[f"J{r}"].value or ""
        ).strip().upper()

        if time_off == "NMUK - PAID ABSENCE - AUTHORISED TIME OFF":
            continue

        calc_tag = str(
            absence_ws[f"G{r}"].value or ""
        ).strip().upper()

        shift_value = 0
        table_ws = wb["TA-BH"]
        for t in range(2, table_ws.max_row + 1):

            table_shift = str(
                table_ws[f"A{t}"].value or ""
            ).strip().upper()

            table_tag = str(
                table_ws[f"C{t}"].value or ""
            ).strip().upper()

            if table_shift == shift_code and table_tag == calc_tag:

                value = table_ws[f"B{t}"].value

                try:
                    shift_value = float(str(value).replace(",", ""))
                except:
                    shift_value = 0

                break

        try:
            reported_qty = float(
                absence_ws[f"E{r}"].value or 0
            )
        except:
            reported_qty = 0

        if abs(reported_qty - 3.941667) < 0.000001:
            absence_ws[f"F{r}"] = shift_value / 2
        else:
            absence_ws[f"F{r}"] = shift_value

    # ------------------------------------------------
    # Add Unit Column Total At Bottom
    # ------------------------------------------------

    total_unit = 0
    last_data_row = 0

    for r in range(3, absence_ws.max_row + 1):

        try:
            unit_val = float(
                absence_ws[f"F{r}"].value or 0
            )

            total_unit += unit_val

            if absence_ws[f"F{r}"].value not in [None, ""]:
                last_data_row = r

        except:
            pass

    absence_ws[f"F{last_data_row + 1}"] = total_unit
# ------------------------------------------------
# B27 Calculation
# Run ONLY if shift code exists
# ------------------------------------------------

if shift_code:

    cntrcthr_rate = 0

    for row in ws2.iter_rows():

        col_a = str(row[0].value or "").strip().upper()

        if "CNTRCTHR" in col_a:

            next_row = row[0].row + 1

            formula_text = str(
                ws2[f"A{next_row}"].value or ""
            )

            if "X" in formula_text.upper():

                try:
                    rate_text = (
                        formula_text.upper()
                        .replace("(", "")
                        .replace(")", "")
                        .split("X")[1]
                        .strip()
                    )

                    cntrcthr_rate = float(rate_text)

                except:
                    cntrcthr_rate = 0

            break

    b27_value = -(total_unit * cntrcthr_rate) + float(ws["C27"].value or 0)

    ws["B27"] = round(b27_value, 2)

    # D27 = B27 - C27
    ws["D27"] = (ws["B27"].value or 0) - (ws["C27"].value or 0)

else:
    print("No shift code found. Skipping B27 calculation.")

# ------------------------------------------------
# RABS C -> C26
# Search Column A in Excelpayslip
# Read value from Column C
# Trailing C = Negative
# ------------------------------------------------

ws["C26"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):

    code = str(row[0].value or "").strip().upper()
    row_num = row[0].row

    if code == "RABS C":

        raw_value = str(
            ws2[f"C{row_num}"].value or ""
        ).strip().upper()

        try:
            if raw_value.endswith("C"):

                amount = float(
                    raw_value[:-1].replace(",", "")
                )

                ws["C26"] = -abs(amount)

            else:

                ws["C26"] = float(
                    raw_value.replace(",", "")
                )

        except:
            ws["C26"] = 0

        break

print("C26 =", ws["C26"].value)
# ------------------------------------------------
# UNPAID ABSENCE
# Update Unit Column F
# ------------------------------------------------

absence_ws = wb["Absence details"]
table_ws = wb["Absence Table"]

# ------------------------------------------------
# Get Shift Code (DF / DL / MF)
# ------------------------------------------------

shift_code = ""

for row in ws2.iter_rows(min_col=1, max_col=1):

    text = str(row[0].value or "").strip().upper()

    if "SHIFT DF" in text:
        shift_code = "DF"
        break

    elif "SHIFT DL" in text:
        shift_code = "DL"
        break

    elif "SHIFT MF" in text:
        shift_code = "MF"
        break

print("Shift Code =", shift_code)
# ------------------------------------------------
# CHECK SHIFT CODE
# ------------------------------------------------

if not shift_code:
    print("No shift code found. Skipping Unpaid Absence calculations.")
else:
   
    # ------------------------------------------------
    # Build Mapping Dictionary
    # Absence Table
    # Col E = Shift
    # Col F = Shift Value
    # Col G = Calculation Tags
    # ------------------------------------------------

    mapping = {}

    for r in range(2, table_ws.max_row + 1):

        shift = str(
            table_ws[f"E{r}"].value or ""
        ).strip().upper()

        calc_tag = str(
            table_ws[f"G{r}"].value or ""
        ).replace(" ", "").strip().upper()

        value = table_ws[f"F{r}"].value

        if not shift or not calc_tag:
            continue

        value_text = str(value).strip().upper()

        if value_text == "NONE":
            mapped_value = 0
        else:
            try:
                mapped_value = float(
                    str(value).replace(",", "")
                )
            except:
                mapped_value = 0

        mapping[(shift, calc_tag)] = mapped_value

    # ------------------------------------------------
    # Update Unit Column F
    # ------------------------------------------------

    for r in range(3, absence_ws.max_row + 1):

        time_off = str(
            absence_ws[f"J{r}"].value or ""
        ).strip().upper()

        if time_off == "NMUK - PAID ABSENCE - AUTHORISED TIME OFF":
            continue

        calc_tag = str(
            absence_ws[f"G{r}"].value or ""
        ).replace(" ", "").strip().upper()

        shift_value = mapping.get(
            (shift_code, calc_tag),
            0
        )

        # --------------------------------------------
        # Reported Quantity Logic
        # --------------------------------------------

        try:
            reported_qty = float(
                absence_ws[f"E{r}"].value or 0
            )
        except:
            reported_qty = 0

        if abs(reported_qty - 3.941667) < 0.000001:
            absence_ws[f"F{r}"] = shift_value / 2
        else:
            absence_ws[f"F{r}"] = shift_value

# ------------------------------------------------
# Calculate Total Unit
# ------------------------------------------------

total_unit = 0

for r in range(3, absence_ws.max_row + 1):

    try:
        total_unit += float(
            absence_ws[f"F{r}"].value or 0
        )
    except:
        pass

print("Total Unit =", total_unit)

# ------------------------------------------------
# Get CNTRCTHR Rate
# ------------------------------------------------

cntrcthr_rate = 0

for row in ws2.iter_rows():

    col_a = str(row[0].value or "").strip().upper()

    if "CNTRCTHR" in col_a:

        next_row = row[0].row + 1

        formula_text = str(
            ws2[f"A{next_row}"].value or ""
        )

        try:
            rate_text = (
                formula_text.upper()
                .replace("(", "")
                .replace(")", "")
                .split("X")[1]
                .strip()
            )

            cntrcthr_rate = float(rate_text)

        except:
            cntrcthr_rate = 0

        break

print("CNTRCTHR Rate =", cntrcthr_rate)

# ------------------------------------------------
# B26 = -(Total Unit * Rate) + C26
# ------------------------------------------------

b26_value = -(total_unit * cntrcthr_rate) + (ws["C26"].value or 0)

ws["B26"] = round(b26_value, 2)

# D26 = B26 - C26
ws["D26"] = (ws["B26"].value or 0) - (ws["C26"].value or 0)

print("B26 =", ws["B26"].value)
print("D26 =", ws["D26"].value)

# ------------------------------------------------
# S.S.P. -> C28
# Search Column A in Excelpayslip
# Value from Column C
# Trailing C = Negative
# ------------------------------------------------

ws["C28"] = 0

for row in ws2.iter_rows(min_col=1, max_col=1):

    code = str(row[0].value or "").strip().upper()

    if code == "S.S.P.":

        raw_value = str(
            ws2[f"C{row[0].row}"].value or ""
        ).strip().upper()

        try:
            if raw_value.endswith("C"):

                amount = float(
                    raw_value[:-1].replace(",", "").strip()
                )

                ws["C28"] = -abs(amount)

            else:

                ws["C28"] = float(
                    raw_value.replace(",", "")
                )

        except:
            ws["C28"] = 0

        break

print("C28 =", ws["C28"].value)

# ------------------------------------------------
# B28 - UK-NMUK Sickness
# ------------------------------------------------

sickness_sum = 0

for r in range(3, absence_ws.max_row + 1):

    time_off = str(
        absence_ws[f"J{r}"].value or ""
    ).strip().upper()

    try:
        reported_qty = float(
            absence_ws[f"E{r}"].value or 0
        )
    except:
        reported_qty = 0

    if time_off == "UK-NMUK SICKNESS" and reported_qty >= 1:
        sickness_sum += reported_qty

# ------------------------------------------------
# B28 - UK-NMUK Sickness
# Option 2 Logic
# Count rows where Reported Quantity >= 1
# ------------------------------------------------

sickness_count = 0

for r in range(3, absence_ws.max_row + 1):

    time_off = str(
        absence_ws[f"J{r}"].value or ""
    ).strip().upper()

    try:
        reported_qty = float(
            absence_ws[f"E{r}"].value or 0
        )
    except:
        reported_qty = 0

    if (
        time_off == "UK-NMUK SICKNESS"
        and reported_qty >= 1
    ):
        sickness_count += 1

# ------------------------------------------------
# B28 = (123.25 / 5) * Count
# ------------------------------------------------

ws["B28"] = round((123.25 / 5) * sickness_count, 2)

# D28 = B28 - C28
ws["D28"] = (ws["B28"].value or 0) - (ws["C28"].value or 0)

print("Sickness Count =", sickness_count)
print("B28 =", ws["B28"].value)
print("D28 =", ws["D28"].value)


# ==========================================================
# COUNCIL TAX -> C47
# Codes in Column G
# Values in Column I
# ==========================================================
council_tax_value = 0

for row_num in range(2, ws2.max_row + 1):

    council_code = str(
        ws2[f"G{row_num}"].value or ""
    ).strip().upper()

    if (
        council_code.startswith("T1")
        or council_code.startswith("T2")
        or council_code.startswith("T3")
    ):

        council_tax_value += float(
            ws2[f"I{row_num}"].value or 0
        )

ws["C47"] = round(council_tax_value, 2)

# ==========================================================
# D47 = B47 - C47
# ==========================================================
ws["D47"] = round(
    (ws["B47"].value or 0) -
    (ws["C47"].value or 0),
    2
)

# ==========================================================
# CHILD MAINTENANCE -> C48
# Codes: DEO D1 to DEO D9
# Column D = Code
# Column F = Value
# ==========================================================
child_maintenance_value = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_code = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if (
        deduction_code == "DEO D1"
        or deduction_code == "DEO D2"
        or deduction_code == "DEO D3"
        or deduction_code == "DEO D4"
        or deduction_code == "DEO D5"
        or deduction_code == "DEO D6"
        or deduction_code == "DEO D7"
        or deduction_code == "DEO D8"
        or deduction_code == "DEO D9"
    ):

        child_maintenance_value += float(
            ws2[f"F{row_num}"].value or 0
        )

ws["C48"] = round(child_maintenance_value, 2)

# ==========================================================
# D48 = B48 - C48
# ==========================================================
ws["D48"] = round(
    (ws["B48"].value or 0) -
    (ws["C48"].value or 0),
    2
)

# ==========================================================
# VC -> C25
# Column A = Code
# Column C = Value
# ==========================================================
vc_value = 0

for row_num in range(2, ws2.max_row + 1):

    code = str(
        ws2[f"A{row_num}"].value or ""
    ).strip().upper()

    if code == "VC":

        vc_value += float(
            ws2[f"C{row_num}"].value or 0
        )

ws["C25"] = round(vc_value, 2)

# ==========================================================
# D25 = B25 - C25
# ==========================================================
ws["D25"] = round(
    (ws["B25"].value or 0) -
    (ws["C25"].value or 0),
    2
)
# ==========================================================
# SERVICE YEARS
#
# Calc Template
# B2 = Leaving Date
# B3 = Hiring Date
#
# Holiday Table
# B31 = Completed Years of Service
# ==========================================================

from datetime import datetime

ws_holiday = wb["Holiday Table"]

hiring_date = ws["B3"].value
leaving_date = ws["B2"].value

service_years = 0

if hiring_date and leaving_date:

    service_years = leaving_date.year - hiring_date.year

    if (
        leaving_date.month,
        leaving_date.day
    ) < (
        hiring_date.month,
        hiring_date.day
    ):
        service_years -= 1

ws_holiday["B31"] = service_years

print("Hiring Date  =", hiring_date)
print("Leaving Date =", leaving_date)
print("Service Years =", service_years)
# ==========================================================
# HOLIDAY CALCULATION
# ==========================================================

holiday_ws = wb["Holiday Table"]

lwd = ws["B2"].value

from datetime import datetime

from datetime import datetime

holiday_ws = wb["Holiday Table"]

lwd = ws["B2"].value

holiday_ws["L2"] = 0

closest_diff = None
closest_value = 0

for row_num in range(2, holiday_ws.max_row + 1):

    # Check Column A -> Return Column B
    date_a = holiday_ws[f"A{row_num}"].value

    if date_a:
        try:
            diff = abs((lwd - date_a).days)

            if closest_diff is None or diff < closest_diff:
                closest_diff = diff
                closest_value = holiday_ws[f"B{row_num}"].value or 0

        except:
            pass

    # Check Column D -> Return Column E
    date_d = holiday_ws[f"D{row_num}"].value

    if date_d:
        try:
            diff = abs((lwd - date_d).days)

            if closest_diff is None or diff < closest_diff:
                closest_diff = diff
                closest_value = holiday_ws[f"E{row_num}"].value or 0

        except:
            pass

holiday_ws["L2"] = closest_value

print("LWD =", lwd)
print("L2 =", holiday_ws["L2"].value)

# ==========================================================
# L3 (LONG SERVICE)
#
# Holiday Table!B31 = Years of Service
# L3 = B31 / 5
# Whole number only (no decimals)
# ==========================================================

service_years = float(holiday_ws["B31"].value or 0)

holiday_ws["L3"] = int(service_years / 5)

print("L3 =", holiday_ws["L3"].value)

# ==========================================================
# L6 = C36
# Holiday Table
# ==========================================================
holiday_ws["L6"] = float(
    holiday_ws["C36"].value or 0
)

# ==========================================================
# L7 = -(F36)
# Holiday Table
# ==========================================================
holiday_ws["L7"] = -float(
    holiday_ws["F36"].value or 0
)

print("L6 =", holiday_ws["L6"].value)
print("L7 =", holiday_ws["L7"].value)

# ==========================================================
# TOTAL HOLIDAY DAYS
# L2:L9
# ==========================================================
holiday_days = 0

for cell_name in ["L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9"]:
    holiday_days += float(
        holiday_ws[cell_name].value or 0
    )

# ==========================================================
# ENTITY
# Pay Group No. 0001 NDE
# Last word = Entity
# ==========================================================
entity = ""

for row_num in range(1, ws2.max_row + 1):

    text_value = str(
        ws2[f"A{row_num}"].value or ""
    ).strip()

    if "PAY GROUP NO." in text_value.upper():

        entity = text_value.split()[-1].upper()
        break

entity_hours = 0

if entity in ["NMUK", "NTCE", "NITCO", "NMPC"]:
    entity_hours = 7.8

elif entity in ["NMGB", "NDE"]:
    entity_hours = 7.5

# ==========================================================
# HOURLY RATE
# Find CNTRCTHR
# Next row contains:
# (169.00 X 14.6493)
# Hourly Rate = 14.6493
# ==========================================================
hourly_rate = 0

for row_num in range(1, ws2.max_row):

    text_value = str(
        ws2[f"A{row_num}"].value or ""
    )

    if "CNTRCTHR" in text_value.upper():

        next_value = str(
            ws2[f"A{row_num + 1}"].value or ""
        )

        next_value = next_value.replace("(", "")
        next_value = next_value.replace(")", "")

        if "X" in next_value:

            parts = next_value.split("X")

            if len(parts) == 2:

                hourly_rate = float(
                    parts[1].strip()
                )

        break

# ==========================================================
# SHIFT PREMIUM
# ==========================================================
shift_percent = 0

for row_num in range(1, ws2.max_row + 1):

    shift_text = str(
        ws2[f"A{row_num}"].value or ""
    ).strip().upper()

    if "SHIFT DF/DT" in shift_text:
        shift_percent = 0.22
        break

    elif "SHIFT DF" in shift_text:
        shift_percent = 0.22
        break

    elif "SHIFT DL" in shift_text:
        shift_percent = 0.12
        break

    elif "SHIFT MF/RT" in shift_text:
        shift_percent = 0.38
        break

    elif "SHIFT WE A" in shift_text:
        shift_percent = 0.30
        break

    elif "SHIFT WE B" in shift_text:
        shift_percent = 0.10
        break

    elif "SHIFT WEC" in shift_text:
        shift_percent = 0.20
        break

    elif "SHIFT WE" in shift_text:
        shift_percent = 0.225
        break

    elif "SHIFT ES" in shift_text:
        shift_percent = 0.1333
        break

    elif "SHIFT MN/SP" in shift_text:
        shift_percent = 0.185
        break

    elif "SHIFT TS" in shift_text:
        shift_percent = 0.27
        break


# ==========================================================
# BALANCE TO PAY
# ==========================================================

if shift_code:

    balance_to_pay = (
        holiday_days *
        entity_hours *
        hourly_rate *
        shift_percent
    )

else:

    balance_to_pay = (
        holiday_days *
        entity_hours *
        hourly_rate
    )

holiday_ws["L10"] = round(balance_to_pay, 2)

# ==========================================================
# COPY TO CALC TEMPLATE
# ==========================================================
ws["B30"] = round(balance_to_pay, 2)

print("L2 =", holiday_ws["L2"].value)
print("L3 =", holiday_ws["L3"].value)
print("L4 =", holiday_ws["L4"].value)
print("L10 =", holiday_ws["L10"].value)
print("B30 =", ws["B30"].value)
# ==========================================================
# D30 = B30 - C30
# ==========================================================

ws["D30"] = round(
    (ws["B30"].value or 0) -
    (ws["C30"].value or 0),
    2
)

print("D30 =", ws["D30"].value)

# ==========================================================
# HOLEXTRA -> C24
# Column A = Code
# Column C = Value
# Source Sheet = Excelpayslip (ws2)
# ==========================================================

holextra_value = 0

for row_num in range(2, ws2.max_row + 1):

    code = str(
        ws2[f"A{row_num}"].value or ""
    ).strip().upper()

    if code == "HOLEXTRA":

        holextra_value += float(
            ws2[f"C{row_num}"].value or 0
        )

ws["C24"] = round(holextra_value, 2)

# ==========================================================
# HOURLY RATE
# Find CNTRCTHR
# Next row contains:
# (169.00 X 14.6493)
# Hourly Rate = 14.6493
# ==========================================================

hourly_rate = 0

for row_num in range(1, ws2.max_row):

    text_value = str(
        ws2[f"A{row_num}"].value or ""
    )

    if "CNTRCTHR" in text_value.upper():

        next_value = str(
            ws2[f"A{row_num + 1}"].value or ""
        )

        next_value = next_value.replace("(", "")
        next_value = next_value.replace(")", "")

        if "X" in next_value:

            parts = next_value.split("X")

            if len(parts) == 2:

                hourly_rate = float(
                    parts[1].strip()
                )

        break

# ==========================================================
# B24 = C24 * 1.5 * Hourly Rate
# ==========================================================

ws["B24"] = round(
    (ws["C24"].value or 0) *
    1.5 *
    hourly_rate,
    2
)

# ==========================================================
# D24 = B24 - C24
# ==========================================================

ws["D24"] = round(
    (ws["B24"].value or 0) -
    (ws["C24"].value or 0),
    2
)

# ------------------------------------------------
# HOL BAL2 -> C83
# Read code from Column G
# Read value from Column I
# Convert value to number
# ------------------------------------------------

ws["C83"] = 0

for row in range(1, ws2.max_row + 1):

    code = str(ws2.cell(row=row, column=7).value or "").strip().upper()  # Column G

    if code == "HOL BAL2":

        raw_value = str(ws2.cell(row=row, column=9).value or "0").strip()  # Column I

        try:
            ws["C83"] = float(raw_value.replace(",", "").rstrip("C"))
        except:
            ws["C83"] = 0

        break

print("C83 =", ws["C83"].value)

# ------------------------------------------------
# Holiday Table O3:O14 Total -> O15
# ------------------------------------------------

holiday_ws = wb["Holiday Table"]

holiday_ws["O15"] = sum(
    float(holiday_ws[f"O{row}"].value or 0)
    for row in range(3, 15)
)

# ------------------------------------------------
# Determine Pay Group Divisor
# ------------------------------------------------

if pay_group in ["NMUK", "NTCE", "NITCO", "NMPC"]:
    divisor = 2028
elif pay_group in ["NMGB", "NDE"]:
    divisor = 1950
else:
    divisor = 0

# ------------------------------------------------
# B31 = (O15 / Divisor) * Rate
# ------------------------------------------------

if divisor and rate:
    ws["B31"] = (
    (float(holiday_ws["O15"].value or 0) / divisor)
    * rate
    * float(holiday_ws["R2"].value or 0)
)
else:
    ws["B31"] = 0

print("Pay Group =", pay_group)
print("Rate =", rate)
print("O15 =", holiday_ws["O15"].value)
print("B31 =", ws["B31"].value)
# D31 = B31 - C31
ws["D31"] = (float(ws["B31"].value or 0)) - (float(ws["C31"].value or 0))

# ------------------------------------------------
# Sum B5:B42 and store in B43
# ------------------------------------------------

ws["B43"] = sum(
    float(ws[f"B{row}"].value or 0)
    for row in range(5, 43)
)

# ------------------------------------------------
# Sum C5:C42 and store in C43
# ------------------------------------------------

ws["C43"] = sum(
    float(ws[f"C{row}"].value or 0)
    for row in range(5, 43)
)

# ------------------------------------------------
# Sum D5:D42 and store in D43
# ------------------------------------------------

ws["D43"] = sum(
    float(ws[f"D{row}"].value or 0)
    for row in range(5, 43)
)
# ------------------------------------------------
# Sum B44:B74 and store in B75
# ------------------------------------------------

ws["B75"] = sum(
    float(ws[f"B{row}"].value or 0)
    for row in range(44, 75)
)

# ------------------------------------------------
# Sum C44:C74 and store in C75
# ------------------------------------------------

ws["C75"] = sum(
    float(ws[f"C{row}"].value or 0)
    for row in range(44, 75)
)

# ------------------------------------------------
# Sum D44:D74 and store in D75
# ------------------------------------------------

ws["D75"] = sum(
    float(ws[f"D{row}"].value or 0)
    for row in range(44, 75)
)
# B76 + B77 -> B78
ws["B78"] = (ws["B76"].value or 0) + (ws["B77"].value or 0)

# C76 + C77 -> C78
ws["C78"] = (ws["C76"].value or 0) + (ws["C77"].value or 0)

# D76 + D77 -> D78
ws["D78"] = (ws["D76"].value or 0) + (ws["D77"].value or 0)

# ==========================================================
# STUDENT LOAN SHEET
# ==========================================================
ws5 = wb["STUDENT LOAN"]

# ==========================================================
# TOTAL GROSS -> B43
# ==========================================================
total_gross = 0

for row_num in range(2, 43):
    value = ws[f"B{row_num}"].value

    if isinstance(value, (int, float)):
        total_gross += value

ws["B43"] = round(total_gross, 2)

# ==========================================================
# ACTUAL STD LOAN FROM EXCELPAYSLIP -> C49
# ==========================================================
actual_std_loan = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_name = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if deduction_name == "STD LOAN":

        actual_std_loan = float(
            ws2[f"F{row_num}"].value or 0
        )

        break

ws["C49"] = round(actual_std_loan, 2)

# ==========================================================
# PLAN DETAILS
# ==========================================================
loan_status = str(ws5["B2"].value or "").strip().upper()
loan_plan = str(ws5["B3"].value or "").strip().upper()

threshold_value = 0
percent_value = 0

if loan_plan == "L1":
    threshold_value = float(ws5["C6"].value or 0)
    percent_value = float(ws5["D6"].value or 0)

elif loan_plan == "L2":
    threshold_value = float(ws5["C7"].value or 0)
    percent_value = float(ws5["D7"].value or 0)

elif loan_plan == "L3":
    threshold_value = float(ws5["C8"].value or 0)
    percent_value = float(ws5["D8"].value or 0)

elif loan_plan == "L4":
    threshold_value = float(ws5["C9"].value or 0)
    percent_value = float(ws5["D9"].value or 0)

# If stored as 9 instead of 0.09
if percent_value > 1:
    percent_value = percent_value / 100

# ==========================================================
# CALCULATED STD LOAN -> B49
# ==========================================================
calculated_std_loan = 0

if loan_status == "ACTIVE":

    if total_gross > threshold_value:

        calculated_std_loan = (
            total_gross - threshold_value
        ) * percent_value

ws["B49"] = int(calculated_std_loan)
ws["D49"] = round(
    float(ws["B49"].value or 0) -
    float(ws["C49"].value or 0),
    2
)

print("D49 =", ws["D49"].value)

# ==========================================================
# PG LOAN ACTUAL VALUE -> C55
# Search Excelpayslip Column D = PG LOAN
# Take value from Column F
# ==========================================================
actual_pg_loan = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_name = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if "PG LOAN" in deduction_name:

        actual_pg_loan = float(
            ws2[f"F{row_num}"].value or 0
        )

        break

ws["C55"] = round(actual_pg_loan, 2)

# ==========================================================
# PG LOAN CALCULATION -> B55
# ==========================================================
calculated_pg_loan = 0

loan_status = str(
    ws5["F2"].value or ""
).strip().upper()

threshold_value = float(
    ws5["C10"].value or 0
)

percent_value = float(
    ws5["D10"].value or 0
)

# If percentage stored as 6 instead of 0.06
if percent_value > 1:
    percent_value = percent_value / 100

print("Total Gross:", total_gross)
print("Status:", loan_status)
print("Threshold:", threshold_value)
print("Percent:", percent_value)

if loan_status == "ACTIVE":

    if total_gross > threshold_value:

        calculated_pg_loan = (
            total_gross - threshold_value
        ) * percent_value

    else:

        calculated_pg_loan = 0

else:

    calculated_pg_loan = 0

# B55 should not contain decimals
ws["B55"] = round(calculated_pg_loan)

# ==========================================================
# DIFFERENCE -> D55
# ==========================================================
ws["D55"] = round(
    (ws["B55"].value or 0)
    - (ws["C55"].value or 0),
    2
)

print("B55 =", ws["B55"].value)
print("C55 =", ws["C55"].value)
print("D55 =", ws["D55"].value)

# ==========================================================
# COURT ORDER -> C46
# Codes: JDEBT J1 to JDEBT J6
# Column D = Code
# Column F = Value
# ==========================================================
court_order_value = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_code = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if (
        deduction_code == "JDEBT J1"
        or deduction_code == "JDEBT J2"
        or deduction_code == "JDEBT J3"
        or deduction_code == "JDEBT J4"
        or deduction_code == "JDEBT J5"
        or deduction_code == "JDEBT J6"
    ):

        court_order_value += float(
            ws2[f"F{row_num}"].value or 0
        )

ws["C46"] = round(court_order_value, 2)

# ==========================================================
# D46 = B46 - C46
# ==========================================================
ws["D46"] = round(
    (ws["B46"].value or 0) -
    (ws["C46"].value or 0),
    2
)

# ==========================================================
# GOODS2 / GOODS CALCULATION
#
# Dec 2025 → Nov 2026 = GOODS2 → B60/C60/D60
# Dec 2026 → Nov 2027 = GOODS  → B61/C61/D61
# Dec 2027 → Nov 2028 = GOODS2 → B60/C60/D60
# Dec 2028 → Nov 2029 = GOODS  → B61/C61/D61
#
# Formula:
# 1000 / 12 * Remaining Months in Cycle
#
# Date is in Calc Template B2
# ==========================================================

leave_date = ws["B2"].value

ws["B60"] = 0
ws["B61"] = 0

if leave_date:

    scheme_start_year = leave_date.year

    if leave_date.month < 12:
        scheme_start_year -= 1

    remaining_months = 12 - leave_date.month

    if leave_date.month == 12:
        remaining_months = 12

    goods_amount = round(
        (1000 / 12) * remaining_months,
        2
    )

    # GOODS2 Cycle
    if scheme_start_year % 2 == 1:
        ws["B60"] = goods_amount

    # GOODS Cycle
    else:
        ws["B61"] = goods_amount

# ==========================================================
# GOODS2 -> C60
# Read from Excelpayslip
# Col D = GOODS2
# Col F = Value
# ==========================================================

goods2_value = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_name = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if deduction_name == "GOODS2":

        goods2_value = float(
            ws2[f"F{row_num}"].value or 0
        )

        break

ws["C60"] = round(goods2_value, 2)

# ==========================================================
# GOODS -> C61
# Read from Excelpayslip
# Col D = GOODS
# Col F = Value
# ==========================================================

goods_value = 0

for row_num in range(2, ws2.max_row + 1):

    deduction_name = str(
        ws2[f"D{row_num}"].value or ""
    ).strip().upper()

    if deduction_name == "GOODS":

        goods_value = float(
            ws2[f"F{row_num}"].value or 0
        )

        break

ws["C61"] = round(goods_value, 2)

# ==========================================================
# DIFFERENCES
# ==========================================================

ws["D60"] = round(
    (ws["B60"].value or 0) -
    (ws["C60"].value or 0),
    2
)

ws["D61"] = round(
    (ws["B61"].value or 0) -
    (ws["C61"].value or 0),
    2
)
# ------------------------------------------------
# Sum B5:B42 and store in B43
# ------------------------------------------------

ws["B43"] = sum(
    float(ws[f"B{row}"].value or 0)
    for row in range(5, 43)
)

# ------------------------------------------------
# Sum C5:C42 and store in C43
# ------------------------------------------------

ws["C43"] = sum(
    float(ws[f"C{row}"].value or 0)
    for row in range(5, 43)
)

# ------------------------------------------------
# Sum D5:D42 and store in D43
# ------------------------------------------------

ws["D43"] = sum(
    float(ws[f"D{row}"].value or 0)
    for row in range(5, 43)
)
# ------------------------------------------------
# Sum B44:B74 and store in B75
# ------------------------------------------------

ws["B75"] = sum(
    float(ws[f"B{row}"].value or 0)
    for row in range(44, 75)
)

# ------------------------------------------------
# Sum C44:C74 and store in C75
# ------------------------------------------------

ws["C75"] = sum(
    float(ws[f"C{row}"].value or 0)
    for row in range(44, 75)
)

# ------------------------------------------------
# Sum D44:D74 and store in D75
# ------------------------------------------------

ws["D75"] = sum(
    float(ws[f"D{row}"].value or 0)
    for row in range(44, 75)
)
# B76 + B77 -> B78
ws["B78"] = (ws["B76"].value or 0) + (ws["B77"].value or 0)

# C76 + C77 -> C78
ws["C78"] = (ws["C76"].value or 0) + (ws["C77"].value or 0)

# D76 + D77 -> D78
ws["D78"] = (ws["D76"].value or 0) + (ws["D77"].value or 0)

# --------------------------------------------
# D Column Variance Calculations
# --------------------------------------------

ws["D44"] = round((float(ws["B44"].value or 0) - float(ws["C44"].value or 0)), 2)
ws["D45"] = round((float(ws["B45"].value or 0) - float(ws["C45"].value or 0)), 2)

ws["D50"] = round((float(ws["B50"].value or 0) - float(ws["C50"].value or 0)), 2)
ws["D51"] = round((float(ws["B51"].value or 0) - float(ws["C51"].value or 0)), 2)
ws["D52"] = round((float(ws["B52"].value or 0) - float(ws["C52"].value or 0)), 2)

ws["D59"] = round((float(ws["B59"].value or 0) - float(ws["C59"].value or 0)), 2)

ws["D64"] = round((float(ws["B64"].value or 0) - float(ws["C64"].value or 0)), 2)
ws["D65"] = round((float(ws["B65"].value or 0) - float(ws["C65"].value or 0)), 2)
ws["D66"] = round((float(ws["B66"].value or 0) - float(ws["C66"].value or 0)), 2)

ws["D71"] = round((float(ws["B71"].value or 0) - float(ws["C71"].value or 0)), 2)
ws["D72"] = round((float(ws["B72"].value or 0) - float(ws["C72"].value or 0)), 2)
ws["D74"] = round((float(ws["B74"].value or 0) - float(ws["C74"].value or 0)), 2)

ws["D80"] = round((float(ws["B80"].value or 0) - float(ws["C80"].value or 0)), 2)
ws["D81"] = round((float(ws["B81"].value or 0) - float(ws["C81"].value or 0)), 2)
ws["D82"] = round((float(ws["B82"].value or 0) - float(ws["C82"].value or 0)), 2)
ws["D83"] = round((float(ws["B83"].value or 0) - float(ws["C83"].value or 0)), 2)

ws["D85"] = round((float(ws["B85"].value or 0) - float(ws["C85"].value or 0)), 2)
ws["D86"] = round((float(ws["B86"].value or 0) - float(ws["C86"].value or 0)), 2)

# --------------------------------------------
# COPY NI!B8 TO Calc Template!C82
# --------------------------------------------

try:
    c82_value = round(float(wb["NI"]["B8"].value or 0), 2)
except:
    c82_value = 0

ws["C82"] = c82_value

print("NI!B8 =", c82_value)
print("Calc Template!C82 =", ws["C82"].value)
# -------------------------------
# SAFE SAVE
# -------------------------------
output_file = file1.replace(".xlsx", "_processed.xlsx")

wb.save(output_file)
print("\n✅ Processing Completed Successfully")
print("Successfully copied all values")
