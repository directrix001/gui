import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32

# ==================================================
# SELECT EXCEL FILE
# ==================================================

root = tk.Tk()
root.withdraw()

excel_file = filedialog.askopenfilename(
    title="Select Excel Workbook",
    filetypes=[("Excel Files", "*.xlsx *.xlsm")]
)

if not excel_file:
    raise SystemExit("No Excel workbook selected.")

# ==================================================
# SELECT PDF FILE
# ==================================================

pdf_file = filedialog.askopenfilename(
    title="Select Payslip PDF",
    filetypes=[("PDF Files", "*.pdf")]
)

if not pdf_file:
    raise SystemExit("No PDF file selected.")

# ==================================================
# OPEN EXCEL
# ==================================================

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

try:

    wb = excel.Workbooks.Open(os.path.abspath(excel_file))

    # ==============================================
    # DELETE OLD QUERY
    # ==============================================

    try:
        wb.Queries.Item("PDFPayslip").Delete()
        print("Old query deleted.")
    except:
        pass

    # ==============================================
    # CREATE PDF QUERY
    # ==============================================

    pdf_path = pdf_file.replace("\\", "\\\\")

    formula = f'''
let
    Source = Pdf.Tables(
        File.Contents("{pdf_path}"),
        [Implementation="1.3"]
    ),
    Page001 = Source{{[Id="Page001"]}}[Data]
in
    Page001
'''

    wb.Queries.Add("PDFPayslip", formula)

    print("Query created.")

    # ==============================================
    # GET OR CREATE EXCELPAYSLIP SHEET
    # ==============================================

    try:
        ws = wb.Worksheets("Excelpayslip")
        print("Excelpayslip sheet found.")
    except:
        ws = wb.Worksheets.Add()
        ws.Name = "Excelpayslip"
        print("Excelpayslip sheet created.")

    # ==============================================
    # CLEAR OLD DATA
    # ==============================================

    ws.Cells.Clear()

    print("Old data cleared.")

    # ==============================================
    # ADD WORKBOOK CONNECTION
    # ==============================================

    connection_string = (
        "OLEDB;"
        "Provider=Microsoft.Mashup.OleDb.1;"
        "Data Source=$Workbook$;"
        "Location=PDFPayslip;"
    )

    qt = ws.QueryTables.Add(
        Connection=connection_string,
        Destination=ws.Range("A1")
    )

    qt.CommandText = "SELECT * FROM [PDFPayslip]"
    qt.BackgroundQuery = False

    print("Refreshing query...")

    qt.Refresh(False)

    excel.CalculateUntilAsyncQueriesDone()

    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()

    # ==============================================
    # SAVE
    # ==============================================

    wb.Save()

    print("Workbook saved successfully.")
    print("Rows Imported:", ws.UsedRange.Rows.Count)
    print("Columns Imported:", ws.UsedRange.Columns.Count)

    messagebox.showinfo(
        "Success",
        "Payslip imported successfully."
    )

except Exception as e:

    print("ERROR:", str(e))

    messagebox.showerror(
        "Error",
        str(e)
    )

finally:

    excel.DisplayAlerts = True