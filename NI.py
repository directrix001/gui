import time
import base64
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import shutil
import re

# -------------------------------
# SELECT EXCEL FILE
# -------------------------------
def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel Files", "*.xlsx")]
    )

excel_file = select_excel_file()

if not excel_file:
    print("No file selected.")
    exit()

print("Excel file selected:", excel_file)

# -------------------------------
# BACKUP FILE
# -------------------------------
backup_file = excel_file.replace(".xlsx", "_backup.xlsx")
shutil.copy(excel_file, backup_file)
print("Backup created:", backup_file)

# -------------------------------
# LOAD WORKBOOK
# -------------------------------
wb = load_workbook(excel_file)
ws_pay = wb["Excelpayslip"]
ws_calc = wb["Calc Template"]

# -------------------------------
# SELENIUM SETUP
# -------------------------------
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
wait = WebDriverWait(driver, 15)

# -------------------------------
# OPEN HMRC NI PAGE
# -------------------------------
url = "https://www.tax.service.gov.uk/guidance/work-out-employer-employee-national-insurance-contributions/input/employee-category-letter?p=1"
driver.get(url)

# -------------------------------
# ACCEPT COOKIES
# -------------------------------
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button"))).click()
except:
    pass

# -------------------------------
# SELECT MONTHLY
# -------------------------------
wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//label[contains(.,'Monthly')]"))).click()

wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//button[contains(.,'Continue')]"))).click()

# -------------------------------
# GET GROSS PAY FROM CALC TEMPLATE
# -------------------------------
ws_calc = wb["Calc Template"]

try:
    gross_pay = round(float(ws_calc["B43"].value or 0), 2)
except (ValueError, TypeError):
    gross_pay = 0.00

print(f"Gross Pay = {gross_pay:.2f}")

# ✅ FIXED INPUT
gross_input = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//input[@type='text']")))

driver.execute_script("arguments[0].scrollIntoView(true);", gross_input)
time.sleep(1)

gross_input.clear()
gross_input.send_keys(gross_pay)

wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//button[contains(.,'Continue')]"))).click()

# -------------------------------
# NI CATEGORY
# -------------------------------
ni_category = "A"
for row in ws_pay.iter_rows():
    if row[3].value and "contribution letter" in str(row[3].value).lower():
        ni_category = str(row[3].value)[-1]

print("NI Category:", ni_category)

cat_input = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//input[@type='text']")))

cat_input.clear()
cat_input.send_keys(ni_category)

wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//button[contains(.,'Continue')]"))).click()

# -------------------------------
# CONTINUE PAGE
# -------------------------------
wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//a[contains(.,'Continue')]"))).click()

time.sleep(2)

# -------------------------------
# EXTRACT VALUES
# -------------------------------
page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
lines = page_text.split("\n")

emp = employer = lel = pt = uel = "0"

for line in lines:
    clean = line.replace("’", "").strip()

    # ✅ Employee NIC
    if "employees nics due" in clean:
        match = re.search(r"[\d,]+\.\d{2}", line)
        if match:
            emp = match.group()

    # ✅ Employer NIC
    elif "employers nics due" in clean:
        match = re.search(r"[\d,]+\.\d{2}", line)
        if match:
            employer = match.group()

    # ✅ Lower Earnings Limit (STRICT match)
    elif clean.startswith("lower earnings limit"):
        match = re.search(r"[\d,]+\.\d{2}", line)
        if match:
            lel = match.group()

    # ✅ Primary Threshold (FIXED)
    elif "earnings above the lower earnings limit" in clean:
        match = re.search(r"[\d,]+\.\d{2}", line)
        if match:
            pt = match.group()

    # ✅ Upper Earnings Limit (FIXED)
    elif "earnings above the primary threshold" in clean:
        match = re.search(r"[\d,]+\.\d{2}", line)
        if match:
            uel = match.group()

# -------------------------------
# CLEAN + CONVERT
# -------------------------------
def val(x):
    try:
        return float(x.replace(",", ""))
    except:
        return 0.0

emp_n = val(emp)
employer_n = val(employer)
lel_n = val(lel)
pt_n = val(pt)
uel_n = val(uel)

print("Extracted Values:")
print("Employee:", emp_n)
print("Employer:", employer_n)
print("LEL:", lel_n)
print("PT:", pt_n)
print("UEL:", uel_n)

# -------------------------------
# WRITE TO EXCEL (NO SAVE HERE)
# -------------------------------
ws_calc["B45"] = emp_n
ws_calc["B85"] = emp_n
ws_calc["B86"] = employer_n

ws_calc["B80"] = lel_n
ws_calc["B81"] = pt_n
ws_calc["B82"] = uel_n

# -------------------------------
# FINAL SAVE (ONLY ONCE ✅)
# -------------------------------
output_file = excel_file.replace(".xlsx", "_updated.xlsx")
wb.save(output_file)

print("✅ Excel saved:", output_file)

# -------------------------------
# SAVE PDF
# -------------------------------
pdf = driver.execute_cdp_cmd("Page.printToPDF", {
    "printBackground": True
})


with open("HMRC_NI_Calculation.pdf", "wb") as f:
    f.write(base64.b64decode(pdf['data']))

print("✅ PDF saved")

driver.quit()