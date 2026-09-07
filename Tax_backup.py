import time
import base64
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

# -------------------------------
# SELECT EXCEL FILE
# -------------------------------
def select_excel_file():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select the Excel file inside Tax Cal",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    return file_path

excel_file = select_excel_file()

if not excel_file:
    print("No file selected.")
    exit()

print("Excel file selected:", excel_file)

# -------------------------------
# LOAD WORKBOOK
# -------------------------------
wb = load_workbook(excel_file, data_only=True)

# -------------------------------
# READ TAX CODE FROM "Excelpayslip"
# -------------------------------
ws = wb["Excelpayslip"]

tax_code_value = None

for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    col_d = row[3].value
    col_e = row[4].value

    if col_d and "tax" in str(col_d).lower() and "code" in str(col_d).lower():
        tax_code_value = str(col_e).strip()
        break

print("Tax Code found:", tax_code_value)

# -------------------------------
# READ DOB FROM "Calc Template" B2
# -------------------------------
ws_date = wb["Calc Template"]
date_value = ws_date["B2"].value

day = date_value.day
month = date_value.month
year = date_value.year

print("DOB:", day, month, year)

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
wait = WebDriverWait(driver, 10)

# -------------------------------
# OPEN HMRC TAX CODE PAGE
# -------------------------------
url = "https://www.tax.service.gov.uk/guidance/paye-tax-calculator/input/tax-code"
driver.get(url)

# -------------------------------
# ACCEPT COOKIES (IF PRESENT)
# -------------------------------
try:
    accept_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Accept')]"))
    )
    accept_button.click()
    print("Cookies accepted.")
    time.sleep(1)
except:
    print("No cookie banner appeared.")

# -------------------------------
# ENTER TAX CODE
# -------------------------------
tax_code_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
)

tax_code_input.clear()
tax_code_input.send_keys(tax_code_value)

time.sleep(1)

continue_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
)
continue_btn.click()

print("Tax Code entered successfully.")

# -------------------------------
# ENTER DOB
# -------------------------------
dob_day = wait.until(
    EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
)
dob_month = driver.find_element(By.XPATH, "(//input[@type='text'])[2]")
dob_year = driver.find_element(By.XPATH, "(//input[@type='text'])[3]")

dob_day.send_keys(str(day))
dob_month.send_keys(str(month))
dob_year.send_keys(str(year))

time.sleep(1)

continue_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
)
continue_btn.click()

print("DOB entered successfully.")

# -------------------------------
# SELECT PAY FREQUENCY → MONTHLY
# -------------------------------
monthly_label = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(normalize-space(),'Monthly')]"))
)
monthly_label.click()

time.sleep(1)

continue_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
)
continue_btn.click()

print("Pay frequency selected: Monthly")

# -------------------------------
# SELECT "NO" FOR WEEK 1 / MONTH 1
# -------------------------------
no_label = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//label[contains(translate(normalize-space(),'CUMULATIVE','cumulative'),'cumulative') "
        "or contains(translate(normalize-space(),'NO','no'),'no')]"
    ))
)
no_label.click()

time.sleep(1)

continue_btn = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
)
continue_btn.click()

print("Selected: Week 1 / Month 1 basis = NO")

# -------------------------------
# ENTER MONTHLY PAY (Calc Template → B43)
# -------------------------------
ws_calc = wb["Calc Template"]

raw_monthly_pay = ws_calc["B43"].value

if raw_monthly_pay is None:
    monthly_pay = "0.00"
else:
    monthly_pay = f"{float(raw_monthly_pay):.2f}"

print("Monthly Pay being entered:", monthly_pay)

pay_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
)

pay_input.clear()
pay_input.send_keys(monthly_pay)

time.sleep(1)

continue_btn = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(text(),'Continue')]")
    )
)

continue_btn.click()

print("Entered monthly pay.")

# -------------------------------
# TOTAL GROSS PAY TO DATE (TAX!B2)
# -------------------------------
page_html = driver.page_source.lower()

if "total gross pay to date" in page_html or "total pay to date" in page_html:
    print("Detected: Total pay to date page")

    tax_ws = wb["TAX"]

    try:
        total_pay_to_date = str(float(tax_ws["B3"].value or 0))
    except (ValueError, TypeError):
        total_pay_to_date = "0"

    print("Total pay to date being entered:", total_pay_to_date)

    total_pay_input = wait.until(
        EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
    )
    total_pay_input.clear()
    total_pay_input.send_keys(total_pay_to_date)

    time.sleep(1)

    continue_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
    )
    continue_btn.click()

    print("Entered total pay to date.")

# -------------------------------
# TOTAL TAX DUE TO DATE (TAX!B2)
# -------------------------------
page_html = driver.page_source.lower()

if "total tax due to date" in page_html or "tax due to date" in page_html:
    print("Detected: Total tax due to date page")

    tax_ws = wb["TAX"]

    try:
        total_tax_due = float(tax_ws["B2"].value or 0)
    except (ValueError, TypeError):
        total_tax_due = 0

    print("Total tax due to date being entered:", total_tax_due)

    tax_due_input = wait.until(
        EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
    )
    tax_due_input.clear()
    tax_due_input.send_keys(str(total_tax_due))

    time.sleep(1)

    continue_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
    )
    continue_btn.click()

    print("Entered total tax due to date.")

# -------------------------------
# REGULATORY LIMIT → ALWAYS 0
# -------------------------------
page_html = driver.page_source.lower()

if "regulatory limit" in page_html or "not deducted" in page_html:
    print("Detected: Regulatory limit page")

    reg_limit_input = wait.until(
        EC.presence_of_element_located((By.XPATH, "(//input[@type='text'])[1]"))
    )
    reg_limit_input.clear()
    reg_limit_input.send_keys("0")

    time.sleep(1)

    continue_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Continue')]"))
    )
    continue_btn.click()

    print("Entered regulatory limit value: 0")

# ==========================================================
# CHECK YOUR ANSWERS PAGE
# ==========================================================

print("\nChecking for 'Check Your Answers' page...")

try:

    # Verify browser session is still alive
    try:
        _ = driver.current_window_handle
    except Exception:
        print("ERROR: Browser session has been closed.")
        raise

    # Give page time to load
    time.sleep(2)

    page_html = driver.page_source.lower()

    if "check your answers" not in page_html:
        print("'Check Your Answers' page not detected.")
        print("Current URL:", driver.current_url)
        print("Current Title:", driver.title)

    else:

        print("Detected: Check Your Answers page")

        continue_btn = None

        # Try button first
        try:
            continue_btn = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//button[contains(.,'Continue')]"
                    )
                )
            )

        except Exception:

            # Try link
            try:
                continue_btn = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//a[contains(.,'Continue')]"
                        )
                    )
                )

            except Exception:

                # Fallback
                continue_btn = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//*[contains(text(),'Continue')]"
                        )
                    )
                )

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            continue_btn
        )

        time.sleep(1)

        driver.execute_script(
            "arguments[0].click();",
            continue_btn
        )

        print("Continue clicked successfully.")

except Exception as e:

    print("\nERROR ON CHECK YOUR ANSWERS PAGE")
    print("Error:", str(e))

    try:
        print("Current URL:", driver.current_url)
        print("Page Title :", driver.title)

        with open(
            "check_answers_debug.html",
            "w",
            encoding="utf-8"
        ) as f:
            f.write(driver.page_source)

        print("HTML saved as check_answers_debug.html")

    except Exception:
        print("Browser session already closed.")

    raise
# -------------------------------
# FINAL PAGE → SAVE ENTIRE PAGE AS PDF
# -------------------------------
print("Saving final page as PDF...")

pdf_data = driver.execute_cdp_cmd(
    "Page.printToPDF",
    {
        "printBackground": True,
        "landscape": False
    }
)

pdf_bytes = base64.b64decode(pdf_data['data'])

with open("HMRC_Tax_Calculation.pdf", "wb") as f:
    f.write(pdf_bytes)

print("PDF saved: HMRC_Tax_Calculation.pdf")

# -------------------------------
# EXTRACT "Tax due at end of current period"
# -------------------------------
print("Extracting 'Tax due at end of current period'...")

try:
    rows = driver.find_elements(By.XPATH, "//dl/div")

    tax_due_value = None

    for row in rows:
        try:
            label = row.find_element(By.TAG_NAME, "dt").text.lower().strip()
            value = row.find_element(By.TAG_NAME, "dd").text.strip()

            if "tax due at end of current period" in label:
                tax_due_value = value
                break
        except:
            continue

    # Fallback if not found
    if tax_due_value is None or tax_due_value == "":
        print("WARNING: Tax due value not found, defaulting to 0")
        tax_due_value = "0"

    print("Extracted Tax Due (raw):", tax_due_value)

except Exception as e:
    print("Error extracting tax due:", str(e))
    tax_due_value = "0"

# -------------------------------
# CLEAN VALUE (remove £, commas)
# -------------------------------
def clean_currency(val):
    try:
        return val.replace("£", "").replace(",", "").strip()
    except:
        return "0"

tax_due_clean = clean_currency(tax_due_value)

print("Cleaned Tax Due:", tax_due_clean)

# -------------------------------
# WRITE VALUE TO EXCEL → B44
# -------------------------------
print("Writing Tax Due to Excel (Calc Template B44)...")

try:
    ws_calc = wb["Calc Template"]

    # Convert safely to float
    try:
        numeric_value = float(tax_due_clean)
    except:
        print("Invalid numeric format, defaulting to 0")
        numeric_value = 0

    ws_calc["B44"] = numeric_value

    wb.save(excel_file)

    print("✅ Successfully written Tax Due to B44:", numeric_value)

except Exception as e:
    print("Error writing to Excel:", str(e))

# -------------------------------
# OPTIONAL: SCREENSHOT (DEBUG / AUDIT)
# -------------------------------
try:
    driver.save_screenshot("tax_due_capture.png")
    print("Screenshot saved: tax_due_capture.png")
except:
    pass

