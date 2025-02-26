import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time

# Load Excel Data
try:
    df = pd.read_excel("data.xlsx", dtype={"P2 CONTACT": str, "P2 FORM SUBMITTED CONTACT": str})
    print("✅ Excel file loaded successfully.")
except Exception as e:
    print(f"❌ Error loading Excel file: {e}")
    exit()

# Initialize WebDriver
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
driver.get("https://gd92600ab354c71-db1c7nt.adb.me-dubai-1.oraclecloudapps.com/ords/r/health/diabetes-camp/survey")

wait = WebDriverWait(driver, 10)

# Function to set value using JavaScript
def set_value_js(element_id, value):
    if pd.notna(value) and value != "":
        script = "document.getElementById(arguments[0]).value = arguments[1];"
        driver.execute_script(script, element_id, value)
    else:
        print(f"⚠ Skipping {element_id} as it is empty or NaN.")

# Function to select dropdown values
def select_dropdown(element_id, value):
    if pd.notna(value) and value != "":
        try:
            dropdown = Select(driver.find_element(By.ID, element_id))
            dropdown.select_by_visible_text(str(value))
        except Exception:
            print(f"⚠ Warning: Could not select {value} in {element_id}, trying JavaScript...")
            set_value_js(element_id, value)
    else:
        print(f"⚠ Skipping {element_id} as it is empty or NaN.")

# Function to search, type, and select value in P2_CODE
def search_and_select_code(dropdown_id, search_input_xpath, search_button_xpath, value):
    if pd.notna(value) and value != "":
        try:
            dropdown = wait.until(EC.element_to_be_clickable((By.ID, dropdown_id)))
            dropdown.click()
            time.sleep(2)

            search_input = wait.until(EC.element_to_be_clickable((By.XPATH, search_input_xpath)))
            search_input.clear()
            search_input.send_keys(value)
            time.sleep(2)

            search_button = wait.until(EC.element_to_be_clickable((By.XPATH, search_button_xpath)))
            search_button.click()
            time.sleep(2)

            suggestion = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/div[3]/ul/li")))
            suggestion.click()
            print(f"✅ Selected: {value}")

        except Exception as e:
            print(f"❌ Error selecting {value} in {dropdown_id}: {e}")
    else:
        print(f"⚠ Skipping {dropdown_id} as it is empty or NaN.")

# Function to click button if P2_CODE has a value
def click_button_if_value_exists(element_id, button_id):
    try:
        value = driver.execute_script(f"return document.getElementById('{element_id}').value;")
        if value.strip():
            print(f"✅ '{element_id}' has a value: {value}. Clicking button '{button_id}'...")
            button = wait.until(EC.element_to_be_clickable((By.ID, button_id)))
            button.click()
        else:
            print(f"⚠ '{element_id}' is empty. Skipping button click.")
    except Exception as e:
        print(f"❌ Error clicking button '{button_id}': {e}")

# Loop through Excel rows and fill the form
for index, row in df.iterrows():
    try:
        search_and_select_code("P2_CODE", "//input[@class='a-PopupLOV-search apex-item-text']", "//button[@class='a-Button a-PopupLOV-doSearch']", row["P2 CODE CONTAINER"])
        
        set_value_js("P2_PATIENT_NAME", row["P2 PATIENT NAME"])
        set_value_js("P2_PATIENT_CNIC", row["P2 PATIENT CNIC"])

        p2_contact_value = str(row["P2 CONTACT"]).zfill(11) if pd.notna(row["P2 CONTACT"]) else ""
        set_value_js("P2_CONTACT", p2_contact_value)

        set_value_js("P2_AGE", row["P2 AGE"])
        set_value_js("P2_HEALTH_HABIT", row["P2 HEALTH HABIT"])
        set_value_js("P2_BODY_WEIGHT", row["P2 BODY WEIGHT"])
        set_value_js("P2_BP_SYSTOLIC", row["P2 BP SYSTOLIC"])
        set_value_js("P2_BP_DIASTOLIC", row["P2 BP DIASTOLIC"])
        set_value_js("P2_RBS", row["P2 RBS"])
        set_value_js("P2_FORM_SUBMITTED", row["P2 FORM SUBMITTED"])

        p2_form_submit_contact = str(row["P2 FORM SUBMITTED CONTACT"]).zfill(11) if pd.notna(row["P2 FORM SUBMITTED CONTACT"]) else ""
        set_value_js("P2_FORM_SUBMITTED_CONTACT", p2_form_submit_contact)

        select_dropdown("P2_GENDER", row["P2 GENDER"])
        select_dropdown("P2_MARITAL_STATUS", row["P2 MARITAL STATUS"])
        select_dropdown("P2_HISTORY", row["P2 HISTORY"])
        select_dropdown("P2_HEALTH_HABIT", row["P2 HEALTH HABIT"])

        # Click button if P2_CODE has a value
        time.sleep(2)
        click_button_if_value_exists("P2_CODE", "B7948174571986929")

        print(f"✅ Form filled for row {index + 1}. Proceeding to next...")
        time.sleep(3)

    except Exception as e:
        print(f"❌ Error on row {index + 1}: {e}")
        continue

print("✅ All records filled. The browser will remain open. Close it manually when done.")
input("🔹 Press Enter to close the browser when you're finished...")