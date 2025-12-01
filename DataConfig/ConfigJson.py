import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import os
from dotenv import load_dotenv
import os
load_dotenv()

user = os.getenv("USER")
password = os.getenv("PASSWORD")
try:
    df_workflow = pd.read_excel(r"Dados/Cópia de Workflow_Profiles_11_13_v1.xlsx")
    df_data = pd.read_csv("Data - 2025-11-26T152141.453.csv")

except Exception as e:
    print(f"impossivel ler o csv/excel[{e}]")
    exit()


options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--no-sandbox")
options.add_argument("--remote-allow-origins=*")
options.add_argument("--ignore-certificate-errors")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()

print("Fazendo login...")
driver.get("https://par1.c2x.com.br/login.aspx")
wait = WebDriverWait(driver, 15)

driver.find_element(By.ID, "vUSERNAME").send_keys(user)
time.sleep(0.5)
driver.find_element(By.ID, "vUSERPASSWORD").send_keys(password)
wait.until(EC.element_to_be_clickable((By.ID, "BTNENTER"))).click()

time.sleep(2)
driver.execute_script("document.body.style.zoom='65%'")
time.sleep(1)