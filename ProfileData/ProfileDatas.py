from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import json
import time 
from dotenv import load_dotenv
import os
load_dotenv()

user = os.getenv("USER")
password = os.getenv("PASSWORD")
# Carrega o JSON
with open(r"profile.json", "r", encoding="utf-8") as f:
    profiles_data = json.load(f)

# Pega os nomes dos profiles
profiles_to_register = [p['profile']['profileName'] for p in profiles_data]

# Pega a lista de chaves (os nomes dos datas) correspondente a cada profile
data_register = [list(d['data'].keys()) for d in profiles_data]

options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()
driver.get("https://par1.c2x.com.br/login.aspx")

# Login
driver.find_element(By.ID, "vUSERNAME").send_keys(user)
time.sleep(1)
driver.find_element(By.ID, "vUSERPASSWORD").send_keys(password)

time.sleep(2)
driver.execute_script("document.body.style.zoom='50%'")
time.sleep(1)

wait = WebDriverWait(driver, 10)
wait.until(EC.element_to_be_clickable((By.ID, "BTNENTER"))).click()

time.sleep(2)
driver.execute_script("document.body.style.zoom='30%'")
time.sleep(1)

# Entra no módulo
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0004"))).click()
time.sleep(3) 

# --- SALVA A URL DA "HOME" DO MÓDULO ---
url_home_do_modulo = driver.current_url
print(f"URL Principal salva: {url_home_do_modulo}")
# -------------------------------------------------------------

# --- LOOP PRINCIPAL ---
for profile_name, lista_datas_deste_profile in zip(profiles_to_register, data_register):
    try:
        # --- RESET INTELIGENTE ---
        print(f"--- Iniciando Profile: {profile_name} ---")
        
        if driver.current_url != url_home_do_modulo:
            driver.get(url_home_do_modulo)
            time.sleep(2) 

        # Pega botão adicionar
        botao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@id='UNNAMEDTABLE2']//a[i[contains(@class,'BotaoAdicionarIconeDark')]]")
            )
        )
        url = botao.get_attribute("href")
        print(f"URL encontrada (Criar Profile): {url}")

        # --- CRIAÇÃO DO PROFILE ---
        driver.get(url)
        time.sleep(1)
        driver.execute_script("document.body.style.zoom='30%'")
        time.sleep(1)
        time.sleep(1)
        # Mudei aqui para visibility também, garante que o campo apareceu
        wait.until(EC.visibility_of_element_located((By.ID, "APLICACAONOME"))).send_keys(profile_name)
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.ID, "BTNTRN_ENTER"))).click()

        # Busca o Profile criado na grid e entra nele
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "GridBlocos"))
        )
        
        try:
            profile_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='{profile_name}']")
                )
            )
            print(f"Profile encontrado e acessando: {profile_name}")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", profile_element)
            profile_element.click()

            time.sleep(2)
            
            # --- CRIAÇÃO DO AMBIENTE PRD ---
            botao = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@id='UNNAMEDTABLE2']//a[i[contains(@class,'BotaoAdicionarIconeDark')]]")
                )
            )
            url = botao.get_attribute("href")
            
            driver.get(url)
            time.sleep(1)
            driver.execute_script("document.body.style.zoom='30%'")
            time.sleep(1)
            # Visibility garante que o campo está visível
            wait.until(EC.visibility_of_element_located((By.ID, "AMBIENTENOME"))).send_keys("PRD") 
            time.sleep(1)   
            wait.until(EC.element_to_be_clickable((By.ID, "BTNTRN_ENTER"))).click()
            
            # Busca o Ambiente PRD e entra nele
            AMBIENTE = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='PRD']")
                )
            )
            AMBIENTE.click()
            time.sleep(2)

            # --- CADASTRO DOS DATAS ---
            
            botao_data = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@id='UNNAMEDTABLE2']//a[i[contains(@class,'BotaoAdicionarIconeDark')]]")
                )
            )
            url_cadastro_data = botao_data.get_attribute("href")
            print(f"URL de cadastro de datas capturada: {url_cadastro_data}")

            # Loop para cadastrar datas
            for nome_data in lista_datas_deste_profile:
                print(f"Cadastrando Data: {nome_data}")
                
                driver.get(url_cadastro_data)
                time.sleep(2)
                driver.execute_script("document.body.style.zoom='30%'")
                time.sleep(1)
                campo_nome = wait.until(EC.visibility_of_element_located((By.ID, "CONJUNTODADOSNOME")))
                time.sleep(0.5)
                campo_nome.click()
                time.sleep(0.5) 
                campo_nome.clear() 
                time.sleep(0.5)
                campo_nome.send_keys(nome_data)
                
                time.sleep(1) 
                
                wait.until(EC.element_to_be_clickable((By.ID, "BTNTRN_ENTER"))).click()
                
                time.sleep(1)
            
            print(f"Todos os datas do profile {profile_name} foram cadastrados!")

        except Exception as e:
            print(f"Erro ao processar profile {profile_name}: {e}")

    except Exception as e:
        print("Erro geral no loop:", e)

# Volta para a tela de login ao final de tudo
print("Automação finalizada. Voltando para a tela de login...")
driver.get("https://par1.c2x.com.br/login.aspx")