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
from dotenv import load_dotenv
import os
load_dotenv()

user = os.getenv("USER")
password = os.getenv("PASSWORD")

print("--- ROBÔ CONFIGURADOR (SEQUÊNCIA CORRETA) ---")

# 1. PREPARAÇÃO DOS DADOS
try:
    df_mapa = pd.read_excel(r"Dados/MAPA_CONFIGURACAO_TEXTO.xlsx")
    df_workflow = pd.read_excel(r"Dados/Cópia de Workflow_Profiles_11_13_v1.xlsx")
    df_data = pd.read_csv(r"Dados/Data - 2025-11-26T152141.453.csv")
    
    df_workflow['skill_pri'] = df_workflow['skill_pri'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_data['skill_no'] = df_data['skill_no'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    df_de_para = pd.merge(df_workflow, df_data[['skill_no', 'skill_name']], left_on='skill_pri', right_on='skill_no', how='left')
    df_de_para = df_de_para.dropna(subset=['skill_name', 'profile'])
    dict_profile = pd.Series(df_de_para.profile.values, index=df_de_para.skill_name).to_dict()
    df_mapa['profile_navegacao'] = df_mapa['skill_name'].map(dict_profile)
    df_final = df_mapa.dropna(subset=['profile_navegacao'])
    
    cols_config = [c for c in df_final.columns if c not in ['skill_name', 'profile_navegacao']]
    print(f"Dados prontos. {len(df_final)} Skills para configurar.")

except Exception as e:
    print(f"ERRO DE ARQUIVO: {e}")
    exit()

# 2. INICIAR NAVEGADOR
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
time.sleep(1)
driver.execute_script("document.body.style.zoom='50%'")
time.sleep(1)
driver.find_element(By.ID, "vUSERNAME").send_keys(user)
time.sleep(1)
driver.find_element(By.ID, "vUSERPASSWORD").send_keys(password)
wait.until(EC.element_to_be_clickable((By.ID, "BTNENTER"))).click()

time.sleep(2)
driver.execute_script("document.body.style.zoom='65%'")
time.sleep(1)

print("Entrando no módulo...")
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0004"))).click()
time.sleep(3)
url_home_modulo = driver.current_url

# 3. LOOP PRINCIPAL
for index, row in df_final.iterrows():
    nome_skill = str(row['skill_name'])
    nome_profile = str(row['profile_navegacao'])
    
    print(f"\n[{index+1}/{len(df_final)}] Skill: {nome_skill}")
    
    try:
        # Navegação
        if driver.current_url != url_home_modulo:
            driver.get(url_home_modulo)
            time.sleep(2)
            driver.execute_script("document.body.style.zoom='65%'")

        try:
            profile_el = wait.until(EC.presence_of_element_located((By.XPATH, f"//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='{nome_profile}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", profile_el)
            profile_el.click()
            time.sleep(1.5)
            
            prd_el = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='PRD']")))
            prd_el.click()
            time.sleep(1.5)
            
            skill_el = wait.until(EC.presence_of_element_located((By.XPATH, f"//span[text()='{nome_skill}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", skill_el)
            skill_el.click()
            time.sleep(3)
        except:
            print(f"    ⚠️ Erro Navegação.")
            continue

        # Configuração
        print("    -> Configurando...")
        
        for col_nome in cols_config:
            valor = row[col_nome]
            
            if pd.isna(valor) or str(valor).strip() == "":
                continue
            
            valor_str = str(valor).strip()
            
            try:
                # 1. Clicar em INSERIR
                btn_ins = wait.until(EC.presence_of_element_located((By.ID, "BTNINSERIR")))
                driver.execute_script("arguments[0].click();", btn_ins)
                time.sleep(1.5)
                
                # 2. Preencher CHAVE (Forçado via JS)
                driver.execute_script(f"document.getElementById('vCONFIGURACAODADOSCHAVE').value = '{col_nome}';")
                time.sleep(1)
                # 3. SELECIONAR TEXTO (CRUCIAL!)
                try:
                    drop = driver.find_element(By.ID, "vTIPOCHAVEID")
                    select = Select(drop)
                    select.select_by_visible_text("Texto")
                    time.sleep(1) # Espera o site reagir e mostrar o campo de texto
                except:
                    print("       ⚠️ Erro ao selecionar dropdown Texto.")
                
                # 4. Preencher VALOR (Agora sim ele deve existir)
                # Espera o campo aparecer na tela
                try:
                    campo_valor = wait.until(EC.visibility_of_element_located((By.ID, "vCAMPOVALORTEXTO")))
                    
                    # Técnica Híbrida (Limpa, Clica e Digita)
                    campo_valor.click()
                    time.sleep(0.5)
                    campo_valor.clear()
                    time.sleep(0.5)
                    # Injeta via JS pra garantir caracteres especiais
                    driver.execute_script(f"arguments[0].value = '{valor_str}';", campo_valor)

                except:
                    print(f"       ❌ Campo de texto não apareceu para {col_nome}!")
                    # Se falhar, tenta salvar vazio ou pula
                
                time.sleep(0.5)

                # 5. SALVAR
                btn_salvar = driver.find_element(By.ID, "BTNTRN_ENTER")
                driver.execute_script("arguments[0].click();", btn_salvar)
                
                time.sleep(3) # Tempo para processar e recarregar
                
            except Exception as e:
                print(f"       ❌ Erro ao processar {col_nome}: {e}")
                driver.get(driver.current_url) 
                time.sleep(2)

    except Exception as e:
        print(f"❌ Erro crítico: {e}")

print("\n--- FIM ---")