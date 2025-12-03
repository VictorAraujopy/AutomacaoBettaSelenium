import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from dotenv import load_dotenv
import os
#ele caminha pelos datasets lendo tudo q tem na pagina e cria o dump do sistema(pega tudo que tem escrito nos datasets)
load_dotenv()
user = os.getenv("USER")
password = os.getenv("PASSWORD")
os.environ['WDM_SSL_VERIFY'] = '0'

print("--- ROBÔ ASPIRADOR: PEGANDO TUDO O QUE HÁ NA TELA ---")

# 1. Carregar Profiles
try:
    df = pd.read_csv(r"Dados/dados_prontos_para_bot.csv")
    lista_profiles = df['profile'].unique()
except:
    print("ERRO CSV. Verifique o caminho.")
    exit()

# 2. Navegador
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--no-sandbox")
options.add_argument("--remote-allow-origins=*")
options.add_argument("--ignore-certificate-errors")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()

# Login
driver.get("https://par1.c2x.com.br/login.aspx")
wait = WebDriverWait(driver, 20)
driver.find_element(By.ID, "vUSERNAME").send_keys(user)
time.sleep(0.5)
driver.find_element(By.ID, "vUSERPASSWORD").send_keys(password)
wait.until(EC.element_to_be_clickable((By.ID, "BTNENTER"))).click()

# Zoom e Módulo
time.sleep(2)
driver.execute_script("document.body.style.zoom='65%'")
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0004"))).click()
time.sleep(3)
url_home_modulo = driver.current_url

dados_brutos = []

# --- LOOP ---
for profile_nome in lista_profiles:
    print(f"\n>> Entrando em: {profile_nome}")
    
    try:
        # Reset
        if driver.current_url != url_home_modulo:
            driver.get(url_home_modulo)
            time.sleep(2)
            driver.execute_script("document.body.style.zoom='65%'")

        # 1. Profile
        try:
            xpath_prof = f"//span[contains(@class, 'AttributeHomeModulesBigTitle') and contains(text(), '{profile_nome}')]"
            el_prof = wait.until(EC.presence_of_element_located((By.XPATH, xpath_prof)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el_prof)
            el_prof.click()
            time.sleep(2)
        except:
            print(f"    ⚠️ Pulei profile {profile_nome}")
            continue

        # 2. PRD
        time.sleep(1)
        driver.execute_script("document.body.style.zoom='10%'")
        time.sleep(1)
        try:
            xpath_prd = "//span[contains(@class, 'AttributeHomeModulesBigTitle') and contains(text(), 'PRD')]"
            el_prd = wait.until(EC.presence_of_element_located((By.XPATH, xpath_prd)))
            el_prd.click()
            print("    Aguardando carregamento...")
            time.sleep(6) # Tempo alto pro sistema carregar tudo
        except:
            print("    ⚠️ PRD não achado")
            continue

        time.sleep(1)
        driver.execute_script("document.body.style.zoom='10%'")
        time.sleep(1)

        # 3. ASPIRADOR DE PÓ (DUMP DE TEXTO)
        print("    🧹 Aspirando textos da tela...")
        
        # Pega todos os spans, divs e links
        elementos = driver.find_elements(By.XPATH, "//span | //div | //a")
        
        count = 0
        for el in elementos:
            try:
                # Pega o texto VISÍVEL
                txt_visivel = el.text.strip()
                # Pega o texto OCULTO (HTML)
                txt_oculto = el.get_attribute("textContent").strip()
                
                # Escolhe o melhor (se tiver visivel, usa. Se nao, usa oculto)
                texto_final = txt_visivel if txt_visivel else txt_oculto
                
                # Se for muito curto ou irrelevante, ignora
                if len(texto_final) < 3: continue
                if texto_final in ["Segmento", "Aplicação", "Ambiente", "Roteamento", "PRD"]: continue
                
                # Salva TUDO no Excel pra gente analisar depois
                dados_brutos.append({
                    'profile': profile_nome,
                    'texto_encontrado': texto_final
                })
                count += 1
            except:
                pass
        
        print(f"    ✅ Capturei {count} textos brutos.")

    except Exception as e:
        print(f"    ❌ Erro: {e}")

# --- SALVAR DUMP ---
if dados_brutos:
    df_dump = pd.DataFrame(dados_brutos)
    df_dump = df_dump.drop_duplicates()
    df_dump.to_excel("DUMP_DO_SISTEMA.xlsx", index=False)
    #ele le tudo nas paginas e salva num excel
    print("\n✅ ARQUIVO 'DUMP_DO_SISTEMA.xlsx' GERADO!")
    print("Abra esse arquivo e procure os nomes dos skills. Se estiverem lá, a gente filtra.")
else:
    print("\n❌ O Aspirador não pegou nada. O sistema está blindado.")