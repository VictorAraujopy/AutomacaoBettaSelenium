import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import os
from dotenv import load_dotenv
import os
load_dotenv()
#usa os dados do excel dados_prontos_para_Bot_corrigido para buscar os ids no sistema atraves dos rotulos fixos
user = os.getenv("USER")
password = os.getenv("PASSWORD")
os.environ['WDM_SSL_VERIFY'] = '0'

print("--- ROBÔ ESPIÃO 4.0: BUSCA POR RÓTULO FIXO (SEGMENTO/APLICAÇÃO) ---")

# 1. Carregar o Mapa
try:
    df = pd.read_csv(r"Dados/dados_prontos_para_bot_CORRIGIDO.csv")
    print(f"Lendo {len(df)} serviços do arquivo.")
except:
    print("ERRO CRÍTICO: Arquivo 'dados_prontos_para_bot.csv' não encontrado.")
    exit()

resultados = []

# 2. Iniciar Navegador
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.maximize_window()

# Login
print("Fazendo login...")
driver.get("https://par1.c2x.com.br/login.aspx")
wait = WebDriverWait(driver, 15)

driver.find_element(By.ID, "vUSERNAME").send_keys(user)
time.sleep(0.5)
driver.find_element(By.ID, "vUSERPASSWORD").send_keys(password)
wait.until(EC.element_to_be_clickable((By.ID, "BTNENTER"))).click()

time.sleep(1)
driver.execute_script("document.body.style.zoom='10%'")
time.sleep(1)
# -------------------

# Entra no módulo
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0004"))).click()
time.sleep(1)
url_home_modulo = driver.current_url
time.sleep(1)
driver.execute_script("document.body.style.zoom='10%'")
time.sleep(1)
# --- FUNÇÃO QUE PEGA O VALOR DO VIZINHO ---
def pegar_valor_do_vizinho(driver, rotulo_fixo):
    """
    1. Procura o texto fixo (ex: 'Aplicação').
    2. Pega o elemento que está logo DEPOIS dele (o vizinho).
    3. Extrai o número desse vizinho.
    """
    try:
        # XPath Genérico Poderoso:
        # Procura qualquer elemento que tenha o texto exato OU contenha o texto
        # E pega o elemento seguinte (following-sibling ou following)
        
        # Tentativa 1: Estrutura de Tabela (TD vizinho) - Muito comum nesses sistemas
        try:
            xpath_tabela = f"//td[contains(text(), '{rotulo_fixo}')]/following-sibling::td[1]"
            elemento = driver.find_element(By.XPATH, xpath_tabela)
        except:
            # Tentativa 2: Estrutura Genérica (Qualquer elemento vizinho)
            xpath_generico = f"//*[contains(text(), '{rotulo_fixo}')]/following::*[1]"
            elemento = driver.find_element(By.XPATH, xpath_generico)
            
        texto_valor = elemento.text.strip() # Ex: "39 -- NUCLEO_NEGOCIOS"
        
        # Extrai o número
        # Pega numero no começo "99 --" ou número sozinho "99"
        match = re.search(r'^(\d+)', texto_valor)
        if match: return match.group(1)
        
        # Se não achou no começo, tenta qualquer número
        match = re.search(r'\d+', texto_valor)
        if match: return match.group(0)
        
        return "?"
        
    except:
        return "?"

# --- LOOP PRINCIPAL ---
for index, row in df.iterrows():
    nome_servico = row['tipo_servico']
    profile_alvo = row['profile']    
    skill_busca  = row['skill_name'] 
    
    print(f"\n[{index+1}/{len(df)}] Processando: {nome_servico}")
    print(f"    Navegando para: {skill_busca}...")
    
    val_final = "ERRO"

    try:
        # 1. Resetar Navegação
        if driver.current_url != url_home_modulo:
            driver.get(url_home_modulo)
            time.sleep(1.5)

        # 2. Navegar
        try:
            time.sleep(1)
            driver.execute_script("document.body.style.zoom='10%'")
            time.sleep(1)
            # Clica no Profile
            profile_el = wait.until(EC.presence_of_element_located((By.XPATH, f"//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='{profile_alvo}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", profile_el)
            profile_el.click()
            time.sleep(1.5)
            time.sleep(1)
            driver.execute_script("document.body.style.zoom='10%'")
            time.sleep(1)
            # Clica no PRD
            wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'AttributeHomeModulesBigTitle') and @title='PRD']"))).click()
            time.sleep(1.5)
            time.sleep(1)
            driver.execute_script("document.body.style.zoom='10%'")
            time.sleep(1)
            # Clica no Skill
            skill_el = wait.until(EC.presence_of_element_located((By.XPATH, f"//span[text()='{skill_busca}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", skill_el)
            skill_el.click()
            time.sleep(3) # Tempo para os dados aparecerem na direita
            
        except:
            print(f"    ⚠️ Erro ao navegar.")
            resultados.append({'tipo_servico': nome_servico, 'valor_final': 'ERRO_NAV'})
            continue

        # 3. EXTRAÇÃO VIA RÓTULOS FIXOS (Agora vai!)
        print("    🕵️ Lendo valores ao lado dos rótulos...")
        
        # Procura "Segmento" -> Pega o vizinho
        seg = pegar_valor_do_vizinho(driver, "Segmento")
        if seg == "?" or seg == "": seg = "9"
        
        # Procura "Aplicação" (Esse é o nome do Profile na sua tela!) -> Pega o vizinho
        prof = pegar_valor_do_vizinho(driver, "Aplicação")
        
        # Procura "Ambiente" -> Pega o vizinho
        amb = pegar_valor_do_vizinho(driver, "Ambiente")
        
        # Procura "ID" -> Pega o vizinho
        dat = pegar_valor_do_vizinho(driver, "ID")
        
        # Monta
        val_final = f"{seg}|{prof}|{amb}|{dat}"
        print(f"    ✅ Capturado: {val_final}")

    except Exception as e:
        print(f"    ❌ Erro: {e}")#
        val_final = "ERRO_GERAL"
    
    resultados.append({
        'tipo_servico': nome_servico,
        'valor_final': val_final
    })

# --- FIM ---
print("\n--- SALVANDO ARQUIVO FINAL... ---")
df_resultado = pd.DataFrame(resultados)
df_resultado.to_excel("ARQUIVO_FINAL_COM_IDS.xlsx", index=False)
print("SUCESSO! Verifique o arquivo 'ARQUIVO_FINAL_COM_IDS.xlsx'.")