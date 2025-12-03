import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv
import os
load_dotenv()
#cadastra os tipo serviço com o arquivo-final-com-ids.xlsx gerado anteriormente
user = os.getenv("USER")
password = os.getenv("PASSWORD")

os.environ['WDM_SSL_VERIFY'] = '0'

print("--- Iniciando Robô de Cadastro ---")

try:
    # MUDANÇA AQUI: Lendo como Excel (.xlsx)
    df = pd.read_excel("ARQUIVO_FINAL_COM_IDS.xlsx")
    
    # FILTRO: Remove se 'valor_final' for vazio (NaN) OU se contiver "ERRO"
    df_limpo = df.dropna(subset=['valor_final']) # Tira vazios
    df_limpo = df_limpo[~df_limpo['valor_final'].astype(str).str.contains("ERRO")] # Tira erros
    
    print(f"Total de itens no arquivo: {len(df)}")
    print(f"Itens válidos para cadastro: {len(df_limpo)}")
    
except Exception as e:
    print(f"ERRO CRÍTICO ao ler o Excel: {e}")
    exit()

# 2. INICIAR NAVEGADOR (Seu código original)
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

time.sleep(2)
driver.execute_script("document.body.style.zoom='30%'")
time.sleep(1)

# Navegação (Seu código original)
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0008"))).click()
time.sleep(1)
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0001"))).click()
time.sleep(1)
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0001"))).click()
time.sleep(1)
wait.until(EC.element_to_be_clickable((By.ID, "LAYOUTOPTIONANDTITLE_0001"))).click()
time.sleep(1)

# --- 3. LOOP DE CADASTRO (AQUI ENTRA A MÁGICA) ---
for index, row in df_limpo.iterrows():
    # Variáveis vindas do Excel
    nome_servico = row['tipo_servico'] 
    valor_ids    = row['valor_final']
    
    print(f"[{index+1}] Cadastrando: {nome_servico} -> {valor_ids}")
    
    try:
        # Clica em INSERIR
        btn_inserir = wait.until(EC.element_to_be_clickable((By.ID, "BTNINSERIR")))
        driver.execute_script("arguments[0].click();", btn_inserir)
        time.sleep(2)
        
        # Preenche o NOME (Usando a variável do Excel)
        campo_chave = wait.until(EC.element_to_be_clickable((By.ID, "vCONFIGURACAODADOSCHAVE")))
        time.sleep(1)
        campo_chave.clear()
        time.sleep(1)
        campo_chave.send_keys(nome_servico)
        time.sleep(2)

        # 2. SELECIONAR "TEXTO" NO DROPDOWN (A MUDANÇA É AQUI)
        try:
            # Acha o elemento select
            elemento_dropdown = wait.until(EC.presence_of_element_located((By.ID, "vTIPOCHAVEID")))
            
            # Cria o objeto Select
            selecao = Select(elemento_dropdown)
            
            # Seleciona a opção que está escrita "Texto"
            selecao.select_by_visible_text("Texto")
            
            time.sleep(1)
        except Exception as e:
            print(f"⚠️ Erro ao selecionar o dropdown: {e}")
        
        # Preenche o VALOR (Usando a variável do Excel)
        # Nota: O ID do campo valor geralmente é vCONFIGURACAODADOSVALOR. Se for diferente, ajuste.
        campo_valor = driver.find_element(By.ID, "vCAMPOVALORTEXTO")
        time.sleep(1)
        campo_valor.clear()
        time.sleep(1)
        campo_valor.send_keys(valor_ids)    # <--- AQUI VAI O CÓDIGO
        
        time.sleep(1)
        
        # Salva (BTNTRN_ENTER)
        btn_salvar = driver.find_element(By.ID, "BTNENTER")
        driver.execute_script("arguments[0].click();", btn_salvar)
        
        # Espera voltar para a lista antes do próximo
        time.sleep(1.5)
        
    except Exception as e:
        print(f"❌ Erro ao cadastrar {nome_servico}: {e}")
        # Tenta voltar para a URL da lista para não quebrar o loop inteiro
        driver.get(driver.current_url)
        time.sleep(2)

print("--- FIM DO PROCESSO ---")