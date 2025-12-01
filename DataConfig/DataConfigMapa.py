import pandas as pd

print("--- BOSS FINAL: PREPARANDO O MAPA DE CONFIGURAÇÃO ---")

# 1. Carregar os arquivos
try:
    # Lendo o Workflow como Excel (ajuste o nome se for .xlsx)
    # ATENÇÃO: Se o arquivo foi enviado com extensão .csv pelo sistema, o pandas pode ter confundido.
    # Se for .xlsx, use read_excel. Se for .csv, use read_csv. 
    # Vou deixar read_excel pois você confirmou que é Excel.
    
    # Se o nome do arquivo for mesmo .csv, troque para pd.read_csv("...", encoding='latin-1', sep=';')
    # Assumindo que é um Excel real (.xlsx):
    df_workflow = pd.read_excel(r"DataConfig/Cópia de Workflow_Profiles_11_13_v1.xlsx") 
    
    df_data = pd.read_csv(r"DataConfig/Data - 2025-11-26T152141.453.csv")
    print("Arquivos carregados com sucesso!")
    
except Exception as e:
    # Fallback: Se der erro lendo como Excel, tenta como CSV (caso o arquivo tenha sido renomeado)
    try:
        print("Tentando ler como CSV...")
        df_workflow = pd.read_csv("Cópia de Workflow_Profiles_11_13_v1.xlsx - Planilha1.csv")
        df_data = pd.read_csv("Data - 2025-11-26T152141.453.csv")
        print("Arquivos carregados via CSV!")
    except Exception as e2:
        print(f"ERRO CRÍTICO: Não consegui ler o arquivo. Detalhe: {e}")
        exit()

# 2. Limpeza de chaves
# Garante que os IDs sejam texto e sem '.0'
if 'skill_pri' in df_workflow.columns:
    df_workflow['skill_pri'] = df_workflow['skill_pri'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

if 'skill_no' in df_data.columns:
    df_data['skill_no'] = df_data['skill_no'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

# 3. Merge para pegar o skill_name (Nome do Skill)
# Trazemos o nome do skill do arquivo Data para dentro do Workflow
df_merged = pd.merge(
    df_workflow,
    df_data[['skill_no', 'skill_name']],
    left_on='skill_pri',
    right_on='skill_no',
    how='left'
)

# Filtra apenas linhas que encontraram um skill correspondente
df_merged = df_merged.dropna(subset=['skill_name'])

# 4. Selecionar colunas de configuração
try:
    cols = df_merged.columns.tolist()
    
    # Acha onde começa a configuração
    idx_inicio = cols.index('HOUR_OF_OPERATION')
    
    # Pega todas as colunas dali pra frente
    config_cols = cols[idx_inicio:]
    
    # Remove colunas de sistema que podem ter vindo junto, se necessário
    if 'skill_name' in config_cols: config_cols.remove('skill_name')
    if 'skill_no' in config_cols: config_cols.remove('skill_no')
    
    # Monta a lista final: Nome do Skill (Chave) + Todas as Configurações
    final_cols = ['skill_name'] + config_cols
    
    df_final = df_merged[final_cols]
    
    print(f"Colunas de configuração identificadas: {len(config_cols)}")
    
except ValueError:
    print("ERRO: Coluna 'HOUR_OF_OPERATION' não encontrada. Verifique o cabeçalho do Excel.")
    exit()

# 5. Remover duplicatas (1 Linha por Skill)
# Se 10 serviços usam o mesmo skill, ficamos com apenas 1 linha de configuração para esse skill.
print(f"Total de linhas brutas: {len(df_final)}")
df_final = df_final.drop_duplicates(subset=['skill_name'])
print(f"Total de Skills únicos para configurar: {len(df_final)}")

# 6. Salvar o Mapa Final
output_filename = "MAPA_CONFIGURACAO_SKILLS.xlsx"
df_final.to_excel(output_filename, index=False)
print(f"\n✅ SUCESSO: Arquivo '{output_filename}' gerado!")