import pandas as pd



print("Lendo arquivos...")

df_excel = pd.read_excel(r"Dados/Workflow_Profiles_11_13_v1.xlsx") 
df_csv = pd.read_csv(r"Dados/DataID.csv")

# 2. Converter as colunas de ID para string e LIMPAR SUJEIRA
# Removemos decimais (.0) e espaços em branco para garantir que o cruzamento funcione
df_excel['skill_pri'] = df_excel['skill_pri'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
df_csv['skill_no'] = df_csv['skill_no'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

print("Cruzando dados...")

# 3. Fazer o cruzamento (Left Join)
# Trazemos o 'skill_name' do CSV baseado no número do skill
df_merged = pd.merge(
    df_excel, 
    df_csv[['skill_no', 'skill_name']], 
    left_on='skill_pri', 
    right_on='skill_no', 
    how='left' 
)

# 4. Filtrar colunas úteis (AQUI MUDOU)
# Pegamos o 'tipo_servico' (nome para cadastrar) e as colunas de navegação
df_final = df_merged[['tipo_servico', 'profile', 'skill_pri', 'skill_name']]

# Filtra apenas os que acharam correspondência no CSV (sem skill_name o robô não acha o item)
df_final = df_final.dropna(subset=['skill_name'])


df_final = df_final.drop_duplicates()

# 6. Salvar
df_final.to_csv("dados_prontos_para_bot.csv", index=False, encoding="utf-8-sig")

print(f"Tabela criada com sucesso! Total de linhas para o robô: {len(df_final)}")
print(df_final.head())