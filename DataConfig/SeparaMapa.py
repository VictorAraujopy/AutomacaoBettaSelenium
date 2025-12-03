import pandas as pd
#separa o mapa de configuração em dois arquivos: um para campos de texto e outro para campos de lista (dropdown)
#MAPA_CONFIGURACAO_TEXTO.xlsx e MAPA_CONFIGURACAO_LISTA.xlsx
print("--- SEPARADOR DE MAPAS (TEXTO vs LISTA) ---")

try:
    # 1. Carregar o Mapa Original
    # (Se o seu arquivo tiver outro nome, ajuste aqui)
    try:
        df = pd.read_excel(r"Dados/MAPA_CONFIGURACAO_SKILLS.xlsx")
    except:
       print("erro")

    print(f"Total de linhas originais: {len(df)}")
    
    # 2. Definir quais colunas são do tipo LISTA (Dropdown)
    colunas_lista = [
        'BALANCEAMENTO_EPS',
        'CHAVE_CALENDARIO_GLOBAL',
        'CHAVE_FRASE_EMERGENCIAL',
        'MODOS_CALLBACK'
    ]
    
    # 3. Identificar todas as colunas de configuração (exceto a chave skill_name)
    todas_colunas = df.columns.tolist()
    if 'skill_name' in todas_colunas:
        todas_colunas.remove('skill_name')
        
    # Se houver colunas de navegação extras (como profile_navegacao), removemos também para limpar
    if 'profile_navegacao' in todas_colunas:
        todas_colunas.remove('profile_navegacao')

    # 4. Separar as colunas de TEXTO (Tudo que não está na lista de 'colunas_lista')
    # List comprehension para manter a ordem original
    colunas_texto = [col for col in todas_colunas if col not in colunas_lista]
    
    print(f"Colunas Tipo LISTA: {len(colunas_lista)}")
    print(f"Colunas Tipo TEXTO: {len(colunas_texto)}")
    
    # 5. Criar os dois DataFrames
    # Mantemos 'skill_name' em ambos para saber de quem é a configuração
    df_texto = df[['skill_name'] + colunas_texto].copy()
    df_lista = df[['skill_name'] + colunas_lista].copy()
    
    # 6. Salvar os arquivos
    file_texto = "MAPA_CONFIGURACAO_TEXTO.xlsx"
    file_lista = "MAPA_CONFIGURACAO_LISTA.xlsx"
    
    df_texto.to_excel(file_texto, index=False)
    df_lista.to_excel(file_lista, index=False)
    
    print(f"\n✅ Arquivo gerado: {file_texto}")
    print(f"✅ Arquivo gerado: {file_lista}")
    
except Exception as e:
    print(f"Erro: {e}")