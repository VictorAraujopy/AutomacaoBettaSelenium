import json
import pandas as pd
import re
#LE o profile.json e crio um excel MAPA_CONFIGURACAO_DATAS_JSON.xlsx com as configurações extraídas do json
print("--- TRADUTOR DE JSON PARA EXCEL (CONFIGURAÇÃO) ---")

# 1. Carregar o profile.json
try:
    with open("profile.json", "r", encoding="utf-8") as f:
        data_json = json.load(f)
    print("Arquivo 'profile.json' carregado.")
except Exception as e:
    print(f"ERRO: Não achei o arquivo profile.json. Detalhe: {e}")
    exit()

# 2. Dicionário de Tradução (De -> Para)
mapa_traducao = {
    'PA': 'ATRIBUTO_PRIORIZADO',
    'PT': 'PRIORIDADE_TRANSFERENCIA',
    'Exp': 'EXPANSAO_SKILLS_BASICAS',
    'MSG_Emerg': 'FRASE_EMERGENCIAL'
    # 'BLC_EPS' e 'TFH' serão ignorados pois não estão aqui
}

lista_final = []

# 3. Processar cada Profile e cada Data
for item in data_json:
    profile_name = item.get('profile', {}).get('profileName', 'Desconhecido')
    datas_dict = item.get('data', {})
    
    for data_name, data_info in datas_dict.items():
        # Pega a string bagunçada: "PA:|PT:200:399|Exp:|TFH:|MSG_Emerg:"
        valor_string = data_info.get('Value', '')
        
        if not valor_string:
            continue
            
        # Quebra nos pipes '|'
        partes = valor_string.split('|')
        
        for parte in partes:
            if ':' in parte:
                # Separa a Chave do Valor (ex: "PT" e "200:399")
                # O split(..., 1) garante que só quebra no primeiro dois pontos
                chave_json, valor_real = parte.split(':', 1)
                
                # Limpa espaços
                chave_json = chave_json.strip()
                valor_real = valor_real.strip()
                
                # Se o valor estiver vazio, PULA (Não vamos configurar vazio)
                if not valor_real:
                    continue
                
                # Verifica se temos tradução para essa chave
                if chave_json in mapa_traducao:
                    chave_sistema = mapa_traducao[chave_json]
                    
                    # Adiciona na lista para o Excel
                    lista_final.append({
                        'data_name': data_name,       # Nome do Tipo de Serviço (ex: HD_ANGELONI)
                        'profile_context': profile_name, # Apenas para referência
                        'chave_config': chave_sistema, # Ex: PRIORIDADE_TRANSFERENCIA
                        'valor_config': valor_real     # Ex: 200:399
                    })

# 4. Salvar o Excel
if lista_final:
    df = pd.DataFrame(lista_final)
    output_file = "MAPA_CONFIGURACAO_DATAS_JSON.xlsx"
    df.to_excel(output_file, index=False)
    print(f"\n✅ SUCESSO! Mapa gerado: '{output_file}'")
    print(f"Total de configurações encontradas: {len(df)}")
    print(df.head())
else:
    print("⚠️ Aviso: Nenhuma configuração válida encontrada (tudo vazio ou chaves desconhecidas).")