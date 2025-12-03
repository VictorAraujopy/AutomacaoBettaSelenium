import pandas as pd
from fuzzywuzzy import process, fuzz
#pega o dump do sistema e corrige os nomes dos DATASETS comparando os dados_prontos_para_bot.csv com o dump do sistema
#cria o arquivo dados_prontos_para_bot_CORRIGIDO.csv que é usado pelo robô de cadastro de tipo de serviço
print("--- CORRETOR DE NOMES (MODO NOME PURO) ---")

try:
    # 1. Carregar os Arquivos
    # Seu arquivo original com os nomes "errados"
    df_origem = pd.read_csv(r"Dados/dados_prontos_para_bot.csv")
    
    # O arquivo que o Robô Aspirador gerou (O que existe de verdade)
    # ATENÇÃO: Ajuste o nome se for .xlsx ou .csv
    
    df_sistema = pd.read_excel("DUMP_DO_SISTEMA.xlsx")
    
        
    print(f"Origem: {len(df_origem)} linhas | Sistema: {len(df_sistema)} linhas")

except Exception as e:
    print(f"❌ Erro ao ler arquivos: {e}")
    exit()

# Prepara a lista de nomes do sistema (remove duplicatas e vazios)
df_sistema['nome_limpo'] = df_sistema['texto_encontrado'].astype(str).str.strip()
df_sistema = df_sistema[df_sistema['nome_limpo'].str.len() > 3] # Remove lixo curto

# Lista final
novos_dados = []

print("\nCorrigindo nomes...")

# 2. Loop de Correção
for index, row in df_origem.iterrows():
    nome_errado = str(row['skill_name']).strip()
    profile = row['profile']
    
    # Filtra o sistema pelo mesmo Profile (para aumentar precisão)
    opcoes = df_sistema[df_sistema['profile'] == profile]
    lista_candidatos = opcoes['nome_limpo'].tolist()
    
    nome_corrigido = nome_errado # Assume o original como padrão
    status = "Nao Encontrado"
    score = 0
    
    if lista_candidatos:
        # Acha o mais parecido
        match = process.extractOne(nome_errado, lista_candidatos, scorer=fuzz.token_sort_ratio)
        if match:
            melhor_nome = match[0]
            score = match[1]
            
            # Se a certeza for boa (> 85%), usamos o nome do sistema
            if score >= 85:
                nome_corrigido = melhor_nome 
                status = "Corrigido"
            elif score >= 60:
                status = f"Duvida ({score}%)"
                # Em caso de dúvida, mantemos o original ou marcamos para você ver
                # nome_corrigido = melhor_nome (Opção: descomentar se quiser arriscar)
    
    # Monta a linha mantendo a estrutura original
    novos_dados.append({
        'tipo_servico': row['tipo_servico'],
        'profile': row['profile'],
        'skill_pri': row['skill_pri'],
        'skill_name': nome_corrigido,  # <--- Nome corrigido
        'STATUS_MATCH': status,        # Pra você conferir
        'SCORE': score
    })

# 3. Salvar
df_final = pd.DataFrame(novos_dados) 
arquivo_saida = "dados_prontos_para_bot_CORRIGIDO.csv"
df_final.to_csv(arquivo_saida, index=False)

print(f"\n✅ PRONTO! Arquivo gerado: '{arquivo_saida}'")
print("Abra esse arquivo e ordene pela coluna 'SCORE' (do menor para o maior) para ver os problemas.")