import pandas as pd
import json

# VARIAVEIS ===================================================================
caminho_arquivo_excel = './excel/original.xlsx'
caminho_arquivo_json = './excel/alunos.json'
colunas_desejadas = ['Código da pessoa', 'Nome da pessoa', 'Email']

# FUNÇÕES =====================================================================
def salvar_dados_em_json(dados, caminho_arquivo):
    try:
        with open(caminho_arquivo, 'w', encoding='utf-8') as json_file:
            json.dump(dados, json_file, indent=2, default=str, ensure_ascii=False)
        print(f"\033[92m✔ Os dados foram salvos com sucesso no arquivo: {caminho_arquivo}\033[0m")
    except Exception as e:
        print(f"\033[91m✖ Ocorreu um erro ao salvar os dados em JSON: {e}\033[0m")

def obter_linhas_sem_palavra_chave(caminho_arquivo_excel, colunas_desejadas):
    linhas_encontradas = {}

    try:
        # Lê todas as planilhas do arquivo Excel
        planilhas = pd.read_excel(caminho_arquivo_excel, sheet_name=None, dtype={'Código da pessoa': str})
        print(f"\033[94mℹ Planilhas lidas: {list(planilhas.keys())}\033[0m")  # Debug: Imprime os nomes das planilhas lidas

        for nome_da_planilha, df in planilhas.items():
            print(f"\033[94mℹ Processando planilha: {nome_da_planilha}\033[0m")  # Debug: Imprime o nome da planilha sendo processada
            print(f"\033[94mℹ Colunas encontradas na planilha: {df.columns.tolist()}\033[0m")  # Debug: Imprime as colunas encontradas

            # Remove espaços extras dos nomes das colunas
            df.columns = df.columns.str.strip()

            # Verifica se as colunas desejadas existem na planilha
            if set(colunas_desejadas).issubset(df.columns):
                print(f"\033[92m✔ Colunas desejadas encontradas na planilha: {nome_da_planilha}\033[0m")  # Debug
                # Seleciona apenas as colunas desejadas
                df = df[colunas_desejadas]

                # Converte colunas Timestamp para formato serializável
                for coluna in df.select_dtypes(include=['datetime64']).columns:
                    df[coluna] = df[coluna].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else None)

                # Remove linhas vazias
                df = df.dropna(how='all')
                print(f"\033[94mℹ Linhas após remoção de vazias: {len(df)}\033[0m")  # Debug: Imprime o número de linhas após remoção

                # Verifica se há linhas para processar após a remoção
                if df.empty:
                    print(f"\033[93m⚠ Nenhuma linha válida encontrada na planilha {nome_da_planilha} para adicionar ao dicionário\033[0m")  # Debug
                    continue  # Pula para a próxima planilha se esta estiver vazia

                # Adiciona as linhas da planilha correspondente ao dicionário
                linhas_encontradas[nome_da_planilha] = df.to_dict('records')

            else:
                print(f"\033[91m✖ As colunas desejadas não foram encontradas na planilha: {nome_da_planilha}\033[0m")  # Debug

        return linhas_encontradas

    except Exception as e:
        print(f"\033[91m✖ Ocorreu um erro ao ler as planilhas: {e}\033[0m")
        return None

# EXECUTANDO O PROGRAMA ========================================================
resultados = obter_linhas_sem_palavra_chave(caminho_arquivo_excel, colunas_desejadas)
if resultados is not None:
    print(f"\033[92m✔ Total de planilhas processadas: {len(resultados)}\033[0m")  # Debug: Imprime o total de planilhas processadas
    salvar_dados_em_json(resultados, caminho_arquivo_json)
else:
    print("\033[91m✖ Nenhuma planilha foi processada.\033[0m")