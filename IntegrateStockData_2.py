#Segundo passo da verificação

import pandas as pd
import os

def atualizar_planilha():
    # Caminho do arquivo da planilha criada anteriormente (raiz do código)
    caminho_planilha = "1.xlsx"

    if not os.path.exists(caminho_planilha):
        print("Nenhuma planilha '1.xlsx' encontrada na raiz do código.")
        return

    # Caminho do arquivo CF_TI-PBI_Estoque.txt
    caminho_txt = r"\\WKRadar-SRV\Relatórios\CF_TI-PBI_Estoque.txt"

    if not os.path.exists(caminho_txt):
        print(f"Arquivo {caminho_txt} não encontrado.")
        return

    # Carregar a planilha existente
    try:
        df_planilha = pd.read_excel(caminho_planilha)
    except Exception as e:
        print(f"Erro ao carregar a planilha '{caminho_planilha}': {e}")
        return

    # Carregar o arquivo TXT
    try:
        df_txt = pd.read_csv(caminho_txt, sep='|', encoding='ISO-8859-1')
    except Exception as e:
        print(f"Erro ao carregar o arquivo '{caminho_txt}': {e}")
        return

    # Selecionar apenas as colunas necessárias do TXT
    colunas_necessarias = ["Código Produto", "Nome do Produto", "Marca", "Qtde Disponível"]
    df_txt = df_txt[colunas_necessarias]

    # Renomear colunas do arquivo TXT para correspondência
    df_txt = df_txt.rename(columns={
        "Código Produto": "Código do produto",
        "Nome do Produto": "Nome",
        "Marca": "Marca",
        "Qtde Disponível": "Quantidade disponivel"
    })

    # Substituir vírgulas por pontos e tratar valores não numéricos
    df_txt["Quantidade disponivel"] = (
        df_txt["Quantidade disponivel"]
        .astype(str)  # Converter para string para processar substituições
        .str.replace(",", ".", regex=False)  # Substituir vírgula por ponto
    )

    # Manter apenas valores numéricos
    df_txt["Quantidade disponivel"] = pd.to_numeric(df_txt["Quantidade disponivel"], errors='coerce').fillna(0).astype(int)

    # Remover linhas do TXT que já existem na planilha
    df_txt = df_txt[~df_txt["Código do produto"].isin(df_planilha["Código do produto"])]

    if df_txt.empty:
        print("Nenhum novo produto a ser adicionado.")
        return

    # Concatenar os novos dados à planilha existente
    df_atualizado = pd.concat([df_planilha, df_txt], ignore_index=True)

    # Salvar o arquivo atualizado
    arquivo_saida = "2.xlsx"
    try:
        df_atualizado.to_excel(arquivo_saida, index=False)
        print(f"Planilha atualizada e salva como '{arquivo_saida}'.")
    except Exception as e:
        print(f"Erro ao salvar a planilha '{arquivo_saida}': {e}")

if __name__ == "__main__":
    atualizar_planilha()
