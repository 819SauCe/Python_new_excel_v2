import pandas as pd
import os

def atualizar_planilha_com_dados(caminho_planilha, caminho_txt_1, caminho_txt_2):
    """
    Atualiza a planilha especificada com dados ausentes usando os arquivos TXT fornecidos.

    Args:
        caminho_planilha (str): Caminho para o arquivo Excel que será atualizado.
        caminho_txt_1 (str): Caminho para o primeiro arquivo TXT de referência.
        caminho_txt_2 (str): Caminho para o segundo arquivo TXT de referência.

    Returns:
        bool: Retorna True se a planilha for atualizada com sucesso, caso contrário False.
    """
    # Verificar se a planilha e os arquivos necessários existem
    if not os.path.exists(caminho_planilha):
        print(f"A planilha '{caminho_planilha}' não foi encontrada.")
        return False

    if not os.path.exists(caminho_txt_1) or not os.path.exists(caminho_txt_2):
        print("Um ou ambos os arquivos TXT necessários não foram encontrados.")
        return False

    # Carregar a planilha existente
    try:
        df_planilha = pd.read_excel(caminho_planilha)
    except Exception as e:
        print(f"Erro ao carregar a planilha '{caminho_planilha}': {e}")
        return False

    # Carregar os arquivos TXT
    try:
        df_estoque_lote = pd.read_csv(caminho_txt_1, sep='|', encoding='ISO-8859-1')
        df_cf_estoque = pd.read_csv(caminho_txt_2, sep='|', encoding='ISO-8859-1')
    except Exception as e:
        print(f"Erro ao carregar os arquivos TXT: {e}")
        return False

    # Renomear colunas para padronização
    colunas_necessarias = ["Código Produto", "Nome do Produto", "Marca", "Inativo"]
    try:
        df_estoque_lote = df_estoque_lote[colunas_necessarias]
        df_cf_estoque = df_cf_estoque[colunas_necessarias]
    except KeyError as e:
        print(f"Erro ao processar colunas necessárias: {e}")
        return False

    # Renomear colunas
    for df in [df_estoque_lote, df_cf_estoque]:
        df.rename(columns={
            "Código Produto": "Código do produto",
            "Nome do Produto": "Nome",
            "Marca": "Marca",
            "Inativo": "Inativo"
        }, inplace=True)

    # Concatenar os dados dos arquivos TXT
    df_referencia = pd.concat([df_estoque_lote, df_cf_estoque], ignore_index=True)

    # Função para buscar valores ausentes
    def buscar_valor(row, coluna):
        if pd.isna(row[coluna]) or row[coluna] == "":
            dados_produto = df_referencia[df_referencia["Código do produto"] == row["Código do produto"]]
            if not dados_produto.empty:
                info_encontrada = dados_produto[coluna].dropna().unique()
                return info_encontrada[0] if len(info_encontrada) > 0 else ""
            return ""
        return row[coluna]

    # Atualizar colunas "Inativo" e "Marca" com dados dos arquivos TXT
    df_planilha["Inativo"] = df_planilha.apply(lambda row: buscar_valor(row, "Inativo"), axis=1)
    df_planilha["Marca"] = df_planilha.apply(lambda row: buscar_valor(row, "Marca"), axis=1)

    # Salvar o resultado na mesma planilha
    # sla
    try:
        df_planilha.to_excel(caminho_planilha, index=False)
        print(f"A planilha '{caminho_planilha}' foi atualizada com sucesso.")
        return True
    except Exception as e:
        print(f"Erro ao salvar a planilha '{caminho_planilha}': {e}")
        return False


def atualizar_planilha_com_remocao():
    """
    Atualiza a planilha '1.xlsx' removendo itens duplicados e adicionando lógica de inatividade.
    """
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
        .astype(str)
        .str.replace(",", ".", regex=False)
    )
    df_txt["Quantidade disponivel"] = pd.to_numeric(df_txt["Quantidade disponivel"], errors='coerce').fillna(0).astype(int)

    # Remover linhas duplicadas e marcar como inativo
    df_txt["Inativo"] = df_txt.apply(
        lambda row: "Inativo" if row["Quantidade disponivel"] == 0 else row.get("Inativo", None), axis=1
    )
    df_txt = df_txt[~df_txt["Código do produto"].isin(df_planilha["Código do produto"])]

    # Concatenar e salvar
    df_atualizado = pd.concat([df_planilha, df_txt], ignore_index=True)
    df_atualizado.to_excel("2.xlsx", index=False)


if __name__ == "__main__":
    caminho_planilha = "2.xlsx"
    caminho_txt_1 = r"\\WKRadar-SRV\Relatórios\TI-PBI_EstoqueLote.txt"
    caminho_txt_2 = r"\\WKRadar-SRV\Relatórios\CF_TI-PBI_Estoque.txt"

    atualizar_planilha_com_remocao()
    atualizar_planilha_com_dados(caminho_planilha, caminho_txt_1, caminho_txt_2)
