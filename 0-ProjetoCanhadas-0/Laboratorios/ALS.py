import pandas as pd
from openpyxl import Workbook

def main(uploaded_file, novo_caminho):
    # Ler o arquivo compartilhador.xlsx
    tabela = pd.read_excel(uploaded_file)

    # URL do arquivo Excel no GitHub
    url = "https://github.com/TecnologiaServmar/ProjetoCanhadas/raw/main/Tabelas%20Consulta/Banco%20de%20Dados/banco%20de%20dados%20-%20CEIMIC.xlsx"

    # Ler o arquivo Excel da URL
    tabela_bd = pd.read_excel(url)

    colunas_merge = ["ANALYTE", "Description"]

    # Pega todos os itens que têm em comum em relação à coluna "Análise" e "Descrição Método", junta em uma planilha só e salva como "Resultado_Merge.xlsx"
    tabela_merge = pd.merge(tabela, tabela_bd, how='left', on=colunas_merge)

    # Formata a coluna SAMPDATE para o padrão desejado
    tabela_merge['SAMPDATE'] = pd.to_datetime(tabela_merge['SAMPDATE'], format='%m/%d/%Y %H:%M').dt.strftime('%d/%m/%Y %H:%M')

    data_frame = tabela_merge

    # Pega da coluna "SAMPLENAME" todos os pms e pega da coluna "Descrição Método" todos as descrições dos métodos (pega os valores únicos e joga na lista)
    lista_pm = data_frame['SAMPLENAME'].unique()
    lista_descricao_metodo = data_frame['Description'].unique()

    # Cria um novo DataFrame para armazenar os resultados filtrados e cria um arquivo excel no openpyxl
    resultado_final = pd.DataFrame()
    workbook = Workbook()

    # Itera sobre as listas e adiciona os resultados filtrados ao novo DataFrame
    for descricao_metodo in lista_descricao_metodo:
        filtro = (data_frame['Description'] == descricao_metodo)
        tabela_merge = data_frame[filtro]

        # Adiciona os títulos das colunas específicas à primeira linha da planilha faz uma lista com as colunas desejadas
        if not tabela_merge.empty:

            #Caso o titulo "descricao_metodo" tiver mais de 31 caracteres ele pega somente os 3 primeiros caracteceres 
            if len(descricao_metodo) > 31:
                aba_titulo = descricao_metodo[:4]
            else:
                aba_titulo = descricao_metodo

            aba_descricao_metodo = workbook.create_sheet(title=aba_titulo)
            colunas_desejadas = ["SAMPLENAME", "SAMPDATE", "CASNUMBER_x", "ANALYTE", "Result", "UNITS", "Description", "Teste"]
            aba_descricao_metodo.append(colunas_desejadas)

            # Adiciona as linhas correspondentes às colunas desejadas
            for linha in tabela_merge[colunas_desejadas].itertuples(index=False):
                aba_descricao_metodo.append(linha)

    # Salva a tabela merge em "Resultado_Merge.xlsx"
    #tabela_merge.to_excel(r'P:\Vitor\Programacao\teste_CEIMIC\Resultado_Merge.xlsx', index=True)

    # Obtém a primeira aba, passando o indice [0] (worksheet)
    primeira_aba = workbook.worksheets[0]

    # Remove a primeira aba que esta vazia
    workbook.remove(primeira_aba)

    # Salva a tabela merge em "Resultado_Final_Com_Abas.xlsx"
    workbook.save(novo_caminho)

if __name__ == "__main__":
    main()