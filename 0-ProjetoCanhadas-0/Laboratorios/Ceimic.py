import pandas as pd
from openpyxl import Workbook


def main(uploaded_file, novo_caminho):
    #Leitor do Arquivo upado pelo usuario
    tabela = pd.read_excel(uploaded_file)

    #Banco de dados em excel da Ceimic 
    url ="https://github.com/TecnologiaServmar/ProjetoCanhadas/raw/main/Tabelas%20Consulta/Banco%20de%20Dados/banco%20de%20dados%20-%20CEIMIC.xlsx"

    #Leitor do banco de Dados
    tabela_bd = pd.read_excel(url)

    #Mesclar a parte de Analyte e Desciption com o Excel do Usuario junto com o Excle do Banco de Dados
    colunas_merge = ["ANALYTE", "Description"]

    #Pega todos os itens que têm em comum em relação à coluna "Análise" e "Descrição Método", junta em uma planilha só
    tabela_merge = pd.merge(tabela, tabela_bd, how='left', on=colunas_merge)

    #Formata a coluna SAMPDATE para o padrão desejado, no caso %d/%m/%Y %H:%M
    try:
    # Tenta converter a coluna SAMPDATE para o formato desejado
        tabela_merge['SAMPDATE'] = pd.to_datetime(tabela_merge['SAMPDATE'], format='%m/%d/%Y %H:%M').dt.strftime('%d/%m/%Y %H:%M')
    except Exception as e:
        # Caso ocorra algum erro, exibe uma mensagem (opcional) e segue o código
        print(f"Erro ao converter a coluna SAMPDATE: {e}")
        pass  # Continua a execução do código
    
    #O novo data_frame recebe os valores mesclados
    data_frame = tabela_merge

    #Pega da coluna "Descrição Método" todos as descrições dos métodos, pega os valores únicos e joga em uma lista
    lista_descricao_metodo = data_frame['Description'].unique()

    #Cria um novo DataFrame para armazenar os resultados filtrados e cria um arquivo excel no openpyxl
    workbook = Workbook()

    #Loop de repetição para filtrar e organizar o dataframe que sera gerado separado por abaas com os nomes da Description
    for descricao_metodo in lista_descricao_metodo:
        filtro = (data_frame['Description'] == descricao_metodo)
        tabela_merge = data_frame[filtro]

        #Adiciona os títulos das colunas específicas à primeira linha da planilha faz uma lista com as colunas desejadas
        if not tabela_merge.empty:

            #Caso o titulo "descricao_metodo" tiver mais de 31 caracteres ele pega somente os 4 primeiros caracteceres 
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

    #Obtém a primeira aba, passando o indice [0] (worksheet)
    primeira_aba = workbook.worksheets[0]

    #Remove a primeira aba que esta vazia
    workbook.remove(primeira_aba)

    #Salva a tabela merge em "Resultado_Final_Com_Abas.xlsx"
    workbook.save(novo_caminho)

if __name__ == "__main__":
    main()