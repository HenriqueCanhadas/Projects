import pandas as pd
from openpyxl import Workbook

def main(uploaded_file,novo_caminho):
    #uploaded_file = r'C:\Users\henrique.canhadas\OneDrive - Servmar Ambientais\Documentos\Codigos\GitHub Tecnologia\ProjetoCanhadas\0-Excel de Exemplos-0\0.1 Arquivo Teste Description.xlsx'
    #novo_caminho = r'C:\Users\henrique.canhadas\Desktop\Teste\0.2 Arquivo TESTE.xlsx'
    #banco_dados = r'C:\Users\henrique.canhadas\Desktop\Teste\banco_dados.xlsx'  

    banco_dados= "https://github.com/TecnologiaServmar/ProjetoCanhadas/blob/main/Tabelas%20Consulta/Tabelas/banco_dados.xlsx"
    # Leitor do Arquivo upado pelo usuário
    tabela = pd.read_excel(uploaded_file)
    
    # Banco de dados em Excel
    banco_dados_df = pd.read_excel(banco_dados)

    # Verifica se a coluna "Description" existe no arquivo do usuário
    if "Description" not in tabela.columns:
        # Se não existir, adiciona a coluna "Description" com base no banco de dados
        tabela["Description"] = None
        
        # Iterar sobre cada valor da coluna "ANALYTE" no arquivo do usuário
        for index, analyte in tabela["ANALYTE"].items():
            # Procurar pelo valor correspondente na coluna "Nome" do Banco de Dados
            grupo_correspondente = banco_dados_df.loc[banco_dados_df["Nome"] == analyte, "Grupo"]
            
            # Se encontrar um valor correspondente, inserir na nova coluna "Description"
            if not grupo_correspondente.empty:
                tabela.at[index, "Description"] = grupo_correspondente.values[0]

        # Substituir valores None na coluna "Description" por "Não Localizado"
        tabela["Description"] = tabela["Description"].fillna("Não Localizado")

    # Verificar se a coluna 'CASNUMBER' existe no arquivo do usuário
    if "CASNUMBER" not in tabela.columns:
        # Se não existir, adiciona a coluna 'CASNUMBER' com valores iniciais como None
        tabela["CASNUMBER_x"] = None
        
        # Iterar sobre cada valor da coluna "ANALYTE" no arquivo do usuário
        for index, analyte in tabela["ANALYTE"].items():
            # Procurar pelo valor correspondente na coluna "Nome" do banco de dados
            cas_correspondente = banco_dados_df.loc[banco_dados_df["Nome"] == analyte, "CAS"]
            
            # Se encontrar um valor correspondente, inserir na nova coluna "CASNUMBER"
            if not cas_correspondente.empty:
                tabela.at[index, "CASNUMBER_x"] = cas_correspondente.values[0]
        # Substituir valores None na coluna "Description" por "Não Localizado"
        tabela["CASNUMBER_x"] = tabela["CASNUMBER_x"].fillna("n.e")

    # Mesclar a parte de Analyte e Description com o Excel do Usuário junto com o Excel do Banco de Dados
    colunas_merge_usuario = ["ANALYTE", "Description"]
    colunas_merge_bd = ["Nome", "Grupo"]
    
    # Pega todos os itens que têm em comum em relação à coluna "ANALYTE" e "Description", junta em uma planilha só
    tabela_merge = pd.merge(tabela, banco_dados_df, how='left', left_on=colunas_merge_usuario, right_on=colunas_merge_bd)
    
    # Formata a coluna SAMPDATE para o padrão desejado, no caso %d/%m/%Y %H:%M
    try:
        # Tenta converter a coluna SAMPDATE para o formato desejado
        tabela_merge['SAMPDATE'] = pd.to_datetime(tabela_merge['SAMPDATE'], format='%m/%d/%Y %H:%M').dt.strftime('%d/%m/%Y %H:%M')
    except Exception as e:
        # Caso ocorra algum erro, exibe uma mensagem (opcional) e segue o código
        print(f"Erro ao converter a coluna SAMPDATE: {e}")
        pass  # Continua a execução do código
    
    # O novo data_frame recebe os valores mesclados
    data_frame = tabela_merge
    
    # Pega da coluna "Description" todos as descrições dos métodos, pega os valores únicos e joga em uma lista
    lista_descricao_metodo = data_frame['Description'].unique()
    
    # Cria um novo DataFrame para armazenar os resultados filtrados e cria um arquivo Excel no openpyxl
    workbook = Workbook()
    
    # Loop de repetição para filtrar e organizar o dataframe que será gerado separado por abas com os nomes da Description
    for descricao_metodo in lista_descricao_metodo:
        filtro = (data_frame['Description'] == descricao_metodo)
        tabela_merge = data_frame[filtro]
        
        # Adiciona os títulos das colunas específicas à primeira linha da planilha faz uma lista com as colunas desejadas
        if not tabela_merge.empty:
            # Caso o título "descricao_metodo" tiver mais de 31 caracteres ele pega somente os 4 primeiros caracteres
            if len(descricao_metodo) > 31:
                aba_titulo = descricao_metodo[:4]
            else:
                aba_titulo = descricao_metodo
            
            aba_descricao_metodo = workbook.create_sheet(title=aba_titulo)
            colunas_desejadas = ["SAMPLENAME", "SAMPDATE", "CASNUMBER_x", "ANALYTE", "Result", "UNITS", "Description"]
            aba_descricao_metodo.append(colunas_desejadas)
            
            # Adiciona as linhas correspondentes às colunas desejadas
            for linha in tabela_merge[colunas_desejadas].itertuples(index=False):
                aba_descricao_metodo.append(linha)
    
    # Obtém a primeira aba, passando o índice [0] (worksheet)
    primeira_aba = workbook.worksheets[0]
    
    # Remove a primeira aba que está vazia
    workbook.remove(primeira_aba)
    
    # Salva a tabela merge em "Resultado_Final_Com_Abas.xlsx"
    workbook.save(novo_caminho)

if __name__ == "__main__":
    main()

