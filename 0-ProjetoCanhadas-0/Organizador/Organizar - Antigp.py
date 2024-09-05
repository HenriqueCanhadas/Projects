import pandas as pd

def main(novo_caminho):
    # Ler o arquivo Excel
    caminho_excel = novo_caminho
    excel = pd.ExcelFile(caminho_excel)

    # Criar um dicionário para armazenar os DataFrames de cada aba
    data_frame_final = {}

    # Loop para verificar cada aba no Excel
    for sheet_name in excel.sheet_names:
        data_frame = pd.read_excel(caminho_excel, sheet_name)

        lista_pm = data_frame['SAMPLENAME'].unique()
        lista_analyte = data_frame['ANALYTE'].unique()

        data_frame_tabelado = pd.DataFrame(index=range(len(lista_pm) + 2), columns=lista_pm)
        data_frame_tabelado.insert(0, 'Parametro', '')  
        data_frame_tabelado.insert(1, 'CAS', '')  
        data_frame_tabelado.insert(2, 'Unidade', '')  

        correspondencia_unidades = {}
        correspondencia_cas = {}

        for coluna in lista_pm:
            for linha, value in enumerate(lista_analyte):
                data_frame_tabelado.at[linha + 1, 'Parametro'] = value
                data_frame_tabelado.at[linha + 1, 'Unidade'] = correspondencia_unidades.get(value, '')
                data_frame_tabelado.at[linha + 1, 'CAS'] = correspondencia_cas.get(value, '')

            valores_totais = data_frame.loc[data_frame['SAMPLENAME'] == coluna, 'ANALYTE'].nunique()

            for linha in range(valores_totais):
                if data_frame['ANALYTE'].unique()[linha] not in correspondencia_unidades:
                    valor_unidades = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == data_frame['ANALYTE'].unique()[linha]), 'UNITS'].iloc[0]
                    correspondencia_unidades[data_frame['ANALYTE'].unique()[linha]] = valor_unidades

                if data_frame['ANALYTE'].unique()[linha] not in correspondencia_cas:
                    valor_cas = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == data_frame['ANALYTE'].unique()[linha]), 'CASNUMBER_x'].iloc[0]
                    correspondencia_cas[data_frame['ANALYTE'].unique()[linha]] = valor_cas

            if coluna in data_frame['SAMPLENAME'].values:
                data_correspondente = data_frame.loc[data_frame['SAMPLENAME'] == coluna, 'SAMPDATE'].iloc[0]
                data_frame_tabelado.at[0, coluna] = data_correspondente

        for coluna in lista_pm:
            if coluna in data_frame['SAMPLENAME'].values:
                for item, valor in enumerate(lista_analyte):
                    if valor in data_frame_tabelado['Parametro'].values:
                        linha_parametro = data_frame_tabelado.index[data_frame_tabelado['Parametro'] == valor][0]
                        resultado_correspondente = data_frame.loc[(data_frame['SAMPLENAME'] == coluna) & (data_frame['ANALYTE'] == valor), 'Result']
                        if not resultado_correspondente.empty:
                            resultado_correspondente = resultado_correspondente.iloc[0]
                            data_frame_tabelado.at[linha_parametro, coluna] = resultado_correspondente
                        else:
                            data_frame_tabelado.at[linha_parametro, coluna] = "n.a"

        # Verificar se a coluna 'Unidade' contém '%'
        if '%' in data_frame_tabelado['Unidade'].values:
            # Filtrar as linhas que contêm '%' na coluna 'Unidade'
            df_percentagem = data_frame_tabelado[data_frame_tabelado['Unidade'] == '%']

            # Criar a aba '%' se ainda não existir
            if '%' not in data_frame_final:
                # Criar a aba '%' com duas linhas vazias no início
                data_frame_final['%'] = pd.DataFrame(index=range(1), columns=data_frame_tabelado.columns)

            # Adicionar as linhas correspondentes a '%' no final da aba a partir da terceira linha
            for index, row in df_percentagem.iterrows():
                # Adicionar a linha inteira da aba atual a partir da terceira linha da nova aba '%'
                last_row_index = len(data_frame_final['%'])
                data_frame_final['%'].loc[last_row_index + 2] = pd.Series(row, index=data_frame_final['%'].columns)

            # Remover as linhas correspondentes da aba original
            data_frame_tabelado = data_frame_tabelado[data_frame_tabelado['Unidade'] != '%']

            # Comparar os nomes das colunas na aba atual com a aba '%' e realizar a operação necessária
            for coluna_atual in data_frame_tabelado.columns:
                # Verificar se o nome da coluna na aba atual está presente na aba '%'
                if coluna_atual in data_frame_final['%'].columns:
                    # Pegar o valor da segunda linha da coluna atual e inserir na aba '%'
                    valor_segunda_linha = data_frame_tabelado[coluna_atual].iloc[0]
                    # Modificação: Usar .loc para evitar atribuição encadeada
                    data_frame_final['%'].loc[0, coluna_atual] = valor_segunda_linha


        # Adiciona a aba original (possivelmente modificada) ao dicionário
        data_frame_final[sheet_name] = data_frame_tabelado

    # Reorganizar o dicionário para mover a aba '%' para o final
    if '%' in data_frame_final:
        data_frame_final['%'] = data_frame_final.pop('%')

    # Salvar o resultado em um novo arquivo Excel
    with pd.ExcelWriter(caminho_excel) as writer:
        for sheet_name, df in data_frame_final.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    main()