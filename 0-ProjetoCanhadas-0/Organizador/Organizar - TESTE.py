import pandas as pd

def main():
    caminho_excel = r"C:\Users\henrique.canhadas\OneDrive - Servmar Ambientais\Documentos\Codigos\GitHub Tecnologia\ProjetoCanhadas\0-Excel de Exemplos-0\DADOS Vapor Solutions resultado.xlsx"
    excel = pd.ExcelFile(caminho_excel)
    data_frame_final = {}
    datas_amostra = {}  # Dicionário para armazenar datas das amostras para cada PM

    for sheet_name in excel.sheet_names:
        data_frame = pd.read_excel(caminho_excel, sheet_name)
        lista_pm = data_frame['SAMPLENAME'].unique()
        lista_analyte = data_frame['ANALYTE'].unique()

        # Inicializando o DataFrame com uma linha extra para as datas
        data_frame_tabelado = pd.DataFrame(index=range(len(lista_analyte) + 1), columns=['Parametro', 'CAS', 'Unidade'] + list(lista_pm))
        correspondencia_unidades = {}
        correspondencia_cas = {}

        for analyte in lista_analyte:
            filtro_analyte = data_frame[data_frame['ANALYTE'] == analyte]
            if not filtro_analyte.empty:
                correspondencia_unidades[analyte] = filtro_analyte['UNITS'].iloc[0]
                correspondencia_cas[analyte] = filtro_analyte['CASNUMBER_x'].iloc[0]

        for i, analyte in enumerate(lista_analyte, start=1):
            data_frame_tabelado.at[i, 'Parametro'] = analyte
            data_frame_tabelado.at[i, 'Unidade'] = correspondencia_unidades.get(analyte, '')
            data_frame_tabelado.at[i, 'CAS'] = correspondencia_cas.get(analyte, '')

        for pm in lista_pm:
            data_frame_pm = data_frame[data_frame['SAMPLENAME'] == pm]
            datas_amostra[pm] = str(data_frame_pm['SAMPDATE'].iloc[0])
            for i, analyte in enumerate(lista_analyte, start=1):
                resultado = data_frame_pm[data_frame_pm['ANALYTE'] == analyte]['Result'].values
                if resultado.size > 0:
                    data_frame_tabelado.at[i, pm] = resultado[0]
                else:
                    data_frame_tabelado.at[i, pm] = "n.a"

        # Aplicando as datas de coleta na primeira linha de cada aba
        for pm, data in datas_amostra.items():
            if pm in data_frame_tabelado.columns:
                data_frame_tabelado.at[0, pm] = data

        # Inclui a lógica para tratar especificamente as linhas com '%'
        # [Sua lógica aqui]

        # Verificar se a coluna 'Unidade' contém '%'
        if '%' in data_frame_tabelado['Unidade'].values:
            df_percentagem = data_frame_tabelado[data_frame_tabelado['Unidade'] == '%']
            if '%' not in data_frame_final:
                data_frame_final['%'] = pd.DataFrame(index=range(1), columns=data_frame_tabelado.columns)
            for index, row in df_percentagem.iterrows():
                last_row_index = len(data_frame_final['%'])
                data_frame_final['%'].loc[last_row_index + 2] = row
            data_frame_tabelado = data_frame_tabelado[data_frame_tabelado['Unidade'] != '%']
            for coluna_atual in data_frame_tabelado.columns:
                if coluna_atual in data_frame_final['%'].columns:
                    valor_segunda_linha = data_frame_tabelado[coluna_atual].iloc[0]
                    data_frame_final['%'].loc[0, coluna_atual] = valor_segunda_linha

        data_frame_final[sheet_name] = data_frame_tabelado

    if '%' in data_frame_final:
        data_frame_final['%'] = data_frame_final.pop('%')

    with pd.ExcelWriter(r"C:\Users\henrique.canhadas\OneDrive - Servmar Ambientais\Documentos\Codigos\GitHub Tecnologia\ProjetoCanhadas\0-Excel de Exemplos-0\DADOS Vapor Solutions testetestes.xlsx") as writer:
        for sheet_name, df in data_frame_final.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    main()
