import streamlit as st
import os
import pandas as pd

#Importar Codigos de cada laboratorio
from Laboratorios import Ceimic
from Laboratorios import Outros
from Laboratorios import ALS
from Laboratorios import Eurofins
from Laboratorios import VaporSolutions

#Importar Codigo para Organizar
from Organizador import Organizar

#Importar Codigos para Analisar os valores
from Analizador import Analise2
from Analizador import Analise3

#-------------------------------------------------------------------------------------------------------
#Funções para verificar as opçoes ente 2 Valores Orientadores ou 3 Valores Orientadores:
#Para 3 Valores Orientadores
def abrir_radiobutton_modal_3_valores(contador):
    #Inicia as ordens de variaveis com valor none
    valor_primario = None
    ordem_planilhas = None
    valor_secundario = None
    ordem_planilhas2 = None
    valor_terceario = None
    ordem_planilhas3 = None

    #Separar 3 colunas para exibição no site 
    col1, col2, col3 = st.columns(3,vertical_alignment="top")

    #Na primeira coluna:
    with col1:
        # Gerar uma chave única para o widget radio, para parte do Streamlit
        chave_radio = f"lab_radio_{contador}"
        #Opções do primeiro Valor Orientador a ser comparado
        escolha1 = st.radio("Primeiro Valor de Referencia:", ["Cetesb", "US EPA", "Lista Holandesa", "Conama-420"], key=chave_radio, index=None)
        #Separação de opções ente Cetesb, US EPA, Lista Holandesa, Conama-420 na Primeira comparação
        if escolha1 == "Cetesb":
            ordem_planilhas = "c"
            chave_radio_cetesb = f"cetesb_radio_{contador}"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Cetesb", ["Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_cetesb, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_primario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 2
            elif escolha_matriz == "Solo Industrial":
                valor_primario = 3
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 4
        elif escolha1 == "US EPA":
            chave_radio_epa = f"epa_radio_{contador}"
            ordem_planilhas = "e"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da US EPA", ["Res Solo", "Res Ar", "Solo para GW 1123", "Água Subterrânea", "Ind Solo", "Ind Air"], key=chave_radio_epa, index=None)
            if escolha_matriz == "Res Solo":
                valor_primario = 1
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 2
            elif escolha_matriz == "Res Ar":
                valor_primario = 3
            elif escolha_matriz == "Solo para GW 1123":
                valor_primario = 4
            elif escolha_matriz == "Ind Solo":
                valor_primario = 5
            elif escolha_matriz == "Ind Air":
                valor_primario = 6
        elif escolha1 == "Lista Holandesa":
            chave_radio_listaholandesa = f"listaholandesa_radio_{contador}"
            ordem_planilhas = "l"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Lista Holandesa", ["Solo Agrícola", "Solo Residencial", "Água Subterrânea"], key=chave_radio_listaholandesa, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_primario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 2
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 3
        elif escolha1 == "Conama-420":
            chave_radio_conama = f"conama_radio_{contador}"
            ordem_planilhas = "o"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Conama-420", ["Solo Prevenção", "Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_conama, index=None)
            if escolha_matriz == "Solo Prevenção":
                valor_primario = 1
            elif escolha_matriz == "Solo Agrícola":
                valor_primario = 2
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 3
            elif escolha_matriz == "Solo Industrial":
                valor_primario = 4
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 5
    
    #Atualiza a chave do Streamlit
    contador = 3

    #Na Segunda coluna:
    with col2:
        # Gerar uma chave única para o widget radio, para parte do Streamlit
        chave_radio = f"lab_radio_{contador}"
        #Opções do segundo Valor Orientador a ser comparado
        escolha = st.radio("Segundo Valor de Referencia:", ["Cetesb", "US EPA", "Lista Holandesa", "Conama-420"], key=chave_radio, index=None)
        #Separação de opções ente Cetesb, US EPA, Lista Holandesa, Conama-420 na Segunda comparação
        if escolha == "Cetesb":
            ordem_planilhas2 = "c"
            chave_radio_cetesb = f"cetesb_radio_{contador}"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Cetesb", ["Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_cetesb,index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_secundario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 2
            elif escolha_matriz == "Solo Industrial":
                valor_secundario = 3
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 4
        elif escolha == "US EPA":
            chave_radio_epa = f"epa_radio_{contador}"
            ordem_planilhas2 = "e"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da US EPA", ["Res Solo", "Res Ar", "Solo para GW 1123", "Água Subterrânea", "Ind Solo", "Ind Air"], key=chave_radio_epa, index=None)
            if escolha_matriz == "Res Solo":
                valor_secundario = 1
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 2
            elif escolha_matriz == "Res Ar":
                valor_secundario = 3
            elif escolha_matriz == "Solo para GW 1123":
                valor_secundario = 4
            elif escolha_matriz == "Ind Solo":
                valor_secundario = 5
            elif escolha_matriz == "Ind Air":
                valor_secundario = 6
        elif escolha == "Lista Holandesa":
            chave_radio_listaholandesa = f"listaholandesa_radio_{contador}"
            ordem_planilhas2 = "l"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Lista Holandesa", ["Solo Agrícola", "Solo Residencial", "Água Subterrânea"], key=chave_radio_listaholandesa, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_secundario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 2
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 3
        elif escolha == "Conama-420":
            chave_radio_conama = f"conama_radio_{contador}"
            ordem_planilhas2 = "o"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Conama-420", ["Solo Prevenção", "Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_conama, index=None)
            if escolha_matriz == "Solo Prevenção":
                valor_secundario = 1
            elif escolha_matriz == "Solo Agrícola":
                valor_secundario = 2
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 3
            elif escolha_matriz == "Solo Industrial":
                valor_secundario = 4
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 5
    
    #Atualiza a chave do Streamlit
    contador = 4

    #Na Terceira coluna:
    with col3:
        # Gerar uma chave única para o widget radio, para parte do Streamlit
        chave_radio = f"lab_radio_{contador}"
        #Opções do terceiro Valor Orientador a ser comparado
        escolha2 = st.radio("Terceiro Valor de Referencia:", ["Cetesb", "US EPA", "Lista Holandesa", "Conama-420"], key=chave_radio, index=None)
        #Separação de opções ente Cetesb, US EPA, Lista Holandesa, Conama-420 na Terceia comparação
        if escolha2 == "Cetesb":
            ordem_planilhas3 = "c"
            chave_radio_cetesb = f"cetesb_radio_{contador}"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Cetesb", ["Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_cetesb, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_terceario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_terceario = 2
            elif escolha_matriz == "Solo Industrial":
                valor_terceario = 3
            elif escolha_matriz == "Água Subterrânea":
                valor_terceario = 4
        elif escolha2 == "US EPA":
            chave_radio_epa = f"epa_radio_{contador}"
            ordem_planilhas3 = "e"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da US EPA", ["Res Solo", "Res Ar", "Solo para GW 1123", "Água Subterrânea", "Ind Solo", "Ind Air"], key=chave_radio_epa, index=None)
            if escolha_matriz == "Res Solo":
                valor_terceario = 1
            elif escolha_matriz == "Água Subterrânea":
                valor_terceario = 2
            elif escolha_matriz == "Res Ar":
                valor_terceario = 3
            elif escolha_matriz == "Solo para GW 1123":
                valor_terceario = 4
            elif escolha_matriz == "Ind Solo":
                valor_terceario = 5
            elif escolha_matriz == "Ind Air":
                valor_terceario = 6
        elif escolha2 == "Lista Holandesa":
            chave_radio_listaholandesa = f"listaholandesa_radio_{contador}"
            ordem_planilhas3 = "l"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Lista Holandesa", ["Solo Agrícola", "Solo Residencial", "Água Subterrânea"], key=chave_radio_listaholandesa, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_terceario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_terceario = 2
            elif escolha_matriz == "Água Subterrânea":
                valor_terceario = 3
        elif escolha2 == "Conama-420":
            chave_radio_conama = f"conama_radio_{contador}"
            ordem_planilhas3 = "o"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Conama-420", ["Solo Prevenção", "Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_conama, index=None)
            if escolha_matriz == "Solo Prevenção":
                valor_terceario = 1
            elif escolha_matriz == "Solo Agrícola":
                valor_terceario = 2
            elif escolha_matriz == "Solo Residencial":
                valor_terceario = 3
            elif escolha_matriz == "Solo Industrial":
                valor_terceario = 4
            elif escolha_matriz == "Água Subterrânea":
                valor_terceario = 5

    #Retorna os valores necessarios selecionados pelo usuario 
    return valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3   
#Para 2 Valores Orientadores   
def abrir_radiobutton_modal_2_valores(contador):
    #Inicia as ordens de variaveis com valor none
    valor_primario = None
    ordem_planilhas = None
    valor_secundario = None
    ordem_planilhas2 = None

    #Separar 2 colunas para exibição no site
    col1, col2 = st.columns(2,vertical_alignment="top")
    
    #Na Primeira coluna:
    with col1:
        # Gerar uma chave única para o widget radio, para parte do Streamlit
        chave_radio = f"lab_radio_{contador}"
        #Opções do primeiro Valor Orientador a ser comparado
        escolha1 = st.radio("Primeiro Valor de Referencia:", ["Cetesb", "US EPA", "Lista Holandesa", "Conama-420"], key=chave_radio, index=None)
        #Separação de opções ente Cetesb, US EPA, Lista Holandesa, Conama-420 na Primeira comparação
        if escolha1 == "Cetesb":
            ordem_planilhas = "c"
            chave_radio_cetesb = f"cetesb_radio_{contador}"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Cetesb", ["Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_cetesb, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_primario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 2
            elif escolha_matriz == "Solo Industrial":
                valor_primario = 3
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 4
        elif escolha1 == "US EPA":
            chave_radio_epa = f"epa_radio_{contador}"
            ordem_planilhas = "e"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da US EPA", ["Res Solo", "Res Ar", "Solo para GW 1123", "Água Subterrânea", "Ind Solo", "Ind Air"], key=chave_radio_epa, index=None)
            if escolha_matriz == "Res Solo":
                valor_primario = 1
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 2
            elif escolha_matriz == "Res Ar":
                valor_primario = 3
            elif escolha_matriz == "Solo para GW 1123":
                valor_primario = 4
            elif escolha_matriz == "Ind Solo":
                valor_primario = 5
            elif escolha_matriz == "Ind Air":
                valor_primario = 6
        elif escolha1 == "Lista Holandesa":
            chave_radio_listaholandesa = f"listaholandesa_radio_{contador}"
            ordem_planilhas = "l"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Lista Holandesa", ["Solo Agrícola", "Solo Residencial", "Água Subterrânea"], key=chave_radio_listaholandesa, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_primario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 2
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 3
        elif escolha1 == "Conama-420":
            chave_radio_conama = f"conama_radio_{contador}"
            ordem_planilhas = "o"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Conama-420", ["Solo Prevenção", "Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_conama, index=None)
            if escolha_matriz == "Solo Prevenção":
                valor_primario = 1
            elif escolha_matriz == "Solo Agrícola":
                valor_primario = 2
            elif escolha_matriz == "Solo Residencial":
                valor_primario = 3
            elif escolha_matriz == "Solo Industrial":
                valor_primario = 4
            elif escolha_matriz == "Água Subterrânea":
                valor_primario = 5

    #Atualiza a chave do Streamlit
    contador = 3
    
    #Na Segunda coluna:
    with col2:
        # Gerar uma chave única para o widget radio, para parte do Streamlit
        chave_radio = f"lab_radio_{contador}"
        #Opções do segundo Valor Orientador a ser comparado
        escolha = st.radio("Segundo Valor de Referencia:", ["Cetesb", "US EPA", "Lista Holandesa", "Conama-420"], key=chave_radio, index=None)
        #Separação de opções ente Cetesb, US EPA, Lista Holandesa, Conama-420 na Segunda comparação
        if escolha == "Cetesb":
            ordem_planilhas2 = "c"
            chave_radio_cetesb = f"cetesb_radio_{contador}"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Cetesb", ["Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_cetesb,index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_secundario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 2
            elif escolha_matriz == "Solo Industrial":
                valor_secundario = 3
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 4
        elif escolha == "US EPA":
            chave_radio_epa = f"epa_radio_{contador}"
            ordem_planilhas2 = "e"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da US EPA", ["Res Solo", "Res Ar", "Solo para GW 1123", "Água Subterrânea", "Ind Solo", "Ind Air"], key=chave_radio_epa, index=None)
            if escolha_matriz == "Res Solo":
                valor_secundario = 1
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 2
            elif escolha_matriz == "Res Ar":
                valor_secundario = 3
            elif escolha_matriz == "Solo para GW 1123":
                valor_secundario = 4
            elif escolha_matriz == "Ind Solo":
                valor_secundario = 5
            elif escolha_matriz == "Ind Air":
                valor_secundario = 6
        elif escolha == "Lista Holandesa":
            chave_radio_listaholandesa = f"listaholandesa_radio_{contador}"
            ordem_planilhas2 = "l"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Lista Holandesa", ["Solo Agrícola", "Solo Residencial", "Água Subterrânea"], key=chave_radio_listaholandesa, index=None)
            if escolha_matriz == "Solo Agrícola":
                valor_secundario = 1
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 2
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 3
        elif escolha == "Conama-420":
            chave_radio_conama = f"conama_radio_{contador}"
            ordem_planilhas2 = "o"
            escolha_matriz = st.radio("Selecione a matriz e/ou o cenário ambiental da Conama-420", ["Solo Prevenção", "Solo Agrícola", "Solo Residencial", "Solo Industrial", "Água Subterrânea"], key=chave_radio_conama, index=None)
            if escolha_matriz == "Solo Prevenção":
                valor_secundario = 1
            elif escolha_matriz == "Solo Agrícola":
                valor_secundario = 2
            elif escolha_matriz == "Solo Residencial":
                valor_secundario = 3
            elif escolha_matriz == "Solo Industrial":
                valor_secundario = 4
            elif escolha_matriz == "Água Subterrânea":
                valor_secundario = 5

    #Retorna os valores necessarios selecionados pelo usuario 
    return valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2
#-------------------------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------
#Funções para acionar os arquivos em ordem escolhida pelo usuario para o Tabelamaneto:
#Para 3 Valores Orientadores
def carregar_analise_3_valores(uploaded_file, novo_caminho, escolha, quantidade_analise, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3):
    #Variavel para a barra de Progresso
    progresso = 0
    
    if quantidade_analise == 3 and escolha == "Ceimic":
        #Rodar o Codigo da Ceimic
        Ceimic.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 3 Analises
        Analise3.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3)
        progresso += 34
        yield progresso

    elif quantidade_analise == 3 and escolha == "ALS":
        #Rodar o Codigo da ALS
        ALS.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 3 Analises
        Analise3.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3)
        progresso += 34
        yield progresso

    elif quantidade_analise == 3 and escolha == "EuroFins":
        #Rodar o Codigo da Eurofins
        Eurofins.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 3 Analises
        Analise3.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3)
        progresso += 34
        yield progresso

    elif quantidade_analise == 3 and escolha == "Vapor Solutions":
        #Rodar o Codigo da Vapor Solutions
        VaporSolutions.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 3 Analises
        Analise3.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3)
        progresso += 34
        yield progresso

    elif quantidade_analise == 3 and escolha == "Outros":
        #Rodar o Codigo da Outros
        Outros.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 3 Analises
        Analise3.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3)
        progresso += 34
        yield progresso
#Para 2 Valores Orientadores  
def carregar_analise_2_valores(uploaded_file, novo_caminho, escolha, quantidade_analise, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2):
    #Variavel para a barra de Progresso
    progresso = 0

    if quantidade_analise == 2 and escolha == "Ceimic":
        #Rodar o Codigo Ceimic
        Ceimic.main(uploaded_file, novo_caminho)
        progresso += 25
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 40
        yield progresso
        #Rodar o Codigo de 2 Analises
        Analise2.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2)
        progresso += 35
        yield progresso

    elif quantidade_analise == 2 and escolha == "ALS":
        #Rodar o Codigo ALS
        ALS.main(uploaded_file, novo_caminho)
        progresso += 25
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 40
        yield progresso
        #Rodar o Codigo de 2 Analises
        Analise2.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2)
        progresso += 35
        yield progresso

    elif quantidade_analise == 2 and escolha == "EuroFins":
        #Rodar o Codigo Eurofins
        Eurofins.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 2 Analises
        Analise2.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2)
        progresso += 34
        yield progresso

    elif quantidade_analise == 2 and escolha == "Vapor Solutions":
        #Rodar o Codigo Vapor Solutions
        VaporSolutions.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 2 Analises
        Analise2.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2)
        progresso += 34
        yield progresso

    elif quantidade_analise == 2 and escolha == "Outros":
        #Rodar o Codigo Outros
        Outros.main(uploaded_file, novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo Organizar
        Organizar.main(novo_caminho)
        progresso += 33
        yield progresso
        #Rodar o Codigo de 2 Analises
        Analise2.main(novo_caminho, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2)
        progresso += 34
        yield progresso
#-------------------------------------------------------------------------------------------------------

def main():
    #Configurações da pagina
    st.markdown(
        """
        <style>
        html, body, [class*="View"] {
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
        }
        .reportview-container {
            margin: 10px;
            flex: 1;
        }
        .main .block-container {
            width: calc(100% - 15px);  /* Subtrai as margens */
            padding: 0;
            margin: -90px;
        }

        /* Estilos responsivos para diferentes tamanhos de tela */
        @media (max-width: 768px) {
            .reportview-container {
                margin: 10px; /* Menor margem para telas menores */
            }
            .main .block-container {
                width: calc(100% - 50px); /* Ajusta a largura para telas menores */
            }
        }
        </style>
        <div style="text-align: center; background-color:#d1d1e4; border-radius: 20px; padding: 10px;">
            <span style="color: black; font-size: 70px; font-weight: bold;">Projeto Canhadas</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Escondendo o menu, rodapé e cabeçalho
    hide_menu_style = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        </style>
        """
    st.markdown(hide_menu_style, unsafe_allow_html=True)

    #Separar pagina em duas Colunas
    left_column, right_column = st.columns(2,vertical_alignment="top")
    
    #Na Coluna Esquerda:
    with left_column:
        #Fazer o Upload do Arquivo Excel para o programa
        uploaded_file = st.file_uploader("Carregue seu arquivo Excel:", type=["xlsx"], key="excel_uploader_1",label_visibility="visible")
        diretório=""
        #Se o diretorio e arquivo forem verdadeiros:
        if uploaded_file:
            df = pd.read_excel(uploaded_file)
            st.write("Dados do arquivo Excel:")
            st.dataframe(df)
            diretório, _ = os.path.split(uploaded_file.name)

        #Na Coluna Direita:
        with right_column:
            
            #Variavel para criar um novo nome para o arquivo
            nome_arquivo = st.text_input("1° Digite o nome do novo arquivo:", key="nome_arquivo")
            if nome_arquivo:
                novo_caminho = os.path.join(diretório, nome_arquivo + ".xlsx")
            else:
                novo_caminho = os.path.join(diretório, ".xlsx")
                st.warning("Por favor, insira um nome para o novo arquivo.")

            # Verifica se um nome de arquivo foi inserido
            if nome_arquivo:
                #Separar pagina em duas Colunas
                col1, col2 = st.columns(2)
                #Na coluna 1 deve ser escolhido o Laboratório
                with col1:
                    escolha = st.radio("2° Escolha qual laboratório a análise deve ser feita:", ["Ceimic","ALS", "EuroFins", "Vapor Solutions", "Outros"], key="escolha_laboratorio_1")
                #Na coluna 2 deve ser escolhido a quantidede de Valores Orientadores
                with col2:
                    # Mapeamento de opções de texto para valores numéricos
                    quantidade_analise_options = {"2 Valores Orientadores": 2, "3 Valores Orientadores": 3}
                    # Usando os textos descritivos no widget, mas obtendo os valores numéricos quando necessário
                    quantidade_analise_texto = st.radio("3° Escolha a quantidade de Valores Orientadores:", list(quantidade_analise_options.keys()), key="quantidade_analise_1", index=None)
                # Verificar se a chave existe no dicionário antes de acessá-la
                if quantidade_analise_texto in quantidade_analise_options:
                    quantidade_analise = quantidade_analise_options[quantidade_analise_texto]
                # Ou outro valor padrão que faça sentido no seu código
                else:
                    quantidade_analise = None

            # Oculta as opções seguintes se o nome do arquivo não for inserido
            else:
                quantidade_analise = None
            
            #Se o usuario escolheu 2 Valores Orientadores sao iniciadas as variaveis necessarias com valores none e é chamada a função abrir_radiobutton_modal_2_valores
            if quantidade_analise == 2:
                valor_primario = None
                ordem_planilhas = None
                valor_secundario = None
                ordem_planilhas2 = None
                valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2 = abrir_radiobutton_modal_2_valores(2)

            #Se o usuario escolheu 3 Valores Orientadores sao iniciadas as variaveis necessarias com valores none e é chamada a função abrir_radiobutton_modal_3_valores
            elif quantidade_analise == 3:
                valor_primario = None
                ordem_planilhas = None
                valor_secundario = None
                ordem_planilhas2 = None
                valor_terceario = None
                ordem_planilhas3 = None
                col1, col2 = st.columns(2)
                valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3 = abrir_radiobutton_modal_3_valores(2)

            #Verifica se os valores retornados da função abrir_radiobutton_modal_2_valores ou abrir_radiobutton_modal_3_valores tem algum valor ainda como none
            if (quantidade_analise == 2 and (valor_primario is not None and ordem_planilhas is not None and valor_secundario is not None and ordem_planilhas2 is not None)) or (quantidade_analise == 3 and (valor_primario is not None and ordem_planilhas is not None and valor_secundario is not None and ordem_planilhas2 is not None and valor_terceario is not None and ordem_planilhas3 is not None)):
                
                #Imprimi uma linha cinza claro na janela
                st.divider()

                #Se o botão Fazer Análise for precisonado
                if st.button("Fazer Análise", type="primary"):
                    #Inicia a barra de progesso zerada
                    progresso_placeholder = st.empty()
                    progresso_bar = progresso_placeholder.progress(0)

                    #Se a quantidade de Valores Orientadores for igual a 2
                    if quantidade_analise == 2:
                        #Chamada a função carregar_analise_2_valores para fazer o Tabelamento dos dados
                        for progresso in carregar_analise_2_valores(uploaded_file, novo_caminho, escolha, quantidade_analise, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2):
                            progresso_bar.progress(progresso)

                    #Se a quantidade de Valores Orientadores for igual a 3
                    elif quantidade_analise == 3:
                        #Chamada a função carregar_analise_3_valores para fazer o Tabelamento dos dados
                        for progresso in carregar_analise_3_valores(uploaded_file, novo_caminho, escolha, quantidade_analise, valor_primario, ordem_planilhas, valor_secundario, ordem_planilhas2, valor_terceario, ordem_planilhas3):
                            progresso_bar.progress(progresso)

                    #Chama a função de download para que o arquivo Tabelado possa ser baixado
                    download_excel(novo_caminho)

#Função para fazer o download do arquivo Tabelado
def download_excel(novo_caminho):
    #Configurar o nome do arquivo para download
    nome_arquivo = os.path.basename(novo_caminho)
    
    with open(novo_caminho, "rb") as file:
        # Ler os bytes do arquivo
        file_bytes = file.read()
    
    #Configurar o botão de download
    with st.spinner("Baixando Excel..."):
        st.success("Análise concluída. Clique abaixo para baixar.")
        st.download_button(
            label="Baixar Resultado da Análise",
            data=file_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
               
if __name__ == "__main__":
    main()