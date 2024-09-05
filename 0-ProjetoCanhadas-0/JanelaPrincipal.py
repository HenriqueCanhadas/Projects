import streamlit as st
import streamlit_authenticator as stauth
import yaml
import requests
import time

#Configurações da pagina
st.set_page_config(layout="wide")
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
            margin: 10px;
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
        <div style="text-align: Left">
            <span style="font-size: 50px; font-weight: bold;">SERVMAR</span>
        </div>
        """,
        unsafe_allow_html=True
)

#Função para exceutar o Projeto Canhadas
def projeto():
    import ProjetoCanhadas
    ProjetoCanhadas.main()

#Função para autenticar o usuário
def authenticate():
    if "authentication_status" not in st.session_state:
        st.session_state["authentication_status"] = None

    config_url = 'https://raw.githubusercontent.com/TecnologiaServmar/ProjetoCanhadas/main/config.yaml'
    try:
        response = requests.get(config_url)
        if response.status_code == 200:
            config = yaml.load(response.content, Loader=yaml.SafeLoader)
        else:
            st.error("Falha ao carregar o arquivo de configuração.")
            return
    except Exception as e:
        st.error(f"Erro ao buscar o arquivo de configuração: {e}")
        return

    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days']
    )
    authenticator.login()

#Função para mensagem temporária de sucessp para o login
def display_temporary_success_message():
    # Aguarda 5 segundos
    time.sleep(3)
    # Exibe a mensagem de sucesso
    success_message = st.success("Login Feito, Seja Bem Vindo!")
    # Aguarda 5 segundos
    time.sleep(3)
    # Remove a mensagem de sucesso
    success_message.empty()


def main():

    hide_menu_style = """
            <style>
            #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
            </style>
            """
    st.markdown(hide_menu_style, unsafe_allow_html=True)

    # Verifica se o status de autenticação está presente na sessão
    if "authentication_status" not in st.session_state:
        st.session_state["authentication_status"] = None

    if st.session_state.get("authentication_status") is None or st.session_state.get("authentication_status") is False:
        authenticate()

    #Verifica o Estado de Login
    if st.session_state.get("authentication_status"):
        # Verifica se a mensagem de sucesso já foi exibida
        if not st.session_state.get("success_message_displayed", False):
            # Exibe a mensagem de sucesso temporária
            display_temporary_success_message()
            # Marca que a mensagem foi exibida para não repetir na próxima execução
            st.session_state["success_message_displayed"] = True
        #Chamam a função do Projeto Canhadas
        projeto()
    #Se a senha ou login estirverem errados imprime a menssagem de erro:
    elif st.session_state.get("authentication_status") is False:
        st.error("Usuário e/ou Senha Incorretos")

if __name__ == "__main__":
    main()
