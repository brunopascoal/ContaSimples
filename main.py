# app_principal.py
import streamlit as st
from speds.speds import run_sped_app
from relatorios_financeiros.relatorios_financeiros import run_relatorios_app
from selecoes_aleatorias.selecoes import run_selecoes_app
from comparativo.tratar_balancetes import run_tratar_balancetes_app
from comparativo.gerar_balancetes import run_gerar_balancetes_app
from comparativo.conferencia_balancetes import run_conferencia_balancetes_app
from comparativo.teste_saldo_inicial import run_teste_saldo_inicial_app
from operadoras.bl_dre_ops import run_bl_e_dre_app
from streamlit_option_menu import option_menu
import base64
import os

# Função para carregar o arquivo CSS
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Caminho do arquivo CSS
current_dir = os.path.dirname(os.path.abspath(__file__))
css_path = os.path.join(current_dir, 'assets', 'styles', 'styles.css')
local_css(css_path)


def main():
 

        # Sidebar para seleção da aplicação
    with st.sidebar:
        st.sidebar.title("Menu")
        app_choice = option_menu(
            "Menu",
            [
                "Home",
                "SPEDs",
                "Relatorios Financeiros",
                "Seleções Aleatórias",
                "Balancetes",
                "BL e DRE OPS",
            ],
            icons=["house", "gear", "bank", "infinity", "cash-coin"],
            menu_icon="cast",
            default_index=0,
        )
        st.caption('Versão atual: 1.0.1')

    # Executar a aplicação selecionada
    if app_choice == "Home":


        current_dir = os.path.dirname(os.path.abspath(__file__))



        st.markdown(f"""
            <div style='display: flex; align-items: center; justify-content: center;'>
                <h1>ContaSimples</h1>
            </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style='display: flex; align-items: center; justify-content: center;'>
                <h2>Sistema para auditores</h2>
            </div>
        """, unsafe_allow_html=True)


        st.divider()  # 👈 Draws a horizontal rule
        st.markdown(f"""
        <div style='display: flex; align-items: center; justify-content: center;'>
                <p>Selecione no menu à esquerda a ferramenta que deseja utilizar e bom trabalho!</p>
            </div>
        """, unsafe_allow_html=True)
    
    elif app_choice == "SPEDs":
        run_sped_app()
    elif app_choice == "Relatorios Financeiros":
        run_relatorios_app()
    elif app_choice == "Seleções Aleatórias":
        run_selecoes_app()
    elif app_choice == "BL e DRE OPS":
        run_bl_e_dre_app()
    elif app_choice == "Balancetes":
        # Usar st.session_state para manter o estado
        if "balancetes_choice" not in st.session_state:
            st.session_state.balancetes_choice = "Tratar Balancetes"  # Valor padrão

        # Criar abas para a seleção
        tab1, tab2, tab3, tab4 = st.tabs(
            ["Tratar Balancetes", "Gerar Balancetes", "Conferencia Balancetes", "Teste de Saldo Inicial"]
        )

        # Atualiza o estado com a aba selecionada e executa a aplicação correspondente
        with tab1:
            st.session_state.balancetes_choice = "Tratar Balancetes"
            run_tratar_balancetes_app()
        with tab2:
            st.session_state.balancetes_choice = "Gerar Balancetes"
            run_gerar_balancetes_app()
        with tab3:
            st.session_state.balancetes_choice = "Conferencia Balancetes"
            run_conferencia_balancetes_app()
        with tab4:
            st.session_state.balancetes_choice = "Teste de Saldo Inicial"
            run_teste_saldo_inicial_app()


if __name__ == "__main__":
        main()
