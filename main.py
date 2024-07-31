# app_principal.py
import streamlit as st
from speds import run_sped_app
from relatorios_financeiros import run_relatorios_app
from selecoes import run_selecoes_app
from comparativo.tratar_balancetes import run_tratar_balancetes_app
from comparativo.gerar_balancetes import run_gerar_balancetes_app
from comparativo.conferencia_balancetes import run_conferencia_balancetes_app
from comparativo.teste_saldo_inicial import run_teste_saldo_inicial_app
from compasso.empresa_unica import run_compasso_app
from compasso.consolidacao import run_compasso_consolidação_app
from bl_dre_ops import run_bl_e_dre_app
from streamlit_option_menu import option_menu


def main():
    # Sidebar para seleção da aplicação
    with st.sidebar:
        # st.image("img_logo.jfif", width=100)
        st.sidebar.title("Menu")
        app_choice = option_menu(
            "Menu",
            [
                "Home",
                "SPEDs",
                "Relatorios Financeiros",
                "Seleções Aleatórias",
                "Balancetes",
                "Relatorios Compasso",
                "BL e DRE OPS",
            ],
            icons=["house", "gear", "bank", "infinity", "cash-coin"],
            menu_icon="cast",
            default_index=0,
        )

    # Executar a aplicação selecionada
    if app_choice == "Home":
        st.title("Sistema DI")
        st.divider()  # 👈 Draws a horizontal rule
        st.write("Bem-vindo ao Sistema do Departamento Interno!")
    elif app_choice == "SPEDs":
        run_sped_app()
    elif app_choice == "Relatorios Financeiros":
        run_relatorios_app()
    elif app_choice == "Seleções Aleatórias":
        run_selecoes_app()
    elif app_choice == "Relatorios Compasso":
        # Usar st.session_state para manter o estado
        if "compasso_choice" not in st.session_state:
            st.session_state.compasso_choice = "Análise de Empresa Única"  # Valor padrão

        compasso_choice = st.sidebar.selectbox(
            "Escolha a aplicação",
            [
                "Análise de Empresa Única",
                "Análise Consolidada de Grupo"
            ],
            index=["Análise de Empresa Única", "Análise Consolidada de Grupo"].index(st.session_state.compasso_choice),
        )

        st.session_state.compasso_choice = compasso_choice  # Atualiza o estado com a escolha do usuário

        if compasso_choice == "Análise de Empresa Única":
            run_compasso_app()
        elif compasso_choice == "Análise Consolidada de Grupo":
            run_compasso_consolidação_app()

    elif app_choice == "BL e DRE OPS":
        run_bl_e_dre_app()
    elif app_choice == "Balancetes":
        # Usar st.session_state para manter o estado
        if "balancetes_choice" not in st.session_state:
            st.session_state.balancetes_choice = "Tratar Balancetes"  # Valor padrão

        balancetes_choice = st.sidebar.selectbox(
            "Escolha a aplicação",
            [
                "Tratar Balancetes",
                "Gerar Balancetes",
                "Conferencia Balancetes",
                "Teste de Saldo Inicial",
            ],
            index=["Tratar Balancetes", "Gerar Balancetes", "Conferencia Balancetes", "Teste de Saldo Inicial"].index(st.session_state.balancetes_choice),
        )

        st.session_state.balancetes_choice = balancetes_choice  # Atualiza o estado com a escolha do usuário

        if balancetes_choice == "Tratar Balancetes":
            run_tratar_balancetes_app()
        elif balancetes_choice == "Gerar Balancetes":
            run_gerar_balancetes_app()
        elif balancetes_choice == "Conferencia Balancetes":
            run_conferencia_balancetes_app()
        elif balancetes_choice == "Teste de Saldo Inicial":
            run_teste_saldo_inicial_app()
    

if __name__ == "__main__":
    main()