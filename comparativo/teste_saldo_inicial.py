import streamlit as st
import pandas as pd
from io import BytesIO

# Função auxiliar para converter um DataFrame em Excel
def run_teste_saldo_inicial_app():
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        return output.getvalue()

    def main():
        # Criando a página com Streamlit
        st.title("Teste de Saldo Inicial")

        # Upload dos arquivos Excel
        uploaded_file_dezembro = st.file_uploader("Escolha o arquivo Excel de Dezembro", type=["xlsx"], key="dezembro")
        uploaded_file_janeiro = st.file_uploader("Escolha o arquivo Excel de Janeiro", type=["xlsx"], key="janeiro")

        if uploaded_file_dezembro and uploaded_file_janeiro:
            # Carregando os DataFrames
            df_dezembro = pd.read_excel(uploaded_file_dezembro)
            df_janeiro = pd.read_excel(uploaded_file_janeiro)

            col_saldo_final_dezembro = st.selectbox("Selecione a coluna de saldo final em Dezembro", df_dezembro.columns, key="saldo_final_dez")
            col_saldo_inicial_janeiro = st.selectbox("Selecione a coluna de saldo inicial em Janeiro", df_janeiro.columns, key="saldo_inicial_jan")

            # Unindo os dataframes com base na coluna "CONTA"
            merged_df = pd.merge(
                df_dezembro[["CONTA", "DESCRIÇÃO", col_saldo_final_dezembro]],
                df_janeiro[["CONTA", col_saldo_inicial_janeiro]],
                on="CONTA",
                how="inner"
            )

            # Comparando os saldos e adicionando resultado na coluna "Conferência"
            merged_df["Conferência"] = merged_df.apply(
                lambda row: "Ok" if row[col_saldo_final_dezembro] == row[col_saldo_inicial_janeiro] else "Analisar",
                axis=1
            )
            merged_df["Mês"] = "Dezembro - Janeiro"
            merged_df = merged_df[
                [
                    #"Mês",
                    "CONTA",
                    "DESCRIÇÃO",
                    col_saldo_final_dezembro,
                    col_saldo_inicial_janeiro,
                    "Conferência",
                ]
            ]

            result = to_excel(merged_df)
            st.download_button(
                label="Baixar resultado como Excel",
                data=result,
                file_name="Teste de Saldo Inicial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    main()

