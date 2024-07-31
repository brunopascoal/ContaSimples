import streamlit as st
import pandas as pd
from io import BytesIO


# Função auxiliar para converter um DataFrame em Excel
def run_conferencia_balancetes_app():
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        return output.getvalue()

    # Criando a página com Streamlit
    st.title("Conferência de Balancetes")

    # Upload do arquivo Excel
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])

    if uploaded_file:
        # Carregando todas as sheets do arquivo
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

        # Criando um DataFrame para armazenar os resultados da conferência
        conferencia_df = pd.DataFrame(
            columns=[
                "Mês",
                "CONTA",
                "DESCRIÇÃO",
                "Saldo Final",
                "Saldo Inicial Próximo Mês",
                "Conferência",
            ]
        )

        # Lista de meses para iteração
        meses = list(all_sheets.keys())

        # Função para comparar saldos
        def compara_saldo(saldo_final_atual, saldo_inicial_proximo):
            return "Sim" if saldo_final_atual == saldo_inicial_proximo else "Não"

        # Iterando pelas sheets e comparando os saldos
        for i in range(len(meses) - 1):
            mes_atual = meses[i]
            proximo_mes = meses[i + 1]

            df_atual = all_sheets[mes_atual]
            df_proximo = all_sheets[proximo_mes]

            col_saldo_final_atual = f"SALDO FINAL {mes_atual.upper()}"
            col_saldo_inicial_proximo = f"SALDO INICIAL {proximo_mes.upper()}"

            merged_df = pd.merge(
                df_atual[["CONTA", "DESCRIÇÃO", col_saldo_final_atual]],
                df_proximo[["CONTA", col_saldo_inicial_proximo]],
                on="CONTA",
                how="inner",
            )

            merged_df["Conferência"] = merged_df.apply(
                lambda row: compara_saldo(
                    row[col_saldo_final_atual], row[col_saldo_inicial_proximo]
                ),
                axis=1,
            )
            merged_df["Mês"] = mes_atual
            merged_df = merged_df[
                [
                    "Mês",
                    "CONTA",
                    "DESCRIÇÃO",
                    col_saldo_final_atual,
                    col_saldo_inicial_proximo,
                    "Conferência",
                ]
            ]

            conferencia_df = pd.concat([conferencia_df, merged_df], ignore_index=True)

        result = to_excel(conferencia_df)
        st.download_button(
            label="Baixar resultado como Excel",
            data=result,
            file_name="conferencia_balancetes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
