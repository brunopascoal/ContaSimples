import streamlit as st
import pandas as pd
import numpy as np


def run_juntar_arquivos_app():
    def process_files(uploaded_files, file_type):
        dfs = []  # Lista para armazenar os dataframes

        for uploaded_file in uploaded_files:
            # Carregar o arquivo em um dataframe
            if file_type == "excel":
                df = pd.read_excel(uploaded_file, header=None)
            else:  # csv
                df = pd.read_csv(uploaded_file, header=None)

            # Determinar o número de colunas e criar nomes de colunas genéricos
            n_colunas = df.shape[1]
            df.columns = [f"Coluna{i+1}" for i in range(n_colunas)]

            # Criar uma linha fictícia com o mesmo número de colunas
            linha_ficticia = pd.DataFrame([[np.nan] * n_colunas], columns=df.columns)

            # Adicionar a linha fictícia ao início de cada DataFrame
            df = pd.concat([linha_ficticia, df], ignore_index=True)

            # Adicionar o dataframe à lista
            dfs.append(df)

        # Concatenar todos os dataframes, ignorando os índices
        df_concatenado = pd.concat(dfs, ignore_index=True)
        return df_concatenado

    def main():
        # Interface Streamlit
        st.title("Concatenador de Arquivos")

        # Permitir ao usuário fazer upload de vários arquivos
        uploaded_files = st.file_uploader(
            "Faça upload dos arquivos para concatenar",
            accept_multiple_files=True,
            type=["xlsx", "csv"],
        )
        file_type = st.selectbox("Tipo do arquivo de saída", ("excel", "csv"))

        if uploaded_files:
            if st.button("Processar"):
                result_df = process_files(uploaded_files, file_type)
                st.dataframe(result_df)

                # Permitir download do arquivo processado
                if file_type == "excel":
                    output = result_df.to_excel(index=False, engine="xlsxwriter")
                    st.download_button(
                        label="Download do Excel processado",
                        data=output,
                        file_name="arquivo_concatenado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:  # csv
                    output = result_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="Download do CSV processado",
                        data=output,
                        file_name="arquivo_concatenado.csv",
                        mime="text/csv",
                    )

    main()
