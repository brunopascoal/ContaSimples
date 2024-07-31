import streamlit as st
import pandas as pd

# Título da aplicação
st.title("Manipulador de Arquivos Dinâmico")

# Upload de múltiplos arquivos
uploaded_files = st.file_uploader(
    "Escolha os arquivos", accept_multiple_files=True, type=["csv", "xlsx"]
)

# Para cada arquivo carregado, leia-o como um DataFrame
dataframes = {
    file.name: pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    for file in uploaded_files
}

# Processamento para cada DataFrame
for filename, df in dataframes.items():
    st.write(f"### {filename}")
    st.dataframe(df)

    # Seleção de colunas para a chave única
    selected_columns = st.multiselect(
        f"Selecione colunas para criar uma chave única em {filename}",
        options=df.columns,
        key=filename,
    )

    if selected_columns:
        # Criação da chave única por concatenação das colunas selecionadas
        df["chave_unica"] = df[selected_columns].astype(str).apply("-".join, axis=1)
        st.write(f"DataFrame com chave única para {filename}:")
        st.dataframe(df[["chave_unica"]])
