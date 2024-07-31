import streamlit as st
import pandas as pd
from io import BytesIO
import io  # IMPORTAÇÔES
import re
import zipfile


def run_gerar_balancetes_app():
    def mes_abreviado(mes):
        meses = [
            "JAN",
            "FEV",
            "MAR",
            "ABR",
            "MAI",
            "JUN",
            "JUL",
            "AGO",
            "SET",
            "OUT",
            "NOV",
            "DEZ",
        ]
        return meses[mes - 1]

    def gerar_sumario(df, mesInicial, mesFinal):
        # Colunas para agrupar e sumarizar
        colunas_sumarizacao = ["CONTA"]
        colunas_inclusao = ["CONTA", "DESCRIÇÃO", "tipo"]
        # Verifica se deve incluir a coluna 'RED'
        if "RED" in df.columns:
            colunas_inclusao.append("RED")
        # Adicionando colunas de saldo inicial, débito, crédito, movimento e saldo final
        colunas_saldo = ["SALDO INICIAL " + mes_abreviado(mesInicial)]
        colunas_saldo += [
            "DEBITO " + mes_abreviado(i) for i in range(mesInicial, mesFinal + 1)
        ]
        colunas_saldo += [
            "CREDITO " + mes_abreviado(i) for i in range(mesInicial, mesFinal + 1)
        ]
        colunas_saldo += [
            "MOVIMENTO " + mes_abreviado(i) for i in range(mesInicial, mesFinal + 1)
        ]
        colunas_saldo.append("SALDO FINAL " + mes_abreviado(mesFinal))
        # Configurando a agregação
        agregacoes = {col: "sum" for col in colunas_saldo}
        for col in colunas_inclusao:
            agregacoes[col] = "first"
        # Realizando a sumarização
        df_sumario = df.groupby(colunas_sumarizacao).agg(agregacoes)
        return df_sumario

    def adicionar_campo_variacao(df, mesInicial, mesSeguinte):
        mesInicialAbrev = mes_abreviado(mesInicial)
        mesSeguinteAbrev = mes_abreviado(mesSeguinte)
        # Nome do novo campo
        nome_campo_variacao = f"VAR_{mesInicialAbrev}"
        # Calculando a variação
        df[nome_campo_variacao] = df.apply(
            lambda row: (
                0
                if row[f"MOVIMENTO {mesInicialAbrev}"] == 0
                else (
                    row[f"MOVIMENTO {mesInicialAbrev}"]
                    - row[f"MOVIMENTO {mesSeguinteAbrev}"]
                )
                / row[f"MOVIMENTO {mesInicialAbrev}"]
            ),
            axis=1,
        )
        # Arredondando os valores para 2 casas decimais
        df[nome_campo_variacao] = df[nome_campo_variacao].round(2)

    def adicionar_campo_saldo_anual(df, mesInicial, mesFinal):
        # Inicializando a coluna de saldo anual
        df["SALDO_ANUAL"] = df[f"SALDO INICIAL {mes_abreviado(mesInicial)}"]
        # Somando os valores de movimento para cada mês
        for i in range(mesInicial, mesFinal + 1):
            mes_abrev = mes_abreviado(i)
            df["SALDO_ANUAL"] += df[f"MOVIMENTO {mes_abrev}"]
        # Arredondando os valores para 2 casas decimais
        df["SALDO_ANUAL"] = df["SALDO_ANUAL"].round(2)

    def adicionar_campo_conferencia_auditoria(df, mesFinal):
        mesFinalAbrev = mes_abreviado(mesFinal)

        # Calculando a conferência de auditoria
        df["CONFERÊNCIA_AUDITORIA"] = (
            df["SALDO_ANUAL"] - df[f"SALDO FINAL {mesFinalAbrev}"]
        )

        # Arredondando os valores para 2 casas decimais
        df["CONFERÊNCIA_AUDITORIA"] = df["CONFERÊNCIA_AUDITORIA"].round(2)

    def reordenar_colunas(df_sumario, mesInicial, mesFinal):
        colunas_base = ["CONTA", "DESCRIÇÃO"]
        colunas_condicionais = [
            col for col in ["tipo", "RED"] if col in df_sumario.columns
        ]
        colunas_meses = []

        for mes in range(mesInicial, mesFinal + 1):
            mes_abrev = mes_abreviado(mes)
            colunas_mes = [
                f"SALDO INICIAL {mes_abrev}" if mes == mesInicial else None,
                f"DEBITO {mes_abrev}",
                f"CREDITO {mes_abrev}",
                f"MOVIMENTO {mes_abrev}",
                f"VAR_{mes_abrev}" if mes != mesFinal else None,
            ]
            colunas_meses.extend(filter(None, colunas_mes))

        colunas_finais = [
            f"SALDO FINAL {mes_abreviado(mesFinal)}",
            "SALDO_ANUAL",
            "CONFERÊNCIA_AUDITORIA",
        ]

        # Combina todas as colunas na ordem desejada
        colunas_ordenadas = (
            colunas_base + colunas_condicionais + colunas_meses + colunas_finais
        )

        # Reordena o DataFrame
        df_reordenado = df_sumario[colunas_ordenadas]

        return df_reordenado

    def main():
        st.title("Gerar Balancetes")

        uploaded_files = st.file_uploader(
            "Escolha os arquivos dos balancetes (JAN a DEZ)",
            accept_multiple_files=True,
            type=["xlsx"],
        )

        numMesUm, numUltimoMes = st.slider(
            "Selecione o intervalo de meses", 1, 12, (1, 12)
        )

        nome_planilhas = {
            1: "jan",
            2: "fev",
            3: "mar",
            4: "abr",
            5: "mai",
            6: "jun",
            7: "jul",
            8: "ago",
            9: "set",
            10: "out",
            11: "nov",
            12: "dez",
        }

        if st.button("Gerar Balancete", key="gerar_balancete") and uploaded_files:
            dfs = []
            for uploaded_file in uploaded_files:
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    sheet_names = xls.sheet_names  # Obtém a lista de todas as planilhas

                    # Determina os nomes das planilhas com base na seleção do usuário
                    for mes in range(numMesUm, numUltimoMes + 1):
                        nome_planilha = nome_planilhas.get(
                            mes
                        )  # Obtém o nome da planilha para o mês
                        if (
                            nome_planilha in sheet_names
                        ):  # Verifica se a planilha existe
                            df = pd.read_excel(xls, sheet_name=nome_planilha)
                            dfs.append(df)
                            st.write(f"DataFrame do mês: {nome_planilha}")
                            st.dataframe(df)
                        else:
                            st.error(
                                f"Não foi possível encontrar a planilha para o mês {mes} ({nome_planilha}). Verifique se o arquivo contém todas as planilhas necessárias."
                            )
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo: {e}")
            st.write("Finalizado... Processando dados")

            if dfs:
                # Combina os DataFrames de cada mês em um único DataFrame
                df_combinado = pd.concat(dfs, ignore_index=True)

                # Processamento dos dados
                df_sumario = gerar_sumario(df_combinado, numMesUm, numUltimoMes)
                for i in range(numMesUm, numUltimoMes):
                    adicionar_campo_variacao(df_sumario, i, i + 1)
                adicionar_campo_saldo_anual(df_sumario, numMesUm, numUltimoMes)
                adicionar_campo_conferencia_auditoria(df_sumario, numUltimoMes)
                df_sumario_reordenado = reordenar_colunas(
                    df_sumario, numMesUm, numUltimoMes
                )
                st.dataframe(df_sumario_reordenado)
                # Botão para baixar o resultado
                output = io.BytesIO()
                df_sumario_reordenado.to_excel(output, index=False)
                output.seek(0)  # Voltar ao início do stream
                st.download_button(
                    label="Baixar Balancete Final",
                    data=output,
                    file_name="Balancete Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.write("Deu ruim:")

        else:
            st.error(
                "Por favor, carregue os arquivos dos balancetes antes de gerar o relatório."
            )

    main()
