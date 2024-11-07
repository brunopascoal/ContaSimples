import streamlit as st
import os
import pandas as pd
import io
from datetime import datetime
from dateutil.relativedelta import relativedelta

def run_sped_app():
    def formatar_data(aaaamm):
        return datetime.strptime(aaaamm, "%Y%m").strftime("%b/%Y").upper()

    def process_sped_file(file_content, user_start_date, user_end_date, seen):
        reg0000_columns = [
            "REG0000", "COD_VER", "COD_FIN", "DT_INI", "DT_FIN", "NOME", "CNPJ", 
            "CPF", "UF", "IE", "COD_MUN", "IM", "SUFRAMA", "IND_PERFIL", "IND_ATIV"
        ]
        Verific_data = []
        duplicates, signatures_absent = [], []

        lines = file_content.split("\n")
        
        for line in lines:
            if line.startswith("|0000|"):
                reg0000 = line.strip().split("|")[1:-1]
                reg0000 += [""] * (len(reg0000_columns) - len(reg0000))  
                reg0000[3] = reg0000[3][4:8] + reg0000[3][2:4]  
                reg0000[4] = reg0000[4][4:8] + reg0000[4][2:4]

                cnpj = reg0000[6]
                periodo = reg0000[3]
                cod_fin = int(reg0000[2])
                key = (cnpj, periodo, cod_fin)

                if key in seen:
                    prev_record = seen[key]
                    prev_cod_fin = int(prev_record[2])  

                    if prev_cod_fin == cod_fin:
                        duplicates.append((cnpj, reg0000[5], f"Duplicado detectado no período {formatar_data(periodo)}"))
                        continue

                    if cod_fin == 1:
                        seen[key] = reg0000

                else:
                    seen[key] = reg0000

            Verific_data.append(seen[key])

        Verific_df = pd.DataFrame(Verific_data, columns=reg0000_columns)
        
        def check_missing_months(records, user_start_date, user_end_date):
            empresas = {}
            arquivos_fora_periodo = []  
            mes_para_numero = {
                "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
                "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
            }
            for _, record in records.iterrows():
                cnpj, dt_ini = record['CNPJ'], record['DT_INI']
                mes_inicial, ano_inicial = int(dt_ini[4:6]), dt_ini[:4]

                file_month_year = datetime.strptime(dt_ini, "%Y%m")
                primeiro_mes, ultimo_mes = datetime.strptime(user_start_date, "%Y%m"), datetime.strptime(user_end_date, "%Y%m")

                if file_month_year < primeiro_mes or file_month_year > ultimo_mes:
                    arquivos_fora_periodo.append(f"Arquivo com CNPJ {cnpj} fora do período: {file_month_year.strftime('%b/%Y')}")

                if cnpj not in empresas:
                    empresas[cnpj] = set()
                
                empresas[cnpj].add((mes_inicial, ano_inicial))

            current_month = primeiro_mes
            meses_no_intervalo = []
            while current_month <= ultimo_mes:
                meses_no_intervalo.append(current_month.strftime("%b/%Y").upper())
                current_month += relativedelta(months=1)

            df_result = pd.DataFrame(columns=['CNPJ', 'Empresa'] + meses_no_intervalo)

            for cnpj, meses_recebidos in empresas.items():
                empresa_nome = records[records['CNPJ'] == cnpj]['NOME'].values[0]
                row = {'CNPJ': cnpj, 'Empresa': empresa_nome}
                for mes in meses_no_intervalo:
                    mes_nome, ano_nome = mes.split('/')
                    mes_numero = mes_para_numero[mes_nome]
                    row[mes] = 'X' if (mes_numero, ano_nome) in meses_recebidos else ''
                
                df_result = pd.concat([df_result, pd.DataFrame([row])], ignore_index=True)

            if arquivos_fora_periodo:
                st.warning("Arquivos fora do período selecionado:")
                for aviso in arquivos_fora_periodo:
                    st.write(aviso)

            return df_result

        def check_digital_signature(file_content, Verific_df):
            last_lines = file_content.splitlines()[-200:]
            signature_indicators = ["ICP-Brasil", "Certificadora", "Autoridade", "AC", "RFB", "BR", "Certsign"]

            if not any(any(indicator in line for indicator in signature_indicators) for line in last_lines):
                empresa_info = Verific_df.iloc[0]
                signatures_absent.append((empresa_info['CNPJ'], empresa_info['NOME'], f"Assinatura digital ausente no período {formatar_data(empresa_info['DT_INI'])}"))

        check_digital_signature(file_content, Verific_df)
        missing_months_df = check_missing_months(Verific_df, user_start_date, user_end_date)

        df_problems = pd.DataFrame(duplicates + signatures_absent, columns=['CNPJ', 'Empresa', 'Situações Detectadas'])

        return Verific_df, df_problems, missing_months_df

    def extract_and_add_0000(file_content, record_type, seen):
        reg0000_columns = [
            "REG0000", "COD_VER", "COD_FIN", "DT_INI", "DT_FIN", "NOME", "CNPJ", "CPF", "UF", "IE", 
            "COD_MUN", "IM", "SUFRAMA", "IND_PERFIL", "IND_ATIV"
        ]
        c100_columns = [
            "REGC100", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "NUM_DOC", "CHV_NFE", 
            "DT_DOC", "DT_E_S", "VL_DOC", "IND_PGTO", "VL_DESC1", "VL_ABAT_NT1", "VL_MERC", "IND_FRT", "VL_FRT",
            "VL_SEG", "VL_OUT_DA", "VL_BC_ICMS1", "VL_ICMS1", "VL_BC_ICMS_ST1", "VL_ICMS_ST1", "VL_IPI1", "VL_PIS1", 
            "VL_COFINS1", "VL_PIS_ST1", "VL_COFINS_ST"
        ]
        c170_columns = [
            "REGC170", "NUM_ITEM", "COD_ITEM", "DESCR_COMPL", "QTD", "UNID", "VL_ITEM", "VL_DESC", "IND_MOV", 
            "CST_ICMS", "CFOP", "COD_NAT", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", 
            "VL_ICMS_ST", "IND_APUR", "CST_IPI", "COD_ENQ", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI", "CST_PIS", 
            "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "ALIQ_PIS1", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", 
            "ALIQ_COFINS", "QUANT_BC_COFINS", "ALIQ_COFINS1", "VL_COFINS", "COD_CTA", "VL_ABAT_NT"
        ]
        c190_columns = [
            "REGC190", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", 
            "VL_ICMS_ST", "VL_RED_BC", "VL_IPI", "COD_OBS"
        ]
        
        combined_data = []
        lines = file_content.split("\n")
        current_c100 = None

        for line in lines:
            if line.startswith("|0000|"):
                reg0000 = line.strip().split("|")[1:-1]
                reg0000 += [""] * (len(reg0000_columns) - len(reg0000))  
                cnpj, periodo, cod_fin = reg0000[6], reg0000[3], int(reg0000[2])

                key = (cnpj, periodo, cod_fin)
                if key in seen:
                    continue  

            elif line.startswith("|C100|"):
                current_c100 = line.strip().split("|")[1:-1]
                current_c100 += [""] * (len(c100_columns) - len(current_c100))

            elif line.startswith(f"|{record_type}|") and current_c100:
                c_record = line.strip().split("|")[1:-1]
                c_columns = c170_columns if record_type == "C170" else c190_columns
                c_record += [""] * (len(c_columns) - len(c_record))

                combined_record = c_record + current_c100 + reg0000
                combined_data.append(combined_record)

        combined_df = pd.DataFrame(
            combined_data,
            columns=(c170_columns if record_type == "C170" else c190_columns) + c100_columns + reg0000_columns
        )
        return combined_df

    def create_all_pivot_tables(combined_df, record_type, writer):
        if record_type == "C170":
            sheets_info = [
                ("C170 por CNPJ para VL_ITEM", 'CNPJ', 'CFOP', 'VL_ITEM', 'All'),
                ("C170 por CNPJ para VL_ICMS", 'CNPJ', 'CFOP', 'VL_ICMS', 'All'),
                ("C170 por DT_FIN para VL_ITEM", 'DT_FIN', 'CFOP', 'VL_ITEM', 'All'),
                ("C170 por DT_FIN para VL_ICMS", 'DT_FIN', 'CFOP', 'VL_ICMS', 'All')
            ]
        elif record_type == "C190":
            sheets_info = [
                ("CFOP 1_2_3 CNPJ VL_OPR", 'CNPJ', 'CFOP', 'VL_OPR', 'Entrada'),
                ("CFOP 1_2_3 CNPJ VL_ICMS", 'CNPJ', 'CFOP', 'VL_ICMS', 'Entrada'),
                ("CFOP 1_2_3 DT_FIN VL_OPR", 'DT_FIN', 'CFOP', 'VL_OPR', 'Entrada'),
                ("CFOP 1_2_3 DT_FIN VL_ICMS", 'DT_FIN', 'CFOP', 'VL_ICMS', 'Entrada'),
                ("CFOP 5_6_7 CNPJ VL_OPR", 'CNPJ', 'CFOP', 'VL_OPR', 'Saída'),
                ("CFOP 5_6_7 CNPJ VL_ICMS", 'CNPJ', 'CFOP', 'VL_ICMS', 'Saída'),
                ("CFOP 5_6_7 DT_FIN VL_OPR", 'DT_FIN', 'CFOP', 'VL_OPR', 'Saída'),
                ("CFOP 5_6_7 DT_FIN VL_ICMS", 'DT_FIN', 'CFOP', 'VL_ICMS', 'Saída')
            ]

        for sheet_name, pivot_column, index_column, pivot_values, cfop_type in sheets_info:
            if cfop_type != 'All':
                filtered_df = filter_cfop(combined_df, cfop_type)
            else:
                filtered_df = combined_df
            pivot_table = create_pivot_table(filtered_df, index_column, pivot_column, pivot_values)
            pivot_table.to_excel(writer, sheet_name=sheet_name)

    def filter_cfop(df, cfop_type):
        return df[df["CFOP"].astype(str).str.startswith(("1", "2", "3") if cfop_type == "Entrada" else ("5", "6", "7"))]

    def create_pivot_table(df, index, columns, values):
        df[values] = pd.to_numeric(df[values], errors="coerce")
        pivot = df.pivot_table(index=index, columns=columns, values=values, aggfunc="sum", fill_value=0)
        return pivot

    st.title("Processador de Arquivos SPED")

    if "verificado" not in st.session_state:
        st.session_state["verificado"] = False

    uploaded_files = st.file_uploader("Selecione os arquivos .txt:", type="txt", accept_multiple_files=True)
    new_record_type = st.sidebar.selectbox("Selecione o tipo de registro", ("C170", "C190"))

    col1, col2 = st.columns(2)
    with col1:
        start_month = st.selectbox("Mês Inicial", [f"{i:02}" for i in range(1, 13)])
    with col2:
        start_year = st.number_input("Ano Inicial", min_value=2000, max_value=2100, value=2024)

    col3, col4 = st.columns(2)
    with col3:
        end_month = st.selectbox("Mês Final", [f"{i:02}" for i in range(1, 13)])
    with col4:
        end_year = st.number_input("Ano Final", min_value=2000, max_value=2100, value=2024)

    user_start_date = f"{start_year}{start_month}"
    user_end_date = f"{end_year}{end_month}"

    if st.button("Verificar Arquivos") and uploaded_files:
        all_problems_dfs, all_missing_months_dfs, seen = [], [], {}
        combined_problems_df = pd.DataFrame()

        for uploaded_file in uploaded_files:
            file_content = uploaded_file.getvalue().decode("iso-8859-1")
            data, problems_df, missing_months_df = process_sped_file(file_content, user_start_date, user_end_date, seen)
            all_problems_dfs.append(problems_df)
            all_missing_months_dfs.append(missing_months_df)

        if all_missing_months_dfs:
            st.warning("Meses Detectados:")
            combined_missing_months_df = pd.concat(all_missing_months_dfs, ignore_index=False)
            consolidated_df = combined_missing_months_df.groupby('CNPJ').max()
            st.session_state['consolidated_df'] = consolidated_df
            st.dataframe(consolidated_df)

        if all_problems_dfs:
            combined_problems_df = pd.concat(all_problems_dfs, ignore_index=True)
            if not combined_problems_df.empty:
                st.warning("Situações Detectadas:")
                st.session_state['combined_problems_df'] = combined_problems_df
                st.dataframe(combined_problems_df)
            else:
                st.success("Nenhuma situação foi detectada.")
        else:
            st.success("Nenhuma situação foi detectada.")

        if 'consolidated_df' in st.session_state or not combined_problems_df.empty:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if 'consolidated_df' in st.session_state:
                    st.session_state['consolidated_df'].to_excel(writer, sheet_name='Meses', index=True)
                if not combined_problems_df.empty:
                    combined_problems_df.to_excel(writer, sheet_name='Situações Detectadas', index=False)
            output.seek(0)
            st.download_button("Baixar Verificações", data=output, file_name='Verificações_Speds.xlsx')

        st.session_state["verificado"] = True

    if st.session_state.get("verificado", False):
        if st.button("Processar Arquivos") and uploaded_files:
            seen = {}
            combined_df = pd.DataFrame()
            for uploaded_file in uploaded_files:
                file_content = uploaded_file.getvalue().decode("iso-8859-1")
                data = extract_and_add_0000(file_content, new_record_type, seen)
                combined_df = pd.concat([combined_df, data])

            # Armazena o DataFrame processado na sessão
            st.session_state["combined_df"] = combined_df
            st.session_state["data_processed"] = True

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="Dados Processados")
                if 'consolidated_df' in st.session_state:
                    st.session_state['consolidated_df'].to_excel(writer, sheet_name='Meses', index=True)
                if 'combined_problems_df' in st.session_state and not st.session_state['combined_problems_df'].empty:
                    st.session_state['combined_problems_df'].to_excel(writer, sheet_name='Situações Detectadas', index=False)

                if new_record_type == "C170":
                    create_all_pivot_tables(combined_df, "C170", writer)
                elif new_record_type == "C190":
                    create_all_pivot_tables(combined_df, "C190", writer)

            output.seek(0)
            st.download_button("Baixar dados e tabelas em Excel", data=output, file_name=f"SPEDS_{new_record_type}.xlsx")

    # Verifica se o DataFrame processado está disponível antes de tentar acessar
    if st.session_state.get("data_processed", False) and "combined_df" in st.session_state:
        st.subheader("Montar Tabela Personalizada")
        combined_df = st.session_state["combined_df"]

        # Opção para o usuário montar sua própria tabela pivot
        pivot_row = st.selectbox("Escolha a linha para a tabela pivot", combined_df.columns)
        pivot_column = st.selectbox("Escolha a coluna para a tabela pivot", combined_df.columns)
        pivot_values = st.selectbox("Escolha os valores para a tabela pivot", combined_df.columns)

        if st.button(f"Gerar Tabela Pivot {new_record_type}"):
            pivot_table = create_pivot_table(combined_df, pivot_row, pivot_column, pivot_values)
            st.write(pivot_table)

            output_pivot = io.BytesIO()
            with pd.ExcelWriter(output_pivot, engine="xlsxwriter") as writer:
                pivot_table.to_excel(writer, index=True, sheet_name=f"Tabela_Pivot_{new_record_type}")
            output_pivot.seek(0)

            st.download_button(
                label=f"Baixar Tabela Pivot {new_record_type} em Excel",
                data=output_pivot,
                file_name=f"Pivot_Table_{new_record_type}.xlsx"
            )

# Função principal
run_sped_app()
