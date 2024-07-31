import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl import Workbook

def run_compasso_app():
    
    def load_data(file, sheet_name=None):
        """ Carrega dados de um arquivo Excel. """
        data = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(data, dict):
            # Se `data` for um dicionário, pegamos a primeira planilha
            data = next(iter(data.values()))
        return data

    def create_excel_download_link(dataframes, filename):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in dataframes.items():
                # Truncar o nome da folha para 31 caracteres e remover caracteres inválidos
                clean_sheet_name = re.sub(r'[\[\]\*:/\\?]', '', sheet_name)[:31]
                df.to_excel(writer, index=True, sheet_name=clean_sheet_name)
        output.seek(0)
        return output

    def get_account_balance(account_df, account_number, account_column, selected_column):
        """Calcula o saldo total para uma conta específica usando a coluna selecionada."""
        if account_number in account_df[account_column].values:
            return account_df[account_df[account_column] == account_number][selected_column].sum()
        else:
            return 0
        
    def get_account_value(account_df, account_name, selected_column):
        if account_name in account_df['DESCRIÇÃO'].values:
            return account_df[account_df['DESCRIÇÃO'] == account_name][selected_column].sum()
        return 0

    def find_missing_accounts(bal_df, base_accounts, account_column):
        return set(bal_df[account_column].unique()) - set(base_accounts)

    def get_movements(df, account_column, debit_column, credit_column):
        df['MOVIMENTO'] = df[debit_column] - df[credit_column]
        return df.groupby(account_column)['MOVIMENTO'].sum()

    def calculate_dre(base_df, movements_2, classification_column, dre_column):
        """Calcula os valores da DRE a partir das classificações fornecidas."""
        def get_classification_sum(df, classification):
            """Retorna a soma dos movimentos para uma classificação específica."""
            accounts = base_df[base_df[dre_column] == classification]['Conta'].unique()
            return df[df.index.isin(accounts)].sum()
        
        dre_data = {
            "Receita líquida de vendas": get_classification_sum(movements_2, "Receita líquida de vendas")*-1,
            "Custo dos produtos vendidos": get_classification_sum(movements_2, "Custo dos produtos vendidos")*-1,
            "Lucro Bruto": None,
            "Despesas Operacionais": get_classification_sum(movements_2, "Despesas Operacionais")*-1,
            "Despesas com Pessoal": get_classification_sum(movements_2, "Despesas com Pessoal")*-1,
            "Despesas Tributárias": get_classification_sum(movements_2, "Despesas Tributárias")*-1,
            "Outras receitas / (despesas) operacionais": get_classification_sum(movements_2, "Outras receitas / (despesas) operacionais")*-1,
            "EBITDA": None,
            "Depreciação e Amortização": get_classification_sum(movements_2, "Depreciação e Amortização")*-1,
            "Outras receitas / (despesas) não operacionais": get_classification_sum(movements_2, "Outras receitas / (despesas) não operacionais")*-1,
            "Lucro/(Prejuízo) antes do resultado financeiro": None,
            "Resultado financeiro": get_classification_sum(movements_2, "Resultado financeiro")*-1,
            "Lucro antes dos impostos": None,
            "Imposto de renda e contribuição social": get_classification_sum(movements_2, "Imposto de renda e contribuição social")*-1,
            "Lucro/(Prejuízo) líquido do exercício": None
        }

        dre_data["Lucro Bruto"] = dre_data["Receita líquida de vendas"] + dre_data["Custo dos produtos vendidos"]
        dre_data["EBITDA"] = (dre_data["Lucro Bruto"] + 
                            dre_data["Despesas Operacionais"] + 
                            dre_data["Despesas com Pessoal"] +
                            dre_data["Despesas Tributárias"] +
                            dre_data["Outras receitas / (despesas) operacionais"])
        
        dre_data["Lucro/(Prejuízo) antes do resultado financeiro"] = (dre_data["EBITDA"] + 
                                                                    dre_data["Depreciação e Amortização"] + 
                                                                    dre_data["Outras receitas / (despesas) não operacionais"])
        
        dre_data["Lucro antes dos impostos"] = dre_data["Lucro/(Prejuízo) antes do resultado financeiro"] + dre_data["Resultado financeiro"]
        dre_data["Lucro/(Prejuízo) líquido do exercício"] = dre_data["Lucro antes dos impostos"] - dre_data["Imposto de renda e contribuição social"]

        dre_df =  pd.DataFrame(dre_data, index=["Valor"]).transpose()

        # Adicionar cálculo de variação vertical
        dre_df["Análise Vertical"] = (dre_df["Valor"] / dre_data["Receita líquida de vendas"]) * 100

        # Arredondar para 2 casas decimais
        dre_df = dre_df.round(2)

        return dre_df



    def main():
        st.title('Relatorios Compasso: Empresa Individual')
        
        #bal_1_file = st.file_uploader("Carregar balancete do ano anterior", type=['xlsx'])
        bal_2_file = st.file_uploader("Carregar balancete atual", type=['xlsx'])
        base_file = st.file_uploader("Carregar arquivo de dados base", type=['xlsx'])


        if bal_2_file and base_file:
            bal_2 = load_data(bal_2_file)
            base_data = load_data(base_file)
            bal_1 = load_data(base_file, sheet_name="Balancete Ant")

            if base_data is not None:
                account_column = 'Conta'
                description_column = 'Classificação'
                segregation_column = 'Segregação'
                index_column = 'Indice'
                Notas_column = 'Descrição da Conta'
                dre_column = 'DRE'
                debit_column_2 = st.selectbox("Escolha a coluna de Débito para o balancete atual", bal_2.columns)
                credit_column_2 = st.selectbox("Escolha a coluna de Crédito para o balancete atual", bal_2.columns)
                selected_column_1 = st.selectbox("Escolha a coluna de saldo para o balancete do ano anterior", bal_1.columns)
                selected_column_2 = st.selectbox("Escolha a coluna de saldo para o balancete atual", bal_2.columns)

                # Verificar se as colunas definidas estão presentes nos DataFrames
                missing_columns = [col for col in [account_column, description_column, segregation_column] if col not in base_data.columns]
                if missing_columns:
                    st.error(f"As seguintes colunas estão ausentes no arquivo base: {', '.join(missing_columns)}")
                    return
                
                # Verificar se as colunas definidas estão presentes nos DataFrames
                missing_columns = [col for col in [account_column, description_column, segregation_column, index_column] if col not in base_data.columns]
                if missing_columns:
                    st.error(f"As seguintes colunas estão ausentes no arquivo base: {', '.join(missing_columns)}")
                    return

                # Verificar contas ausentes nos balancetes
                base_accounts = base_data[account_column].unique()
                missing_in_bal_1 = find_missing_accounts(bal_1, base_accounts, account_column)
                missing_in_bal_2 = find_missing_accounts(bal_2, base_accounts, account_column)

                if missing_in_bal_1 or missing_in_bal_2:
                    missing_accounts_message = []
                    if missing_in_bal_1:
                        missing_accounts_message.append(f"Contas ausentes no arquivo base, mas presentes no balancete do ano anterior: {', '.join(map(str, sorted(missing_in_bal_1)))}")
                    if missing_in_bal_2:
                        missing_accounts_message.append(f"Contas ausentes no arquivo base, mas presentes no balancete atual: {', '.join(map(str, sorted(missing_in_bal_2)))}")
                    st.warning(" ".join(missing_accounts_message))

                interim_results = []
                for idx, row in base_data.iterrows():
                    account_number = row[account_column]
                    description = row[description_column]
                    segregation = row[segregation_column]
                    index_value = row[index_column]
                    notas = row[Notas_column]
                    balance_1 = get_account_balance(bal_1, account_number, account_column, selected_column_1)
                    balance_2 = get_account_balance(bal_2, account_number, account_column, selected_column_2)

                    interim_results.append({
                        "CONTA": account_number,
                        "DESCRIÇÃO": description,
                        "SEGREGAÇÃO": segregation,
                        "INDICE": index_value,
                        "NOTAS": notas,
                        "SALDO BALANCETE ANTERIOR": balance_1,
                        "SALDO BALANCETE ATUAL": balance_2
                    })

                interim_df = pd.DataFrame(interim_results)

                # Agrupar e sumarizar os resultados por descrição
                grouped_df = interim_df.groupby(['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO']).agg({
                'SALDO BALANCETE ANTERIOR': 'sum',
                'SALDO BALANCETE ATUAL': 'sum'
                }).reset_index()

                grouped_df['DIFERENÇA'] = grouped_df['SALDO BALANCETE ATUAL'] - grouped_df['SALDO BALANCETE ANTERIOR']
                grouped_df['%'] = ((grouped_df['SALDO BALANCETE ANTERIOR'] - grouped_df['SALDO BALANCETE ATUAL']) / grouped_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar para Ativo e Passivo
                active_df = grouped_df[grouped_df['SEGREGAÇÃO'].str.contains('Ativo')].sort_values('INDICE')
                passive_df = grouped_df[grouped_df['SEGREGAÇÃO'].str.contains('Passivo')].sort_values('INDICE')

                # Somar os valores totais de Ativo e Passivo
                total_active = active_df['SALDO BALANCETE ANTERIOR'].sum()
                total_passive = passive_df['SALDO BALANCETE ANTERIOR'].sum()
                difference_ant = total_active - total_passive

                # Somar os valores totais de Ativo e Passivo
                total_active = active_df['SALDO BALANCETE ATUAL'].sum()
                total_passive = passive_df['SALDO BALANCETE ATUAL'].sum()
                difference_at = total_active - total_passive

                st.write("Comparação de Saldos Ativo:")
                active_df = active_df.round(2)
                st.dataframe(active_df)

                st.write("Comparação de Saldos Passivo:")
                passive_df = passive_df.round(2)
                st.dataframe(passive_df)

                st.write("**Conferencia BL Anterior:** R$ {:.2f}".format(difference_ant))
                st.write("**Conferencia BL Atual:** R$ {:.2f}".format(difference_at))


                classified_area_dataframes = {}

                # Filtrar apenas as contas que possuem classificação não nula
                classified_accounts = base_data.dropna(subset=[description_column])

                # Adicionar informações e saldos aos DataFrames por área/classificação
                for _, row in classified_accounts.iterrows():
                    area = row[description_column]
                    account_number = row[account_column]
                    description = row[Notas_column]
                    index_value = row[index_column]
                    segregation_area = row[segregation_column]

                    balance = get_account_balance(bal_2, account_number, account_column, selected_column_2)
                    balance_ant = get_account_balance(bal_1, account_number, account_column, selected_column_1)

                    if area not in classified_area_dataframes:
                        classified_area_dataframes[area] = []

                    classified_area_dataframes[area].append({
                        "Indice": index_value,
                        #"Segregação" : segregation_area,
                        #"Conta": account_number,
                        "Descrição": description,
                        f"SALDO ANT": balance_ant,
                        f"SALDO ATUAL": balance
                    })

                # DRE
                # Verificar se as colunas definidas estão presentes nos DataFrames
                missing_columns = [col for col in [account_column, description_column, segregation_column, index_column, dre_column] if col not in base_data.columns]
                if missing_columns:
                    st.error(f"As seguintes colunas estão ausentes no arquivo base: {', '.join(missing_columns)}")
                    return
                
                # Calcular movimentos para cada balancete
                movements_2 = get_movements(bal_2, account_column, debit_column_2, credit_column_2)

                # Calcular a DRE
                dre_df = calculate_dre(base_data, movements_2, account_column, dre_column)
        

                st.write("Demonstração do Resultado do Exercício (DRE):")
                st.dataframe(dre_df)

                # Consolidação dos saldos de todas as filiais em colunas únicas
                dfs_to_download = {
                    "DRE": dre_df,
                    "Ativo Individual": active_df,
                    "Passivo Individual": passive_df
                }

                # Notas Explicativas:
                for area, data in sorted(classified_area_dataframes.items(), key=lambda x: min([d['Indice'] for d in x[1]])):
                    area_df = pd.DataFrame(data)
                    area_df['Variação'] = area_df['SALDO ATUAL'] - area_df['SALDO ANT']
                    area_df['%'] = ((area_df['Variação'] / area_df['SALDO ANT']).replace([np.inf, -np.inf], 0)) * 100

                    # Ordenar DataFrame pela coluna 'Indice'
                    area_df = area_df.sort_values(by='Indice')

                    st.write(f"{area}")
                    st.dataframe(area_df)

                    # Adicionando ao dicionário de DataFrames para download
                    dfs_to_download[area] = area_df
            
            excel_data = create_excel_download_link(dfs_to_download, "Relatorios_Compasso.xlsx")
            st.download_button(label="Baixar Relatorio", data=excel_data, file_name="Relatorios_Compasso.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


    main()