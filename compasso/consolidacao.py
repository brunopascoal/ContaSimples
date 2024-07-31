import streamlit as st
import pandas as pd
import numpy as np
import io
import re

def run_compasso_consolidação_app():
    
    def load_data(file, sheet_name=None):
        """ Carrega dados de um arquivo Excel. """
        data = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(data, dict):
            # Se `data` for um dicionário, pegamos a primeira planilha
            data = next(iter(data.values()))
        return data

    def convert_to_numeric(column):
        """ Tenta converter uma coluna para numérico, tratando erros. """
        return pd.to_numeric(column, errors='coerce')

    def get_account_balance(account_df, account_number, account_column, selected_column):
        """Calcula o saldo total para uma conta específica usando a coluna selecionada."""
        if account_number in account_df[account_column].values:
            account_data = account_df[account_df[account_column] == account_number]
            if selected_column in account_data.columns:
                account_data[selected_column] = convert_to_numeric(account_data[selected_column])
                return account_data[selected_column].sum()
            else:
                return 0
        else:
            return 0
        
    def find_missing_accounts(df, base_accounts, account_column):
        """Encontra contas que estão no DataFrame `df` mas não em `base_accounts`."""
        bal_accounts = df[account_column].unique()
        return [account for account in bal_accounts if account not in base_accounts]

    def create_excel_download_link(dataframes, filename):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in dataframes.items():
                # Truncar o nome da folha para 31 caracteres e remover caracteres inválidos
                clean_sheet_name = re.sub(r'[\[\]\*:/\\?]', '', sheet_name)[:31]
                df.to_excel(writer, index=False, sheet_name=clean_sheet_name)
        output.seek(0)
        return output

    def main():
        st.title('Relatorios Compasso: Grupo de Empresas')

        num_filiais = st.number_input("Quantas filiais?", min_value=1, step=1)
        bal_files = []
        base_files = []
        balancete_ant_files = []
        branch_names = []
        selected_columns = []

        for i in range(int(num_filiais)):
            bal_file = st.file_uploader(f"Carregar balancete da filial {i+1}", type=['xlsx'])
            base_file = st.file_uploader(f"Carregar arquivo de dados base para o balancete da filial {i+1}", type=['xlsx'])
            if bal_file and base_file:
                bal_df = load_data(bal_file)
                base_df = load_data(base_file)
                balancete_ant_df = load_data(base_file, sheet_name="Balancete Ant")
                
                bal_files.append(bal_df)
                base_files.append(base_df)
                balancete_ant_files.append(balancete_ant_df)
                
                branch_name = bal_file.name.replace('.xlsx', '')
                branch_names.append(branch_name)
                selected_column = st.selectbox(f"Escolha a coluna de saldo para o balancete da {branch_name}", bal_df.columns)
                selected_columns.append(selected_column)
                selected_column_ant = 'Saldo atual'

        if len(bal_files) == num_filiais and len(base_files) == num_filiais:
            account_column = 'Conta'
            description_column = 'Classificação'
            index_column = 'Indice'
            segregation_column = 'Segregação'
            description_cont = 'Descrição da Conta'

            results = []
            results_ant = []
            missing_accounts_messages = []

            classified_area_dataframes = {}

            for i in range(int(num_filiais)):
                base = base_files[i]
                bal = bal_files[i]
                balancete_ant = balancete_ant_files[i]
                selected_column = selected_columns[i]
                branch_name = branch_names[i]

                # Verificar contas ausentes nos balancetes
                base_accounts = base[account_column].unique()
                missing_in_bal_1 = find_missing_accounts(balancete_ant, base_accounts, account_column)
                missing_in_bal_2 = find_missing_accounts(bal, base_accounts, account_column)

                if missing_in_bal_1 or missing_in_bal_2:
                    if missing_in_bal_1:
                        missing_accounts_messages.append(f"Contas ausentes no arquivo base, mas presentes no balancete 'Balancete Ant' da {branch_name}: {', '.join(map(str, sorted(missing_in_bal_1)))}")
                    if missing_in_bal_2:
                        missing_accounts_messages.append(f"Contas ausentes no arquivo base, mas presentes no balancete atual da {branch_name}: {', '.join(map(str, sorted(missing_in_bal_2)))}")


                # Filtrar apenas as contas que possuem classificação não nula
                classified_accounts = base.dropna(subset=[description_column])

                # Adicionar informações e saldos aos DataFrames por área/classificação
                for _, row in classified_accounts.iterrows():
                    area = row[description_column]
                    account_number = row[account_column]
                    description = row[description_cont]
                    index_value = row[index_column]
                    segregation_area = row[segregation_column]

                    balance = get_account_balance(bal, account_number, account_column, selected_column)
                    balance_ant = get_account_balance(balancete_ant, account_number, account_column, selected_column_ant)

                    if area not in classified_area_dataframes:
                        classified_area_dataframes[area] = []

                    classified_area_dataframes[area].append({
                        "Indice": index_value,
                        "Arquivo de Origem": f"Balancete {branch_name}",
                        "Segregação" : segregation_area,
                        "Conta": account_number,
                        "Descrição": description,
                        f"SALDO {branch_name} ANT": balance_ant,
                        f"SALDO {branch_name}": balance
                    })


                for idx, row in base.iterrows():
                    account_number = row[account_column]
                    description = row[description_column]
                    index_value = row[index_column]
                    segregation = row[segregation_column]
                    balance = get_account_balance(bal, account_number, account_column, selected_column)
                    balance_ant = get_account_balance(balancete_ant, account_number, account_column, selected_column_ant)

                    results.append({
                        "CONTA": account_number,
                        "DESCRIÇÃO": description,
                        "INDICE": index_value,
                        "SEGREGAÇÃO": segregation,
                        f"SALDO {branch_name}": balance
                    })

                    results_ant.append({
                        "CONTA": account_number,
                        "DESCRIÇÃO": description,
                        "INDICE": index_value,
                        "SEGREGAÇÃO": segregation,
                        f"SALDO {branch_name} ANT": balance_ant
                    })

            results_df = pd.DataFrame(results)
            results_ant_df = pd.DataFrame(results_ant)

            # Mostrar mensagens de contas ausentes
            if missing_accounts_messages:
                st.warning(" ".join(missing_accounts_messages))

            # Agrupar e sumarizar os resultados por descrição
            group_columns = ['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO'] + [f'SALDO {branch_name}' for branch_name in branch_names]
            grouped_df = results_df.groupby(['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO']).agg({col: 'sum' for col in group_columns if 'SALDO' in col}).reset_index()

            group_columns_ant = ['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO'] + [f'SALDO {branch_name} ANT' for branch_name in branch_names]
            grouped_ant_df = results_ant_df.groupby(['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO']).agg({col: 'sum' for col in group_columns_ant if 'SALDO' in col}).reset_index()

            grouped_df['Consolidado Atual'] = grouped_df[[f'SALDO {branch_name}' for branch_name in branch_names]].sum(axis=1)
            grouped_ant_df['Consolidado Ant'] = grouped_ant_df[[f'SALDO {branch_name} ANT' for branch_name in branch_names]].sum(axis=1)

        

            st.write("Ativo Individual:")
            active_df = grouped_df[grouped_df['SEGREGAÇÃO'].str.contains('Ativo')].sort_values('INDICE')
            active_df = active_df.round(2)
            st.dataframe(active_df)

            st.write("Passivo Individual:")
            passive_df = grouped_df[grouped_df['SEGREGAÇÃO'].str.contains('Passivo')].sort_values('INDICE')
            saldo_columns = [col for col in passive_df.columns if 'SALDO' in col] + ['Consolidado Atual']
            passive_df[saldo_columns] = passive_df[saldo_columns] * -1
            passive_df = passive_df.round(2)
            st.dataframe(passive_df)


            # Criar DataFrames comparativos para Ativo e Passivo
            comparativo_df = pd.merge(grouped_df[['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO', 'Consolidado Atual']],
                                    grouped_ant_df[['INDICE', 'Consolidado Ant']],
                                    on='INDICE', how='left')
            
            # Reorganizar as colunas para que 'Consolidado Ant' venha antes de 'Consolidado Atual'
            comparativo_df = comparativo_df[['INDICE', 'DESCRIÇÃO', 'SEGREGAÇÃO', 'Consolidado Ant', 'Consolidado Atual']]

            # Calculo de Variações
            comparativo_df['Variação'] = comparativo_df['Consolidado Atual'] - comparativo_df['Consolidado Ant']
            comparativo_df['%'] = ((comparativo_df['Variação'] / comparativo_df['Consolidado Ant']).replace([np.inf, -np.inf], 0)) * 100


            # Filtrar para Ativo e Passivo nos dados comparativos
            comparativo_ativo_df = comparativo_df[comparativo_df['SEGREGAÇÃO'].str.contains('Ativo')].sort_values('INDICE')
            comparativo_passivo_df = comparativo_df[comparativo_df['SEGREGAÇÃO'].str.contains('Passivo')].sort_values('INDICE')

            # Multiplicar os saldos do passivo por -1, incluindo a coluna de variação
            saldo_columns = ['Consolidado Ant', 'Consolidado Atual', 'Variação']
            comparativo_passivo_df[saldo_columns] = comparativo_passivo_df[saldo_columns] * -1

            st.write("Ativo Consolidado:")
            comparativo_ativo_df = comparativo_ativo_df.round(2)
            st.dataframe(comparativo_ativo_df)

            st.write("Passivo Consolidado:")
            comparativo_passivo_df = comparativo_passivo_df.round(2)
            st.dataframe(comparativo_passivo_df)

            # Consolidação dos saldos de todas as filiais em colunas únicas
            dfs_to_download = {
                "Ativo Consolidado": comparativo_ativo_df,
                "Passivo Consolidado": comparativo_passivo_df,
                "Ativo Individual": active_df,
                "Passivo Individual": passive_df
            }
            
            for area, data in sorted(classified_area_dataframes.items(), key=lambda x: min([d['Indice'] for d in x[1]])):
                area_df = pd.DataFrame(data)
                # Verificar se todas as colunas de saldo estão presentes
                saldo_atual_columns = [f"SALDO {branch_name}" for branch_name in branch_names if f"SALDO {branch_name}" in area_df.columns]
                saldo_ant_columns = [f"SALDO {branch_name} ANT" for branch_name in branch_names if f"SALDO {branch_name} ANT" in area_df.columns]
                
                if saldo_ant_columns:
                    area_df['Saldo Anterior'] = area_df[saldo_ant_columns].sum(axis=1)
                else:
                    area_df['Saldo Anterior'] = 0

                if saldo_atual_columns:
                    area_df['Saldo Atual'] = area_df[saldo_atual_columns].sum(axis=1)
                else:
                    area_df['Saldo Atual'] = 0


                # Multiplicar os valores das colunas "Saldo Atual" e "Saldo Anterior" por -1 se a "Segregação" contiver "Passivo"
                if area_df['Segregação'].str.contains('Passivo', case=False).any():
                    area_df['Saldo Anterior'] = area_df['Saldo Anterior'] * -1
                    area_df['Saldo Atual'] = area_df['Saldo Atual'] * -1

                # Remover colunas de saldo individual após consolidação
                area_df = area_df.drop(columns=saldo_atual_columns + saldo_ant_columns, errors='ignore')
                
                area_df['Variação'] = area_df['Saldo Atual'] - area_df['Saldo Anterior']
                area_df['%'] = ((area_df['Variação'] / area_df['Saldo Anterior']).replace([np.inf, -np.inf], 0)) * 100

                # Ordenar DataFrame pela coluna 'Indice'
                area_df = area_df.sort_values(by='Indice')
                # Ordenar DataFrame pela coluna 
                area_df = area_df.sort_values(by='Arquivo de Origem')

                st.write(f"{area}")
                st.dataframe(area_df)
                
                # Adicionando ao dicionário de DataFrames para download
                dfs_to_download[area] = area_df

            excel_data = create_excel_download_link(dfs_to_download, "Consolidacao_Compasso.xlsx")
            st.download_button(label="Baixar Planilhas Consolidadas", data=excel_data, file_name="Consolidacao_Compasso.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    main()
