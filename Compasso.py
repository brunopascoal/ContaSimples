import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook

def run_compasso_app():

    def load_data(file):
        return pd.read_excel(file)

    def save_excel(dataframes, sheet_names, file_path='Dataframes_Exportados.xlsx'):
        """Salva os dataframes em um arquivo Excel, cada um em uma aba separada."""
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for dataframe, sheet_name in zip(dataframes, sheet_names):
                dataframe.to_excel(writer, sheet_name=sheet_name, index=True)
        return file_path

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
        st.title('Relatorios Compasso')
        
        bal_1_file = st.file_uploader("Carregar balancete do ano anterior", type=['xlsx'])
        bal_2_file = st.file_uploader("Carregar balancete atual", type=['xlsx'])
        base_file = st.file_uploader("Carregar arquivo de dados base", type=['xlsx'])

        if bal_1_file and bal_2_file and base_file:
            bal_1 = load_data(bal_1_file)
            bal_2 = load_data(bal_2_file)
            base_data = load_data(base_file)

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
                grouped_df['%'] = ((grouped_df['DIFERENÇA'] / grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100

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

                # Notas Explicativas
                Notas_df = interim_df.groupby(['CONTA','DESCRIÇÃO', 'NOTAS']).agg({
                'SALDO BALANCETE ANTERIOR': 'sum',
                'SALDO BALANCETE ATUAL': 'sum'
                }).reset_index()

                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                caixa_e_equivalentes_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Caixa')].sort_values('CONTA')

                st.write("Caixa e Equivalentes de Caixa:")
                caixa_e_equivalentes_df = caixa_e_equivalentes_df.round(2)
                st.dataframe(caixa_e_equivalentes_df)

                # Contas a receber
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar
                Contas_a_receber_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Contas a receber')].sort_values('CONTA')

                st.write("Contas a Receber:")
                Contas_a_receber_df = Contas_a_receber_df.round(2)
                st.dataframe(Contas_a_receber_df)

                # Impostos a recuperar
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar
                Impostos_a_recuperar_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Impostos a recuperar')].sort_values('CONTA')

                st.write("Impostos a recuperar:")
                Impostos_a_recuperar_df = Impostos_a_recuperar_df.round(2)
                st.dataframe(Impostos_a_recuperar_df)

                # Adiantamentos
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Adiantamentos_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Adiantamentos')].sort_values('CONTA')

                st.write("Adiantamentos:")
                Adiantamentos_df = Adiantamentos_df.round(2)
                st.dataframe(Adiantamentos_df)

                # Despesas antecipadas
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Despesas_antecipadas_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Despesas antecipadas')].sort_values('CONTA')

                st.write("Despesas antecipadas:")
                Despesas_antecipadas_df = Despesas_antecipadas_df.round(2)
                st.dataframe(Despesas_antecipadas_df)

                # Consórcios
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Consórcios_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Consórcios')].sort_values('CONTA')

                st.write("Consórcios:")
                Consórcios_df = Consórcios_df.round(2)
                st.dataframe(Consórcios_df)

                # Imobilizado
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Imobilizado_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Imobilizado')].sort_values('CONTA')

                st.write("Imobilizado:")
                Imobilizado_df = Imobilizado_df.round(2)
                st.dataframe(Imobilizado_df)

                # Empréstimos e financiamentos
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Empréstimos_e_financiamentos_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Empréstimos e financiamentos')].sort_values('CONTA')

                st.write("Empréstimos e financiamentos:")
                Empréstimos_e_financiamentos_df = Empréstimos_e_financiamentos_df.round(2)
                st.dataframe(Empréstimos_e_financiamentos_df)

                # Contas a Pagar
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Contas_a_Pagar_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Contas a pagar')].sort_values('CONTA')

                st.write("Contas a Pagar:")
                Contas_a_Pagar_df = Contas_a_Pagar_df.round(2)
                st.dataframe(Contas_a_Pagar_df)

                
                # Obrigações trabalhistas
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Obrigações_trabalhistas_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Obrigações trabalhistas')].sort_values('CONTA')

                st.write("Obrigações trabalhistas:")
                Obrigações_trabalhistas_df = Obrigações_trabalhistas_df.round(2)
                st.dataframe(Obrigações_trabalhistas_df)

                # Obrigações tributárias
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Obrigações_tributárias_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Obrigações tributárias')].sort_values('CONTA')

                st.write("Obrigações tributárias:")
                Obrigações_tributárias_df = Obrigações_tributárias_df.round(2)
                st.dataframe(Obrigações_tributárias_df)

                # Outros débitos
                Notas_df['DIFERENÇA'] = Notas_df['SALDO BALANCETE ATUAL'] - Notas_df['SALDO BALANCETE ANTERIOR']
                Notas_df['%'] = ((Notas_df['SALDO BALANCETE ANTERIOR'] - Notas_df['SALDO BALANCETE ATUAL']) / Notas_df['SALDO BALANCETE ATUAL']).replace([np.inf, -np.inf], 0) * 100

                # Filtrar 
                Outros_débitos_df = Notas_df[Notas_df['DESCRIÇÃO'].str.contains('Outros débitos')].sort_values('CONTA')

                st.write("Outros débitos:")
                Outros_débitos_df = Outros_débitos_df.round(2)
                st.dataframe(Outros_débitos_df)

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


            if st.button('Exportar Excel'):
                        file_path = save_excel(
                            [active_df, passive_df, caixa_e_equivalentes_df, Contas_a_receber_df, Impostos_a_recuperar_df, Adiantamentos_df, Despesas_antecipadas_df, Consórcios_df,
                            Imobilizado_df, Empréstimos_e_financiamentos_df, Contas_a_Pagar_df, Obrigações_trabalhistas_df, Obrigações_tributárias_df, Outros_débitos_df, dre_df],  # Lista dos seus DataFrames
                            ["Ativo", "Passivo", "Caixa e Equivalentes", "Contas a Receber", "Impostos a Recuperar", "Adiantamentos", "Despesas Antecipadas", "Consórcios",
                            "Imobilizado", "Empréstimos e Financiamentos", "Contas a Pagar", "Obrigações trabalhistas", "Obrigações tributárias", "Outros débitos", "DRE" ]  # Nomes das abas correspondentes
                        )
                        st.success('Exportado com sucesso! Clique abaixo para baixar o arquivo.')
                        with open(file_path, "rb") as file:
                            btn = st.download_button(
                                label="Baixar Excel",
                                data=file,
                                file_name='Relatorio Compasso.xlsx',
                                mime='application/vnd.ms-excel'
                            )
    main()
