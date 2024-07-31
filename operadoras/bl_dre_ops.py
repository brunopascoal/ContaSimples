import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook


def run_bl_e_dre_app():

    def load_data(file, sheet_name=None):
        data = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(data, dict):
            # Retorna o primeiro dataframe do dicionário, se houver múltiplas planilhas
            return next(iter(data.values()))
        return data

    def save_excel(dataframes, sheet_names, file_path='Dataframes_Exportados.xlsx'):
        """Salva os dataframes em um arquivo Excel, cada um em uma aba separada."""
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for dataframe, sheet_name in zip(dataframes, sheet_names):
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
        return file_path

    def get_account_balance(account_df, account_number, selected_column):
        if account_number in account_df['CONTA'].values:
            return account_df[account_df['CONTA'] == account_number][selected_column].sum()
        else:
            return None

    def calculate_liquidity_ratio(disponible, financial_applications, current_liabilities):
        if current_liabilities == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return (disponible + financial_applications) / (current_liabilities*-1)

    def calculate_current_ratio(current_assets, current_liabilities):
        if current_liabilities == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return current_assets / (current_liabilities*-1)

    def calculate_quick_ratio(current_assets, inventories, current_liabilities):
        if current_liabilities == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return (current_assets - inventories) / (current_liabilities*-1)

    def calculate_endividamento(total_assets, total_liabilities):
        if total_liabilities == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return total_assets / (total_liabilities*-1)

    def calculate_general_liquidity_ratio(current_assets, long_term_receivables, current_liabilities, long_term_debt):
        total_assets = current_assets + long_term_receivables
        total_liabilities = current_liabilities + long_term_debt
        if total_liabilities == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return total_assets / (total_liabilities*-1)

    def annualize_balance(balance, current_date, previous_december):
        days_difference = (current_date - previous_december).days
        if days_difference > 0 and days_difference < 365:
            annualized_balance = (balance / (days_difference / 30)) * 12
            return annualized_balance
        return balance

    def calculate_ebitda(resultado_bruto, contraprestacoes_efetivas):
        return resultado_bruto + contraprestacoes_efetivas

    def calculate_ebitda_index(ebitda, contraprestacoes_efetivas):
        if contraprestacoes_efetivas == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return (ebitda / contraprestacoes_efetivas)

    def calculate_net_margin(resultado_liquido, contraprestacoes_efetivas):
        if contraprestacoes_efetivas == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return (resultado_liquido / contraprestacoes_efetivas) * 100

    def calculate_revenue_vs_cost(eventos_indenizaveis, contraprestacoes_efetivas):
        if contraprestacoes_efetivas == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return ((eventos_indenizaveis * -1) / contraprestacoes_efetivas) * 100

    def calculate_expenses_vs_financial(despesas_financeiras, emprestimos_pagar):
        if emprestimos_pagar == 0:
            return 'Indefinido'  # Evitar divisão por zero
        return (despesas_financeiras * -1) / (emprestimos_pagar * -1)

    # Função de formatação
    def format_number(x):
        return "{:,.2f}".format(x)

    def main():
        st.title('Revisão Analitica das Operadoras de Saude')
        
        # Seleção do tipo de visita
        visit_type = st.selectbox("Selecione o tipo da visita:", ["Preliminar", "Final"])
        
        bal_1_file = st.file_uploader("Carregar balancete do ano anterior", type=['xlsx'])
        bal_2_file = st.file_uploader("Carregar balancete atual", type=['xlsx'])
        ops_file = st.file_uploader("Carregar arquivo BL e DRE OPS", type=['xlsx'])

        if bal_1_file and bal_2_file and ops_file:
            bal_1 = load_data(bal_1_file)
            bal_2 = load_data(bal_2_file)
            ops_data = load_data(ops_file)
            dre_data = load_data(ops_file, sheet_name='DRE')

            # Permitir ao usuário escolher a coluna de saldo para cada balancete
            selected_column_1 = st.selectbox("Escolha a coluna de saldo para o balancete do ano anterior", bal_1.columns)
            selected_column_2 = st.selectbox("Escolha a coluna de saldo para o balancete atual", bal_2.columns)

            # Perguntar ao usuário qual é a data do balancete atual e a data de dezembro do ano anterior
            if visit_type == "Preliminar":
                current_date = st.date_input("Escolha a data do balancete atual")
                previous_december = st.date_input("Escolha a data de dezembro do ano anterior")

            # Processamento de comparação de saldos
            balance_details = {}
            interim_results = []

            # Pergunta para o usuário inserir o valor de variação máxima permitida
            max_variation = st.number_input("Materialidade")
            percent_variation_threshold = 15.0

            for idx, row in ops_data.iterrows():
                index = row["ÍNDICE"]
                account_number = row['CONTA']
                description = row['DESCRIÇÃO']
                balance_1 = get_account_balance(bal_1, account_number, selected_column_1)
                balance_2 = get_account_balance(bal_2, account_number, selected_column_2)
                
                # Incluir apenas se a conta tem saldo em pelo menos um dos balancetes
                if balance_1 != 0 or balance_2 != 0:
                    interim_results.append({
                        "ÍNDICE": index,
                        "CONTA": account_number,
                        "DESCRIÇÃO": description,
                        "SALDO BALANCETE ATUAL": balance_2,
                        "SALDO BALANCETE ANTERIOR": balance_1
                    })
                    balance_details[description] = [balance_1, balance_2]

            # Criar DataFrame a partir dos resultados intermediários
            interim_df = pd.DataFrame(interim_results).sort_values(by="ÍNDICE")

            # Agrupar e sumarizar os resultados por descrição
            grouped_df = interim_df.groupby(['ÍNDICE', 'DESCRIÇÃO']).agg({
                'SALDO BALANCETE ATUAL': 'sum',
                'SALDO BALANCETE ANTERIOR': 'sum'
            }).reset_index()

            grouped_df['DIFERENÇA'] = grouped_df['SALDO BALANCETE ATUAL'] - grouped_df['SALDO BALANCETE ANTERIOR']
            grouped_df['%'] = ((grouped_df['DIFERENÇA'] / grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100
            filtered_df = grouped_df[(grouped_df['SALDO BALANCETE ANTERIOR'] != 0) | (grouped_df['SALDO BALANCETE ATUAL'] != 0)]

            # Adicionar coluna "VERIFICAR" baseada nas duas condições
            filtered_df['VERIFICAR'] = np.where(
                (abs(filtered_df['DIFERENÇA']) > max_variation) |
                (abs(filtered_df['%']) > percent_variation_threshold),
                "Verificar", 
                ""
            )

            filtered_df['DESCRIÇÃO'] = filtered_df['DESCRIÇÃO'].str.strip()

            st.write("BL:")
            filtered_df = filtered_df.round(2)
            st.dataframe(filtered_df)

            # Processamento da DRE
            dre_results = []

            for idx, row in dre_data.iterrows():
                index = row["ÍNDICE"]
                account_number = row['CONTA']
                description = row['DESCRIÇÃO']
                balance_1 = get_account_balance(bal_1, account_number, selected_column_1)
                balance_2 = get_account_balance(bal_2, account_number, selected_column_2)
                
                if balance_1 is not None or balance_2 is not None:
                    dre_results.append({
                        "ÍNDICE": index,
                        "CONTA": account_number,
                        "DESCRIÇÃO": description,
                        "SALDO BALANCETE ATUAL": (balance_2 if balance_2 is not None else 0) * -1,
                        "SALDO BALANCETE ANTERIOR": (balance_1 if balance_1 is not None else 0) * -1
                    })

            dre_df = pd.DataFrame(dre_results)

            # Verificar se a coluna 'ÍNDICE' existe antes de ordenar
            if 'ÍNDICE' in dre_df.columns:
                dre_df = dre_df.sort_values(by="ÍNDICE")
            else:
                st.error("Coluna 'ÍNDICE' não encontrada na DRE. Verifique os dados de entrada.")

            # Agrupar e sumarizar os resultados por descrição
            dre_grouped_df = dre_df.groupby(['ÍNDICE', 'DESCRIÇÃO']).agg({
                'SALDO BALANCETE ATUAL': 'sum',
                'SALDO BALANCETE ANTERIOR': 'sum'
            }).reset_index()

            # Calcular subtotais
            dre_grouped_df = dre_grouped_df.round(2)

            dre_grouped_df['DESCRIÇÃO'] = dre_grouped_df['DESCRIÇÃO'].str.strip()

            # Anualizar os saldos se o tipo de visita for "Preliminar"
            if visit_type == "Preliminar":
                dre_grouped_df['SALDO ANUALIZADO'] = dre_grouped_df['SALDO BALANCETE ATUAL'].apply(lambda x: annualize_balance(x, current_date, previous_december))

                    # RESULTADO DAS OPERAÇÕES COM PLANOS DE ASSISTÊNCIA À SAÚDE
                resultado_op_saude_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_op_saude_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos'])]['SALDO BALANCETE ATUAL'].sum()
                resultado_op_saude_anualizado = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos'])]['SALDO ANUALIZADO'].sum()

                # RESULTADO BRUTO
                resultado_bruto_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_bruto_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora'])]['SALDO BALANCETE ATUAL'].sum()
                resultado_bruto_anualizado = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora'])]['SALDO ANUALIZADO'].sum()

                # RESULTADO ANTES DOS IMPOSTOS E PARTICIPAÇÕES
                resultado_antes_impostos_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora',
                    'Despesas de Comercialização',
                    'Despesas Administrativas',
                    'Resultado Financeiro Líquido',
                    'Resultado Patrimonial',
                    'Resultado com Seguro e Resseguro'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_antes_impostos_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora',
                    'Despesas de Comercialização',
                    'Despesas Administrativas',
                    'Resultado Financeiro Líquido',
                    'Resultado Patrimonial',
                    'Resultado com Seguro e Resseguro'])]['SALDO BALANCETE ATUAL'].sum()
                resultado_antes_impostos_anualizado = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora',
                    'Despesas de Comercialização',
                    'Despesas Administrativas',
                    'Resultado Financeiro Líquido',
                    'Resultado Patrimonial',
                    'Resultado com Seguro e Resseguro'])]['SALDO ANUALIZADO'].sum()

                #dre_grouped_df['DIFERENÇA'] = dre_grouped_df['SALDO ANUALIZADO'] - dre_grouped_df['SALDO BALANCETE ANTERIOR']
                #dre_grouped_df['%'] = ((dre_grouped_df['DIFERENÇA'] / dre_grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100
            
            else:
                        # RESULTADO DAS OPERAÇÕES COM PLANOS DE ASSISTÊNCIA À SAÚDE
                resultado_op_saude_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_op_saude_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos'])]['SALDO BALANCETE ATUAL'].sum()

                # RESULTADO BRUTO
                resultado_bruto_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_bruto_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora'])]['SALDO BALANCETE ATUAL'].sum()
                

                # RESULTADO ANTES DOS IMPOSTOS E PARTICIPAÇÕES
                resultado_antes_impostos_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora',
                    'Despesas de Comercialização',
                    'Despesas Administrativas',
                    'Resultado Financeiro Líquido',
                    'Resultado Patrimonial',
                    'Resultado com Seguro e Resseguro'])]['SALDO BALANCETE ANTERIOR'].sum()
                resultado_antes_impostos_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'].isin([
                    'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde',
                    'Eventos Indenizáveis Líquidos / Sinistros Retidos',
                    'Outras Receitas Operacionais de Planos de Assistência à Saúde',
                    'Receitas de Assistência à Saúde Não Relacionadas com Planos de Saúde da Operadora',
                    'Tributos Diretos de Outras Atividades de Assistência à Saúde',
                    'Outras Despesas Operacionais com Plano de Assistência à Saúde',
                    'Outras Despesas Oper. de Assist. à Saúde Não Rel. com Planos de Saúde da Operadora',
                    'Despesas de Comercialização',
                    'Despesas Administrativas',
                    'Resultado Financeiro Líquido',
                    'Resultado Patrimonial',
                    'Resultado com Seguro e Resseguro'])]['SALDO BALANCETE ATUAL'].sum()
                

                #dre_grouped_df['DIFERENÇA'] = dre_grouped_df['SALDO BALANCETE ATUAL'] - dre_grouped_df['SALDO BALANCETE ANTERIOR']
                #dre_grouped_df['%'] = ((dre_grouped_df['DIFERENÇA'] / dre_grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100

            # Adicionar subtotais ao dataframe
            if visit_type == "Preliminar":
                subtotais = pd.DataFrame({
                    'ÍNDICE': [10, 28, 40],
                    'DESCRIÇÃO': [
                        'RESULTADO DAS OPERAÇÕES COM PLANOS DE ASSISTÊNCIA À SAÚDE',
                        'RESULTADO BRUTO',
                        'RESULTADO ANTES DOS IMPOSTOS E PARTICIPAÇÕES'],
                    'SALDO BALANCETE ATUAL': [
                        resultado_op_saude_anualizado,
                        resultado_bruto_anualizado,
                        resultado_antes_impostos_anualizado],
                    'SALDO BALANCETE ANTERIOR': [
                        resultado_op_saude_anterior,
                        resultado_bruto_anterior,
                        resultado_antes_impostos_anterior],
                    'SALDO ANUALIZADO': [
                        resultado_op_saude_anualizado,
                        resultado_bruto_anualizado,
                        resultado_antes_impostos_anualizado]
                })
            else:
                subtotais = pd.DataFrame({
                    'ÍNDICE': [10, 28, 40],
                    'DESCRIÇÃO': [
                        'RESULTADO DAS OPERAÇÕES COM PLANOS DE ASSISTÊNCIA À SAÚDE',
                        'RESULTADO BRUTO',
                        'RESULTADO ANTES DOS IMPOSTOS E PARTICIPAÇÕES'],
                    'SALDO BALANCETE ATUAL': [
                        resultado_op_saude_atual,
                        resultado_bruto_atual,
                        resultado_antes_impostos_atual],
                    'SALDO BALANCETE ANTERIOR': [
                        resultado_op_saude_anterior,
                        resultado_bruto_anterior,
                        resultado_antes_impostos_anterior]
                })


            dre_grouped_df = pd.concat([dre_grouped_df, subtotais]).sort_values(by='ÍNDICE').reset_index(drop=True)

            if visit_type == "Preliminar":
                dre_grouped_df['DIFERENÇA'] = dre_grouped_df['SALDO ANUALIZADO'] - dre_grouped_df['SALDO BALANCETE ANTERIOR']
                dre_grouped_df['%'] = ((dre_grouped_df['DIFERENÇA'] / dre_grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100

            else:
                dre_grouped_df['DIFERENÇA'] = dre_grouped_df['SALDO BALANCETE ATUAL'] - dre_grouped_df['SALDO BALANCETE ANTERIOR']
                dre_grouped_df['%'] = ((dre_grouped_df['DIFERENÇA'] / dre_grouped_df['SALDO BALANCETE ANTERIOR']).replace([np.inf, -np.inf], 0)) * 100
        
            dre_grouped_df = dre_grouped_df[(dre_grouped_df['SALDO BALANCETE ANTERIOR'] != 0) | (dre_grouped_df['SALDO BALANCETE ATUAL'] != 0)]
            
            # Adicionar coluna "VERIFICAR" baseada nas duas condições
            dre_grouped_df['VERIFICAR'] = np.where(
                (abs(dre_grouped_df['DIFERENÇA']) > max_variation) |
                (abs(dre_grouped_df['%']) > percent_variation_threshold),
                "Verificar", 
                ""
            )

            st.write("DRE:")
            if visit_type == "Preliminar":
                    st.dataframe(dre_grouped_df)
            else:
                if 'SALDO ANUALIZADO' in dre_grouped_df.columns:
                    st.dataframe(dre_grouped_df.drop(columns=['SALDO ANUALIZADO']))
                else:
                    st.dataframe(dre_grouped_df)   


            # Lista para armazenar os resultados dos índices
            indices_result = []

            # Calcular Índice de Liquidez Imediata
            liquidity_ratio_1 = calculate_liquidity_ratio(balance_details["Disponível "][0], balance_details["Aplicações Financeiras"][0], balance_details["PASSIVO CIRCULANTE  "][0])
            liquidity_ratio_2 = calculate_liquidity_ratio(balance_details["Disponível "][1], balance_details["Aplicações Financeiras"][1], balance_details["PASSIVO CIRCULANTE  "][1])
            indices_result.append({
                "Índice": "Índice de Liquidez Imediata",
                "BALANCETE DEZ": liquidity_ratio_1,
                "BALANCETE ATUAL": liquidity_ratio_2,
                "Análise Saldo Dez": "Bom" if liquidity_ratio_1 >= 1 else "Ruim",
                "Análise Saldo Atual": "Bom" if liquidity_ratio_2 >= 1 else "Ruim"
            })

            # Calcular Índice de Liquidez Corrente
            current_ratio_1 = calculate_current_ratio(balance_details["ATIVO CIRCULANTE  "][0], balance_details["PASSIVO CIRCULANTE  "][0])
            current_ratio_2 = calculate_current_ratio(balance_details["ATIVO CIRCULANTE  "][1], balance_details["PASSIVO CIRCULANTE  "][1])
            indices_result.append({
                "Índice": "Índice de Liquidez Corrente",
                "BALANCETE DEZ": current_ratio_1,
                "BALANCETE ATUAL": current_ratio_2,
                "Análise Saldo Dez": "Bom" if current_ratio_1 >= 1 else "Ruim",
                "Análise Saldo Atual": "Bom" if current_ratio_2 >= 1 else "Ruim"
            })

            # Calcular Índice de Liquidez Seca
            quick_ratio_1 = calculate_quick_ratio(balance_details["ATIVO CIRCULANTE  "][0], balance_details["Bens e Títulos a Receber "][0], balance_details["PASSIVO CIRCULANTE  "][0])
            quick_ratio_2 = calculate_quick_ratio(balance_details["ATIVO CIRCULANTE  "][1], balance_details["Bens e Títulos a Receber "][1], balance_details["PASSIVO CIRCULANTE  "][1])
            indices_result.append({
                "Índice": "Índice de Liquidez Seca",
                "BALANCETE DEZ": quick_ratio_1,
                "BALANCETE ATUAL": quick_ratio_2,
                "Análise Saldo Dez": "Bom" if quick_ratio_1 >= 1 else "Ruim",
                "Análise Saldo Atual": "Bom" if quick_ratio_2 >= 1 else "Ruim"
            })

            # Calcular Índice de Liquidez Geral
            general_liquidity_ratio_1 = calculate_general_liquidity_ratio(
                balance_details["ATIVO CIRCULANTE  "][0], balance_details["Realizável a Longo Prazo  "][0], balance_details["PASSIVO CIRCULANTE  "][0], balance_details["PASSIVO NÃO CIRCULANTE "][0])
            general_liquidity_ratio_2 = calculate_general_liquidity_ratio(
                balance_details["ATIVO CIRCULANTE  "][1], balance_details["Realizável a Longo Prazo  "][1], balance_details["PASSIVO CIRCULANTE  "][1], balance_details["PASSIVO NÃO CIRCULANTE "][1])
            indices_result.append({
                "Índice": "Índice de Liquidez Geral",
                "BALANCETE DEZ": general_liquidity_ratio_1,
                "BALANCETE ATUAL": general_liquidity_ratio_2,
                "Análise Saldo Dez": "Bom" if general_liquidity_ratio_1 >= 1 else "Ruim",
                "Análise Saldo Atual": "Bom" if general_liquidity_ratio_2 >= 1 else "Ruim"
            })

            # Calcular Índice de Grau de Endividamento
            general_endividamento_1 = calculate_endividamento(balance_details["ATIVO  "][0], balance_details["PASSIVO  "][0])
            general_endividamento_2 = calculate_endividamento(balance_details["ATIVO  "][1], balance_details["PASSIVO  "][1])
            indices_result.append({
                "Índice": "Grau de Endividamento",
                "BALANCETE DEZ": general_endividamento_1,
                "BALANCETE ATUAL": general_endividamento_2,
                "Análise Saldo Dez": "Bom" if general_endividamento_1 >= 1 else "Ruim",
                "Análise Saldo Atual": "Bom" if general_endividamento_2 >= 1 else "Ruim"
            })

            # Calcular Índice de PL Negativo
            pl_negativo_1 = balance_details["PATRIMÔNIO LÍQUIDO / PATRIMÔNIO SOCIAL "][0]
            pl_negativo_2 = balance_details["PATRIMÔNIO LÍQUIDO / PATRIMÔNIO SOCIAL "][1]
            indices_result.append({
                "Índice": "Índice de PL Negativo",
                "BALANCETE DEZ": pl_negativo_1,
                "BALANCETE ATUAL": pl_negativo_2,
                "Análise Saldo Dez": "Ruim" if pl_negativo_1 < 0 else "Bom",
                "Análise Saldo Atual": "Ruim" if pl_negativo_2 < 0 else "Bom"
            })

            # Criar DataFrame consolidado
            indices_df = pd.DataFrame(indices_result)
            st.write("Índices BL:")
            st.dataframe(indices_df)

            # Atual
            resultado_liquido_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'RESULTADO LÍQUIDO']['SALDO BALANCETE ATUAL'].sum()
            despesas_financeiras_atual = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Despesas Financeiras']['SALDO BALANCETE ATUAL'].sum()
            emprestimos_circulante_atual = filtered_df[filtered_df['DESCRIÇÃO'] == 'Empréstimos e Financiamentos a Pagar']['SALDO BALANCETE ATUAL'].sum()
            emprestimos_nao_circulante_atual = filtered_df[filtered_df['DESCRIÇÃO'] == 'Empréstimos e Financiamentos a Pagar não circulante']['SALDO BALANCETE ATUAL'].sum()
            contraprestacoes = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde']['SALDO BALANCETE ATUAL'].sum()
            eventos_indenizaveis = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Eventos Indenizáveis Líquidos / Sinistros Retidos']['SALDO BALANCETE ATUAL'].sum()

            ebitda_atual = calculate_ebitda(resultado_bruto_atual, contraprestacoes)
            ebitda_index_atual = calculate_ebitda_index(ebitda_atual, contraprestacoes)
            margem_liquida_atual = calculate_net_margin(resultado_liquido_atual, contraprestacoes)
            receita_vs_custo_atual = calculate_revenue_vs_cost(eventos_indenizaveis, contraprestacoes)
            despesas_vs_financ_atual = calculate_expenses_vs_financial(despesas_financeiras_atual, emprestimos_circulante_atual + emprestimos_nao_circulante_atual)

            # Anterior
            resultado_liquido_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'RESULTADO LÍQUIDO']['SALDO BALANCETE ANTERIOR'].sum()
            despesas_financeiras_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Despesas Financeiras']['SALDO BALANCETE ANTERIOR'].sum()
            emprestimos_circulante_anterior = filtered_df[filtered_df['DESCRIÇÃO'] == 'Empréstimos e Financiamentos a Pagar']['SALDO BALANCETE ANTERIOR'].sum()
            emprestimos_nao_circulante_anterior = filtered_df[filtered_df['DESCRIÇÃO'] == 'Empréstimos e Financiamentos a Pagar não circulante']['SALDO BALANCETE ANTERIOR'].sum()
            contraprestacoes_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Contraprestações Efetivas / Prêmios Ganhos de Plano de Assistência à Saúde']['SALDO BALANCETE ANTERIOR'].sum()
            eventos_indenizaveis_anterior = dre_grouped_df[dre_grouped_df['DESCRIÇÃO'] == 'Eventos Indenizáveis Líquidos / Sinistros Retidos']['SALDO BALANCETE ANTERIOR'].sum()

            ebitda_anterior = calculate_ebitda(resultado_bruto_anterior, contraprestacoes_anterior)
            ebitda_index_anterior = calculate_ebitda_index(ebitda_anterior, contraprestacoes_anterior)
            margem_liquida_anterior = calculate_net_margin(resultado_liquido_anterior, contraprestacoes_anterior)
            receita_vs_custo_anterior = calculate_revenue_vs_cost(eventos_indenizaveis_anterior, contraprestacoes_anterior)
            despesas_vs_financ_anterior = calculate_expenses_vs_financial(despesas_financeiras_anterior, emprestimos_circulante_anterior + emprestimos_nao_circulante_anterior)

            indices_dre_result = [
                {"Índice": "EBTIDA", "Atual": ebitda_atual, "Anterior": ebitda_anterior},
                {"Índice": "Índice EBTIDA", "Atual": ebitda_index_atual, "Anterior": ebitda_index_anterior},
                {"Índice": "Margem Líquida", "Atual": margem_liquida_atual, "Anterior": margem_liquida_anterior},
                {"Índice": "Receita Líquida vs Custo", "Atual": receita_vs_custo_atual, "Anterior": receita_vs_custo_anterior},
                {"Índice": "Despesas vs Empréstimos Financeiros", "Atual": despesas_vs_financ_atual, "Anterior": despesas_vs_financ_anterior},
            ]

            indices_dre_df = pd.DataFrame(indices_dre_result)

            # Aplicar a formatação ao DataFrame para exibição
            formatted_indices_dre_df = indices_dre_df.applymap(lambda x: format_number(x) if isinstance(x, (int, float)) else x)
            
            st.write("Índices DRE:")
            st.dataframe(formatted_indices_dre_df)

            if st.button('Exportar Excel'):
                file_path = save_excel(
                    [filtered_df, dre_grouped_df, indices_df, formatted_indices_dre_df],  # Lista dos seus DataFrames
                    ["BL", "DRE", "Índices BL", "Índices DRE"]  # Nomes das abas correspondentes
                )
                st.success('Exportado com sucesso! Clique abaixo para baixar o arquivo.')
                with open(file_path, "rb") as file:
                    btn = st.download_button(
                        label="Baixar Excel",
                        data=file,
                        file_name='Revisão Analitica das OPS.xlsx',
                        mime='application/vnd.ms-excel'
                    )

    main()
