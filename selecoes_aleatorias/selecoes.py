import streamlit as st
import pandas as pd
import numpy as np
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage, ImageDraw, ImageFont
import io
from streamlit import download_button
import os
import json

def create_text_image(
    text, font_name="arial.ttf", font_size=14, img_width=500, img_height=500
):
    # Função para criar uma imagem de texto
    image = PILImage.new("RGB", (img_width, img_height), "white")
    draw = ImageDraw.Draw(image)

    # Ajustar o caminho para a raiz do projeto e depois para o diretório assets/fonts
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(current_dir, os.pardir))  # Sobe um nível
    font_path = os.path.join(project_root, 'assets', 'fonts', font_name)
    
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        st.warning("Fonte Arial não encontrada. Usando fonte padrão.")
        font = ImageFont.load_default()
    
    draw.text((10, 10), text, fill="black", font=font)
    return image

def add_image_to_excel(image, workbook, sheet_name="Imagem de Informações"):
    # Função para adicionar uma aba de imagem ao Workbook
    img_sheet = workbook.create_sheet(title=sheet_name)
    img_buffer = io.BytesIO()
    image.save(img_buffer, format="PNG")
    img_buffer.seek(0)
    img = OpenpyxlImage(img_buffer)
    img_sheet.add_image(img, "A1")

def generate_excel(dataframe, text_info, client_name):
    # Gerando o workbook e adicionando dados
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Dados"

    for r in dataframe_to_rows(dataframe, index=False, header=True):
        ws.append(r)

    # Criação da imagem com informações
    image = create_text_image(text_info)

    # Adicionando a imagem ao Excel
    add_image_to_excel(image, workbook)

    # Salva o workbook em um buffer
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output

def get_client_list(selection_type):

    if selection_type == "Seleções":
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, 'clients.txt')
        # Abra o arquivo e processe as linhas, garantindo que está sendo lido corretamente
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                return sorted([line.strip() for line in file if line.strip()])
        except FileNotFoundError:
            print(f"Arquivo 'clients.txt' não foi encontrado no caminho {file_path}.")
            return []
    else:
        return []
    
def get_selections_list(selection_type):
    if selection_type == "Seleções":
        return {
                "Preliminar": [
                    "100.50. - Seleção caixa e equivalente de caixa",
                    "105.50. - Seleção contas a receber",
                    "110.50. - Seleção estoque",
                    "115.50. - Seleção impostos e contribuições a recuperar",
                    "120.50. - Seleção Partes relacionadas",
                    "125.50. - Seleção Outros ativos",
                    "130.50. - Seleção Despesas antecipadas",
                    "135.50. - Seleção Depósitos judiciais",
                    "145.50. - Seleção Investimentos",
                    "155.50. - Seleção Imobilizado",
                    "160.50. - Seleção Intangível",
                    "200.50. - Seleção Empréstimos e financiamentos",
                    "205.50. - Seleção Contas a pagar",
                    "210.50. - Seleção Obrigações sociais e trabalhistas",
                    "215.50. - Seleção Obrigações fiscais",
                    "220.50. - Seleção Dividendos ou juros de capital próprio",
                    "225.50. - Seleção Partes relacionadas",
                    "230.50. - Seleção Provisão para contingências",
                    "235.50. - Seleção Outros passivos",
                    "240.50. - Seleção Patrimônio líquido",
                    "300.50. - Teste de voucher - Receita",
                    "305.50. - Teste de voucher - Custo",
                    "310.50. - Teste de voucher - Despesas comerciais",
                    "315.50. - Teste de voucher - Despesas gerais e administrativas",
                    "320.50. - Teste de voucher - Outras receitas",
                    "320.50.. - Teste de voucher - Outras despesas",
                    "325.50. - Teste de voucher - Receitas financeiras",
                    "325.50.. - Teste de voucher - Despesas financeiras"
                ],
                "Final": [
                    "100. - Seleção caixa e equivalente de caixa",
                    "105. - Seleção contas a receber",
                    "110. - Seleção estoque",
                    "115. - Seleção impostos e contribuições a recuperar",
                    "120. - Seleção Partes relacionadas",
                    "125. - Seleção Outros ativos",
                    "130. - Seleção Despesas antecipadas",
                    "135. - Seleção Depósitos judiciais",
                    "145. - Seleção Investimentos",
                    "155. - Seleção Imobilizado",
                    "160. - Seleção Intangível",
                    "200. - Seleção Empréstimos e financiamentos",
                    "205. - Seleção Contas a pagar",
                    "210. - Seleção Obrigações sociais e trabalhistas",
                    "215. - Seleção Obrigações fiscais",
                    "220. - Seleção Dividendos ou juros de capital próprio",
                    "225. - Seleção Partes relacionadas",
                    "230. - Seleção Provisão para contingências",
                    "235. - Seleção Outros passivos",
                    "240. - Seleção Patrimônio líquido",
                    "300. - Teste de voucher - Receita",
                    "305. - Teste de voucher - Custo",
                    "310. - Teste de voucher - Despesas comerciais",
                    "315. - Teste de voucher - Despesas gerais e administrativas",
                    "320. - Teste de voucher - Outras receitas",
                    "320.. - Teste de voucher - Outras despesas",
                    "325. - Teste de voucher - Receitas financeiras",
                    "325.. - Teste de voucher - Despesas financeiras"
                ],
                "Controle Interno": [
                    "T100 - Seleção folha de pagamento",
                    "T200 - Seleção faturamento",
                    "T300 - Seleção compras",
                    "T400 - Seleção custo"
                ]
        }
    else:
        return []
    
def process_file(uploaded_file, selection_type, tab_key):
    # Função para processar o arquivo enviado e executar a lógica principal

    if uploaded_file is not None:

        clientes = get_client_list(selection_type)
        selecoes = get_selections_list(selection_type)

        nome_cliente = st.selectbox("Escolha o cliente:", clientes, key=f"cliente_{tab_key}")

        if isinstance(selecoes, dict):
            etapa_trabalho = st.selectbox("Selecione a etapa:", list(selecoes.keys()), key=f"etapa_{tab_key}")
            nome_selecoes = st.selectbox("Selecione a categoria:", selecoes[etapa_trabalho])

        else:
            etapa_trabalho = st.selectbox("Selecione a etapa:", ["PPA"], key=f"etapa_{tab_key}")
            nome_selecoes = st.selectbox("Selecione a categoria:", selecoes, key=f"nome_selecoes_{tab_key}")


        df = load_excel_file(uploaded_file, tab_key)

        df = apply_filters(df, tab_key)

        df_maiores_valores = select_largest_values(df, tab_key)
        if df_maiores_valores is not None:
            df = df.drop(df_maiores_valores.index)

        seed = get_seed(tab_key)

        amostras_selecionadas = select_random_samples(df, seed, tab_key)

        display_and_download_results(
            amostras_selecionadas,
            df_maiores_valores,
            nome_cliente,
            nome_selecoes,
            etapa_trabalho,
            seed,
            uploaded_file.name,
            tab_key,
        )
    else:
        st.write("Por favor, faça o upload de um arquivo Excel para começar.")

def load_excel_file(uploaded_file, tab_key):
    # Função para carregar o arquivo Excel e selecionar a aba
    if "uploaded_filename" not in st.session_state:
        st.session_state.uploaded_filename = None
    if "selected_sheet" not in st.session_state:
        st.session_state.selected_sheet = None

    if st.session_state.uploaded_filename != uploaded_file.name:
        st.session_state.uploaded_filename = uploaded_file.name
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        st.session_state.selected_sheet = st.selectbox(
            "Escolha a aba do Excel:", sheet_names, key=f"sheet_{tab_key}"
        )
        st.session_state.df = xl.parse(st.session_state.selected_sheet)
    else:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        new_selected_sheet = st.selectbox(
            "Escolha a aba do Excel:",
            sheet_names,
            index=sheet_names.index(st.session_state.selected_sheet),
            key=f"sheet_{tab_key}_reload"
        )
        if new_selected_sheet != st.session_state.selected_sheet:
            st.session_state.selected_sheet = new_selected_sheet
            st.session_state.df = xl.parse(new_selected_sheet)

    return st.session_state.df

def apply_filters(df, tab_key):

    if tab_key == 'tab2':
        st.markdown("### Filtros de Dados")
        desconsiderar_valores = st.checkbox(
            "Desconsiderar valores com base em outro arquivo?",
            key=f"desconsiderar_valores_{tab_key}"
        )
        if desconsiderar_valores:
            df = exclude_values(df, tab_key)

    coluna_para_filtrar = st.selectbox(
        "Escolha a coluna para filtrar:", ["Nenhuma"] + list(df.columns), key=f"coluna_filtrar_{tab_key}"
    )
    if coluna_para_filtrar != "Nenhuma":
        tipo_de_filtro = st.radio(
            "Tipo de Filtro:", ["Maior Que", "Menor Que"]
        )
        valor_filtro = st.number_input("Valor de Referência:", value=0, key=f"valor_filtro_{tab_key}")

        if tipo_de_filtro == "Maior Que":
            df = df[df[coluna_para_filtrar] > valor_filtro]
        elif tipo_de_filtro == "Menor Que":
            df = df[df[coluna_para_filtrar] < valor_filtro]
    return df

def exclude_values(df, tab_key):
    # Função para excluir valores com base em outro arquivo
    exclusion_file = st.file_uploader(
        "Faça o upload do arquivo com os registros a serem excluídos",
        type=["xlsx"],
    )

    if exclusion_file is not None:
        exclusion_xl = pd.ExcelFile(exclusion_file)
        exclusion_sheet_names = exclusion_xl.sheet_names
        selected_exclusion_sheet = st.selectbox(
            "Escolha a aba do Excel para exclusão:", exclusion_sheet_names
        )
        exclusion_df = exclusion_xl.parse(selected_exclusion_sheet)

        exclusion_column = st.selectbox(
            "Coluna no arquivo de exclusão:",
            exclusion_df.columns,
            key=f"exclusion_column_{tab_key}"
        )
        base_column = st.selectbox(
            "Coluna correspondente no arquivo principal:", df.columns, key=f"base_column_{tab_key}"
        )

        if st.button("Excluir registros"):
            exclusion_values = set(exclusion_df[exclusion_column])
            df = df[~df[base_column].isin(exclusion_values)]
            st.success("Registros excluídos com sucesso!")
    return df

def select_largest_values(df, tab_key):
    # Função para selecionar os maiores valores
    st.markdown("### Seleção de Maiores Valores")
    selecionar_maiores = st.checkbox(
        "Selecionar maiores valores antes da amostragem?",
        key=f"selecionar_maiores_{tab_key}"

    )

    if selecionar_maiores:
        coluna_maiores_valores = st.selectbox(
            "Escolha a coluna para os maiores valores:", df.columns, key=f"coluna_maiores_{tab_key}"
        )
        qtd_maiores = st.number_input(
            "Quantidade de maiores valores", min_value=1, max_value=len(df), value=5, key=f"qtd_maiores_{tab_key}"
        )
        maiores_valores = df.nlargest(qtd_maiores, coluna_maiores_valores)
        st.markdown("#### Maiores valores selecionados:")
        st.dataframe(maiores_valores)
        output = io.BytesIO()
        maiores_valores.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)
        nome_maiores = f"Maiores valores selecionados.xlsx"
        download_button(
            label="Baixar seleção maiores valores",
            data=output,
            file_name=nome_maiores,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return maiores_valores
    else:
        return None

def get_seed(tab_key):
    # Função para obter ou gerar a seed

    # st.markdown("### Configurações de Seed")
    # seed_input = st.text_input(
    #     "Digite uma seed para reproduzir uma seleção anterior ou deixe em branco para uma nova:",
    #     key=f"seed_input_{tab_key}"
    # )

    usar_seed = False
    # if seed_input:
    #     try:
    #         seed = int(seed_input)
    #         usar_seed = True
    #         np.random.seed(seed)
    #     except ValueError:
    #         st.error("A seed deve ser um número inteiro.")
    if not usar_seed:
        seed = np.random.randint(0, 100000)
        np.random.seed(seed)
    return seed

def select_random_samples(df, seed, tab_key):
    # Função para selecionar amostras aleatórias
    st.markdown("### Seleção Aleatória de Amostras")
    n_amostras = st.number_input(
        "Número de amostras", min_value=1, value=10, key=f"n_amostras_{tab_key}"
    )
    coluna_para_evitar_repeticao = st.selectbox(
        'Escolha a coluna para evitar repetição (ou "Nenhuma"):',
        ["Nenhuma"] + list(df.columns), key=f"evitar_repeticao_{tab_key}"
    )

    if coluna_para_evitar_repeticao != "Nenhuma":
        grouped = df.groupby(coluna_para_evitar_repeticao, group_keys=False)
        amostras_selecionadas = grouped.apply(lambda x: x.sample(1))
        n_amostras = min(n_amostras, len(amostras_selecionadas))
        amostras_selecionadas = amostras_selecionadas.sample(
            n=n_amostras, random_state=seed
        )
    else:
        amostras_selecionadas = df.sample(n=n_amostras, random_state=seed)
    return amostras_selecionadas

def log_action(username, action_description):
    # Define o caminho para o arquivo de log
    log_file_path = "logs.json"
    
    # Cria um dicionário com os dados do log
    log_entry = {
        "username": username,
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "action": action_description
    }
    
    # Verifica se o arquivo de log existe e se está vazio
    if os.path.exists(log_file_path):
        if os.path.getsize(log_file_path) > 0:
            # Se o arquivo não estiver vazio, carrega os logs existentes
            with open(log_file_path, "r", encoding="utf-8") as file:
                logs = json.load(file)
        else:
            # Se o arquivo estiver vazio, inicializa uma lista de logs vazia
            logs = []
    else:
        # Se o arquivo não existir, inicializa uma lista de logs vazia
        logs = []

    # Adiciona a nova entrada de log
    logs.append(log_entry)

    # Salva todos os logs de volta no arquivo
    with open(log_file_path, "w", encoding="utf-8") as file:
        json.dump(logs, file, indent=4, ensure_ascii=False)

def display_and_download_results(
    amostras_selecionadas,
    df_maiores_valores,
    nome_cliente,
    nome_selecoes,
    etapa_trabalho,
    seed,
    uploaded_filename,
    tab_key, # Adicione o tab_key aqui

):
    # Função para exibir e permitir o download dos resultados
    st.markdown(f"#### Amostras aleatórias geradas")

    selection_info = f"Data e Hora da Seleção: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    selection_info += f"Etapa do trabalho: {etapa_trabalho}\n"
    selection_info += f"Semente Utilizada: {seed}\n"
    selection_info += f"Caminho do Arquivo da seleção: {uploaded_filename}\n"
    selection_info += f"Aba Utilizada: {st.session_state.selected_sheet}\n"
    # Outros detalhes podem ser adicionados aqui

    st.dataframe(amostras_selecionadas)


    output = generate_excel(amostras_selecionadas, selection_info, nome_cliente)



    def handle_download_click():
        # Exibe um toast de notificação
        st.toast("✅ Download realizado!")

        # Gera o log
        log_action(
            st.session_state.username, 
            f"Gerou amostras aleatórias para cliente {nome_cliente} e seleção {nome_selecoes} usando seed {seed}. Detalhes: {selection_info}"
        )

    st.markdown("## Confira os dados antes de baixar")
    st.divider()
    st.markdown(f"""
            #### Nome do cliente: {nome_cliente}  
            #### Etapa: {etapa_trabalho}  
            #### Seleção: {nome_selecoes}  
            #### Nome do arquivo: {uploaded_filename}  
            #### Sheet: {st.session_state.selected_sheet}  
            #### Número de amostras: {len(amostras_selecionadas)}  
            """)

            # Botão de download com múltiplas ações
    st.download_button(
        label="Baixar amostras",
        data=output,
        file_name=f"Amostra Aleatória - {nome_selecoes} - {nome_cliente} - {etapa_trabalho} - {seed}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="⬇️ Clique aqui para baixar as amostras aleatórias solicitadas",
        on_click=handle_download_click  # Chama a função que executa as ações
    )

def run_selecoes_app():


    st.markdown("""## Seleção de Amostras   
                """)

    # Armazena a aba selecionada no session_state
    if "selected_tab" not in st.session_state:
        st.session_state.selected_tab = "Seleções"  # Define uma aba padrão

    # Exibe as abas e atualiza o estado quando uma aba for selecionada
    tab1, tab2, tab3 = st.tabs(["📙 Seleções", "📗 Seleções PPA", "📄 Documentação"])

    with tab1:
        st.header("Seleção Aleatória", divider=True)

        st.session_state.selected_tab = "Seleções"
        uploaded_file = st.file_uploader("Faça o upload de seu arquivo Excel", type=["xlsx"], key="upload_df_tab1")
        if uploaded_file:

            process_file(uploaded_file, "Seleções", "tab1")


        
    with tab3:
        st.header("Documentação Seleções Aleatórias", divider=True)





