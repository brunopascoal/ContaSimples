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


def run_selecoes_app():
    # Função para criar uma imagem de texto
    def create_text_image(
        text, font_path="arial.ttf", font_size=14, img_width=500, img_height=500
    ):
        image = PILImage.new("RGB", (img_width, img_height), "white")
        draw = ImageDraw.Draw(image)
        font = ImageFont.truetype(font_path, font_size)
        draw.text((10, 10), text, fill="black", font=font)
        return image

    # Função para adicionar uma aba de imagem ao Workbook
    def add_image_to_excel(image, workbook, sheet_name="Imagem de Informações"):
        img_sheet = workbook.create_sheet(title=sheet_name)
        img_buffer = io.BytesIO()
        image.save(img_buffer, format="PNG")
        img_buffer.seek(0)
        img = OpenpyxlImage(img_buffer)
        img_sheet.add_image(img, "A1")

    # Gerando o workbook e adicionando dados
    def generate_excel(dataframe, text_info, client_name):
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

    st.title("Seleção de Amostras")

    uploaded_file = st.sidebar.file_uploader(
        "Faça o upload de seu arquivo Excel", type=["xlsx"]
    )

    # Utilizando session_state para manter o estado dos dados
    if "df" not in st.session_state:
        st.session_state.df = None
        st.session_state.seed_used = None
        st.session_state.uploaded_filename = None
        st.session_state.selected_sheet = None

    if uploaded_file is not None:
        clientes = [
            "APAS",
            "BS",
            "UA",
            "UAV",
            "UB",
            "UG",
            "UL",
            "UP",
            "UPN",
            "UPP",
            "URG",
            "URP",
            "URSG",
            "USBA",
            "USC",
            "USP",
            "UVR",
        ]
        nome_cliente = st.sidebar.selectbox("Escolha o cliente:", clientes)

        selecoes = [
            "T600.1.2.1 - Individual 123",
            "T600.1.2.1 - Coletivo 123",
            "T600.1.2.2 - 124",
            "T600.1.2.3 - PJ Eventos a liquidar",
            "T600.1.2.3 - PF Eventos a liquidar",
            "T600.1.2.4 - PJ Eventos subsequentes",
            "T600.1.2.4 - PF Eventos subsequentes",
            "T600.1.2.5 - 214",
            "T600.1.2.6 - 2132",
            "T600.1.2.7 - Seleção remissão",
            "T600.1.15 - Capital Referente ao Risco de Crédito",
        ]
        nome_selecoes = st.sidebar.selectbox("Escolha a seleção:", selecoes)

        if st.session_state.uploaded_filename != uploaded_file.name:
            st.session_state.uploaded_filename = uploaded_file.name
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = (
                xl.sheet_names
            )  # Obtém todas as abas (sheets) do arquivo Excel
            st.session_state.selected_sheet = st.sidebar.selectbox(
                "Escolha a aba do Excel:", sheet_names
            )
            st.session_state.df = xl.parse(
                st.session_state.selected_sheet
            )  # Carrega a aba selecionada
        else:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            # Sempre mostra a selectbox, mas mantém a seleção anterior se a aba ainda não foi mudada
            new_selected_sheet = st.sidebar.selectbox(
                "Escolha a aba do Excel:",
                sheet_names,
                index=sheet_names.index(st.session_state.selected_sheet),
            )
            if new_selected_sheet != st.session_state.selected_sheet:
                st.session_state.selected_sheet = new_selected_sheet
                st.session_state.df = xl.parse(
                    new_selected_sheet
                )  # Recarrega se a aba selecionada mudar

        df = st.session_state.df

        st.sidebar.markdown("### Filtros de Dados")
        desconsiderar_valores = st.sidebar.checkbox(
            "Desconsiderar valores com base em outro arquivo?"
        )
        if desconsiderar_valores:
            exclusion_file = st.sidebar.file_uploader(
                "Faça o upload do arquivo com os registros a serem excluídos",
                type=["xlsx"],
            )

            if exclusion_file is not None:
                exclusion_xl = pd.ExcelFile(exclusion_file)
                exclusion_sheet_names = exclusion_xl.sheet_names
                selected_exclusion_sheet = st.sidebar.selectbox(
                    "Escolha a aba do Excel para exclusão:", exclusion_sheet_names
                )
                st.session_state.exclusion_df = exclusion_xl.parse(
                    selected_exclusion_sheet
                )
                st.session_state.exclusion_ready = True

                exclusion_column = st.sidebar.selectbox(
                    "Coluna no arquivo de exclusão:",
                    st.session_state.exclusion_df.columns,
                )
                base_column = st.sidebar.selectbox(
                    "Coluna correspondente no arquivo principal:", df.columns
                )

                if st.sidebar.button("Excluir registros"):
                    exclusion_values = set(
                        st.session_state.exclusion_df[exclusion_column]
                    )
                    df = df[~df[base_column].isin(exclusion_values)]
                    st.session_state.df = df
                    st.success("Registros excluídos com sucesso!")

                # Adicionando um botão de instrução para alerta

        coluna_para_filtrar = st.sidebar.selectbox(
            "Escolha a coluna para filtrar:", ["Nenhuma"] + list(df.columns)
        )
        if coluna_para_filtrar != "Nenhuma":
            tipo_de_filtro = st.sidebar.radio(
                "Tipo de Filtro:", ["Maior Que", "Menor Que"]
            )
            valor_filtro = st.sidebar.number_input("Valor de Referência:", value=0)

            if tipo_de_filtro == "Maior Que":
                df = df[df[coluna_para_filtrar] > valor_filtro]
            elif tipo_de_filtro == "Menor Que":
                df = df[df[coluna_para_filtrar] < valor_filtro]
        else:
            tipo_de_filtro = "NA"
            valor_filtro = "NA"
        st.sidebar.markdown("### Configurações de Seed")
        seed = st.sidebar.text_input(
            "Digite uma seed para reproduzir uma seleção anterior ou deixe em branco para uma nova:"
        )

        usar_seed = False
        if seed:
            try:
                seed = int(seed)
                usar_seed = True
                np.random.seed(seed)
            except ValueError:
                st.sidebar.error("A seed deve ser um número inteiro.")

        st.sidebar.markdown("### Seleção de Maiores Valores")
        selecionar_maiores = st.sidebar.checkbox(
            "Selecionar maiores valores antes da amostragem?"
        )

        if selecionar_maiores:
            coluna_maiores_valores = st.sidebar.selectbox(
                "Escolha a coluna para os maiores valores:", df.columns
            )
            qtd_maiores = st.sidebar.number_input(
                "Quantidade de maiores valores", min_value=1, max_value=len(df), value=5
            )
            maiores_valores = df.nlargest(qtd_maiores, coluna_maiores_valores)
            df = df.drop(maiores_valores.index)
            st.markdown("#### Maiores valores selecionados:")
            st.dataframe(maiores_valores)
            output = io.BytesIO()
            # Salvar o DataFrame no objeto de buffer Excel
            maiores_valores.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)  # Rebobinar o buffer
            nome_maiores = f"Maiores valores - {nome_cliente}.xlsx"
            # Criar o botão de download para o arquivo Excel
            download_button(
                label="Baixar seleção maiores valores",
                data=output,
                file_name=nome_maiores,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.sidebar.markdown("### Seleção Aleatória de Amostras")
        n_amostras = st.sidebar.number_input(
            "Número de amostras", min_value=1, max_value=30, value=10
        )
        coluna_para_evitar_repeticao = st.sidebar.selectbox(
            'Escolha a coluna para evitar repetição (ou "Nenhuma"):',
            ["Nenhuma"] + list(df.columns),
        )

        # if st.button("Gerar Amostras"):
        if not usar_seed:
            seed = np.random.randint(0, 100000)
            np.random.seed(seed)
        amostras_selecionadas = df.sample(n=n_amostras, replace=False)
        if coluna_para_evitar_repeticao != "Nenhuma":
            # Agrupa pelo valor da coluna selecionada e escolhe um registro aleatório de cada grupo
            # Ao usar groupby, garantimos que cada grupo é tratado separadamente
            grouped = df.groupby(coluna_para_evitar_repeticao, group_keys=False)
            # Aplicando a função de sample em cada grupo para pegar um registro aleatório
            amostras_selecionadas = grouped.apply(lambda x: x.sample(1))
            # Amostra aleatória dos registros selecionados, caso haja mais grupos do que amostras necessárias
            n_amostras = min(n_amostras, len(amostras_selecionadas))
            amostras_selecionadas = amostras_selecionadas.sample(
                n=n_amostras, random_state=seed
            )
        else:
            # Amostragem sem considerar a coluna para evitar repetição
            amostras_selecionadas = df.sample(n=n_amostras, random_state=seed)

        st.markdown(f" #### Amostras aleatórias geradas")

        selection_info = f"Data e Hora da Seleção: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        selection_info += f"Semente Utilizada: {seed}\n"
        selection_info += (
            f"Caminho do Arquivo da seleção: {st.session_state.uploaded_filename}\n"
        )
        selection_info += f"Aba Utilizada: {st.session_state.selected_sheet}\n"
        selection_info += f"Coluna Filtrada: {coluna_para_filtrar if coluna_para_filtrar != 'Nenhuma' else 'Nenhuma'}\n"
        selection_info += (
            f"Tipo de Filtro: {tipo_de_filtro if tipo_de_filtro else 'Nenhum'}\n"
        )
        selection_info += f"Valor de Referência para Filtro: {valor_filtro if valor_filtro is not None else 'Nenhum'}\n"
        selection_info += f"Coluna para Maiores Valores: {coluna_maiores_valores if selecionar_maiores else 'Não aplicável'}\n"
        selection_info += f"Quantidade de Maiores Valores: {qtd_maiores if selecionar_maiores else 'Não aplicável'}\n"
        selection_info += f"Número de Amostras Selecionadas: {n_amostras}\n"
        selection_info += f"Coluna para Evitar Repetição: {coluna_para_evitar_repeticao if coluna_para_evitar_repeticao != 'Nenhuma' else 'Não aplicável'}\n"
        st.dataframe(amostras_selecionadas)

        output = generate_excel(amostras_selecionadas, selection_info, nome_cliente)

        st.download_button(
            label="Baixar amostras selecionadas",
            data=output,
            file_name=f"Amostra Aleatória - {nome_selecoes} - {nome_cliente} - {seed}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    else:
        st.write("Por favor, faça o upload de um arquivo Excel para começar.")
