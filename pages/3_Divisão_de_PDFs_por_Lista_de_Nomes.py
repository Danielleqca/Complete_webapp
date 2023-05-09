import os
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
import streamlit as st
import zipfile
from utils import bg_page
st.set_page_config(
    page_title="Divisão de PDFs por Lista de Nomes",
    page_icon='qca_logo_2.png',
    layout="wide",
)

bg_page('bg_dark.png')

hide_menu = """
<style>
#MainMenu {
    visibility:visible;
}

footer {
    visibility:visible;
}

footer:before {
    content:'Desenvolvido pela Eficiência Jurídica - Controladoria Jurídica';
    display:block;
    position:relative;
    color:#6c6a76;
}
</style>
"""
with st.sidebar:
    st.image('qca_logo_2.png')
    st.title('Divisão de PDFs por Lista de Nomes')
    st.info('Esse projeto irá ajudar você a fazer a separação do arquivo em PDF de forma mais eficiente e automática.')            

st.title("Divisão de PDFs por Lista de Nomes")
st.markdown(hide_menu, unsafe_allow_html=True)

st.write("Este aplicativo separa um arquivo PDF em vários arquivos PDFs menores com base em uma lista de nomes.")

# Inserir lista de nomes
uploaded_file = st.file_uploader("Insira a lista de nomes em formato Excel:", type= "xlsx")
if uploaded_file is not None:
    beneficiarios_estag = pd.read_excel(uploaded_file, engine='openpyxl')

    # Criando uma nova lista de estagiários beneficiários para utilizar como parâmetro
    lista_ben_estag = []
    for i in beneficiarios_estag.iloc[:, 0]:
        lista_ben_estag.append(i)

    # Inserir arquivo PDF
    uploaded_pdf_file = st.file_uploader("Insira o arquivo PDF a ser separado:", type=["pdf"])
    if uploaded_pdf_file is not None:
        # Nome do arquivo PDF de entrada
        pdf_file = uploaded_pdf_file.name

        # Criar uma pasta temporária para salvar os arquivos PDFs
        temp_dir = "arquivos"
        if not os.path.exists(temp_dir):
            os.mkdir(temp_dir)

        # Para cada página do arquivo PDF
        for page_num in range(len(PdfReader(uploaded_pdf_file).pages)):
            # Obter o texto da página atual
            page = PdfReader(uploaded_pdf_file).pages[page_num]
            text = page.extract_text()
            for name in lista_ben_estag:
                # Verificar se o nome da pessoa está presente na página atual
                if name in text:
                    # Criar um novo arquivo PDF para a pessoa
                    output_pdf = PdfWriter()
                    output_pdf.add_page(page)

                    # Salvar o novo arquivo PDF com o nome da pessoa
                    filename = f"{name}.pdf"
                    output_filename = os.path.join(temp_dir, filename)
                    with open(output_filename, 'wb') as f:
                        output_pdf.write(f)

        # Criar um arquivo zip contendo todos os arquivos PDFs separados
        zip_filename = "pdfs_separados.zip"
        with zipfile.ZipFile(zip_filename, "w") as zip:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    zip.write(os.path.join(root, file))

        # Remover a pasta temporária
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                os.remove(os.path.join(root, file))
        os.rmdir(temp_dir)

        # Exibir um botão para baixar o arquivo zip contendo os PDFs separados
        with open(zip_filename, "rb") as f:
            bytes = f.read()
            st.download_button(
                label="Baixar todos os PDFs",
                data=bytes,
                file_name=zip_filename,
                mime="application/zip",
            )

        # Exibir uma mensagem de conclusão
        st.success("Todos os PDFs foram separados e salvos em um arquivo zip.")
        
    else:
        st.warning("Por favor, insira o arquivo PDF a ser separado.")

