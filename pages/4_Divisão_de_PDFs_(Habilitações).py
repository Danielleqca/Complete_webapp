import os
import pandas as pd
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
import streamlit as st
import zipfile 
import re
from utils import bg_page
st.set_page_config(
    page_title="Divisão de PDFs (Habilitações)",
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
    st.title('Divisão de PDFs (Habilitações)')
    st.info('Esse projeto irá ajudar você a fazer a separação do arquivo em PDF de forma mais eficiente e automática.')            

# Nome do arquivo PDF de entrada
st.title('Divisão de PDFs (Habilitações)')
st.markdown(hide_menu, unsafe_allow_html=True)

# Inserir arquivo PDF
uploaded_pdf_file = st.file_uploader("Insira o arquivo PDF a ser separado:", type=["pdf"])
if uploaded_pdf_file is not None:
    # Nome do arquivo PDF de entrada
    pdf_file = uploaded_pdf_file.name

    # criando uma pasta temporaria pra salvar os pdfs separados
    temp_dir = 'arquivos'
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)

    for page_num in range(len(PdfReader(uploaded_pdf_file).pages)):
        page = PdfReader(uploaded_pdf_file).pages[page_num]
        text = page.extract_text()
        # Busca o número do processo na página atual
        process_number = re.search(r"\d{7,8}-\d{2}\.\d{4}\.\d{1,2}\.\d{2,3}\.\d{3,4}", text)
        if process_number:
            # Criar um novo arquivo PDF para a pessoa
            output_pdf = PdfWriter()
            output_pdf.add_page(page)

            # Salvar o novo arquivo PDF
            output_filename = os.path.join(temp_dir, f'{process_number.group(0)}.pdf')
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
