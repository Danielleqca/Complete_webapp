import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import re
import pandas as pd
import os
import string
import zipfile
import tempfile
import openpyxl
import unicodedata

# Título do aplicativo
st.title('Aplicativo de Geração de PDFs')

# Widget para carregar o arquivo Excel
excel_file = st.file_uploader('Carregue o arquivo Excel', type=['xlsx'])

# Widget para carregar o arquivo PDF original
pdf_file = st.file_uploader('Carregue o arquivo PDF original', type=['pdf'])

if excel_file is not None and pdf_file is not None:
    # Lógica para processar o Excel e PDF original
    df = pd.read_excel(excel_file, engine='openpyxl')
    pdf_original = PdfReader(pdf_file)

    def clean_filename(filename):
        # Adicionando caracteres especiais à lista de caracteres válidos
        valid_chars = "-_().& %s%s" % (string.ascii_letters, string.digits)
        valid_chars += 'çÇãÃõÕáéíóúýÁÉÍÓÚÝâêîôûÂÊÎÔÛàèìòùÀÈÌÒÙäëïöüÿÄËÏÖÜ'
        # Substituindo caracteres especiais por equivalentes sem acento
        filename = ''.join([c if c in valid_chars else ' ' for c in filename])
        # Removendo espaços duplos e espaços no início e final do nome do arquivo
        cleaned_filename = ' '.join(filename.split())
        return cleaned_filename
    
    # Crie uma pasta temporária para salvar os PDFs gerados
    with tempfile.TemporaryDirectory() as temp_dir:
        for pagina_num in range(len(pdf_original.pages)):
            if pagina_num >= len(df):
                st.text(f"Aviso: Não há dados correspondentes para a página {pagina_num + 1}. Ignorando a página.")
                continue
        
            nf = df.at[pagina_num, 'Título Pirâmide']   # NF
            nomeCorrespondente = df.at[pagina_num, 'Correspondente'] # NOME DO CORRESPONDENTE
            nProcesso = df.at[pagina_num, 'Número Processo']    # NUMERO DO PROCESSO
            idSolicitacao = df.at[pagina_num, ' Id Solicitacao']    # ID DA SOLICITACAO

            # Crie um novo objeto PDF para cada página
            pdf_novo = PdfWriter()

            # Adicione a página atual ao novo objeto PDF
            pdf_novo.add_page(pdf_original.pages[pagina_num])

            # Crie o nome do arquivo de saída com base nas informações do colaborador e número do processo
            nome_arquivo_saida = f'NF {nf} - {nomeCorrespondente} - {nProcesso} - ID {idSolicitacao}.pdf'
            nome_arquivo_saida = clean_filename(nome_arquivo_saida)
            caminho_arquivo_saida = os.path.join(temp_dir, nome_arquivo_saida)

            # Salve o PDF gerado na pasta temporária
            with open(caminho_arquivo_saida, 'wb') as novo_pdf:
                pdf_novo.write(novo_pdf)

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