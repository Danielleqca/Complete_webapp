import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import streamlit as st
import os
import zipfile
import shutil
from utils import bg_page

st.set_page_config(
    page_title="Divisão de abas em novas planilhas",
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
    st.title('Divisão de abas em novas planilhas')
    st.info('Esse projeto irá ajudar você a fazer a separação da planilha de forma mais eficiente e automática.')
    

# Título da página
st.title("Divisão de abas em novas planilhas")

# Upload do arquivo
st.subheader("1. Faça o upload do arquivo a ser separado:")
st.markdown(hide_menu, unsafe_allow_html=True)

file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

# Seleção das colunas
st.subheader("2. Selecione as colunas que deseja manter:")
if file is not None:
    df = pd.read_excel(file)
    colunas = st.multiselect("Selecione as colunas", options=list(df.columns) + ["Todas"])
    # Verifica se a opção "Todas" foi selecionada e inclui todas as colunas
    if "Todas" in colunas:
        colunas = list(df.columns)
    # Separação das planilhas
    if len(colunas) > 0:
        st.subheader("3. Separar planilhas")
        if os.path.exists("planilhas_separadas.zip"):
            os.remove("planilhas_separadas.zip")
        if os.path.exists("output"):
            shutil.rmtree("output")
        equipes = {}
        for equipe in df.iloc[:,0].unique():
            equipes[equipe] = df[df.iloc[:,0] == equipe].reset_index(drop=True)[colunas]

        # Salva os arquivos separados em uma pasta
        os.makedirs("output", exist_ok=True)
        for equipe, df_equipe in equipes.items():
            df_equipe.to_excel(f"output/{equipe}.xlsx", index=False)

        # Cria uma pasta zipada com os arquivos separados
        with zipfile.ZipFile("planilhas_separadas.zip", mode="w") as z:
            for file in os.listdir("output"):
                z.write(f"output/{file}", arcname=file)
        st.success("Planilhas separadas com sucesso!")
        st.download_button(label="Separar planilhas", data=open("planilhas_separadas.zip", "rb").read(), file_name="planilhas_separadas.zip", mime="application/zip")
        # st.experimental_memo.clear()

