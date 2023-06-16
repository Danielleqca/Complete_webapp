import os
import win32com.client as win32
import pandas as pd
from email.mime.image import MIMEImage
import base64
import streamlit as st
import warnings
import openpyxl
import pythoncom

warnings.filterwarnings('ignore')

def enviar_email():
    # controle de arquivos
    BASE_DIR = os.path.dirname(os.path.abspath('__file__'))
    IMG_DIR = os.path.join(BASE_DIR, 'img')
    DATA_DIR = os.path.join(BASE_DIR, 'data')
    ARQUIVO_DIR = os.path.join(DATA_DIR, 'envio_parabens.xlsx')
    caminho_imagem = os.path.join(BASE_DIR, IMG_DIR, 'Publicações.png') # mudar o nome da img

    # abrir planilha com base de dados com tratamento
    df = pd.read_excel(ARQUIVO_DIR)
    df.EMAIL = df.EMAIL.fillna('0')
    df.IMAGEM = df.IMAGEM.fillna('0')
    df['CENTROS DE CUSTO'] = df['CENTROS DE CUSTO'].fillna('0')
    df.ID = df.ID.astype('object')

    cont = 1
    for i, contato in enumerate(df['EMAIL']):
        indice = df.loc[i, "CENTROS DE CUSTO"]

        # obter caminho completo da imagem
        nome_imagem = df.loc[i, "IMAGEM"]
        caminho_imagem = caminho_imagem

        pythoncom.CoInitialize()
        # criar a integração com o outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um email
        email = outlook.CreateItem(0)

        # configurar as informações do e-mail
        email.To = contato
        email.cc = "rhuansilva@queirozcavalcanti.adv.br"
        email.Subject = f"Auditoria CJ - Publicação - {indice}"

        # Codificar a imagem em formato base64
        with open(caminho_imagem, 'rb') as img_file:
            img_data = img_file.read()
        img_base64 = base64.b64encode(img_data).decode('utf-8')

        # Corpo do e-mail
        email.HTMLBody = f"""
        <img src="data:image/png;base64,{img_base64}" alt="Imagem">
        <p>À disposição!<p>
        <p>Atenciosamente,</p>
        <p>{nome_selecionado} - Núcleo de Gestão de Dados - Controladoria Jurídica.</p>
        """

        # Enviar o e-mail
        email.Send()

        st.write(f"Email Enviado {cont}")
        cont += 1

def selecionar_arquivo():
    arquivo = st.file_uploader("Selecionar Arquivo", type=['xlsx'])
    if arquivo is not None:
        df = pd.read_excel(arquivo, engine='openpyxl')
        st.write(f"Arquivo selecionado: {arquivo.name}")
        st.dataframe(df)


st.set_page_config(page_title='Envio Automático de E-mails',
                    layout='wide')

with st.sidebar:
    st.image('https://www.onepointltd.com/wp-content/uploads/2019/12/shutterstock_1166533285-Converted-02.png')
    st.title('Envio Automático de E-mails')
    # choices = st.selectbox('Escolha a auditoria:', ('Obrigações', 'PA', 'Prazo', 'Garantia', 'Publicação', 'Prazo de Audiência', 'Com Subsídio',
    # 'Sem Estratégia', 'Mensagens', 'Orientações de Audiência', 'Valor a Provisionar, Processos Ativos', 'Revisão Cadastro', 'Acordos'))
    st.info('Esse projeto irá ajudar você a fazer o envio de e-mails de forma mais eficiente e automática.')            



st.markdown("### Envio Automático de E-mails - Publicação")
st.markdown('#### Importe uma planilha')
#
#  Lista de nomes de pessoas
nomes = ['Liliane Germano', 'Lívia Freitas', 'Rebecca Fidélis', 'Mila Ramos', 'João Paulo']

# Criar a lista suspensa
nome_selecionado = st.selectbox('Selecione um nome:', nomes)

# Exibir o nome selecionado
st.write('Nome selecionado:', nome_selecionado)

uploaded_file = st.file_uploader('Base de dados - Centro de Custo:', type='xlsx')
st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
st.markdown('---')
if uploaded_file is not None:
    agendados = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(agendados)
    tratar_button = st.button('Enviar E-mail')
    if tratar_button:
        agendados_tratado = enviar_email()
        st.success('E-mail enviado!')
