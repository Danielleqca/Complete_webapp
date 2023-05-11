import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, time, datetime, timedelta
import os

import base64
from io import StringIO, BytesIO
from utils import bg_page

def generate_excel_download_link(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_TRATADA_GERAL_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link_bahia(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_TRATADA_BAHIA_{hoje}.xlsx">Download Bahia</a>'
    return st.markdown(href, unsafe_allow_html=True)

# -- BASE PRAZOS DIARIOS AGENDADOS
def tratamento_base(base, base_centro_custo):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados.
    '''
    hoje = date.today()
    df = pd.read_excel(base)
    # dropando colunas que nao serao usadas
    centro_custo = pd.read_excel(base_centro_custo)
    df_ajustado = pd.merge(df, centro_custo, on='Célula', how='left')
    df_ajustado.drop(['ID da Parte Adversa', 'Justificativa Estratégia Defesa', 'Centro de Custo_x', 'Gestor', 'Diretoria', 'Escritorio', 'Considerar no Score?',
       'Centro de custo Score'], axis=1, inplace=True)

    # reorganizar as colunas
    df_ajustado = df_ajustado[['ID', 'NPC', 'Parte Interessada',
       'Parte Adversa', 'Número Processo', 'Número do Processo Antigo',
       'Órgão', 'Nº Foro', 'Foro', 'Comarca', 'UF',
       'Processo Fisico ou Virtual', 'Sistema Eletrônico',
       'Sistema de Acompanhamento', 'Juiz', 'Corréu', 'Advogado Adverso',
       'UF Advogado Adverso', 'OAB Advogado Adverso', 'Escritório Adverso',
       'Valor Causa', 'Assunto', 'Origem', 'Tipo de Ação', 'Rito', 'Instância',
       'Data Distribuição', 'Data Citação', 'Data de Recebimento',
       'Data Cadastro', 'Usuário Cadastro', 'Data Revisão', 'Data Aceite',
       'Data e Hora do Encerramento do Processo', 'Data da Baixa', 'Contrato',
       'Projeto', 'Ramo', 'Produto', 'Objeto', 'Sub-Objeto', 'Estratégia', 'Tutela Deferida', 'Fase', 'Célula',
       'Advogado Responsável', 'Segmento', 'Cliente', 'Centro de Custo_y', 'Tipo Processo',
       'Status', 'Processo Estratégico?', 'Processo Problemático?',
       'Usuário Revisão', 'Responsável pela análise de encerramento',
       'Data Verificação', 'Origem.1', 'Data Solicitação Cliente',
       'Tipo Encerramento', 'Observação de Encerramento',
       'Responsável pela Solicitação do Encerramento ao Cliente', 'Valor Pago',
       'Observação', 'CPF/CNPJ - Parte Adversa', 'CPF/CNPJ - Cliente',
       'CPF/CNPJ - Parte Interessada', 'Data de Reativação do Processo',
       'Responsável pela Reativação do Processo', 'Forma de Cadastro',
       'Tipo de Recurso', 'Segredo de Justiça?', 'Importado por Planilha',
       'Iniciar Fluxo de Revisão de Processo por Importação ?',
       'Iniciar Fluxo de Prazo Automático por Importação ?', 'Massificado?',
       'Revelia?', 'Valor Pedido', 'Local de Trabalho', 'Função',
       'Valor a Provisionar', 'Encerrado Por']]
    
    df_ajustado.rename(columns = {'Centro de Custo_y': 'Centro de Custo'}, inplace=True)

    df_ajustado["CPF/CNPJ - Cliente"].fillna(df_ajustado["CPF/CNPJ - Parte Interessada"], inplace=True)

    df_final = df_ajustado.loc[(df_ajustado['Centro de Custo'] != 'BANCO INTER') &
                                (df_ajustado['Centro de Custo'] != 'LENOVO') &
                                (df_ajustado['Centro de Custo'] != 'CONTENCIOSO 6') &
                                (df_ajustado['Centro de Custo'] != 'VIA REGULATORIO') &
                                ~(df_ajustado['Centro de Custo'].isna())]

    df_bahia = df_ajustado[(df_ajustado['Centro de Custo'] == 'BANCO INTER') |
                                (df_ajustado['Centro de Custo'] == 'LENOVO') |
                                (df_ajustado['Centro de Custo'] == 'CONTENCIOSO 6') |
                                (df_ajustado['Centro de Custo'] == 'VIA REGULATORIO')]
    
    # df_final.to_excel(f'BASE_TRATADA_AGENDADOS_{hoje}.xlsx', index=False, engine='openpyxl')
    return df_final

# -- BASE BAHIA
def tratamento_base_bahia(base, base_centro_custo):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados.
    '''
    hoje = date.today()
    df = pd.read_excel(base)
    
    centro_custo = pd.read_excel(base_centro_custo)
    df_ajustado = pd.merge(df, centro_custo, on='Célula', how='left')
    df_ajustado.drop(['ID da Parte Adversa', 'Justificativa Estratégia Defesa', 'Centro de Custo_x', 'Gestor', 'Diretoria', 'Escritorio', 'Considerar no Score?',
       'Centro de custo Score'], axis=1, inplace=True)

    # reorganizar as colunas
    df_ajustado = df_ajustado[['ID', 'NPC', 'Parte Interessada',
       'Parte Adversa', 'Número Processo', 'Número do Processo Antigo',
       'Órgão', 'Nº Foro', 'Foro', 'Comarca', 'UF',
       'Processo Fisico ou Virtual', 'Sistema Eletrônico',
       'Sistema de Acompanhamento', 'Juiz', 'Corréu', 'Advogado Adverso',
       'UF Advogado Adverso', 'OAB Advogado Adverso', 'Escritório Adverso',
       'Valor Causa', 'Assunto', 'Origem', 'Tipo de Ação', 'Rito', 'Instância',
       'Data Distribuição', 'Data Citação', 'Data de Recebimento',
       'Data Cadastro', 'Usuário Cadastro', 'Data Revisão', 'Data Aceite',
       'Data e Hora do Encerramento do Processo', 'Data da Baixa', 'Contrato',
       'Projeto', 'Ramo', 'Produto', 'Objeto', 'Sub-Objeto', 'Estratégia', 'Tutela Deferida', 'Fase', 'Célula',
       'Advogado Responsável', 'Segmento', 'Cliente', 'Centro de Custo_y', 'Tipo Processo',
       'Status', 'Processo Estratégico?', 'Processo Problemático?',
       'Usuário Revisão', 'Responsável pela análise de encerramento',
       'Data Verificação', 'Origem.1', 'Data Solicitação Cliente',
       'Tipo Encerramento', 'Observação de Encerramento',
       'Responsável pela Solicitação do Encerramento ao Cliente', 'Valor Pago',
       'Observação', 'CPF/CNPJ - Parte Adversa', 'CPF/CNPJ - Cliente',
       'CPF/CNPJ - Parte Interessada', 'Data de Reativação do Processo',
       'Responsável pela Reativação do Processo', 'Forma de Cadastro',
       'Tipo de Recurso', 'Segredo de Justiça?', 'Importado por Planilha',
       'Iniciar Fluxo de Revisão de Processo por Importação ?',
       'Iniciar Fluxo de Prazo Automático por Importação ?', 'Massificado?',
       'Revelia?', 'Valor Pedido', 'Local de Trabalho', 'Função',
       'Valor a Provisionar', 'Encerrado Por']]
    
    df_ajustado.rename(columns = {'Centro de Custo_y': 'Centro de Custo'}, inplace=True)
    
    df_ajustado["CPF/CNPJ - Cliente"].fillna(df_ajustado["CPF/CNPJ - Parte Interessada"], inplace=True)

    df_final = df_ajustado.loc[(df_ajustado['Centro de Custo'] != 'BANCO INTER') &
                                (df_ajustado['Centro de Custo'] != 'LENOVO') &
                                (df_ajustado['Centro de Custo'] != 'CONTENCIOSO 6') &
                                (df_ajustado['Centro de Custo'] != 'VIA REGULATORIO') &
                                ~(df_ajustado['Centro de Custo'].isna())]

    df_bahia = df_ajustado.loc[(df_ajustado['Centro de Custo'] == 'BANCO INTER') |
                                (df_ajustado['Centro de Custo'] == 'LENOVO') |
                                (df_ajustado['Centro de Custo'] == 'CONTENCIOSO 6') |
                                (df_ajustado['Centro de Custo'] == 'VIA REGULATORIO')]

    # df_final.to_excel(f'BASE_TRATADA_AGENDADOS_{hoje}.xlsx', index=False, engine='openpyxl')
    return df_bahia


st.set_page_config(
    page_title="Tratamento da Base de Dados (Serviços Internos)",
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
    st.title('Tratamento da Base de Dados (Serviços Internos)')
    st.info('Esse projeto irá ajudar você a fazer o tratamento das bases de dados de forma mais eficiente e automática.')            

hoje = date.today()

st.title("Tratamento da Base de Dados (Serviços Internos)")
st.subheader('Importe uma planilha')
st.markdown(hide_menu, unsafe_allow_html=True)


centro_de_custo = st.file_uploader('Importe o arquivo de Centro de Custo - Auditorias Diárias:', type='xlsx')
uploaded_file = st.file_uploader('Base de dados:', type='xlsx')
st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
st.markdown('---')
if uploaded_file is not None:
    gerais = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(gerais)
    tratar_button = st.button('Tratamento das bases')
    if tratar_button:
        gerais_tratado = tratamento_base(uploaded_file, centro_de_custo)
        pendentes_tratado = tratamento_base_bahia(uploaded_file, centro_de_custo)
        generate_excel_download_link(gerais_tratado)
        generate_excel_download_link_bahia(pendentes_tratado)
        st.success('As bases foram tratadas e estão disponíveis para download.')

# if choices == 'Bahia':
#     centro_de_custo = st.file_uploader('Importe o arquivo de Centro de Custo - Auditorias Diárias:', type='xlsx')
#     file_pendentes = st.file_uploader('Base de dados - BAHIA:', type='xlsx')
#     st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
#     st.markdown('---')
#     if file_pendentes is not None:
#         bahia = pd.read_excel(file_pendentes, engine='openpyxl')
#         st.dataframe(bahia)
#         tratar_pendentes_button = st.button('Tratamento dos dados da Bahia')
#         if tratar_pendentes_button:
#             pendentes_tratado = tratamento_base_bahia(file_pendentes, centro_de_custo)
#             generate_excel_download_link_bahia(pendentes_tratado)
#             st.success('A base foi tratada e está disponível para download.')

