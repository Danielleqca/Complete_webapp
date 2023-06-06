import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, time, datetime, timedelta
import os

import base64
from io import StringIO, BytesIO

def generate_excel_download_link_nao_tratadas(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_NAO_TRATADAS_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link_ag_providencias(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_AG_PROVIDENCIAS_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)


def concatena_bases(base1, base2):
    df1 = pd.read_excel(base1)
    df2 = pd.read_excel(base2)
    df_concatenado = pd.concat([df1, df2], ignore_index=True)
    return df_concatenado


def download_publi(dataset1, dataset2):
    hoje = date.today()    
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_GERAL_PUBLICAÇÕES_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)



# -- BASE NÃO TRATADAS
def tratamento_nao_tratadas(base_nao_tratadas, base_processos):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados referentes publicações não tratadas.
    '''
    hoje = date.today()

    # importar a base de dados nao tratadas
    df = pd.read_excel(base_nao_tratadas)

    # dropando a coluna "conteudo" pra melhor eficiencia do processo
    df = df.drop(['Conteúdo'], axis=1)

    # importar a base de dados de fluxos de processos juridicos
    processos = pd.read_excel(base_processos)

    # procv para buscar as informações de centro de custo
    df_ajustado = pd.merge(df, processos, left_on='ID Grupo de Pesquisa', right_on='Grupo de Pesquisa', how='left', indicator=True)
    
    df_ajustado.drop(['NPC',
       'ID Grupo de Pesquisa', 'Célula_x', 'Centro de Custo_x', '_merge'], axis=1, inplace=True)

    df_ajustado.rename(columns = {'Centro de Custo_y': 'Centro de Custo', 'Célula_y': 'Célula'}, inplace=True)
    
    df_ajustado = df_ajustado[['Centro de Custo','Célula','Data Diário','Data Recebimento','Número Processo','ID Publicação','ID Processo','ID Publicação', 'Diário','Responsável Publicação','Status Publicação','Nome Encontrado','Data da Entrada da Publicação no Seven','Data do Tratamento','Grupo de Pesquisa']]

    df_final = df_ajustado[(df_ajustado['Centro de Custo'] != 'AMBEV') &
                                (df_ajustado['Centro de Custo'] != 'CCB MASSIFICADO') &
                                (df_ajustado['Centro de Custo'] != 'Apresentação QCA') &
                                (df_ajustado['Centro de Custo'] != 'EQUIPE PARAIBA') &
                                (df_ajustado['Centro de Custo'] != 'MONGERAL') &
                                (df_ajustado['Centro de Custo'] != 'MARITIMO E PORTUARIO') &
                                (df_ajustado['Centro de Custo'] != ' ') &
                                (df_ajustado['Centro de Custo'] != 'ADMINISTRACAO JUDICIAL') &
                                (df_ajustado['Centro de Custo'] != 'DIREITO PUBLICO')]
    
    df_remove_tributarios_consultivos = df_final[ (df_final['Célula'].str.contains('Tributário') == True) |
                                        (df_final['Célula'].str.contains('Consultivo') == True) ]

    df_final = df_final.drop(df_remove_tributarios_consultivos.index) 

    df_final = df_final.drop_duplicates(subset=['ID Publicação'])

    # df_final.to_excel(f'BASE_TRATADA_AGENDADOS_{hoje}.xlsx', index=False, engine='openpyxl')
    return df_final

# -- BASE NÃO TRATADAS
def tratamento_ag_providencia(base_ag_providencia, base_centro_custo):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados referentes publicações não tratadas.
    '''
    hoje = date.today()

    # importar a base de dados nao tratadas
    df2 = pd.read_excel(base_ag_providencia)

    # dropando a coluna "conteudo" pra melhor eficiencia do processo
    df2 = df2.drop(['Conteúdo'], axis=1)

    # importar a base de dados de fluxos de processos juridicos
    cc = pd.read_excel(base_centro_custo)

    # procv para buscar as informações de centro de custo
    df_ajustado2 = pd.merge(df2, cc, on='Célula', how='left')
    
    df_ajustado2.drop(['Centro de Custo_x', 'Gestor', 'Diretoria', 'Escritorio',
       'Considerar no Score?', 'Centro de custo Score'], axis=1, inplace=True)

    df_ajustado2.rename(columns = {'Centro de Custo_y': 'Centro de Custo'}, inplace=True)
    
    # ajustar a ordem das colunas para ficar igual ao que o cliente deseja
    df_ajustado2 = df_ajustado2[['Centro de Custo','Célula','Data Diário','Data Recebimento','Número Processo','ID Publicação','ID Processo','ID Publicação', 'Diário','Responsável Publicação','Status Publicação','Nome Encontrado','Data da Entrada da Publicação no Seven','Data do Tratamento']]

    df_final2 = df_ajustado2[(df_ajustado2['Centro de Custo'] != 'AMBEV') &
                                (df_ajustado2['Centro de Custo'] != 'CCB MASSIFICADO') &
                                (df_ajustado2['Centro de Custo'] != 'Apresentação QCA') &
                                (df_ajustado2['Centro de Custo'] != 'EQUIPE PARAIBA') &
                                (df_ajustado2['Centro de Custo'] != 'MONGERAL') &
                                (df_ajustado2['Centro de Custo'] != 'MARITIMO E PORTUARIO') &
                                (df_ajustado2['Centro de Custo'] != ' ') &
                                (df_ajustado2['Centro de Custo'] != 'ADMINISTRACAO JUDICIAL') &
                                (df_ajustado2['Centro de Custo'] != 'DIREITO PUBLICO')]
    
    df_remove_tributarios_consultivos = df_final2[ (df_final2['Célula'].str.contains('Tributário') == True) |
                                        (df_final2['Célula'].str.contains('Consultivo') == True) ]

    df_final2 = df_final2.drop(df_remove_tributarios_consultivos.index)  

    # df_final2.to_excel(f'BASE_TRATADA_AGENDADOS_{hoje}.xlsx', index=False, engine='openpyxl')
    return df_final2

st.set_page_config(page_title='Tratamento Automático',
                    layout='wide')

with st.sidebar:
    st.image('https://www.onepointltd.com/wp-content/uploads/2020/03/inno2.png')
    st.title('Tratamento das Planilhas Automático')
    choices = st.radio('Escolha o tratamento:', ('Não Tratadas', 'Aguardando Providências'))
    st.info('Esse projeto irá ajudar você a fazer o tratamento das bases de dados de forma mais eficiente e automática.')            

hoje = date.today()

st.markdown("### Tratamento Automático 📊")
st.markdown('#### Importe uma planilha')

if choices == 'Não Tratadas':
    processos = st.file_uploader('Importe o arquivo de Fluxo de Processos Jurídicos:', type='xlsx')
    uploaded_file = st.file_uploader('Base de dados - NÃO TRATADAS:', type='xlsx')
    st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
    st.markdown('---')
    if uploaded_file is not None:
        nao_tratadas = pd.read_excel(uploaded_file, engine='openpyxl')
        st.dataframe(nao_tratadas)
        tratar_button = st.button('Tratamento das Não Tratadas')
        if tratar_button:
            nao_tratadas_tratado = tratamento_nao_tratadas(uploaded_file, processos)
            generate_excel_download_link_nao_tratadas(nao_tratadas_tratado)
            st.success('A base foi tratada e está disponível para download.')

if choices == 'Aguardando Providências':
    centro_de_custo = st.file_uploader('Importe o arquivo de Centro de Custo - Auditorias Diárias:', type='xlsx')
    file_pendentes = st.file_uploader('Base de dados - AGUARDANDO PROVIDÊNCIAS:', type='xlsx')
    st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
    st.markdown('---')
    if file_pendentes is not None:
        pendentes = pd.read_excel(file_pendentes, engine='openpyxl')
        st.dataframe(pendentes)
        tratar_pendentes_button = st.button('Tratamento dos Aguardando Providências')
        if tratar_pendentes_button:
            pendentes_tratado = tratamento_ag_providencia(file_pendentes, centro_de_custo)
            generate_excel_download_link_ag_providencias(pendentes_tratado)
            st.success('A base foi tratada e está disponível para download.')

