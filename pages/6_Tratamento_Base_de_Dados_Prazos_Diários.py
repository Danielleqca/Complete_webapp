import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, time, datetime, timedelta
import os

import base64
from io import StringIO, BytesIO
from utils import bg_page
import streamlit_theme

streamlit_theme.set_theme({'theme': 'config.toml'})

def generate_excel_download_link_agendados(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_TRATADA_AGENDADOS_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link_pendentes(df):
    hoje = date.today()
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="BASE_TRATADA_PENDENTES_{hoje}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

# -- BASE PRAZOS DIARIOS AGENDADOS
def tratamento_prazos_diarios_agendados(base_agendados, base_centro_custo):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados referentes aos prazos agendados.
    '''
    hoje = date.today()
    df = pd.read_excel(base_agendados)
    # dropando colunas que nao serao usadas
    df = df.drop(['NPC', 'Número do Órgão', 'Sistema Acompanhamento', 'Responsável Cancelamento',
                    'ID Obrigação', 'Observação ', 'Observação da conclusão', 'Projeto',
                    'Link de Acesso à Audiência'], axis=1)
    # criando duas novas colunas vazias
    df[' '] = ' '
    df[' '] = ' '
    # selecionar o tipo de atividade = diligencia e nao selecionar os protocolos e apagar todas as linhas
    df_remove = df[(df['Tipo de Atividade'] == 'Diligência') &
                        (df['Sub Tipo Compromisso'].str.contains('Protocolo') == False)]
    df = df.drop(df_remove.index)
    centro_custo = pd.read_excel(base_centro_custo)
    df_ajustado = pd.merge(df, centro_custo, on='Célula', how='left')
    df_ajustado.drop(['Centro de Custo_x', 'Gestor', 'Diretoria', 'Escritorio'], axis=1, inplace=True)

    df_ajustado = df_ajustado[['ID Processo', 'ID Prazo', 'Data Cadastro Prazo', 'Data do Prazo','Hora do Prazo', 'Tipo de Atividade', 'Sub Tipo Compromisso','Parte Cliente', 'Parte Adversa', 'Número do Processo', 'Órgão',
                                    'Comarca', 'UF', 'Tipo de Processo', 'Sistema Eletrônico','Status do Prazo', 'Data Inicio Compromisso', 'Data Conclusão','Data Protocolo', 'Data Auditoria Protocolo', 'Prazo para Protocolo',
                                    'Data Prazo Automático', 'Data Revisão', 'Prazo Revisão','Data de Reprovação', 'Data Cancelamento','Motivo do cancelamento do Agendamento', 'Observação de Cancelamento',
                                    'Fase', 'Estratégia', 'Objeto', 'Sub-Objeto', 'Produto','Adv. Responsável Processo', 'Responsável Cadastro','Responsavel Prazo', 'Responsável Revisão', 'Responsável Conclusão',
                                    'Responsável Protocolo', 'Responsável Auditoria Protocolo', 'Célula','Centro de Custo_y','Cliente', 'Segmento', 'Forma Abertura', 'ID PA', 'ID Prazo Origem',
                                    ' ', ' ', 'Tem serviço contratado?', 'Id serviço', 'Tipo','Subtipo do serviço', 'Data da contratação', 'Hora da contratação','Status do serviço', 'Status Sugestão alteração/cancelamento',
                                    'Vinculado ao Recurso?', 'ID Recurso']]
    
    df_ajustado.rename(columns = {'Centro de Custo_y': 'Centro de Custo'}, inplace=True)

    df_final = df_ajustado[(df_ajustado['Centro de Custo'] != 'AMBEV') &
                                (df_ajustado['Centro de Custo'] != 'CCB MASSIFICADO') &
                                (df_ajustado['Centro de Custo'] != 'Apresentação QCA') &
                                (df_ajustado['Centro de Custo'] != 'EQUIPE PARAIBA') &
                                (df_ajustado['Centro de Custo'] != 'MONGERAL') &
                                (df_ajustado['Centro de Custo'] != 'MARITIMO E PORTUARIO') &
                                (df_ajustado['Centro de Custo'] != ' ') &
                                (df_ajustado['Centro de Custo'] != 'ADMINISTRACAO JUDICIAL')]
    
    df_remove_tributarios = df_final[df_final['Célula'].str.contains('Tributário') == True]

    df_final = df_final.drop(df_remove_tributarios.index)

    # exemplar de como alterar o formato da data e ainda deixar a tipagem como data
    # df['date'] = pd.to_datetime(df["date"].dt.strftime('%Y-%m'))

    df_final['Data Cadastro Prazo'] = pd.to_datetime(df_final['Data Cadastro Prazo'], format = ('%d/%m/%Y'))
    # df_final['Data Cadastro Prazo'] = df_final['Data Cadastro Prazo'].dt.strftime('%d/%m/%Y')
    
    df_final['Data do Prazo'] = pd.to_datetime(df_final['Data do Prazo'], format = ('%d/%m/%Y'))
    # df_final['Data do Prazo'] = df_final['Data do Prazo'].dt.strftime('%d/%m/%Y')

    df_final['Data Inicio Compromisso'] = pd.to_datetime(df_final['Data Inicio Compromisso'], format = ('%d/%m/%Y'))
    # df_final['Data Inicio Compromisso'] = df_final['Data Inicio Compromisso'].dt.strftime('%d/%m/%Y')

    df_final['Data Conclusão'] = pd.to_datetime(df_final['Data Conclusão'], format = ('%d/%m/%Y'))
    # df_final['Data Conclusão'] = df_final['Data Conclusão'].dt.strftime('%d/%m/%Y')

    df_final['Data Protocolo'] = pd.to_datetime(df_final['Data Protocolo'], format = ('%d/%m/%Y'))
    # df_final['Data Protocolo'] = df_final['Data Protocolo'].dt.strftime('%d/%m/%Y')

    df_final['Data Auditoria Protocolo'] = pd.to_datetime(df_final['Data Auditoria Protocolo'], format = ('%d/%m/%Y'))
    # df_final['Data Auditoria Protocolo'] = df_final['Data Auditoria Protocolo'].dt.strftime('%d/%m/%Y')

    df_final['Prazo para Protocolo'] = pd.to_datetime(df_final['Prazo para Protocolo'], format = ('%d/%m/%Y'))
    # df_final['Prazo para Protocolo'] = df_final['Prazo para Protocolo'].dt.strftime('%d/%m/%Y')

    df_final['Data Prazo Automático'] = pd.to_datetime(df_final['Data Prazo Automático'], format = ('%d/%m/%Y'))
    # df_final['Data Prazo Automático'] = df_final['Data Prazo Automático'].dt.strftime('%d/%m/%Y')
    
    df_final['Data Revisão'] = pd.to_datetime(df_final['Data Revisão'], format = ('%d/%m/%Y'))
    # df_final['Data Revisão'] = df_final['Data Revisão'].dt.strftime('%d/%m/%Y')
    
    df_final['Prazo Revisão'] = pd.to_datetime(df_final['Prazo Revisão'], format = ('%d/%m/%Y'))
    # df_final['Prazo Revisão'] = df_final['Prazo Revisão'].dt.strftime('%d/%m/%Y')
    
    df_final['Data Cancelamento'] = pd.to_datetime(df_final['Data Cancelamento'], format = ('%d/%m/%Y'))
    # df_final['Data Cancelamento'] = df_final['Data Cancelamento'].dt.strftime('%d/%m/%Y')
    
    df_final['Data da contratação'] = pd.to_datetime(df_final['Data da contratação'], format = ('%d/%m/%Y'))
    # df_final['Data da contratação'] = df_final['Data da contratação'].dt.strftime('%d/%m/%Y')
    

    # df_final['Data Cadastro Prazo'] = df_final.loc[:, ('Data Cadastro Prazo')].dt.strftime("%d/%m/%Y")
    # df_final['Data do Prazo'] = df_final.loc[:, ('Data do Prazo')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Inicio Compromisso')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Conclusão')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Protocolo')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Auditoria Protocolo')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Prazo para Protocolo')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Prazo Automático')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Revisão')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Prazo Revisão')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data Cancelamento')].dt.strftime("%d/%m/%Y")
    # df_final[''] = df_final.loc[:, ('Data da contratação')].dt.strftime("%d/%m/%Y")

    # df_final.to_excel(f'BASE_TRATADA_AGENDADOS_{hoje}.xlsx', index=False, engine='openpyxl')
    return df_final

def tratamento_prazos_diarios_pendentes(base_pendentes, base_centro_custo):
    '''
    Essa função vai auxiliar no tratamento automático da base de dados referentes aos prazos pendentes.
    '''
    hoje = date.today()
    df2 = pd.read_excel(base_pendentes, engine='openpyxl')

    df2 = df2.drop(['NPC', 'Número do Órgão', 'Sistema Acompanhamento', 'Responsável Cancelamento',
                            'ID Obrigação', 'Observação ', 'Observação da conclusão', 'Projeto', 'Link de Acesso à Audiência'], axis=1)   # dropando as colunas desnecessarias
    # criacao de novas colunas vazias para colocar entre as colunas id prazo origem e tem servico contratado
    df2[' '] = ' '
    df2[' '] = ' '
    # selecionar o tipo de atividade = diligencia e nao selecionar os protocolos e apagar todas as linhas
    df2_remove = df2[(df2['Tipo de Atividade'] == 'Diligência') &
        (df2['Sub Tipo Compromisso'].str.contains('Protocolo') == False)]

    df2 = df2.drop(df2_remove.index)
    centro_custo = pd.read_excel(base_centro_custo, engine='openpyxl')
    df2_ajustado = pd.merge(df2, centro_custo, on='Célula', how='left')
    df2_ajustado.drop(['Centro de Custo_x', 'Gestor', 'Diretoria', 'Escritorio'], axis=1, inplace=True)

    df2_ajustado = df2_ajustado[['ID Processo', 'ID Prazo', 'Data Cadastro Prazo', 'Data do Prazo','Hora do Prazo', 'Tipo de Atividade', 'Sub Tipo Compromisso','Parte Cliente', 'Parte Adversa', 'Número do Processo', 'Órgão',
        'Comarca', 'UF', 'Tipo de Processo', 'Sistema Eletrônico','Status do Prazo', 'Data Inicio Compromisso', 'Data Conclusão','Data Protocolo', 'Data Auditoria Protocolo', 'Prazo para Protocolo',
        'Data Prazo Automático', 'Data Revisão', 'Prazo Revisão','Data de Reprovação', 'Data Cancelamento','Motivo do cancelamento do Agendamento', 'Observação de Cancelamento',
        'Fase', 'Estratégia', 'Objeto', 'Sub-Objeto', 'Produto','Adv. Responsável Processo', 'Responsável Cadastro','Responsavel Prazo', 'Responsável Revisão', 'Responsável Conclusão',
        'Responsável Protocolo', 'Responsável Auditoria Protocolo', 'Célula','Centro de Custo_y','Cliente', 'Segmento', 'Forma Abertura', 'ID PA', 'ID Prazo Origem',
        ' ', ' ', 'Tem serviço contratado?', 'Id serviço', 'Tipo','Subtipo do serviço', 'Data da contratação', 'Hora da contratação','Status do serviço', 'Status Sugestão alteração/cancelamento',
        'Vinculado ao Recurso?', 'ID Recurso']]

    df2_ajustado.rename(columns = {'Centro de Custo_y': 'Centro de Custo'}, inplace=True)

    # Acrescentar "Administração Judicial" que não estará mais presente no escritório
    df2_final = df2_ajustado[(df2_ajustado['Centro de Custo'] != 'AMBEV') &
                (df2_ajustado['Centro de Custo'] != 'CCB MASSIFICADO') &
                (df2_ajustado['Centro de Custo'] != 'Apresentação QCA') &
                (df2_ajustado['Centro de Custo'] != 'EQUIPE PARAIBA') &
                (df2_ajustado['Centro de Custo'] != 'MONGERAL') &
                (df2_ajustado['Centro de Custo'] != 'MARITIMO E PORTUARIO') &
                (df2_ajustado['Centro de Custo'] != ' ') &
                (df2_ajustado['Centro de Custo'] != 'ADMINISTRACAO JUDICIAL')]
    
    df_remove_tributarios2 = df2_final[df2_final['Célula'].str.contains('Tributário') == True]

    df2_final = df2_final.drop(df_remove_tributarios2.index)

    # Filtragens das datas
    filtro1 = (((df2_final['Status do Prazo'] == 'Pendente') | (df2_final['Status do Prazo'] == 'Nao Protocolada')) & (df2_final['Data Prazo Automático'].isnull()))
    df2_final.loc[filtro1, "Data Prazo Automático"] = df2_final["Data do Prazo"]

    filtro2 = df2_final['Status do Prazo'] == 'Pendente'
    df2_final.loc[filtro2, "Data Prazo Automático"] = df2_final["Data do Prazo"]

    filtro3 = df2_final['Status do Prazo'] == 'Aguardando Auditoria'
    df2_final.loc[filtro3, "Data Prazo Automático"] = df2_final['Prazo para Protocolo'] + timedelta(days = 2)

    filtro4 = ((df2_final['Status do Prazo'] == 'Aguardando Revisão') & (df2_final['Prazo Revisão'] < pd.to_datetime(hoje)))
    df2_final.loc[filtro4, "Data Prazo Automático"] = df2_final['Prazo para Protocolo']

    filtro5 = ((df2_final['Status do Prazo'] == 'Aguardando Revisão') & (df2_final['Prazo Revisão'] >= pd.to_datetime(hoje)))
    df2_final.loc[filtro5, "Data Prazo Automático"] = df2_final['Prazo Revisão']    

    df2_final['Data Cadastro Prazo'] = pd.to_datetime(df2_final['Data Cadastro Prazo'], format = ('%d/%m/%Y'))
    # df2_final['Data Cadastro Prazo'] = df2_final['Data Cadastro Prazo'].dt.strftime('%d/%m/%Y')
    
    df2_final['Data do Prazo'] = pd.to_datetime(df2_final['Data do Prazo'], format = ('%d/%m/%Y'))
    # df2_final['Data do Prazo'] = df2_final['Data do Prazo'].dt.strftime('%d/%m/%Y')

    df2_final['Data Inicio Compromisso'] = pd.to_datetime(df2_final['Data Inicio Compromisso'], format = ('%d/%m/%Y'))
    # df2_final['Data Inicio Compromisso'] = df2_final['Data Inicio Compromisso'].dt.strftime('%d/%m/%Y')

    df2_final['Data Conclusão'] = pd.to_datetime(df2_final['Data Conclusão'], format = ('%d/%m/%Y'))
    # df2_final['Data Conclusão'] = df2_final['Data Conclusão'].dt.strftime('%d/%m/%Y')

    df2_final['Data Protocolo'] = pd.to_datetime(df2_final['Data Protocolo'], format = ('%d/%m/%Y'))
    # df2_final['Data Protocolo'] = df2_final['Data Protocolo'].dt.strftime('%d/%m/%Y')

    df2_final['Data Auditoria Protocolo'] = pd.to_datetime(df2_final['Data Auditoria Protocolo'], format = ('%d/%m/%Y'))
    # df2_final['Data Auditoria Protocolo'] = df2_final['Data Auditoria Protocolo'].dt.strftime('%d/%m/%Y')

    df2_final['Prazo para Protocolo'] = pd.to_datetime(df2_final['Prazo para Protocolo'], format = ('%d/%m/%Y'))
    # df2_final['Prazo para Protocolo'] = df2_final['Prazo para Protocolo'].dt.strftime('%d/%m/%Y')

    df2_final['Data Prazo Automático'] = pd.to_datetime(df2_final['Data Prazo Automático'], format = ('%d/%m/%Y'))
    # df2_final['Data Prazo Automático'] = df2_final['Data Prazo Automático'].dt.strftime('%d/%m/%Y')
    
    df2_final['Data Revisão'] = pd.to_datetime(df2_final['Data Revisão'], format = ('%d/%m/%Y'))
    # df2_final['Data Revisão'] = df2_final['Data Revisão'].dt.strftime('%d/%m/%Y')
    
    df2_final['Prazo Revisão'] = pd.to_datetime(df2_final['Prazo Revisão'], format = ('%d/%m/%Y'))
    # df2_final['Prazo Revisão'] = df2_final['Prazo Revisão'].dt.strftime('%d/%m/%Y')
    
    df2_final['Data Cancelamento'] = pd.to_datetime(df2_final['Data Cancelamento'], format = ('%d/%m/%Y'))
    # df2_final['Data Cancelamento'] = df2_final['Data Cancelamento'].dt.strftime('%d/%m/%Y')
    
    df2_final['Data da contratação'] = pd.to_datetime(df2_final['Data da contratação'], format = ('%d/%m/%Y'))
    # df2_final['Data da contratação'] = df2_final['Data da contratação'].dt.strftime('%d/%m/%Y')

    # # Fazer a filtragem das datas anteriores a 2020 e remover as linhas
    # df_remove_data_under_2020 = df2_final['Data Prazo Automático'] < pd.to_datetime(date(2021,1,1))

    # df2_final = df2_final.drop(df_remove_data_under_2020.index)

    # df2_final.to_excel(f'BASE_TRATADA_PENDENTES_{hoje}.xlsx', index=False, engine='openpyxl')
    return df2_final

st.set_page_config(
    page_title="Tratamento da Base de Dados - Prazos Diários",
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
    st.title('Tratamento da Base de Dados - Prazos Diários')
    choices = st.radio('Escolha o tratamento:', ('Agendados', 'Pendentes'))
    st.info('Esse projeto irá ajudar você a fazer o tratamento das bases de dados de forma mais eficiente e automática.')            

hoje = date.today()

st.title("Tratamento da Base de Dados - Prazos Diários")
st.subheader('Importe uma planilha')
st.markdown(hide_menu, unsafe_allow_html=True)

if choices == 'Agendados':
    centro_de_custo = st.file_uploader('Importe o arquivo de Centro de Custo - Auditorias Diárias:', type='xlsx')
    uploaded_file = st.file_uploader('Base de dados - AGENDADOS:', type='xlsx')
    st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
    st.markdown('---')
    if uploaded_file is not None:
        agendados = pd.read_excel(uploaded_file, engine='openpyxl')
        st.dataframe(agendados)
        tratar_button = st.button('Tratamento dos agendados')
        if tratar_button:
            agendados_tratado = tratamento_prazos_diarios_agendados(uploaded_file, centro_de_custo)
            generate_excel_download_link_agendados(agendados_tratado)
            st.success('A base foi tratada e está disponível para download.')

if choices == 'Pendentes':
    centro_de_custo = st.file_uploader('Importe o arquivo de Centro de Custo - Auditorias Diárias:', type='xlsx')
    file_pendentes = st.file_uploader('Base de dados - PENDENTES:', type='xlsx')
    st.warning('⚠️ Os arquivos precisam ser no formato Excel (.xlsx)')
    st.markdown('---')
    if file_pendentes is not None:
        pendentes = pd.read_excel(file_pendentes, engine='openpyxl')
        st.dataframe(pendentes)
        tratar_pendentes_button = st.button('Tratamento dos pendentes')
        if tratar_pendentes_button:
            pendentes_tratado = tratamento_prazos_diarios_pendentes(file_pendentes, centro_de_custo)
            generate_excel_download_link_pendentes(pendentes_tratado)
            st.success('A base foi tratada e está disponível para download.')

