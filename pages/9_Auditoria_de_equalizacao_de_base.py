import pandas as pd
import streamlit as st
import io
from utils import bg_page

# Configurações da Página no Streamlit
st.set_page_config(
    page_title="Auditoria de Equalização de Base",
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
    st.title('Auditoria de Equalização de Base')
    st.info('Automação Desenvolvida para Automatizar parte da Auditoria de Eq. de Base.')
    
st.title("Equalização de Base")

st.subheader("1. Faça o upload da base do cliente:")
st.markdown(hide_menu, unsafe_allow_html=True)
base_cliente = st.file_uploader("Selecione o arquivo da base do cliente", type=["xlsx"])

st.subheader("2. Faça o upload da base do Performa:")
st.markdown(hide_menu, unsafe_allow_html=True)
base_performa = st.file_uploader("Selecione o arquivo da base do Performa", type=["xlsx"])

st.subheader("Instruções para uso correto da automação:")

st.markdown("""
**Para conseguir utilizar a automação com precisão, certifique-se de fazer as alterações necessárias na base do cliente. São elas:**

1. Se encontrar a coluna correspondente ao **'NPC'** da base do Performa na base do cliente, renomeie-a para **'Cód. Causa'**. Caso a base do cliente não tenha essa informação, prossiga para a próxima etapa.

2. Altere o nome da coluna que representa o **'Número Processo'** na base do cliente para **'Numeração Única'**. 

**Essa padronização é fundamental para a correta identificação e acompanhamento dos processos.**
""")

# Initialize session state variables
if 'processado' not in st.session_state:
    st.session_state.processado = False
if 'resultados_nulos_performa' not in st.session_state:
    st.session_state.resultados_nulos_performa = None
if 'resultados_nulos_cliente' not in st.session_state:
    st.session_state.resultados_nulos_cliente = None

class Automacao:
    def __init__(self):
        self.df_cliente = None
        self.df_performa = None

    def normalizar_num_processo(self, df, coluna_num_processo):
        if df is self.df_performa:
            nova_coluna = f'{coluna_num_processo} na Base do Performa'
        elif df is self.df_cliente:
            nova_coluna = f'{coluna_num_processo} na Base do Cliente'
        df[nova_coluna] = df[coluna_num_processo]
        
        df[coluna_num_processo] = df[coluna_num_processo].astype(str)	
        df[coluna_num_processo] = df[coluna_num_processo].str.replace(r'[^\w\s]', '', regex=True)
        df[coluna_num_processo] = df[coluna_num_processo].str.replace(r'\s+', '', regex=True)
        
    def preencher_nulos(self):
        if 'Cód. Causa' not in self.df_cliente.columns:
            self.df_cliente['Cód. Causa'] = '0'
        try:
            # Preencher valores nulos no df_performa
            self.df_performa['NPC'].fillna('N/A', inplace=True)
            self.df_performa['Número Processo'].fillna('N/A', inplace=True)

            # Preencher valores nulos no df_cliente
            self.df_cliente['Cód. Causa'].fillna('N/A', inplace=True)
            self.df_cliente['Numeração Única'].fillna('N/A', inplace=True)
        except:
            pass
        
    def processar_automacao(self):
        if self.df_cliente is None or self.df_performa is None:
            st.error("Por favor, faça o upload de ambos os arquivos antes de processar.")
            return
        
        st.spinner("Processando dados...")
        self.normalizar_num_processo(self.df_cliente, 'Numeração Única')
        self.normalizar_num_processo(self.df_performa, 'Número Processo')
        
        self.preencher_nulos()
        
        # Converter colunas NPC/Cód. Causa e Número Processo/Numeração Única para o mesmo tipo (str) antes do merge
        try:
            self.df_performa['NPC'] = self.df_performa['NPC'].astype(str)
            self.df_cliente['Cód. Causa'] = self.df_cliente['Cód. Causa'].astype(str)
        except:
            pass
        try:
            self.df_performa['Número Processo'] = self.df_performa['Número Processo'].astype(str)
            self.df_cliente['Numeração Única'] = self.df_cliente['Numeração Única'].astype(str)
        except:
            pass
            
        # TRATAMENTOS BASE PERFORMA
        self.df_performa = self.df_performa[['ID', 'NPC', 'Número Processo', 'Data Cadastro', 'Data Revisão', 'Fase', 'Célula', 'Advogado Responsável', 'Cliente', 'Centro de Custo',
                                            'Tipo Processo', 'Status', 'Número Processo na Base do Performa']]

        try:
            df_performa_procv = pd.merge(self.df_performa, self.df_cliente[['Numeração Única']], left_on='Número Processo', right_on='Numeração Única', how="left")
        except Exception as e:
            st.error(f'Não foi possível realizar o PROCV na coluna de Número Processo da base do Performa: {e}')
            return

        try:
            df_performa_procv = pd.merge(df_performa_procv, self.df_cliente[['Cód. Causa']], left_on='NPC', right_on='Cód. Causa', how="left")
        except Exception as e:
            st.error(f'Não foi possível realizar o PROCV na coluna NPC da base do Performa: {e}')
            return

        condition = df_performa_procv['Cód. Causa'].isna() & df_performa_procv['Numeração Única'].isna()
        st.session_state.resultados_nulos_performa = df_performa_procv[condition].copy()
        st.session_state.n_resultados_nulos_performa = st.session_state.resultados_nulos_performa.shape[0]
        st.session_state.total_base_performa = self.df_performa.shape[0]

        # TRATAMENTOS BASE CLIENTE
        try:
            df_cliente_procv = pd.merge(self.df_cliente, self.df_performa[['Número Processo']], left_on='Numeração Única', right_on='Número Processo', how="left")
        except Exception as e:
            st.error(f'Não foi possível realizar o PROCV na coluna de Numeração Única da base do cliente: {e}')
            return

        try:
            df_cliente_procv = pd.merge(df_cliente_procv, self.df_performa[['NPC']], left_on='Cód. Causa', right_on='NPC', how="left")
        except Exception as e:
            st.error(f'Não foi possível realizar o PROCV na coluna Cód. Causa da base do cliente: {e}')
            return

        condition = df_cliente_procv['NPC'].isna() & df_cliente_procv['Número Processo'].isna()
        st.session_state.resultados_nulos_cliente = df_cliente_procv[condition].copy()
        st.session_state.n_resultados_nulos_cliente = st.session_state.resultados_nulos_cliente.shape[0]
        st.session_state.total_base_cliente = self.df_cliente.shape[0]

    def gerar_arquivo_excel(self, df, nome_arquivo):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        buffer.seek(0)
        return buffer

    def exibir_resultados(self):
        if st.session_state.total_base_performa == 0 or st.session_state.total_base_cliente == 0:
            st.error("Por favor, processe os dados antes de exibir os resultados.")
            return
        
        st.write(f'Base do Performa: {st.session_state.total_base_performa} processos')
        st.write(f'Base do Cliente: {st.session_state.total_base_cliente} processos')
        diferenca = st.session_state.total_base_performa - st.session_state.total_base_cliente
        st.write(f'Diferença: {diferenca}')
        st.write(f'Temos {st.session_state.n_resultados_nulos_performa} processos na base do Performa que não encontramos no cliente.')
        st.write(f'Temos {st.session_state.n_resultados_nulos_cliente} processos na base do cliente que não encontramos no Performa.')

# Criar uma instância da classe
automacao = Automacao()

# Botão de processar automação
if st.button('Iniciar Automação'):
    if base_cliente and base_performa:
        automacao.df_cliente = pd.read_excel(base_cliente)
        automacao.df_performa = pd.read_excel(base_performa)
        automacao.processar_automacao()
        st.session_state.processado = True

        # Exibir resultados automaticamente após o processamento
        automacao.exibir_resultados()
    else:
        st.error("Por favor, faça o upload de ambos os arquivos antes de processar.")

# Exibir botões apenas após o processamento
if st.session_state.processado:
    if st.session_state.resultados_nulos_performa is not None:
        resultado_performa = automacao.gerar_arquivo_excel(st.session_state.resultados_nulos_performa, 'base_performa_auditada.xlsx')
        st.download_button(
            label="Download Base do Performa Auditada",
            data=resultado_performa,
            file_name="base_performa_auditada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    if st.session_state.resultados_nulos_cliente is not None:
        resultado_cliente = automacao.gerar_arquivo_excel(st.session_state.resultados_nulos_cliente, 'base_cliente_auditada.xlsx')
        st.download_button(
            label="Download Base do Cliente Auditada",
            data=resultado_cliente,
            file_name="base_cliente_auditada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
