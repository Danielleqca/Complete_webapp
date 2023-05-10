import streamlit as st

from utils import bg_page

st.set_page_config(
    page_title="Home",
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

st.markdown('''
    # Bem vindo ao QCA DataBoost!
''')
col1, col2 = st.columns(2)
with col1:
    st.markdown('''
        ###### Acelere seus projetos administrativos com o QCA DataBoost. Simplifique a separação de planilhas, PDFs e o tratamento de bases de dados de forma automática e eficiente, economizando tempo e aumentando sua produtividade. Experimente agora mesmo e descubra como o QCA DataBoost pode revolucionar a maneira como você trabalha.
    ''')

st.markdown(hide_menu, unsafe_allow_html=True)

st.write('\n')
st.write('\n')
with st.container():
    st.markdown("""
        ##### Com a barra lateral, é possível acessar rapidamente as diversas funcionalidades do sistema. 👈""")
    st.write('\n')
    st.markdown("""
        ##### Para ter acesso exclusivo a cada sistema de automatização, segue links abaixo:
        1. [Divisão das abas em novas planilhas](https://separacaoplanilhas2.streamlit.app/)
        2. [Divisão das planilhas por partição](https://excel-subfiles-python.streamlit.app/)
        3. [Divisão de PDFs por Lista de Nomes e Habilitações](https://separacao-pdfs.streamlit.app/)
        4. [Tratamento da Base de Dados (Serviços Internos)](https://tratamento-basededados.streamlit.app/)
        5. [Tratamento da Base de Dados (Prazos Diários)](https://tratamento-automatico-prazos-diarios.streamlit.app/)
    """)