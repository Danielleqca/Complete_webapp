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
    visibility:hidden;
}

footer {
    visibility:visible;
    content: 'Teste teste teste';
}

footer:after {
    content:'Desenvolvido pela Eficiência Jurídica - Controladoria Jurídica';
    display:block;
    position:relative;
    color:grey;
}
</style>
"""

st.markdown('''
    # Bem vindo à QCA DataBoost!
''')
col1, col2 = st.columns(2)
with col1:
    st.markdown('''
        ###### Acelere seus projetos administrativos com o QCA DataBoost. Simplifique a separação de planilhas, PDFs e o tratamento de bases de dados de forma automática e eficiente, economizando tempo e aumentando sua produtividade. Experimente agora mesmo e descubra como o QCA DataBoost pode revolucionar a maneira como você trabalha.
    ''')

st.markdown(hide_menu, unsafe_allow_html=True)
