import streamlit as st
import pandas as pd
import io

# 1. Configuração da página e Layout UX
st.set_page_config(
    page_title="Gestor de Planilhas", 
    page_icon="📊", 
    layout="centered"
)

# CSS Customizado para deixar os botões e a tela mais bonitos
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    h1 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    </style>
""", unsafe_allow_html=True)

# Função para converter o dataframe para Excel e permitir o download
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Planilha1')
    return output.getvalue()

st.title("📊 Limpeza de Planilhas Inteligente")
st.markdown("Faça o upload da sua planilha (ex: **TATA 1K.xlsx**) para identificar e tratar dados duplicados.")

# 2. Gerenciamento de Estado (Para saber em qual tela o usuário está)
if 'step' not in st.session_state:
    st.session_state.step = 'upload'
if 'df' not in st.session_state:
    st.session_state.df = None

# --- TELA 1: UPLOAD ---
if st.session_state.step == 'upload':
    uploaded_file = st.file_uploader("Arraste ou selecione seu arquivo Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        st.session_state.df = pd.read_excel(uploaded_file)
        st.session_state.step = 'analise'
        st.rerun()

# --- TELA 2: ANÁLISE DE DUPLICADOS ---
if st.session_state.step == 'analise':
    df = st.session_state.df
    
    # Identifica linhas onde todos os valores das 3 colunas são iguais (mostra todas as ocorrências)
    duplicados = df[df.duplicated(keep=False)]
    
    if duplicados.empty:
        st.success("✨ Parabéns! Nenhuma linha duplicada foi encontrada na sua planilha.")
        if st.button("Sair / Anexar Outra"):
            st.session_state.step = 'upload'
            st.session_state.df = None
            st.rerun()
    else:
        st.warning(f"⚠️ Encontramos {df.duplicated().sum()} linhas duplicadas. Veja abaixo:")
        st.dataframe(duplicados, use_container_width=True)
        
        st.markdown("### O que você deseja fazer?")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("🗑️ Excluir", type="primary", help="Remove as duplicatas e mantém apenas a primeira"):
                st.session_state.step = 'excluir'
                st.rerun()
        with col2:
            if st.button("✏️ Alterar", help="Edite a planilha manualmente"):
                st.session_state.step = 'alterar'
                st.rerun()
        with col3:
            if st.button("🚪 Sair", help="Volta para a tela de upload"):
                st.session_state.step = 'upload'
                st.session_state.df = None
                st.rerun()

# --- TELA 3: EXCLUIR DUPLICADOS ---
if st.session_state.step == 'excluir':
    df = st.session_state.df
    # Remove duplicatas mantendo a primeira ocorrência
    df_clean = df.drop_duplicates()
    
    st.success("✅ Duplicatas excluídas com sucesso! Apenas registros únicos foram mantidos.")
    st.dataframe(df_clean, use_container_width=True)
    
    excel_data = to_excel(df_clean)
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📥 Baixar Planilha Limpa", data=excel_data, file_name="planilha_sem_duplicatas.xlsx")
    with col2:
        if st.button("Voltar ao Início"):
            st.session_state.step = 'upload'
            st.session_state.df = None
            st.rerun()

# --- TELA 4: ALTERAR DADOS ---
if st.session_state.step == 'alterar':
    st.info("✏️ Clique duas vezes em qualquer célula abaixo para alterar os valores ou adicione/remova linhas.")
    df = st.session_state.df
    
    # Tabela interativa onde o usuário pode editar os dados como no Excel
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    
    excel_data = to_excel(edited_df)
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📥 Baixar Planilha Alterada", data=excel_data, file_name="planilha_alterada.xlsx")
    with col2:
        if st.button("Voltar ao Início"):
            st.session_state.step = 'upload'
            st.session_state.df = None
            st.rerun()