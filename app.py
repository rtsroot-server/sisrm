import streamlit as st
import pandas as pd
import io
import re

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

# Função: Formatar o Telefone para ( 21 ) XXXXX-XXXX
def formatar_telefone(numero):
    if pd.isna(numero) or numero == "":
        return numero
    
    # Remove tudo que não for número (tira espaços, letras, traços antigos)
    numeros_limpos = re.sub(r'\D', '', str(numero))
    
    if len(numeros_limpos) == 0:
        return numero
        
    # Pega os últimos 9 dígitos (para ignorar códigos de país ou DDDs antigos)
    if len(numeros_limpos) >= 9:
        cel = numeros_limpos[-9:]
        return f"( 21 ) {cel[:5]}-{cel[5:]}"
    # Se a pessoa digitou só 8 números (esqueceu o 9), ele adiciona o 9
    elif len(numeros_limpos) == 8:
        return f"( 21 ) 9{numeros_limpos[:4]}-{numeros_limpos[4:]}"
    else:
        return numero # Se for um número inválido/curto demais, mantém original

st.title("📊 Limpeza de Planilhas Inteligente")
st.markdown("Faça o upload da sua planilha (ex: **TATA 1K.xlsx**) para identificar e tratar dados duplicados.")

# 2. Gerenciamento de Estado
if 'step' not in st.session_state:
    st.session_state.step = 'upload'
if 'df' not in st.session_state:
    st.session_state.df = None

# --- TELA 1: UPLOAD ---
if st.session_state.step == 'upload':
    uploaded_file = st.file_uploader("Arraste ou selecione seu arquivo Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        # CORREÇÃO APLICADA AQUI: 
        # Convertendo os nomes das colunas para minúsculo na hora da verificação
        # Isso impede erros caso o cabeçalho esteja como TELEFONE, Telefone, telefone, etc.
        colunas_telefone = ['telefone', 'celular', 'contato', 'numero']
        for col in df.columns:
            if str(col).lower().strip() in colunas_telefone:
                df[col] = df[col].apply(formatar_telefone)
        
        st.session_state.df = df
        st.session_state.step = 'analise'
        st.rerun()

# --- TELA 2: ANÁLISE DE DUPLICADOS ---
if st.session_state.step == 'analise':
    df = st.session_state.df
    
    duplicados = df[df.duplicated(keep=False)]
    
    if duplicados.empty:
        st.success("✨ Parabéns! Nenhuma linha duplicada foi encontrada na sua planilha.")
        st.dataframe(df, use_container_width=True)
        
        excel_data = to_excel(df)
        st.download_button("📥 Baixar Planilha Formatada", data=excel_data, file_name="planilha_formatada.xlsx")
        
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
