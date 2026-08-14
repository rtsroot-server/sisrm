import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. Configuração da página - Novo Nome
st.set_page_config(
    page_title="Central de Dados", 
    page_icon="📱", 
    layout="centered"
)

# 2. CSS Customizado
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp { background-color: #f4f7f6; }
    h1, h2, h3 { color: #1e3d59; font-family: 'Segoe UI', sans-serif; text-align: center; }
    .stButton>button { 
        width: 100%; border-radius: 25px; font-weight: bold; 
        background-color: #17c3b2; color: white; border: none;
        padding: 12px 24px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #13a294; transform: translateY(-2px); box-shadow: 0 6px 8px rgba(0,0,0,0.15);
    }
    [data-testid="stFileUploadDropzone"] {
        border-radius: 20px; border: 2px dashed #17c3b2; background-color: #ffffff;
        padding: 30px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);
    }
    .stAlert { border-radius: 15px; }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE LÓGICA ---

def formatar_telefone(numero):
    if pd.isna(numero) or str(numero).strip() == "":
        return ""
    numeros_limpos = re.sub(r'\D', '', str(numero))
    if len(numeros_limpos) == 0:
        return ""
    if len(numeros_limpos) == 11:
        return f"( {numeros_limpos[:2]} ) {numeros_limpos[2:7]}-{numeros_limpos[7:]}"
    elif len(numeros_limpos) == 10:
        return f"( {numeros_limpos[:2]} ) 9{numeros_limpos[2:6]}-{numeros_limpos[6:]}"
    elif len(numeros_limpos) == 9:
        return f"( 21 ) {numeros_limpos[:5]}-{numeros_limpos[5:]}"
    elif len(numeros_limpos) == 8:
        return f"( 21 ) 9{numeros_limpos[:4]}-{numeros_limpos[4:]}"
    else:
        return numeros_limpos

def limpar_nome_e_bairro(nome_raw, bairro_raw):
    nome = str(nome_raw).strip()
    bairro = str(bairro_raw).strip()
    
    if " - " in nome or "-" in nome:
        partes = nome.split("-", 1)
        nome = partes[0].strip()
        bairro_extraido = partes[1].strip()
        if not bairro or bairro.lower() == 'nan':
            bairro = bairro_extraido

    bairro_lower = bairro.lower()
    bairro_invalido = False
    bairro_sem_espacos = bairro.replace(" ", "")
    
    if not bairro or bairro_lower == 'nan':
        bairro_invalido = True
    elif re.search(r'\d', bairro): 
        bairro_invalido = True
    elif len(bairro_sem_espacos) <= 3: 
        bairro_invalido = True
    elif "." in bairro and len(bairro_sem_espacos) <= 5: 
        bairro_invalido = True
    else:
        palavras_proibidas = ['gato', 'cachorro', 'cadela', 'adotou', 'adotante', 'devolveu', 'macho', 'fêmea', 'femea', 'filhote', 'pet', 'aa.c']
        for palavra in palavras_proibidas:
            if palavra in bairro_lower:
                bairro_invalido = True
                break
                
    if bairro_invalido:
        bairro = "Sem Bairro"
        
    return nome.title(), bairro.title()

def processar_planilha(df):
    tel_col = None
    nome_col = None
    bairro_col = None
    
    for c in df.columns:
        c_lower = str(c).lower().strip()
        if c_lower in ['telefone', 'celular', 'contato', 'numero']:
            tel_col = c
        elif 'nome' in c_lower or 'cliente' in c_lower:
            nome_col = c
        elif c_lower in ['região', 'regiao', 'bairro', 'unnamed: 1']:
            bairro_col = c

    dados_padronizados = []
    
    for _, row in df.iterrows():
        n_raw = str(row[nome_col]) if nome_col and not pd.isna(row[nome_col]) else "Sem Nome"
        b_raw = str(row[bairro_col]) if bairro_col and not pd.isna(row[bairro_col]) else ""
        t_raw = str(row[tel_col]) if tel_col and not pd.isna(row[tel_col]) else ""
        
        nome_limpo, bairro_limpo = limpar_nome_e_bairro(n_raw, b_raw)
        tel_formatado = formatar_telefone(t_raw)
        
        if tel_formatado != "":
            dados_padronizados.append({
                "Região/Bairro": bairro_limpo,
                "Nome": nome_limpo,
                "Telefone": tel_formatado
            })
            
    return pd.DataFrame(dados_padronizados)

def gerar_excel_final(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        bairros = df['Região/Bairro'].unique()
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="17C3B2", end_color="17C3B2", fill_type="solid")
        
        abas_usadas = {} 
        
        for bairro in bairros:
            nome_limpo_excel = re.sub(r'[\\/*?:\[\]]', '-', str(bairro)).strip()
            nome_aba_base = nome_limpo_excel[:31]
            if not nome_aba_base:
                nome_aba_base = "Sem Bairro"
                
            nome_aba = nome_aba_base
            contador = 1
            while nome_aba in abas_usadas.values():
                sufixo = str(contador)
                nome_aba = nome_aba_base[:31 - len(sufixo)] + sufixo
                contador += 1
                
            abas_usadas[bairro] = nome_aba
                
            df_bairro = df[df['Região/Bairro'] == bairro].copy()
            df_bairro['Agendamento de Visita'] = "" 
            
            df_bairro = df_bairro[['Região/Bairro', 'Nome', 'Telefone', 'Agendamento de Visita']]
            df_bairro.to_excel(writer, index=False, sheet_name=nome_aba)
            
            worksheet = writer.sheets[nome_aba]
            for col_num, column_title in enumerate(df_bairro.columns, 1):
                col_letra = get_column_letter(col_num)
                
                if column_title == 'Nome':
                    worksheet.column_dimensions[col_letra].width = 35
                elif column_title == 'Agendamento de Visita':
                    worksheet.column_dimensions[col_letra].width = 40
                else:
                    worksheet.column_dimensions[col_letra].width = 25
                
                celula = worksheet.cell(row=1, column=col_num)
                celula.font = header_font
                celula.fill = header_fill
                celula.alignment = Alignment(horizontal='center', vertical='center')

    return output.getvalue()

# --- GERENCIAMENTO DE ESTADO ---
if 'df_final' not in st.session_state:
    st.session_state.df_final = None

if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

# --- TELA PRINCIPAL (UI) ---

st.markdown("<h1>📱 Central de Dados</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #555;'>Unifique planilhas e arquivos de texto, formate contatos e organize automaticamente.</p>", unsafe_allow_html=True)

col1, col2 = st.columns([3, 1])

with col1:
    # Atualizado para aceitar csv e txt além de xlsx
    uploaded_files = st.file_uploader(
        "Toque ou arraste os arquivos aqui", 
        type=["xlsx", "csv", "txt"], 
        accept_multiple_files=True,
        key=str(st.session_state.uploader_key)
    )

with col2:
    st.markdown("<div style='margin-top: 35px;'></div>", unsafe_allow_html=True)
    if st.button("🗑️ Limpar Arquivos"):
        st.session_state.uploader_key += 1 
        st.session_state.df_final = None
        st.rerun()

if uploaded_files:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Processar Contatos", type="primary"):
        with st.spinner("Lendo e padronizando arquivos..."):
            lista_dfs = []
            
            for file in uploaded_files:
                extensao = file.name.split('.')[-1].lower()
                
                try:
                    # Lógica inteligente de leitura baseada no formato do arquivo
                    if extensao == 'csv':
                        try:
                            # Tenta ler CSV padrão separado por vírgula
                            df_bruto = pd.read_csv(file)
                        except:
                            # Se falhar, tenta ler separado por ponto e vírgula (padrão Excel BR)
                            file.seek(0)
                            df_bruto = pd.read_csv(file, sep=';', encoding='latin-1')
                    elif extensao == 'txt':
                        # TXTs geralmente vêm separados por tabulação (Tab)
                        df_bruto = pd.read_csv(file, sep='\t')
                    else:
                        # Padrão XLSX
                        df_bruto = pd.read_excel(file)
                    
                    df_limpo = processar_planilha(df_bruto)
                    lista_dfs.append(df_limpo)
                    
                except Exception as e:
                    st.error(f"Não foi possível ler o arquivo {file.name}. Verifique se ele não está corrompido.")

            if lista_dfs:
                df_unificado = pd.concat(lista_dfs, ignore_index=True)
                
                total_antes = len(df_unificado)
                df_final = df_unificado.drop_duplicates(subset=['Telefone'], keep='first')
                df_final = df_final.sort_values(by=['Região/Bairro', 'Nome'])
                
                total_depois = len(df_final)
                duplicadas = total_antes - total_depois
                
                st.session_state.df_final = df_final
                st.success(f"✨ Pronto! {duplicadas} telefones repetidos foram excluídos.")

if st.session_state.df_final is not None:
    df_resultado = st.session_state.df_final
    
    st.markdown("### 📊 Pré-visualização")
    st.dataframe(df_resultado, use_container_width=True)
    
    excel_pronto = gerar_excel_final(df_resultado)
    
    st.markdown("<br>", unsafe_allow_html=True)
    # Novo nome do botão conforme solicitado
    st.download_button(
        label="📥 Baixar Planilha", 
        data=excel_pronto, 
        file_name="Central_de_Dados_Limpa.xlsx", 
        type="primary"
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Iniciar Novo Processo"):
        st.session_state.df_final = None
        st.session_state.uploader_key += 1 
        st.rerun()
