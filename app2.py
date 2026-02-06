import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import json
import os

# Configuração da página
st.set_page_config(
    page_title="Expedição ARM Recap",
    page_icon="🔍",
    layout="wide"
)

LAST_ACCESS_FILE = "last_access.json"

TEMAS = {
    "Palmeiras": {
        "background": "#0a3d1f", "text": "#e8ffe8", "accent": "#00ff41",
        "secondary": "#1a5d3f", "card_bg": "rgba(16, 61, 31, 0.6)",
        "input_bg": "#0f4d29", "button_text": "#003d1a"
    },
    "Corinthians": {
        "background": "#1a1a1a", "text": "#f5f5f5", "accent": "#ffffff",
        "secondary": "#333333", "card_bg": "rgba(40, 40, 40, 0.6)",
        "input_bg": "#2a2a2a", "button_text": "#000000"
    },
    "São Paulo": {
        "background": "#6b0f2a", "text": "#fff8dc", "accent": "#ffed4e",
        "secondary": "#8b1538", "card_bg": "rgba(107, 15, 42, 0.6)",
        "input_bg": "#8b1538", "button_text": "#4a0515"
    },
    "Santos": {
        "background": "#e8e8e8", "text": "#1a1a1a", "accent": "#2c2c2c",
        "secondary": "#c0c0c0", "card_bg": "rgba(255, 255, 255, 0.7)",
        "input_bg": "#f5f5f5", "button_text": "#ffffff"
    }
}

def get_last_access():
    if os.path.exists(LAST_ACCESS_FILE):
        with open(LAST_ACCESS_FILE, 'r') as f:
            data = json.load(f)
            return data.get('last_access')
    return None

def save_current_access():
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    try:
        with open(LAST_ACCESS_FILE, 'w') as f:
            json.dump({'last_access': current_time}, f)
    except Exception:
        pass  # No ambiente cloud, pode falhar ao gravar, ignore
    return current_time

def aplicar_tema(tema_nome):
    tema = TEMAS[tema_nome]
    st.markdown(f"""
        <style>
        .stApp {{
            background-color: {tema['background']};
        }}
        /* Aqui segue o CSS customizado... (igual ao original) */
        </style>
    """, unsafe_allow_html=True)

last_access = get_last_access()
current_access = save_current_access()

st.sidebar.title("🎨 Configurações")
tema_selecionado = st.sidebar.selectbox(
    "Escolha o Tema:", list(TEMAS.keys()), index=0
)
aplicar_tema(tema_selecionado)

st.title("🔍 Expedição ARM Recap")

now = datetime.now()
data_abreviada = now.strftime("%d/%b/%y").upper()
hora_atual = now.strftime("%H:%M:%S")

col_clock1, col_clock2 = st.columns(2)
with col_clock1:
    st.markdown(
        f'<div class="digital-clock">📅 {data_abreviada} | ⏰ {hora_atual}</div>',
        unsafe_allow_html=True
    )
with col_clock2:
    if last_access:
        st.markdown(
            f'<div class="digital-clock">🕐 ÚLTIMO: {last_access}</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div class="digital-clock">🕐 PRIMEIRO ACESSO</div>',
            unsafe_allow_html=True
        )
st.markdown("---")

ARQUIVO_EXCEL = "CX PTMs.xlsx"
df = None

if os.path.exists(ARQUIVO_EXCEL):
    try:
        df = pd.read_excel(ARQUIVO_EXCEL)
        st.success(f"✅ Arquivo '{ARQUIVO_EXCEL}' carregado automaticamente!")
    except Exception as e:
        st.error(f"❌ Erro ao abrir o arquivo '{ARQUIVO_EXCEL}': {e}")

if df is None:
    st.warning(f"⚠️ Arquivo '{ARQUIVO_EXCEL}' não encontrado no repositório.")
    st.info("📁 Por favor, faça o upload do arquivo Excel:")
    uploaded_file = st.file_uploader("Carregar arquivo Excel", type=['xlsx'])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("✅ Arquivo carregado com sucesso!")
        except Exception as e:
            st.error(f"❌ Erro ao ler o arquivo enviado: {e}")

if df is not None:
    with st.expander("ℹ️ Informações do Arquivo"):
        st.write(f"**Total de linhas:** {len(df)}")
        st.write(f"**Total de colunas:** {len(df.columns)}")
        st.write(f"**Colunas disponíveis:** {' - '.join(df.columns)}")
  st.markdown("---")
    st.subheader("🔎 Buscar Números")
  busca_keys = [f'busca{i}' for i in range(1, 5)]
    for k in busca_keys:
        if k not in st.session_state:
            st.session_state[k] = ""
  cols = st.columns(2)
    numero_busca = []
    for i in range(2):
        numero_busca.append(
            cols[0].text_input(
                f"Busca {i+1}:", placeholder="Ex: 8455",
                key=f"input_busca{i+1}", value=st.session_state[busca_keys[i]]
            )
        )
        numero_busca.append(
            cols[1].text_input(
                f"Busca {i+3}:", placeholder="Ex: 9122",
                key=f"input_busca{i+3}", value=st.session_state[busca_keys[i+2]]
            )
        )
    col_btn1, col_btn2 = st.columns([2, 1])
    buscar = col_btn1.button("🔍 Buscar Todos", use_container_width=True)
    limpar = col_btn2.button("🗑️ Limpar", use_container_width=True)
    if limpar:
        for k in busca_keys:
            st.session_state[k] = ""
        st.rerun()
    for idx, val in enumerate(numero_busca):
        st.session_state[busca_keys[idx]] = val
  def buscar_numero(numero):
        resultados = []
        for coluna in df.columns:
            coluna_numerica = pd.to_numeric(df[coluna], errors='coerce')
            if pd.notnull(coluna_numerica).any() and numero in coluna_numerica.values:
                linhas = df[coluna_numerica == numero].index.tolist()
                for linha in linhas:
                    resultados.append({
                        'Número': numero,
                        'Posição': coluna,
                        'Linha': linha + 2
                    })
        return resultados
  if buscar:
        numeros_busca = [n.strip() for n in numero_busca if n.strip()]
        if numeros_busca:
            st.markdown("---")
            st.subheader("📊 Resultados da Busca")
            todos_resultados = []
            for numero_str in numeros_busca:
                try:
                    numero = int(numero_str)
                    resultados = buscar_numero(numero)
                    if resultados:
                        st.success(f"✅ Número **{numero}** - Encontrado {len(resultados)} resultado(s)")
                        for resultado in resultados:
                            col_res1, col_res2, col_res3 = st.columns([1, 2, 1])
                            with col_res1:
                                st.metric("Número", f"{resultado['Número']}")
                            with col_res2:
                                st.metric("Posição", resultado['Posição'])
                            with col_res3:
                                st.metric("Linha Excel", f"{resultado['Linha']}")
                            todos_resultados.append(resultado)
                        st.markdown("---")
                    else:
                        st.warning(f"⚠️ Número **{numero}** não encontrado.")
                except ValueError:
                    st.error(f"❌ '{numero_str}' não é um número válido!")
            if todos_resultados:
                st.subheader("📋 Tabela Consolidada de Resultados")
                df_resultados = pd.DataFrame(todos_resultados)
                st.dataframe(df_resultados, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ Por favor digite pelo menos um número para buscar!")
  st.markdown("---")
    with st.expander("📊 Visualizar Dados Completos"):
        st.dataframe(df, use_container_width=True)
else:
    st.markdown("---")
    st.subheader("📖 Como usar:")
    st.markdown("""
    1. **Escolha um tema** na barra lateral  
    2. **Adicione o arquivo 'CX PTMs.xlsx'** na raiz do repositório GitHub   
    3. **Digite até 4 números** que deseja encontrar simultaneamente  
    4. **Clique em Buscar Todos** para localizar os números  
    5. Use o botão **Limpar** para apagar todos os campos de busca  
    6. O sistema mostrará em qual **posição (coluna)** cada número foi encontrado  
    """)

st.markdown("""
    <div class="footer">
        Desenvolvido por Djalma A Barbosa 2026. Direitos reservados.
    </div>
""", unsafe_allow_html=True)