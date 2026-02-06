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

# Arquivo para armazenar o último acesso
LAST_ACCESS_FILE = "last_access.json"

# Definição dos temas com cores melhoradas para melhor legibilidade
TEMAS = {
    "Palmeiras": {
        "background": "#0a3d1f",
        "text": "#e8ffe8",
        "accent": "#00ff41",
        "secondary": "#1a5d3f",
        "card_bg": "rgba(16, 61, 31, 0.6)",
        "input_bg": "#0f4d29",
        "button_text": "#003d1a"
    },
    "Corinthians": {
        "background": "#1a1a1a",
        "text": "#f5f5f5",
        "accent": "#ffffff",
        "secondary": "#333333",
        "card_bg": "rgba(40, 40, 40, 0.6)",
        "input_bg": "#2a2a2a",
        "button_text": "#000000"
    },
    "São Paulo": {
        "background": "#6b0f2a",
        "text": "#fff8dc",
        "accent": "#ffed4e",
        "secondary": "#8b1538",
        "card_bg": "rgba(107, 15, 42, 0.6)",
        "input_bg": "#8b1538",
        "button_text": "#4a0515"
    },
    "Santos": {
        "background": "#e8e8e8",
        "text": "#1a1a1a",
        "accent": "#2c2c2c",
        "secondary": "#c0c0c0",
        "card_bg": "rgba(255, 255, 255, 0.7)",
        "input_bg": "#f5f5f5",
        "button_text": "#ffffff"
    }
}

def get_last_access():
    """Recupera a data do último acesso"""
    if os.path.exists(LAST_ACCESS_FILE):
        with open(LAST_ACCESS_FILE, 'r') as f:
            data = json.load(f)
            return data.get('last_access')
    return None

def save_current_access():
    """Salva a data e hora do acesso atual"""
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(LAST_ACCESS_FILE, 'w') as f:
        json.dump({'last_access': current_time}, f)
    return current_time

def aplicar_tema(tema_nome):
    """Aplica o tema selecionado"""
    tema = TEMAS[tema_nome]
    
    # CSS customizado para o tema com melhor contraste e legibilidade
    st.markdown(f"""
        <style>
        .stApp {{
            background-color: {tema['background']};
        }}
        
        h1, h2, h3, h4, h5, h6 {{
            color: {tema['text']} !important;
            font-weight: {"bold" if tema_nome == "Santos" else "600"} !important;
        }}
        
        p, span, div, label {{
            color: {tema['text']} !important;
        }}
        
        .stTextInput > div > div > input {{
            color: {tema['text']} !important;
            background-color: {tema['input_bg']} !important;
            border: 2px solid {tema['accent']} !important;
            font-size: 16px !important;
            font-weight: 500 !important;
        }}
        
        .stSelectbox > div > div > div {{
            color: {tema['text']} !important;
            background-color: {tema['input_bg']} !important;
            border: 2px solid {tema['accent']} !important;
        }}
        
        .stButton > button {{
            background-color: {tema['accent']} !important;
            color: {tema['button_text']} !important;
            font-weight: bold !important;
            border: none !important;
            font-size: 16px !important;
            padding: 10px 20px !important;
        }}
        
        .stButton > button:hover {{
            opacity: 0.8 !important;
            transform: scale(1.02) !important;
        }}
        
        .stMetric {{
            background-color: {tema['card_bg']} !important;
            padding: 15px !important;
            border-radius: 8px !important;
            border: 1px solid {tema['accent']} !important;
        }}
        
        .stMetric label {{
            color: {tema['text']} !important;
            font-weight: {"bold" if tema_nome == "Santos" else "600"} !important;
            font-size: 14px !important;
        }}
        
        .stMetric > div {{
            color: {tema['accent']} !important;
            font-weight: bold !important;
            font-size: 24px !important;
        }}
        
        .digital-clock {{
            font-family: 'Courier New', monospace;
            font-size: 24px;
            font-weight: bold;
            color: {tema['accent']};
            text-align: center;
            padding: 15px;
            background-color: {tema['card_bg']};
            border-radius: 10px;
            border: 2px solid {tema['accent']};
            letter-spacing: 2px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }}
        
        .footer {{
            position: fixed;
            bottom: 10px;
            right: 20px;
            font-size: 10px;
            color: rgba(128, 128, 128, 0.4);
            font-style: italic;
        }}
        
        .stDataFrame {{
            background-color: {tema['card_bg']} !important;
        }}
        
        div[data-testid="stExpander"] {{
            background-color: {tema['card_bg']} !important;
            border: 1px solid {tema['accent']} !important;
            border-radius: 5px !important;
        }}
        
        div[data-testid="stExpander"] summary {{
            color: {tema['text']} !important;
            font-weight: {"bold" if tema_nome == "Santos" else "600"} !important;
        }}
        
        .stAlert {{
            background-color: {tema['card_bg']} !important;
            color: {tema['text']} !important;
            border: 1px solid {tema['accent']} !important;
        }}
        
        .stSuccess, .stInfo, .stWarning, .stError {{
            color: {tema['text']} !important;
        }}
        
        /* Melhorar contraste dos textos em alertas */
        .stSuccess > div, .stInfo > div, .stWarning > div, .stError > div {{
            color: {tema['text']} !important;
            font-weight: 500 !important;
        }}
        
        /* Sidebar */
        section[data-testid="stSidebar"] {{
            background-color: {tema['secondary']} !important;
        }}
        
        section[data-testid="stSidebar"] * {{
            color: {tema['text']} !important;
        }}
        </style>
    """, unsafe_allow_html=True)

# Obter informações de acesso
last_access = get_last_access()
current_access = save_current_access()

# Seleção de tema
st.sidebar.title("🎨 Configurações")
tema_selecionado = st.sidebar.selectbox(
    "Escolha o Tema:",
    list(TEMAS.keys()),
    index=0
)

# Aplicar tema
aplicar_tema(tema_selecionado)

# Título
st.title("🔍 Expedição ARM Recap")

# Relógio digital moderno
now = datetime.now()
data_abreviada = now.strftime("%d/%b/%y").upper()
hora_atual = now.strftime("%H:%M:%S")

col_clock1, col_clock2 = st.columns(2)
with col_clock1:
    st.markdown(f"""
        <div class="digital-clock">
            📅 {data_abreviada} | ⏰ {hora_atual}
        </div>
    """, unsafe_allow_html=True)

with col_clock2:
    if last_access:
        st.markdown(f"""
            <div class="digital-clock">
                🕐 ÚLTIMO: {last_access}
            </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
            <div class="digital-clock">
                🕐 PRIMEIRO ACESSO
            </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# Carregar arquivo automaticamente
ARQUIVO_EXCEL = "CX PTMs.xlsx"

try:
    # Verificar se o arquivo existe
    if os.path.exists(ARQUIVO_EXCEL):
        # Carregar o arquivo Excel automaticamente
        df = pd.read_excel(ARQUIVO_EXCEL)
        
        st.success(f"✅ Arquivo '{ARQUIVO_EXCEL}' carregado automaticamente!")
        
        # Mostrar informações do arquivo
        with st.expander("ℹ️ Informações do Arquivo"):
            st.write(f"**Total de linhas:** {len(df)}")
            st.write(f"**Total de colunas:** {len(df.columns)}")
            st.write(f"**Colunas disponíveis:** {' - '.join(df.columns.tolist())}")
        
        st.markdown("---")
        
        # Campo de busca com múltiplas opções
        st.subheader("🔎 Buscar Números")
        
        # Inicializar session state para os campos de busca
        if 'busca1' not in st.session_state:
            st.session_state.busca1 = ""
        if 'busca2' not in st.session_state:
            st.session_state.busca2 = ""
        if 'busca3' not in st.session_state:
            st.session_state.busca3 = ""
        if 'busca4' not in st.session_state:
            st.session_state.busca4 = ""
        
        # Criar 4 campos de busca
        col1, col2 = st.columns(2)
        
        with col1:
            numero_busca1 = st.text_input(
                "Busca 1:",
                placeholder="Ex: 8455",
                key="input_busca1",
                value=st.session_state.busca1
            )
            
            numero_busca2 = st.text_input(
                "Busca 2:",
                placeholder="Ex: 6140",
                key="input_busca2",
                value=st.session_state.busca2
            )
        
        with col2:
            numero_busca3 = st.text_input(
                "Busca 3:",
                placeholder="Ex: 9122",
                key="input_busca3",
                value=st.session_state.busca3
            )
            
            numero_busca4 = st.text_input(
                "Busca 4:",
                placeholder="Ex: 5621",
                key="input_busca4",
                value=st.session_state.busca4
            )
        
        # Botões de ação
        col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
        
        with col_btn1:
            buscar = st.button("🔍 Buscar Todos", type="primary", use_container_width=True)
        
        with col_btn2:
            limpar = st.button("🗑️ Limpar", use_container_width=True)
        
        # Função para limpar os campos
        if limpar:
            st.session_state.busca1 = ""
            st.session_state.busca2 = ""
            st.session_state.busca3 = ""
            st.session_state.busca4 = ""
            st.rerun()
        
        # Atualizar session state
        st.session_state.busca1 = numero_busca1
        st.session_state.busca2 = numero_busca2
        st.session_state.busca3 = numero_busca3
        st.session_state.busca4 = numero_busca4
        
        # Função para buscar número
        def buscar_numero(numero):
            resultados = []
            for coluna in df.columns:
                coluna_numerica = pd.to_numeric(df[coluna], errors='coerce')
                if numero in coluna_numerica.values:
                    linhas = df[coluna_numerica == numero].index.tolist()
                    for linha in linhas:
                        resultados.append({
                            'Número': numero,
                            'Posição': coluna,
                            'Linha': linha + 2
                        })
            return resultados
        
        # Realizar busca
        if buscar:
            numeros_busca = [numero_busca1, numero_busca2, numero_busca3, numero_busca4]
            numeros_busca = [n.strip() for n in numeros_busca if n.strip()]
            
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
                
                # Tabela consolidada
                if todos_resultados:
                    st.subheader("📋 Tabela Consolidada de Resultados")
                    df_resultados = pd.DataFrame(todos_resultados)
                    st.dataframe(
                        df_resultados,
                        use_container_width=True,
                        hide_index=True
                    )
            else:
                st.warning("⚠️ Por favor digite pelo menos um número para buscar!")
        
        # Visualizar dados completos (opcional)
        st.markdown("---")
        with st.expander("📊 Visualizar Dados Completos"):
            st.dataframe(df, use_container_width=True)
    
    else:
        # Se o arquivo não existir, mostrar opção de upload
        st.warning(f"⚠️ Arquivo '{ARQUIVO_EXCEL}' não encontrado no repositório.")
        st.info("📁 Por favor, faça o upload do arquivo Excel:")
        
        uploaded_file = st.file_uploader("Carregar arquivo Excel", type=['xlsx'])
        
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            st.success("✅ Arquivo carregado com sucesso!")
            st.info("💡 **Dica:** Para carregar automaticamente, adicione o arquivo 'CX PTMs.xlsx' na raiz do seu repositório GitHub.")
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
            
except Exception as e:
    st.error(f"❌ Erro ao carregar o arquivo: {str(e)}")
    st.info("💡 Certifique-se de que o arquivo 'CX PTMs.xlsx' está na raiz do repositório GitHub.")

# Footer discreto
st.markdown("""
    <div class="footer">
        Desenvolvido por Djalma A Barbosa 2026. Direitos reservados.
    </div>
""", unsafe_allow_html=True)