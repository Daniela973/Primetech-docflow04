import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
from docx import Document

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Prime Tech | DocFlow", 
    page_icon="📄", 
    layout="wide"
)

# --- ESTILIZAÇÃO VISUAL (PADRÃO PRIME TECH) ---
st.markdown("""
<style>
    .stApp {
        background-color: #0b0c10;
        color: #ffffff;
    }
    h1, h2, h3, h4 {
        color: #00ffff !important;
    }
    p, label, span, div, .stMarkdown {
        color: #ffffff !important;
    }
    [data-testid="stSidebar"] {
        background-color: #12141a;
        border-right: 1px solid #1f2833;
    }
    div.stButton > button {
        background: linear-gradient(135deg, #00d2ff 0%, #3a7bd5 100%);
        color: #0b0c10;
        border-radius: 8px;
        border: none;
        font-weight: bold;
        box-shadow: 0 4px 10px rgba(0, 210, 255, 0.3);
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #3a7bd5 0%, #00d2ff 100%);
        color: #ffffff;
    }
    input, textarea, select {
        background-color: #1f2833 !important;
        color: #ffffff !important;
        border: 1px solid #2c353d !important;
    }
    .logo-container {
        text-align: center;
        padding: 10px;
        background: #0b0c10;
        border-radius: 10px;
        border: 1px solid #1f2833;
        margin-bottom: 15px;
    }
    .logo-titulo {
        font-size: 20px;
        font-weight: 900;
        color: #00ffff;
        letter-spacing: 2px;
        margin: 0;
        text-shadow: 0 0 10px rgba(0, 255, 255, 0.5);
    }
    .logo-sub {
        font-size: 10px;
        color: #ffffff;
        letter-spacing: 1px;
        margin-top: 2px;
    }
</style>
""", unsafe_allow_html=True)

# --- DIRETÓRIOS DO SISTEMA ---
PASTA_SAIDA = "saida_documentos"
if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

# --- CONTROLO DE AUTENTICAÇÃO SEGURA (LOGIN) ---
if "autenticado_p4" not in st.session_state:
    st.session_state.autenticado_p4 = False

if not st.session_state.autenticado_p4:
    st.markdown("""
        <div style="text-align: center; padding: 20px;">
            <div style="font-size: 32px; font-weight: 900; color: #00ffff; text-shadow: 0 0 15px rgba(0, 255, 255, 0.4);">PRIME TECH</div>
            <div style="font-size: 14px; color: #ffffff; letter-spacing: 2px;">DOCFLOW — ACESSO RESTRITO</div>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        senha = st.text_input("Digite a palavra-passe de acesso:", type="password", key="senha_p4")
        if st.button("Entrar no DocFlow", use_container_width=True):
            senha_correta = st.secrets.get("SENHA_ADMIN", "admin123")
            if senha == senha_correta:
                st.session_state.autenticado_p4 = True
                st.rerun()
            else:
                st.error("❌ Palavra-passe incorreta!")
    st.stop()

# --- BARRA LATERAL ---
st.sidebar.markdown("""
    <div class="logo-container">
        <div class="logo-titulo">PRIME TECH</div>
        <div style="background: linear-gradient(90deg, transparent, #00ffff, transparent); height: 2px; margin: 5px 0;"></div>
        <div class="logo-sub">DOCFLOW AUTOMATION</div>
    </div>
""", unsafe_allow_html=True)

menu = st.sidebar.radio("Navegação", [
    "🏠 Dashboard",
    "👥 Clientes",
    "📋 Modelos de Documentos",
    "✍️ Gerar Documento (Word/PDF)",
    "📦 Geração em Lote (Excel)",
    "📊 Relatórios & Histórico",
    "⚙️ Configurações"
])

if st.sidebar.button("🚪 Terminar Sessão"):
    st.session_state.autenticado_p4 = False
    st.rerun()

# --- 1. DASHBOARD ---
if menu == "🏠 Dashboard":
    st.title("📄 Prime Tech DocFlow")
    st.markdown("Central Inteligente de Automação de Documentos e Geração de Contratos.")
    st.markdown("---")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📄 Documentos Gerados", "12")
    col2.metric("📝 Modelos Ativos", "3")
    col3.metric("👥 Clientes Registados", "8")
    col4.metric("⚡ Taxa de Sucesso", "100%")
    
    st.markdown("---")
    st.info("💡 **Dica:** Utilize o menu lateral para gerir modelos, cadastrar clientes ou gerar novos documentos de forma automatizada.")

# --- 2. CLIENTES ---
elif menu == "👥 Clientes":
    st.title("👥 Gestão de Clientes")
    st.markdown("Cadastre e consulte os dados dos clientes para preenchimento rápido de contratos.")
    
    with st.form("form_cliente"):
        col1, col2 = st.columns(2)
        with col1:
            nome = st.text_input("Nome Completo / Empresa")
            cpf = st.text_input("CPF / CNPJ")
        with col2:
            email = st.text_input("E-mail")
            telefone = st.text_input("Telefone / Telemóvel")
            
        salvar_cli = st.form_submit_button("Guardar Cliente")
        if salvar_cli:
            if nome:
                st.success(f"✔ Cliente `{nome}` registado com sucesso!")
            else:
                st.warning("⚠️ O campo Nome é obrigatório.")

# --- 3. MODELOS ---
elif menu == "📋 Modelos de Documentos":
    st.title("📋 Modelos de Documentos")
    st.markdown("Modelos padrão disponíveis para automação.")
    
    st.markdown("""
    * **Contrato de Prestação de Serviços** (Variáveis: `{cliente}`, `{cpf}`, `{servico}`, `{valor}`, `{data}`)
    * **Recibo de Pagamento** (Variáveis: `{cliente}`, `{valor}`, `{data}`)
    * **Orçamento Comercial** (Variáveis: `{cliente}`, `{servico}`, `{valor}`, `{data}`)
    """)

# --- 4. GERAR DOCUMENTO ---
elif menu == "✍️ Gerar Documento (Word/PDF)":
    st.title("✍️ Gerador Automático de Documentos")
    st.markdown("Preencha os dados abaixo para gerar o seu documento em formato Word (.docx).")
    
    with st.form("form_geracao"):
        modelo_escolhido = st.selectbox("Selecione o Modelo", ["Contrato de Prestação de Serviços", "Recibo de Pagamento", "Orçamento Comercial"])
        
        cliente = st.text_input("Nome do Cliente", "João da Silva")
        cpf = st.text_input("CPF / CNPJ", "123.456.789-00")
        servico = st.text_input("Descrição do Serviço", "Manutenção de Computador e Suporte Técnico")
        valor = st.text_input("Valor (R$ / €)", "350,00")
        data_doc = st.text_input("Data de Emissão", datetime.now().strftime("%d/%m/%Y"))
        
        gerar = st.form_submit_button("Gerar Documento Word")
        
        if gerar:
            try:
                # Criar documento Word via python-docx
                doc = Document()
                doc.add_heading(f"PRIME TECH - {modelo_escolhido.upper()}", 0)
                
                doc.add_paragraph(f"Emitido em: {data_doc}\n")
                doc.add_heading("1. Dados do Contratante", level=1)
                doc.add_paragraph(f"Nome / Razão Social: {cliente}")
                doc.add_paragraph(f"CPF / CNPJ: {cpf}")
                
                doc.add_heading("2. Descrição do Serviço", level=1)
                doc.add_paragraph(f"Serviço Prestado: {servico}")
                doc.add_paragraph(f"Valor Total: {valor}")
                
                doc.add_heading("3. Termos e Condições", level=1)
                doc.add_paragraph("O presente documento rege-se pelos termos acordados entre as partes, assegurando a qualidade e o cumprimento dos prazos estabelecidos pela Prime Tech Solutions.")
                
                doc.add_paragraph("\n\n___________________________________\nAssinatura do Responsável")
                
                # Guardar ficheiro
                nome_arquivo_saida = f"Documento_{cliente.replace(' ', '_')}.docx"
                caminho_completo = os.path.join(PASTA_SAIDA, nome_arquivo_saida)
                doc.save(caminho_completo)
                
                st.success(f"✔ Documento gerado com sucesso!")
                
                # Botão de download
                with open(caminho_completo, "rb") as f:
                    st.download_button(
                        label="📥 Descarregar Documento Word (.docx)",
                        data=f,
                        file_name=nome_arquivo_saida,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"❌ Erro ao gerar documento: {e}")

# --- 5. GERAÇÃO EM LOTE ---
elif menu == "📦 Geração em Lote (Excel)":
    st.title("📦 Geração em Lote a partir de Excel")
    st.info("Carregue uma planilha com múltiplos clientes para gerar dezenas de contratos ou recibos automaticamente de uma só vez.")
    st.file_uploader("Carregar Ficheiro de Clientes (.xlsx)", type=["xlsx", "csv"])

# --- 6. RELATÓRIOS ---
elif menu == "📊 Relatórios & Histórico":
    st.title("📊 Relatórios e Estatísticas")
    st.info("Registo de todos os documentos gerados pelo sistema.")

# --- 7. CONFIGURAÇÕES ---
elif menu == "⚙️ Configurações":
    st.title("⚙️ Configurações da Empresa")
    st.info("Personalize dados da empresa, logótipos e pastas padrão de salvamento.")
