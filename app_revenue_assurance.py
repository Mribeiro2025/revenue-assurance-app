import os
import json
import datetime
import pandas as pd
import openpyxl
import streamlit as st

# 1. Configuração Inicial da Página
st.set_page_config(
    page_title="Grupo Arbaitman | Revenue Assurance",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_DASHBOARD = "Dashboard_Revenue_Assurance_Consolidado.xlsx"
ARQUIVO_USUARIOS = "usuarios_autorizados.json"

# 2. Gerenciador de Usuários e Persistência de Cadastro
USUARIOS_PADRAO = {
    "mribeiro": {"senha": "123", "nome": "Marcos Ribeiro", "perfil": "Compliance", "status": "APROVADO", "data_solicitacao": "2026-08-31"},
    "compliance1": {"senha": "123", "nome": "Compliance - Auditoria 01", "perfil": "Compliance", "status": "APROVADO", "data_solicitacao": "2026-08-31"},
    "operacao": {"senha": "123", "nome": "Equipe Operacional", "perfil": "Operacao", "status": "APROVADO", "data_solicitacao": "2026-08-31"},
    "backoffice": {"senha": "123", "nome": "Atendimento Backoffice", "perfil": "Operacao", "status": "APROVADO", "data_solicitacao": "2026-08-31"}
}

def carregar_usuarios():
    if not os.path.exists(ARQUIVO_USUARIOS):
        with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
            json.dump(USUARIOS_PADRAO, f, ensure_ascii=False, indent=4)
        return USUARIOS_PADRAO
    try:
        with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return USUARIOS_PADRAO

def salvar_usuarios(dict_users):
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(dict_users, f, ensure_ascii=False, indent=4)

usuarios_db = carregar_usuarios()

# 3. Estilização CSS Corporativa
st.markdown("""
    <style>
        .stApp { background-color: #f8f9fa; }
        
        /* Botões Principais */
        div.stButton > button[kind="primary"], div.stButton > button {
            background-color: #002060 !important;
            color: #ffffff !important;
            border-radius: 6px !important;
            border: none !important;
            font-weight: 600 !important;
        }
        div.stButton > button:hover {
            background-color: #001040 !important;
            color: #ffffff !important;
        }
        
        /* Cabeçalho */
        .header-box {
            background: linear-gradient(135deg, #002060 0%, #003366 100%);
            padding: 15px 25px;
            border-radius: 10px;
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .header-box h1 { color: #ffffff !important; margin: 0; font-size: 24px; font-weight: 700; }
        .header-box p { color: #d0e0ff !important; margin-top: 4px; font-size: 13px; margin-bottom: 0; }
        
        /* Card de Login */
        .login-card {
            background-color: #ffffff;
            padding: 30px 35px;
            border-radius: 12px;
            border-top: 6px solid #002060;
            box-shadow: 0 8px 24px rgba(0,0,0,0.08);
            margin-top: 20px;
        }
        
        /* Métricas KPI */
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 14px;
            border-left: 5px solid #002060;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        }

        /* TAGS DO MULTISELECT EM CINZA CORPORATIVO */
        span[data-baseweb="tag"] {
            background-color: #6c757d !important;
            color: #ffffff !important;
            border-radius: 4px !important;
        }
        
        /* Marca Textual Corporativa */
        .brand-header {
            font-size: 22px;
            font-weight: 800;
            color: #002060;
            letter-spacing: 1px;
            text-align: center;
            margin-bottom: 5px;
        }
    </style>
""", unsafe_allow_html=True)

def renderizar_marca():
    st.markdown('<div class="brand-header">GRUPO ARBAITMAN</div>', unsafe_allow_html=True)

# 4. Autenticação e Sessão
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario_atual"] = None
    st.session_state["perfil_atual"] = None
    st.session_state["login_user_id"] = None

if not st.session_state["autenticado"]:
    col_left, col_center, col_right = st.columns([1, 1.2, 1])
    
    with col_center:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        renderizar_marca()
        st.markdown("<p style='text-align: center; color: #6c757d; font-size: 13px; margin-bottom: 20px;'>Maringá Turismo | Portal de Revenue Assurance</p>", unsafe_allow_html=True)
        
        aba_login, aba_redefinir, aba_solicitar = st.tabs(["🔐 Entrar", "🔑 Esqueci a Senha", "📝 Solicitar Acesso"])
        
        with aba_login:
            user_input = st.text_input("Usuário de Acesso:", key="login_user").strip().lower()
            pass_input = st.text_input("Senha:", type="password", key="login_pass").strip()
            btn_entrar = st.button("🚀 Acessar Sistema", use_container_width=True, type="primary")
            
            if btn_entrar:
                if user_input in usuarios_db:
                    dados_u = usuarios_db[user_input]
                    if dados_u["senha"] == pass_input:
                        if dados_u.get("status") == "APROVADO":
                            st.session_state["autenticado"] = True
                            st.session_state["usuario_atual"] = dados_u["nome"]
                            st.session_state["perfil_atual"] = dados_u["perfil"]
                            st.session_state["login_user_id"] = user_input
                            st.rerun()
                        elif dados_u.get("status") == "PENDENTE":
                            st.warning("⏳ Sua solicitação de acesso está **Pendente de Aprovação** pelo Compliance.")
                        else:
                            st.error("❌ Acesso não autorizado.")
                    else:
                        st.error("❌ Senha incorreta.")
                else:
                    st.error("❌ Usuário não cadastrado.")

        with aba_redefinir:
            st.caption("Redefinição direta de senha de acesso.")
            with st.form("form_redefinir_senha_login"):
                user_reset = st.text_input("Informe seu Usuário de Acesso:").strip().lower()
                nova_senha_login = st.text_input("Nova Senha:", type="password")
                confirma_senha_login = st.text_input("Confirme a Nova Senha:", type="password")
                btn_redefinir_senha = st.form_submit_button("🔄 Redefinir Senha", use_container_width=True)

                if btn_redefinir_senha:
                    if not user_reset:
                        st.error("⚠️ Digite o usuário de acesso.")
                    elif user_reset not in usuarios_db:
                        st.error("❌ Usuário não encontrado no sistema.")
                    elif not nova_senha_login.strip():
                        st.error("⚠️ Digite a nova senha.")
                    elif nova_senha_login != confirma_senha_login:
                        st.error("❌ As senhas não coincidem.")
                    else:
                        usuarios_db[user_reset]["senha"] = nova_senha_login.strip()
                        salvar_usuarios(usuarios_db)
                        st.success("✅ Senha redefinida com sucesso! Clique na aba '🔐 Entrar' para acessar.")

        with aba_solicitar:
            st.caption("Solicitação formal de acesso para novos colaboradores.")
            novo_nome = st.text_input("Nome Completo:")
            novo_user = st.text_input("Usuário Desejado (ex: nome.sobrenome):").strip().lower()
            nova_senha = st.text_input("Crie uma Senha:", type="password")
            novo_perfil = st.selectbox("Perfil Solicitado:", options=["Operacao", "Compliance"])
            
            st.markdown("""
                <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 10px; border-radius: 6px; font-size: 11px; color: #495057; max-height: 90px; overflow-y: scroll; margin-bottom: 10px;">
                    <b>TERMO DE CONFIDENCIALIDADE - GRUPO ARBAITMAN</b><br>
                    O usuário declara estar ciente do caráter confidencial das informações de Revenue Assurance. Qualquer alteração ou exportação de dados é monitorada na trilha de auditoria.
                </div>
            """, unsafe_allow_html=True)
            
            termo_aceito = st.checkbox("Li e aceito os termos de sigilo")
            btn_cadastrar = st.button("📩 Enviar Solicitação", use_container_width=True)
            
            if btn_cadastrar:
                if not novo_nome or not novo_user or not nova_senha:
                    st.error("⚠️ Preencha todos os campos do formulário.")
                elif not termo_aceito:
                    st.error("⚠️ Aceite o termo de sigilo para continuar.")
                elif novo_user in usuarios_db:
                    st.error("❌ Usuário já existente.")
                else:
                    usuarios_db[novo_user] = {
                        "senha": nova_senha,
                        "nome": novo_nome,
                        "perfil": novo_perfil,
                        "status": "PENDENTE",
                        "data_solicitacao": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    salvar_usuarios(usuarios_db)
                    st.success("✅ Solicitação enviada! Aguarde a liberação pelo Compliance.")

        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# 5. Leitura de Bases e Tratamento de Colunas
def padronizar_e_deduplicar_colunas(df, origem=""):
    if df is None or df.empty: return pd.DataFrame()
    df = df.loc[:, ~df.columns.duplicated()].copy()
    
    renomear = {}
    for col in df.columns:
        c_str = str(col).strip().lower()
        if c_str in ["gerente", "gerentes", "área resp. operação", "area resp. operacao", "area_resp_operacao"]:
            renomear[col] = "Área Resp. Operação"
        elif c_str == "setor":
            renomear[col] = "Setor"
        elif c_str in ["consultor_lemon", "emissor", "consultor"]:
            renomear[col] = "Consultor_Lemon"
            
    df_out = df.rename(columns=renomear).copy()
    df_out = df_out.loc[:, ~df_out.columns.duplicated()].copy()
    
    if "Área Resp. Operação" not in df_out.columns: df_out["Área Resp. Operação"] = "Não Atribuído"
    if "Setor" not in df_out.columns: df_out["Setor"] = "Geral"
    if "Consultor_Lemon" not in df_out.columns: df_out["Consultor_Lemon"] = "-"
    if "Localizador_Sistema" not in df_out.columns: df_out["Localizador_Sistema"] = "-"
    if "Status_Geral" not in df_out.columns: df_out["Status_Geral"] = "Pendente de Lançamento"
    df_out["Origem_Aba"] = origem
    return df_out

@st.cache_data(ttl=2)
def carregar_bases():
    if not os.path.exists(ARQUIVO_DASHBOARD):
        return None, None, None, None, None, "ARQUIVO_NAO_ENCONTRADO"
    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD)
        aba_names = xls.sheet_names
        
        df_m = pd.read_excel(xls, sheet_name="99_Base_Divergencias_Geral") if "99_Base_Divergencias_Geral" in aba_names else pd.DataFrame()
        df_d = pd.read_excel(xls, sheet_name="98_OK_Divergencia_Operacao") if "98_OK_Divergencia_Operacao" in aba_names else pd.DataFrame()
        df_s = pd.read_excel(xls, sheet_name="98_OK_Sem_Divergencia_Concil") if "98_OK_Sem_Divergencia_Concil" in aba_names else pd.DataFrame()
        df_b = pd.read_excel(xls, sheet_name="99_Suporte backoffice") if "99_Suporte backoffice" in aba_names else pd.DataFrame()
        
        df_master = padronizar_e_deduplicar_colunas(df_m, "99_Geral")
        df_div_op = padronizar_e_deduplicar_colunas(df_d, "98_Divergencia_Operacao")
        df_sem_div = padronizar_e_deduplicar_colunas(df_s, "98_Sem_Divergencia")
        df_backoffice = padronizar_e_deduplicar_colunas(df_b, "99_Backoffice")
        
        if "00_Log_Auditoria" in aba_names:
            df_log = pd.read_excel(xls, sheet_name="00_Log_Auditoria")
            df_log = df_log.loc[:, ~df_log.columns.duplicated()].copy()
        else:
            df_log = pd.DataFrame(columns=["Data_Hora", "Bilhete", "Usuario_Acao", "Status_Anterior", "Novo_Status", "Area_Anterior", "Nova_Area", "Observacao", "Tipo_Interacao"])
            
        return df_master, df_div_op, df_sem_div, df_backoffice, df_log, "OK"
    except PermissionError:
        return None, None, None, None, None, "ARQUIVO_BLOQUEADO"
    except Exception as e:
        return None, None, None, None, None, str(e)

df_master, df_div_op, df_sem_div, df_backoffice, df_log_master, status_carga = carregar_bases()

if status_carga == "ARQUIVO_BLOQUEADO":
    st.error(f"⚠️ O arquivo **'{ARQUIVO_DASHBOARD}'** está em uso. Feche a planilha para atualizar.")
    if st.button("🔄 Recarregar Dados"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

if (df_master is None or df_master.empty) and (df_div_op is None or df_div_op.empty):
    st.error(f"⚠️ Base consolidada de divergências não encontrada.")
    st.stop()

COL_GERENTE = "Área Resp. Operação"
COL_SETOR = "Setor"
COL_EMISSOR = "Consultor_Lemon"

df_list = [d for d in [df_master, df_div_op] if not d.empty]
df_acao_total = pd.concat(df_list, ignore_index=True, sort=False) if len(df_list) > 0 else pd.DataFrame()

if "Bilhetes" in df_acao_total.columns:
    df_acao_total["Bilhetes_Str"] = df_acao_total["Bilhetes"].astype(str).str.strip()
    df_acao_total = df_acao_total.drop_duplicates(subset=["Bilhetes_Str"], keep="first")
else:
    df_acao_total["Bilhetes_Str"] = ""

usuario_log_formatado = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"

# 6. Cabeçalho Principal do Dashboard
col_hdr1, col_hdr2 = st.columns([3, 1])
with col_hdr1:
    st.markdown(f"""
        <div class="header-box">
            <div>
                <h1>GRUPO ARBAITMAN | Revenue Assurance</h1>
                <p>Maringá Turismo — Conectado como: <b>{st.session_state['usuario_atual']}</b> | Perfil: <b>{st.session_state['perfil_atual']}</b></p>
            </div>
        </div>
    """, unsafe_allow_html=True)

with col_hdr2:
    if st.button("🔄 Recarregar / Atualizar Dados", use_container_width=True, type="primary"):
        st.cache_data.clear()
        st.rerun()

if "msg_sucesso" in st.session_state:
    st.success(st.session_state["msg_sucesso"])
    del st.session_state["msg_sucesso"]

# Sidebar / Filtros Operacionais
renderizar_marca()
st.sidebar.markdown(f"### 👤 {st.session_state['usuario_atual']}")

col_btn_sair, col_btn_senha = st.sidebar.columns(2)
with col_btn_sair:
    if st.button("🔒 Sair"):
        st.session_state["autenticado"] = False
        st.session_state["usuario_atual"] = None
        st.session_state["perfil_atual"] = None
        st.rerun()

with col_btn_senha:
    btn_mudar_senha = st.button("🔑 Senha")

if btn_mudar_senha or st.session_state.get("abrir_modal_senha", False):
    st.session_state["abrir_modal_senha"] = True
    with st.sidebar.expander("🔑 Alterar Minha Senha", expanded=True):
        u_id_atual = st.session_state["login_user_id"]
        with st.form("form_senha_sidebar"):
            senha_atual_input = st.text_input("Senha Atual:", type="password")
            nova_senha_input = st.text_input("Nova Senha:", type="password")
            confirma_senha_input = st.text_input("Confirmar Nova Senha:", type="password")
            btn_upd_pwd = st.form_submit_button("💾 Atualizar Senha", use_container_width=True)
            
            if btn_upd_pwd:
                if usuarios_db[u_id_atual]["senha"] != senha_atual_input:
                    st.error("❌ Senha atual incorreta.")
                elif not nova_senha_input.strip():
                    st.error("⚠️ Digite uma nova senha.")
                elif nova_senha_input != confirma_senha_input:
                    st.error("❌ As senhas não coincidem.")
                else:
                    usuarios_db[u_id_atual]["senha"] = nova_senha_input.strip()
                    salvar_usuarios(usuarios_db)
                    st.session_state["msg_sucesso"] = "✅ Senha alterada com sucesso!"
                    st.session_state["abrir_modal_senha"] = False
                    st.rerun()

st.sidebar.markdown("---")
st.sidebar.title("🔍 Filtros Operacionais")

opcoes_setor = sorted(list(df_acao_total[COL_SETOR].dropna().astype(str).unique()))
setor_sel = st.sidebar.multiselect("Setor Responsável:", options=opcoes_setor, default=opcoes_setor)

df_f_1 = df_acao_total[df_acao_total[COL_SETOR].astype(str).isin(setor_sel)]

opcoes_gerente = sorted(list(df_f_1[COL_GERENTE].dropna().astype(str).unique()))
gerente_sel = st.sidebar.multiselect("Gerente (Área Resp. Operação):", options=opcoes_gerente, default=opcoes_gerente)

df_f_2 = df_f_1[df_f_1[COL_GERENTE].astype(str).isin(gerente_sel)]

opcoes_emissor = sorted(list(df_f_2[COL_EMISSOR].dropna().astype(str).unique()))
emissor_sel = st.sidebar.multiselect("Emissor / Consultor (OBT):", options=opcoes_emissor, default=opcoes_emissor)

df_f_3 = df_f_2[df_f_2[COL_EMISSOR].astype(str).isin(emissor_sel)]

tipos_emissao = sorted(list(df_f_3["Tipo_Emissao_Lemon"].dropna().astype(str).unique())) if "Tipo_Emissao_Lemon" in df_f_3.columns else []
emissao_sel = st.sidebar.multiselect("Tipo de Emissão (OBT):", options=tipos_emissao, default=tipos_emissao)

df_f_4 = df_f_3[df_f_3["Tipo_Emissao_Lemon"].astype(str).isin(emissao_sel)] if "Tipo_Emissao_Lemon" in df_f_3.columns else df_f_3

cias = sorted(list(df_f_4["CIA"].dropna().astype(str).unique())) if "CIA" in df_f_4.columns else []
cia_sel = st.sidebar.multiselect("Companhia Aérea:", options=cias, default=cias)

def aplicar_filtros_globais(df):
    if df is None or df.empty: return df
    m = pd.Series(True, index=df.index)
    if COL_SETOR in df.columns and setor_sel: m = m & df[COL_SETOR].astype(str).isin(setor_sel)
    if COL_GERENTE in df.columns and gerente_sel: m = m & df[COL_GERENTE].astype(str).isin(gerente_sel)
    if COL_EMISSOR in df.columns and emissor_sel: m = m & df[COL_EMISSOR].astype(str).isin(emissor_sel)
    if "Tipo_Emissao_Lemon" in df.columns and emissao_sel: m = m & df["Tipo_Emissao_Lemon"].astype(str).isin(emissao_sel)
    if "CIA" in df.columns and cia_sel: m = m & df["CIA"].astype(str).isin(cia_sel)
    return df[m].copy()

df_master_filtrado = aplicar_filtros_globais(df_master)
df_div_op_filtrado = aplicar_filtros_globais(df_div_op)
df_sem_div_filtrado = aplicar_filtros_globais(df_sem_div)
df_acao_filtrado = aplicar_filtros_globais(df_acao_total)

df_backoffice_dinamico = pd.concat([
    df_backoffice,
    df_master[df_master[COL_GERENTE].astype(str).str.lower() == "suporte backoffice"],
    df_div_op[df_div_op[COL_GERENTE].astype(str).str.lower() == "suporte backoffice"]
], ignore_index=True).drop_duplicates(subset=["Bilhetes"], keep="last") if "Bilhetes" in df_master.columns else df_backoffice

df_backoffice_filtrado = aplicar_filtros_globais(df_backoffice_dinamico)

abas_nomes = [
    "📊 Dashboard & KPIs",
    "🎯 Tratativa Operacional (Geral)",
    "⚠️ Divergência Operação (CIAs/HOT + Tratativa Direta)",
    "✅ Sem Divergência (Conciliação)",
    "🎧 Suporte Backoffice",
    "⚖️ Réplica da Auditoria",
    "📋 Visão Geral da Base Total"
]

if st.session_state["perfil_atual"] == "Compliance":
    abas_nomes.insert(6, "📜 Trilha de Auditoria (Exclusivo Compliance)")
    abas_nomes.insert(7, "⚙️ Gestão de Acessos & Aprovações")

abas_objetos = st.tabs(abas_nomes)

gerentes_base_unicos = sorted([g for g in df_acao_total[COL_GERENTE].dropna().astype(str).unique() if g not in ["Suporte backoffice", "Não Atribuído", "-"]])

# ABA 0: DASHBOARD
with abas_objetos[0]:
    st.subheader("📊 Painel Executivo e Métricas de Controladoria")
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Bilhetes/LOCs em Ação", f"{len(df_acao_filtrado):,}")
    k2.metric("Tarifa Pendente (R$)", f"R$ {df_acao_filtrado['A credito'].sum():,.2f}" if "A credito" in df_acao_filtrado.columns else "R$ 0.00")
    k3.metric("Taxas Pendentes (R$)", f"R$ {df_acao_filtrado['Taxa'].sum():,.2f}" if "Taxa" in df_acao_filtrado.columns else "R$ 0.00")
    k4.metric("Receita em Risco (R$)", f"R$ {df_acao_filtrado['Incentivo'].sum():,.2f}" if "Incentivo" in df_acao_filtrado.columns else "R$ 0.00")
    
    st.markdown("---")
    
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.markdown("##### 📌 Distribuição de Volumetria por Status Geral")
        if not df_acao_filtrado.empty and "Status_Geral" in df_acao_filtrado.columns:
            chart_status = df_acao_filtrado["Status_Geral"].value_counts().reset_index()
            chart_status.columns = ["Status", "Quantidade"]
            st.bar_chart(data=chart_status, x="Status", y="Quantidade", color="#002060")
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    with col_chart2:
        st.markdown("##### 👤 Top Gerentes Responsáveis por Volume de Pendências")
        if not df_acao_filtrado.empty and COL_GERENTE in df_acao_filtrado.columns:
            chart_gerente = df_acao_filtrado[COL_GERENTE].value_counts().head(8).reset_index()
            chart_gerente.columns = ["Gerente", "Quantidade"]
            st.bar_chart(data=chart_gerente, x="Gerente", y="Quantidade", color="#003366")
        else:
            st.info("Sem dados para exibir o gráfico.")
            
    st.markdown("---")
    
    col_chart3, col_chart4 = st.columns(2)
    with col_chart3:
        st.markdown("##### ✈️ Concentração por Companhia Aérea (CIA)")
        if not df_acao_filtrado.empty and "CIA" in df_acao_filtrado.columns:
            chart_cia = df_acao_filtrado["CIA"].value_counts().reset_index()
            chart_cia.columns = ["Companhia Aérea", "Quantidade"]
            st.dataframe(chart_cia, use_container_width=True, hide_index=True)
            
    with col_chart4:
        st.markdown("##### 🏢 Volume de Divergências por Setor")
        if not df_acao_filtrado.empty and COL_SETOR in df_acao_filtrado.columns:
            chart_setor = df_acao_filtrado[COL_SETOR].value_counts().reset_index()
            chart_setor.columns = ["Setor", "Quantidade"]
            st.dataframe(chart_setor, use_container_width=True, hide_index=True)

# ABA 1: TRATATIVA OPERACIONAL GERAL
with abas_objetos[1]:
    st.subheader("📝 Módulo de Resolução e Detalhamento Operacional")
    
    if len(df_master_filtrado) == 0:
        st.warning("Nenhum bilhete encontrado na base geral para os filtros selecionados.")
    else:
        lista_busca_geral = df_master_filtrado["Bilhetes"].astype(str).tolist()
        opcao_sel_m = st.selectbox("Procure ou Selecione o Bilhete / LOC para Tratativa:", options=lista_busca_geral, key="sb_geral")
        
        row_m = df_master_filtrado[df_master_filtrado["Bilhetes"].astype(str) == opcao_sel_m].iloc[0]
        
        st.markdown(f"""
            <div style="background-color: #fff3cd; border-left: 6px solid #ffc107; padding: 15px; border-radius: 8px; margin: 15px 0;">
                <h4 style="margin: 0; color: #856404;">⚠️ Diagnóstico: Bilhete/LOC {row_m['Bilhetes']}</h4>
                <p style="margin: 5px 0 0 0; color: #856404;">
                    <b>Área / Gerente Resp.:</b> {row_m.get(COL_GERENTE, '-')} | 
                    <b>CIA:</b> {row_m.get('CIA', '-')} | 
                    <b>Status Geral:</b> {row_m.get('Status_Geral', 'Pendente')}
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("##### 📋 Detalhes do Bilhete e Valores de Emissão")
        cols_det = ["Ponto de venda", "Área Resp. Operação", "CIA", "Bilhetes", "Localizador_Sistema", "Status_Sistema", "Data Emissão", "Pagto", "A vista", "A credito", "Taxa", "Comissão", "Taxa DU", "Desc.", "Incentivo", "VL. Líquido", "Status_Geral"]
        cols_presentes = [c for c in cols_det if c in row_m.index]
        st.dataframe(pd.DataFrame([row_m[cols_presentes]]), use_container_width=True, hide_index=True)
        st.markdown("---")
        
        lista_gerentes_combo = sorted(list(set(gerentes_base_unicos + ["Keli Santi", "Guilherme Silva", "Fabiano Souza", "Ivanete Bertasol", "Jaime Schnaider"]))) + ["Outro Gerente..."]
        gerente_atual_row = str(row_m.get(COL_GERENTE, ""))
        idx_g = lista_gerentes_combo.index(gerente_atual_row) if gerente_atual_row in lista_gerentes_combo else (len(lista_gerentes_combo) - 1)

        with st.form("form_tratativa_geral"):
            col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1.5, 1.2])
            
            with col_a:
                novo_status = st.selectbox("Status da Tratativa:", options=["Já Lançado no ERP", "Pendente de Lançamento", "Aguardando TI", "Cancelado / Devolvido"])
            
            with col_b:
                areas_opcoes = ["Operação", "Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private"]
                area_atual = str(row_m.get(COL_GERENTE, "Operação"))
                index_area = areas_opcoes.index(area_atual) if area_atual in areas_opcoes else 0
                nova_area = st.selectbox("Área Responsável (Reatribuir):", options=areas_opcoes, index=index_area)
            
            with col_c:
                if nova_area == "Operação":
                    gerente_indicado_sel = st.selectbox("Gerente Responsável:", options=lista_gerentes_combo, index=idx_g)
                    novo_gerente_texto = st.text_input("Escreva o Nome do Novo Gerente:", placeholder="Digite o nome se selecionou 'Outro Gerente...'")
                else:
                    st.text_input("Gerente Responsável:", value=f"N/A ({nova_area})", disabled=True)
                    gerente_indicado_sel = nova_area
                    novo_gerente_texto = ""

            with col_d:
                num_chamado = st.text_input("Nº do Chamado / Ticket (Obrigatório se Suporte Backoffice):", placeholder="Ex: INC-98472")
                
            obs_detalhe = st.text_area("Observações e Detalhes da Solução:", value="", placeholder="Digite aqui as observações desta tratativa...")
            btn_salvar_g = st.form_submit_button("💾 Salvar Tratativa Operacional")
            
            if btn_salvar_g:
                if nova_area == "Operação":
                    if gerente_indicado_sel == "Outro Gerente...":
                        gerente_final = novo_gerente_texto.strip()
                    else:
                        gerente_final = gerente_indicado_sel
                else:
                    gerente_final = nova_area

                if nova_area == "Suporte backoffice" and not num_chamado.strip():
                    st.error("⚠️ Para reatribuir ao **Suporte backoffice**, é OBRIGATÓRIO informar o Número do Chamado!")
                elif nova_area == "Operação" and not gerente_final:
                    st.error("⚠️ Por favor, informe o Nome do Gerente no campo de texto abaixo do seletor!")
                else:
                    idx = df_master[df_master["Bilhetes"].astype(str) == opcao_sel_m].index
                    
                    novo_log = pd.DataFrame([{
                        "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Bilhete": opcao_sel_m,
                        "Usuario_Acao": usuario_log_formatado,
                        "Status_Anterior": row_m.get("Status_Geral", "Pendente"),
                        "Novo_Status": novo_status,
                        "Area_Anterior": row_m.get(COL_GERENTE, "Operação"),
                        "Nova_Area": gerente_final,
                        "Observacao": obs_detalhe,
                        "Tipo_Interacao": "Tratativa Geral"
                    }])

                    if novo_status == "Já Lançado no ERP":
                        row_upd = row_m.copy()
                        row_upd["Status_Geral"] = "Já Lançado no ERP"
                        row_upd[COL_GERENTE] = gerente_final
                        row_upd["Obs. Operação"] = f"[Chamado: {num_chamado}] {obs_detalhe}" if num_chamado else obs_detalhe
                        row_upd["Status_Divergencia"] = "Valores Corretos"
                        
                        df_master = df_master.drop(idx)
                        df_sem_div = pd.concat([df_sem_div, pd.DataFrame([row_upd])], ignore_index=True)
                        msg_res = "transferido para a aba 'Sem Divergência'"
                    else:
                        df_master.loc[idx, "Status_Geral"] = novo_status
                        df_master.loc[idx, COL_GERENTE] = gerente_final
                        df_master.loc[idx, "Obs. Operação"] = f"[Chamado: {num_chamado}] {obs_detalhe}" if num_chamado else obs_detalhe
                        msg_res = "atualizado com sucesso"

                    df_back_atualizado = df_master[df_master[COL_GERENTE].astype(str) == "Suporte backoffice"].copy()
                    df_log_updated = pd.concat([df_log_master, novo_log], ignore_index=True)
                    
                    try:
                        with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_master.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
                            df_sem_div.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
                            if not df_back_atualizado.empty:
                                df_back_atualizado.to_excel(writer, sheet_name="99_Suporte backoffice", index=False)
                            df_log_updated.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)
                        
                        st.session_state["msg_sucesso"] = f"✅ Bilhete {opcao_sel_m} {msg_res} por {usuario_log_formatado} (Gerente Atribuído: {gerente_final})!"
                        st.cache_data.clear()
                        st.rerun()
                    except PermissionError:
                        st.error("❌ O arquivo Excel está aberto em outro programa. Feche a planilha para salvar.")

        st.markdown("---")
        st.markdown(f"### 📊 Lista Completa dos Casos Filtrados ({len(df_master_filtrado)} registros)")
        st.dataframe(df_master_filtrado, use_container_width=True, hide_index=True)

# ABA 2: DIVERGÊNCIA OPERAÇÃO
with abas_objetos[2]:
    st.subheader("⚠️ Base 98 - Divergência de Operação / CIAs Aéreas / Arquivos HOT")
    
    if len(df_div_op_filtrado) == 0:
        st.warning("Nenhuma divergência de operação encontrada.")
    else:
        lista_busca_div = df_div_op_filtrado["Bilhetes"].astype(str).tolist()
        bilhete_div_sel = st.selectbox("Selecione o Bilhete / LOC para Tratativa Direta:", options=lista_busca_div, key="sb_div_op")
        row_d = df_div_op_filtrado[df_div_op_filtrado["Bilhetes"].astype(str) == bilhete_div_sel].iloc[0]
        
        with st.form("form_tratativa_div_op"):
            c_x, c_y, c_z = st.columns(3)
            with c_x:
                novo_status_d = st.selectbox("Status da Tratativa:", options=["CIA Aérea Corrigida", "Já Lançado no ERP", "Pendente de Lançamento", "Aguardando TI", "Cancelado / Devolvido"])
            with c_y:
                cia_corrigida_d = st.text_input("CIA Aérea Corrigida:", value=str(row_d.get("CIA", "")))
            with c_z:
                num_chamado_d = st.text_input("Nº do Chamado / Ticket (se houver):")
                
            obs_d = st.text_area("Observações e Justificativas:", value="", placeholder="Digite aqui a justificativa ou correção realizada...")
            btn_salvar_d = st.form_submit_button("💾 Salvar Ação nesta Divergência")
            
            if btn_salvar_d:
                idx_d = df_div_op[df_div_op["Bilhetes"].astype(str) == bilhete_div_sel].index
                texto_obs_d = f"[Correção CIA: {cia_corrigida_d}] " + (f"[Chamado: {num_chamado_d}] " if num_chamado_d else "") + obs_d
                
                novo_log_d = pd.DataFrame([{
                    "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Bilhete": bilhete_div_sel,
                    "Usuario_Acao": usuario_log_formatado,
                    "Status_Anterior": row_d.get("Status_Geral", "Pendente"),
                    "Novo_Status": novo_status_d,
                    "Area_Anterior": row_d.get(COL_GERENTE, "Operação"),
                    "Nova_Area": row_d.get(COL_GERENTE, "Operação"),
                    "Observacao": texto_obs_d,
                    "Tipo_Interacao": "Tratativa Divergência CIA"
                }])

                if novo_status_d in ["Já Lançado no ERP", "CIA Aérea Corrigida"]:
                    row_upd_d = row_d.copy()
                    row_upd_d["Status_Geral"] = novo_status_d
                    if cia_corrigida_d: row_upd_d["CIA"] = cia_corrigida_d
                    row_upd_d["Obs. Operação"] = texto_obs_d
                    row_upd_d["Status_Divergencia"] = "Valores Corretos"
                    
                    df_div_op = df_div_op.drop(idx_d)
                    df_sem_div = pd.concat([df_sem_div, pd.DataFrame([row_upd_d])], ignore_index=True)
                    msg_res_d = "resolvida e transferida para 'Sem Divergência'"
                else:
                    df_div_op.loc[idx_d, "Status_Geral"] = novo_status_d
                    if cia_corrigida_d: df_div_op.loc[idx_d, "CIA"] = cia_corrigida_d
                    df_div_op.loc[idx_d, "Obs. Operação"] = texto_obs_d
                    msg_res_d = "atualizada com sucesso"

                df_log_updated = pd.concat([df_log_master, novo_log_d], ignore_index=True)
                
                try:
                    with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_div_op.to_excel(writer, sheet_name="98_OK_Divergencia_Operacao", index=False)
                        df_sem_div.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
                        df_log_updated.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)
                    
                    st.session_state["msg_sucesso"] = f"✅ Divergência {bilhete_div_sel} {msg_res_d} por {usuario_log_formatado}!"
                    st.cache_data.clear()
                    st.rerun()
                except PermissionError:
                    st.error("❌ O arquivo Excel está aberto em outro programa.")
                    
        st.markdown("---")
        st.markdown(f"### 📊 Lista Completa das Divergências ({len(df_div_op_filtrado)} registros)")
        st.dataframe(df_div_op_filtrado, use_container_width=True, hide_index=True)

# ABA 3: SEM DIVERGÊNCIA
with abas_objetos[3]:
    st.subheader("✅ Base 98 - Bilhetes Prontos para Conciliação Operacional")
    st.dataframe(df_sem_div_filtrado, use_container_width=True, hide_index=True)

# ABA 4: SUPORTE BACKOFFICE
with abas_objetos[4]:
    st.subheader("🎧 Base 99 - Chamados Atribuídos ao Suporte Backoffice")
    
    if len(df_backoffice_filtrado) == 0:
        st.info("Nenhum chamado pendente no Suporte Backoffice no momento.")
    else:
        lista_back_bilhetes = df_backoffice_filtrado["Bilhetes"].astype(str).tolist()
        bilhete_back_sel = st.selectbox("Selecione o Chamado / Bilhete para Tratativa do Suporte:", options=lista_back_bilhetes, key="sb_backoffice")
        row_back = df_backoffice_filtrado[df_backoffice_filtrado["Bilhetes"].astype(str) == bilhete_back_sel].iloc[0]

        with st.form("form_solucao_backoffice"):
            acao_back = st.selectbox("Ação do Suporte / Auditoria:", options=["Informar que está Correto (Mover para Sem Divergência)", "Devolver para Tratativa Operacional"])
            
            if acao_back.startswith("Devolver"):
                area_devolucao_bk = st.selectbox("Área Operacional de Destino:", options=["Operação", "Central de Eventos", "Concierge/Lazer", "Unique", "Private"])
            else:
                area_original = str(row_back.get(COL_GERENTE, "Operação"))
                if area_original.lower() == "suporte backoffice": area_original = "Operação"
                st.text_input("Área Operacional Preservada:", value=area_original, disabled=True)
                area_devolucao_bk = area_original

            obs_back = st.text_area("Parecer do Suporte Backoffice / Auditoria:", value="", placeholder="Digite aqui o parecer técnico do suporte...")
            btn_salvar_back = st.form_submit_button("💾 Salvar Resolução do Suporte")

            if btn_salvar_back:
                idx_m_bk = df_master[df_master["Bilhetes"].astype(str) == bilhete_back_sel].index if "Bilhetes" in df_master.columns else []
                
                if acao_back.startswith("Informar"):
                    row_upd_bk = row_back.copy()
                    row_upd_bk["Status_Geral"] = "Já Lançado no ERP"
                    row_upd_bk[COL_GERENTE] = area_devolucao_bk
                    row_upd_bk["Status_Divergencia"] = "Valores Corretos"
                    row_upd_bk["Obs. Operação"] = f"[Correto pelo Suporte]: {obs_back}"
                    
                    if len(idx_m_bk) > 0: df_master = df_master.drop(idx_m_bk)
                    df_sem_div = pd.concat([df_sem_div, pd.DataFrame([row_upd_bk])], ignore_index=True)
                    
                    novo_status_log = "Já Lançado no ERP (Sem Divergência)"
                    area_destino_log = area_devolucao_bk
                    msg_sucesso_bk = f"🎉 Chamado {bilhete_back_sel} resolvido com área '{area_devolucao_bk}' e movido para **Sem Divergência**!"
                else:
                    if len(idx_m_bk) > 0:
                        df_master.loc[idx_m_bk, "Status_Geral"] = "Pendente de Lançamento"
                        df_master.loc[idx_m_bk, COL_GERENTE] = area_devolucao_bk
                        df_master.loc[idx_m_bk, "Obs. Operação"] = f"[Devolvido pelo Suporte]: {obs_back}"
                        
                    novo_status_log = f"Devolvido para {area_devolucao_bk}"
                    area_destino_log = area_devolucao_bk
                    msg_sucesso_bk = f"🔄 Chamado {bilhete_back_sel} devolvido com sucesso para **{area_devolucao_bk}**!"

                novo_log_bk = pd.DataFrame([{
                    "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Bilhete": bilhete_back_sel,
                    "Usuario_Acao": usuario_log_formatado,
                    "Status_Anterior": row_back.get("Status_Geral", "Pendente"),
                    "Novo_Status": novo_status_log,
                    "Area_Anterior": "Suporte backoffice",
                    "Nova_Area": area_destino_log,
                    "Observacao": obs_back,
                    "Tipo_Interacao": "Tratativa Suporte Backoffice"
                }])

                df_backoffice_atualizado = df_master[df_master[COL_GERENTE].astype(str) == "Suporte backoffice"].copy()
                df_log_updated = pd.concat([df_log_master, novo_log_bk], ignore_index=True)

                try:
                    with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_master.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
                        df_sem_div.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
                        if not df_backoffice_atualizado.empty:
                            df_backoffice_atualizado.to_excel(writer, sheet_name="99_Suporte backoffice", index=False)
                        df_log_updated.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)

                    st.session_state["msg_sucesso"] = msg_sucesso_bk
                    st.cache_data.clear()
                    st.rerun()
                except PermissionError:
                    st.error("❌ O arquivo Excel está aberto em outro programa.")

        st.markdown("---")
        st.markdown(f"### 📊 Lista Completa dos Chamados ({len(df_backoffice_filtrado)} registros)")
        st.dataframe(df_backoffice_filtrado, use_container_width=True, hide_index=True)

# ABA 5: RÉPLICA DA AUDITORIA
with abas_objetos[5]:
    st.subheader("⚖️ Módulo de Contestação e Réplica da Auditoria")
    bilhetes_com_tratativa = df_acao_filtrado[df_acao_filtrado["Status_Geral"] != "Pendente de Lançamento"]
    if len(bilhetes_com_tratativa) > 0:
        bilhete_rep = st.selectbox("Selecione o Bilhete para Réplica:", options=bilhetes_com_tratativa["Bilhetes_Str"].tolist())
        row_rep = df_acao_filtrado[df_acao_filtrado["Bilhetes_Str"] == bilhete_rep].iloc[0]
        
        with st.form("form_replica_auditoria"):
            status_auditoria = st.selectbox("Decisão da Auditoria:", options=["Contestado / Recusado", "Aprovado / Conciliado"])
            area_devolucao = st.selectbox("Devolver para Área / Gerente:", options=["Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private", "Operação"])
            motivo_replica = st.text_area("Justificativa da Auditoria:", value="", placeholder="Digite a justificativa para a réplica...")
            btn_replica = st.form_submit_button("🚨 Enviar Apontamento")
            
            if btn_replica:
                idx_m_rep = df_master[df_master["Bilhetes"].astype(str).str.strip() == bilhete_rep].index if "Bilhetes" in df_master.columns else []
                novo_log_rep = pd.DataFrame([{
                    "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Bilhete": bilhete_rep,
                    "Usuario_Acao": usuario_log_formatado,
                    "Status_Anterior": row_rep.get("Status_Geral", "-"),
                    "Novo_Status": f"Contestado ({area_devolucao})",
                    "Area_Anterior": row_rep.get(COL_GERENTE, "-"),
                    "Nova_Area": area_devolucao,
                    "Observacao": motivo_replica,
                    "Tipo_Interacao": "Réplica da Auditoria"
                }])
                
                if len(idx_m_rep) > 0:
                    df_master.loc[idx_m_rep, "Status_Geral"] = f"Contestado ({area_devolucao})"
                    df_master.loc[idx_m_rep, COL_GERENTE] = area_devolucao

                df_log_updated = pd.concat([df_log_master, novo_log_rep], ignore_index=True)
                
                try:
                    with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_master.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
                        df_log_updated.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)
                    st.session_state["msg_sucesso"] = f"✅ Contestação registrada por {usuario_log_formatado}!"
                    st.cache_data.clear()
                    st.rerun()
                except PermissionError:
                    st.error("❌ Feche a planilha para salvar.")

        st.markdown("---")
        st.markdown(f"### 📊 Lista Completa dos Casos em Réplica ({len(bilhetes_com_tratativa)} registros)")
        st.dataframe(bilhetes_com_tratativa, use_container_width=True, hide_index=True)

# ABAS EXCLUSIVAS DO COMPLIANCE
idx_aba_compliance = 6

if st.session_state["perfil_atual"] == "Compliance":
    # ABA TRILHA DE AUDITORIA
    with abas_objetos[idx_aba_compliance]:
        st.subheader("📜 Histórico Completo de Alterações e Trilha de Auditoria")
        
        with st.expander("🗑️ Módulo Master de Gerenciamento e Exclusão de Logs"):
            if not df_log_master.empty:
                indices_log = df_log_master.index.tolist()
                log_opcao = st.selectbox(
                    "Selecione o registro de log que deseja EXCLUIR:",
                    options=indices_log,
                    format_func=lambda i: f"Linha {i} | Data: {df_log_master.loc[i, 'Data_Hora']} | Bilhete: {df_log_master.loc[i, 'Bilhete']} | Usuário: {df_log_master.loc[i, 'Usuario_Acao']}"
                )
                if st.button("❌ Excluir Registro de Log Selecionado"):
                    df_log_master = df_log_master.drop(log_opcao).reset_index(drop=True)
                    try:
                        with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_log_master.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)
                        st.session_state["msg_sucesso"] = "🗑️ Registro de log excluído com sucesso!"
                        st.cache_data.clear()
                        st.rerun()
                    except PermissionError:
                        st.error("❌ O arquivo Excel está aberto em outro programa.")
            else:
                st.info("Nenhum registro de log para exclusão.")
                
        st.dataframe(df_log_master, use_container_width=True, hide_index=False)
    idx_aba_compliance += 1

    # ABA GESTÃO DE ACESSOS & APROVAÇÕES
    with abas_objetos[idx_aba_compliance]:
        st.subheader("⚙️ Central de Aprovações de Acesso e Governança")
        
        usuarios_atuais = carregar_usuarios()
        pendentes = {k: v for k, v in usuarios_atuais.items() if v.get("status") == "PENDENTE"}
        
        if not pendentes:
            st.info("🎉 Nenhuma solicitação de acesso pendente no momento.")
        else:
            for u_id, u_info in pendentes.items():
                st.markdown(f"""
                    <div style="background-color: #ffffff; border-left: 5px solid #002060; padding: 15px; border-radius: 8px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);">
                        <b>Nome:</b> {u_info['nome']} | <b>Usuário:</b> {u_id} | <b>Perfil Solicitado:</b> {u_info['perfil']}<br>
                        <small style="color: #6c757d;">Data da Solicitação: {u_info.get('data_solicitacao', '-')}</small>
                    </div>
                """, unsafe_allow_html=True)
                
                col_ap, col_rej, _ = st.columns([1, 1, 4])
                with col_ap:
                    if st.button(f"✅ Aprovar {u_id}", key=f"btn_ap_{u_id}"):
                        usuarios_atuais[u_id]["status"] = "APROVADO"
                        salvar_usuarios(usuarios_atuais)
                        st.session_state["msg_sucesso"] = f"✅ Usuário {u_id} aprovado com sucesso!"
                        st.rerun()
                with col_rej:
                    if st.button(f"❌ Rejeitar {u_id}", key=f"btn_rej_{u_id}"):
                        usuarios_atuais[u_id]["status"] = "REJEITADO"
                        salvar_usuarios(usuarios_atuais)
                        st.session_state["msg_sucesso"] = f"🚫 Usuário {u_id} rejeitado!"
                        st.rerun()

    idx_aba_compliance += 1

# ABA VISÃO GERAL BASE TOTAL
with abas_objetos[idx_aba_compliance]:
    st.subheader("📋 Visão Geral da Base Total de Divergências")
    st.dataframe(df_acao_filtrado, use_container_width=True, hide_index=True)