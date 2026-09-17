import datetime
import io
import json
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import streamlit as st
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. CONFIGURAÇÃO INICIAL E ESTILOS CSS
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Revenue Assurance",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_DASHBOARD = "Dashboard_Revenue_Assurance_Consolidado.xlsx"
ARQUIVO_USUARIOS = "usuarios_autorizados.json"
PASTA_ENVIO = "Envio"
PASTA_INPUTS_LOCAL = r"C:\Users\mribeiro1\MARINGA TURISMO\Maringá Turismo - PLANEJAMENTO ESTRATEGICO (1)\01-Planejamento Estratégico\Auditoria de Bilhetes\inputs"

USUARIOS_PADRAO = {
    "mribeiro": {
        "senha": "123",
        "nome": "Marcos Ribeiro",
        "perfil": "Master",
        "status": "APROVADO",
        "data_solicitacao": "2026-08-31",
    },
    "fellipe": {
        "senha": "123",
        "nome": "Fellipe Fernandes",
        "perfil": "Master",
        "status": "APROVADO",
        "data_solicitacao": "2026-08-31",
    },
    "backoffice": {
        "senha": "123",
        "nome": "Atendimento Backoffice",
        "perfil": "Operacao",
        "status": "APROVADO",
        "data_solicitacao": "2026-08-31",
    },
}

st.markdown(
    """
    <style>
        .stApp { background-color: #f8f9fa; }
        .main-header {
            background: linear-gradient(135deg, #002060 0%, #003366 100%);
            padding: 18px 25px;
            border-radius: 10px;
            color: white;
            margin-bottom: 15px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        .main-header h1 { color: #ffffff !important; margin: 0; font-size: 22px; font-weight: 700; }
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border-radius: 8px;
            padding: 12px 18px;
            border-left: 5px solid #002060;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        }
        div.stButton > button {
            background-color: #002060 !important;
            color: white !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# ==============================================================================
# CONEXÃO COM SUPABASE (POSTGRESQL NUVEM)
# ==============================================================================
def get_db_engine():
    """Conecta ao Supabase buscando a URI configurada nos Secrets do Streamlit."""
    try:
        db_url = st.secrets["postgres"]["url"]
        return create_engine(db_url, pool_pre_ping=True)
    except Exception as e:
        st.error(f"⚠️ Erro ao obter credenciais do Supabase. Verifique os Secrets do Streamlit. ({e})")
        return None

def carregar_tratativas_db():
    """Busca as tratativas ativas diretamente no Supabase."""
    engine = get_db_engine()
    if not engine:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao"])
    
    query = "SELECT bilhete, status_geral, area_resp, obs_operacao FROM tratativas"
    try:
        return pd.read_sql(query, engine)
    except Exception:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao"])

def salvar_tratativa_supabase(bilhete, status_geral, area_resp, obs_operacao, usuario):
    """Executa o UPSERT da tratativa individual no Supabase."""
    engine = get_db_engine()
    if not engine:
        return

    upsert_sql = text("""
        INSERT INTO tratativas (bilhete, status_geral, area_resp, obs_operacao, usuario_modificacao, data_modificacao)
        VALUES (:bilhete, :status_geral, :area_resp, :obs_operacao, :usuario, NOW())
        ON CONFLICT (bilhete) DO UPDATE SET
            status_geral = EXCLUDED.status_geral,
            area_resp = EXCLUDED.area_resp,
            obs_operacao = EXCLUDED.obs_operacao,
            usuario_modificacao = EXCLUDED.usuario_modificacao,
            data_modificacao = NOW();
    """)
    
    with engine.begin() as conn:
        conn.execute(upsert_sql, {
            "bilhete": str(bilhete).strip().replace(".0", ""),
            "status_geral": status_geral,
            "area_resp": area_resp,
            "obs_operacao": obs_operacao,
            "usuario": usuario
        })

def registrar_log_supabase(logs_list):
    """Persiste a trilha de auditoria na tabela log_auditoria do Supabase."""
    engine = get_db_engine()
    if not engine or not logs_list:
        return
    
    df_logs = pd.DataFrame(logs_list)
    df_logs.rename(columns={
        "Data_Hora": "data_hora",
        "Bilhete": "bilhete",
        "Usuario_Acao": "usuario_acao",
        "Status_Anterior": "status_anterior",
        "Novo_Status": "novo_status",
        "Area_Anterior": "area_anterior",
        "Nova_Area": "nova_area",
        "Observacao": "observacao",
        "Tipo_Interacao": "tipo_interacao"
    }, inplace=True)
    
    try:
        df_logs.to_sql("log_auditoria", engine, if_exists="append", index=False)
    except Exception as e:
        st.warning(f"Aviso ao gravar log no Supabase: {e}")

# ==============================================================================
# 2. GESTÃO DE USUÁRIOS E SEGURANÇA MASTER
# ==============================================================================
def carregar_usuarios():
    if not os.path.exists(ARQUIVO_USUARIOS):
        with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
            json.dump(USUARIOS_PADRAO, f, ensure_ascii=False, indent=4)
        return USUARIOS_PADRAO
    try:
        with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return USUARIOS_PADRAO

def salvar_usuarios(dict_users):
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as f:
        json.dump(dict_users, f, ensure_ascii=False, indent=4)

usuarios_db = carregar_usuarios()

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_atual" not in st.session_state:
    st.session_state["usuario_atual"] = None
if "perfil_atual" not in st.session_state:
    st.session_state["perfil_atual"] = None
if "login_user_id" not in st.session_state:
    st.session_state["login_user_id"] = None

def e_master():
    u_id = st.session_state.get("login_user_id", "")
    perfil = st.session_state.get("perfil_atual", "")
    return u_id == "mribeiro" or perfil in ["Master", "Compliance"]

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown("<div class='main-header'><h1>✈️ Grupo Arbaitman | Login</h1></div>", unsafe_allow_html=True)
        tab_log, tab_pwd, tab_req = st.tabs(["🔐 Entrar", "🔑 Esqueci a Senha", "📝 Solicitar Acesso"])
        
        with tab_log:
            u_input = st.text_input("Usuário:", key="l_user").strip().lower()
            p_input = st.text_input("Senha:", type="password", key="l_pass").strip()
            if st.button("Acessar Portal", type="primary"):
                if u_input in usuarios_db and usuarios_db[u_input]["senha"] == p_input:
                    if usuarios_db[u_input].get("status", "APROVADO") == "APROVADO":
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_atual"] = usuarios_db[u_input]["nome"]
                        st.session_state["perfil_atual"] = usuarios_db[u_input]["perfil"]
                        st.session_state["login_user_id"] = u_input
                        st.rerun()
                    else:
                        st.warning("⏳ Seu usuário está aguardando aprovação do Gestor Master.")
                else:
                    st.error("Usuário ou senha incorretos.")

        with tab_pwd:
            u_reset = st.text_input("Usuário cadastrado:").strip().lower()
            n_pass = st.text_input("Nova Senha:", type="password")
            c_pass = st.text_input("Confirme a Nova Senha:", type="password")
            if st.button("Redefinir Senha"):
                if u_reset in usuarios_db and n_pass and n_pass == c_pass:
                    usuarios_db[u_reset]["senha"] = n_pass
                    salvar_usuarios(usuarios_db)
                    st.success("✅ Senha redefinida com sucesso!")
                else:
                    st.error("Verifique os dados informados.")

        with tab_req:
            r_nome = st.text_input("Nome Completo:")
            r_user = st.text_input("Usuário Desejado:").strip().lower()
            r_pass = st.text_input("Senha de Acesso:", type="password")
            r_perf = st.selectbox("Perfil Solicitado:", ["Operacao"])
            if st.button("Enviar Solicitação"):
                if r_user and r_pass and r_nome:
                    usuarios_db[r_user] = {
                        "senha": r_pass, "nome": r_nome, "perfil": r_perf,
                        "status": "PENDENTE", "data_solicitacao": datetime.datetime.now().strftime("%Y-%m-%d")
                    }
                    salvar_usuarios(usuarios_db)
                    st.success("✅ Solicitação enviada! Aguarde a liberação do usuário Master.")
    st.stop()

# ==============================================================================
# 3. TRATAMENTO DE DADOS, ARMAZENAMENTO E CARGA
# ==============================================================================
def clean_str(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    return re.sub(r"\.0$", "", s)

def padronizar_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Bilhetes", "Ponto de venda", "Status_Geral", "Área Resp. Operação",
            "Obs. Operação", "Setor", "Data Emissão", "CIA", "Taxa", "A vista", "A credito",
            "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao"
        ])
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    if "Bilhetes" not in df_out.columns:
        cols_b = [c for c in df_out.columns if "bilhete" in str(c).lower() or "ticket" in str(c).lower()]
        df_out = df_out.rename(columns={cols_b[0]: "Bilhetes"}) if cols_b else df_out.assign(Bilhetes="")
    df_out["Bilhetes"] = df_out["Bilhetes"].apply(clean_str)
    
    for c in ["Ponto de venda", "Área Resp. Operação", "Obs. Operação", "Setor", "Status_Geral", "CIA", "Data Emissão"]:
        if c not in df_out.columns: df_out[c] = "-"
    
    for c in ["Taxa", "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa", "Receita_Sistema", "Dif_Receita"]:
        if c in df_out.columns: df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)

    for c in ["Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao"]:
        if c not in df_out.columns: df_out[c] = "-"

    df_out["Dt_Parsed"] = pd.to_datetime(df_out["Data Emissão"], format="mixed", dayfirst=True, errors="coerce")
    
    cols_prioritarias = [
        "Bilhetes", "Ponto de venda", "CIA", "Data Emissão", "Status_Geral",
        "Área Resp. Operação", "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao", "Obs. Operação"
    ]
    existentes = [c for c in cols_prioritarias if c in df_out.columns]
    outras = [c for c in df_out.columns if c not in existentes]
    return df_out[existentes + outras]

def mesclar_com_supabase(df_in):
    """Sobrescreve o status e observações com os dados salvos no Supabase."""
    df_db = carregar_tratativas_db()
    if df_in is None or df_in.empty or df_db.empty:
        return df_in

    df_in["Bilhete_Clean"] = df_in["Bilhetes"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df_db["Bilhete_Clean"] = df_db["bilhete"].astype(str).str.strip()

    df_merged = pd.merge(
        df_in,
        df_db[["Bilhete_Clean", "status_geral", "area_resp", "obs_operacao"]],
        on="Bilhete_Clean",
        how="left"
    )

    df_merged["Status_Geral"] = df_merged["status_geral"].fillna(df_merged["Status_Geral"])
    df_merged["Área Resp. Operação"] = df_merged["area_resp"].fillna(df_merged["Área Resp. Operação"])
    df_merged["Obs. Operação"] = df_merged["obs_operacao"].fillna(df_merged["Obs. Operação"])

    df_merged.drop(columns=["Bilhete_Clean", "status_geral", "area_resp", "obs_operacao"], inplace=True)
    return df_merged

def rotear_bases_backoffice(df_m, df_d, df_s, df_b):
    df_m = padronizar_df(df_m)
    df_d = padronizar_df(df_d)
    df_s = padronizar_df(df_s)
    df_b = padronizar_df(df_b)

    def e_backoffice_row(r):
        st_val = str(r.get("Status_Geral", "")).strip().lower()
        ar_val = str(r.get("Área Resp. Operação", "")).strip().lower()
        se_val = str(r.get("Setor", "")).strip().lower()
        obs_val = str(r.get("Obs. Operação", "")).strip().lower()

        if "backoffice" in st_val or "suporte" in st_val or "encaminhado" in st_val:
            return True
        if any(k in ar_val for k in ["backoffice", "suporte", "benner", "katia", "ti"]):
            return True
        if any(k in se_val for k in ["backoffice", "suporte"]):
            return True
        if any(p in obs_val for p in ["ticket", "chamado", "backoffice", "suporte", "benner", "erro de integração", "erro integração", "aguardando suporte"]):
            return True
        return False

    def e_resolvido_row(r):
        st_val = str(r.get("Status_Geral", "")).strip().lower()
        return any(term in st_val for term in ["já lançado", "conciliado", "regularizado", "sem divergência", "ok"])

    df_all = pd.concat([df_m, df_d, df_b], ignore_index=True)
    if df_all.empty or "Bilhetes" not in df_all.columns:
        return df_m, df_d, df_s, df_b

    df_all = df_all.drop_duplicates(subset=["Bilhetes"], keep="last")

    mask_bo = df_all.apply(e_backoffice_row, axis=1)
    mask_ok = df_all.apply(e_resolvido_row, axis=1)

    df_b_novo = df_all[mask_bo & ~mask_ok].copy()
    if not df_b_novo.empty:
        df_b_novo["Setor"] = "Suporte backoffice"
        df_b_novo["Área Resp. Operação"] = df_b_novo["Área Resp. Operação"].apply(
            lambda x: "Suporte Backoffice" if str(x).strip() in ["-", "", "Não Mapeado", "nan", "None"] else x
        )

    df_s_novos = df_all[mask_ok].copy()
    if not df_s_novos.empty:
        df_s = pd.concat([df_s, df_s_novos], ignore_index=True).drop_duplicates(subset=["Bilhetes"], keep="last")

    df_resto = df_all[~mask_bo & ~mask_ok].copy()
    
    if "Status_Divergencia" in df_resto.columns:
        mask_div = df_resto["Status_Divergencia"].astype(str).str.contains("Divergência", case=False, na=False)
        df_d_novo = df_resto[mask_div].copy()
        df_m_novo = df_resto[~mask_div].copy()
    else:
        df_m_novo = df_resto
        df_d_novo = pd.DataFrame()

    return padronizar_df(df_m_novo), padronizar_df(df_d_novo), padronizar_df(df_s), padronizar_df(df_b_novo)

@st.cache_data(ttl=60)
def carregar_bases():
    vazio = padronizar_df(None)
    
    if not os.path.exists(ARQUIVO_DASHBOARD):
        return vazio, vazio, vazio, vazio, pd.DataFrame()

    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD, engine="openpyxl")
        m = pd.read_excel(xls, "99_Base_Divergencias_Geral") if "99_Base_Divergencias_Geral" in xls.sheet_names else None
        d = pd.read_excel(xls, "98_OK_Divergencia_Operacao") if "98_OK_Divergencia_Operacao" in xls.sheet_names else None
        s = pd.read_excel(xls, "98_OK_Sem_Divergencia_Concil") if "98_OK_Sem_Divergencia_Concil" in xls.sheet_names else None
        b = pd.read_excel(xls, "99_Suporte backoffice") if "99_Suporte backoffice" in xls.sheet_names else None
        l = pd.read_excel(xls, "00_Log_Auditoria") if "00_Log_Auditoria" in xls.sheet_names else pd.DataFrame()

        # Aplica Supabase por cima do Excel
        m = mesclar_com_supabase(m)
        d = mesclar_com_supabase(d)
        s = mesclar_com_supabase(s)
        b = mesclar_com_supabase(b)

        m_p, d_p, s_p, b_p = rotear_bases_backoffice(m, d, s, b)
        return m_p, d_p, s_p, b_p, l
    except Exception as e:
        st.error(f"⚠️ Erro ao carregar as bases de dados: {e}")
        return vazio, vazio, vazio, vazio, pd.DataFrame()

def gerar_excel_estilizado(df_export, nome_aba="Relatorio_Filtrado"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(columns=["Aviso"], data=[["Nenhum registro encontrado para os filtros selecionados"]])
    
    cols_remover = [c for c in ["Dt_Parsed"] if c in df_export.columns]
    df_clean = df_export.drop(columns=cols_remover) if cols_remover else df_export.copy()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name=nome_aba[:30], index=False)
    
    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    ws = wb.active
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()

with st.spinner("🔄 Carregando bases de dados do painel e conectando ao Supabase..."):
    df_master, df_div_op, df_sem_div, df_backoffice, df_log = carregar_bases()

# ==============================================================================
# 4. SIDEBAR E FILTROS OPERACIONAIS
# ==============================================================================
st.sidebar.title("Navegação")
st.sidebar.write(f"👤 **{st.session_state['usuario_atual']}** ({st.session_state['perfil_atual']})")

if st.sidebar.button("🔑 Alterar Minha Senha"):
    with st.sidebar.form("form_pwd"):
        s_atu = st.text_input("Senha Atual:", type="password")
        s_nov = st.text_input("Nova Senha:", type="password")
        if st.form_submit_button("Salvar"):
            u_id = st.session_state["login_user_id"]
            if usuarios_db[u_id]["senha"] == s_atu and s_nov:
                usuarios_db[u_id]["senha"] = s_nov
                salvar_usuarios(usuarios_db)
                st.sidebar.success("Senha alterada com sucesso!")

if st.sidebar.button("🔒 Sair"):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Filtros Operacionais")

d_inicio = st.sidebar.date_input("Data Inicial:", value=datetime.date(2025, 1, 1), format="DD/MM/YYYY")
d_fim = st.sidebar.date_input("Data Final:", value=datetime.date(2026, 12, 31), format="DD/MM/YYYY")

df_todos = pd.concat([df_master, df_div_op, df_sem_div, df_backoffice], ignore_index=True)

filtro_gerente = st.sidebar.multiselect("Gerente / Área Resp.:", options=sorted(df_todos["Área Resp. Operação"].dropna().unique()))
filtro_setor = st.sidebar.multiselect("Setor:", options=sorted(df_todos["Setor"].dropna().unique()))
filtro_cia = st.sidebar.multiselect("Companhia Aérea:", options=sorted(df_todos["CIA"].dropna().unique()))

def aplicar_filtros(df):
    if df.empty: return df
    m = pd.Series(True, index=df.index)
    if "Dt_Parsed" in df.columns:
        m = m & (df["Dt_Parsed"].dt.date >= d_inicio) & (df["Dt_Parsed"].dt.date <= d_fim)
    if filtro_gerente: m = m & df["Área Resp. Operação"].isin(filtro_gerente)
    if filtro_setor: m = m & df["Setor"].isin(filtro_setor)
    if filtro_cia: m = m & df["CIA"].isin(filtro_cia)
    return df[m]

f_master = aplicar_filtros(df_master)
f_div_op = aplicar_filtros(df_div_op)
f_sem_div = aplicar_filtros(df_sem_div)
f_backoffice = aplicar_filtros(df_backoffice)

# ==============================================================================
# 5. HEADER PRINCIPAL E CARDS DE KPIS
# ==============================================================================
st.markdown("<div class='main-header'><h1>✈️ Garantia de Receita do Portal | Auditoria de Bilhetes</h1></div>", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Pendentes de Lançamento", f"{len(f_master):,}")
c2.metric("Com Divergência Operacional", f"{len(f_div_op):,}")
c3.metric("Sem Divergência (OK)", f"{len(f_sem_div):,}")
c4.metric("Backoffice", f"{len(f_backoffice):,}")

st.markdown("---")

# ==============================================================================
# 6. ESTRUTURA DE ABAS
# ==============================================================================
abas = [
    "📊 Dashboard Executivo",
    "📋 1. Falta de Lançamento",
    "⚠️ 2. Erros de Valores & CIA",
    "🎧 3. Suporte Backoffice",
    "✅ 4. Sem Divergência (OK)"
]

if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Auditoria", "📥 Carga de Relatórios (Lotes)"])

aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# ABA 0: DASHBOARD EXECUTIVO
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.subheader("📊 Painel Executivo de Revenue Assurance")
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        st.markdown("##### ⚠️ Visão de Divergências Operacionais")
        if not f_div_op.empty and "Status_Divergencia" in f_div_op.columns:
            df_div_chart = f_div_op["Status_Divergencia"].value_counts().reset_index()
            df_div_chart.columns = ["Status_Divergencia", "count"]
            fig_div = px.bar(df_div_chart, x="count", y="Status_Divergencia", orientation="h", text="count", color_discrete_sequence=["#002060"])
            st.plotly_chart(fig_div, width="stretch")
        else:
            st.info("Nenhuma divergência registrada no período.")

    with col_d2:
        st.markdown("##### 👤 Pendências por Gerente Responsável")
        if not f_master.empty:
            df_ger_chart = f_master["Área Resp. Operação"].value_counts().head(8).reset_index()
            df_ger_chart.columns = ["Gerente", "count"]
            fig_ger = px.bar(df_ger_chart, x="Gerente", y="count", text="count", color_discrete_sequence=["#00509d"])
            st.plotly_chart(fig_ger, width="stretch")

# ------------------------------------------------------------------------------
# FUNÇÃO PARA RENDERIZAÇÃO E TRATATIVA VIA SUPABASE
# ------------------------------------------------------------------------------
def renderizar_modulo_tratativa(df_filtrado, nome_base, key_prefix):
    st.download_button(
        f"📥 Extrair Relatório ({nome_base})",
        data=gerar_excel_estilizado(df_filtrado, nome_base),
        file_name=f"{nome_base}_Filtrado.xlsx",
        key=f"btn_dl_{key_prefix}"
    )
    
    bilhetes_lista = df_filtrado["Bilhetes"].tolist()
    if bilhetes_lista:
        bilhet_sel = st.multiselect("Selecione um ou mais Bilhetes para Tratativa:", options=bilhetes_lista, key=f"ms_{key_prefix}")
        if bilhet_sel:
            df_sel = df_filtrado[df_filtrado["Bilhetes"].isin(bilhet_sel)].iloc[0]
            st.info(
                f"📌 **Dados do Primeiro Bilhete Selecionado:** CIA: `{df_sel.get('CIA', '-')}` | "
                f"Taxa: `R$ {df_sel.get('Taxa', 0.0):,.2f}` | Gerente Atual: `{df_sel.get('Área Resp. Operação', '-')}` | "
                f"Status Atual: `{df_sel.get('Status_Geral', '-')}`"
            )
            
            with st.form(f"form_trat_{key_prefix}"):
                c_s1, c_s2 = st.columns(2)
                with c_s1:
                    opcoes_factiveis = [
                        "Já Lançado no ERP",
                        "Pendente de Lançamento",
                        "Divergência de Tarifa e Taxa",
                        "Divergência de Receita",
                        "Divergência de CIA Aérea",
                        "Encaminhado para Suporte Backoffice",
                        "Aguardando Retorno da Operação",
                        "Cancelado / Reembolsado"
                    ]
                    status_atual_val = str(df_sel.get("Status_Geral", "Pendente de Lançamento"))
                    idx_default = opcoes_factiveis.index(status_atual_val) if status_atual_val in opcoes_factiveis else 0
                    n_status = st.selectbox("Novo Status (Obrigatório):", options=opcoes_factiveis, index=idx_default, key=f"st_{key_prefix}")
                
                with c_s2:
                    n_area = st.text_input("Nova Área Responsável / Gerente:", value=str(df_sel.get("Área Resp. Operação", "")), key=f"ar_{key_prefix}")
                
                n_obs = st.text_area("Observação / Justificativa Detalhada:", value=str(df_sel.get("Obs. Operação", "")), key=f"obs_{key_prefix}")
                
                if st.form_submit_button("💾 Salvar e Atualizar Bilhetes"):
                    with st.spinner("⚡ Gravando diretamente no banco Supabase..."):
                        agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                        
                        novos_logs = []
                        for b_sel in bilhet_sel:
                            # 1. Salva no Supabase
                            salvar_tratativa_supabase(
                                bilhete=b_sel,
                                status_geral=n_status,
                                area_resp=n_area,
                                obs_operacao=n_obs,
                                usuario=usr_str
                            )
                            
                            novos_logs.append({
                                "Data_Hora": agora_str,
                                "Bilhete": b_sel,
                                "Usuario_Acao": usr_str,
                                "Status_Anterior": df_sel.get("Status_Geral", "-"),
                                "Novo_Status": n_status,
                                "Area_Anterior": df_sel.get("Área Resp. Operação", "-"),
                                "Nova_Area": n_area,
                                "Observacao": n_obs,
                                "Tipo_Interacao": f"Tratativa Web ({nome_base})"
                            })
                        
                        # 2. Grava a auditoria no Supabase
                        registrar_log_supabase(novos_logs)

                        st.cache_data.clear()
                        st.success("✅ Tratativa gravada no Supabase em nuvem com sucesso!")
                        st.rerun()

    st.markdown("---")
    st.dataframe(df_filtrado, width="stretch", hide_index=True)

# ------------------------------------------------------------------------------
# ABA 1: FALTA DE LANÇAMENTO
# ------------------------------------------------------------------------------
with aba_sel[1]:
    st.subheader("📋 1. Pendências de Lançamento no ERP")
    renderizar_modulo_tratativa(f_master, "Falta_de_Lancamento", "fl")

# ------------------------------------------------------------------------------
# ABA 2: ERROS DE VALORES E CIA AÉREA
# ------------------------------------------------------------------------------
with aba_sel[2]:
    st.subheader("⚠️ 2. Divergências de Valores e Companhias Aéreas")
    renderizar_modulo_tratativa(f_div_op, "Erros_Valores_CIA", "ev")

# ------------------------------------------------------------------------------
# ABA 3: SUPORTE BACKOFFICE
# ------------------------------------------------------------------------------
with aba_sel[3]:
    st.subheader("🎧 3. Chamados em Análise no Suporte Backoffice")
    renderizar_modulo_tratativa(f_backoffice, "Suporte_Backoffice", "sb")

# ------------------------------------------------------------------------------
# ABA 4: SEM DIVERGÊNCIA (OK)
# ------------------------------------------------------------------------------
with aba_sel[4]:
    st.subheader("✅ 4. Bilhetes Conciliados e Lançados")
    st.dataframe(f_sem_div, width="stretch", hide_index=True)

# ------------------------------------------------------------------------------
# ABAS EXCLUSIVAS DO MASTER
# ------------------------------------------------------------------------------
if e_master():
    # ABA 5: GESTÃO DE ACESSOS
    with aba_sel[5]:
        st.subheader("⚙️ Central de Aprovação de Acessos")
        pendentes = {k: v for k, v in usuarios_db.items() if v.get("status") == "PENDENTE"}
        if pendentes:
            for u_k, u_v in pendentes.items():
                st.write(f"👤 **{u_v['nome']}** (`{u_k}`) | Perfil: **{u_v['perfil']}**")
                ca1, ca2, _ = st.columns([1, 1, 4])
                if ca1.button(f"✅ Aprovar {u_k}"):
                    usuarios_db[u_k]["status"] = "APROVADO"
                    salvar_usuarios(usuarios_db)
                    st.success("Usuário aprovado!")
                    st.rerun()
                if ca2.button(f"❌ Rejeitar {u_k}"):
                    usuarios_db[u_k]["status"] = "REJEITADO"
                    salvar_usuarios(usuarios_db)
                    st.rerun()
        else:
            st.info("Nenhuma solicitação de acesso pendente.")

    # ABA 6: LOG DE AUDITORIA
    with aba_sel[6]:
        st.subheader("📜 Trilha de Auditoria do Supabase")
        engine_sb = get_db_engine()
        if engine_sb:
            try:
                df_log_db = pd.read_sql("SELECT * FROM log_auditoria ORDER BY id DESC LIMIT 500", engine_sb)
                st.dataframe(df_log_db, width="stretch", hide_index=True)
            except Exception:
                st.info("Nenhum registro de log encontrado na tabela log_auditoria do Supabase.")

    # ABA 7: CARGA DE RELATÓRIOS
    with aba_sel[7]:
        st.subheader("📥 Carga de Relatórios Semanais")
        st.info("Envie aqui relatórios adicionais em lote para alimentar a visualização.")