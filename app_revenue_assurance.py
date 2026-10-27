import datetime
import io
import json
import os
import re
import time
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from sqlalchemy import create_engine, text

# ==============================================================================
# 1. CONFIGURAÇÃO INICIAL E ESTILOS CSS RECOMPACTADOS (CABEÇALHO E BOTÕES)
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Portal de Conciliação Aérea FP&A",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_DASHBOARD_PADRAO = "Dashboard_Revenue_Assurance_Consolidado.xlsx"

DEPARA_CIAS = {
    "G3": "GOL (G3)", "AD": "AZUL (AD)", "LA": "LATAM (LA)", "JJ": "LATAM (JJ)",
    "TP": "TAP Portugal (TP)", "AA": "American Airlines (AA)", "AF": "Air France (AF)",
    "KL": "KLM (KL)", "LH": "Lufthansa (LH)", "UA": "United Airlines (UA)",
    "DL": "Delta Air Lines (DL)", "AM": "Aeromexico (AM)", "AV": "Avianca (AV)",
    "AR": "Aerolíneas Argentinas (AR)", "CM": "Copa Airlines (CM)", "IB": "Iberia (IB)",
    "EK": "Emirates (EK)", "QR": "Qatar Airways (QR)", "LX": "Swiss Air (LX)",
    "AZ": "ITA Airways (AZ)", "AC": "Air Canada (AC)", "SA": "South African (SA)",
    "OB": "Boliviana de Aviación (OB)", "HR": "Hahn Air (HR)", "15": "15 - Outras CIAs"
}

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }
        .stApp { background-color: #f8fafc; }
        
        /* Redução de Padding da Página */
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 1rem !important;
            max-width: 98% !important;
        }
        
        /* Cabeçalho Compacto C-Level */
        .main-header {
            background: linear-gradient(135deg, #001e4d 0%, #003366 100%);
            padding: 14px 24px !important;
            border-radius: 8px;
            color: white;
            margin-bottom: 12px !important;
            box-shadow: 0 4px 14px rgba(0, 32, 96, 0.15);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .main-header h1 { 
            color: #ffffff !important; 
            margin: 0; 
            font-size: 20px !important; 
            font-weight: 700; 
            letter-spacing: -0.3px;
        }
        .main-header p { 
            color: #93c5fd !important; 
            margin: 2px 0 0 0; 
            font-size: 12px !important; 
        }
        
        /* Badge Fontes de Dados Compacto */
        .datasource-badge {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-left: 4px solid #002060;
            border-radius: 6px;
            padding: 8px 14px !important;
            margin-bottom: 12px !important;
            font-size: 12px !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.02);
        }
        
        /* Banners das Abas Operacionais */
        .section-banner {
            background: #ffffff;
            border-left: 4px solid #002060;
            padding: 8px 14px !important;
            border-radius: 6px;
            margin-bottom: 12px !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.02);
        }
        .section-banner h3 { margin: 0; color: #002060; font-size: 15px !important; font-weight: 700; }
        .section-banner p { margin: 2px 0 0 0; color: #64748b; font-size: 11px !important; }
        
        /* Métricas Compactas */
        div[data-testid="stMetric"] {
            background: #ffffff;
            border-radius: 8px;
            padding: 10px 14px !important;
            border-top: 3px solid #002060;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        }
        div[data-testid="stMetricLabel"] { font-size: 11px !important; color: #64748b; font-weight: 500; }
        div[data-testid="stMetricValue"] { font-size: 20px !important; font-weight: 700; color: #002060; }
        
        /* Botões Principais */
        div.stButton > button {
            background: linear-gradient(135deg, #002060 0%, #003366 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            font-size: 12px !important;
            padding: 6px 16px !important;
            box-shadow: 0 2px 6px rgba(0,32,96,0.15) !important;
            transition: all 0.2s ease-in-out !important;
        }
        div.stButton > button:hover {
            background: linear-gradient(135deg, #003366 0%, #004080 100%) !important;
            box-shadow: 0 4px 10px rgba(0,32,96,0.25) !important;
            transform: translateY(-1px);
        }
        
        /* Botões de Download */
        div[data-testid="stDownloadButton"] > button {
            background: #ffffff !important;
            color: #002060 !important;
            border: 1px solid #002060 !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            font-size: 12px !important;
            padding: 6px 14px !important;
        }
        div[data-testid="stDownloadButton"] > button:hover {
            background: #f1f5f9 !important;
            color: #001030 !important;
        }

        /* Tabela Dataframe */
        div[data-testid="stDataFrame"] {
            background-color: #ffffff;
            border-radius: 8px !important;
            padding: 4px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
            border: 1px solid #cbd5e1;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# ==============================================================================
# 2. CONEXÃO DIRETA COM SUPABASE & CACHE AUTOMÁTICO
# ==============================================================================
def get_db_engine():
    try:
        if "postgres" in st.secrets and "url" in st.secrets["postgres"]:
            db_url = st.secrets["postgres"]["url"]
            if db_url.startswith("postgresql://"):
                db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
            return create_engine(db_url, pool_pre_ping=True)
    except Exception as e:
        st.error(f"Erro na ligação ao banco de dados Supabase: {e}")
    return None

def sanitizar_bilhete(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"[\[\]]", "", s)
    s = re.sub(r"\.0$", "", s)
    return s.strip()

@st.cache_data(ttl=180, show_spinner="Carregando base unificada do Supabase...")
def carregar_dados_supabase_direto():
    engine = get_db_engine()
    if not engine:
        return pd.DataFrame()
    
    query = "SELECT * FROM vw_conciliacao_mestra;"
    try:
        df = pd.read_sql(query, engine)
        if not df.empty and "Bilhetes" in df.columns:
            df["Bilhete_Clean"] = df["Bilhetes"].apply(sanitizar_bilhete)
            df = df.drop_duplicates("Bilhete_Clean", keep="first").reset_index(drop=True)
        return df
    except Exception as e:
        st.warning(f"⚠️ Não foi possível ler a visão 'vw_conciliacao_mestra': {e}")
        return pd.DataFrame()

def carregar_tratativas_db():
    engine = get_db_engine()
    if not engine:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "data_modificacao"])
    
    query = "SELECT bilhete, status_geral, area_resp, obs_operacao, setor, gerentes, data_modificacao FROM tratativas"
    try:
        df_db = pd.read_sql(query, engine)
        if not df_db.empty and "bilhete" in df_db.columns:
            df_db["bilhete_clean"] = df_db["bilhete"].apply(sanitizar_bilhete)
            if "data_modificacao" in df_db.columns:
                df_db["data_modificacao"] = pd.to_datetime(df_db["data_modificacao"], errors="coerce")
                df_db = df_db.sort_values("data_modificacao", ascending=False)
            df_db = df_db.drop_duplicates("bilhete_clean", keep="first").reset_index(drop=True)
            df_db["bilhete"] = df_db["bilhete_clean"]
            df_db.drop(columns=["bilhete_clean"], inplace=True, errors="ignore")
        return df_db
    except Exception as e:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "data_modificacao"])

def salvar_tratativas_lote_supabase(df_lote, usuario):
    engine = get_db_engine()
    if not engine or df_lote.empty:
        return False, "Conexão com o Supabase indisponível.", 0, 0, 0
    
    df_existentes = carregar_tratativas_db()
    set_existentes = set(df_existentes["bilhete"].apply(sanitizar_bilhete).tolist()) if not df_existentes.empty else set()
    
    dados_lote = []
    qtd_atualizados, qtd_novos = 0, 0
    for _, r in df_lote.iterrows():
        b_clean = sanitizar_bilhete(r.get("Bilhetes", r.get("bilhete", "")))
        if b_clean and b_clean.lower() not in ["nan", "none", "", "-"]:
            obs_limpa = str(r.get("Obs. Operação", r.get("obs_operacao", ""))).replace("Sem tratativa na operação", "").strip()
            st_val = str(r.get("Status_Geral", r.get("status_geral", ""))).strip()
            ar_val = str(r.get("Área Resp. Operação", r.get("area_resp", ""))).strip()
            setor_val = str(r.get("Setor", r.get("setor", "-"))).strip()
            gerente_val = str(r.get("Gerentes", r.get("gerentes", ar_val))).strip()
            
            if st_val in ["-", "", "None", "nan"]:
                st_val = "Pendente de Lançamento (Não Consta)"
            if ar_val in ["-", "", "None", "nan"]:
                ar_val = "Operação"
            if gerente_val in ["-", "", "None", "nan"]:
                gerente_val = ar_val
                
            if b_clean in set_existentes:
                qtd_atualizados += 1
            else:
                qtd_novos += 1
                
            dados_lote.append({
                "bilhete": b_clean,
                "status_geral": st_val,
                "area_resp": ar_val,
                "obs_operacao": obs_limpa,
                "setor": setor_val,
                "gerentes": gerente_val,
                "usuario": str(usuario)
            })
            
    if not dados_lote:
        return False, "Nenhum bilhete válido para atualização.", 0, 0, 0
        
    sql_upsert = text("""
        INSERT INTO tratativas (bilhete, status_geral, area_resp, obs_operacao, setor, gerentes, usuario_modificacao, data_modificacao)
        VALUES (:bilhete, :status_geral, :area_resp, :obs_operacao, :setor, :gerentes, :usuario, NOW())
        ON CONFLICT (bilhete) DO UPDATE SET
            status_geral = EXCLUDED.status_geral,
            area_resp = EXCLUDED.area_resp,
            obs_operacao = EXCLUDED.obs_operacao,
            setor = EXCLUDED.setor,
            gerentes = EXCLUDED.gerentes,
            usuario_modificacao = EXCLUDED.usuario_modificacao,
            data_modificacao = NOW();
    """)
    try:
        with engine.begin() as conn:
            conn.execute(sql_upsert, dados_lote)
        st.cache_data.clear()
        msg = f"✅ Sincronização concluída! {qtd_atualizados} atualizado(s) e {qtd_novos} novo(s) inserido(s) no Supabase."
        return True, msg, qtd_atualizados, qtd_novos, 0
    except Exception as e:
        return False, f"Erro ao salvar tratativas: {e}", 0, 0, len(dados_lote)

def registrar_log_supabase(logs_list):
    engine = get_db_engine()
    if not engine or not logs_list:
        return
    df_logs = pd.DataFrame(logs_list)
    df_logs.rename(columns={
        "Data_Hora": "data_hora", "Bilhete": "bilhete", "Usuario_Acao": "usuario_action",
        "Status_Anterior": "status_anterior", "Novo_Status": "novo_status",
        "Area_Anterior": "area_anterior", "Nova_Area": "nova_area",
        "Observacao": "observacao", "Tipo_Interacao": "tipo_interacao"
    }, inplace=False)
    try:
        df_logs.to_sql("log_auditoria", engine, if_exists="append", index=False)
    except Exception:
        pass

# ==============================================================================
# 3. GESTÃO DE USUÁRIOS & AUTENTICAÇÃO
# ==============================================================================
def carregar_usuarios_supabase():
    dict_u = {}
    if os.path.exists("usuarios_autorizados.json"):
        try:
            with open("usuarios_autorizados.json", "r", encoding="utf-8") as f:
                data_json = json.load(f)
                for k, v in data_json.items():
                    dict_u[str(k).strip().lower()] = {
                        "senha": str(v.get("senha", "")).strip(),
                        "nome": str(v.get("nome", "")).strip(),
                        "perfil": str(v.get("perfil", "Operacao")).strip(),
                        "status": str(v.get("status", "APROVADO")).strip().upper(),
                    }
        except Exception:
            pass
            
    engine = get_db_engine()
    if engine:
        try:
            df_u = pd.read_sql("SELECT usuario, senha, nome, perfil, status FROM usuarios", engine)
            for _, row in df_u.iterrows():
                u_key = str(row["usuario"]).strip().lower()
                dict_u[u_key] = {
                    "senha": str(row["senha"]).strip(),
                    "nome": str(row["nome"]).strip(),
                    "perfil": str(row["perfil"]).strip(),
                    "status": str(row["status"]).strip().upper(),
                }
        except Exception:
            pass
    return dict_u

def salvar_usuario_supabase(u_id, senha, nome, perfil, status="PENDENTE"):
    engine = get_db_engine()
    if not engine:
        return False, "Conexão com a base de dados indisponível."
    sql = text("""
        INSERT INTO usuarios (usuario, senha, nome, perfil, status, data_solicitacao)
        VALUES (:u, :p, :n, :perf, :st, CURRENT_DATE)
        ON CONFLICT (usuario) DO UPDATE SET
            senha = EXCLUDED.senha, nome = EXCLUDED.nome, perfil = EXCLUDED.perfil, status = EXCLUDED.status;
    """)
    try:
        with engine.begin() as conn:
            conn.execute(sql, {"u": u_id, "p": senha, "n": nome, "perf": perfil, "st": status})
        return True, "Operação realizada com sucesso."
    except Exception as e:
        return False, f"Erro no banco: {e}"

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
    return u_id in ["mribeiro", "fellipe", "ffernandes"] or perfil in ["Master", "Compliance"]

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown(
            """
            <div class='main-header'>
                <div>
                    <h1>✈️ Grupo Arbaitman</h1>
                    <p>Portal de Conciliação Aérea & Gestão de Bilhetes Pendentes (FP&A)</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        tab_log, tab_req = st.tabs(["🔐 Entrar no Portal", "📝 Solicitar Acesso"])
        with tab_log:
            u_input = st.text_input("Usuário:", key="l_user").strip().lower()
            p_input = str(st.text_input("Senha:", type="password", key="l_pass")).strip()
            if st.button("Acessar Portal", type="primary"):
                usuarios_db = carregar_usuarios_supabase()
                if u_input in usuarios_db and usuarios_db[u_input]["senha"] == p_input:
                    if usuarios_db[u_input].get("status") == "APROVADO":
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_atual"] = usuarios_db[u_input]["nome"]
                        st.session_state["perfil_atual"] = usuarios_db[u_input]["perfil"]
                        st.session_state["login_user_id"] = u_input
                        st.rerun()
                    else:
                        st.warning("⏳ Seu usuário está aguardando aprovação do Gestor Master.")
                elif u_input == "mribeiro" and p_input in ["14052013", "123"]:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = "Marcos Ribeiro"
                    st.session_state["perfil_atual"] = "Master"
                    st.session_state["login_user_id"] = "mribeiro"
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        with tab_req:
            r_nome = st.text_input("Nome Completo:")
            r_user = st.text_input("Usuário Desejado:").strip().lower()
            r_pass = str(st.text_input("Senha de Acesso:", type="password")).strip()
            if st.button("Enviar Solicitação"):
                if r_user and r_pass and r_nome:
                    ok, msg = salvar_usuario_supabase(r_user, r_pass, r_nome, "Operacao", status="PENDENTE")
                    if ok:
                        st.success("✅ Solicitação enviada!")
                    else:
                        st.error(msg)
    st.stop()

# ==============================================================================
# 4. ROTEAMENTO DE BASES & PROCESSAMENTO
# ==============================================================================
def clean_str(val):
    if pd.isna(val) or val is None: 
        return "-"
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    if s.lower() in ["none", "nan", "null", "<na>", ""]:
        return "-"
    return s

def padronizar_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Bilhetes", "Ponto de venda", "Status_Geral", "Área Resp. Operação",
            "Obs. Operação", "Setor", "Data Emissão", "CIA", "Taxa", "A vista", "A credito",
            "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao", "Emissor", 
            "Consultor", "Status_Sistema", "Status_Divergencia", "Gerentes", "Fornecedor_Sistema", "Fonte_Relatorio"
        ])
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    for c in df_out.columns:
        if df_out[c].dtype == object or df_out[c].dtype == 'string':
            df_out[c] = df_out[c].apply(clean_str)
    
    if "CIA" in df_out.columns:
        df_out["CIA"] = df_out["CIA"].astype(str).str.upper().str.strip().replace(DEPARA_CIAS)
        
    for c in ["Taxa", "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa", "Receita_Sistema", "Dif_Receita", "Comissão", "Taxa DU", "Incentivo"]:
        if c in df_out.columns:
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)
            
    df_out["Dt_Parsed"] = pd.to_datetime(df_out["Data Emissão"], format="mixed", dayfirst=True, errors="coerce")
    return df_out

def rotear_bases_mestra(df_master):
    if df_master is None or df_master.empty:
        vazio = padronizar_df(None)
        return vazio, vazio, vazio, vazio, vazio, vazio
        
    df_master = padronizar_df(df_master)
    
    st_geral_str = df_master["Status_Geral"].astype(str).str.lower().str.strip()
    st_div_str = df_master["Status_Divergencia"].astype(str).str.lower().str.strip()
    
    termos_ok = ["já lançado", "ja lancado", "emitido e lançado", "emitido e lancado", "lançado", "lancado", "conciliado", "valores corretos", "regularizado", "ok"]
    
    mask_ok = st_geral_str.apply(lambda s: any(t in s for t in termos_ok)) | st_div_str.str.contains("valores corretos|sem divergência", na=False)
    
    df_ok = df_master[mask_ok].copy()
    df_rest = df_master[~mask_ok].copy()
    
    def e_bo(r):
        ar_val = str(r.get("Área Resp. Operação", "")).strip().lower()
        obs_val = str(r.get("Obs. Operação", "")).strip().lower()
        return any(k in ar_val or k in obs_val for k in ["katia", "kátia", "backoffice", "suporte"])
        
    mask_bo = df_rest.apply(e_bo, axis=1)
    df_bo = df_rest[mask_bo].copy()
    df_rest = df_rest[~mask_bo].copy()
    
    def e_eventos(r):
        ar_val = str(r.get("Área Resp. Operação", "")).strip().lower()
        return "eventos" in ar_val
        
    mask_evt = df_rest.apply(e_eventos, axis=1)
    df_eventos = df_rest[mask_evt].copy()
    df_rest = df_rest[~mask_evt].copy()
    
    def e_falta(r):
        st_val = str(r.get("Status_Geral", "")).strip().lower()
        return "falta de lançamento" in st_val or "não consta" in st_val
        
    mask_falta = df_rest.apply(e_falta, axis=1)
    df_falta = df_rest[mask_falta].copy()
    df_erros = df_rest[~mask_falta].copy()
    
    df_lemontech = pd.DataFrame()
    return df_falta, df_erros, df_bo, df_eventos, df_lemontech, df_ok

# ==============================================================================
# 5. SIDEBAR & FILTROS
# ==============================================================================
st.sidebar.title("Grupo Arbaitman")
st.sidebar.caption("Conciliação Aérea FP&A v3.7 (Integração Supabase Direct)")
st.sidebar.write(f"👤 **{st.session_state['usuario_atual']}** ({st.session_state['perfil_atual']})")

if st.sidebar.button("🔄 Sincronizar / Atualizar"):
    st.cache_data.clear()
    st.rerun()

if st.sidebar.button("🔒 Sair"):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")

df_mestra = carregar_dados_supabase_direto()
df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div = rotear_bases_mestra(df_mestra)

st.sidebar.subheader("🔍 Filtros Operacionais")
opts_fonte_rel = sorted(df_mestra["Fonte_Relatorio"].dropna().unique()) if "Fonte_Relatorio" in df_mestra.columns else []
filtro_fonte_rel = st.sidebar.multiselect("Fonte do Relatório:", options=opts_fonte_rel, placeholder="Todas")
filtro_gerente = st.sidebar.multiselect("Gerente / Área Resp.:", options=sorted(df_mestra["Área Resp. Operação"].dropna().unique()) if "Área Resp. Operação" in df_mestra.columns else [], placeholder="Todos")
filtro_cia = st.sidebar.multiselect("Companhia Aérea:", options=sorted(df_mestra["CIA"].dropna().unique()) if "CIA" in df_mestra.columns else [], placeholder="Todas")

def aplicar_filtros(df):
    if df is None or df.empty:
        return df
    m = pd.Series(True, index=df.index)
    if filtro_fonte_rel and "Fonte_Relatorio" in df.columns:
        m = m & df["Fonte_Relatorio"].isin(filtro_fonte_rel)
    if filtro_gerente and "Área Resp. Operação" in df.columns:
        m = m & df["Área Resp. Operação"].isin(filtro_gerente)
    if filtro_cia and "CIA" in df.columns:
        m = m & df["CIA"].isin(filtro_cia)
    return df[m].reset_index(drop=True)

f_falta = aplicar_filtros(df_falta)
f_erros = aplicar_filtros(df_erros)
f_backoffice = aplicar_filtros(df_backoffice)
f_eventos = aplicar_filtros(df_eventos)
f_lemon_virt = aplicar_filtros(df_lemon_virt)
f_sem_div = aplicar_filtros(df_sem_div)

# ==============================================================================
# 6. HEADER PRINCIPAL E METRICAS COMPACTAS
# ==============================================================================
st.markdown(
    """
    <div class='main-header'>
        <div>
            <h1>✈️ Grupo Arbaitman | Portal de Conciliação Aérea (FP&A)</h1>
            <p>Leitura Direta da Visão Unificada 'vw_conciliacao_mestra' no Supabase PostgreSQL</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Pendentes ERP", f"{len(f_falta):,}")
c2.metric("Erros Val/CIA", f"{len(f_erros):,}")
c3.metric("Backoffice", f"{len(f_backoffice):,}")
c4.metric("Central Eventos", f"{len(f_eventos):,}")
c5.metric("Emissor Virtual", f"{len(f_lemon_virt):,}")
c6.metric("Conciliados (OK)", f"{len(f_sem_div):,}")

st.markdown("---")

abas = [
    "📊 Dashboard Executivo",
    "📋 1. Falta de Lançamento (Maringá)",
    "⚠️ 2. Erros de Valores & CIA (Maringá)",
    "🎧 3. Suporte Backoffice",
    "🎪 4. Central de Eventos",
    "🤖 5. Emissor Virtual Lemontech",
    "✅ 6. Bilhetes Conciliados (Sem Divergência)"
]
if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Histórico"])

aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# DASHBOARD EXECUTIVO
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.markdown("<div class='section-banner'><h3>📊 Painel Executivo C-Level FP&A</h3></div>", unsafe_allow_html=True)
    df_pendentes_todas = pd.concat([f_falta, f_erros, f_backoffice, f_eventos, f_lemon_virt], ignore_index=True)
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        if not df_pendentes_todas.empty and "Status_Geral" in df_pendentes_todas.columns:
            fig_pie = px.pie(df_pendentes_todas["Status_Geral"].value_counts().reset_index(), values="count", names="Status_Geral", title="Pendências por Status")
            fig_pie.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_pie, use_container_width=True)
    with col_d2:
        if not df_pendentes_todas.empty and "CIA" in df_pendentes_todas.columns:
            fig_cia = px.bar(df_pendentes_todas["CIA"].value_counts().head(8).reset_index(), x="CIA", y="count", title="Top 8 CIAs com Pendências")
            fig_cia.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_cia, use_container_width=True)

# ------------------------------------------------------------------------------
# MÓDULO DE TRATATIVA DIRETA LINHA A LINHA
# ------------------------------------------------------------------------------
def renderizar_modulo_tratativa(df_filtrado, nome_base, key_prefix):
    if df_filtrado.empty:
        st.info("Nenhum bilhete pendente nesta categoria.")
        return
        
    termo_busca = st.text_input("🔍 Buscar nesta tela (Bilhete, LOC, Cliente, Gerente):", key=f"src_{key_prefix}")
    df_exib = df_filtrado.copy()
    if termo_busca.strip():
        term = termo_busca.strip().lower()
        cols_s = [c for c in ["Bilhetes", "Localizador_Sistema", "Ponto de venda", "Área Resp. Operação", "CIA"] if c in df_exib.columns]
        m_s = pd.Series(False, index=df_exib.index)
        for col in cols_s:
            m_s |= df_exib[col].astype(str).str.lower().str.contains(term, na=False)
        df_exib = df_exib[m_s]
        
    st.markdown(f"**Registros Visíveis:** {len(df_exib)}")
    bilhetes_lista = df_exib["Bilhetes"].tolist() if "Bilhetes" in df_exib.columns else []
    
    bilhet_sel = st.multiselect("Selecione Bilhetes para Atualizar:", options=bilhetes_lista, key=f"ms_{key_prefix}")
    
    if bilhet_sel:
        df_sel = df_exib[df_exib["Bilhetes"].isin(bilhet_sel)]
        df_pr = df_sel.iloc[0]
        with st.form(f"form_trat_{key_prefix}"):
            c1, c2 = st.columns(2)
            with c1:
                n_status = st.selectbox("Novo Status:", ["Já Lançado no ERP", "Emitido e Lançado", "Pendente de Lançamento (Não Consta)", "Em Análise Backoffice", "Divergência de Tarifa", "Cancelado / Reembolsado"], key=f"st_{key_prefix}")
            with c2:
                n_area = st.text_input("Nova Área Responsável / Gerente:", value=str(df_pr.get("Área Resp. Operação", "Operação")), key=f"ar_{key_prefix}")
            n_obs = st.text_area("Observação / Justificativa:", key=f"obs_{key_prefix}")
            
            if st.form_submit_button("💾 Salvar Tratativa no Supabase"):
                usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                df_up = pd.DataFrame({
                    "Bilhetes": df_sel["Bilhetes"].values,
                    "Status_Geral": [n_status] * len(df_sel),
                    "Área Resp. Operação": [n_area] * len(df_sel),
                    "Obs. Operação": [n_obs] * len(df_sel),
                    "Setor": df_sel["Setor"].values if "Setor" in df_sel.columns else ["Operação"] * len(df_sel),
                    "Gerentes": [n_area] * len(df_sel)
                })
                ok, msg, _, _, _ = salvar_tratativas_lote_supabase(df_up, usuario=usr_str)
                if ok:
                    st.success(msg)
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg)
                    
    st.dataframe(df_exib, use_container_width=True, hide_index=True)

with aba_sel[1]:
    renderizar_modulo_tratativa(f_falta, "Falta_de_Lancamento", "fl")
with aba_sel[2]:
    renderizar_modulo_tratativa(f_erros, "Erros_Valores_CIA", "ev")
with aba_sel[3]:
    renderizar_modulo_tratativa(f_backoffice, "Suporte_Backoffice", "sb")
with aba_sel[4]:
    renderizar_modulo_tratativa(f_eventos, "Central_de_Eventos", "ce")
with aba_sel[5]:
    renderizar_modulo_tratativa(f_lemon_virt, "Emissor_Virtual_Lemontech", "evl")
with aba_sel[6]:
    st.dataframe(f_sem_div, use_container_width=True, hide_index=True)