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
# 1. CONFIGURAÇÃO INICIAL E ESTILOS CSS EXECUTIVOS
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Portal de Conciliação Aérea (FP&A)",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_DASHBOARD_PADRAO = "Dashboard_Revenue_Assurance_Consolidado.xlsx"

# Dicionário de Mapeamento das Companhias Aéreas
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

# Dicionário de Padronização de Nomes de Gerentes
DEPARA_GERENTES = {
    "FABIANO SOUZA": "Fabiano Souza", "FABIANO": "Fabiano Souza",
    "GUILHERME SILVA": "Guilherme Silva", "GUILHERME": "Guilherme Silva",
    "IVANETE BERTASOL": "Ivanete Bertasol", "IVANETE": "Ivanete Bertasol",
    "KELI SANTI": "Keli Santi", "KELI": "Keli Santi",
    "MARCELO PEREIRA": "Marcelo Pereira", "MARCELO": "Marcelo Pereira",
    "ROSANGELA PALLU": "Rosangela Pallu", "ROSANGELA": "Rosangela Pallu",
    "JAIME SCHNAIDER": "Jaime Schnaider", "JAIME": "Jaime Schnaider",
    "SILVANA CELANI": "Silvana Celani", "SILVANA": "Silvana Celani",
    "KATIA MARTINS": "Katia Martins", "KÁTIA MARTINS": "Katia Martins",
    "KATIA": "Katia Martins", "KÁTIA": "Katia Martins",
    "SUPORTE BACKOFFICE": "Suporte Backoffice", "BACKOFFICE": "Suporte Backoffice",
    "CENTRAL DE EVENTOS": "Central de Eventos", "EVENTOS": "Central de Eventos"
}

def padronizar_nome_gerente(nome_raw):
    """Unifica variações de nomes de gerentes (ex: GUILHERME SILVA, Guilherme -> Guilherme Silva)"""
    if not nome_raw or str(nome_raw).strip().lower() in ["nan", "none", "", "-", "não mapeado", "nao mapeado"]:
        return "Não Mapeado"
    
    s = str(nome_raw).strip().upper()
    s_clean = re.sub(r"\s+", " ", s)
    
    for k, v in DEPARA_GERENTES.items():
        if k == s_clean or k in s_clean:
            return v
            
    return s_clean.title()

def consolidar_emissor_consultor(row):
    """Cria uma string unificada de Emissor e Consultor sem duplicar nomes"""
    emis = str(row.get("Emissor") or row.get("Emissor_Reserva_Lemon") or "").strip()
    cons = str(row.get("Consultor") or row.get("Consultor_Lemon") or "").strip()
    
    elementos = []
    for val in [emis, cons]:
        val_clean = val.replace("Sem consultor", "").replace("Sem emissor", "").strip()
        if val_clean and val_clean.lower() not in ["nan", "none", "", "-", "null"]:
            if val_clean not in elementos:
                elementos.append(val_clean)
                
    if elementos:
        return " / ".join(elementos)
    return "-"

# Injeção de CSS
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        * { font-family: 'Inter', sans-serif; }
        .stApp { background-color: #f8fafc; }
        
        .main-header {
            background: linear-gradient(135deg, #001e4d 0%, #003366 50%, #004080 100%);
            padding: 22px 30px;
            border-radius: 12px;
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 8px 20px rgba(0, 32, 96, 0.2);
        }
        .main-header h1 { color: #ffffff !important; margin: 0; font-size: 24px; font-weight: 700; }
        .main-header p { color: #b3d1ff !important; margin: 4px 0 0 0; font-size: 13px; }
        
        .section-banner {
            background: #ffffff;
            border-left: 5px solid #002060;
            padding: 14px 20px;
            border-radius: 8px;
            margin-bottom: 18px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .section-banner h3 { margin: 0; color: #002060; font-size: 18px; font-weight: 700; }
        .section-banner p { margin: 2px 0 0 0; color: #64748b; font-size: 12px; }
        
        .chart-card {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 18px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
            border: 1px solid #e2e8f0;
            margin-bottom: 18px;
        }
        .chart-card h5 {
            color: #002060;
            font-weight: 700;
            font-size: 15px;
            margin-top: 0;
            margin-bottom: 12px;
        }
        
        div[data-testid="stMetric"] {
            background: #ffffff;
            border-radius: 8px;
            padding: 14px;
            border-top: 4px solid #002060;
            box-shadow: 0 2px 10px rgba(0,0,0,0.03);
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# 2. FUNÇÕES DE SUPABASE E CONEXÃO DE DADOS
# ==============================================================================
def sanitizar_bilhete(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"[\[\]]", "", s)
    s = re.sub(r"\.0$", "", s)
    return s.strip()

def get_db_engine():
    try:
        if "postgres" in st.secrets and "url" in st.secrets["postgres"]:
            db_url = st.secrets["postgres"]["url"]
            if db_url.startswith("postgresql://"):
                db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
            return create_engine(db_url, pool_pre_ping=True)
    except Exception as e:
        st.error(f"Erro na ligação ao banco de dados: {e}")
    return None

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
    except Exception:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "data_modificacao"])

def salvar_tratativas_lote_supabase(df_lote, usuario):
    engine = get_db_engine()
    if not engine or df_lote.empty:
        return False, "Conexão com o banco de dados indisponível.", 0, 0, 0
    df_existentes = carregar_tratativas_db()
    set_existentes = set(df_existentes["bilhete"].apply(sanitizar_bilhete).tolist()) if not df_existentes.empty else set()
    dados_lote = []
    qtd_atualizados = 0
    qtd_novos = 0
    
    for _, r in df_lote.iterrows():
        b_clean = sanitizar_bilhete(r.get("Bilhetes", r.get("bilhete", "")))
        if b_clean and b_clean.lower() not in ["nan", "none", "", "-"]:
            obs_limpa = str(r.get("Obs. Operação", r.get("obs_operacao", ""))).replace("Sem tratativa na operação", "").strip()
            st_val = str(r.get("Status_Geral", r.get("status_geral", ""))).strip()
            ar_val = padronizar_nome_gerente(r.get("Área Resp. Operação", r.get("area_resp", "")))
            setor_val = str(r.get("Setor", r.get("setor", "-"))).strip()
            gerente_val = padronizar_nome_gerente(r.get("Gerentes", r.get("gerentes", ar_val)))
            
            if st_val in ["-", "", "None", "nan"]:
                st_val = "Pendente de Lançamento (Não Consta)"
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
        return True, f"✅ Sincronizado com sucesso! {qtd_atualizados} atualizados e {qtd_novos} novos.", qtd_atualizados, qtd_novos, 0
    except Exception as e:
        return False, f"Erro ao salvar tratativas: {e}", 0, 0, len(dados_lote)

def registrar_log_supabase(logs_list):
    engine = get_db_engine()
    if not engine or not logs_list:
        return
    df_logs = pd.DataFrame(logs_list)
    df_logs.rename(columns={
        "Data_Hora": "data_hora", "Bilhete": "bilhete", "Usuario_Acao": "usuario_acao",
        "Status_Anterior": "status_anterior", "Novo_Status": "novo_status",
        "Area_Anterior": "area_anterior", "Nova_Area": "nova_area",
        "Observacao": "observacao", "Tipo_Interacao": "tipo_interacao"
    }, inplace=True, errors="ignore")
    try:
        with engine.begin() as conn:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS log_auditoria (
                    id SERIAL PRIMARY KEY, data_hora TIMESTAMP, bilhete VARCHAR(255),
                    usuario_acao VARCHAR(255), status_anterior VARCHAR(255), novo_status VARCHAR(255),
                    area_anterior VARCHAR(255), nova_area VARCHAR(255), observacao TEXT, tipo_interacao VARCHAR(255)
                );
            """))
        df_logs.to_sql("log_auditoria", engine, if_exists="append", index=False)
    except Exception:
        pass

# ==============================================================================
# 3. AUTENTICAÇÃO E USUÁRIOS
# ==============================================================================
def carregar_usuarios_supabase():
    dict_u = {}
    engine = get_db_engine()
    if engine:
        try:
            df_u = pd.read_sql("SELECT usuario, senha, nome, perfil, status FROM usuarios", engine)
            for _, row in df_u.iterrows():
                u_key = str(row["usuario"]).strip().lower()
                dict_u[u_key] = {
                    "senha": str(row["senha"]).strip(), "nome": str(row["nome"]).strip(),
                    "perfil": str(row["perfil"]).strip(), "status": str(row["status"]).strip().upper()
                }
        except Exception:
            pass
    return dict_u

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
                <h1>✈️ Grupo Arbaitman</h1>
                <p>Portal de Conciliação Aérea & Gestão de Bilhetes Pendentes (FP&A)</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        u_input = st.text_input("Usuário:", key="l_user").strip().lower()
        p_input = str(st.text_input("Senha:", type="password", key="l_pass")).strip()
        
        if st.button("Acessar Portal", type="primary"):
            usuarios_db = carregar_usuarios_supabase()
            if u_input in usuarios_db and usuarios_db[u_input]["senha"] == p_input:
                st.session_state["autenticado"] = True
                st.session_state["usuario_atual"] = usuarios_db[u_input]["nome"]
                st.session_state["perfil_atual"] = usuarios_db[u_input]["perfil"]
                st.session_state["login_user_id"] = u_input
                st.rerun()
            elif u_input == "mribeiro" and p_input in ["14052013", "123"]:
                st.session_state["autenticado"] = True
                st.session_state["usuario_atual"] = "Marcos Ribeiro"
                st.session_state["perfil_atual"] = "Master"
                st.session_state["login_user_id"] = "mribeiro"
                st.rerun()
            elif u_input in ["fellipe", "ffernandes"] and p_input == "123":
                st.session_state["autenticado"] = True
                st.session_state["usuario_atual"] = "Fellipe Fernandes"
                st.session_state["perfil_atual"] = "Master"
                st.session_state["login_user_id"] = u_input
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
    st.stop()

# ==============================================================================
# 4. PADRONIZAÇÃO DE DADOS E ROTEADOR DE ABAS
# ==============================================================================
def clean_str(val):
    if pd.isna(val) or val is None: return "-"
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    if s.lower() in ["none", "nan", "null", "<na>", ""]: return "-"
    return s

def padronizar_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
        
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    
    for c in df_out.columns:
        if df_out[c].dtype == object or df_out[c].dtype == 'string':
            df_out[c] = df_out[c].apply(clean_str)
            
    if "Obs. Operação" in df_out.columns:
        df_out["Obs. Operação"] = df_out["Obs. Operação"].astype(str).str.replace("Sem tratativa na operação", "", regex=False).str.strip()
        
    # Padronização de Gerente/Área Responsável
    for col_g in ["Área Resp. Operação", "Gerentes"]:
        if col_g in df_out.columns:
            df_out[col_g] = df_out[col_g].apply(padronizar_nome_gerente)
            
    # Criação da Coluna Consolidada 'Emissor / Consultor'
    df_out["Emissor / Consultor"] = df_out.apply(consolidar_emissor_consultor, axis=1)
    
    if "CIA" in df_out.columns:
        df_out["CIA"] = df_out["CIA"].astype(str).str.upper().str.strip().replace(DEPARA_CIAS)
        
    for c in ["Taxa", "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa"]:
        if c in df_out.columns:
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)
            
    df_out["Dt_Parsed"] = pd.to_datetime(df_out.get("Data Emissão"), format="mixed", dayfirst=True, errors="coerce")
    return df_out

def mesclar_com_supabase(df_excel):
    if df_excel is None or df_excel.empty: return df_excel
    df_db = carregar_tratativas_db()
    if df_db.empty: return df_excel
    
    df_excel["Bilhete_Clean"] = df_excel["Bilhetes"].apply(sanitizar_bilhete)
    df_db["Bilhete_Clean"] = df_db["bilhete"].apply(sanitizar_bilhete)
    
    if "data_modificacao" in df_db.columns:
        df_db = df_db.sort_values("data_modificacao", ascending=False)
    df_db = df_db.drop_duplicates("Bilhete_Clean", keep="first")
    
    cols_sup_desejadas = ["status_geral", "area_resp", "obs_operacao", "setor", "gerentes"]
    cols_presentes = [c for c in cols_sup_desejadas if c in df_db.columns]
    
    df_merged = pd.merge(df_excel, df_db[["Bilhete_Clean"] + cols_presentes], on="Bilhete_Clean", how="left")
    
    muta_map = [
        ("Status_Geral", "status_geral"), ("Área Resp. Operação", "area_resp"),
        ("Obs. Operação", "obs_operacao"), ("Setor", "setor"), ("Gerentes", "gerentes")
    ]
    for col_excel, col_sb in muta_map:
        if col_sb in df_merged.columns and col_excel in df_merged.columns:
            mask_valido = df_merged[col_sb].notna() & ~df_merged[col_sb].astype(str).str.strip().isin(["-", "", "None", "nan"])
            df_merged[col_excel] = df_merged[col_sb].where(mask_valido, df_merged[col_excel])
            
    df_merged = df_merged.drop(columns=[c for c in cols_presentes if c in df_merged.columns])
    return df_merged

def rotear_bases_mestra(df_master):
    """
    Roteador estrito ajustado conforme especificação:
    1. Suporte Backoffice: 100% dos casos com setor/gerente 'Suporte Backoffice' OU menção a chamado aberto na obs.
    2. Central de Eventos: setor/gerente 'Central de Eventos'.
    3. Emissor Virtual Lemontech: Consultor/Emissor com 'Emissor Virtual Lemontech'.
    4. Falta de Lançamento (Maringá): Status não consta / pendente.
    5. Erros de Valores & CIA (Maringá): Divergências financeiras do ERP.
    6. Bilhetes Solicitados Conciliação: Conciliados (OK).
    """
    if df_master is None or df_master.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_master = padronizar_df(df_master)
    
    st_geral_str = df_master["Status_Geral"].astype(str).str.lower().str.strip()
    st_div_str = df_master.get("Status_Divergencia", pd.Series("")).astype(str).str.lower().str.strip()
    
    mask_ok = st_geral_str.str.contains("já lançado|emitido e lançado|lançado|conciliado|valores corretos|regularizado", na=False) | \
              st_div_str.str.contains("valores corretos|sem divergência", na=False)
              
    df_ok = df_master[mask_ok].copy()
    df_pend = df_master[~mask_ok].copy()
    
    # 1. SUPORTE BACKOFFICE (REGRA ESTRITA 100% PRIORITÁRIA)
    def e_suporte_backoffice(r):
        ar_val = str(r.get("Área Resp. Operação", "")).lower()
        ger_val = str(r.get("Gerentes", "")).lower()
        obs_val = str(r.get("Obs. Operação", "")).lower()
        st_val = str(r.get("Status_Geral", "")).lower()
        
        tem_area_bo = "backoffice" in ar_val or "backoffice" in ger_val or "katia" in ar_val or "kátia" in ar_val
        tem_obs_bo = any(k in obs_val or k in st_val for k in ["chamado", "ticket", "suporte", "backoffice", "análise backoffice", "analise backoffice"])
        return tem_area_bo or tem_obs_bo

    mask_bo = df_pend.apply(e_suporte_backoffice, axis=1)
    df_bo = df_pend[mask_bo].copy()
    df_rest = df_pend[~mask_bo].copy()
    
    # 2. CENTRAL DE EVENTOS
    def e_central_eventos(r):
        ar_val = str(r.get("Área Resp. Operação", "")).lower()
        ger_val = str(r.get("Gerentes", "")).lower()
        return "eventos" in ar_val or "eventos" in ger_val

    mask_evt = df_rest.apply(e_central_eventos, axis=1)
    df_eventos = df_rest[mask_evt].copy()
    df_rest = df_rest[~mask_evt].copy()
    
    # 3. EMISSOR VIRTUAL LEMONTECH
    def e_emissor_virtual(r):
        emis_val = str(r.get("Emissor / Consultor", "")).lower()
        ponto_val = str(r.get("Ponto de venda", "")).lower()
        return "virtual" in emis_val or "lemontech" in emis_val or "virtual" in ponto_val

    mask_lemon = df_rest.apply(e_emissor_virtual, axis=1)
    df_lemon_virt = df_rest[mask_lemon].copy()
    df_rest = df_rest[~mask_lemon].copy()
    
    # 4. FALTA DE LANÇAMENTO (MARINGÁ)
    def e_falta_lancamento(r):
        st_val = str(r.get("Status_Geral", "")).lower()
        return "falta de lançamento" in st_val or "não consta" in st_val or "nao consta" in st_val or st_val in ["-", "", "none"]

    mask_falta = df_rest.apply(e_falta_lancamento, axis=1)
    df_falta = df_rest[mask_falta].copy()
    df_erros = df_rest[~mask_falta].copy()
    
    return df_falta, df_erros, df_bo, df_eventos, df_lemon_virt, df_ok

@st.cache_data(ttl=300, show_spinner=False)
def carregar_bases():
    if os.path.exists(ARQUIVO_DASHBOARD_PADRAO):
        fonte = ARQUIVO_DASHBOARD_PADRAO
    else:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    try:
        xls = pd.ExcelFile(fonte, engine="openpyxl")
        frames = []
        for sheet in xls.sheet_names:
            if sheet.startswith("99_") or sheet.startswith("98_"):
                df_sheet = pd.read_excel(xls, sheet_name=sheet)
                frames.append(df_sheet)
                
        if not frames: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        df_m = pd.concat(frames, ignore_index=True)
        if "Bilhetes" in df_m.columns:
            df_m["Bilhete_Clean"] = df_m["Bilhetes"].apply(sanitizar_bilhete)
            df_m = df_m[df_m["Bilhete_Clean"].astype(str).str.len() > 3]
            df_m = df_m.drop_duplicates(subset=["Bilhete_Clean"], keep="first").reset_index(drop=True)
            
        df_m = mesclar_com_supabase(df_m)
        return rotear_bases_mestra(df_m)
    except Exception as e:
        st.error(f"Erro ao carregar base: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ==============================================================================
# 5. CARREGAMENTO E FILTROS OPERACIONAIS
# ==============================================================================
df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div = carregar_bases()
df_todos_unificados = pd.concat([df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div], ignore_index=True)

st.sidebar.title("Grupo Arbaitman")
st.sidebar.caption("Conciliação Aérea FP&A v3.8")
st.sidebar.write(f"👤 **{st.session_state['usuario_atual']}** ({st.session_state['perfil_atual']})")

if st.sidebar.button("🔄 Recarregar Dados"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Filtros Operacionais")

opts_gerentes = sorted(list(set(df_todos_unificados["Área Resp. Operação"].dropna().unique()))) if not df_todos_unificados.empty else []
filtro_gerente = st.sidebar.multiselect("Gerente Responsável:", options=opts_gerentes, placeholder="Todos os Gerentes")

opts_cias = sorted(list(set(df_todos_unificados["CIA"].dropna().unique()))) if not df_todos_unificados.empty else []
filtro_cia = st.sidebar.multiselect("Companhia Aérea:", options=opts_cias, placeholder="Todas as CIAs")

def aplicar_filtros(df):
    if df is None or df.empty: return df
    m = pd.Series(True, index=df.index)
    if filtro_gerente:
        m = m & df["Área Resp. Operação"].isin(filtro_gerente)
    if filtro_cia:
        m = m & df["CIA"].isin(filtro_cia)
    return df[m].reset_index(drop=True)

f_falta = aplicar_filtros(df_falta)
f_erros = aplicar_filtros(df_erros)
f_backoffice = aplicar_filtros(df_backoffice)
f_eventos = aplicar_filtros(df_eventos)
f_lemon_virt = aplicar_filtros(df_lemon_virt)
f_sem_div = aplicar_filtros(df_sem_div)

# ==============================================================================
# 6. BANNER E CARDS TOP
# ==============================================================================
st.markdown(
    """
    <div class='main-header'>
        <h1>✈️ Grupo Arbaitman | Portal de Conciliação Aérea & Gestão de Bilhetes Pendentes (FP&A)</h1>
        <p>Leitura direta e atualizada da Visão de Conciliação Aérea no Supabase PostgreSQL</p>
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
c6.metric("Solicitados Conciliação", f"{len(f_sem_div):,}")

st.markdown("---")

abas = [
    "📊 Dashboard Executivo",
    "📋 1. Falta de Lançamento (Maringá)",
    "⚠️ 2. Erros de Valores & CIA (Maringá)",
    "🎧 3. Suporte Backoffice",
    "🎪 4. Central de Eventos",
    "🤖 5. Emissor Virtual Lemontech",
    "✅ 6. Bilhetes Solicitados Conciliação"
]
if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Histórico", "📥 Carga em Lote Avançada"])

aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# ABA 0: DASHBOARD EXECUTIVO COM ÍNDICES NOS GRÁFICOS
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.markdown("<div class='section-banner'><h3>📊 Painel Executivo C-Level & Desempenho por Gerente</h3></div>", unsafe_allow_html=True)
    
    df_pendentes = pd.concat([f_falta, f_erros, f_backoffice, f_eventos, f_lemon_virt], ignore_index=True)
    
    # Linha 1: Gráfico de Rosca de Status + Top CIAs
    g1, g2 = st.columns(2)
    
    with g1:
        st.markdown('<div class="chart-card"><h5>⚠️ Pendências por Status Operacional</h5>', unsafe_allow_html=True)
        if not df_pendentes.empty:
            df_st = df_pendentes["Status_Geral"].value_counts().reset_index()
            df_st.columns = ["Status", "Quantidade"]
            
            fig_pie = px.pie(
                df_st, values="Quantidade", names="Status", hole=0.5,
                color_discrete_sequence=["#002060", "#0050b3", "#0086c0", "#fa8c16", "#d90429"]
            )
            fig_pie.update_traces(
                textposition='inside',
                textinfo='value+percent', # ÍNDICES DENTRO DA FATIA
                hovertemplate='<b>%{label}</b><br>Quantidade: %{value:,}<br>Proporção: %{percent}'
            )
            fig_pie.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Sem pendências.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with g2:
        st.markdown('<div class="chart-card"><h5>✈️ Top 8 CIAs Aéreas com Pendências</h5>', unsafe_allow_html=True)
        if not df_pendentes.empty:
            df_cia = df_pendentes["CIA"].value_counts().head(8).reset_index()
            df_cia.columns = ["Companhia", "Quantidade"]
            
            fig_bar = px.bar(df_cia, x="Companhia", y="Quantidade", text="Quantidade", color_discrete_sequence=["#002060"])
            fig_bar.update_traces(
                textposition='outside', # ÍNDICES EM CIMA DA BARRA
                texttemplate='%{value:,}',
                hovertemplate='<b>%{x}</b><br>Pendências: %{y:,}'
            )
            fig_bar.update_layout(height=320, margin=dict(l=10, r=10, t=25, b=10), xaxis_title="", yaxis_title="Qtd Bilhetes")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Sem pendências.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Linha 2: Desempenho por Gerente (Quantidade) + Exposição Financeira (R$)
    g3, g4 = st.columns(2)
    
    with g3:
        st.markdown('<div class="chart-card"><h5>👤 Volume de Pendências por Gerente Responsável</h5>', unsafe_allow_html=True)
        if not df_pendentes.empty:
            df_ger = df_pendentes["Área Resp. Operação"].value_counts().reset_index()
            df_ger.columns = ["Gerente", "Quantidade"]
            
            fig_ger = px.bar(df_ger, x="Gerente", y="Quantidade", text="Quantidade", color_discrete_sequence=["#0050b3"])
            fig_ger.update_traces(
                textposition='outside', # ÍNDICES EM CIMA DA BARRA
                texttemplate='%{value:,}',
                hovertemplate='<b>%{x}</b><br>Pendências: %{y:,}'
            )
            fig_ger.update_layout(height=330, margin=dict(l=10, r=10, t=25, b=10), xaxis_title="", yaxis_title="Qtd Bilhetes")
            st.plotly_chart(fig_ger, use_container_width=True)
        else:
            st.info("Sem pendências.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with g4:
        st.markdown('<div class="chart-card"><h5>💼 Exposição Financeira (R$) por Gerente Responsável</h5>', unsafe_allow_html=True)
        if not df_pendentes.empty:
            df_pendentes["Valor_Total"] = df_pendentes["A vista"] + df_pendentes["A credito"] + df_pendentes["Taxa"]
            df_fin = df_pendentes.groupby("Área Resp. Operação")["Valor_Total"].sum().reset_index()
            df_fin = df_fin.sort_values("Valor_Total", ascending=False)
            df_fin["Texto_Fmt"] = df_fin["Valor_Total"].apply(lambda x: f"R$ {x:,.0f}")
            
            fig_fin = px.bar(df_fin, x="Área Resp. Operação", y="Valor_Total", text="Texto_Fmt", color_discrete_sequence=["#fa8c16"])
            fig_fin.update_traces(
                textposition='outside', # ÍNDICES VALOR FORMATADO EM CIMA DA BARRA
                hovertemplate='<b>%{x}</b><br>Valor Pendente: R$ %{y:,.2f}'
            )
            fig_fin.update_layout(height=330, margin=dict(l=10, r=10, t=25, b=10), xaxis_title="", yaxis_title="Valor (R$)")
            st.plotly_chart(fig_fin, use_container_width=True)
        else:
            st.info("Sem pendências.")
        st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# MÓDULO REUTILIZÁVEL DE TABELAS
# ------------------------------------------------------------------------------
def renderizar_tabela_tratativas(df_tab, nome_base, key_pfx):
    if df_tab.empty:
        st.info("Nenhum bilhete pendente nesta categoria.")
        return
        
    termo = st.text_input("🔍 Buscar por Bilhete, Gerente, Emissor ou Localizador:", key=f"src_{key_pfx}")
    df_exib = df_tab.copy()
    
    if termo.strip():
        t_low = termo.strip().lower()
        cols_s = [c for c in ["Bilhetes", "Localizador_Sistema", "Área Resp. Operação", "Emissor / Consultor", "Ponto de venda"] if c in df_exib.columns]
        m_s = pd.Series(False, index=df_exib.index)
        for col in cols_s:
            m_s |= df_exib[col].astype(str).str.lower().str.contains(t_low, na=False)
        df_exib = df_exib[m_s]
        
    st.write(f"**Registros Visíveis:** {len(df_exib)}")
    
    bilhet_sel = st.multiselect("Selecione Bilhetes para Atualizar:", options=df_exib["Bilhetes"].tolist(), key=f"ms_{key_pfx}")
    
    if bilhet_sel:
        df_sel = df_exib[df_exib["Bilhetes"].isin(bilhet_sel)]
        st.info(f"🎯 **{len(df_sel)} bilhete(s) selecionado(s)** para atualização em lote.")
        
        with st.form(f"form_{key_pfx}"):
            c1, c2 = st.columns(2)
            with c1:
                st_novo = st.selectbox("Novo Status:", [
                    "Já Lançado no ERP", "Emitido e Lançado", "Pendente de Lançamento (Não Consta)",
                    "Em Análise Backoffice", "Divergência de Tarifa", "Divergência de Taxa", "Cancelado / Reembolsado"
                ], key=f"st_{key_pfx}")
            with c2:
                opts_g = sorted(list(set(DEPARA_GERENTES.values())))
                ar_nova = st.selectbox("Nova Área Responsável / Gerente:", opts_g, key=f"ar_{key_pfx}")
                
            obs_nova = st.text_area("Observação / Justificativa:", key=f"obs_{key_pfx}")
            
            if st.form_submit_button("💾 Salvar Tratativa no Supabase"):
                usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                df_up = pd.DataFrame({
                    "Bilhetes": df_sel["Bilhetes"].values,
                    "Status_Geral": [st_novo] * len(df_sel),
                    "Área Resp. Operação": [ar_nova] * len(df_sel),
                    "Obs. Operação": [obs_nova] * len(df_sel),
                    "Setor": df_sel.get("Setor", pd.Series("-")).values,
                    "Gerentes": [ar_nova] * len(df_sel)
                })
                ok, msg, _, _, _ = salvar_tratativas_lote_supabase(df_up, usr_str)
                if ok:
                    st.success(msg)
                    time.sleep(1)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(msg)

    # Organização e Ordenação das Colunas para Exibição
    cols_ordem = ["Bilhetes", "Ponto de venda", "Área Resp. Operação", "Emissor / Consultor", "CIA", "Status_Geral", "Data Emissão", "A vista", "A credito", "Taxa", "Obs. Operação"]
    cols_existentes = [c for c in cols_ordem if c in df_exib.columns] + [c for c in df_exib.columns if c not in cols_ordem and c not in ["Dt_Parsed", "Bilhete_Clean"]]
    
    st.dataframe(df_exib[cols_existentes], use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# ABAS OPERACIONAIS
# ------------------------------------------------------------------------------
with aba_sel[1]:
    st.markdown("<div class='section-banner'><h3>📋 1. Falta de Lançamento (Maringá)</h3></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_falta, "Falta_Lancamento", "fl")

with aba_sel[2]:
    st.markdown("<div class='section-banner'><h3>⚠️ 2. Erros de Valores & CIA (Maringá)</h3></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_erros, "Erros_Valores", "ev")

with aba_sel[3]:
    st.markdown("<div class='section-banner'><h3>🎧 3. Suporte Backoffice</h3></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_backoffice, "Suporte_Backoffice", "sb")

with aba_sel[4]:
    st.markdown("<div class='section-banner'><h3>🎪 4. Central de Eventos</h3></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_eventos, "Central_Eventos", "ce")

with aba_sel[5]:
    st.markdown("<div class='section-banner'><h3>🤖 5. Emissor Virtual Lemontech</h3></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_lemon_virt, "Emissor_Virtual", "evl")

with aba_sel[6]:
    st.markdown("<div class='section-banner'><h3>✅ 6. Bilhetes Solicitados Conciliação</h3></div>", unsafe_allow_html=True)
    st.dataframe(f_sem_div, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# ABAS MASTER: HISTÓRICO, ACESSOS E CARGA EM LOTE
# ------------------------------------------------------------------------------
if e_master():
    with aba_sel[7]:
        st.subheader("⚙️ Gestão de Acessos")
        usuarios_db = carregar_usuarios_supabase()
        if usuarios_db:
            df_usr = pd.DataFrame.from_dict(usuarios_db, orient="index").reset_index()
            df_usr.columns = ["Usuário", "Senha", "Nome", "Perfil", "Status"]
            st.dataframe(df_usr[["Usuário", "Nome", "Perfil", "Status"]], use_container_width=True, hide_index=True)

    with aba_sel[8]:
        st.subheader("📜 Log de Histórico Supabase")
        engine_sb = get_db_engine()
        if engine_sb:
            try:
                df_log_db = pd.read_sql("SELECT * FROM log_auditoria ORDER BY id DESC LIMIT 300", engine_sb)
                st.dataframe(df_log_db, use_container_width=True, hide_index=True)
            except Exception:
                st.info("Sem registros no log.")

    with aba_sel[9]:
        st.subheader("📥 Carga de Retornos de Cobrança em Lote")
        st.markdown("Suba o arquivo `.xlsx` ou `.csv` enviado pelos Gerentes para atualizar bilhetes em lote no Supabase.")
        
        arq_up = st.file_uploader("Selecione o arquivo de retorno:", type=["xlsx", "csv"], key="upl_lote")
        if arq_up:
            try:
                df_lote_raw = pd.read_csv(arq_up, dtype=str) if arq_up.name.endswith(".csv") else pd.read_excel(arq_up, dtype=str)
                st.success(f"Arquivo lido com sucesso! Total de {len(df_lote_raw)} linhas.")
                st.dataframe(df_lote_raw.head(10), use_container_width=True)
                
                if st.button("🚀 Confirmar e Atualizar no Supabase"):
                    usr_str = f"CargaLote_{st.session_state['usuario_atual']}"
                    ok, msg, n_at, n_nv, _ = salvar_tratativas_lote_supabase(df_lote_raw, usr_str)
                    if ok:
                        st.success(msg)
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(msg)
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {e}")