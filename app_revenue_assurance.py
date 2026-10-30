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
# 1. CONFIGURAÇÃO INICIAL E ESTILOS CSS EXECUTIVOS (C-LEVEL STYLING)
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Portal de Conciliação Aérea (FP&A)",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)
ARQUIVO_DASHBOARD_PADRAO = "Dashboard_Revenue_Assurance_Consolidado.xlsx"
TEMPO_MAXIMO_SESSAO_HORAS = 8.0 

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

DEPARA_GERENTES = {
    "FABIANO SOUZA": "Fabiano Souza", "FABIANO": "Fabiano Souza", "Lazer": "Fabiano Souza", "LAZER": "Fabiano Souza", "CONCIERGE": "Fabiano Souza", "Concierge": "Fabiano Souza",
    "GUILHERME SILVA": "Guilherme Silva", "GUILHERME": "Guilherme Silva",
    "IVANETE BERTASOL": "Ivanete Bertasol", "IVANETE": "Ivanete Bertasol",
    "KELI SANTI": "Keli Santi", "KELI": "Keli Santi",
    "MARCELO PEREIRA": "Marcelo Pereira", "MARCELO": "Marcelo Pereira",
    "ROSANGELA PALLU": "Rosangela Pallu", "ROSANGELA": "Rosangela Pallu",
    "JAIME SCHNAIDER": "Jaime Schnaider", "JAIME": "Jaime Schnaider",
    "SILVANA CELANI": "Silvana Celani", "SILVANA": "Silvana Celani", "Private": "Silvana Celani", "PRIVATE": "Silvana Celani",
    "KATIA MARTINS": "Katia Martins", "KÁTIA MARTINS": "Katia Martins", "KATIA": "Katia Martins",
    "JOAO PAULO": "João Paulo", "JOÃO PAULO": "João Paulo",
    "SUPORTE BACKOFFICE": "Suporte Backoffice", "BACKOFFICE": "Suporte Backoffice",
    "CENTRAL DE EVENTOS": "Central de Eventos", "EVENTOS": "Central de Eventos",
    "EMISSOR VIRTUAL": "Emissor Virtual Lemontech", "EMISSOR VIRTUAL LEMONTECH": "Emissor Virtual Lemontech",
    "LEMONTECH VIRTUAL": "Emissor Virtual Lemontech", "VIRTUAL LEMONTECH": "Emissor Virtual Lemontech", "LEMONTECH": "Emissor Virtual Lemontech"
}

def padronizar_nome_gerente(nome_raw):
    if not nome_raw or str(nome_raw).strip().lower() in ["nan", "none", "", "-", "não mapeado", "nao mapeado"]:
        return "Não Mapeado"
    s = str(nome_raw).strip().upper()
    s_clean = re.sub(r"\s+", " ", s)
    for k, v in DEPARA_GERENTES.items():
        if k == s_clean or k in s_clean:
            return v
    return s_clean.title()

def consolidar_emissor_consultor(row):
    emis = str(row.get("Emissor") or row.get("Emissor_Reserva_Lemon") or "").strip()
    emis_clean = emis.replace("Sem emissor", "").replace("Sem consultor", "").strip()
    
    if emis_clean and emis_clean.lower() not in ["nan", "none", "", "-", "null"]:
        return emis_clean
    cons = str(row.get("Consultor") or row.get("Consultor_Lemon") or "").strip()
    cons_clean = cons.replace("Sem consultor", "").replace("Sem emissor", "").strip()
    
    if cons_clean and cons_clean.lower() not in ["nan", "none", "", "-", "null"]:
        return cons_clean
    return "-"

# Injeção de CSS
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        * { font-family: 'Inter', -apple-system, sans-serif; }
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
        
        .fonte-pill {
            background: rgba(255, 255, 255, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.3);
            color: #ffffff;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 500;
            display: inline-block;
            margin-right: 6px;
            margin-top: 8px;
        }
        .update-badge {
            background: rgba(255, 255, 255, 0.25);
            border: 1px solid rgba(255, 255, 255, 0.4);
            color: #ffffff;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            display: inline-block;
            margin-top: 8px;
            float: right;
        }
        div[data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #f1f5f9;
            padding: 8px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            overflow-x: auto;
            white-space: nowrap;
        }
        button[data-baseweb="tab"] {
            height: 40px;
            border-radius: 8px !important;
            padding: 8px 16px !important;
            background-color: transparent !important;
            color: #334155 !important;
            font-weight: 600 !important;
            font-size: 13px !important;
            border: 1px solid transparent !important;
            transition: all 0.2s ease-in-out !important;
        }
        button[data-baseweb="tab"]:hover {
            background-color: #e2e8f0 !important;
            color: #002060 !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #002060 !important;
            color: #ffffff !important;
            border-radius: 8px !important;
            box-shadow: 0 4px 12px rgba(0, 32, 96, 0.25) !important;
        }
        div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {
            display: none !important;
        }
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
        .chart-card h5 { color: #002060; font-weight: 700; font-size: 15px; margin-top: 0; margin-bottom: 12px; }
        
        div[data-testid="stMetric"] {
            background: #ffffff;
            border-radius: 8px;
            padding: 14px;
            border-top: 4px solid #002060;
            box-shadow: 0 2px 10px rgba(0,0,0,0.03);
        }
        
        div[data-testid="stDataFrame"] {
            background-color: #ffffff;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# 2. CONEXÃO COM BANCO DE DADOS, MIGRAÇÃO & LOGS DE ACESSO
# ==============================================================================
def sanitizar_bilhete(val):
    if pd.isna(val) or val is None: return ""
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
            engine = create_engine(db_url, pool_pre_ping=True)
            
            try:
                with engine.begin() as conn:
                    conn.execute(text("""
                        CREATE TABLE IF NOT EXISTS tratativas (
                            bilhete VARCHAR(255) PRIMARY KEY,
                            status_geral VARCHAR(255),
                            area_resp VARCHAR(255),
                            obs_operacao TEXT,
                            setor VARCHAR(255),
                            gerentes VARCHAR(255),
                            usuario_modificacao VARCHAR(255),
                            data_modificacao TIMESTAMP
                        );
                        ALTER TABLE tratativas ADD COLUMN IF NOT EXISTS sv VARCHAR(255);
                        ALTER TABLE tratativas ADD COLUMN IF NOT EXISTS fop VARCHAR(255);
                        ALTER TABLE tratativas ADD COLUMN IF NOT EXISTS rloc_cia VARCHAR(255);
                        ALTER TABLE tratativas ADD COLUMN IF NOT EXISTS cliente VARCHAR(255);
                        
                        CREATE TABLE IF NOT EXISTS log_acessos (
                            id SERIAL PRIMARY KEY,
                            usuario VARCHAR(255),
                            nome VARCHAR(255),
                            data_acesso TIMESTAMP DEFAULT NOW()
                        );
                    """))
            except Exception: pass
            return engine
    except Exception as e:
        st.error(f"Erro na conexão com o banco de dados: {e}")
    return None

def registrar_log_acesso(usuario_id, nome_usuario):
    engine = get_db_engine()
    if engine:
        try:
            with engine.begin() as conn:
                conn.execute(
                    text("INSERT INTO log_acessos (usuario, nome, data_acesso) VALUES (:usuario, :nome, NOW())"),
                    {"usuario": str(usuario_id).strip().lower(), "nome": str(nome_usuario).strip()}
                )
        except Exception: pass

def obter_relatorio_acessos_db():
    engine = get_db_engine()
    if not engine:
        return pd.DataFrame(), pd.DataFrame()
    try:
        query_log = "SELECT usuario, nome, data_acesso FROM log_acessos ORDER BY data_acesso DESC LIMIT 500"
        df_logs = pd.read_sql(query_log, engine)
        if not df_logs.empty:
            df_logs["data_acesso"] = pd.to_datetime(df_logs["data_acesso"])
            
        df_users = pd.DataFrame()
        try:
            df_users = pd.read_sql("SELECT usuario, nome, perfil, status FROM usuarios", engine)
            df_users["usuario"] = df_users["usuario"].astype(str).str.strip().str.lower()
        except Exception: pass
        if not df_logs.empty and not df_users.empty:
            df_agg = df_logs.groupby("usuario").agg(
                ultimo_acesso=("data_acesso", "max"),
                total_acessos=("data_acesso", "count")
            ).reset_index()
            
            df_m = pd.merge(df_users, df_agg, on="usuario", how="left")
        elif not df_users.empty:
            df_m = df_users.copy()
            df_m["ultimo_acesso"] = pd.NaT
            df_m["total_acessos"] = 0
        else:
            df_m = pd.DataFrame()
        return df_m, df_logs
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

def carregar_tratativas_db():
    engine = get_db_engine()
    if not engine:
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "sv", "fop", "rloc_cia", "cliente", "data_modificacao"])
    query = "SELECT bilhete, status_geral, area_resp, obs_operacao, setor, gerentes, sv, fop, rloc_cia, cliente, data_modificacao FROM tratativas"
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
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "sv", "fop", "rloc_cia", "cliente", "data_modificacao"])

def salvar_tratativas_lote_supabase(df_lote, usuario):
    engine = get_db_engine()
    if not engine or df_lote.empty:
        return False, "Conexão com o banco de dados indisponível.", 0, 0, 0
    df_existentes = carregar_tratativas_db()
    set_existentes = set(df_existentes["bilhete"].apply(sanitizar_bilhete).tolist()) if not df_existentes.empty else set()
    dados_lote = []
    qtd_atualizados, qtd_novos = 0, 0
    
    for _, r in df_lote.iterrows():
        b_clean = sanitizar_bilhete(r.get("Bilhetes", r.get("bilhete", "")))
        if b_clean and b_clean.lower() not in ["nan", "none", "", "-"]:
            obs_limpa = str(r.get("Obs. Operação", r.get("obs_operacao", ""))).replace("Sem tratativa na operação", "").strip()
            st_val = str(r.get("Status_Geral", r.get("status_geral", ""))).strip()
            ar_val = padronizar_nome_gerente(r.get("Área Resp. Operação", r.get("area_resp", "")))
            setor_val = str(r.get("Setor", r.get("setor", "-"))).strip()
            gerente_val = padronizar_nome_gerente(r.get("Gerentes", r.get("gerentes", ar_val)))
            
            sv_val = str(r.get("SV", r.get("sv", r.get("Solicitação", "-")))).strip()
            fop_val = str(r.get("FOP", r.get("fop", r.get("Pagto", "-")))).strip()
            rloc_val = str(r.get("Rloc_CIA", r.get("rloc_cia", r.get("Rloc_Cia", "-")))).strip()
            cliente_val = str(r.get("Cliente", r.get("cliente", ""))).strip()
            
            if st_val in ["-", "", "None", "nan"]:
                st_val = "Pendente de Lançamento (Não Consta)"
            if b_clean in set_existentes:
                qtd_atualizados += 1
            else:
                qtd_novos += 1
            dados_lote.append({
                "bilhete": b_clean, "status_geral": st_val, "area_resp": ar_val,
                "obs_operacao": obs_limpa, "setor": setor_val, "gerentes": gerente_val,
                "sv": sv_val, "fop": fop_val, "rloc_cia": rloc_val, "cliente": cliente_val, "usuario": str(usuario)
            })
            
    if not dados_lote:
        return False, "Nenhum bilhete válido para atualização.", 0, 0, 0
        
    sql_upsert = text("""
        INSERT INTO tratativas (bilhete, status_geral, area_resp, obs_operacao, setor, gerentes, sv, fop, rloc_cia, cliente, usuario_modificacao, data_modificacao)
        VALUES (:bilhete, :status_geral, :area_resp, :obs_operacao, :setor, :gerentes, :sv, :fop, :rloc_cia, :cliente, :usuario, NOW())
        ON CONFLICT (bilhete) DO UPDATE SET
            status_geral = EXCLUDED.status_geral, area_resp = EXCLUDED.area_resp, obs_operacao = EXCLUDED.obs_operacao,
            setor = EXCLUDED.setor, gerentes = EXCLUDED.gerentes, sv = EXCLUDED.sv, fop = EXCLUDED.fop,
            rloc_cia = EXCLUDED.rloc_cia, cliente = EXCLUDED.cliente, usuario_modificacao = EXCLUDED.usuario_modificacao, data_modificacao = NOW();
    """)
    try:
        with engine.begin() as conn:
            conn.execute(sql_upsert, dados_lote)
        st.cache_data.clear()
        return True, f"✅ Sincronizado com sucesso! {qtd_atualizados} atualizados e {qtd_novos} novos.", qtd_atualizados, qtd_novos, 0
    except Exception as e:
        return False, f"Erro ao salvar tratativas: {e}", 0, 0, len(dados_lote)

# ==============================================================================
# 3. GERADOR DE EXCEL ESTILIZADO
# ==============================================================================
def gerar_excel_estilizado(df_export, nome_aba="Relatorio", nome_gerente="Geral"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(columns=["Aviso"], data=[["Nenhum registro encontrado"]])
        
    cols_remover = [c for c in ["Dt_Parsed", "Bilhete_Clean", "Aba_Origem"] if c in df_export.columns]
    df_clean = df_export.drop(columns=cols_remover) if cols_remover else df_export.copy()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    nome_sheet = re.sub(r'[\*\?:/\[\]]', '', nome_aba)[:30]
    ws.title = nome_sheet if nome_sheet else "Relatorio"
    ws.views.sheetView[0].showGridLines = True
    
    header_title_fill = PatternFill(start_color="001E4D", end_color="001E4D", fill_type="solid")
    header_title_font = Font(name="Calibri", color="FFFFFF", bold=True, size=14)
    subtitle_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    subtitle_font = Font(name="Calibri", color="002060", bold=True, size=10)
    
    max_cols = max(len(df_clean.columns), 6)
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    cell_title = ws.cell(row=1, column=1, value="✈️ Grupo Arbaitman | Relatório Executivo de Conciliação Aérea (FP&A)")
    cell_title.fill = header_title_fill
    cell_title.font = header_title_font
    cell_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    
    dt_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    gerente_label = f"Gerente Responsável: {nome_gerente}" if nome_gerente and nome_gerente != "Geral" else "Escopo: Geral (Todos os Gerentes)"
    sub_txt = f"Aba: {nome_aba.replace('_', ' ')}  |  {gerente_label}  |  Emissão: {dt_str}"
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
    cell_sub = ws.cell(row=2, column=1, value=sub_txt)
    cell_sub.fill = subtitle_fill
    cell_sub.font = subtitle_font
    cell_sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 10
    
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
    )
    fill_alerta = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_alerta = Font(color="9C0006", bold=True)
    
    currency_cols = [
        "A vista", "A credito", "Tarifa Total Emitida", "Tarifa_Sistema", "Dif_Tarifa",
        "Taxa", "Taxa_Sistema", "Dif_Taxa", "Comissão", "Taxa DU", "Incentivo",
        "Receita Total Emitida", "Receita_Sistema", "Dif_Receita"
    ]
    
    for col_idx, col_name in enumerate(df_clean.columns, 1):
        cell = ws.cell(row=4, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[4].height = 24
    
    for r_idx, row_data in enumerate(df_clean.values, start=5):
        ws.row_dimensions[r_idx].height = 19
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            col_name = df_clean.columns[c_idx - 1]
            if col_name in currency_cols and val not in ["-", None, ""]:
                try:
                    cell.value = float(val)
                    cell.number_format = "R$ #,##0.00"
                except (ValueError, TypeError): pass
            if col_name.startswith("Dif_") and val not in ["-", None, ""]:
                try:
                    if abs(float(val)) >= 0.01:
                        cell.fill = fill_alerta
                        cell.font = font_alerta
                except (ValueError, TypeError): pass
                
    for col_idx in range(1, len(df_clean.columns) + 1):
        col_letter = get_column_letter(col_idx)
        len_vals = [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(4, min(ws.max_row + 1, 100))]
        max_len = max(len_vals) if len_vals else 10
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 42)
        
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# 4. AUTENTICAÇÃO E PERSISTÊNCIA DE SESSÃO
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
        except Exception: pass
    return dict_u

def alterar_senha_usuario(usuario_id, senha_atual, nova_senha):
    engine = get_db_engine()
    if not engine:
        return False, "Banco de dados indisponível."
    try:
        with engine.begin() as conn:
            res = conn.execute(
                text("SELECT senha FROM usuarios WHERE LOWER(usuario) = LOWER(:usuario)"),
                {"usuario": usuario_id}
            ).fetchone()
            if not res:
                return False, "Usuário não encontrado."
            senha_banco = str(res[0]).strip()
            if senha_banco != str(senha_atual).strip():
                return False, "A senha atual informada está incorreta."
                
            conn.execute(
                text("UPDATE usuarios SET senha = :nova_senha WHERE LOWER(usuario) = LOWER(:usuario)"),
                {"nova_senha": str(nova_senha).strip(), "usuario": usuario_id}
            )
            return True, "🔑 Senha alterada com sucesso!"
    except Exception as e:
        return False, f"Erro ao atualizar senha no banco: {e}"

if "autenticado" not in st.session_state: st.session_state["autenticado"] = False
if "usuario_atual" not in st.session_state: st.session_state["usuario_atual"] = None
if "perfil_atual" not in st.session_state: st.session_state["perfil_atual"] = None
if "login_user_id" not in st.session_state: st.session_state["login_user_id"] = None

def validar_sessao_url():
    session_user_param = st.query_params.get("session_user")
    session_time_param = st.query_params.get("session_time")
    if not session_user_param or not session_time_param:
        return False, None
    try:
        ts_login = float(session_time_param)
        ts_atual = time.time()
        horas_decorridas = (ts_atual - ts_login) / 3600.0
        if horas_decorridas > TEMPO_MAXIMO_SESSAO_HORAS:
            st.query_params.clear()
            st.warning(f"⚠️ Sua sessão expirou após {int(TEMPO_MAXIMO_SESSAO_HORAS)} horas de inatividade.")
            return False, None
        else:
            st.query_params["session_time"] = str(ts_atual)
            return True, str(session_user_param).strip().lower()
    except Exception:
        st.query_params.clear()
        return False, None

if not st.session_state["autenticado"]:
    sessao_valida, u_key = validar_sessao_url()
    if sessao_valida and u_key:
        usuarios_db = carregar_usuarios_supabase()
        if u_key in usuarios_db:
            st.session_state["autenticado"] = True
            st.session_state["usuario_atual"] = usuarios_db[u_key]["nome"]
            st.session_state["perfil_atual"] = usuarios_db[u_key]["perfil"]
            st.session_state["login_user_id"] = u_key
            registrar_log_acesso(u_key, usuarios_db[u_key]["nome"])
        elif u_key == "mribeiro":
            st.session_state["autenticado"] = True
            st.session_state["usuario_atual"] = "Marcos Ribeiro"
            st.session_state["perfil_atual"] = "Master"
            st.session_state["login_user_id"] = "mribeiro"
            registrar_log_acesso("mribeiro", "Marcos Ribeiro")

def e_master():
    u_id = st.session_state.get("login_user_id", "")
    perfil = st.session_state.get("perfil_atual", "")
    return u_id in ["mribeiro", "fellipe", "ffernandes"] or perfil in ["Master", "Compliance"]

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown("<div class='main-header'><h1>✈️ Grupo Arbaitman</h1><p>Portal de Conciliação Aérea & Gestão de Bilhetes (FP&A)</p></div>", unsafe_allow_html=True)
        u_input = st.text_input("Usuário:", key="l_user").strip().lower()
        p_input = str(st.text_input("Senha:", type="password", key="l_pass")).strip()
        if st.button("Acessar Portal", type="primary"):
            usuarios_db = carregar_usuarios_supabase()
            if u_input in usuarios_db and usuarios_db[u_input]["senha"] == p_input:
                st.session_state["autenticado"] = True
                st.session_state["usuario_atual"] = usuarios_db[u_input]["nome"]
                st.session_state["perfil_atual"] = usuarios_db[u_input]["perfil"]
                st.session_state["login_user_id"] = u_input
                st.query_params["session_user"] = u_input
                st.query_params["session_time"] = str(time.time())
                registrar_log_acesso(u_input, usuarios_db[u_input]["nome"])
                st.rerun()
            elif u_input == "mribeiro" and p_input in ["14052013", "123"]:
                st.session_state["autenticado"] = True
                st.session_state["usuario_atual"] = "Marcos Ribeiro"
                st.session_state["perfil_atual"] = "Master"
                st.session_state["login_user_id"] = "mribeiro"
                st.query_params["session_user"] = "mribeiro"
                st.query_params["session_time"] = str(time.time())
                registrar_log_acesso("mribeiro", "Marcos Ribeiro")
                st.rerun()
            else: st.error("Usuário ou senha incorretos.")
    st.stop()

# ==============================================================================
# 5. ROTEADOR DE BASES E COMPATIBILIDADE DE COLUNAS
# ==============================================================================
def clean_str(val):
    if pd.isna(val) or val is None: return "-"
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    return "-" if s.lower() in ["none", "nan", "null", "<na>", ""] else s

def padronizar_df(df):
    if df is None or df.empty: return pd.DataFrame()
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    
    for c in df_out.columns:
        if df_out[c].dtype == object or df_out[c].dtype == 'string':
            df_out[c] = df_out[c].apply(clean_str)
            
    if "Obs. Operação" in df_out.columns:
        df_out["Obs. Operação"] = df_out["Obs. Operação"].astype(str).str.replace("Sem tratativa na operação", "", regex=False).str.strip()
        
    for col_g in ["Área Resp. Operação", "Gerentes"]:
        if col_g in df_out.columns:
            df_out[col_g] = df_out[col_g].apply(padronizar_nome_gerente)
            
    df_out["Emissor / Consultor"] = df_out.apply(consolidar_emissor_consultor, axis=1)
    
    cols_sv_candidatas = ["SV", "Solicitação", "Solicitacao", "Solicitacao_Extrato", "Solicitacao_Lemon"]
    sv_series = pd.Series("-", index=df_out.index)
    for c_sv in cols_sv_candidatas:
        if c_sv in df_out.columns:
            val_clean = df_out[c_sv].astype(str).str.strip()
            mask_val = val_clean.notna() & ~val_clean.str.lower().isin(["nan", "none", "", "-", "null", "<na>"])
            sv_series = sv_series.where(~mask_val, val_clean)
    df_out["SV"] = sv_series

    for c_fop in ["Forma de Recebimento", "Pagto", "FOP"]:
        if c_fop in df_out.columns and "FOP" not in df_out.columns:
            df_out["FOP"] = df_out[c_fop]
    if "FOP" not in df_out.columns: df_out["FOP"] = "-"

    for c_rloc in ["Rloc_Cia", "Localizador_Sistema", "Rloc_CIA"]:
        if c_rloc in df_out.columns and "Rloc_CIA" not in df_out.columns:
            df_out["Rloc_CIA"] = df_out[c_rloc]
    if "Rloc_CIA" not in df_out.columns: df_out["Rloc_CIA"] = "-"

    if "CIA" in df_out.columns:
        df_out["CIA"] = df_out["CIA"].astype(str).str.upper().str.strip().replace(DEPARA_CIAS)
        
    for c in ["Taxa", "A vista", "A credito", "Comissão", "Taxa DU", "Incentivo", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa"]:
        if c in df_out.columns:
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)
            
    df_out["Tarifa Total Emitida"] = df_out["A vista"] + df_out["A credito"]
    df_out["Receita Total Emitida"] = df_out["Comissão"] + df_out["Taxa DU"] + df_out["Incentivo"]
    
    def converter_data_flexivel(val):
        if pd.isna(val) or val in ["-", "", "nan", "None", None]:
            return pd.NaT
        if isinstance(val, (int, float)):
            try: return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')
            except Exception: return pd.NaT
        try: return pd.to_datetime(val, dayfirst=True, errors="coerce")
        except Exception: return pd.NaT

    if "Data Emissão" in df_out.columns:
        df_out["Dt_Parsed"] = df_out["Data Emissão"].apply(converter_data_flexivel)
    else:
        df_out["Dt_Parsed"] = pd.NaT

    hoje = pd.to_datetime(datetime.date.today())
    dias = (hoje - df_out["Dt_Parsed"]).dt.days
    df_out["Dias_Pendente"] = dias
    
    def calc_aging(d):
        if pd.isna(d): return "Sem Data"
        if d <= 30: return "Até 30 dias"
        elif d <= 60: return "31 a 60 dias"
        elif d <= 90: return "61 a 90 dias"
        else: return "> 90 dias (Crítico)"
    df_out["Faixa_Aging"] = dias.apply(calc_aging)
    
    return df_out

def reordenar_colunas(df, modo="completo"):
    if df.empty: return df
    if modo == "falta_lancamento":
        cols_desejadas = [
            "Bilhetes", "Rloc_CIA", "FOP", "Área Resp. Operação", "Código Iata", "SV", "Ponto de venda", "Cliente", "Emissor / Consultor",
            "CIA", "Data Emissão", "Status_Geral", "Obs. Operação",
            "A vista", "A credito", "Tarifa Total Emitida", "Taxa", "Comissão", "Taxa DU", "Incentivo", "Receita Total Emitida"
        ]
    else:
        cols_desejadas = [
            "Bilhetes", "Rloc_CIA", "FOP", "Área Resp. Operação", "Código Iata", "SV", "Ponto de venda", "Cliente", "Emissor / Consultor",
            "CIA", "Localizador_Sistema", "Data Emissão", "Status_Geral", "Status_Divergencia", "Status_Cia", "Obs. Operação",
            "A vista", "A credito", "Tarifa Total Emitida", "Taxa", "Comissão", "Taxa DU", "Incentivo", "Receita Total Emitida",
            "Fornecedor_Sistema", "Status_Sistema", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa", "Receita_Sistema", "Dif_Receita"
        ]
        
    cols_finais = [c for c in cols_desejadas if c in df.columns]
    cols_protegidas = ["Dt_Parsed", "Dias_Pendente", "Faixa_Aging", "Bilhete_Clean", "Aba_Origem"]
    cols_sobra = [c for c in df.columns if c not in cols_finais and c not in cols_protegidas]
    
    if modo != "falta_lancamento":
        cols_finais.extend(cols_sobra)
        
    for cp in cols_protegidas:
        if cp in df.columns and cp not in cols_finais:
            cols_finais.append(cp)
            
    return df[cols_finais]

def rotear_bases_mestra(df_master):
    if df_master is None or df_master.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_master = padronizar_df(df_master)
    
    st_geral_str = df_master["Status_Geral"].astype(str).str.lower().str.strip()
    st_div_str = df_master.get("Status_Divergencia", pd.Series("")).astype(str).str.lower().str.strip()
    
    mask_ok = st_geral_str.str.contains("já lançado|emitido e lançado|lançado|conciliado|valores corretos|regularizado", na=False) | \
              st_div_str.str.contains("valores corretos|sem divergência", na=False)
              
    df_ok = df_master[mask_ok].copy()
    df_pend = df_master[~mask_ok].copy()
    
    # 2. Concur
    def e_concur(r):
        obs_val = str(r.get("Obs. Operação", "")).lower()
        st_val = str(r.get("Status_Geral", "")).lower()
        fonte_val = str(r.get("Fonte_Relatorio", "")).lower()
        return "concur" in obs_val or "concur" in st_val or "concur" in fonte_val
    mask_concur = df_pend.apply(e_concur, axis=1)
    df_concur = df_pend[mask_concur].copy()
    df_rest = df_pend[~mask_concur].copy()
    
    # 3. Emissor Virtual Lemontech
    def e_emissor_virtual_lemon(r):
        st_val = str(r.get("Status_Geral", "")).lower().strip()
        obs_val = str(r.get("Obs. Operação", "")).lower().strip()
        emis_cons = str(r.get("Emissor / Consultor", "")).lower().strip()
        emis_raw = str(r.get("Emissor", "")).lower().strip()
        ar_val = str(r.get("Área Resp. Operação", "")).lower().strip()
        
        if "emissor virtual" in st_val or "virtual lemontech" in st_val or "emissor virtual lemontech" in st_val:
            return True
        if "emissor virtual" in ar_val or "virtual lemontech" in ar_val:
            return True
            
        text_combined = f"{emis_cons} {emis_raw} {obs_val} {st_val}"
        termos = ["emissor virtual", "virtual lemon", "lemontech virtual", "robô lemon", "robo lemon"]
        if any(t in text_combined for t in termos):
            return True
        return False

    mask_lemon = df_rest.apply(e_emissor_virtual_lemon, axis=1)
    df_lemon_virt = df_rest[mask_lemon].copy()
    df_rest = df_rest[~mask_lemon].copy()

    # 4. Suporte Backoffice
    def e_suporte_backoffice(r):
        ar_val = str(r.get("Área Resp. Operação", "")).lower()
        ger_val = str(r.get("Gerentes", "")).lower()
        obs_val = str(r.get("Obs. Operação", "")).lower()
        st_val = str(r.get("Status_Geral", "")).lower()
        return "backoffice" in ar_val or "backoffice" in ger_val or "katia" in ar_val or any(k in obs_val or k in st_val for k in ["chamado", "ticket", "suporte"])
    mask_bo = df_rest.apply(e_suporte_backoffice, axis=1)
    df_bo = df_rest[mask_bo].copy()
    df_rest = df_rest[~mask_bo].copy()
    
    # 5. Central de Eventos
    mask_evt = df_rest.apply(lambda r: "eventos" in str(r.get("Área Resp. Operação", "")).lower() or "eventos" in str(r.get("Gerentes", "")).lower(), axis=1)
    df_eventos = df_rest[mask_evt].copy()
    df_rest = df_rest[~mask_evt].copy()
    
    # 6. Falta de Lançamento vs Erros de Valores
    mask_falta = df_rest.apply(lambda r: "falta de lançamento" in str(r.get("Status_Geral", "")).lower() or "não consta" in str(r.get("Status_Geral", "")).lower() or str(r.get("Status_Geral", "")).lower() in ["-", "", "none"], axis=1)
    df_falta = df_rest[mask_falta].copy()
    df_erros = df_rest[~mask_falta].copy()
    
    df_falta = reordenar_colunas(df_falta, modo="falta_lancamento")
    df_concur = reordenar_colunas(df_concur, modo="falta_lancamento")
    df_erros = reordenar_colunas(df_erros, modo="completo")
    df_bo = reordenar_colunas(df_bo, modo="completo")
    df_eventos = reordenar_colunas(df_eventos, modo="completo")
    df_lemon_virt = reordenar_colunas(df_lemon_virt, modo="completo")
    df_ok = reordenar_colunas(df_ok, modo="completo")
    
    return df_falta, df_concur, df_erros, df_bo, df_eventos, df_lemon_virt, df_ok

@st.cache_data(ttl=60, show_spinner=False)
def carregar_bases():
    dt_atualizacao = datetime.datetime.now()
    df_m = pd.DataFrame()
    engine = get_db_engine()
    
    # 1. Leitura direta e sem filtro intermediário da View Mestra do Supabase
    if engine:
        try:
            query = "SELECT * FROM vw_conciliacao_mestra"
            df_m = pd.read_sql(query, engine)
            
            if not df_m.empty:
                if "Ultima_Alteracao" in df_m.columns and df_m["Ultima_Alteracao"].notna().any():
                    dt_val = pd.to_datetime(df_m["Ultima_Alteracao"]).max()
                    if pd.notna(dt_val):
                        dt_atualizacao = dt_val
        except Exception as e:
            try:
                query_base = "SELECT * FROM tb_conciliacao_base"
                df_m = pd.read_sql(query_base, engine)
            except Exception: pass

    # 2. Fallback para arquivo Excel local caso o banco esteja indisponível
    if df_m.empty:
        if os.path.exists(ARQUIVO_DASHBOARD_PADRAO):
            fonte = ARQUIVO_DASHBOARD_PADRAO
            try:
                mod_time = os.path.getmtime(ARQUIVO_DASHBOARD_PADRAO)
                dt_atualizacao = datetime.datetime.fromtimestamp(mod_time)
            except Exception: pass
            
            try:
                xls = pd.ExcelFile(fonte, engine="openpyxl")
                frames = []
                for sheet in xls.sheet_names:
                    if sheet.startswith("99_") or sheet.startswith("98_"):
                        df_sheet = pd.read_excel(xls, sheet_name=sheet)
                        frames.append(df_sheet)
                if frames:
                    df_m = pd.concat(frames, ignore_index=True)
            except Exception: pass

    if df_m.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), dt_atualizacao.strftime("%d/%m/%Y às %H:%M:%S")

    # 3. Limpeza, padronização e garantia do formato dos bilhetes
    if "Bilhetes" in df_m.columns:
        df_m["Bilhete_Clean"] = df_m["Bilhetes"].apply(sanitizar_bilhete)
        df_m = df_m[df_m["Bilhete_Clean"].astype(str).str.len() > 3]
        df_m = df_m.drop_duplicates(subset=["Bilhete_Clean"], keep="first").reset_index(drop=True)
    elif "bilhete" in df_m.columns:
        df_m["Bilhete_Clean"] = df_m["bilhete"].apply(sanitizar_bilhete)
        df_m = df_m[df_m["Bilhete_Clean"].astype(str).str.len() > 3]
        df_m = df_m.drop_duplicates(subset=["Bilhete_Clean"], keep="first").reset_index(drop=True)
        df_m["Bilhetes"] = df_m["Bilhete_Clean"]

    df_m = padronizar_df(df_m)
    
    # 4. Roteamento dinâmico para as abas operacionais
    f_falta, f_concur, f_erros, f_bo, f_evt, f_lem, f_ok = rotear_bases_mestra(df_m)
    
    dt_str = dt_atualizacao.strftime("%d/%m/%Y às %H:%M:%S") if isinstance(dt_atualizacao, (datetime.datetime, pd.Timestamp)) else str(dt_atualizacao)
    return f_falta, f_concur, f_erros, f_bo, f_evt, f_lem, f_ok, dt_str

# ==============================================================================
# 6. CARREGAMENTO E FILTROS OPERACIONAIS
# ==============================================================================
df_falta, df_concur, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div, dt_ultima_atualizacao_str = carregar_bases()
df_todos_unificados = pd.concat([df_falta, df_concur, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div], ignore_index=True)

st.sidebar.title("Grupo Arbaitman")
st.sidebar.caption("Conciliação Aérea FP&A v5.0")
st.sidebar.write(f"👤 **{st.session_state['usuario_atual']}** ({st.session_state['perfil_atual']})")
st.sidebar.info(f"🕒 **Última Atualização:**\n{dt_ultima_atualizacao_str}")

with st.sidebar.expander("🔑 Alterar Minha Senha"):
    with st.form("form_alterar_senha_sb"):
        pass_act = st.text_input("Senha Atual:", type="password", key="pass_act")
        pass_new = st.text_input("Nova Senha:", type="password", key="pass_new")
        pass_cnf = st.text_input("Confirmar Nova Senha:", type="password", key="pass_cnf")
        if st.form_submit_button("Atualizar Senha"):
            if not pass_act or not pass_new:
                st.error("Preencha todos os campos.")
            elif pass_new != pass_cnf:
                st.error("As senhas não coincidem.")
            elif len(pass_new) < 3:
                st.error("A nova senha deve ter no mínimo 3 caracteres.")
            else:
                u_id = st.session_state.get("login_user_id", "")
                ok, msg = alterar_senha_usuario(u_id, pass_act, pass_new)
                if ok: st.success(msg)
                else: st.error(msg)

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    if st.button("🔄 Recarregar", key="btn_recarregar_sb"):
        st.cache_data.clear()
        st.rerun()
with col_sb2:
    if st.button("🚪 Sair", key="btn_logout_sb"):
        st.session_state["autenticado"] = False
        st.session_state["usuario_atual"] = None
        st.session_state["perfil_atual"] = None
        st.session_state["login_user_id"] = None
        st.query_params.clear()
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
    if filtro_gerente and "Área Resp. Operação" in df.columns:
        m = m & df["Área Resp. Operação"].isin(filtro_gerente)
    if filtro_cia and "CIA" in df.columns:
        m = m & df["CIA"].isin(filtro_cia)
    return df[m].reset_index(drop=True)

f_falta = aplicar_filtros(df_falta)
f_concur = aplicar_filtros(df_concur)
f_erros = aplicar_filtros(df_erros)
f_backoffice = aplicar_filtros(df_backoffice)
f_eventos = aplicar_filtros(df_eventos)
f_lemon_virt = aplicar_filtros(df_lemon_virt)
f_sem_div = aplicar_filtros(df_sem_div)
df_todos_filtrados = pd.concat([f_falta, f_concur, f_erros, f_backoffice, f_eventos, f_lemon_virt, f_sem_div], ignore_index=True)

# ==============================================================================
# 7. BANNER PRINCIPAL COM FONTES DE DADOS E MÉTRICAS
# ==============================================================================
st.markdown(
    f"""
    <div class='main-header'>
        <div style='display: flex; justify-content: space-between; align-items: flex-start;'>
            <div>
                <h1>✈️ Grupo Arbaitman | Portal Executivo de Conciliação Aérea (FP&A)</h1>
                <p>Painel Estratégico de Revenue Assurance & Gestão Unificada de Pendências</p>
                <div style='margin-top: 10px;'>
                    <span class='fonte-pill'>🏢 ERP Benner (Acumulado)</span>
                    <span class='fonte-pill'>🤖 OBT Lemontech (Extrato)</span>
                    <span class='fonte-pill'>✈️ Cias Aéreas (AZUL, GOL, BSP, HOT)</span>
                    <span class='fonte-pill'>🌐 GDS Sabre</span>
                </div>
            </div>
            <div style='text-align: right;'>
                <span class='update-badge'>🕒 Dados Atualizados em: {dt_ultima_atualizacao_str}</span>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("Pendentes ERP", f"{len(f_falta):,}")
c2.metric("Erro Concur 🚨", f"{len(f_concur):,}")
c3.metric("Erros Val/CIA", f"{len(f_erros):,}")
c4.metric("Backoffice", f"{len(f_backoffice):,}")
c5.metric("Central Eventos", f"{len(f_eventos):,}")
c6.metric("Emissor Virtual", f"{len(f_lemon_virt):,}")
c7.metric("Solicitados Conciliação", f"{len(f_sem_div):,}")

st.markdown("---")

abas = [
    "📊 Dashboard Executivo C-Level",
    "📋 1. Falta de Lançamento (Maringá)",
    "🚨 2. Erro de Integração Concur",
    "⚠️ 3. Erros de Valores & CIA (Maringá)",
    "🎧 4. Suporte Backoffice",
    "🎪 5. Central de Eventos",
    "🤖 6. Emissor Virtual Lemontech",
    "✅ 7. Bilhetes Solicitados Conciliação"
]
if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Histórico", "📥 Carga em Lote Avançada"])
aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# ABA 0: DASHBOARD EXECUTIVO C-LEVEL
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.markdown("<div class='section-banner'><h3>📊 Visão Financeira Global & Indicadores de Performance (FP&A)</h3></div>", unsafe_allow_html=True)
    
    df_pendentes = pd.concat([f_falta, f_concur, f_erros, f_backoffice, f_eventos, f_lemon_virt], ignore_index=True)
    
    tot_processados = len(df_todos_filtrados)
    tot_pendentes = len(df_pendentes)
    tot_conciliados = len(f_sem_div)
    taxa_concil = (tot_conciliados / tot_processados * 100) if tot_processados > 0 else 0.0
    
    tarifa_total_emitida = df_todos_filtrados["Tarifa Total Emitida"].sum() if not df_todos_filtrados.empty and "Tarifa Total Emitida" in df_todos_filtrados.columns else 0.0
    taxa_total_emitida = df_todos_filtrados["Taxa"].sum() if not df_todos_filtrados.empty and "Taxa" in df_todos_filtrados.columns else 0.0
    receita_total_emitida = df_todos_filtrados["Receita Total Emitida"].sum() if not df_todos_filtrados.empty and "Receita Total Emitida" in df_todos_filtrados.columns else 0.0
    ticket_medio_pendente = (df_pendentes["Tarifa Total Emitida"].sum() / len(df_pendentes)) if not df_pendentes.empty else 0.0
    
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Vol. Total Bilhetes", f"{tot_processados:,}")
    k2.metric("Taxa Conciliação", f"{taxa_concil:.1f}%")
    k3.metric("Tarifa Total Emitida", f"R$ {tarifa_total_emitida:,.2f}")
    k4.metric("Taxa Total Emitida", f"R$ {taxa_total_emitida:,.2f}")
    k5.metric("Receita Total Emitida", f"R$ {receita_total_emitida:,.2f}")
    k6.metric("Ticket Médio Pendente", f"R$ {ticket_medio_pendente:,.2f}")
    
    st.markdown("---")
    
    g1, g2 = st.columns(2)
    with g1:
        st.markdown('<div class="chart-card"><h5>📊 Distribuição de Bilhetes Pendentes por Aba do Sistema</h5>', unsafe_allow_html=True)
        dados_abas = [
            {"Aba": "1. Falta Lançamento", "Quantidade": len(f_falta)},
            {"Aba": "2. Erro Concur 🚨", "Quantidade": len(f_concur)},
            {"Aba": "3. Erros de Valores", "Quantidade": len(f_erros)},
            {"Aba": "4. Backoffice", "Quantidade": len(f_backoffice)},
            {"Aba": "5. Central Eventos", "Quantidade": len(f_eventos)},
            {"Aba": "6. Emissor Virtual", "Quantidade": len(f_lemon_virt)}
        ]
        df_abas = pd.DataFrame(dados_abas)
        df_abas = df_abas[df_abas["Quantidade"] > 0]
        if not df_abas.empty:
            fig_pie = px.pie(df_abas, values="Quantidade", names="Aba", hole=0.5, color_discrete_sequence=["#002060", "#d90429", "#fa8c16", "#0050b3", "#0086c0", "#722ed1"])
            fig_pie.update_traces(textposition='inside', textinfo='value+percent', hovertemplate='<b>%{label}</b><br>Volume: %{value:,}<br>Representação: %{percent}')
            fig_pie.update_layout(height=330, margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_pie, use_container_width=True)
        else: st.info("Sem pendências cadastradas.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with g2:
        st.markdown('<div class="chart-card"><h5>✈️ Volume de Pendências por CIA Aérea</h5>', unsafe_allow_html=True)
        if not df_pendentes.empty:
            df_cia = df_pendentes["CIA"].value_counts().head(8).reset_index()
            df_cia.columns = ["Companhia", "Quantidade"]
            fig_bar = px.bar(df_cia, x="Companhia", y="Quantidade", text="Quantidade", color_discrete_sequence=["#002060"])
            fig_bar.update_traces(textposition='outside', texttemplate='%{value:,}', hovertemplate='<b>%{x}</b><br>Pendências: %{y:,}')
            fig_bar.update_layout(height=330, margin=dict(l=10, r=10, t=25, b=10), xaxis_title="", yaxis_title="Qtd Bilhetes")
            st.plotly_chart(fig_bar, use_container_width=True)
        else: st.info("Sem pendências.")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-card"><h5>⏳ Aging / Curva de Envelhecimento das Pendências</h5>', unsafe_allow_html=True)
    if not df_pendentes.empty and "Faixa_Aging" in df_pendentes.columns:
        df_aging = df_pendentes["Faixa_Aging"].value_counts().reset_index()
        df_aging.columns = ["Faixa", "Quantidade"]
        ordem_aging = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "> 90 dias (Crítico)", "Sem Data"]
        df_aging["Faixa"] = pd.Categorical(df_aging["Faixa"], categories=ordem_aging, ordered=True)
        df_aging = df_aging.sort_values("Faixa")
        
        g_ag1, g_ag2 = st.columns([2.2, 1])
        with g_ag1:
            fig_aging = px.bar(df_aging, x="Faixa", y="Quantidade", text="Quantidade", color="Faixa",
                               color_discrete_map={"Até 30 dias": "#52c41a", "31 a 60 dias": "#1890ff", "61 a 90 dias": "#fa8c16", "> 90 dias (Crítico)": "#f5222d", "Sem Data": "#bfbfbf"})
            fig_aging.update_traces(textposition='outside', texttemplate='%{value:,}')
            fig_aging.update_layout(height=300, margin=dict(l=10, r=10, t=25, b=10), showlegend=False, xaxis_title="", yaxis_title="Qtd Bilhetes")
            st.plotly_chart(fig_aging, use_container_width=True)
            
        with g_ag2:
            st.write("**Resumo Financeiro por Aging:**")
            df_pendentes["Valor_Total_Tmp"] = df_pendentes["A vista"] + df_pendentes["A credito"] + df_pendentes["Taxa"]
            df_aging_fin = df_pendentes.groupby("Faixa_Aging", observed=False)["Valor_Total_Tmp"].sum().reset_index()
            df_aging_fin.columns = ["Faixa", "Valor (R$)"]
            df_aging_fin["Valor (R$)"] = df_aging_fin["Valor (R$)"].apply(lambda x: f"R$ {x:,.2f}")
            st.dataframe(df_aging_fin, use_container_width=True, hide_index=True)
    else: st.info("Sem dados de aging disponíveis.")
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# MÓDULO REUTILIZÁVEL DE TABELAS E AÇÕES EM LOTE
# ------------------------------------------------------------------------------
def renderizar_tabela_tratativas(df_tab, nome_base, key_pfx):
    if df_tab.empty:
        st.info("Nenhum bilhete nesta categoria.")
        return
        
    gerente_nome_export = "Geral"
    if filtro_gerente and len(filtro_gerente) == 1:
        gerente_nome_export = filtro_gerente[0]
    elif "Área Resp. Operação" in df_tab.columns:
        unq_g = list(set(df_tab["Área Resp. Operação"].dropna().tolist()))
        if len(unq_g) == 1:
            gerente_nome_export = unq_g[0]
            
    fn_gerente_s = re.sub(r'\s+', '_', gerente_nome_export)
    nome_arquivo_excel = f"Relatorio_Conciliacao_{nome_base}_{fn_gerente_s}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    
    c_s1, c_s2 = st.columns([3, 1])
    with c_s1:
        termo = st.text_input("🔍 Buscar por Bilhete, Localizador (Loc / Rloc), Gerente, Cliente, Emissor ou SV:", key=f"src_{key_pfx}")
    with c_s2:
        st.write("")
        st.download_button(
            "📥 Extrair Relatório Executivo (.xlsx)",
            data=gerar_excel_estilizado(df_tab, nome_aba=nome_base, nome_gerente=gerente_nome_export),
            file_name=nome_arquivo_excel,
            key=f"dl_{key_pfx}"
        )
        
    df_exib = df_tab.copy()
    if termo.strip():
        t_low = termo.strip().lower()
        cols_s = [
            c for c in [
                "Bilhetes", "Rloc_CIA", "Localizador_Sistema", "Localizador", "Rloc_Cia",
                "Área Resp. Operação", "Emissor / Consultor", "Ponto de venda", "Cliente", "SV", "FOP"
            ] if c in df_exib.columns
        ]
        m_s = pd.Series(False, index=df_exib.index)
        for col in cols_s:
            m_s |= df_exib[col].astype(str).str.lower().str.contains(t_low, na=False)
        df_exib = df_exib[m_s]
        
    st.write(f"**Registros Visíveis:** {len(df_exib):,}")
    bilhet_sel = st.multiselect("Selecione Bilhetes para Atualização em Lote:", options=df_exib["Bilhetes"].tolist(), key=f"ms_{key_pfx}")
    
    if bilhet_sel:
        df_sel = df_exib[df_exib["Bilhetes"].isin(bilhet_sel)]
        st.info(f"🎯 **{len(df_sel)} bilhete(s) selecionado(s)** para atualização de tratativa.")
        
        is_single = len(df_sel) == 1
        primeira_linha = df_sel.iloc[0] if is_single else {}
        
        OPCAO_MANTER_ST = "-- Manter Status Original --"
        opts_st = [
            OPCAO_MANTER_ST, "Já Lançado no ERP", "Emitido e Lançado", "Pendente de Lançamento (Não Consta)",
            "Erro de Integração Concur", "Em Análise Backoffice", "Emissor Virtual Lemontech", "Divergência de Tarifa", "Divergência de Taxa", "Cancelado / Reembolsado"
        ]
        idx_st = opts_st.index(str(primeira_linha.get("Status_Geral", "")).strip()) if is_single and str(primeira_linha.get("Status_Geral", "")).strip() in opts_st else 0
            
        OPCAO_MANTER_GER = "-- Manter Gerente Original --"
        opts_g = [OPCAO_MANTER_GER] + sorted(list(set(DEPARA_GERENTES.values())))
        ar_atual = padronizar_nome_gerente(primeira_linha.get("Área Resp. Operação", primeira_linha.get("Gerentes", ""))) if is_single else ""
        idx_g = opts_g.index(ar_atual) if is_single and ar_atual in opts_g else 0
            
        rloc_init = str(primeira_linha.get("Rloc_CIA", "")).strip() if is_single else ""
        if rloc_init in ["-", "None", "nan", "null"]: rloc_init = ""
        
        obs_init = str(primeira_linha.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip() if is_single else ""
        if obs_init in ["-", "None", "nan", "null"]: obs_init = ""
        
        with st.form(f"form_{key_pfx}"):
            c1, c2, c3 = st.columns(3)
            with c1:
                st_novo = st.selectbox("Novo Status:", opts_st, index=idx_st, key=f"st_{key_pfx}")
            with c2:
                ar_nova = st.selectbox("Nova Área Responsável / Gerente:", opts_g, index=idx_g, key=f"ar_{key_pfx}")
            with c3:
                rloc_novo = st.text_input("Novo Rloc_CIA (Em branco mantém original):", value=rloc_init, key=f"rloc_{key_pfx}")
                
            obs_nova = st.text_area("Observação / Justificativa Detalhada:", value=obs_init, key=f"obs_{key_pfx}")
            
            if st.form_submit_button("💾 Salvar Tratativa"):
                usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                list_up = []
                for _, row_item in df_sel.iterrows():
                    st_final = row_item.get("Status_Geral", "Pendente de Lançamento (Não Consta)") if st_novo == OPCAO_MANTER_ST else st_novo
                    ar_final = padronizar_nome_gerente(row_item.get("Área Resp. Operação", row_item.get("Gerentes", "Não Mapeado"))) if ar_nova == OPCAO_MANTER_GER else ar_nova
                    rloc_final = str(row_item.get("Rloc_CIA", "-")).strip() if not rloc_novo.strip() else rloc_novo.strip()
                    obs_final = str(row_item.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip() if not obs_nova.strip() else obs_nova.strip()
                    
                    list_up.append({
                        "Bilhetes": row_item.get("Bilhetes"),
                        "Status_Geral": st_final,
                        "Área Resp. Operação": ar_final,
                        "Obs. Operação": obs_final,
                        "Setor": row_item.get("Setor", "-"),
                        "Gerentes": ar_final,
                        "SV": row_item.get("SV", "-"),
                        "FOP": row_item.get("FOP", "-"),
                        "Rloc_CIA": rloc_final,
                        "Cliente": row_item.get("Cliente", "")
                    })
                    
                df_up = pd.DataFrame(list_up)
                ok, msg, _, _, _ = salvar_tratativas_lote_supabase(df_up, usr_str)
                if ok:
                    st.success(msg)
                    time.sleep(1)
                    st.cache_data.clear()
                    st.rerun()
                else: st.error(msg)

    currency_cols_app = [
        c for c in [
            "A vista", "A credito", "Tarifa Total Emitida", "Taxa", "Comissão", "Taxa DU", "Incentivo",
            "Receita Total Emitida", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa",
            "Receita_Sistema", "Dif_Receita"
        ] if c in df_exib.columns
    ]
    def formatar_moeda_br(val):
        if pd.isna(val) or val is None or str(val).strip() in ["-", "", "nan", "None", "<na>"]:
            return "-"
        try:
            v = float(val)
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError): return str(val)

    df_styled = df_exib.style
    if currency_cols_app:
        df_styled = df_styled.format(formatar_moeda_br, subset=currency_cols_app)
    st.dataframe(df_styled, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# ABAS OPERACIONAIS
# ------------------------------------------------------------------------------
with aba_sel[1]:
    st.markdown("<div class='section-banner'><h3>📋 1. Falta de Lançamento no ERP (Maringá)</h3><p>Emissões pendentes de lançamento no ERP Benner.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_falta, "Falta_Lancamento", "fl")

with aba_sel[2]:
    st.markdown("<div class='section-banner'><h3>🚨 2. Erro de Integração Concur</h3><p>Falhas de integração com a plataforma Concur.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_concur, "Erro_Concur", "conc")

with aba_sel[3]:
    st.markdown("<div class='section-banner'><h3>⚠️ 3. Erros de Valores & CIA (Maringá)</h3><p>Divergências de Tarifa, Taxa ou Companhia Aérea.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_erros, "Erros_Valores", "ev")

with aba_sel[4]:
    st.markdown("<div class='section-banner'><h3>🎧 4. Suporte Backoffice</h3><p>Análise da equipe de Backoffice.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_backoffice, "Suporte_Backoffice", "sb")

with aba_sel[5]:
    st.markdown("<div class='section-banner'><h3>🎪 5. Central de Eventos</h3><p>Emissões do setor de Eventos.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_eventos, "Central_Eventos", "ce")

with aba_sel[6]:
    st.markdown("<div class='section-banner'><h3>🤖 6. Emissor Virtual Lemontech</h3><p>Emissões via robô Lemontech.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_lemon_virt, "Emissor_Virtual", "evl")

with aba_sel[7]:
    st.markdown("<div class='section-banner'><h3>✅ 7. Bilhetes Solicitados Conciliação</h3><p>Bilhetes conciliados sem divergências.</p></div>", unsafe_allow_html=True)
    renderizar_tabela_tratativas(f_sem_div, "Bilhetes_Conciliados", "ok")

# ------------------------------------------------------------------------------
# ABAS MASTER
# ------------------------------------------------------------------------------
if e_master():
    with aba_sel[8]:
        st.subheader("⚙️ Gestão de Acessos & Frequência de Uso")
        df_gestao_usr, df_historico_acessos = obter_relatorio_acessos_db()
        if not df_gestao_usr.empty:
            st.dataframe(df_gestao_usr, use_container_width=True, hide_index=True)
        else: st.info("Sem dados de acessos.")

    with aba_sel[9]:
        st.subheader("📜 Log de Histórico de Auditoria")
        _, df_logs_det = obter_relatorio_acessos_db()
        if not df_logs_det.empty:
            st.dataframe(df_logs_det, use_container_width=True, hide_index=True)

    with aba_sel[10]:
        st.subheader("📥 Carga de Retornos de Cobrança em Lote Avançada")
        arq_up = st.file_uploader("Arquivo de retorno (.xlsx ou .csv):", type=["xlsx", "csv"], key="upl_lote_adv")
        if arq_up:
            try:
                df_lote_raw = pd.read_csv(arq_up, dtype=str) if arq_up.name.endswith(".csv") else pd.read_excel(arq_up, dtype=str)
                st.dataframe(df_lote_raw.head(10), use_container_width=True, hide_index=True)
                if st.button("🚀 Processar Lote", type="primary"):
                    usr_str = f"CargaLote_{st.session_state['usuario_atual']}"
                    ok, msg, n_at, n_nv, _ = salvar_tratativas_lote_supabase(df_lote_raw, usr_str)
                    if ok:
                        st.success(msg)
                        st.cache_data.clear()
                        st.rerun()
            except Exception as e: st.error(f"Erro ao ler arquivo: {e}")