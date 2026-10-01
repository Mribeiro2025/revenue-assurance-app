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
# 1. CONFIGURAÇÃO INICIAL E ESTILOS CSS
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Portal de Conciliação Aérea & Bilhetes Pendentes",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)
ARQUIVO_DASHBOARD = "Dashboard_Revenue_Assurance_Consolidado.xlsx"

st.markdown(
    """
    <style>
        .stApp { background-color: #f8f9fa; }
        .main-header {
            background: linear-gradient(135deg, #002060 0%, #003366 100%);
            padding: 22px 28px;
            border-radius: 12px;
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        .main-header h1 { color: #ffffff !important; margin: 0; font-size: 24px; font-weight: 700; }
        .main-header p { color: #d0e1fd !important; margin: 4px 0 0 0; font-size: 13px; }
        
        .section-banner {
            background-color: #ffffff;
            border-left: 6px solid #002060;
            padding: 14px 20px;
            border-radius: 8px;
            margin-bottom: 18px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        }
        .section-banner h3 { margin: 0; color: #002060; font-size: 18px; font-weight: 600; }
        .section-banner p { margin: 3px 0 0 0; color: #666; font-size: 12px; }
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
        .highlight-card {
            background-color: #eef6ff;
            border: 2px solid #004080;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 15px;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# ==============================================================================
# 2. FUNÇÃO DE SANITIZAÇÃO E CONEXÃO COM SUPABASE
# ==============================================================================
def sanitizar_bilhete(val):
    """Padroniza identificadores removendo colchetes, decimais (.0) e espaços."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"[\[\]]", "", s)
    s = re.sub(r"\.0$", "", s)
    return s.strip()

def get_db_engine():
    """Conecta ao Supabase com suporte aos Secrets e driver SQLAlchemy."""
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
    """Lê do Supabase e mantém o estado mais recente por bilhete."""
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
        st.error(f"⚠️ Erro ao consultar a tabela 'tratativas' no Supabase: {e}")
        return pd.DataFrame(columns=["bilhete", "status_geral", "area_resp", "obs_operacao", "setor", "gerentes", "data_modificacao"])

def salvar_tratativas_lote_supabase(df_lote, usuario):
    """
    Executa UPSERT na tabela tratativas do Supabase.
    Atualiza bilhetes existentes e insere novos registros.
    """
    engine = get_db_engine()
    if not engine or df_lote.empty:
        return False, "Conexão com a base de dados indisponível.", 0, 0, 0

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

    sucessos = 0
    erros = 0

    for item in dados_lote:
        try:
            with engine.connect() as conn:
                conn.execute(sql_upsert, item)
                conn.commit()
                sucessos += 1
        except Exception:
            erros += 1

    st.cache_data.clear()
    msg = f"✅ Sincronização concluída! {qtd_atualizados} bilhete(s) atualizado(s) e {qtd_novos} novo(s) inserido(s)."
    return True, msg, qtd_atualizados, qtd_novos, erros

def registrar_log_supabase(logs_list):
    """Grava o histórico na tabela log_auditoria do Supabase."""
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
        with engine.begin() as conn:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS log_auditoria (
                    id SERIAL PRIMARY KEY,
                    data_hora TIMESTAMP,
                    bilhete VARCHAR(255),
                    usuario_acao VARCHAR(255),
                    status_anterior VARCHAR(255),
                    novo_status VARCHAR(255),
                    area_anterior VARCHAR(255),
                    nova_area VARCHAR(255),
                    observacao TEXT,
                    tipo_interacao VARCHAR(255)
                );
            """))
        df_logs.to_sql("log_auditoria", engine, if_exists="append", index=False)
    except Exception:
        pass

# ==============================================================================
# 3. GESTÃO DE USUÁRIOS
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
                        "data_solicitacao": str(v.get("data_solicitacao", ""))
                    }
        except Exception:
            pass
            
    engine = get_db_engine()
    if engine:
        query = "SELECT usuario, senha, nome, perfil, status, data_solicitacao FROM usuarios"
        try:
            df_u = pd.read_sql(query, engine)
            for _, row in df_u.iterrows():
                u_key = str(row["usuario"]).strip().lower()
                dict_u[u_key] = {
                    "senha": str(row["senha"]).strip(),
                    "nome": str(row["nome"]).strip(),
                    "perfil": str(row["perfil"]).strip(),
                    "status": str(row["status"]).strip().upper(),
                    "data_solicitacao": str(row["data_solicitacao"])
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
            senha = EXCLUDED.senha,
            nome = EXCLUDED.nome,
            perfil = EXCLUDED.perfil,
            status = EXCLUDED.status;
    """)
    try:
        with engine.begin() as conn:
            conn.execute(sql, {"u": u_id, "p": senha, "n": nome, "perf": perfil, "st": status})
        return True, "Operação realizada com sucesso."
    except Exception as e:
        return False, f"Erro no banco: {e}"

def atualizar_senha_supabase(u_id, nova_senha):
    engine = get_db_engine()
    if not engine:
        return False, "Conexão com a base de dados indisponível."
    
    sql = text("UPDATE usuarios SET senha = :p WHERE usuario = :u;")
    try:
        with engine.begin() as conn:
            conn.execute(sql, {"u": u_id, "p": nova_senha})
        return True, "Senha redefinida com sucesso."
    except Exception as e:
        return False, f"Erro no banco: {e}"

def atualizar_status_usuario_supabase(u_id, novo_status):
    engine = get_db_engine()
    if not engine:
        return False
    
    sql = text("UPDATE usuarios SET status = :st WHERE usuario = :u;")
    try:
        with engine.begin() as conn:
            conn.execute(sql, {"u": u_id, "st": novo_status})
        return True
    except Exception:
        return False

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
        tab_log, tab_pwd, tab_req = st.tabs(["🔐 Entrar", "🔑 Esqueci a Senha", "📝 Solicitar Acesso"])
        
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
                elif u_input in ["fellipe", "ffernandes"] and p_input == "123":
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = "Fellipe Fernandes"
                    st.session_state["perfil_atual"] = "Master"
                    st.session_state["login_user_id"] = u_input
                    st.rerun()
                elif u_input == "compliance1" and p_input == "123":
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = "Compliance FP&A"
                    st.session_state["perfil_atual"] = "Compliance"
                    st.session_state["login_user_id"] = "compliance1"
                    st.rerun()
                elif u_input == "backoffice" and p_input == "123":
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = "Atendimento Backoffice"
                    st.session_state["perfil_atual"] = "Operacao"
                    st.session_state["login_user_id"] = "backoffice"
                    st.rerun()
                elif u_input == "operacao" and p_input == "123":
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_atual"] = "Equipe Operacional"
                    st.session_state["perfil_atual"] = "Operacao"
                    st.session_state["login_user_id"] = "operacao"
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        with tab_pwd:
            u_reset = st.text_input("Usuário cadastrado:").strip().lower()
            n_pass = str(st.text_input("Nova Senha Alfanumérica:", type="password")).strip()
            c_pass = str(st.text_input("Confirme a Nova Senha:", type="password")).strip()
            if st.button("Redefinir Senha"):
                usuarios_db = carregar_usuarios_supabase()
                if u_reset in usuarios_db and n_pass and n_pass == c_pass:
                    ok, msg = atualizar_senha_supabase(u_reset, n_pass)
                    if ok:
                        st.success("✅ Senha redefinida no Supabase!")
                    else:
                        st.error(f"Erro ao atualizar: {msg}")
                else:
                    st.error("Verifique se o usuário existe e se as senhas coincidem.")
        with tab_req:
            r_nome = st.text_input("Nome Completo:")
            r_user = st.text_input("Usuário Desejado:").strip().lower()
            r_pass = str(st.text_input("Senha de Acesso:", type="password")).strip()
            r_perf = st.selectbox("Perfil Solicitado:", ["Operacao"])
            if st.button("Enviar Solicitação"):
                if r_user and r_pass and r_nome:
                    ok, msg = salvar_usuario_supabase(r_user, r_pass, r_nome, r_perf, status="PENDENTE")
                    if ok:
                        st.success("✅ Solicitação enviada!")
                    else:
                        st.error(f"Erro ao salvar solicitação: {msg}")
    st.stop()

# ==============================================================================
# 4. SOBREPOSIÇÃO DOS DADOS DO SUPABASE SOBRE O RELATÓRIO BRUTO
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
            "Emissor_Reserva_Lemon", "Consultor_Lemon", "Consultor", "Status_Sistema", 
            "Status_Divergencia", "Gerentes", "Fornecedor_Sistema", "Status_Cia"
        ])
    
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    
    if "Bilhetes" not in df_out.columns:
        cols_b = [c for c in df_out.columns if "bilhete" in str(c).lower() or "ticket" in str(c).lower()]
        df_out = df_out.rename(columns={cols_b[0]: "Bilhetes"}) if cols_b else df_out.assign(Bilhetes="")
    
    for c in df_out.columns:
        if df_out[c].dtype == object or df_out[c].dtype == 'string':
            df_out[c] = df_out[c].apply(clean_str)
            
    if "Obs. Operação" in df_out.columns:
        df_out["Obs. Operação"] = df_out["Obs. Operação"].astype(str).str.replace("Sem tratativa na operação", "", regex=False).str.strip()
        
    status_map = ["nao_consta", "não_consta", "nan", "-", "", "none"]
    df_out["Status_Geral"] = df_out["Status_Geral"].apply(
        lambda x: "Pendente de Lançamento (Não Consta)" if str(x).lower().strip() in status_map else str(x)
    )
    for c in ["Taxa", "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa", "Receita_Sistema", "Dif_Receita", "Tarifa_Total"]:
        if c in df_out.columns: 
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)
            
    df_out["Dt_Parsed"] = pd.to_datetime(df_out["Data Emissão"], format="mixed", dayfirst=True, errors="coerce")
    df_out["Dt_Mod_Parsed"] = pd.to_datetime(df_out.get("Data_Modificacao"), format="mixed", dayfirst=True, errors="coerce")
    
    return df_out

def mesclar_com_supabase(df_excel):
    if df_excel is None or df_excel.empty:
        return df_excel
    df_db = carregar_tratativas_db()
    if df_db.empty:
        return df_excel
    
    df_excel["Bilhete_Clean"] = df_excel["Bilhetes"].apply(sanitizar_bilhete)
    df_db["Bilhete_Clean"] = df_db["bilhete"].apply(sanitizar_bilhete)
    
    if "data_modificacao" in df_db.columns:
        df_db = df_db.sort_values("data_modificacao", ascending=False)
    df_db = df_db.drop_duplicates("Bilhete_Clean", keep="first")
    
    cols_sup_desejadas = ["status_geral", "area_resp", "obs_operacao", "setor", "gerentes"]
    cols_presentes = [c for c in cols_sup_desejadas if c in df_db.columns]
    
    df_merged = pd.merge(
        df_excel,
        df_db[["Bilhete_Clean"] + cols_presentes],
        on="Bilhete_Clean",
        how="left"
    )
    
    muta_map = [
        ("Status_Geral", "status_geral"),
        ("Área Resp. Operação", "area_resp"),
        ("Obs. Operação", "obs_operacao"),
        ("Setor", "setor"),
        ("Gerentes", "gerentes")
    ]
    for col_excel, col_sb in muta_map:
        if col_sb in df_merged.columns and col_excel in df_merged.columns:
            mask_valido = df_merged[col_sb].notna() & ~df_merged[col_sb].astype(str).str.strip().isin(["-", "", "None", "nan", "NULL", "EMPTY"])
            df_merged[col_excel] = df_merged[col_sb].where(mask_valido, df_merged[col_excel])
            
    cols_drop = [c for c in cols_presentes if c in df_merged.columns]
    df_merged.drop(columns=cols_drop, inplace=True, errors="ignore")
    return df_merged

def rotear_bases_mestra(df_master):
    if df_master is None or df_master.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_master = padronizar_df(df_master)
    def e_apenas_divergencia_receita(st_div):
        s = str(st_div).lower().strip()
        if "diverg" not in s and "erro" not in s:
            return False
        s_sem_diverg = s.replace("divergência", "").replace("divergencia", "").strip()
        return "receita" in s_sem_diverg and not any(x in s_sem_diverg for x in ["tarifa", "taxa", "cia", "companhia"])
        
    st_geral_str = df_master["Status_Geral"].astype(str).str.lower()
    st_div_str = df_master["Status_Divergencia"].astype(str).str.lower()
    
    mask_tem_divergencia_real = st_div_str.str.contains("divergência|divergencia|erro", na=False) & \
                                ~df_master["Status_Divergencia"].apply(e_apenas_divergencia_receita)
                                
    mask_status_lancado = (
        st_geral_str.str.contains("já lançado|emitido e lançado|lançado|conciliado|valores corretos|regularizado", na=False) |
        st_div_str.str.contains("valores corretos|sem divergência|sem divergencia", na=False)
    )
    
    mask_ok = (mask_status_lancado | df_master["Status_Divergencia"].apply(e_apenas_divergencia_receita)) & ~mask_tem_divergencia_real
    
    df_ok = df_master[mask_ok].copy()
    df_rest = df_master[~mask_ok].copy()
    
    def e_bo_flexivel(r):
        ar_val = str(r.get("Área Resp. Operação", "")).strip().lower()
        obs_val = str(r.get("Obs. Operação", "")).strip().lower()
        st_val = str(r.get("Status_Geral", "")).strip().lower()
        ger_val = str(r.get("Gerentes", "")).strip().lower()
        setor_val = str(r.get("Setor", "")).strip().lower()
        origem_val = str(r.get("Aba_Origem", "")).strip().lower()
        return any(k in ar_val or k in ger_val or k in setor_val or k in origem_val or k in obs_val or k in st_val
                   for k in ["katia", "kátia", "backoffice", "suporte"])
                    
    mask_bo = df_rest.apply(e_bo_flexivel, axis=1)
    df_bo = df_rest[mask_bo].copy()
    df_rest = df_rest[~mask_bo].copy()
    
    mask_evt = (
        df_rest["Setor"].astype(str).str.lower().str.contains("eventos", na=False) |
        df_rest["Área Resp. Operação"].astype(str).str.lower().str.contains("eventos", na=False) |
        df_rest["Gerentes"].astype(str).str.lower().str.contains("eventos", na=False)
    )
    df_eventos = df_rest[mask_evt].copy()
    df_rest = df_rest[~mask_evt].copy()
    
    def e_emissor_virtual(r):
        for col in ["Emissor", "Emissor_Reserva_Lemon", "Consultor_Lemon", "Consultor"]:
            val = str(r.get(col, "")).lower().strip()
            if "virtual" in val or "lemontech" in val:
                return True
        return False
        
    mask_lemon = df_rest.apply(e_emissor_virtual, axis=1)
    df_lemon_virt = df_rest[mask_lemon].copy()
    df_rest = df_rest[~mask_lemon].copy()
    
    mask_consta_benner = ~df_rest["Status_Sistema"].astype(str).str.upper().str.contains("NAO_CONSTA|NÃO_CONSTA", na=False)
    mask_tem_divergencia = df_rest["Status_Divergencia"].astype(str).str.lower().str.contains("divergência|divergencia|erro", na=False)
    
    mask_erros = mask_consta_benner | mask_tem_divergencia
    df_erros = df_rest[mask_erros].copy()
    df_falta = df_rest[~mask_erros].copy()
    
    return df_falta, df_erros, df_bo, df_eventos, df_lemon_virt, df_ok

@st.cache_data(ttl=15)
def carregar_bases():
    if not os.path.exists(ARQUIVO_DASHBOARD):
        vazio = padronizar_df(None)
        return vazio, vazio, vazio, vazio, vazio, vazio, pd.DataFrame()
    
    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD, engine="openpyxl")
        frames = []
        
        for sheet in xls.sheet_names:
            if sheet.startswith("99_") or sheet.startswith("98_"):
                df_sheet = pd.read_excel(xls, sheet_name=sheet)
                df_sheet["Aba_Origem"] = sheet
                frames.append(df_sheet)
                
        df_log_arq = pd.read_excel(xls, "00_Log_Auditoria") if "00_Log_Auditoria" in xls.sheet_names else pd.DataFrame()
        
        if not frames:
            vazio = padronizar_df(None)
            return vazio, vazio, vazio, vazio, vazio, vazio, df_log_arq
            
        df_m = pd.concat(frames, ignore_index=True)
        
        if "Bilhetes" in df_m.columns:
            df_m["Bilhete_Clean"] = df_m["Bilhetes"].apply(sanitizar_bilhete)
            df_m = df_m.drop_duplicates(subset=["Bilhete_Clean"], keep="first").reset_index(drop=True)
            
        df_m = mesclar_com_supabase(df_m)
        f_falta, f_erros, f_bo, f_evt, f_lem, f_ok = rotear_bases_mestra(df_m)
        return f_falta, f_erros, f_bo, f_evt, f_lem, f_ok, df_log_arq
    except Exception as e:
        st.error(f"⚠️ Erro ao carregar as bases do Dashboard: {e}")
        vazio = padronizar_df(None)
        return vazio, vazio, vazio, vazio, vazio, vazio, pd.DataFrame()

def gerar_excel_estilizado(df_export, nome_aba="Relatorio"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(columns=["Aviso"], data=[["Nenhum registro encontrado para os filtros selecionados"]])
    
    cols_remover = [c for c in ["Dt_Parsed", "Dt_Mod_Parsed", "Aba_Origem", "Bilhete_Clean"] if c in df_export.columns]
    df_clean = df_export.drop(columns=cols_remover) if cols_remover else df_export.copy()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name=nome_aba[:30], index=False)
    
    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    ws = wb.active
    ws.views.sheetView[0].showGridLines = True
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )
    fill_alerta = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_alerta = Font(color="9C0006", bold=True)
    currency_cols = ["A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa", "Taxa_Sistema", "Dif_Taxa", "Comissão", "Taxa DU", "Desc.", "Incentivo", "Receita_Sistema", "Dif_Receita", "VL. Líquido", "Tarifa_Total"]
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            col_name = str(ws.cell(row=1, column=cell.column).value or "")
            if col_name in currency_cols and cell.value not in ["-", None, ""]:
                try:
                    cell.value = float(cell.value)
                    cell.number_format = "R$ #,##0.00"
                except (ValueError, TypeError):
                    pass
            if col_name.startswith("Dif_") and cell.value not in ["-", None, ""]:
                try:
                    if abs(float(cell.value)) >= 0.01:
                        cell.fill = fill_alerta
                        cell.font = font_alerta
                except (ValueError, TypeError):
                    pass
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        len_vals = [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, min(ws.max_row + 1, 100))]
        max_len = max(len_vals) if len_vals else 10
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()

with st.spinner("🔄 Conectando ao Supabase e recarregando o painel..."):
    df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div, df_log = carregar_bases()

# ==============================================================================
# 5. SIDEBAR E FILTROS OPERACIONAIS
# ==============================================================================
st.sidebar.title("Grupo Arbaitman")
st.sidebar.caption("Conciliação Aérea FP&A v3.5")
st.sidebar.write(f"👤 **{st.session_state['usuario_atual']}** ({st.session_state['perfil_atual']})")
with st.sidebar.expander("🔑 Alterar Minha Senha"):
    with st.form("form_pwd_side"):
        s_atu = str(st.text_input("Senha Atual:", type="password")).strip()
        s_nov = str(st.text_input("Nova Senha Alfanumérica:", type="password")).strip()
        if st.form_submit_button("Salvar Nova Senha"):
            u_id = st.session_state["login_user_id"]
            usuarios_db = carregar_usuarios_supabase()
            if u_id in usuarios_db and usuarios_db[u_id]["senha"] == s_atu and s_nov:
                ok, msg = atualizar_senha_supabase(u_id, s_nov)
                if ok:
                    st.success("✅ Senha alterada com sucesso!")
                else:
                    st.error(f"Erro ao salvar senha: {msg}")
            else:
                st.error("Senha atual incorreta.")
if st.sidebar.button("🔒 Sair"):
    st.session_state["autenticado"] = False
    st.rerun()
st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Filtros Operacionais")
d_inicio = st.sidebar.date_input("Data Inicial:", value=datetime.date(2024, 1, 1), format="DD/MM/YYYY")
d_fim = st.sidebar.date_input("Data Final:", value=datetime.date(2026, 12, 31), format="DD/MM/YYYY")
df_todos = pd.concat([df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div], ignore_index=True)
filtro_gerente = st.sidebar.multiselect("Gerente / Área Resp.:", options=sorted(df_todos["Área Resp. Operação"].dropna().unique()), placeholder="Todos")
filtro_setor = st.sidebar.multiselect("Setor:", options=sorted(df_todos["Setor"].dropna().unique()), placeholder="Todos")
filtro_cia = st.sidebar.multiselect("Companhia Aérea:", options=sorted(df_todos["CIA"].dropna().unique()), placeholder="Todas")

def aplicar_filtros(df):
    if df is None or df.empty:
        return df
    m = pd.Series(True, index=df.index)
    
    if "Dt_Parsed" in df.columns and d_inicio and d_fim:
        has_dt = df["Dt_Parsed"].notna()
        m_dt = pd.Series(True, index=df.index)
        if has_dt.any():
            dt_dates = df.loc[has_dt, "Dt_Parsed"].dt.date
            m_dt.loc[has_dt] = (dt_dates >= d_inicio) & (dt_dates <= d_fim)
        m = m & m_dt
    if filtro_gerente:
        m = m & df["Área Resp. Operação"].isin(filtro_gerente)
    if filtro_setor:
        m = m & df["Setor"].isin(filtro_setor)
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
# 6. HEADER PRINCIPAL (COM MARCA ATUALIZADA)
# ==============================================================================
st.markdown(
    """
    <div class='main-header'>
        <h1>✈️ Grupo Arbaitman | Portal de Conciliação Aérea & Gestão de Bilhetes Pendentes (FP&A)</h1>
        <p>Plataforma Executiva para Gestão de Divergências, Lançamentos em ERP e Monitoramento de SLA de Resolução</p>
    </div>
    """,
    unsafe_allow_html=True
)
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Pendentes de ERP", f"{len(f_falta):,}")
c2.metric("Erros Valores/CIA", f"{len(f_erros):,}")
c3.metric("Backoffice", f"{len(f_backoffice):,}")
c4.metric("Central de Eventos", f"{len(f_eventos):,}")
c5.metric("Emissor Virtual", f"{len(f_lemon_virt):,}")
c6.metric("Conciliados (OK)", f"{len(f_sem_div):,}")
st.markdown("---")

# ==============================================================================
# 7. ESTRUTURA DE ABAS
# ==============================================================================
abas = [
    "📊 Dashboard Executivo",
    "📋 1. Falta de Lançamento",
    "⚠️ 2. Erros de Valores & CIA",
    "🎧 3. Suporte Backoffice",
    "🎪 4. Central de Eventos",
    "🤖 5. Emissor Virtual Lemontech",
    "✅ 6. Bilhetes Conciliados (Sem Divergência)"
]
if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Histórico", "📥 Carga em Lote Avançada"])
aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# ABA 0: DASHBOARD EXECUTIVO EXPANDIDO COM MAIS KPIS
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.markdown(
        """
        <div class='section-banner'>
            <h3>📊 Painel Executivo C-Level & Desempenho de Conciliação</h3>
            <p>Visão estratégica do volume financeiro pendente, taxa de resolução, SLAs operacionais e pendências por companhia aérea.</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    df_pendentes_todas = pd.concat([f_falta, f_erros, f_backoffice, f_eventos, f_lemon_virt], ignore_index=True)
    df_todas_casos = pd.concat([df_pendentes_todas, f_sem_div], ignore_index=True)
    
    total_casos = len(df_todas_casos)
    total_conciliado = len(f_sem_div)
    total_pendente = len(df_pendentes_todas)
    taxa_resolucao = (total_conciliado / total_casos * 100) if total_casos > 0 else 0.0
    
    # Cálculo do Montante Financeiro Pendente (Tarifa + Taxas)
    val_pendente_total = 0.0
    for col_v in ["Tarifa_Total", "A vista", "A credito", "Tarifa_Sistema"]:
        if col_v in df_pendentes_todas.columns:
            val_pendente_total += df_pendentes_todas[col_v].sum()
            break
            
    df_sla = f_sem_div.copy()
    tempo_medio_dias = 0.0
    if not df_sla.empty and "Dt_Parsed" in df_sla.columns and "Dt_Mod_Parsed" in df_sla.columns:
        df_sla["Dias_Resolucao"] = (df_sla["Dt_Mod_Parsed"] - df_sla["Dt_Parsed"]).dt.days
        df_sla_valido = df_sla[df_sla["Dias_Resolucao"] >= 0]
        if not df_sla_valido.empty:
            tempo_medio_dias = df_sla_valido["Dias_Resolucao"].mean()
            
    # Linha 1 de KPIs Executivos
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Total Bilhetes Processados", f"{total_casos:,}")
    kpi2.metric("Bilhetes Conciliados (OK)", f"{total_conciliado:,}")
    kpi3.metric("Bilhetes Pendentes", f"{total_pendente:,}")
    kpi4.metric("Índice de Conciliação (%)", f"{taxa_resolucao:.1f}%")
    kpi5.metric("Valor Total Pendente (R$)", f"R$ {val_pendente_total:,.2f}")
    
    st.markdown("---")
    
    # Linha de Gráficos Executivos
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.markdown("##### ⚠️ Distribuição das Pendências por Status")
        if not df_pendentes_todas.empty and "Status_Geral" in df_pendentes_todas.columns:
            df_st_chart = df_pendentes_todas["Status_Geral"].value_counts().reset_index()
            df_st_chart.columns = ["Status", "Quantidade"]
            fig_pie = px.pie(df_st_chart, values="Quantidade", names="Status", hole=0.4, color_discrete_sequence=px.colors.qualitative.Bold)
            fig_pie.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=320)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Nenhuma pendência registrada para os filtros selecionados.")
            
    with col_d2:
        st.markdown("##### ✈️ Pendências por Companhia Aérea (Top 8)")
        if not df_pendentes_todas.empty and "CIA" in df_pendentes_todas.columns:
            df_cia_chart = df_pendentes_todas["CIA"].value_counts().head(8).reset_index()
            df_cia_chart.columns = ["Companhia", "Bilhetes Pendentes"]
            fig_cia = px.bar(df_cia_chart, x="Companhia", y="Bilhetes Pendentes", text="Bilhetes Pendentes", color_discrete_sequence=["#003366"])
            fig_cia.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=320)
            st.plotly_chart(fig_cia, use_container_width=True)
            
    col_d3, col_d4 = st.columns(2)
    with col_d3:
        st.markdown("##### 👤 Top 8 Gerentes por Volume de Pendências")
        if not df_pendentes_todas.empty and "Área Resp. Operação" in df_pendentes_todas.columns:
            df_ger_chart = df_pendentes_todas["Área Resp. Operação"].value_counts().head(8).reset_index()
            df_ger_chart.columns = ["Gerente", "Quantidade"]
            fig_ger = px.bar(df_ger_chart, x="Gerente", y="Quantidade", text="Quantidade", color_discrete_sequence=["#002060"])
            fig_ger.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=300)
            st.plotly_chart(fig_ger, use_container_width=True)
            
    with col_d4:
        st.markdown("##### ⏱️ SLA Médio de Solução por Gerente (Dias)")
        if not df_sla.empty and "Dias_Resolucao" in df_sla.columns and "Área Resp. Operação" in df_sla.columns:
            df_sla_ger = df_sla.groupby("Área Resp. Operação")["Dias_Resolucao"].mean().reset_index()
            df_sla_ger.columns = ["Gerente", "Dias_Medios"]
            df_sla_ger = df_sla_ger.sort_values("Dias_Medios", ascending=False).head(8)
            fig_sla = px.bar(df_sla_ger, x="Gerente", y="Dias_Medios", text_auto=".1f", color_discrete_sequence=["#d90429"])
            fig_sla.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=300)
            st.plotly_chart(fig_sla, use_container_width=True)
        else:
            st.info("Aguardando mais atualizações salvas para calcular a média histórica de SLA.")

# ------------------------------------------------------------------------------
# FUNÇÃO REUTILIZÁVEL DE TRATATIVA LINHA A LINHA
# ------------------------------------------------------------------------------
def renderizar_modulo_tratativa(df_filtrado, nome_base, key_prefix):
    if df_filtrado.empty:
        st.info("Nenhum bilhete pendente nesta categoria.")
        return
        
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        termo_busca = st.text_input("🔍 Buscar nesta tela (por Bilhete, LOC, Cliente, Passageiro ou Gerente):", key=f"src_{key_prefix}")
    with col_h2:
        st.write("")
        st.write("")
        st.download_button(
            f"📥 Extrair Relatório ({nome_base})",
            data=gerar_excel_estilizado(df_filtrado, nome_base),
            file_name=f"{nome_base}_Filtrado.xlsx",
            key=f"btn_dl_{key_prefix}"
        )
        
    df_exib = df_filtrado.copy()
    if termo_busca.strip():
        term = termo_busca.strip().lower()
        cols_search = [c for c in ["Bilhetes", "Localizador_Sistema", "Rloc_Cia", "Ponto de venda", "Gerentes", "Área Resp. Operação", "Consultor", "Cliente", "Setor"] if c in df_exib.columns]
        mask_src = pd.Series(False, index=df_exib.index)
        for col in cols_search:
            mask_src |= df_exib[col].astype(str).str.lower().str.contains(term, na=False)
        df_exib = df_exib[mask_src]
        
    st.markdown(f"**Registros Visíveis:** {len(df_exib)}")
    bilhetes_lista = df_exib["Bilhetes"].tolist() if "Bilhetes" in df_exib.columns else []
    
    bilhet_sel = st.multiselect(
        "Selecione um ou mais Bilhetes para Tratativa Linha a Linha:",
        options=bilhetes_lista,
        placeholder="Selecione um ou mais bilhetes para atualizar...",
        key=f"ms_{key_prefix}"
    )
    
    if bilhet_sel:
        df_sel_cards = df_exib[df_exib["Bilhetes"].isin(bilhet_sel)]
        st.markdown('<div class="highlight-card">', unsafe_allow_html=True)
        st.markdown(f"#### 🎯 Informações do(s) Bilhete(s) Selecionado(s) ({len(df_sel_cards)})")
        for _, r_card in df_sel_cards.iterrows():
            obs_card = str(r_card.get('Obs. Operação', '')).replace("Sem tratativa na operação", "").strip() or '-'
            st.markdown(f"""
            * **Bilhete:** `{r_card.get('Bilhetes')}` | **CIA:** {r_card.get('CIA')} | **Ponto Venda:** {r_card.get('Ponto de venda')}
            * **Gerente Atual:** {r_card.get('Área Resp. Operação')} | **Setor:** {r_card.get('Setor')} | **Status Atual:** `{r_card.get('Status_Geral')}`
            * **Obs. Atual:** {obs_card}
            """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        df_primeiro = df_sel_cards.iloc[0]
        with st.form(f"form_trat_{key_prefix}"):
            c_s1, c_s2 = st.columns(2)
            with c_s1:
                opcoes_factiveis = [
                    "Já Lançado no ERP",
                    "Emitido e Lançado",
                    "Pendente de Lançamento (Não Consta)",
                    "Em Análise Backoffice",
                    "Divergência de Tarifa",
                    "Divergência de Taxa",
                    "Divergência de CIA Aérea",
                    "Cancelado / Reembolsado"
                ]
                status_atual_val = str(df_primeiro.get("Status_Geral", "Pendente de Lançamento (Não Consta)"))
                idx_default = opcoes_factiveis.index(status_atual_val) if status_atual_val in opcoes_factiveis else 2
                n_status = st.selectbox("Novo Status (Obrigatório):", options=opcoes_factiveis, index=idx_default, key=f"st_{key_prefix}")
            
            with c_s2:
                lista_gerentes = sorted(list(set(df_todos["Área Resp. Operação"].dropna().unique()))) if "Área Resp. Operação" in df_todos.columns else ["Operação"]
                idx_ger = lista_gerentes.index(df_primeiro.get("Área Resp. Operação")) if df_primeiro.get("Área Resp. Operação") in lista_gerentes else 0
                n_area = st.selectbox("Nova Área Responsável / Gerente:", options=lista_gerentes, index=idx_ger, key=f"ar_{key_prefix}")
            
            obs_init = str(df_primeiro.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip()
            n_obs = st.text_area("Observação / Justificativa Detalhada:", value=obs_init, key=f"obs_{key_prefix}")
            
            if st.form_submit_button("💾 Salvar Tratativa no Supabase"):
                usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                df_up = pd.DataFrame({
                    "Bilhetes": df_sel_cards["Bilhetes"].values,
                    "Status_Geral": [n_status] * len(df_sel_cards),
                    "Área Resp. Operação": [n_area] * len(df_sel_cards),
                    "Obs. Operação": [n_obs] * len(df_sel_cards),
                    "Setor": df_sel_cards["Setor"].values,
                    "Gerentes": [n_area] * len(df_sel_cards)
                })
                
                ok, msg, _, _, _ = salvar_tratativas_lote_supabase(df_up, usuario=usr_str)
                if ok:
                    logs_lote = []
                    for _, row_b in df_sel_cards.iterrows():
                        logs_lote.append({
                            "Data_Hora": agora_str, "Bilhete": row_b.get("Bilhetes"), "Usuario_Acao": usr_str,
                            "Status_Anterior": row_b.get("Status_Geral", "-"), "Novo_Status": n_status,
                            "Area_Anterior": row_b.get("Área Resp. Operação", "-"), "Nova_Area": n_area,
                            "Observacao": n_obs, "Tipo_Interacao": f"Tratativa Direta ({nome_base})"
                        })
                    registrar_log_supabase(logs_lote)
                    st.success(msg)
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg)
                    
    df_tbl_final = df_exib.copy()
    if bilhet_sel:
        df_tbl_final["🎯 Destaque"] = df_tbl_final["Bilhetes"].isin(bilhet_sel).map({True: "⭐ SELECIONADO", False: ""})
        cols_order = ["🎯 Destaque"] + [c for c in df_tbl_final.columns if c != "🎯 Destaque"]
        df_tbl_final = df_tbl_final[cols_order].sort_values("🎯 Destaque", ascending=False)
        
    st.markdown("---")
    st.dataframe(df_tbl_final, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# ABAS OPERACIONAIS
# ------------------------------------------------------------------------------
with aba_sel[1]:
    st.markdown("<div class='section-banner'><h3>📋 1. Falta de Lançamento no ERP</h3></div>", unsafe_allow_html=True)
    renderizar_modulo_tratativa(f_falta, "Falta_de_Lancamento", "fl")

with aba_sel[2]:
    st.markdown("<div class='section-banner'><h3>⚠️ 2. Divergências de Valores & CIA</h3></div>", unsafe_allow_html=True)
    renderizar_modulo_tratativa(f_erros, "Erros_Valores_CIA", "ev")

with aba_sel[3]:
    st.markdown("<div class='section-banner'><h3>🎧 3. Suporte Backoffice</h3></div>", unsafe_allow_html=True)
    renderizar_modulo_tratativa(f_backoffice, "Suporte_Backoffice", "sb")

with aba_sel[4]:
    st.markdown("<div class='section-banner'><h3>🎪 4. Central de Eventos</h3></div>", unsafe_allow_html=True)
    renderizar_modulo_tratativa(f_eventos, "Central_de_Eventos", "ce")

with aba_sel[5]:
    st.markdown("<div class='section-banner'><h3>🤖 5. Emissor Virtual Lemontech</h3></div>", unsafe_allow_html=True)
    renderizar_modulo_tratativa(f_lemon_virt, "Emissor_Virtual_Lemontech", "evl")

# ------------------------------------------------------------------------------
# ABA 6: BILHETES CONCILIADOS (SEM DIVERGÊNCIA) - APRIMORADA
# ------------------------------------------------------------------------------
with aba_sel[6]:
    st.markdown(
        """
        <div class='section-banner'>
            <h3>✅ 6. Módulo de Bilhetes Conciliados (Sem Divergência)</h3>
            <p>Base de bilhetes validados, sem divergências ativas e integrados no ERP. Utilize os filtros avançados abaixo para análises manuais e extração de relatórios.</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Filtros para Análise Manual
    f_c1, f_c2, f_c3 = st.columns(3)
    with f_c1:
        opts_ger = sorted(f_sem_div["Área Resp. Operação"].dropna().unique().tolist()) if "Área Resp. Operação" in f_sem_div.columns else []
        filtro_ok_ger = st.multiselect("Filtrar por Gerente:", options=opts_ger, key="f_ok_ger")
    with f_c2:
        opts_cia = sorted(f_sem_div["CIA"].dropna().unique().tolist()) if "CIA" in f_sem_div.columns else []
        filtro_ok_cia = st.multiselect("Filtrar por CIA Aérea:", options=opts_cia, key="f_ok_cia")
    with f_c3:
        opts_setor = sorted(f_sem_div["Setor"].dropna().unique().tolist()) if "Setor" in f_sem_div.columns else []
        filtro_ok_setor = st.multiselect("Filtrar por Setor:", options=opts_setor, key="f_ok_setor")
        
    df_ok_view = f_sem_div.copy()
    if filtro_ok_ger:
        df_ok_view = df_ok_view[df_ok_view["Área Resp. Operação"].isin(filtro_ok_ger)]
    if filtro_ok_cia:
        df_ok_view = df_ok_view[df_ok_view["CIA"].isin(filtro_ok_cia)]
    if filtro_ok_setor:
        df_ok_view = df_ok_view[df_ok_view["Setor"].isin(filtro_ok_setor)]
        
    col_ok1, col_ok2 = st.columns([3, 1])
    with col_ok1:
        s_ok = st.text_input("🔍 Buscar em Bilhetes Conciliados (por Bilhete, LOC ou Passageiro):", key="s_ok")
    with col_ok2:
        st.write("")
        st.write("")
        st.download_button(
            "📥 Extrair Relatório Executivo (Sem Divergência)",
            data=gerar_excel_estilizado(df_ok_view, "Bilhetes_Conciliados"),
            file_name="Bilhetes_Conciliados_Filtrado.xlsx",
            key="btn_dl_ok"
        )
        
    if s_ok.strip():
        term = s_ok.strip().lower()
        cols_ok_s = [c for c in ["Bilhetes", "Localizador_Sistema", "Rloc_Cia", "Ponto de venda", "Gerentes", "Consultor"] if c in df_ok_view.columns]
        m_s_ok = pd.Series(False, index=df_ok_view.index)
        for col in cols_ok_s:
            m_s_ok |= df_ok_view[col].astype(str).str.lower().str.contains(term, na=False)
        df_ok_view = df_ok_view[m_s_ok]

    st.markdown(f"**Bilhetes Conciliados Exibidos:** {len(df_ok_view):,}")
    st.dataframe(df_ok_view, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# ABAS EXCLUSIVAS DO MASTER
# ------------------------------------------------------------------------------
if e_master():
    with aba_sel[7]:
        st.subheader("⚙️ Central de Aprovação de Acessos")
        usuarios_db = carregar_usuarios_supabase()
        pendentes = {k: v for k, v in usuarios_db.items() if v.get("status") == "PENDENTE"}
        if pendentes:
            for u_k, u_v in pendentes.items():
                st.write(f"👤 **{u_v['nome']}** (`{u_k}`) | Perfil: **{u_v['perfil']}**")
                ca1, ca2, _ = st.columns([1, 1, 4])
                if ca1.button(f"✅ Aprovar {u_k}"):
                    if atualizar_status_usuario_supabase(u_k, "APROVADO"):
                        st.success(f"Usuário {u_k} aprovado!")
                        st.rerun()
                if ca2.button(f"❌ Rejeitar {u_k}"):
                    if atualizar_status_usuario_supabase(u_k, "REJEITADO"):
                        st.success(f"Usuário {u_k} rejeitado.")
                        st.rerun()
        else:
            st.info("Nenhuma solicitação de acesso pendente.")
            
    with aba_sel[8]:
        st.subheader("📜 Histórico de Modificações no Supabase")
        engine_sb = get_db_engine()
        if engine_sb:
            try:
                df_log_db = pd.read_sql("SELECT * FROM log_auditoria ORDER BY id DESC LIMIT 500", engine_sb)
                st.dataframe(df_log_db, use_container_width=True, hide_index=True)
            except Exception:
                st.info("Nenhum registro de histórico encontrado na tabela log_auditoria do Supabase.")
                
    # --------------------------------------------------------------------------
    # ABA 9: CARGA EM LOTE COM ESTATÍSTICAS AVANÇADAS
    # --------------------------------------------------------------------------
    with aba_sel[9]:
        st.subheader("📥 Carga de Relatórios de Retorno (Processamento em Lote Avançado)")
        st.markdown("Envie uma planilha `.xlsx` ou `.csv` para atualização massiva no Supabase com análise estatística prévia.")
        st.warning("🔒 **Sincronização em Lote Ativa:** Os bilhetes e seus campos serão atualizados diretamente no Supabase mantendo a última versão.")
        
        arq_upload = st.file_uploader("Selecione o arquivo de retorno:", type=["xlsx", "xls", "csv"], key="uploader_lote")
        if arq_upload:
            try:
                if arq_upload.name.endswith(".csv"):
                    df_up_raw = pd.read_csv(arq_upload, dtype=str)
                else:
                    df_up_raw = pd.read_excel(arq_upload, dtype=str)
                cols_up = df_up_raw.columns.tolist()
                
                col_b = next((c for c in cols_up if c in ["Bilhetes", "Bilhete"] or any(x in str(c).lower() for x in ["bilhete", "ticket"])), None)
                col_st = next((c for c in cols_up if str(c).strip().lower() in ["status_geral", "status geral", "status", "novo status"]), None)
                col_ar = next((c for c in cols_up if any(x in str(c).lower() for x in ["área resp", "area resp", "responsavel", "gerente", "área"])), None)
                col_obs = next((c for c in cols_up if any(x in str(c).lower() for x in ["obs", "observação", "observacao", "justificativa"])), None)
                col_setor = next((c for c in cols_up if str(c).lower() == "setor"), None)
                col_gerentes = next((c for c in cols_up if str(c).lower() in ["gerentes", "gerente"]), None)
                
                if not col_b:
                    st.error("⚠️ O arquivo precisa conter ao menos uma coluna identificadora de 'Bilhete'.")
                else:
                    df_up_raw["Bilhete_Clean"] = df_up_raw[col_b].apply(sanitizar_bilhete)
                    
                    df_existentes = carregar_tratativas_db()
                    set_existentes = set(df_existentes["bilhete"].apply(sanitizar_bilhete).tolist()) if not df_existentes.empty else set()

                    lote_alteracoes = []
                    novos_logs = []
                    agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                    
                    tot_lote = 0
                    tot_para_atualizar = 0
                    tot_novos = 0

                    for _, r_v in df_up_raw.iterrows():
                        b_code = r_v["Bilhete_Clean"]
                        if not b_code or b_code.lower() in ["nan", "none", "-", ""]:
                            continue
                            
                        tot_lote += 1
                        if b_code in set_existentes:
                            tot_para_atualizar += 1
                            tipo_registro = "Atualização de Existente"
                        else:
                            tot_novos += 1
                            tipo_registro = "Novo Registro"

                        st_novo = str(r_v.get(col_st, "Pendente de Lançamento (Não Consta)")).strip() if col_st else "Pendente de Lançamento (Não Consta)"
                        ar_novo = str(r_v.get(col_ar, "Operação")).strip() if col_ar else "Operação"
                        obs_novo = str(r_v.get(col_obs, "")).replace("Sem tratativa na operação", "").strip() if col_obs else ""
                        setor_novo = str(r_v.get(col_setor, "-")).strip() if col_setor else "-"
                        gerentes_novo = str(r_v.get(col_gerentes, ar_novo)).strip() if col_gerentes else ar_novo
                        
                        lote_alteracoes.append({
                            "Bilhetes": b_code,
                            "Status_Geral": st_novo,
                            "Área Resp. Operação": ar_novo,
                            "Obs. Operação": obs_novo,
                            "Setor": setor_novo,
                            "Gerentes": gerentes_novo,
                            "Ação no Supabase": tipo_registro
                        })
                        novos_logs.append({
                            "Data_Hora": agora_str,
                            "Bilhete": b_code,
                            "Usuario_Acao": usr_str,
                            "Status_Anterior": "-",
                            "Novo_Status": st_novo,
                            "Area_Anterior": "-",
                            "Nova_Area": ar_novo,
                            "Observacao": obs_novo,
                            "Tipo_Interacao": f"Carga em Lote ({tipo_registro})"
                        })
                            
                    df_lote_prep = pd.DataFrame(lote_alteracoes)
                    
                    st.markdown("### 📊 Relatório Estatístico Pré-Carga")
                    
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Total de Bilhetes", tot_lote)
                    m2.metric("Bilhetes a Atualizar no Supabase", tot_para_atualizar)
                    m3.metric("Novos Bilhetes a Inserir", tot_novos)
                    
                    duplicados_arquivo = df_lote_prep.duplicated(subset=["Bilhetes"]).sum() if not df_lote_prep.empty else 0
                    m4.metric("Duplicados no Arquivo", duplicados_arquivo)
                    
                    if not df_lote_prep.empty:
                        st.markdown("#### 📈 Estatísticas do Arquivo Importado")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            st.markdown("**Distribuição por Status de Destino:**")
                            df_st_lote = df_lote_prep["Status_Geral"].value_counts().reset_index()
                            df_st_lote.columns = ["Status", "Quantidade"]
                            st.dataframe(df_st_lote, use_container_width=True, hide_index=True)
                        with col_e2:
                            st.markdown("**Distribuição por Área Responsável / Gerente:**")
                            df_ar_lote = df_lote_prep["Área Resp. Operação"].value_counts().reset_index()
                            df_ar_lote.columns = ["Gerente", "Quantidade"]
                            st.dataframe(df_ar_lote, use_container_width=True, hide_index=True)

                        st.markdown("**Amostra dos Dados a Serem Sincronizados no Supabase:**")
                        st.dataframe(df_lote_prep, use_container_width=True)
                        
                        if st.button("🚀 Confirmar e Enviar Atualizações para o Supabase", key="btn_confirmar_lote"):
                            bar_prog = st.progress(0, text="Sincronizando com o banco de dados Supabase...")
                            
                            ok, msg, n_at, n_nv, n_err = salvar_tratativas_lote_supabase(df_lote_prep, usuario=f"Carga_Lote_{usr_str}")
                            
                            bar_prog.progress(50, text="Gravando trilha de histórico...")
                            registrar_log_supabase(novos_logs)
                            
                            bar_prog.progress(100, text="Concluído!")
                            st.success(msg)
                            
                            time.sleep(1.5)
                            st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")