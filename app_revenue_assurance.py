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
# CONEXÃO COM SUPABASE (POSTGRESQL NUVEM)
# ==============================================================================
def get_db_engine():
    """Conecta ao Supabase buscando a URI configurada nos Secrets do Streamlit."""
    try:
        if "postgres" in st.secrets and "url" in st.secrets["postgres"]:
            db_url = st.secrets["postgres"]["url"]
            return create_engine(db_url, pool_pre_ping=True)
    except Exception:
        pass
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

def salvar_tratativas_lote_supabase(df_lote, usuario):
    """Executa o UPSERT de um lote de tratativas no Supabase."""
    engine = get_db_engine()
    if not engine or df_lote.empty:
        return False, "Conexão com o banco de dados indisponível."

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

    dados_lote = []
    for _, r in df_lote.iterrows():
        b_clean = re.sub(r"\.0$", "", str(r.get("Bilhetes", "")).strip())
        if b_clean and b_clean.lower() not in ["nan", "none", ""]:
            obs_limpa = str(r.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip()
            dados_lote.append({
                "bilhete": b_clean,
                "status_geral": str(r.get("Status_Geral", "")),
                "area_resp": str(r.get("Área Resp. Operação", "")),
                "obs_operacao": obs_limpa,
                "usuario": str(usuario)
            })

    if not dados_lote:
        return False, "Nenhum bilhete válido para atualização."

    try:
        with engine.begin() as conn:
            conn.execute(upsert_sql, dados_lote)
        st.cache_data.clear()
        return True, f"✅ {len(dados_lote)} bilhete(s) atualizados com sucesso no Supabase!"
    except Exception as e:
        return False, f"⚠️ Erro ao salvar no banco: {e}"

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
    except Exception:
        pass

# ==============================================================================
# 2. GESTÃO DE USUÁRIOS NO SUPABASE
# ==============================================================================
def carregar_usuarios_supabase():
    """Busca os usuários cadastrados na tabela usuarios do Supabase."""
    engine = get_db_engine()
    if not engine:
        return {}
    
    query = "SELECT usuario, senha, nome, perfil, status, data_solicitacao FROM usuarios"
    try:
        df_u = pd.read_sql(query, engine)
        dict_u = {}
        for _, row in df_u.iterrows():
            dict_u[str(row["usuario"]).strip().lower()] = {
                "senha": str(row["senha"]).strip(),
                "nome": str(row["nome"]).strip(),
                "perfil": str(row["perfil"]).strip(),
                "status": str(row["status"]).strip().upper(),
                "data_solicitacao": str(row["data_solicitacao"])
            }
        return dict_u
    except Exception:
        return {}

def salvar_usuario_supabase(u_id, senha, nome, perfil, status="PENDENTE"):
    """Insere ou atualiza um usuário no Supabase."""
    engine = get_db_engine()
    if not engine:
        return False, "Conexão com o banco indisponível."
    
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
    """Atualiza a senha de um usuário no Supabase."""
    engine = get_db_engine()
    if not engine:
        return False, "Conexão com o banco indisponível."
    
    sql = text("UPDATE usuarios SET senha = :p WHERE usuario = :u;")
    try:
        with engine.begin() as conn:
            conn.execute(sql, {"u": u_id, "p": nova_senha})
        return True, "Senha redefinida com sucesso."
    except Exception as e:
        return False, f"Erro no banco: {e}"

def atualizar_status_usuario_supabase(u_id, novo_status):
    """Atualiza o status (APROVADO/REJEITADO) de um usuário no Supabase."""
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

# Inicialização de Sessão
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
    return u_id in ["mribeiro", "fellipe"] or perfil in ["Master", "Compliance"]

if not st.session_state["autenticado"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        st.markdown("<div class='main-header'><h1>✈️ Grupo Arbaitman | Login</h1></div>", unsafe_allow_html=True)
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
                        st.success("✅ Senha redefinida no Supabase! Você já pode entrar com a nova senha.")
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
                        st.success("✅ Solicitação enviada! Aguarde a liberação do usuário Master.")
                    else:
                        st.error(f"Erro ao salvar solicitação: {msg}")
    st.stop()

# ==============================================================================
# 3. TRATAMENTO DE DADOS (Correção Segura Sem Erro de Float/Attribute)
# ==============================================================================
def clean_str(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    return re.sub(r"\.0$", "", s)

def padronizar_df(df):
    """Garante a integridade total do DataFrame e padroniza os status de forma 100% segura."""
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Bilhetes", "Ponto de venda", "Status_Geral", "Área Resp. Operação",
            "Obs. Operação", "Setor", "Data Emissão", "CIA", "Taxa", "A vista", "A credito",
            "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao", "Emissor", "Emissor_Reserva_Lemon",
            "Status_Sistema", "Status_Divergencia", "Gerentes"
        ])
    
    df_out = df.copy().loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    
    if "Bilhetes" not in df_out.columns:
        cols_b = [c for c in df_out.columns if "bilhete" in str(c).lower() or "ticket" in str(c).lower()]
        df_out = df_out.rename(columns={cols_b[0]: "Bilhetes"}) if cols_b else df_out.assign(Bilhetes="")
    df_out["Bilhetes"] = df_out["Bilhetes"].apply(clean_str)
    
    if "Obs. Operação" in df_out.columns:
        df_out["Obs. Operação"] = df_out["Obs. Operação"].fillna("").astype(str).str.replace("Sem tratativa na operação", "", regex=False).str.strip()

    for c in ["Ponto de venda", "Área Resp. Operação", "Obs. Operação", "Setor", "Status_Geral", "CIA", "Data Emissão", "Emissor", "Emissor_Reserva_Lemon", "Status_Sistema", "Status_Divergencia", "Gerentes"]:
        if c not in df_out.columns:
            df_out[c] = "-"
        else:
            df_out[c] = df_out[c].fillna("-").astype(str).str.strip()

    status_map = ["nao_consta", "não_consta", "nan", "-", "", "none"]
    df_out["Status_Geral"] = df_out["Status_Geral"].apply(
        lambda x: "Pendente de Lançamento (Não Consta)" if str(x).lower().strip() in status_map else str(x)
    )

    for c in ["Taxa", "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa_Sistema", "Dif_Taxa", "Receita_Sistema", "Dif_Receita", "Tarifa_Total"]:
        if c in df_out.columns: df_out[c] = pd.to_numeric(df_out[c], errors="coerce").fillna(0.0)

    for c in ["Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao"]:
        if c not in df_out.columns: df_out[c] = "-"

    df_out["Dt_Parsed"] = pd.to_datetime(df_out["Data Emissão"], format="mixed", dayfirst=True, errors="coerce")
    return df_out

def mesclar_com_supabase(df_in):
    """Sobrescreve o status e observações com os dados salvos no Supabase mantendo todas as colunas."""
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

# ==============================================================================
# ROTEAMENTO ESTRITO COM AJUSTE DE EXCLUSIVIDADE DE RECEITA
# ==============================================================================
def rotear_bases_mestra(df_master):
    if df_master.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_master = padronizar_df(df_master)

    def e_apenas_divergencia_receita(st_div):
        s = str(st_div).lower().strip()
        if "diverg" not in s and "erro" not in s:
            return False
        
        s_sem_diverg = s.replace("divergência", "").replace("divergencia", "").strip()
        tem_receita = "receita" in s_sem_diverg
        outros_erros = any(x in s_sem_diverg for x in ["tarifa", "taxa", "cia", "companhia"])
        return tem_receita and not outros_erros

    # 1. Sem Divergência / Conciliados (Tela 6)
    mask_valores_corretos = df_master["Status_Geral"].astype(str).str.lower().str.contains("já lançado|conciliado|valores corretos|regularizado") | \
                           df_master["Status_Divergencia"].astype(str).str.lower().str.contains("valores corretos")
    
    mask_so_receita = df_master["Status_Divergencia"].apply(e_apenas_divergencia_receita)

    mask_ok = mask_valores_corretos | mask_so_receita

    df_ok = df_master[mask_ok].copy()
    df_rest = df_master[~mask_ok].copy()

    # 2. Suporte Backoffice (Tela 3 - Regra Flexível de Alimentação)
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

    # 3. Central de Eventos (Tela 4)
    mask_evt = df_rest["Setor"].astype(str).str.lower().str.contains("eventos") | \
               df_rest["Área Resp. Operação"].astype(str).str.lower().str.contains("eventos") | \
               df_rest["Gerentes"].astype(str).str.lower().str.contains("eventos")
    df_eventos = df_rest[mask_evt].copy()
    df_rest = df_rest[~mask_evt].copy()

    # 4. Emissor Virtual Lemontech (Tela 5)
    mask_lemon = df_rest["Emissor"].astype(str).str.lower().str.contains("virtual") | \
                 df_rest["Emissor_Reserva_Lemon"].astype(str).str.lower().str.contains("virtual")
    df_lemon_virt = df_rest[mask_lemon].copy()
    df_rest = df_rest[~mask_lemon].copy()

    # 5. Erros de Valores & CIA (Tela 2)
    mask_consta_benner = ~df_rest["Status_Sistema"].astype(str).str.upper().str.contains("NAO_CONSTA|NÃO_CONSTA")
    mask_tem_divergencia = df_rest["Status_Divergencia"].astype(str).str.lower().str.contains("divergência|divergencia|erro")
    
    mask_erros = mask_consta_benner | mask_tem_divergencia
    df_erros = df_rest[mask_erros].copy()
    
    # 6. Falta de Lançamento (Tela 1)
    df_falta = df_rest[~mask_erros].copy()

    return df_falta, df_erros, df_bo, df_eventos, df_lemon_virt, df_ok

@st.cache_data(ttl=60)
def carregar_bases():
    if not os.path.exists(ARQUIVO_DASHBOARD):
        vazio = padronizar_df(None)
        return vazio, vazio, vazio, vazio, vazio, vazio, pd.DataFrame()

    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD, engine="openpyxl")
        frames = []
        for sheet in xls.sheet_names:
            if sheet.startswith("98_") or sheet.startswith("99_"):
                df_s = pd.read_excel(xls, sheet_name=sheet)
                df_s["Aba_Origem"] = sheet
                frames.append(df_s)

        df_log_arq = pd.read_excel(xls, "00_Log_Auditoria") if "00_Log_Auditoria" in xls.sheet_names else pd.DataFrame()

        if not frames:
            vazio = padronizar_df(None)
            return vazio, vazio, vazio, vazio, vazio, vazio, df_log_arq

        df_m = pd.concat(frames, ignore_index=True)
        
        if "Bilhetes" in df_m.columns:
            df_m = df_m.drop_duplicates(subset=["Bilhetes"], keep="first").reset_index(drop=True)

        df_m = mesclar_com_supabase(df_m)

        f_falta, f_erros, f_bo, f_evt, f_lem, f_ok = rotear_bases_mestra(df_m)
        return f_falta, f_erros, f_bo, f_evt, f_lem, f_ok, df_log_arq

    except Exception as e:
        st.error(f"⚠️ Erro ao carregar as bases de dados: {e}")
        vazio = padronizar_df(None)
        return vazio, vazio, vazio, vazio, vazio, vazio, pd.DataFrame()

# Gerador de Relatório Estilizado em Excel
def gerar_excel_estilizado(df_export, nome_aba="Relatorio"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(columns=["Aviso"], data=[["Nenhum registro encontrado para os filtros selecionados"]])
    
    cols_remover = [c for c in ["Dt_Parsed", "Aba_Origem"] if c in df_export.columns]
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

with st.spinner("🔄 Carregando bases de dados do painel e conectando ao Supabase..."):
    df_falta, df_erros, df_backoffice, df_eventos, df_lemon_virt, df_sem_div, df_log = carregar_bases()

# ==============================================================================
# 4. SIDEBAR E FILTROS OPERACIONAIS
# ==============================================================================
st.sidebar.title("Navegação")
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
                    st.success("✅ Senha alterada com sucesso no Supabase!")
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

f_falta = aplicar_filtros(df_falta)
f_erros = aplicar_filtros(df_erros)
f_backoffice = aplicar_filtros(df_backoffice)
f_eventos = aplicar_filtros(df_eventos)
f_lemon_virt = aplicar_filtros(df_lemon_virt)
f_sem_div = aplicar_filtros(df_sem_div)

# ==============================================================================
# 5. HEADER PRINCIPAL E CARDS DE KPIS
# ==============================================================================
st.markdown("<div class='main-header'><h1>✈️ Garantia de Receita do Portal | Auditoria de Bilhetes</h1></div>", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Pendentes de ERP", f"{len(f_falta):,}")
c2.metric("Erros Valores/CIA", f"{len(f_erros):,}")
c3.metric("Backoffice", f"{len(f_backoffice):,}")
c4.metric("Central de Eventos", f"{len(f_eventos):,}")
c5.metric("Sem Divergência (OK)", f"{len(f_sem_div):,}")

st.markdown("---")

# ==============================================================================
# 6. ESTRUTURA DE ABAS
# ==============================================================================
abas = [
    "📊 Dashboard Executivo",
    "📋 1. Falta de Lançamento",
    "⚠️ 2. Erros de Valores & CIA",
    "🎧 3. Suporte Backoffice",
    "🎪 4. Central de Eventos",
    "🤖 5. Emissor Virtual Lemontech",
    "✅ 6. Sem Divergência (OK)"
]

if e_master():
    abas.extend(["⚙️ Gestão de Acessos", "📜 Log de Auditoria", "📥 Carga de Relatórios (Lotes)"])

aba_sel = st.tabs(abas)

# ------------------------------------------------------------------------------
# ABA 0: DASHBOARD EXECUTIVO C-LEVEL
# ------------------------------------------------------------------------------
with aba_sel[0]:
    st.subheader("📊 Painel Executivo de Revenue Assurance")
    
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        st.markdown("##### ⚠️ Distribuição de Pendências por Status")
        df_todas_pend = pd.concat([f_falta, f_erros, f_backoffice, f_eventos, f_lemon_virt], ignore_index=True)
        if not df_todas_pend.empty and "Status_Geral" in df_todas_pend.columns:
            df_st_chart = df_todas_pend["Status_Geral"].value_counts().reset_index()
            df_st_chart.columns = ["Status", "Quantidade"]
            fig_pie = px.pie(df_st_chart, values="Quantidade", names="Status", hole=0.4, color_discrete_sequence=px.colors.qualitative.Bold)
            fig_pie.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=320)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Nenhuma pendência registrada.")

    with col_d2:
        st.markdown("##### 👤 Top 8 Gerentes por Volume de Pendências")
        if not df_todas_pend.empty and "Área Resp. Operação" in df_todas_pend.columns:
            df_ger_chart = df_todas_pend["Área Resp. Operação"].value_counts().head(8).reset_index()
            df_ger_chart.columns = ["Gerente", "Quantidade"]
            fig_ger = px.bar(df_ger_chart, x="Gerente", y="Quantidade", text="Quantidade", color_discrete_sequence=["#002060"])
            fig_ger.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=320)
            st.plotly_chart(fig_ger, use_container_width=True)

    col_d3, col_d4 = st.columns(2)
    with col_d3:
        st.markdown("##### 🏢 Volume de Pendências por Unidade / Setor")
        if not df_todas_pend.empty and "Setor" in df_todas_pend.columns:
            df_set = df_todas_pend["Setor"].value_counts().reset_index()
            df_set.columns = ["Setor", "Quantidade"]
            fig_set = px.bar(df_set, x="Setor", y="Quantidade", color="Setor", text_auto=True)
            fig_set.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=300, showlegend=False)
            st.plotly_chart(fig_set, use_container_width=True)

    with col_d4:
        st.markdown("##### ✈️ Pendências por Companhia Aérea (Top 8)")
        if not df_todas_pend.empty and "CIA" in df_todas_pend.columns:
            df_cia = df_todas_pend["CIA"].value_counts().head(8).reset_index()
            df_cia.columns = ["CIA", "Quantidade"]
            fig_cia = px.bar(df_cia, x="CIA", y="Quantidade", color_discrete_sequence=["#00509d"], text_auto=True)
            fig_cia.update_layout(margin=dict(l=10, r=10, t=20, b=20), height=300)
            st.plotly_chart(fig_cia, use_container_width=True)

# ------------------------------------------------------------------------------
# FUNÇÃO REUTILIZÁVEL DE TRATATIVA COM DESTAQUE E BUSCA
# ------------------------------------------------------------------------------
def renderizar_modulo_tratativa(df_filtrado, nome_base, key_prefix):
    if df_filtrado.empty:
        st.info("Nenhum bilhete pendente nesta categoria.")
        return

    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        termo_busca = st.text_input(f"🔍 Buscar nesta tela (por Bilhete, LOC, Cliente, Passageiro ou Gerente):", key=f"src_{key_prefix}")
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
        cols_search = [c for c in ["Bilhetes", "Localizador_Sistema", "Rloc_Cia", "Ponto de venda", "Gerentes", "Área Resp. Operação", "Consultor", "Cliente"] if c in df_exib.columns]
        mask_src = pd.Series(False, index=df_exib.index)
        for col in cols_search:
            mask_src |= df_exib[col].astype(str).str.lower().str.contains(term, na=False)
        df_exib = df_exib[mask_src]

    st.markdown(f"**Registros Visíveis:** {len(df_exib)}")

    bilhetes_lista = df_exib["Bilhetes"].tolist() if "Bilhetes" in df_exib.columns else []
    bilhet_sel = st.multiselect("Selecione um ou mais Bilhetes para Tratativa:", options=bilhetes_lista, key=f"ms_{key_prefix}")

    if bilhet_sel:
        df_sel_cards = df_exib[df_exib["Bilhetes"].isin(bilhet_sel)]
        st.markdown('<div class="highlight-card">', unsafe_allow_html=True)
        st.markdown(f"#### 🎯 Informações do(s) Bilhete(s) Selecionado(s) ({len(df_sel_cards)})")
        for _, r_card in df_sel_cards.iterrows():
            obs_card = str(r_card.get('Obs. Operação', '')).replace("Sem tratativa na operação", "").strip() or '-'
            st.markdown(f"""
            * **Bilhete:** `{r_card.get('Bilhetes')}` | **CIA:** {r_card.get('CIA')} | **Ponto Venda:** {r_card.get('Ponto de venda')}
            * **Gerente Atual:** {r_card.get('Área Resp. Operação')} | **Status Atual:** `{r_card.get('Status_Geral')}`
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
                idx_default = opcoes_factiveis.index(status_atual_val) if status_atual_val in opcoes_factiveis else 0
                n_status = st.selectbox("Novo Status (Obrigatório):", options=opcoes_factiveis, index=idx_default, key=f"st_{key_prefix}")
            
            with c_s2:
                lista_gerentes = sorted(list(set(df_todos["Área Resp. Operação"].dropna().unique()))) if "Área Resp. Operação" in df_todos.columns else ["Operação"]
                idx_ger = lista_gerentes.index(df_primeiro.get("Área Resp. Operação")) if df_primeiro.get("Área Resp. Operação") in lista_gerentes else 0
                n_area = st.selectbox("Nova Área Responsável / Gerente:", options=lista_gerentes, index=idx_ger, key=f"ar_{key_prefix}")
            
            obs_init = str(df_primeiro.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip()
            n_obs = st.text_area("Observação / Justificativa Detalhada:", value=obs_init, key=f"obs_{key_prefix}")
            
            if st.form_submit_button("💾 Salvar e Atualizar Bilhetes"):
                usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                df_up = pd.DataFrame({
                    "Bilhetes": bilhet_sel,
                    "Status_Geral": [n_status] * len(bilhet_sel),
                    "Área Resp. Operação": [n_area] * len(bilhet_sel),
                    "Obs. Operação": [n_obs] * len(bilhet_sel)
                })

                ok, msg = salvar_tratativas_lote_supabase(df_up, usuario=usr_str)
                if ok:
                    logs_lote = []
                    for b_s in bilhet_sel:
                        logs_lote.append({
                            "Data_Hora": agora_str, "Bilhete": b_s, "Usuario_Acao": usr_str,
                            "Status_Anterior": df_primeiro.get("Status_Geral", "-"), "Novo_Status": n_status,
                            "Area_Anterior": df_primeiro.get("Área Resp. Operação", "-"), "Nova_Area": n_area,
                            "Observacao": n_obs, "Tipo_Interacao": f"Tratativa Web ({nome_base})"
                        })
                    registrar_log_supabase(logs_lote)
                    st.success(msg)
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
    st.subheader("📋 1. Pendências de Lançamento no ERP")
    renderizar_modulo_tratativa(f_falta, "Falta_de_Lancamento", "fl")

with aba_sel[2]:
    st.subheader("⚠️ 2. Divergências de Valores e Companhias Aéreas")
    renderizar_modulo_tratativa(f_erros, "Erros_Valores_CIA", "ev")

with aba_sel[3]:
    st.subheader("🎧 3. Chamados em Análise no Suporte Backoffice")
    renderizar_modulo_tratativa(f_backoffice, "Suporte_Backoffice", "sb")

with aba_sel[4]:
    st.subheader("🎪 4. Central de Eventos")
    renderizar_modulo_tratativa(f_eventos, "Central_de_Eventos", "ce")

with aba_sel[5]:
    st.subheader("🤖 5. Emissor Virtual Lemontech")
    renderizar_modulo_tratativa(f_lemon_virt, "Emissor_Virtual_Lemontech", "evl")

with aba_sel[6]:
    st.subheader("✅ 6. Bilhetes Conciliados e Lançados")
    col_ok1, col_ok2 = st.columns([3, 1])
    with col_ok1:
        s_ok = st.text_input("🔍 Buscar em Bilhetes Conciliados:", key="s_ok")
    with col_ok2:
        st.write("")
        st.write("")
        st.download_button(
            "📥 Extrair Relatório (Sem Divergência)",
            data=gerar_excel_estilizado(f_sem_div, "Sem_Divergencia"),
            file_name="Sem_Divergencia_Filtrado.xlsx"
        )
    df_ok_view = f_sem_div.copy()
    if s_ok.strip():
        term = s_ok.strip().lower()
        df_ok_view = df_ok_view[df_ok_view["Bilhetes"].astype(str).str.lower().str.contains(term, na=False)]
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
                    else:
                        st.error("Erro ao aprovar no banco de dados.")
                if ca2.button(f"❌ Rejeitar {u_k}"):
                    if atualizar_status_usuario_supabase(u_k, "REJEITADO"):
                        st.success(f"Usuário {u_k} rejeitado.")
                        st.rerun()
                    else:
                        st.error("Erro ao rejeitar no banco de dados.")
        else:
            st.info("Nenhuma solicitação de acesso pendente.")

    with aba_sel[8]:
        st.subheader("📜 Trilha de Auditoria do Supabase")
        engine_sb = get_db_engine()
        if engine_sb:
            try:
                df_log_db = pd.read_sql("SELECT * FROM log_auditoria ORDER BY id DESC LIMIT 500", engine_sb)
                st.dataframe(df_log_db, use_container_width=True, hide_index=True)
            except Exception:
                st.info("Nenhum registro de log encontrado na tabela log_auditoria do Supabase.")

    with aba_sel[9]:
        st.subheader("📥 Carga de Relatórios de Retorno (Processamento em Lote Protegido)")
        st.markdown("Envie uma planilha `.xlsx` ou `.csv` contendo as colunas de **Bilhete**, **Novo Status**, **Área Responsável** e **Observação** para atualização em massa no Supabase.")
        st.warning("🔒 **Proteção de Escopo Ativa:** Somente bilhetes pertencentes à base auditada original serão atualizados.")

        arq_upload = st.file_uploader("Selecione o arquivo de retorno:", type=["xlsx", "xls", "csv"])
        if arq_upload:
            try:
                if arq_upload.name.endswith(".csv"):
                    df_up_raw = pd.read_csv(arq_upload, dtype=str)
                else:
                    df_up_raw = pd.read_excel(arq_upload, dtype=str)

                col_b = next((c for c in df_up_raw.columns if any(x in str(c).lower() for x in ["bilhete", "ticket"])), None)
                col_st = next((c for c in df_up_raw.columns if any(x in str(c).lower() for x in ["status", "novo status"])), None)
                col_ar = next((c for c in df_up_raw.columns if any(x in str(c).lower() for x in ["gerente", "área", "area", "responsavel"])), None)
                col_obs = next((c for c in df_up_raw.columns if any(x in str(c).lower() for x in ["obs", "observação"])), None)

                if not col_b or not col_st:
                    st.error("⚠️ O arquivo precisa conter ao menos uma coluna de 'Bilhete' e uma de 'Status'.")
                else:
                    df_up_raw["Bilhete_Clean"] = df_up_raw[col_b].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                    df_todos["Bilhete_Clean"] = df_todos["Bilhetes"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

                    df_validos = pd.merge(
                        df_up_raw,
                        df_todos[["Bilhete_Clean", "Status_Geral", "Área Resp. Operação", "Obs. Operação"]].drop_duplicates("Bilhete_Clean"),
                        on="Bilhete_Clean",
                        how="inner",
                        suffixes=("_Novo", "_Atual")
                    )

                    qtd_total_arq = len(df_up_raw)
                    qtd_validos = len(df_validos)
                    qtd_fora = qtd_total_arq - qtd_validos

                    novos_logs = []
                    lote_alteracoes = []
                    qtd_iguais = 0

                    agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"

                    for _, r_v in df_validos.iterrows():
                        b_code = r_v["Bilhete_Clean"]
                        st_novo = str(r_v.get(col_st, "")).strip()
                        ar_novo = str(r_v.get(col_ar, r_v.get("Área Resp. Operação", ""))).strip()
                        obs_novo = str(r_v.get(col_obs, r_v.get("Obs. Operação", ""))).strip()

                        st_atual = str(r_v.get("Status_Geral", "")).strip()
                        ar_atual = str(r_v.get("Área Resp. Operação", "")).strip()
                        obs_atual = str(r_v.get("Obs. Operação", "")).strip()

                        if st_novo == st_atual and ar_novo == ar_atual and obs_novo == obs_atual:
                            qtd_iguais += 1
                        else:
                            lote_alteracoes.append({
                                "Bilhetes": b_code,
                                "Status_Geral": st_novo,
                                "Área Resp. Operação": ar_novo,
                                "Obs. Operação": obs_novo
                            })
                            novos_logs.append({
                                "Data_Hora": agora_str,
                                "Bilhete": b_code,
                                "Usuario_Acao": usr_str,
                                "Status_Anterior": st_atual,
                                "Novo_Status": st_novo,
                                "Area_Anterior": ar_atual,
                                "Nova_Area": ar_novo,
                                "Observacao": obs_novo,
                                "Tipo_Interacao": "Carga em Lote (Excel)"
                            })

                    st.markdown("### 📊 Relatório de Pré-Validação da Carga em Lote")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Linhas na Planilha", qtd_total_arq)
                    m2.metric("🟢 Com Alteração", len(lote_alteracoes))
                    m3.metric("🟡 Sem Alteração (Mantidos)", qtd_iguais)
                    m4.metric("🔴 Ignorados (Fora da Base)", qtd_fora)

                    if len(lote_alteracoes) > 0:
                        st.markdown("**Amostra de Bilhetes que Serão Atualizados:**")
                        st.dataframe(pd.DataFrame(lote_alteracoes).head(10), use_container_width=True)

                        if st.button("🚀 Confirmar e Processar Atualizações no Supabase"):
                            bar_prog = st.progress(0, text="Sincronizando com o banco de dados...")
                            
                            df_lote_final = pd.DataFrame(lote_alteracoes)
                            ok, msg = salvar_tratativas_lote_supabase(df_lote_final, usuario=f"Carga_Lote_{usr_str}")
                            
                            bar_prog.progress(50, text="Gravando trilha de auditoria...")
                            registrar_log_supabase(novos_logs)
                            
                            bar_prog.progress(100, text="Concluído!")
                            st.success(f"✅ Processamento concluído! {len(lote_alteracoes)} bilhetes foram atualizados no Supabase e registrados na Trilha de Auditoria.")
                            st.rerun()
                    else:
                        st.info("ℹ️ Nenhuma alteração pendente. Todos os bilhetes válidos da planilha já possuem exatamente as mesmas tratativas registradas.")

            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")