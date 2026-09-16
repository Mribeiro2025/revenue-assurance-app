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
ARQUIVO_MEMORIA = "Historico_Tratativas.csv"
PASTA_ENVIO = "Envio"

USUARIOS_PADRAO = {
    "mribeiro": {
        "senha": "123",
        "nome": "Marcos Ribeiro",
        "perfil": "Master",
        "status": "APROVADO",
        "data_solicitacao": "2026-08-31",
    },
    "operacao": {
        "senha": "123",
        "nome": "Equipe Operacional",
        "perfil": "Operacao",
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
# 3. TRATAMENTO DE DADOS E CARGA (COM ENGINE EXPLÍCITO OPENPYXL)
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

@st.cache_data(ttl=300)
def carregar_bases():
    vazio = padronizar_df(None)
    
    if not os.path.exists(ARQUIVO_DASHBOARD):
        return vazio, vazio, vazio, vazio, pd.DataFrame()
    
    if os.path.getsize(ARQUIVO_DASHBOARD) == 0:
        st.warning("⚠️ O arquivo de consolidado está vazio ou sendo gerado. Rode a auditoria novamente.")
        return vazio, vazio, vazio, vazio, pd.DataFrame()

    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD, engine="openpyxl")
        m = pd.read_excel(xls, "99_Base_Divergencias_Geral") if "99_Base_Divergencias_Geral" in xls.sheet_names else None
        d = pd.read_excel(xls, "98_OK_Divergencia_Operacao") if "98_OK_Divergencia_Operacao" in xls.sheet_names else None
        s = pd.read_excel(xls, "98_OK_Sem_Divergencia_Concil") if "98_OK_Sem_Divergencia_Concil" in xls.sheet_names else None
        b = pd.read_excel(xls, "99_Suporte backoffice") if "99_Suporte backoffice" in xls.sheet_names else None
        l = pd.read_excel(xls, "00_Log_Auditoria") if "00_Log_Auditoria" in xls.sheet_names else pd.DataFrame()
        return padronizar_df(m), padronizar_df(d), padronizar_df(s), padronizar_df(b), l
    except Exception as e:
        st.error(f"⚠️ O arquivo de dados está temporariamente indisponível ou corrompido. Execute o script de auditoria para restaurá-lo. ({e})")
        return vazio, vazio, vazio, vazio, pd.DataFrame()

def sincronizar_memoria_csv(df_m, df_d, df_s, df_b):
    try:
        df_todos = pd.concat([df_m, df_d, df_s, df_b], ignore_index=True)
        cols_mem = [
            c for c in [
                "Bilhetes", "Status_Geral", "Área Resp. Operação", "Obs. Operação",
                "Setor", "Ponto de venda", "Código Iata", "Data_Modificacao",
                "Usuario_Modificacao", "Ultima_Alteracao"
            ] if c in df_todos.columns
        ]
        df_mem = df_todos[cols_mem].drop_duplicates(subset=["Bilhetes"], keep="last")
        
        tmp_csv = f"{ARQUIVO_MEMORIA}.tmp"
        df_mem.to_csv(tmp_csv, index=False, encoding="utf-8-sig")
        if os.path.exists(ARQUIVO_MEMORIA):
            os.remove(ARQUIVO_MEMORIA)
        os.rename(tmp_csv, ARQUIVO_MEMORIA)
    except Exception as e:
        st.warning(f"Aviso de sincronização da memória CSV: {e}")

def salvar_copia_retorno_automatica(novos_logs_list, pasta_destino=PASTA_ENVIO):
    if not novos_logs_list:
        return
    try:
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)
            
        agora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = os.path.join(pasta_destino, f"Relatorio_Alteracoes_{agora}.xlsx")
        
        df_alteracoes = pd.DataFrame(novos_logs_list)
        df_alteracoes.rename(columns={
            "Bilhete": "Nº Bilhete / LOC",
            "Status_Anterior": "Status Anterior",
            "Novo_Status": "Novo Status",
            "Area_Anterior": "Área Anterior",
            "Nova_Area": "Nova Área / Gerente",
            "Observacao": "Observação Aplicada",
            "Tipo_Interacao": "Ação Executada"
        }, inplace=True)
        
        with pd.ExcelWriter(nome_arquivo, engine="openpyxl") as writer:
            df_alteracoes.to_excel(writer, sheet_name="Relatorio_Alteracoes", index=False)
            
        print(f"📁 Cópia do relatório salva automaticamente em: {nome_arquivo}")
    except Exception as e:
        print(f"⚠️ Aviso ao salvar cópia automática do relatório: {e}")

def salvar_bases(df_m, df_d, df_s, df_b, df_l):
    try:
        cols_drop = ["Dt_Parsed"]
        def clean_df(df_in):
            if df_in is None: return pd.DataFrame()
            return df_in.drop(columns=cols_drop, errors="ignore")

        df_m_c, df_d_c, df_s_c, df_b_c, df_l_c = clean_df(df_m), clean_df(df_d), clean_df(df_s), clean_df(df_b), clean_df(df_l)

        with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine="openpyxl") as writer:
            df_m_c.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
            df_d_c.to_excel(writer, sheet_name="98_OK_Divergencia_Operacao", index=False)
            df_s_c.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
            df_b_c.to_excel(writer, sheet_name="99_Suporte backoffice", index=False)
            df_l_c.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)

        sincronizar_memoria_csv(df_m_c, df_d_c, df_s_c, df_b_c)

        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar arquivo consolidado: {e}")
        return False

def ingestar_lote_semanal(arquivos_novos, df_m, df_d, df_s, df_b, df_log_atual):
    """
    Ingestão com container de status em tempo real, progresso detalhado e retorno do relatório de auditoria do lote.
    """
    novos_registros_brutos = []
    novos_logs = []
    qtd_alteracoes_relatorio = 0
    total_linhas_lidas = 0
    agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    usr_str = f"{st.session_state.get('usuario_atual', 'Sistema')} ({st.session_state.get('login_user_id', 'master')})"

    total_arquivos = len(arquivos_novos)

    with st.status("📊 Processando arquivos em lote e mapeando tratativas...", expanded=True) as status_box:
        barra_progresso = st.progress(0, text="Iniciando processamento dos arquivos...")
        place_info = st.empty()

        for idx_arq, arq in enumerate(arquivos_novos):
            pct = int(((idx_arq + 1) / total_arquivos) * 100)
            barra_progresso.progress(pct, text=f"📂 Lendo arquivo {idx_arq + 1} de {total_arquivos}: '{arq.name}' ({pct}%)")

            xls = pd.ExcelFile(arq, engine="openpyxl")
            for sheet in xls.sheet_names:
                df_temp = pd.read_excel(xls, sheet_name=sheet)
                total_linhas_lidas += len(df_temp)
                cols_clean = [str(c).strip() for c in df_temp.columns]

                place_info.markdown(
                    f"🔹 **Arquivo Atual:** `{arq.name}` | **Aba:** `{sheet}` | **Linhas:** `{len(df_temp):,}`\n\n"
                    f"📈 **Total Acumulado:** `{total_linhas_lidas:,}` linhas analisadas | `{qtd_alteracoes_relatorio:,}` alterações mapeadas."
                )

                if "Nº Bilhete / LOC" in cols_clean or "Novo Status" in cols_clean:
                    col_b = next((c for c in df_temp.columns if "bilhete" in str(c).lower() or "loc" in str(c).lower()), None)
                    col_sn = next((c for c in df_temp.columns if "novo status" in str(c).lower()), None)
                    col_sa = next((c for c in df_temp.columns if "status anterior" in str(c).lower()), None)
                    col_an = next((c for c in df_temp.columns if "nova área" in str(c).lower() or "gerente" in str(c).lower()), None)
                    col_aa = next((c for c in df_temp.columns if "área anterior" in str(c).lower()), None)
                    col_obs = next((c for c in df_temp.columns if "observação" in str(c).lower() or "obs" in str(c).lower()), None)
                    col_act = next((c for c in df_temp.columns if "ação" in str(c).lower()), None)

                    all_target_dfs = [df_m, df_d, df_s, df_b]
                    for _, r in df_temp.iterrows():
                        b_val = str(r[col_b]).strip().replace(".0", "") if col_b and pd.notna(r[col_b]) else ""
                        if not b_val or b_val.lower() in ["nan", "none", ""]: 
                            continue

                        st_novo = str(r[col_sn]).strip() if col_sn and pd.notna(r[col_sn]) else ""
                        st_ant = str(r[col_sa]).strip() if col_sa and pd.notna(r[col_sa]) else "-"
                        ar_nova = str(r[col_an]).strip() if col_an and pd.notna(r[col_an]) else ""
                        ar_ant = str(r[col_aa]).strip() if col_aa and pd.notna(r[col_aa]) else "-"
                        obs_val = str(r[col_obs]).strip() if col_obs and pd.notna(r[col_obs]) else ""
                        act_val = str(r[col_act]).strip() if col_act and pd.notna(r[col_act]) else "Ingestão via Relatório de Retorno"

                        atualizado = False
                        for df_t in all_target_dfs:
                            if df_t is not None and not df_t.empty and "Bilhetes" in df_t.columns:
                                mask = df_t["Bilhetes"].astype(str).str.strip() == b_val
                                if mask.any():
                                    if st_novo and st_novo != "-": df_t.loc[mask, "Status_Geral"] = st_novo
                                    if ar_nova and ar_nova != "-": df_t.loc[mask, "Área Resp. Operação"] = ar_nova
                                    if obs_val and obs_val != "-": df_t.loc[mask, "Obs. Operação"] = obs_val
                                    if "Data_Modificacao" in df_t.columns: df_t.loc[mask, "Data_Modificacao"] = agora_str
                                    if "Usuario_Modificacao" in df_t.columns: df_t.loc[mask, "Usuario_Modificacao"] = usr_str
                                    atualizado = True

                        if atualizado or st_novo:
                            qtd_alteracoes_relatorio += 1
                            novos_logs.append({
                                "Data_Hora": agora_str,
                                "Bilhete": b_val,
                                "Usuario_Acao": usr_str,
                                "Status_Anterior": st_ant,
                                "Novo_Status": st_novo,
                                "Area_Anterior": ar_ant,
                                "Nova_Area": ar_nova,
                                "Observacao": obs_val,
                                "Tipo_Interacao": act_val
                            })
                else:
                    col_bilhete = [c for c in df_temp.columns if "bilhete" in str(c).lower() or "ticket" in str(c).lower()]
                    if col_bilhete:
                        df_temp = df_temp.rename(columns={col_bilhete[0]: "Bilhetes"})
                        df_temp["Bilhetes"] = df_temp["Bilhetes"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                        novos_registros_brutos.append(df_temp)

        status_box.update(label=f"✅ Processamento de {total_arquivos} arquivo(s) finalizado com sucesso!", state="complete", expanded=False)

    df_novos_logs = pd.DataFrame(novos_logs)
    df_log_consolidado = pd.concat([df_log_atual, df_novos_logs], ignore_index=True) if not df_novos_logs.empty else df_log_atual

    if novos_registros_brutos:
        df_lote = pd.concat(novos_registros_brutos, ignore_index=True)
        bilhetes_conciliados = set(df_s["Bilhetes"].astype(str)) if df_s is not None and "Bilhetes" in df_s.columns else set()
        df_lote_filtrado = df_lote[~df_lote["Bilhetes"].isin(bilhetes_conciliados)].copy()
        
        df_master_atualizado = pd.concat([df_m, df_lote_filtrado], ignore_index=True)
        df_master_atualizado = df_master_atualizado.drop_duplicates(subset=["Bilhetes"], keep="first")
        return padronizar_df(df_master_atualizado), df_log_consolidado, df_novos_logs, total_linhas_lidas, len(df_lote_filtrado) + qtd_alteracoes_relatorio

    return df_m, df_log_consolidado, df_novos_logs, total_linhas_lidas, qtd_alteracoes_relatorio

def gerar_excel_estilizado(df_export, nome_aba="Relatorio_Filtrado"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(columns=["Aviso"], data=[["Nenhum registro encontrado para os filtros selecionados"]])
    
    cols_remover = [c for c in ["Dt_Parsed"] if c in df_export.columns]
    df_clean = df_export.drop(columns=cols_remover) if cols_remover else df_export.copy()

    cols_prioritarias = [
        "Data_Hora", "Bilhete", "Usuario_Acao", "Status_Anterior", "Novo_Status",
        "Area_Anterior", "Nova_Area", "Observacao", "Tipo_Interacao",
        "Bilhetes", "Ponto de venda", "CIA", "Fornecedor_Sistema", "Status_Cia",
        "Status_Geral", "Status_Divergencia", "Área Resp. Operação", "Data Emissão",
        "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa", "Taxa_Sistema", "Dif_Taxa",
        "Comissão", "Taxa DU", "Desc.", "Incentivo", "Receita_Sistema", "Dif_Receita", "VL. Líquido",
        "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao", "Obs. Operação"
    ]
    existentes = [c for c in cols_prioritarias if c in df_clean.columns]
    outras = [c for c in df_clean.columns if c not in existentes]
    df_clean = df_clean[existentes + outras]

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name=nome_aba[:30], index=False)
    
    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    ws = wb.active
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    diff_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    diff_font = Font(color="B25900", bold=True, size=10)
    
    warn_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    warn_font = Font(color="C00000", bold=True, size=10)
    
    normal_font = Font(size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    currency_cols = [
        "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa", "Taxa_Sistema", "Dif_Taxa",
        "Comissão", "Taxa DU", "Desc.", "Incentivo", "Receita_Sistema", "Dif_Receita", "VL. Líquido"
    ]

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    max_row = ws.max_row
    max_col = ws.max_column
    col_names = [str(ws.cell(row=1, column=c).value or "") for c in range(1, max_col + 1)]

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        col_name = col_names[col_idx - 1]
        
        len_vals = [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, max_row + 1)]
        max_len = max(len_vals) if len_vals else 10
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        for r in range(2, max_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = thin_border
            cell.font = normal_font

            val_str = str(cell.value or "").strip()

            if col_name in currency_cols:
                try:
                    if val_str not in ["-", "", "None"]:
                        val_num = float(cell.value)
                        cell.value = val_num
                        cell.number_format = "R$ #,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        
                        if "Dif_" in col_name and abs(val_num) > 0.01:
                            cell.fill = diff_fill
                            cell.font = diff_font
                except Exception:
                    pass

            if col_name in ["Status_Divergencia", "Status_Cia"] and ("Divergência" in val_str or val_str not in ["OK", "Valores Corretos", "-", ""]):
                cell.fill = warn_fill
                cell.font = warn_font

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()

with st.spinner("🔄 Carregando bases de dados do painel..."):
    df_master, df_div_op, df_sem_div, df_backoffice, df_log = carregar_bases()

# ==============================================================================
# 4. SIDEBAR E FILTROS BRASILEIROS POR DATA
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

st.sidebar.markdown("**📅 Período da Emissão:**")
d_inicio = st.sidebar.date_input("Data Inicial:", value=datetime.date(2025, 1, 1), format="DD/MM/YYYY")
d_fim = st.sidebar.date_input("Data Final:", value=datetime.date(2026, 12, 31), format="DD/MM/YYYY")

df_todos = pd.concat([df_master, df_div_op, df_sem_div], ignore_index=True)

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
        st.markdown("##### ⚠️ Visão de Divergências Operacionais (Maior para o Menor)")
        if not f_div_op.empty and "Status_Divergencia" in f_div_op.columns:
            df_div_chart = f_div_op["Status_Divergencia"].value_counts().reset_index()
            df_div_chart.columns = ["Status_Divergencia", "count"]
            df_div_chart = df_div_chart.sort_values(by="count", ascending=True)
            
            fig_div = px.bar(
                df_div_chart,
                x="count", y="Status_Divergencia", orientation="h",
                text="count",
                color_discrete_sequence=["#002060"]
            )
            fig_div.update_traces(textposition="outside")
            fig_div.update_layout(height=380, xaxis_title="Qtd. Bilhetes", yaxis_title="", margin=dict(l=10, r=30, t=20, b=20))
            st.plotly_chart(fig_div, width="stretch")
        else:
            st.info("Nenhuma divergência registrada no período.")

    with col_d2:
        st.markdown("##### 👤 Pendências por Gerente Responsável")
        if not f_master.empty:
            df_ger_chart = f_master["Área Resp. Operação"].value_counts().head(8).reset_index()
            df_ger_chart.columns = ["Gerente", "count"]
            
            fig_ger = px.bar(
                df_ger_chart,
                x="Gerente", y="count", text="count",
                color_discrete_sequence=["#00509d"]
            )
            fig_ger.update_traces(textposition="outside")
            fig_ger.update_layout(height=380, xaxis_title="", yaxis_title="Qtd. Bilhetes", margin=dict(l=10, r=10, t=20, b=20))
            st.plotly_chart(fig_ger, width="stretch")

# ------------------------------------------------------------------------------
# FUNÇÃO PARA RENDERIZAÇÃO E TRATATIVA COM SALVAMENTO CORRETO
# ------------------------------------------------------------------------------
def renderizar_modulo_tratativa(df_filtrado, nome_base, key_prefix, df_global_ref_name):
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
                f"Status Atual: `{df_sel.get('Status_Geral', '-')}` | Última Modificação: `{df_sel.get('Data_Modificacao', '-')}`"
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
                    with st.spinner("⚡ Atualizando registros e gravando histórico..."):
                        global df_master, df_div_op, df_sem_div, df_backoffice, df_log
                        
                        agora_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        usr_str = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"
                        
                        if df_global_ref_name == "df_master":
                            target_df = df_master
                        elif df_global_ref_name == "df_div_op":
                            target_df = df_div_op
                        elif df_global_ref_name == "df_backoffice":
                            target_df = df_backoffice
                        else:
                            target_df = df_sem_div

                        mask_global = target_df["Bilhetes"].isin(bilhet_sel)
                        
                        target_df.loc[mask_global, "Status_Geral"] = n_status
                        target_df.loc[mask_global, "Área Resp. Operação"] = n_area
                        target_df.loc[mask_global, "Obs. Operação"] = n_obs
                        target_df.loc[mask_global, "Data_Modificacao"] = agora_str
                        target_df.loc[mask_global, "Usuario_Modificacao"] = usr_str
                        target_df.loc[mask_global, "Ultima_Alteracao"] = f"Status: '{n_status}' | Obs: {n_obs}"

                        novos_logs = []
                        for b_sel in bilhet_sel:
                            novos_logs.append({
                                "Data_Hora": agora_str,
                                "Bilhete": b_sel,
                                "Usuario_Acao": usr_str,
                                "Status_Anterior": df_sel.get("Status_Geral", "-"),
                                "Novo_Status": n_status,
                                "Area_Anterior": df_sel.get("Área Resp. Operação", "-"),
                                "Nova_Area": n_area,
                                "Observacao": n_obs,
                                "Tipo_Interacao": f"Tratativa ({nome_base})"
                            })
                        df_log = pd.concat([df_log, pd.DataFrame(novos_logs)], ignore_index=True)

                        if n_status in ["Já Lançado no ERP", "Sem Divergência (OK)"]:
                            movidos = target_df[mask_global].copy()
                            if df_global_ref_name == "df_master":
                                df_master = df_master[~mask_global].reset_index(drop=True)
                            elif df_global_ref_name == "df_div_op":
                                df_div_op = df_div_op[~mask_global].reset_index(drop=True)
                            elif df_global_ref_name == "df_backoffice":
                                df_backoffice = df_backoffice[~mask_global].reset_index(drop=True)
                            
                            df_sem_div = pd.concat([df_sem_div, movidos], ignore_index=True)
                        elif n_status == "Encaminhado para Suporte Backoffice" and df_global_ref_name != "df_backoffice":
                            movidos = target_df[mask_global].copy()
                            if df_global_ref_name == "df_master":
                                df_master = df_master[~mask_global].reset_index(drop=True)
                            elif df_global_ref_name == "df_div_op":
                                df_div_op = df_div_op[~mask_global].reset_index(drop=True)
                            elif df_global_ref_name == "df_sem_div":
                                df_sem_div = df_sem_div[~mask_global].reset_index(drop=True)
                            
                            df_backoffice = pd.concat([df_backoffice, movidos], ignore_index=True)
                        
                        if salvar_bases(df_master, df_div_op, df_sem_div, df_backoffice, df_log):
                            salvar_copia_retorno_automatica(novos_logs, pasta_destino=PASTA_ENVIO)
                            st.success("✅ Tratativas salvas, registradas e cópia gerada automaticamente na pasta!")
                            st.rerun()

    st.markdown("---")
    st.markdown(f"##### 📄 Lista Completa dos Bilhetes - {nome_base} ({len(df_filtrado)} registros)")
    st.dataframe(df_filtrado, width="stretch", hide_index=True)

# ------------------------------------------------------------------------------
# ABA 1: FALTA DE LANÇAMENTO
# ------------------------------------------------------------------------------
with aba_sel[1]:
    st.subheader("📋 1. Pendências de Lançamento no ERP")
    renderizar_modulo_tratativa(f_master, "Falta_de_Lancamento", "fl", "df_master")

# ------------------------------------------------------------------------------
# ABA 2: ERROS DE VALORES E CIA AÉREA
# ------------------------------------------------------------------------------
with aba_sel[2]:
    st.subheader("⚠️ 2. Divergências de Valores e Companhias Aéreas")
    renderizar_modulo_tratativa(f_div_op, "Erros_Valores_CIA", "ev", "df_div_op")

# ------------------------------------------------------------------------------
# ABA 3: SUPORTE BACKOFFICE
# ------------------------------------------------------------------------------
with aba_sel[3]:
    st.subheader("🎧 3. Chamados em Análise no Suporte Backoffice")
    renderizar_modulo_tratativa(f_backoffice, "Suporte_Backoffice", "sb", "df_backoffice")

# ------------------------------------------------------------------------------
# ABA 4: SEM DIVERGÊNCIA (OK)
# ------------------------------------------------------------------------------
with aba_sel[4]:
    st.subheader("✅ 4. Bilhetes Conciliados e Lançados")
    st.download_button(
        "📥 Extrair Relatório (Sem Divergência)",
        data=gerar_excel_estilizado(f_sem_div, "Sem_Divergencia"),
        file_name="Sem_Divergencia_Filtrado.xlsx",
        key="btn_dl_sem_div"
    )
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
        st.subheader("📜 Trilha de Auditoria e Histórico Completo de Alterações")
        
        col_la1, col_la2 = st.columns([3, 1])
        with col_la1:
            st.write(f"Total de eventos de auditoria registrados: **{len(df_log):,}**")
        with col_la2:
            if not df_log.empty:
                st.download_button(
                    "📥 Baixar Trilha de Auditoria (.xlsx)",
                    data=gerar_excel_estilizado(df_log, "Trilha_Auditoria_Completa"),
                    file_name=f"Trilha_Auditoria_Completa_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    key="btn_dl_log_completo",
                    type="primary"
                )
        st.dataframe(df_log, width="stretch", hide_index=True)

    # ABA 7: CARGA DE RELATÓRIOS (AO LADO DA AUDITORIA)
    with aba_sel[7]:
        st.subheader("📥 Carga de Relatórios Semanais em Lote (Acesso Restrito ao Master)")
        arqs_semanais = st.file_uploader(
            "Arraste ou selecione os arquivos Excel de retorno (.xlsx):",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="uploader_master_semanal"
        )
        if arqs_semanais and st.button("🚀 Ingestar e Atualizar Lote Semanal", type="primary"):
            df_m_novo, df_log_novo, df_novos_logs, total_lidos, qtd_adicionados = ingestar_lote_semanal(
                arqs_semanais, df_master, df_div_op, df_sem_div, df_backoffice, df_log
            )
            if qtd_adicionados > 0 or not df_novos_logs.empty:
                with st.spinner("💾 Consolidando e salvando bases do painel..."):
                    if salvar_bases(df_m_novo, df_div_op, df_sem_div, df_backoffice, df_log_novo):
                        st.session_state["ultimo_relatorio_auditoria"] = df_novos_logs
                        st.session_state["qtd_ultimos_alterados"] = qtd_adicionados
                        st.session_state["total_ultimos_lidos"] = total_lidos
                        st.success(f"🎉 Ingestão finalizada! {qtd_adicionados:,} registros/tratativas atualizados de {total_lidos:,} linhas lidas.")

        # Exibe o Relatório de Auditoria da Carga Atual (se houver carga recente na sessão)
        if "ultimo_relatorio_auditoria" in st.session_state and not st.session_state["ultimo_relatorio_auditoria"].empty:
            df_audit_carga = st.session_state["ultimo_relatorio_auditoria"]
            st.markdown("---")
            st.markdown("### 📋 Relatório de Auditoria das Alterações Processadas na Carga")
            
            cm1, cm2, cm3 = st.columns(3)
            cm1.metric("Total de Linhas Analisadas", f"{st.session_state.get('total_ultimos_lidos', 0):,}")
            cm2.metric("Alterações Aplicadas nesta Carga", f"{st.session_state.get('qtd_ultimos_alterados', 0):,}")
            cm3.metric("Data/Hora da Operação", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))

            st.download_button(
                "📥 Baixar Relatório de Auditoria desta Carga (Excel)",
                data=gerar_excel_estilizado(df_audit_carga, "Relatorio_Auditoria_Carga"),
                file_name=f"Relatorio_Auditoria_Carga_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                key="btn_dl_audit_carga_atual",
                type="primary"
            )

            st.dataframe(df_audit_carga, width="stretch", hide_index=True)