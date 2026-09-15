import datetime
import io
import json
import os
import re
import unicodedata
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
import streamlit as st

# ==============================================================================
# 1. CONFIGURAÇÃO INICIAL DA PÁGINA & ESTILOS CSS MODERNOS
# ==============================================================================
st.set_page_config(
    page_title="Grupo Arbaitman | Revenue Assurance",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_DASHBOARD = "Dashboard_Revenue_Assurance_Consolidado.xlsx"
ARQUIVO_USUARIOS = "usuarios_autorizados.json"

USUARIOS_PADRAO = {
    "mribeiro": {
        "senha": "123",
        "nome": "Marcos Ribeiro",
        "perfil": "Compliance",
        "status": "APROVADO",
        "data_solicitacao": "2026-08-31",
    },
    "compliance1": {
        "senha": "123",
        "nome": "Compliance - Auditoria 01",
        "perfil": "Compliance",
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

        div.stButton > button[kind="primary"], div.stButton > button {
            background-color: #002060 !important;
            color: #ffffff !important;
            border-radius: 8px !important;
            border: none !important;
            font-weight: 600 !important;
            padding: 8px 16px !important;
            transition: all 0.2s ease-in-out !important;
        }
        div.stButton > button:hover {
            background-color: #001040 !important;
            color: #ffffff !important;
            box-shadow: 0 4px 10px rgba(0, 32, 96, 0.25) !important;
        }

        .header-box {
            background: linear-gradient(135deg, #002060 0%, #003366 100%);
            padding: 18px 25px;
            border-radius: 12px;
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.08);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .header-box h1 { color: #ffffff !important; margin: 0; font-size: 24px; font-weight: 700; }
        .header-box p { color: #d0e0ff !important; margin-top: 4px; font-size: 13px; margin-bottom: 0; }

        [data-testid="stSidebar"] {
            background-color: #ffffff !important;
            border-right: 1px solid #e9ecef;
        }
        
        .user-card-sidebar {
            background: linear-gradient(135deg, #f0f4f9 0%, #e8eef6 100%);
            border-radius: 10px;
            padding: 14px;
            margin-bottom: 15px;
            border-left: 4px solid #002060;
            box-shadow: 0 2px 5px rgba(0,0,0,0.03);
        }

        div[data-testid="stSidebar"] div[role="radiogroup"] {
            gap: 6px;
        }
        div[data-testid="stSidebar"] div[role="radiogroup"] label {
            background-color: #f8f9fa;
            border: 1px solid #e2e8f0;
            border-radius: 8px !important;
            padding: 10px 14px !important;
            margin-bottom: 2px;
            transition: all 0.2s ease-in-out;
            cursor: pointer;
            width: 100%;
        }
        div[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
            background-color: #edf2f7;
            border-color: #cbd5e1;
            transform: translateX(3px);
        }
        div[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
            background: linear-gradient(135deg, #002060 0%, #003366 100%) !important;
            color: #ffffff !important;
            border-color: #002060 !important;
            font-weight: 600 !important;
            box-shadow: 0 4px 10px rgba(0, 32, 96, 0.2) !important;
        }
        div[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] p {
            color: #ffffff !important;
        }

        .login-card {
            background-color: #ffffff;
            padding: 30px 35px;
            border-radius: 12px;
            border-top: 6px solid #002060;
            box-shadow: 0 8px 24px rgba(0,0,0,0.08);
            margin-top: 20px;
        }
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 16px;
            border-left: 5px solid #002060;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }

        [data-baseweb="tag"], span[data-baseweb="tag"], div[data-baseweb="tag"] {
            background-color: #002060 !important;
            color: #ffffff !important;
            border-radius: 4px !important;
        }
        [data-baseweb="tag"] span { color: #ffffff !important; }

        .brand-header {
            font-size: 22px;
            font-weight: 800;
            color: #002060;
            letter-spacing: 1px;
            text-align: center;
            margin-bottom: 5px;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# ==============================================================================
# 2. FUNÇÕES DE SUPORTE, LIMPEZA VETORIZADA, REGRAS BSP HOT E BUSCA FLEXÍVEL
# ==============================================================================
def normalize_str(s):
    if not s:
        return ""
    s_unaccent = "".join(
        c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"[^a-z0-9]", "", s_unaccent.lower())


def clean_str_strict(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    if s.lower() in ["nan", "none", "null", "<na>"]:
        return ""
    return s


def extract_keys(val):
    if pd.isna(val) or val is None:
        return []
    s = clean_str_strict(val).upper()
    s = s.replace("[", "").replace("]", "").replace("R$", "").strip()
    if not s or s in ("NAN", "NONE", "NULL", "0", "-"):
        return []

    keys = set()
    clean_str = re.sub(r"[^A-Z0-9]", "", s)
    if clean_str:
        keys.add(clean_str)
        clean_no_zeros = clean_str.lstrip("0")
        if clean_no_zeros:
            keys.add(clean_no_zeros)
        if len(clean_str) == 13 and clean_str.isdigit():
            keys.add(clean_str[3:])
        if len(clean_str) >= 10 and clean_str.isdigit():
            keys.add(clean_str[-10:])

    tokens = re.split(r"[\s/\\,-]+", s)
    for tok in tokens:
        t_clean = re.sub(r"[^A-Z0-9]", "", tok)
        if t_clean:
            keys.add(t_clean)
            if t_clean.isdigit():
                keys.add(t_clean.lstrip("0"))
                if len(t_clean) >= 10:
                    keys.add(t_clean[-10:])

    return list(keys)


def find_column(df, candidates):
    if df is None or df.columns.empty:
        return None
    cols_norm = {c: normalize_str(c) for c in df.columns}

    # 1º passe: correspondência exata normalizada
    for cand in candidates:
        cand_norm = normalize_str(cand)
        for c, c_norm in cols_norm.items():
            if cand_norm == c_norm:
                return c

    # 2º passe: correspondência parcial
    for cand in candidates:
        cand_norm = normalize_str(cand)
        if not cand_norm:
            continue
        for c, c_norm in cols_norm.items():
            if cand_norm in c_norm:
                return c
    return None


def reordenar_colunas_visivel(df):
    if df is None or df.empty:
        return df

    cols_prioritarias = [
        "Ponto de venda",
        "Área Resp. Operação",
        "Obs. Operação",
        "Status_Geral",
        "Bilhetes",
        "Localizador_Sistema",
        "Rloc_Cia",
        "CIA",
        "📄 Origem / Arquivo",
        "Status_Divergencia",
        "Tipo_Inconsistencia",
        "Data Emissão",
        "Setor",
        "Consultor_Lemon",
    ]

    existentes = [c for c in cols_prioritarias if c in df.columns]
    outras = [c for c in df.columns if c not in existentes]

    return df[existentes + outras]


def detect_hot(row):
    campos_busca = [
        "CIA", "Origem_Aba", "Ponto de venda", "Tipo_Emissao_Lemon",
        "Setor", "Sistema", "Produto", "Fornecedor", "Arquivo"
    ]
    for col in campos_busca:
        val = str(row.get(col, "")).strip().upper()
        if "HOT" in val or "BSP" in val:
            return True
    return False


def categorizar_tipo_inconsistencia(row):
    origem = str(row.get("Origem_Aba", ""))
    status_div = str(row.get("Status_Divergencia", "")).strip()
    status_sis = str(row.get("Status_Sistema", "")).strip()
    status_geral = str(row.get("Status_Geral", "")).strip()

    is_hot = detect_hot(row)

    if is_hot:
        if any(term in status_div for term in ["Total", "Divergência Total", "Divergência no Total", "Divergência de Total"]):
            return "📄 BSP HOT - Divergência no Valor Total"
        elif status_div in ["Valores Corretos", "Sem_Divergencia", "", "nan"] or "Sem_Divergencia" in origem or "Já Lançado" in status_geral:
            return "Sem Divergência (Conciliado - BSP HOT)"
        else:
            return "Sem Divergência (Conciliado - BSP HOT Alocação OK)"

    if (
        origem == "99_Geral"
        or status_sis == "NAO_CONSTA"
        or "Pendente" in status_geral
    ):
        return "Pendente de Lançamento no ERP"
    elif status_div and status_div not in ["nan", "Valores Corretos", ""]:
        return status_div
    elif "Sem_Divergencia" in origem or status_div == "Valores Corretos":
        return "Sem Divergência (Conciliado)"
    else:
        return "Pendente de Lançamento no ERP"


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


def renderizar_marca():
    st.markdown(
        '<div class="brand-header">GRUPO ARBAITMAN</div>', unsafe_allow_html=True
    )


def gerar_excel_formatado(df_export, nome_aba="Relatorio_Filtrado"):
    buffer = io.BytesIO()
    if df_export is None or df_export.empty:
        df_export = pd.DataFrame(
            columns=["Aviso"],
            data=[["Nenhum registro encontrado para os filtros selecionados"]],
        )

    cols_remover = [
        c
        for c in ["Dt_Parsed", "Mes_Ano_Sort", "Bilhetes_Str"]
        if c in df_export.columns
    ]
    df_clean = (
        df_export.drop(columns=cols_remover) if cols_remover else df_export.copy()
    )
    df_clean = reordenar_colunas_visivel(df_clean)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name=nome_aba[:30], index=False)

    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    ws = wb.active

    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    hot_fill = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
    hot_font = Font(color="002060", bold=True, size=10)
    normal_font = Font(size=10)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    currency_cols = [
        "A vista",
        "A credito",
        "Taxa",
        "Comissão",
        "Taxa DU",
        "Desc.",
        "Incentivo",
        "VL. Líquido",
    ]
    ws.views.sheetView[0].showGridLines = True

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    max_row = ws.max_row
    max_col = ws.max_column

    col_names = [str(ws.cell(row=1, column=c).value or "") for c in range(1, max_col + 1)]
    hot_col_idx = None
    for idx_c, c_n in enumerate(col_names, 1):
        if c_n in ["É_HOT", "📄 Origem / Arquivo", "Origem_HOT"]:
            hot_col_idx = idx_c
            break

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        col_name = col_names[col_idx - 1]

        len_vals = [
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, max_row + 1)
        ]
        max_len = max(len_vals) if len_vals else 10
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        for r in range(2, max_row + 1):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = thin_border
            cell.font = normal_font

            is_hot_row = False
            if hot_col_idx:
                val_h = str(ws.cell(row=r, column=hot_col_idx).value or "")
                if "HOT" in val_h.upper() or val_h in ["True", "1"]:
                    is_hot_row = True

            if is_hot_row:
                cell.fill = hot_fill
                if col_idx == hot_col_idx or col_idx == 1:
                    cell.font = hot_font

            if col_name in currency_cols:
                try:
                    if cell.value is not None and str(cell.value).strip() not in ["-", ""]:
                        cell.value = float(cell.value)
                        cell.number_format = "R$ #,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                except:
                    pass

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# ==============================================================================
# 3. AUTENTICAÇÃO E SESSÃO DO USUÁRIO
# ==============================================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_atual" not in st.session_state:
    st.session_state["usuario_atual"] = None
if "perfil_atual" not in st.session_state:
    st.session_state["perfil_atual"] = None
if "login_user_id" not in st.session_state:
    st.session_state["login_user_id"] = None

if not st.session_state["autenticado"]:
    col_left, col_center, col_right = st.columns([1, 1.2, 1])

    with col_center:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        renderizar_marca()
        st.markdown(
            "<p style='text-align: center; color: #6c757d; font-size: 13px;"
            " margin-bottom: 20px;'>Maringá Turismo | Portal de Revenue"
            " Assurance</p>",
            unsafe_allow_html=True,
        )

        aba_login, aba_redefinir, aba_solicitar = st.tabs(
            ["🔐 Entrar", "🔑 Esqueci a Senha", "📝 Solicitar Acesso"]
        )

        with aba_login:
            user_input = (
                st.text_input("Usuário de Acesso:", key="login_user").strip().lower()
            )
            pass_input = st.text_input(
                "Senha:", type="password", key="login_pass"
            ).strip()
            btn_entrar = st.button("🚀 Acessar Sistema", type="primary")

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
                            st.warning(
                                "⏳ Sua solicitação de acesso está **Pendente de"
                                " Aprovação** pelo Compliance."
                            )
                        else:
                            st.error("❌ Acesso não autorizado.")
                    else:
                        st.error("❌ Senha incorreta.")
                else:
                    st.error("❌ Usuário não cadastrado.")

        with aba_redefinir:
            st.caption("Redefinição direta de senha de acesso.")
            with st.form("form_redefinir_senha_login"):
                user_reset = (
                    st.text_input("Informe seu Usuário de Acesso:").strip().lower()
                )
                nova_senha_login = st.text_input("Nova Senha:", type="password")
                confirma_senha_login = st.text_input(
                    "Confirme a Nova Senha:", type="password"
                )
                btn_redefinir_senha = st.form_submit_button("🔄 Redefinir Senha")

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
                        st.success(
                            "✅ Senha redefinida com sucesso! Clique na aba '🔐 Entrar' para acessar."
                        )

        with aba_solicitar:
            st.caption("Solicitação formal de acesso para novos colaboradores.")
            novo_nome = st.text_input("Nome Completo:")
            novo_user = (
                st.text_input("Usuário Desejado (ex: nome.sobrenome):")
                .strip()
                .lower()
            )
            nova_senha = st.text_input("Crie uma Senha:", type="password")
            novo_perfil = st.selectbox(
                "Perfil Solicitado:", options=["Operacao", "Compliance"]
            )

            st.markdown(
                """
                <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 10px; border-radius: 6px; font-size: 11px; color: #495057; max-height: 90px; overflow-y: scroll; margin-bottom: 10px;">
                    <b>TERMO DE CONFIDENCIALIDADE - GRUPO ARBAITMAN</b><br>
                    O usuário declara estar ciente do caráter confidencial das informações de Revenue Assurance. Qualquer alteração ou exportação de dados é monitorada na trilha de auditoria.
                </div>
                """,
                unsafe_allow_html=True,
            )

            termo_aceito = st.checkbox("Li e aceito os termos de sigilo")
            btn_cadastrar = st.button("📩 Enviar Solicitação")

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
                        "data_solicitacao": datetime.datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                    }
                    salvar_usuarios(usuarios_db)
                    st.success(
                        "✅ Solicitação enviada! Aguarde a liberação pelo Compliance."
                    )

        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# ==============================================================================
# 4. PADRONIZAÇÃO, CRONOLOGIA E PROCESSAMENTO VETORIZADO
# ==============================================================================
MAPA_MESES = {
    "01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr",
    "05": "Mai", "06": "Jun", "07": "Jul", "08": "Ago",
    "09": "Set", "10": "Out", "11": "Nov", "12": "Dez",
}


def padronizar_gerentes_e_setores_vector(series):
    s = series.astype(str).str.strip()
    s = s.replace({
        "Rosangela Pallu": "Rosângela Pallu",
        "ROSANGELA PALLU": "Rosângela Pallu",
        "ROSÂNGELA PALLU": "Rosângela Pallu",
        "CENTRAL DE EVENTOS": "Central de Eventos",
        "central de eventos": "Central de Eventos",
        "Jaime Schinaider": "Jaime Schnaider",
        "JAIME SCHNAIDER": "Jaime Schnaider",
        "Silvana Celane": "Silvana Celani",
        "SILVANA CELANI": "Silvana Celani",
        "nan": "Não Atribuído",
        "": "Não Atribuído",
        "None": "Não Atribuído",
    })
    mask_back = s.str.lower().str.contains("suporte|benner|katia", na=False)
    s.loc[mask_back] = "Suporte Backoffice"
    return s


def padronizar_e_deduplicar_colunas(df, origem=""):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df = df.loc[:, ~df.columns.duplicated()].copy()

    cols_norm = {c: normalize_str(c) for c in df.columns}

    has_area_resp = any(cn in ["arearespoperacao", "arearesponsavel"] for cn in cols_norm.values())
    has_gerentes = any(cn in ["gerente", "gerentes"] for cn in cols_norm.values())

    if has_area_resp and has_gerentes:
        cols_to_drop = [c for c, cn in cols_norm.items() if cn in ["gerente", "gerentes"]]
        df = df.drop(columns=cols_to_drop, errors="ignore")
        cols_norm = {c: normalize_str(c) for c in df.columns}

    col_pv = find_column(df, ["ponto de venda", "ponto_venda", "cliente", "ponto de venda / cliente"])
    col_bil = find_column(df, ["bilhetes", "bilhete", "nº bilhete", "bilhete/rloc", "ticket"])

    mask_descarte = pd.Series(False, index=df.index)
    if col_pv:
        mask_descarte = mask_descarte | df[col_pv].astype(str).str.strip().str.upper().str.contains("TOTAL", na=False)
    if col_bil:
        mask_descarte = (
            mask_descarte
            | df[col_bil].astype(str).str.strip().str.upper().str.contains("TOTAL", na=False)
            | df[col_bil].isna()
        )

    df = df[~mask_descarte].copy()
    cols_norm = {c: normalize_str(c) for c in df.columns}

    col_area_resp_orig = find_column(df, ["área responsável", "area responsavel", "área resp. operação", "area resp. operacao", "area_resp_operacao"])

    renomear = {}
    for col in df.columns:
        c_norm = cols_norm[col]
        if c_norm in ["arearespoperacao", "arearesponsavel"]:
            renomear[col] = "Área Resp. Operação"
        elif c_norm in ["gerente", "gerentes"]:
            renomear[col] = "Área Resp. Operação"
        elif c_norm in ["obs", "obsoperacao", "observacao", "observacoes"]:
            renomear[col] = "Obs. Operação"
        elif c_norm == "setor":
            renomear[col] = "Setor"
        elif c_norm in ["consultorlemon", "emissor", "consultor"]:
            renomear[col] = "Consultor_Lemon"
        elif c_norm in ["bilhete", "bilhetes", "nbilhete", "bilheterloc"]:
            renomear[col] = "Bilhetes"
        elif c_norm in ["rloccia", "rloc", "localizadorcia"]:
            renomear[col] = "Rloc_Cia"
        elif c_norm in ["dataemissao", "dataemissao"]:
            renomear[col] = "Data Emissão"
        elif c_norm in ["pontodevenda", "pontovenda", "cliente"]:
            renomear[col] = "Ponto de venda"

    df_out = df.rename(columns=renomear).copy()
    df_out = df_out.loc[:, ~df_out.columns.duplicated()].copy()

    for c, default_val in [
        ("Ponto de venda", "Geral"),
        ("Área Resp. Operação", "Não Atribuído"),
        ("Setor", "Geral"),
        ("Consultor_Lemon", "-"),
        ("Localizador_Sistema", "-"),
        ("Status_Geral", "Pendente de Lançamento"),
        ("Obs. Operação", ""),
        ("Data Emissão", "-"),
    ]:
        if c not in df_out.columns:
            df_out[c] = default_val

    if "Rloc_Cia" not in df_out.columns:
        df_out["Rloc_Cia"] = df_out["Localizador_Sistema"]

    df_out["Bilhetes"] = df_out["Bilhetes"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df_out["Localizador_Sistema"] = df_out["Localizador_Sistema"].astype(str).str.strip()
    df_out["Rloc_Cia"] = df_out["Rloc_Cia"].astype(str).str.strip()

    df_out["Área Resp. Operação"] = padronizar_gerentes_e_setores_vector(df_out["Área Resp. Operação"])
    df_out["Setor"] = np.where(
        df_out["Setor"].astype(str).str.upper() == "CENTRAL DE EVENTOS",
        "Central de Eventos",
        df_out["Setor"],
    )

    val_area_orig = df[col_area_resp_orig].astype(str).str.lower() if col_area_resp_orig and col_area_resp_orig in df.columns else pd.Series("", index=df.index)
    val_ger_orig = df_out["Área Resp. Operação"].astype(str).str.lower()
    val_obs_orig = df_out["Obs. Operação"].fillna("").astype(str).str.lower()

    mascara_suporte = (
        val_area_orig.str.contains("suporte benner|suporte backoffice|suporte ti|backoffice", na=False)
        | val_ger_orig.str.contains("suporte benner|suporte backoffice|katia martins", na=False)
        | val_obs_orig.str.contains("ticket 375|chamado no backoffice|chamado backoffice|erro de integração|erro integração|não integrou|não integração|aguardando suporte", na=False)
    )

    df_out.loc[mascara_suporte, "Área Resp. Operação"] = "Suporte Backoffice"
    df_out["Origem_Aba"] = origem

    df_out["É_HOT"] = df_out.apply(detect_hot, axis=1)
    df_out["📄 Origem / Arquivo"] = np.where(df_out["É_HOT"], "📄 BSP HOT (IATA - Apenas Total)", "✈️ Bilhete Regular")

    df_out["Tipo_Inconsistencia"] = df_out.apply(categorizar_tipo_inconsistencia, axis=1)

    dt_parsed = pd.to_datetime(df_out["Data Emissão"], format="mixed", dayfirst=True, errors="coerce")
    df_out["Dt_Parsed"] = dt_parsed

    df_out["Mes_Ano_Sort"] = np.where(
        dt_parsed.notna(),
        dt_parsed.dt.strftime("%Y-%m"),
        "0000-00"
    )

    m_num = dt_parsed.dt.strftime("%m")
    y_num = dt_parsed.dt.strftime("%Y")
    
    df_out["Mes_Ano_Label"] = np.where(
        dt_parsed.notna(),
        m_num.map(MAPA_MESES) + "/" + y_num,
        "Acumulado / Sem Data"
    )

    df_out = df_out.sort_values(by="Dt_Parsed", ascending=False, na_position="last")
    return reordenar_colunas_visivel(df_out)


# ==============================================================================
# 5. CARREGAMENTO COM CACHE DINÂMICO & PERSISTÊNCIA CENTRALIZADA
# ==============================================================================
def _get_file_mtime(filename):
    """Mapeia a data/hora exata da última modificação para invalidar a memória em tempo real."""
    return os.path.getmtime(filename) if os.path.exists(filename) else 0.0


@st.cache_data(show_spinner=False)
def carregar_bases_cached(mtime_key):
    if not os.path.exists(ARQUIVO_DASHBOARD):
        return None, None, None, None, None, "ARQUIVO_NAO_ENCONTRADO"
    try:
        xls = pd.ExcelFile(ARQUIVO_DASHBOARD, engine="openpyxl")
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
            df_log = pd.DataFrame(
                columns=[
                    "Data_Hora", "Bilhete", "Usuario_Acao", "Status_Anterior",
                    "Novo_Status", "Area_Anterior", "Nova_Area", "Observacao", "Tipo_Interacao"
                ]
            )

        return df_master, df_div_op, df_sem_div, df_backoffice, df_log, "OK"
    except PermissionError:
        return None, None, None, None, None, "ARQUIVO_BLOQUEADO"
    except Exception as e:
        return None, None, None, None, None, str(e)


def carregar_bases():
    return carregar_bases_cached(_get_file_mtime(ARQUIVO_DASHBOARD))


def salvar_base_consolidada(df_m, df_d, df_s, df_b, df_l):
    """Salva a base consolidada atualizando TODAS as abas do Excel mantendo a integridade das referências e novas linhas."""
    try:
        with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine="openpyxl") as writer:
            # 1. Pareto de Clientes
            if df_m is not None and not df_m.empty:
                df_m_clean = df_m.copy()
                cols_tarifa = [c for c in ["A vista", "A credito"] if c in df_m_clean.columns]
                df_m_clean["Tarifa_Total"] = df_m_clean[cols_tarifa].apply(pd.to_numeric, errors="coerce").sum(axis=1) if cols_tarifa else 0
                cols_rec = [c for c in ["Incentivo", "Comissão", "Taxa DU"] if c in df_m_clean.columns]
                df_m_clean["Receita_Total"] = df_m_clean[cols_rec].apply(pd.to_numeric, errors="coerce").sum(axis=1) if cols_rec else 0
                
                pareto_df = df_m_clean.groupby(["Setor", "Ponto de venda"]).agg(
                    Qtd_Bilhetes=("Bilhetes", "count"),
                    Tarifa_Pendente_R_=("Tarifa_Total", "sum"),
                    Taxa_Pendente_R_=("Taxa", "sum") if "Taxa" in df_m_clean.columns else ("Tarifa_Total", "count"),
                    Receita_Pendente_R_=("Receita_Total", "sum")
                ).reset_index().rename(columns={
                    "Ponto de venda": "Cliente / Ponto de Venda",
                    "Tarifa_Pendente_R_": "Tarifa_Pendente_R$",
                    "Taxa_Pendente_R_": "Taxa_Pendente_R$",
                    "Receita_Pendente_R_": "Receita_Pendente_R$"
                }).sort_values(by=["Qtd_Bilhetes", "Tarifa_Pendente_R$"], ascending=False)
                
                pareto_df.to_excel(writer, sheet_name="01_Pareto_Cliente", index=False)

            # 2. Abas Mestres (Garantindo que novos registros e referências antigas sejam salvos limpos)
            if df_d is not None:
                df_d_save = reordenar_colunas_visivel(df_d.drop_duplicates(subset=["Bilhetes"], keep="last")) if "Bilhetes" in df_d.columns else df_d
                df_d_save.to_excel(writer, sheet_name="98_OK_Divergencia_Operacao", index=False)
            if df_s is not None:
                df_s_save = reordenar_colunas_visivel(df_s.drop_duplicates(subset=["Bilhetes"], keep="last")) if "Bilhetes" in df_s.columns else df_s
                df_s_save.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
            if df_m is not None:
                df_m_save = reordenar_colunas_visivel(df_m.drop_duplicates(subset=["Bilhetes"], keep="last")) if "Bilhetes" in df_m.columns else df_m
                df_m_save.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)

            # 3. Abas Setoriais (Sincronizadas dinamicamente a partir da df_m atualizada)
            if df_m is not None and not df_m.empty and "Setor" in df_m.columns:
                df_m[df_m["Setor"] == "Suporte backoffice"].to_excel(writer, sheet_name="99_Suporte backoffice", index=False)
                df_m[df_m["Setor"] == "Central de Eventos"].to_excel(writer, sheet_name="99_Central de Eventos", index=False)
                df_m[df_m["Setor"] == "Concierge/Lazer"].to_excel(writer, sheet_name="99_Concierge-Lazer", index=False)
                df_m[df_m["Setor"] == "Unique"].to_excel(writer, sheet_name="99_Unique", index=False)
                df_m[df_m["Setor"] == "Private"].to_excel(writer, sheet_name="99_Private", index=False)
                df_m[df_m["Setor"] == "Operação"].to_excel(writer, sheet_name="99_Operação", index=False)

            # 4. Log de Auditoria
            if df_l is not None:
                df_l.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)

        return True, "Dashboard salvo com sucesso em TODAS as abas do Excel!"
    except Exception as e:
        return False, str(e)


# ==============================================================================
# 6. CARGA E MONTAGEM DA BASE DE DADOS GLOBAL
# ==============================================================================
df_master, df_div_op, df_sem_div, df_backoffice, df_log_master, status_carga = carregar_bases()

if status_carga == "ARQUIVO_BLOQUEADO":
    st.error(f"⚠️ O arquivo **'{ARQUIVO_DASHBOARD}'** está em uso. Feche a planilha para continuar.")
    if st.button("🔄 Recarregar Dados"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

if (df_master is None or df_master.empty) and (df_div_op is None or df_div_op.empty):
    st.error("⚠️ Base consolidada de divergências não encontrada.")
    st.stop()

COL_GERENTE = "Área Resp. Operação"
COL_SETOR = "Setor"
COL_EMISSOR = "Consultor_Lemon"

df_list = [d for d in [df_master, df_div_op, df_sem_div] if not d.empty]
df_acao_total = pd.concat(df_list, ignore_index=True, sort=False) if df_list else pd.DataFrame()

if "Bilhetes" in df_acao_total.columns:
    df_acao_total["Bilhetes_Str"] = df_acao_total["Bilhetes"].apply(clean_str_strict)

usuario_log_formatado = f"{st.session_state['usuario_atual']} ({st.session_state['login_user_id']})"

# Header principal
col_hdr1, col_hdr2 = st.columns([3, 1])
with col_hdr1:
    st.markdown(
        f"""
        <div class="header-box">
            <div>
                <h1>GRUPO ARBAITMAN | Revenue Assurance</h1>
                <p>Maringá Turismo — Conectado como: <b>{st.session_state['usuario_atual']}</b> | Perfil: <b>{st.session_state['perfil_atual']}</b></p>
            </div>
        </div>
    """,
        unsafe_allow_html=True,
    )

with col_hdr2:
    if st.button("🔄 Recarregar Memória", type="primary"):
        st.cache_data.clear()
        st.rerun()

if "msg_sucesso" in st.session_state:
    st.success(st.session_state["msg_sucesso"])
    del st.session_state["msg_sucesso"]

renderizar_marca()

# ==============================================================================
# 7. SIDEBAR & MENU DE NAVEGAÇÃO
# ==============================================================================
st.sidebar.markdown(
    f"""
    <div class="user-card-sidebar">
        <div style="font-size: 11px; color: #6c757d; font-weight: 700; text-transform: uppercase;">Usuário Conectado</div>
        <div style="font-size: 15px; font-weight: 700; color: #002060;">👤 {st.session_state['usuario_atual']}</div>
        <div style="font-size: 12px; color: #495057;">Perfil: <b>{st.session_state['perfil_atual']}</b></div>
    </div>
    """,
    unsafe_allow_html=True,
)

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
            btn_upd_pwd = st.form_submit_button("💾 Atualizar Senha")

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

opcoes_navegacao = [
    "📊 Dashboard & KPIs",
    "🎯 Tratativa Operacional (Geral)",
    "⚠️ Divergência Operação (CIAs/BSP HOT)",
    "✅ Sem Divergência (Conciliação)",
    "🎧 Suporte Backoffice",
    "⚖️ Réplica da Auditoria",
    "📋 Visão Geral da Base Total",
]

if st.session_state["perfil_atual"] == "Compliance":
    opcoes_navegacao.insert(6, "📜 Trilha de Auditoria")
    opcoes_navegacao.insert(7, "⚙️ Gestão de Acessos")

if "aba_selecionada" not in st.session_state:
    st.session_state["aba_selecionada"] = "🎯 Tratativa Operacional (Geral)"

aba_atual = st.sidebar.radio(
    "📌 Menu de Navegação:",
    options=opcoes_navegacao,
    key="aba_selecionada",
)

st.sidebar.markdown("---")
st.sidebar.title("🔍 Filtros Operacionais")

# ==============================================================================
# 8. FILTROS GLOBAIS COM ORDENAÇÃO CRONOLÓGICA E FILTRO BSP HOT
# ==============================================================================
df_meses_ord = (
    df_acao_total[df_acao_total["Mes_Ano_Label"].notna()]
    .drop_duplicates(subset=["Mes_Ano_Sort", "Mes_Ano_Label"])
    .sort_values(by="Mes_Ano_Sort", ascending=True)
)

opcoes_meses_ordenadas = [
    m for m in df_meses_ord["Mes_Ano_Label"].tolist() if m != "Acumulado / Sem Data"
]
if "Acumulado / Sem Data" in df_acao_total["Mes_Ano_Label"].values:
    opcoes_meses_ordenadas.append("Acumulado / Sem Data")

filtro_hot = st.sidebar.radio(
    "📄 Origem / Tipo de Arquivo:",
    options=["Todos", "📄 Somente BSP HOT (IATA - Apenas Total)", "✈️ Somente Bilhetes Regular"],
    index=0,
)

mes_sel = st.sidebar.multiselect("📅 Mês de Emissão:", options=opcoes_meses_ordenadas, default=[])

df_f_mes = (
    df_acao_total[df_acao_total["Mes_Ano_Label"].isin(mes_sel)]
    if len(mes_sel) > 0
    else df_acao_total
)

opcoes_setor = sorted(list(df_f_mes[COL_SETOR].dropna().astype(str).unique()))
setor_sel = st.sidebar.multiselect("Setor Responsável:", options=opcoes_setor, default=[])

df_f_1 = df_f_mes[df_f_mes[COL_SETOR].astype(str).isin(setor_sel)] if len(setor_sel) > 0 else df_f_mes

opcoes_gerente = sorted(list(df_f_1[COL_GERENTE].dropna().astype(str).unique()))
gerente_sel = st.sidebar.multiselect("Gerente (Área Resp. Operação):", options=opcoes_gerente, default=[])

df_f_2 = df_f_1[df_f_1[COL_GERENTE].astype(str).isin(gerente_sel)] if len(gerente_sel) > 0 else df_f_1

opcoes_emissor = sorted(list(df_f_2[COL_EMISSOR].dropna().astype(str).unique()))
emissor_sel = st.sidebar.multiselect("Emissor / Consultor (OBT):", options=opcoes_emissor, default=[])

df_f_3 = df_f_2[df_f_2[COL_EMISSOR].astype(str).isin(emissor_sel)] if len(emissor_sel) > 0 else df_f_2

tipos_emissao = (
    sorted(list(df_f_3["Tipo_Emissao_Lemon"].dropna().astype(str).unique()))
    if "Tipo_Emissao_Lemon" in df_f_3.columns
    else []
)
emissao_sel = st.sidebar.multiselect("Tipo de Emissão (OBT):", options=tipos_emissao, default=[])

df_f_4 = (
    df_f_3[df_f_3["Tipo_Emissao_Lemon"].astype(str).isin(emissao_sel)]
    if len(emissao_sel) > 0 and "Tipo_Emissao_Lemon" in df_f_3.columns
    else df_f_3
)

cias = sorted(list(df_f_4["CIA"].dropna().astype(str).unique())) if "CIA" in df_f_4.columns else []
cia_sel = st.sidebar.multiselect("Companhia Aérea:", options=cias, default=[])


def aplicar_filtros_globais(df):
    if df is None or df.empty:
        return df
    m = pd.Series(True, index=df.index)
    if "É_HOT" in df.columns:
        if filtro_hot == "📄 Somente BSP HOT (IATA - Apenas Total)":
            m = m & (df["É_HOT"] == True)
        elif filtro_hot == "✈️ Somente Bilhetes Regular":
            m = m & (df["É_HOT"] == False)

    if "Mes_Ano_Label" in df.columns and len(mes_sel) > 0:
        m = m & df["Mes_Ano_Label"].isin(mes_sel)
    if COL_SETOR in df.columns and len(setor_sel) > 0:
        m = m & df[COL_SETOR].astype(str).isin(setor_sel)
    if COL_GERENTE in df.columns and len(gerente_sel) > 0:
        m = m & df[COL_GERENTE].astype(str).isin(gerente_sel)
    if COL_EMISSOR in df.columns and len(emissor_sel) > 0:
        m = m & df[COL_EMISSOR].astype(str).isin(emissor_sel)
    if "Tipo_Emissao_Lemon" in df.columns and len(emissao_sel) > 0:
        m = m & df["Tipo_Emissao_Lemon"].astype(str).isin(emissao_sel)
    if "CIA" in df.columns and len(cia_sel) > 0:
        m = m & df["CIA"].astype(str).isin(cia_sel)
    return df[m].copy()


df_master_filtrado = aplicar_filtros_globais(df_master)
df_div_op_filtrado = aplicar_filtros_globais(df_div_op)
df_sem_div_filtrado = aplicar_filtros_globais(df_sem_div)
df_acao_filtrado = aplicar_filtros_globais(df_acao_total)

mascara_back_m = df_master[COL_GERENTE].astype(str).str.lower().str.contains("suporte backoffice|suporte benner|katia martins", na=False) if not df_master.empty else pd.Series(False)
mascara_back_d = df_div_op[COL_GERENTE].astype(str).str.lower().str.contains("suporte backoffice|suporte benner|katia martins", na=False) if not df_div_op.empty else pd.Series(False)

df_b_copy = df_backoffice.copy()
df_m_copy = df_master[mascara_back_m].copy() if not df_master.empty else pd.DataFrame()
df_d_copy = df_div_op[mascara_back_d].copy() if not df_div_op.empty else pd.DataFrame()

df_backoffice_dinamico = (
    pd.concat([df_b_copy, df_m_copy, df_d_copy], ignore_index=True).drop_duplicates(subset=["Bilhetes"], keep="last")
    if "Bilhetes" in df_master.columns and not df_master.empty
    else df_backoffice
)

df_backoffice_filtrado = aplicar_filtros_globais(df_backoffice_dinamico)

gerentes_base_unicos = sorted([
    g for g in df_acao_total[COL_GERENTE].dropna().astype(str).unique()
    if str(g).lower() not in ["suporte backoffice", "suporte benner", "não atribuído", "-", "katia martins"]
]) if not df_acao_total.empty and COL_GERENTE in df_acao_total.columns else []

dt_str_export = datetime.datetime.now().strftime("%Y%m%d_%H%M")

# ==============================================================================
# 9. CONTEÚDO DAS ABAS / NAVEGAÇÃO
# ==============================================================================

# ABA 0: DASHBOARD INTERATIVO
if aba_atual == "📊 Dashboard & KPIs":
    st.subheader("📊 Painel Executivo e Métricas Globais (100% da Base Auditada)")

    a_vista_sum = pd.to_numeric(df_acao_filtrado["A vista"], errors="coerce").fillna(0.0).sum() if "A vista" in df_acao_filtrado.columns else 0.0
    a_credito_sum = pd.to_numeric(df_acao_filtrado["A credito"], errors="coerce").fillna(0.0).sum() if "A credito" in df_acao_filtrado.columns else 0.0
    val_tarifa = a_vista_sum + a_credito_sum

    val_taxa = pd.to_numeric(df_acao_filtrado["Taxa"], errors="coerce").fillna(0.0).sum() if "Taxa" in df_acao_filtrado.columns else 0.0

    inc_sum = pd.to_numeric(df_acao_filtrado["Incentivo"], errors="coerce").fillna(0.0).sum() if "Incentivo" in df_acao_filtrado.columns else 0.0
    com_sum = pd.to_numeric(df_acao_filtrado["Comissão"], errors="coerce").fillna(0.0).sum() if "Comissão" in df_acao_filtrado.columns else 0.0
    tdu_sum = pd.to_numeric(df_acao_filtrado["Taxa DU"], errors="coerce").fillna(0.0).sum() if "Taxa DU" in df_acao_filtrado.columns else 0.0
    val_receita = inc_sum + com_sum + tdu_sum

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Bilhetes Auditados em Ação", f"{len(df_acao_filtrado):,}")
    k2.metric("Tarifa Pendente (À vista + À crédito)", f"R$ {val_tarifa:,.2f}")
    k3.metric("Taxas Pendentes (Taxas)", f"R$ {val_taxa:,.2f}")
    k4.metric("Receita em Risco (Incentivo + Comissões)", f"R$ {val_receita:,.2f}")

    st.markdown("---")

    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.markdown("##### ⚠️ Classificação Detalhada dos Tipos de Inconsistência")
        if not df_acao_filtrado.empty and "Tipo_Inconsistencia" in df_acao_filtrado.columns:
            df_inc = df_acao_filtrado["Tipo_Inconsistencia"].value_counts().reset_index()
            df_inc.columns = ["Inconsistencia", "Quantidade"]

            fig_inc = px.bar(
                df_inc, x="Quantidade", y="Inconsistencia", orientation="h",
                text="Quantidade", color="Inconsistencia",
                color_discrete_sequence=["#002060", "#2b9348", "#c1121f", "#7209b7", "#4361ee", "#4cc9f0"],
            )
            fig_inc.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
            fig_inc.update_layout(
                height=380, xaxis_title="Qtd. Bilhetes", yaxis_title="",
                showlegend=False, margin=dict(l=10, r=30, t=20, b=20),
            )
            st.plotly_chart(fig_inc, use_container_width=True)
        else:
            st.info("Sem dados para exibir o gráfico.")

    with col_chart2:
        st.markdown("##### 👤 Volumetria Total por Gerente Responsável")
        if not df_acao_filtrado.empty and COL_GERENTE in df_acao_filtrado.columns:
            df_ger = df_acao_filtrado[COL_GERENTE].value_counts().head(10).reset_index()
            df_ger.columns = ["Gerente", "Quantidade"]

            fig_bar = px.bar(
                df_ger, x="Gerente", y="Quantidade", text="Quantidade",
                color="Quantidade", color_continuous_scale=["#00509d", "#002060"],
            )
            fig_bar.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
            fig_bar.update_layout(
                height=380, xaxis_title="", yaxis_title="Qtd. Bilhetes",
                coloraxis_showscale=False, margin=dict(l=10, r=10, t=20, b=20),
                xaxis=dict(tickangle=-25),
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Sem dados para exibir o gráfico.")

    st.markdown("---")

    st.markdown("##### 📈 Volumetria Mensal de Pendências por Setor Responsável")
    if not df_acao_filtrado.empty and "Mes_Ano_Label" in df_acao_filtrado.columns:
        df_mes_setor = (
            df_acao_filtrado.groupby(["Mes_Ano_Sort", "Mes_Ano_Label", COL_SETOR])
            .size()
            .reset_index(name="Quantidade")
            .sort_values(by="Mes_Ano_Sort", ascending=True)
        )

        fig_line = px.bar(
            df_mes_setor, x="Mes_Ano_Label", y="Quantidade", color=COL_SETOR,
            barmode="group", text="Quantidade",
            color_discrete_sequence=["#002060", "#0077b6", "#7209b7", "#4361ee", "#4cc9f0", "#f72585"],
        )
        fig_line.update_traces(textposition="outside", cliponaxis=False)
        fig_line.update_layout(
            height=350, xaxis_title="Mês da Emissão", yaxis_title="Volume de Bilhetes",
            legend_title="Setor", margin=dict(l=10, r=10, t=20, b=20),
        )
        fig_line.update_xaxes(
            categoryorder="array",
            categoryarray=opcoes_meses_ordenadas
        )
        st.plotly_chart(fig_line, use_container_width=True)

# ABA 1: TRATATIVA OPERACIONAL (GERAL)
elif aba_atual == "🎯 Tratativa Operacional (Geral)":
    st.subheader("📝 Módulo de Resolução Operacional (Atribuição Individual ou em Lote)")

    # UPLOAD COM BARRA DE PROGRESSO & RELATÓRIO DE ALTERAÇÕES (INCLUI INGESTÃO DE NOVAS LINHAS E PRESERVAÇÃO DE REFERÊNCIAS ANTIGAS)
    with st.expander("📥 Carga de Retornos Gerenciais (Upload Otimizado de Planilhas de Gerentes)", expanded=False):
        st.caption("Suba as planilhas enviadas pelos gerentes. O sistema atualiza os registros existentes mantendo referências e INGESTIONA novas linhas automaticamente.")
        
        arquivos_retorno = st.file_uploader(
            "Arraste ou selecione os arquivos Excel de retorno dos gerentes:",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="uploader_retornos_gerenciais"
        )

        if arquivos_retorno and st.button("🚀 Processar Retornos e Atualizar Dashboard"):
            total_atualizados = 0
            total_novos_inseridos = 0
            novos_logs_retorno = []
            relatorio_modificados = []

            tot_arqs = len(arquivos_retorno)
            bar_progresso = st.progress(0, text="⚡ Iniciando leitura e indexação dos arquivos...")

            with st.status("⚡ Processando arquivos de retorno em alta velocidade...", expanded=True) as status:
                st.write("🔍 Indexando base de dados usando os índices reais para busca em milissegundos...")
                
                map_index = {}
                for target_name, target_df in [("df_master", df_master), ("df_div_op", df_div_op), ("df_sem_div", df_sem_div)]:
                    if target_df is not None and not target_df.empty:
                        for idx in target_df.index:
                            b_val = target_df.loc[idx, "Bilhetes"] if "Bilhetes" in target_df.columns else ""
                            loc_sys = target_df.loc[idx, "Localizador_Sistema"] if "Localizador_Sistema" in target_df.columns else ""
                            rloc_cia = target_df.loc[idx, "Rloc_Cia"] if "Rloc_Cia" in target_df.columns else ""
                            
                            keys_all = set(extract_keys(b_val) + extract_keys(loc_sys) + extract_keys(rloc_cia))
                            for k in keys_all:
                                if k not in map_index:
                                    map_index[k] = []
                                map_index[k].append((target_name, idx))

                for num_arq, arq in enumerate(arquivos_retorno, 1):
                    prog_pct = int((num_arq / tot_arqs) * 100)
                    bar_progresso.progress(num_arq / tot_arqs, text=f"📂 Processando arquivo {num_arq} de {tot_arqs} ({prog_pct}%)...")
                    st.write(f"📄 Lendo arquivo {num_arq}/{tot_arqs}: **{arq.name}**...")
                    try:
                        wb_check = openpyxl.load_workbook(arq, read_only=True, data_only=True)
                        all_sheets = wb_check.sheetnames
                        wb_check.close()
                        
                        target_sheets = [s for s in all_sheets if any(term in s.upper() for term in ["NÃO ENCONTRADOS", "NAO ENCONTRADOS", "NÃO CONCILIADOS", "NAO CONCILIADOS", "PENDENTES", "RETORNO"])]
                        if not target_sheets:
                            target_sheets = all_sheets

                        for sheet_name in target_sheets:
                            st.write(f"    ↳ Processando aba: **{sheet_name}**...")
                            df_ret = pd.read_excel(arq, sheet_name=sheet_name)
                            
                            col_bil = find_column(df_ret, ["bilhetes", "bilhete", "nº bilhete", "bilhete/rloc", "localizador", "loc", "ticket", "rloc_cia", "rloc"])
                            if not col_bil:
                                continue

                            df_ret = df_ret.dropna(subset=[col_bil]).copy()
                            if df_ret.empty:
                                continue

                            col_obs = find_column(df_ret, ["obs. operação", "obs", "observação", "observacoes", "observacao", "observações", "tratativa", "resolução", "parecer", "justificativa", "obs operação", "detalhes"])
                            col_area = find_column(df_ret, ["área resp. operação", "area resp. operacao", "área responsável", "area responsavel", "gerentes", "gerente", "setor", "área"])
                            col_status = find_column(df_ret, ["status da tratativa", "status_geral", "status geral", "status", "situação", "situacao", "resultado"])

                            ret_records = df_ret.to_dict("records")
                            for r_ret in ret_records:
                                b_ret = r_ret.get(col_bil, "")
                                keys_ret = extract_keys(b_ret)
                                if not keys_ret:
                                    continue

                                nova_obs = clean_str_strict(r_ret.get(col_obs, "")) if col_obs else ""
                                nova_area = clean_str_strict(r_ret.get(col_area, "")) if col_area else ""
                                novo_status_raw = clean_str_strict(r_ret.get(col_status, "")) if col_status else ""

                                if nova_obs.lower() in ["sem tratativa na operação", "nan", "-", "none", "null"]:
                                    nova_obs = ""
                                if nova_area.lower() in ["nan", "não atribuído", "-", "none", "null"]:
                                    nova_area = ""

                                is_lancado = False
                                if any(w in novo_status_raw.lower() for w in ["já lançado", "lançado", "conciliado", "ok", "lançado no erp"]):
                                    is_lancado = True
                                elif any(w in nova_obs.lower() for w in ["lançado no benner", "lançado erp", "já lançado", "cadastrado no benner", "regularizado"]):
                                    is_lancado = True

                                matched_targets = set()
                                for k in keys_ret:
                                    if k in map_index:
                                        for item in map_index[k]:
                                            matched_targets.add(item)

                                # CASO 1: REGISTRO JÁ EXISTE (Atualiza mantendo referências antigas)
                                if matched_targets:
                                    for target_name, idx in matched_targets:
                                        if target_name == "df_master" and idx in df_master.index:
                                            target_df = df_master
                                        elif target_name == "df_div_op" and idx in df_div_op.index:
                                            target_df = df_div_op
                                        elif target_name == "df_sem_div" and idx in df_sem_div.index:
                                            target_df = df_sem_div
                                        else:
                                            continue

                                        obs_ant = clean_str_strict(target_df.loc[idx, "Obs. Operação"]) if "Obs. Operação" in target_df.columns else ""
                                        area_ant = clean_str_strict(target_df.loc[idx, COL_GERENTE]) if COL_GERENTE in target_df.columns else ""
                                        status_ant = clean_str_strict(target_df.loc[idx, "Status_Geral"]) if "Status_Geral" in target_df.columns else ""

                                        alterou = False
                                        area_final = area_ant
                                        if nova_area and nova_area != area_ant:
                                            area_final = padronizar_gerentes_e_setores_vector(pd.Series([nova_area])).iloc[0]
                                            target_df.loc[idx, COL_GERENTE] = area_final
                                            alterou = True

                                        obs_final = obs_ant
                                        if nova_obs and nova_obs != obs_ant:
                                            obs_final = nova_obs
                                            target_df.loc[idx, "Obs. Operação"] = obs_final
                                            alterou = True

                                        b_str_val = clean_str_strict(target_df.loc[idx, "Bilhetes"])
                                        novo_status_final = "Já Lançado no ERP" if is_lancado else status_ant

                                        if is_lancado and status_ant != "Já Lançado no ERP":
                                            target_df.loc[idx, "Status_Geral"] = "Já Lançado no ERP"
                                            target_df.loc[idx, "Status_Divergencia"] = "Valores Corretos"
                                            target_df.loc[idx, "Tipo_Inconsistencia"] = categorizar_tipo_inconsistencia(target_df.loc[idx])
                                            alterou = True

                                            if target_name in ["df_master", "df_div_op"]:
                                                row_moved = target_df.loc[[idx]].copy()
                                                row_moved["Status_Geral"] = "Já Lançado no ERP"
                                                row_moved["Status_Divergencia"] = "Valores Corretos"
                                                row_moved["Tipo_Inconsistencia"] = categorizar_tipo_inconsistencia(row_moved.iloc[0])
                                                if nova_area:
                                                    row_moved[COL_GERENTE] = area_final
                                                if nova_obs:
                                                    row_moved["Obs. Operação"] = obs_final

                                                if target_name == "df_master":
                                                    df_master = df_master.drop(idx)
                                                elif target_name == "df_div_op":
                                                    df_div_op = df_div_op.drop(idx)

                                                df_sem_div = pd.concat([df_sem_div, row_moved], ignore_index=True)

                                        if alterou:
                                            total_atualizados += 1
                                            novos_logs_retorno.append({
                                                "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                "Bilhete": b_str_val,
                                                "Usuario_Acao": usuario_log_formatado,
                                                "Status_Anterior": status_ant,
                                                "Novo_Status": novo_status_final,
                                                "Area_Anterior": area_ant,
                                                "Nova_Area": area_final,
                                                "Observacao": f"[Retorno Aba: {sheet_name}]: {obs_final}",
                                                "Tipo_Interacao": "Carga de Retorno Gerencial",
                                            })
                                            relatorio_modificados.append({
                                                "Aba de Origem": sheet_name,
                                                "Nº Bilhete / LOC": b_str_val,
                                                "Área Anterior": area_ant,
                                                "Nova Área / Gerente": area_final,
                                                "Status Anterior": status_ant,
                                                "Novo Status": novo_status_final,
                                                "Observação Aplicada": obs_final,
                                                "Ação Executada": "Conciliado (Lançado ERP)" if is_lancado else "Tratativa Atualizada"
                                            })

                                # CASO 2: REGISTRO NOVO (Alimentação de novas linhas)
                                else:
                                    row_nova_dict = {str(k): v for k, v in r_ret.items() if pd.notna(v)}
                                    b_novo = clean_str_strict(b_ret)
                                    row_nova_dict["Bilhetes"] = b_novo
                                    row_nova_dict["Obs. Operação"] = nova_obs
                                    row_nova_dict[COL_GERENTE] = nova_area if nova_area else "Não Atribuído"
                                    row_nova_dict["Status_Geral"] = "Já Lançado no ERP" if is_lancado else "Pendente de Lançamento"
                                    
                                    df_nova_row = pd.DataFrame([row_nova_dict])
                                    df_nova_row = padronizar_e_deduplicar_colunas(df_nova_row, origem=f"Upload_{sheet_name}")
                                    
                                    if is_lancado:
                                        df_sem_div = pd.concat([df_sem_div, df_nova_row], ignore_index=True)
                                        target_new_name = "df_sem_div"
                                        new_idx = df_sem_div.index[-1]
                                    else:
                                        df_master = pd.concat([df_master, df_nova_row], ignore_index=True)
                                        target_new_name = "df_master"
                                        new_idx = df_master.index[-1]

                                    # Atualiza o mapa de busca para evitar duplicatas futuras
                                    for k in keys_ret:
                                        if k not in map_index:
                                            map_index[k] = []
                                        map_index[k].append((target_new_name, new_idx))

                                    total_novos_inseridos += 1
                                    total_atualizados += 1
                                    
                                    novos_logs_retorno.append({
                                        "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "Bilhete": b_novo,
                                        "Usuario_Acao": usuario_log_formatado,
                                        "Status_Anterior": "NOVO REGISTRO",
                                        "Novo_Status": "Já Lançado no ERP" if is_lancado else "Pendente de Lançamento",
                                        "Area_Anterior": "-",
                                        "Nova_Area": row_nova_dict[COL_GERENTE],
                                        "Observacao": f"[Nova Linha Carga Aba: {sheet_name}]: {nova_obs}",
                                        "Tipo_Interacao": "Inclusão de Novo Bilhete via Carga",
                                    })
                                    relatorio_modificados.append({
                                        "Aba de Origem": sheet_name,
                                        "Nº Bilhete / LOC": b_novo,
                                        "Área Anterior": "-",
                                        "Nova Área / Gerente": row_nova_dict[COL_GERENTE],
                                        "Status Anterior": "N/A (Novo Registro)",
                                        "Novo Status": "Já Lançado no ERP" if is_lancado else "Pendente de Lançamento",
                                        "Observação Aplicada": nova_obs,
                                        "Ação Executada": "Novo Registro Inserido"
                                    })

                    except Exception as e:
                        st.error(f"Erro ao processar o arquivo '{arq.name}': {str(e)}")

                bar_progresso.progress(1.0, text="✅ Processamento e sincronização concluídos (100%)!")
                status.update(label="✅ Processamento concluído com sucesso!", state="complete")

            if total_atualizados > 0:
                mascara_back_upd = df_master[COL_GERENTE].astype(str).str.lower().str.contains("suporte backoffice|suporte benner|katia martins", na=False) if not df_master.empty else pd.Series(False)
                df_back_atualizado = df_master[mascara_back_upd].copy() if not df_master.empty else pd.DataFrame()
                df_log_updated = pd.concat([df_log_master, pd.DataFrame(novos_logs_retorno)], ignore_index=True)

                sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, df_back_atualizado, df_log_updated)
                if sucesso_save:
                    st.session_state["msg_sucesso"] = f"🎉 Processamento Finalizado! {total_atualizados} registro(s) processados ({total_novos_inseridos} novos inseridos) e salvos no Dashboard."
                    st.session_state["relatorio_modificados"] = relatorio_modificados
                    st.rerun()
                else:
                    st.error(err_msg)
            else:
                st.warning("⚠️ Nenhuma nova alteração válida foi encontrada nas abas processadas.")

    if "relatorio_modificados" in st.session_state and st.session_state["relatorio_modificados"]:
        df_rel_mod = pd.DataFrame(st.session_state["relatorio_modificados"])
        st.markdown(
            f"""
            <div style="background-color: #e6f4ea; border-left: 5px solid #2b9348; padding: 15px; border-radius: 8px; margin-top: 15px; margin-bottom: 15px;">
                <h4 style="color: #1e4620; margin: 0;">✅ Retornos Processados e Salvos na Planilha Oficial!</h4>
                <p style="color: #2b9348; margin-top: 4px; margin-bottom: 0;">Foram realizadas <b>{len(df_rel_mod)} alteração(ões) / inclusão(ões)</b>. Baixe o relatório detalhado abaixo:</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        col_rm1, col_rm2, col_rm3 = st.columns([2.5, 1.2, 1])
        with col_rm1:
            st.markdown(f"##### 📄 Relação Detalhada das Alterações e Novas Linhas ({len(df_rel_mod)} registros)")
        with col_rm2:
            st.download_button(
                label="📥 Extrair Relatório (Excel)",
                data=gerar_excel_formatado(df_rel_mod, "Relatorio_Alteracoes"),
                file_name=f"Relatorio_Alteracoes_Processadas_{dt_str_export}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_exp_casos_modificados_top",
                type="primary"
            )
        with col_rm3:
            if st.button("❌ Fechar Painel", key="btn_fechar_painel_relatorio"):
                del st.session_state["relatorio_modificados"]
                st.rerun()

        st.dataframe(reordenar_colunas_visivel(df_rel_mod), hide_index=True)

    st.markdown("---")

    if len(df_master_filtrado) == 0:
        st.warning("Nenhum bilhete encontrado na base geral para os filtros selecionados.")
    else:
        lista_busca_geral = df_master_filtrado["Bilhetes"].apply(clean_str_strict).tolist()

        bilhetes_selecionados = st.multiselect(
            "⚡ Selecione UM ou MAIS Bilhetes / LOCs para Alteração em Lote:",
            options=lista_busca_geral,
            default=[lista_busca_geral[0]] if lista_busca_geral else [],
            key="sb_geral_multi",
        )

        if not bilhetes_selecionados:
            st.info("💡 Selecione ao menos um bilhete para realizar a tratativa.")
        else:
            st.info(f"⚡ **{len(bilhetes_selecionados)} bilhete(s) selecionado(s)** para atualização simultânea.")

            df_previa = df_master_filtrado[
                df_master_filtrado["Bilhetes"].apply(clean_str_strict).isin(bilhetes_selecionados)
            ]
            
            cols_desejadas_g = ["Ponto de venda", "Área Resp. Operação", "Obs. Operação", "CIA", "📄 Origem / Arquivo", "Bilhetes", "Localizador_Sistema", "Rloc_Cia", "Status_Geral"]
            cols_exibir_g = [c for c in cols_desejadas_g if c in df_previa.columns]
            
            st.dataframe(df_previa[cols_exibir_g], hide_index=True)

            with st.form("form_tratativa_geral_multi"):
                col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1.5, 1.2])
                with col_a:
                    novo_status = st.selectbox(
                        "Status da Tratativa:",
                        options=["Já Lançado no ERP", "Pendente de Lançamento", "Aguardando TI", "Cancelado / Devolvido"],
                    )
                with col_b:
                    areas_opcoes = ["Operação", "Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private"]
                    nova_area = st.selectbox("Área Responsável (Reatribuir Todos):", options=areas_opcoes, index=0)

                with col_c:
                    if nova_area == "Private":
                        st.text_input("Gerente Responsável:", value="Silvana Celani", disabled=True)
                        gerente_indicado_sel, novo_gerente_texto = "Silvana Celani", ""
                    elif nova_area == "Unique":
                        st.text_input("Gerente Responsável:", value="Jaime Schnaider", disabled=True)
                        gerente_indicado_sel, novo_gerente_texto = "Jaime Schnaider", ""
                    elif nova_area == "Concierge/Lazer":
                        st.text_input("Gerente Responsável:", value="Fabiano Souza", disabled=True)
                        gerente_indicado_sel, novo_gerente_texto = "Fabiano Souza", ""
                    elif nova_area == "Central de Eventos":
                        st.text_input("Gerente Responsável:", value="Alexandre Souza", disabled=True)
                        gerente_indicado_sel, novo_gerente_texto = "Alexandre Souza", ""
                    elif nova_area == "Suporte backoffice":
                        st.text_input("Gerente Responsável:", value="Suporte Backoffice", disabled=True)
                        gerente_indicado_sel, novo_gerente_texto = "Suporte Backoffice", ""
                    else:
                        gerentes_reservados = ["Silvana Celani", "Jaime Schnaider", "Fabiano Souza", "Alexandre Souza", "Central de Eventos", "Suporte Backoffice", "Katia Martins"]
                        gerentes_operacao_puros = sorted(list(set([
                            g for g in gerentes_base_unicos + ["Keli Santi", "Guilherme Silva", "Ivanete Bertasol", "Rosângela Pallu", "Marcelo Pereira"]
                            if g not in gerentes_reservados
                        ])))
                        if "Outro Gerente..." not in gerentes_operacao_puros:
                            gerentes_operacao_puros.append("Outro Gerente...")

                        gerente_indicado_sel = st.selectbox("Gerente Responsável:", options=gerentes_operacao_puros, index=0)
                        novo_gerente_texto = st.text_input("Escreva o Nome do Novo Gerente:", placeholder="Nome completo...") if gerente_indicado_sel == "Outro Gerente..." else ""

                with col_d:
                    num_chamado = st.text_input("Nº do Chamado / Ticket (Obrigatório para Suporte):", placeholder="Ex: INC-98472")

                obs_detalhe = st.text_area("Observações e Detalhes da Solução (Aplicado a Todos):", value="", placeholder="Observações da tratativa...")
                btn_salvar_g = st.form_submit_button(f"💾 Salvar Tratativa em Lote ({len(bilhetes_selecionados)} bilhetes)")

                if btn_salvar_g:
                    gerente_final = novo_gerente_texto.strip() if nova_area == "Operação" and gerente_indicado_sel == "Outro Gerente..." else gerente_indicado_sel

                    if nova_area == "Suporte backoffice" and not num_chamado.strip():
                        st.error("⚠️ Para reatribuir ao **Suporte backoffice**, é OBRIGATÓRIO informar o Número do Chamado!")
                    elif nova_area == "Operação" and not gerente_final:
                        st.error("⚠️ Por favor, informe o Nome do Gerente!")
                    else:
                        texto_obs_final = f"[Chamado: {num_chamado}] {obs_detalhe}" if num_chamado else obs_detalhe

                        novos_logs_list = []
                        mascara_selecionados = df_master["Bilhetes"].apply(clean_str_strict).isin(bilhetes_selecionados)
                        idxs_para_atualizar = df_master[mascara_selecionados].index

                        for b_item in bilhetes_selecionados:
                            r_item = df_master[df_master["Bilhetes"].apply(clean_str_strict) == b_item].iloc[0]
                            novos_logs_list.append({
                                "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Bilhete": b_item,
                                "Usuario_Acao": usuario_log_formatado,
                                "Status_Anterior": r_item.get("Status_Geral", "Pendente"),
                                "Novo_Status": novo_status,
                                "Area_Anterior": r_item.get(COL_GERENTE, "Operação"),
                                "Nova_Area": gerente_final,
                                "Observacao": obs_detalhe,
                                "Tipo_Interacao": "Tratativa Geral em Lote",
                            })

                        novo_log_df = pd.DataFrame(novos_logs_list)

                        if novo_status == "Já Lançado no ERP":
                            rows_upd = df_master.loc[idxs_para_atualizar].copy()
                            rows_upd["Status_Geral"] = "Já Lançado no ERP"
                            rows_upd[COL_GERENTE] = gerente_final
                            rows_upd["Obs. Operação"] = texto_obs_final
                            rows_upd["Status_Divergencia"] = "Valores Corretos"
                            rows_upd["Tipo_Inconsistencia"] = rows_upd.apply(categorizar_tipo_inconsistencia, axis=1)

                            df_master = df_master.drop(idxs_para_atualizar)
                            df_sem_div = pd.concat([df_sem_div, rows_upd], ignore_index=True)
                            msg_res = f"🎉 {len(bilhetes_selecionados)} bilhete(s) movidos para 'Sem Divergência'!"
                        else:
                            df_master.loc[idxs_para_atualizar, "Status_Geral"] = novo_status
                            df_master.loc[idxs_para_atualizar, COL_GERENTE] = gerente_final
                            df_master.loc[idxs_para_atualizar, "Obs. Operação"] = texto_obs_final
                            msg_res = f"✅ {len(bilhetes_selecionados)} bilhete(s) atualizados com sucesso!"

                        mascara_back_upd = df_master[COL_GERENTE].astype(str).str.lower().str.contains("suporte backoffice|suporte benner|katia martins", na=False)
                        df_back_atualizado = df_master[mascara_back_upd].copy()
                        df_log_updated = pd.concat([df_log_master, novo_log_df], ignore_index=True)

                        sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, df_back_atualizado, df_log_updated)
                        if sucesso_save:
                            st.session_state["msg_sucesso"] = f"{msg_res} (Atribuídos a: {gerente_final})"
                            st.rerun()
                        else:
                            st.error(err_msg)

    st.markdown("---")
    col_t1, col_e1 = st.columns([3, 1])
    with col_t1:
        st.markdown(f"### 📊 Lista Completa dos Casos Filtrados ({len(df_master_filtrado)} registros)")
    with col_e1:
        st.download_button(
            label="📥 Exportar Excel Formatado",
            data=gerar_excel_formatado(df_master_filtrado, "Base_Geral"),
            file_name=f"Relatorio_Base_Geral_Filtrado_{dt_str_export}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_geral",
        )
    st.dataframe(reordenar_colunas_visivel(df_master_filtrado), hide_index=True)

# ABA 2: DIVERGÊNCIA OPERAÇÃO - LOTE
elif aba_atual == "⚠️ Divergência Operação (CIAs/BSP HOT)":
    st.subheader("⚠️ Base 98 - Divergência de Operação / CIAs Aéreas / BSP HOT (IATA)")
    if len(df_div_op_filtrado) == 0:
        st.warning("Nenhuma divergência de operação encontrada para os filtros selecionados.")
    else:
        lista_busca_div = df_div_op_filtrado["Bilhetes"].apply(clean_str_strict).tolist()
        bilhetes_div_sel = st.multiselect(
            "Selecione UM ou MAIS Bilhetes para Tratativa em Lote:",
            options=lista_busca_div,
            default=[lista_busca_div[0]] if lista_busca_div else [],
            key="sb_div_op_multi",
        )

        if bilhetes_div_sel:
            st.info(f"⚡ **{len(bilhetes_div_sel)} divergência(s) selecionada(s)** para resolução.")
            df_previa_div = df_div_op_filtrado[df_div_op_filtrado["Bilhetes"].apply(clean_str_strict).isin(bilhetes_div_sel)]
            
            cols_desejadas_d = ["Ponto de venda", "Área Resp. Operação", "Obs. Operação", "CIA", "📄 Origem / Arquivo", "Bilhetes", "Rloc_Cia", "Status_Divergencia", "Status_Geral"]
            cols_exibir_d = [c for c in cols_desejadas_d if c in df_previa_div.columns]
            
            st.dataframe(df_previa_div[cols_exibir_d], hide_index=True)

            with st.form("form_tratativa_div_op"):
                c_x, c_y, c_z = st.columns(3)
                with c_x:
                    novo_status_d = st.selectbox(
                        "Status da Tratativa:",
                        options=["CIA Aérea Corrigida", "Já Lançado no ERP", "Pendente de Lançamento", "Aguardando TI", "Cancelado / Devolvido"],
                    )
                with c_y:
                    cia_corrigida_d = st.text_input("CIA Aérea Corrigida (opcional):", value="")
                with c_z:
                    num_chamado_d = st.text_input("Nº do Chamado / Ticket (se houver):")

                obs_d = st.text_area("Observações e Justificativas:", value="", placeholder="Justificativa ou correção...")
                btn_salvar_d = st.form_submit_button("💾 Salvar Ação nas Divergências")

                if btn_salvar_d:
                    mascara_div_sel = df_div_op["Bilhetes"].apply(clean_str_strict).isin(bilhetes_div_sel)
                    idxs_d = df_div_op[mascara_div_sel].index

                    texto_obs_d = (
                        (f"[Correção CIA: {cia_corrigida_d}] " if cia_corrigida_d else "")
                        + (f"[Chamado: {num_chamado_d}] " if num_chamado_d else "")
                        + obs_d
                    )

                    novos_logs_d = []
                    for b_div in bilhetes_div_sel:
                        r_d = df_div_op[df_div_op["Bilhetes"].apply(clean_str_strict) == b_div].iloc[0]
                        g_d = r_d.get(COL_GERENTE, "Operação")
                        novos_logs_d.append({
                            "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Bilhete": b_div,
                            "Usuario_Acao": usuario_log_formatado,
                            "Status_Anterior": r_d.get("Status_Geral", "Pendente"),
                            "Novo_Status": novo_status_d,
                            "Area_Anterior": g_d,
                            "Nova_Area": g_d,
                            "Observacao": texto_obs_d,
                            "Tipo_Interacao": "Tratativa Divergência CIA em Lote",
                        })

                    novo_log_d_df = pd.DataFrame(novos_logs_d)

                    if novo_status_d in ["Já Lançado no ERP", "CIA Aérea Corrigida"]:
                        rows_upd_d = df_div_op.loc[idxs_d].copy()
                        rows_upd_d["Status_Geral"] = novo_status_d
                        if cia_corrigida_d:
                            rows_upd_d["CIA"] = cia_corrigida_d
                        rows_upd_d["Obs. Operação"] = texto_obs_d
                        rows_upd_d["Status_Divergencia"] = "Valores Corretos"
                        rows_upd_d["Tipo_Inconsistencia"] = rows_upd_d.apply(categorizar_tipo_inconsistencia, axis=1)

                        df_div_op = df_div_op.drop(idxs_d)
                        df_sem_div = pd.concat([df_sem_div, rows_upd_d], ignore_index=True)
                        msg_res_d = f"🎉 {len(bilhetes_div_sel)} divergência(s) resolvidas e transferidas para 'Sem Divergência'!"
                    else:
                        df_div_op.loc[idxs_d, "Status_Geral"] = novo_status_d
                        if cia_corrigida_d:
                            df_div_op.loc[idxs_d, "CIA"] = cia_corrigida_d
                        df_div_op.loc[idxs_d, "Obs. Operação"] = texto_obs_d
                        msg_res_d = f"✅ {len(bilhetes_div_sel)} divergência(s) atualizadas com sucesso!"

                    df_log_updated = pd.concat([df_log_master, novo_log_d_df], ignore_index=True)

                    sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, None, df_log_updated)
                    if sucesso_save:
                        st.session_state["msg_sucesso"] = msg_res_d
                        st.rerun()
                    else:
                        st.error(err_msg)

    st.markdown("---")
    col_t2, col_e2 = st.columns([3, 1])
    with col_t2:
        st.markdown(f"### 📊 Lista Completa das Divergências ({len(df_div_op_filtrado)} registros)")
    with col_e2:
        st.download_button(
            label="📥 Exportar Excel Formatado",
            data=gerar_excel_formatado(df_div_op_filtrado, "Divergencias_Operacao"),
            file_name=f"Relatorio_Divergencias_Operacao_Filtrado_{dt_str_export}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_div",
        )
    st.dataframe(reordenar_colunas_visivel(df_div_op_filtrado), hide_index=True)

# ABA 3: SEM DIVERGÊNCIA
elif aba_atual == "✅ Sem Divergência (Conciliação)":
    st.subheader("✅ Base 98 - Bilhetes Prontos para Conciliação Operacional")
    if st.session_state.get("perfil_atual") == "Compliance":
        with st.expander("🛡️ Módulo do Compliance — Retirar Linha e Devolver para Tratativa Operacional", expanded=False):
            if df_sem_div_filtrado.empty or "Bilhetes" not in df_sem_div_filtrado.columns:
                st.info("Nenhum bilhete disponível para devolução nesta visualização.")
            else:
                lista_bilhetes_sd = df_sem_div_filtrado["Bilhetes"].apply(clean_str_strict).tolist()
                bilhete_devolver_sel = st.selectbox("Selecione o Bilhete / LOC para Devolução:", options=lista_bilhetes_sd, key="sb_devolucao_compliance")
                row_sd = df_sem_div_filtrado[df_sem_div_filtrado["Bilhetes"].apply(clean_str_strict) == bilhete_devolver_sel].iloc[0]

                with st.form("form_devolucao_compliance"):
                    st.write(f"**Bilhete:** {row_sd.get('Bilhetes', '-')} | **Rloc CIA:** {row_sd.get('Rloc_Cia', '-')} | **Gerente Atual:** {row_sd.get(COL_GERENTE, '-')} | **CIA:** {row_sd.get('CIA', '-')}")
                    area_devolucao_comp = st.selectbox("Devolver para Área:", options=["Operação", "Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private"])
                    motivo_devolucao = st.text_area("Motivo/Justificativa da Devolução (Obrigatório):", placeholder="Motivo técnico da recusa...")
                    btn_devolver_sd = st.form_submit_button("🔄 Confirmar Devolução para Operação")

                    if btn_devolver_sd:
                        if not motivo_devolucao.strip():
                            st.error("⚠️ É obrigatório preencher a justificativa da devolução!")
                        else:
                            mascara_sd_bilhete = df_sem_div["Bilhetes"].apply(clean_str_strict) == str(bilhete_devolver_sel).strip()
                            idx_sd = df_sem_div[mascara_sd_bilhete].index
                            texto_obs_comp = f"[Devolvido pelo Compliance]: {motivo_devolucao}"
                            area_final_comp = "Suporte Backoffice" if area_devolucao_comp == "Suporte backoffice" else area_devolucao_comp

                            rows_para_devolver = df_sem_div.loc[idx_sd].copy()
                            rows_para_devolver["Status_Geral"] = "Pendente de Lançamento"
                            rows_para_devolver[COL_GERENTE] = area_final_comp
                            rows_para_devolver["Obs. Operação"] = texto_obs_comp
                            rows_para_devolver["Tipo_Inconsistencia"] = rows_para_devolver.apply(categorizar_tipo_inconsistencia, axis=1)

                            df_sem_div = df_sem_div.drop(idx_sd)
                            df_master = pd.concat([df_master, rows_para_devolver], ignore_index=True)

                            novo_log_comp = pd.DataFrame([{
                                "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Bilhete": bilhete_devolver_sel,
                                "Usuario_Acao": usuario_log_formatado,
                                "Status_Anterior": "Já Lançado no ERP (Sem Divergência)",
                                "Novo_Status": "Pendente de Lançamento (Devolvido pelo Compliance)",
                                "Area_Anterior": row_sd.get(COL_GERENTE, "-"),
                                "Nova_Area": area_final_comp,
                                "Observacao": motivo_devolucao,
                                "Tipo_Interacao": "Devolução do Compliance",
                            }])

                            df_log_updated = pd.concat([df_log_master, novo_log_comp], ignore_index=True)

                            sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, None, df_log_updated)
                            if sucesso_save:
                                st.session_state["msg_sucesso"] = f"🔄 Bilhete {bilhete_devolver_sel} devolvido para {area_final_comp}!"
                                st.rerun()
                            else:
                                st.error(err_msg)

    col_t3, col_e3 = st.columns([3, 1])
    with col_t3:
        st.markdown(f"### 📊 Lista Completa dos Bilhetes Sem Divergência ({len(df_sem_div_filtrado)} registros)")
    with col_e3:
        st.download_button(
            label="📥 Exportar Excel Formatado",
            data=gerar_excel_formatado(df_sem_div_filtrado, "Sem_Divergencia"),
            file_name=f"Relatorio_Sem_Divergencia_Filtrado_{dt_str_export}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_sem_div",
        )
    st.dataframe(reordenar_colunas_visivel(df_sem_div_filtrado), hide_index=True)

# ABA 4: SUPORTE BACKOFFICE - LOTE
elif aba_atual == "🎧 Suporte Backoffice":
    st.subheader("🎧 Base 99 - Chamados Atribuídos ao Suporte Backoffice / Suporte Benner")
    if len(df_backoffice_filtrado) == 0:
        st.info("Nenhum chamado pendente no Suporte Backoffice no momento.")
    else:
        lista_back_bilhetes = df_backoffice_filtrado["Bilhetes"].apply(clean_str_strict).tolist()
        bilhetes_back_sel = st.multiselect(
            "Selecione UM ou MAIS Chamados / Bilhetes para Tratativa do Suporte:",
            options=lista_back_bilhetes,
            default=[lista_back_bilhetes[0]] if lista_back_bilhetes else [],
            key="sb_backoffice_multi",
        )

        if bilhetes_back_sel:
            st.info(f"⚡ **{len(bilhetes_back_sel)} chamado(s) selecionado(s)** para solução.")

            df_previa_bk = df_backoffice_filtrado[df_backoffice_filtrado["Bilhetes"].apply(clean_str_strict).isin(bilhetes_back_sel)]
            
            cols_desejadas_bk = ["Ponto de venda", "Área Resp. Operação", "Obs. Operação", "CIA", "📄 Origem / Arquivo", "Bilhetes", "Rloc_Cia", "Status_Geral"]
            cols_exibir_bk = [c for c in cols_desejadas_bk if c in df_previa_bk.columns]
            
            st.dataframe(df_previa_bk[cols_exibir_bk], hide_index=True)

            with st.form("form_solucao_backoffice"):
                acao_back = st.selectbox(
                    "Ação do Suporte / Auditoria:",
                    options=["Informar que está Correto (Mover para Sem Divergência)", "Devolver para Tratativa Operacional"],
                )

                area_devolucao_bk = (
                    st.selectbox("Área Operacional de Destino:", options=["Operação", "Central de Eventos", "Concierge/Lazer", "Unique", "Private"])
                    if acao_back.startswith("Devolver")
                    else "Suporte Backoffice"
                )

                obs_back = st.text_area("Parecer do Suporte Backoffice / Auditoria (Aplicado a Todos):", value="", placeholder="Parecer técnico...")
                btn_salvar_back = st.form_submit_button(f"💾 Salvar Resolução do Suporte ({len(bilhetes_back_sel)} chamados)")

                if btn_salvar_back:
                    novos_logs_bk = []
                    for b_bk in bilhetes_back_sel:
                        r_bk = df_backoffice_filtrado[df_backoffice_filtrado["Bilhetes"].apply(clean_str_strict) == b_bk].iloc[0]

                        mascara_bk_bilhete = (
                            df_master["Bilhetes"].apply(clean_str_strict) == str(b_bk).strip()
                            if "Bilhetes" in df_master.columns
                            else []
                        )
                        idx_m_bk = df_master[mascara_bk_bilhete].index if len(mascara_bk_bilhete) > 0 else []

                        if acao_back.startswith("Informar"):
                            row_upd_bk = r_bk.copy()
                            row_upd_bk["Status_Geral"] = "Já Lançado no ERP"
                            row_upd_bk[COL_GERENTE] = area_devolucao_bk
                            row_upd_bk["Status_Divergencia"] = "Valores Corretos"
                            row_upd_bk["Tipo_Inconsistencia"] = categorizar_tipo_inconsistencia(row_upd_bk)
                            texto_obs_bk = f"[Correto pelo Suporte]: {obs_back}"
                            row_upd_bk["Obs. Operação"] = texto_obs_bk

                            if len(idx_m_bk) > 0:
                                df_master = df_master.drop(idx_m_bk)
                            df_sem_div = pd.concat([df_sem_div, pd.DataFrame([row_upd_bk])], ignore_index=True)

                            novo_status_log, area_destino_log = "Já Lançado no ERP (Sem Divergência)", area_devolucao_bk
                        else:
                            texto_obs_bk = f"[Devolvido pelo Suporte]: {obs_back}"
                            if len(idx_m_bk) > 0:
                                df_master.loc[idx_m_bk, "Status_Geral"] = "Pendente de Lançamento"
                                df_master.loc[idx_m_bk, COL_GERENTE] = area_devolucao_bk
                                df_master.loc[idx_m_bk, "Obs. Operação"] = texto_obs_bk

                            novo_status_log, area_destino_log = f"Devolvido para {area_devolucao_bk}", area_devolucao_bk

                        novos_logs_bk.append({
                            "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Bilhete": b_bk,
                            "Usuario_Acao": usuario_log_formatado,
                            "Status_Anterior": r_bk.get("Status_Geral", "Pendente"),
                            "Novo_Status": novo_status_log,
                            "Area_Anterior": "Suporte Backoffice",
                            "Nova_Area": area_destino_log,
                            "Observacao": obs_back,
                            "Tipo_Interacao": "Tratativa Suporte Backoffice em Lote",
                        })

                    mascara_back_upd = df_master[COL_GERENTE].astype(str).str.lower().str.contains("suporte backoffice|suporte benner|katia martins", na=False)
                    df_backoffice_atualizado = df_master[mascara_back_upd].copy()
                    df_log_updated = pd.concat([df_log_master, pd.DataFrame(novos_logs_bk)], ignore_index=True)

                    sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, df_backoffice_atualizado, df_log_updated)
                    if sucesso_save:
                        st.session_state["msg_sucesso"] = f"🎉 {len(bilhetes_back_sel)} chamado(s) processados com sucesso!"
                        st.rerun()
                    else:
                        st.error(err_msg)

    st.markdown("---")
    col_t4, col_e4 = st.columns([3, 1])
    with col_t4:
        st.markdown(f"### 📊 Lista Completa dos Chamados no Suporte ({len(df_backoffice_filtrado)} registros)")
    with col_e4:
        st.download_button(
            label="📥 Exportar Excel Formatado",
            data=gerar_excel_formatado(df_backoffice_filtrado, "Suporte_Backoffice"),
            file_name=f"Relatorio_Suporte_Backoffice_Filtrado_{dt_str_export}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_back",
        )
    st.dataframe(reordenar_colunas_visivel(df_backoffice_filtrado), hide_index=True)

# ABA 5: RÉPLICA DA AUDITORIA
elif aba_atual == "⚖️ Réplica da Auditoria":
    st.subheader("⚖️ Módulo de Contestação e Réplica da Auditoria")
    bilhetes_com_tratativa = df_acao_filtrado[df_acao_filtrado["Status_Geral"] != "Pendente de Lançamento"]
    if len(bilhetes_com_tratativa) > 0:
        bilhete_rep = st.selectbox("Selecione o Bilhete para Réplica:", options=bilhetes_com_tratativa["Bilhetes_Str"].tolist())
        row_rep = df_acao_filtrado[df_acao_filtrado["Bilhetes_Str"] == bilhete_rep].iloc[0]

        with st.form("form_replica_auditoria"):
            status_auditoria = st.selectbox("Decisão da Auditoria:", options=["Contestado / Recusado", "Aprovado / Conciliado"])
            area_devolucao = st.selectbox("Devolver para Área / Gerente:", options=["Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private", "Operação"])
            motivo_replica = st.text_area("Justificativa da Auditoria:", value="", placeholder="Justificativa...")
            btn_replica = st.form_submit_button("🚨 Enviar Apontamento")

            if btn_replica:
                mascara_rep_bilhete = df_master["Bilhetes"].apply(clean_str_strict) == str(bilhete_rep).strip() if "Bilhetes" in df_master.columns else []
                idx_m_rep = df_master[mascara_rep_bilhete].index if len(mascara_rep_bilhete) > 0 else []
                texto_obs_rep = f"[Contestado pela Auditoria]: {motivo_replica}"
                area_rep_final = "Suporte Backoffice" if area_devolucao == "Suporte backoffice" else area_devolucao

                novo_log_rep = pd.DataFrame([{
                    "Data_Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Bilhete": bilhete_rep,
                    "Usuario_Acao": usuario_log_formatado,
                    "Status_Anterior": row_rep.get("Status_Geral", "-"),
                    "Novo_Status": f"Contestado ({area_rep_final})",
                    "Area_Anterior": row_rep.get(COL_GERENTE, "-"),
                    "Nova_Area": area_rep_final,
                    "Observacao": motivo_replica,
                    "Tipo_Interacao": "Réplica da Auditoria",
                }])

                if len(idx_m_rep) > 0:
                    df_master.loc[idx_m_rep, "Status_Geral"] = f"Contestado ({area_rep_final})"
                    df_master.loc[idx_m_rep, COL_GERENTE] = area_rep_final
                    df_master.loc[idx_m_rep, "Obs. Operação"] = texto_obs_rep

                df_log_updated = pd.concat([df_log_master, novo_log_rep], ignore_index=True)

                sucesso_save, err_msg = salvar_base_consolidada(df_master, df_div_op, df_sem_div, None, df_log_updated)
                if sucesso_save:
                    st.session_state["msg_sucesso"] = f"✅ Contestação registrada por {usuario_log_formatado}!"
                    st.rerun()
                else:
                    st.error(err_msg)

    st.markdown("---")
    st.markdown(f"### 📊 Lista Completa dos Casos em Réplica ({len(bilhetes_com_tratativa)} registros)")
    st.dataframe(reordenar_colunas_visivel(bilhetes_com_tratativa), hide_index=True)

# ABA 6: TRILHA DE AUDITORIA (COMPLIANCE)
elif aba_atual == "📜 Trilha de Auditoria" and st.session_state["perfil_atual"] == "Compliance":
    st.subheader("📜 Histórico Completo de Alterações e Trilha de Auditoria")
    with st.expander("🗑️ Módulo Master de Gerenciamento e Exclusão de Logs"):
        if not df_log_master.empty:
            indices_log = df_log_master.index.tolist()
            log_opcao = st.selectbox(
                "Selecione o registro de log que deseja EXCLUIR:",
                options=indices_log,
                format_func=lambda i: f"Linha {i} | Data: {df_log_master.loc[i, 'Data_Hora']} | Bilhete: {df_log_master.loc[i, 'Bilhete']} | Usuário: {df_log_master.loc[i, 'Usuario_Acao']}",
            )
            if st.button("❌ Excluir Registro de Log Selecionado"):
                df_log_master = df_log_master.drop(log_opcao).reset_index(drop=True)
                sucesso_save, err_msg = salvar_base_consolidada(None, None, None, None, df_log_master)
                if sucesso_save:
                    st.session_state["msg_sucesso"] = "🗑️ Registro de log excluído com sucesso!"
                    st.rerun()
                else:
                    st.error(err_msg)
        else:
            st.info("Nenhum registro de log para exclusão.")

    st.dataframe(df_log_master, hide_index=False)

# ABA 7: GESTÃO DE ACESSOS (COMPLIANCE)
elif aba_atual == "⚙️ Gestão de Acessos" and st.session_state["perfil_atual"] == "Compliance":
    st.subheader("⚙️ Central de Aprovações de Acesso e Governança")
    usuarios_atuais = carregar_usuarios()
    pendentes = {k: v for k, v in usuarios_atuais.items() if v.get("status") == "PENDENTE"}

    if not pendentes:
        st.info("🎉 Nenhuma solicitação de acesso pendente no momento.")
    else:
        for u_id, u_info in pendentes.items():
            st.markdown(
                f"""
                <div style="background-color: #ffffff; border-left: 5px solid #002060; padding: 15px; border-radius: 8px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);">
                    <b>Nome:</b> {u_info['nome']} | <b>Usuário:</b> {u_id} | <b>Perfil Solicitado:</b> {u_info['perfil']}<br>
                    <small style="color: #6c757d;">Data da Solicitação: {u_info.get('data_solicitacao', '-')}</small>
                </div>
                """,
                unsafe_allow_html=True,
            )

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

# ABA VISÃO GERAL
elif aba_atual == "📋 Visão Geral da Base Total":
    col_tv, col_ev = st.columns([3, 1])
    with col_tv:
        st.subheader("📋 Visão Geral da Base Total de Divergências")
    with col_ev:
        st.download_button(
            label="📥 Exportar Excel Formatado",
            data=gerar_excel_formatado(df_acao_filtrado, "Base_Total_Filtrada"),
            file_name=f"Relatorio_Base_Total_Filtrado_{dt_str_export}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_exp_total",
        )
    st.dataframe(reordenar_colunas_visivel(df_acao_filtrado), hide_index=True)