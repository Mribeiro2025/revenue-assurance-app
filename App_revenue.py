import os
import datetime
import pandas as pd
import openpyxl
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Revenue Assurance | Portal de Tratativas",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .stApp { background-color: #f4f6f9; }
        .header-box {
            background: linear-gradient(135deg, #002060 0%, #004080 100%);
            padding: 24px;
            border-radius: 12px;
            color: white;
            margin-bottom: 25px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        .header-box h1 { color: #ffffff !important; margin: 0; font-size: 28px; font-weight: 700; }
        .header-box p { color: #d0e0ff !important; margin-top: 5px; font-size: 14px; }
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 16px;
            border-left: 5px solid #002060;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
    </style>
""", unsafe_allow_html=True)

ARQUIVO_DASHBOARD = "Dashboard_Revenue_Assurance_Consolidado.xlsx"

def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace("[", "").replace("]", "").replace("R$", "").strip()
    if not s or s == "-": return 0.0
    if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
    elif "," in s: s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

@st.cache_data(ttl=5)
def carregar_dados():
    if not os.path.exists(ARQUIVO_DASHBOARD):
        return None
    xls = pd.ExcelFile(ARQUIVO_DASHBOARD)
    df_furo = pd.read_excel(xls, sheet_name="99_Base_Divergencias_Geral")
    return df_furo

df_master = carregar_dados()

if df_master is None:
    st.error(f"⚠️ Base **{ARQUIVO_DASHBOARD}** não encontrada! Execute primeiro a auditoria Python.")
    st.stop()

st.markdown("""
    <div class="header-box">
        <h1>✈️ Portal Executivo de Revenue Assurance</h1>
        <p>Sistema Inteligente de Tratativa, Reatribuição de Chamados e Conciliação Operacional</p>
    </div>
""", unsafe_allow_html=True)

st.sidebar.title("🔍 Filtros Operacionais")
setores = list(df_master["Setor"].dropna().unique())
setor_sel = st.sidebar.multiselect("Setor Responsável:", options=setores, default=setores)

tipos_emissao = list(df_master["Tipo_Emissao_Lemon"].dropna().unique())
emissao_sel = st.sidebar.multiselect("Tipo de Emissão (OBT):", options=tipos_emissao, default=tipos_emissao)

cias = list(df_master["CIA"].dropna().unique())
cia_sel = st.sidebar.multiselect("Companhia Aérea:", options=cias, default=cias)

df_filtrado = df_master[
    (df_master["Setor"].isin(setor_sel)) &
    (df_master["Tipo_Emissao_Lemon"].isin(emissao_sel)) &
    (df_master["CIA"].isin(cia_sel))
].copy()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Bilhetes Pendentes", f"{len(df_filtrado):,}")
c2.metric("Tarifa Pendente (R$)", f"R$ {df_filtrado[\x27A credito\x27].sum():,.2f}")
c3.metric("Taxas Pendentes (R$)", f"R$ {df_filtrado[\x27Taxa\x27].sum():,.2f}")
c4.metric("Receita em Risco (R$)", f"R$ {df_filtrado[\x27Incentivo\x27].sum():,.2f}")

st.markdown("---")

aba_form, aba_graficos, aba_tabela = st.tabs([
    "🎯 Form de Tratativa Individual", 
    "📊 Gráficos & KPIs Dinâmicos", 
    "📋 Visão Geral da Base"
])

with aba_form:
    st.subheader("📝 Módulo de Resolução e Detalhamento por Bilhete")
    if len(df_filtrado) == 0:
        st.warning("Nenhum bilhete encontrado com os filtros atuais.")
    else:
        lista_bilhetes = df_filtrado["Bilhetes"].astype(str).tolist()
        bilhete_selecionado = st.selectbox("Pesquise ou Selecione o Bilhete:", options=lista_bilhetes)
        row = df_filtrado[df_filtrado["Bilhetes"].astype(str) == bilhete_selecionado].iloc[0]
        
        st.markdown(f"""
            <div style="background-color: #fff3cd; border-left: 6px solid #ffc107; padding: 15px; border-radius: 8px; margin: 15px 0;">
                <h4 style="margin: 0; color: #856404;">⚠️ Diagnóstico da Auditoria: {row[\x27Bilhetes\x27]}</h4>
                <p style="margin: 5px 0 0 0; color: #856404;">
                    <b>Status Geral:</b> {row.get(\x27Status_Geral\x27, \x27Pendente\x27)} | 
                    <b>CIA:</b> {row.get(\x27CIA\x27, \x27-\x27)} | 
                    <b>Data Emissão:</b> {row.get(\x27Data Emissão\x27, \x27-\x27)}
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        f1, f2, f3, f4 = st.columns(4)
        f1.text_input("Tarifa A Crédito", f"R$ {clean_num(row.get(\x27A credito\x27)):,.2f}", disabled=True)
        f2.text_input("Taxas", f"R$ {clean_num(row.get(\x27Taxa\x27)):,.2f}", disabled=True)
        f3.text_input("Receita/Incentivo", f"R$ {clean_num(row.get(\x27Incentivo\x27)):,.2f}", disabled=True)
        f4.text_input("Consultor (OBT Lemon)", str(row.get("Consultor_Lemon", "-")), disabled=True)
        
        st.markdown("### 🛠️ Preenchimento de Tratativa da Operação")
        with st.form("form_tratativa"):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                novo_status = st.selectbox("Status da Tratativa:", options=["Pendente de Lançamento", "Já Lançado no ERP", "Aguardando TI", "Cancelado / Devolvido"])
            with col_b:
                areas_opcoes = ["Suporte backoffice", "Central de Eventos", "Concierge/Lazer", "Unique", "Private", "Operação"]
                area_atual = row.get("Área Resp. Operação", "Operação")
                idx_area = areas_opcoes.index(area_atual) if area_atual in areas_opcoes else 0
                nova_area = st.selectbox("Área Responsável:", options=areas_opcoes, index=idx_area)
            with col_c:
                num_chamado = st.text_input("Nº do Chamado / Ticket (TI):", placeholder="Ex: INC-98472")
            
            obs_detalhe = st.text_area("Observações e Detalhes:", value=str(row.get("Obs. Operação", "")))
            btn_salvar = st.form_submit_button("💾 Salvar Tratativa deste Bilhete")
            
            if btn_salvar:
                idx_m = df_master[df_master["Bilhetes"].astype(str) == bilhete_selecionado].index
                df_master.loc[idx_m, "Status_Geral"] = novo_status
                df_master.loc[idx_m, "Área Resp. Operação"] = nova_area
                df_master.loc[idx_m, "Obs. Operação"] = f"[Chamado: {num_chamado}] - {obs_detalhe}" if num_chamado else obs_detalhe
                df_master.loc[idx_m, "Setor"] = nova_area
                
                with pd.ExcelWriter(ARQUIVO_DASHBOARD, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_master.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
                st.success(f"✅ Tratativa do bilhete {bilhete_selecionado} atualizada com sucesso!")
                st.cache_data.clear()

with aba_graficos:
    st.subheader("📊 Análise Dinâmica das Pendências")
    g1, g2 = st.columns(2)
    with g1:
        fig_setor = px.bar(
            df_filtrado.groupby("Setor")["A credito"].sum().reset_index(),
            x="Setor", y="A credito", title="Tarifa Pendente por Setor (R$)", color="Setor"
        )
        st.plotly_chart(fig_setor, use_container_width=True)
    with g2:
        fig_obt = px.pie(df_filtrado, names="Tipo_Emissao_Lemon", values="A credito", title="Distribuição por Tipo de Emissão (OBT)", hole=0.4)
        st.plotly_chart(fig_obt, use_container_width=True)

with aba_tabela:
    st.subheader("📋 Tabela Geral de Auditoria")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

