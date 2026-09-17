import csv
import datetime
import glob
import os
import re
import shutil
import sqlite3
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from tqdm import tqdm

DIR_ATUAL = (
    os.path.dirname(os.path.abspath(__file__))
    if "__file__" in globals()
    else os.getcwd()
)
os.chdir(DIR_ATUAL)

print("=" * 75)
print(" MOTOR DE AUDITORIA FP&A ADVANCED - REVENUE ASSURANCE ".center(75, "="))
print("=" * 75)

PASTA_ENVIO = "Envio"
PASTA_INPUTS_LOCAL = r"C:\Users\mribeiro1\MARINGA TURISMO\Maringá Turismo - PLANEJAMENTO ESTRATEGICO (1)\01-Planejamento Estratégico\Auditoria de Bilhetes\inputs"

MAPA_CIAS = {
    "JJ": "LATAM", "LA": "LATAM", "PZ": "LATAM", "4C": "LATAM", "XL": "LATAM", "4M": "LATAM",
    "G3": "GOL", "AD": "AZUL", "TP": "TAP", "CM": "COPA", "AA": "AMERICAN", "UA": "UNITED",
    "DL": "DELTA", "AF": "AIR FRANCE", "KL": "KLM", "AR": "AEROLINEAS", "IB": "IBERIA",
    "UX": "EUROPA", "BA": "BRITISH", "AV": "AVIANCA", "O6": "AVIANCA", "AC": "CANADA",
    "AM": "AEROMEXICO", "QR": "QATAR", "EK": "EMIRATES", "TK": "TURKISH", "LH": "LUFTHANSA",
    "LX": "SWISS", "ET": "ETHIOPIAN", "AT": "MAROC", "SA": "SOUTH AFRICAN", "OB": "BOLIVIANA",
    "PY": "PARAGUAY", "HR": "HAHN", "AZ": "ITA AIRWAYS", "JA": "JETSMART",
}


# ==============================================================================
# CONEXÃO E AUXILIARES DO SUPABASE / BANCO DE DADOS NUVEM
# ==============================================================================
def obter_engine_supabase():
    """Obtém a conexão com o Supabase com suporte multiplataforma a versões do Python."""
    try:
        from sqlalchemy import create_engine
        
        # Suporte para Python < 3.11 e >= 3.11
        try:
            import tomllib
        except ImportError:
            try:
                import tomli as tomllib
            except ImportError:
                import toml as tomllib

        secrets_path = os.path.join(DIR_ATUAL, ".streamlit", "secrets.toml")
        if os.path.exists(secrets_path):
            with open(secrets_path, "rb") as f:
                secrets = tomllib.load(f)
                db_url = secrets.get("postgres", {}).get("url")
                if db_url:
                    return create_engine(db_url, pool_pre_ping=True)
    except Exception as e:
        pass
    return None


def auto_converter_benner_parquet():
    """Garante conversão e atualização automática do Acumulado.xlsx -> Parquet."""
    arq_xlsx = "Acumulado.xlsx"
    arq_parquet = "Acumulado.parquet"

    if not os.path.exists(arq_xlsx):
        return

    precisa_converter = False
    if not os.path.exists(arq_parquet):
        precisa_converter = True
    else:
        mtime_xlsx = os.path.getmtime(arq_xlsx)
        mtime_parquet = os.path.getmtime(arq_parquet)
        if mtime_xlsx > mtime_parquet:
            precisa_converter = True

    if precisa_converter:
        print("\n⏳ [AUTO-CONVERTER] Atualizando 'Acumulado.parquet' a partir de 'Acumulado.xlsx'...")
        try:
            df_benner = pd.read_excel(arq_xlsx, sheet_name="Planilha1")
            df_benner.to_parquet(arq_parquet, engine="pyarrow", compression="snappy")
            print("⚡ [AUTO-CONVERTER] Base Parquet sincronizada com sucesso!")
        except Exception as e:
            print(f"⚠️ Erro ao converter Parquet automaticamente: {e}")


def clean_num(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).replace("[", "").replace("]", "").replace("R$", "").strip()
    if not s or s == "-":
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def clean_str_strict(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    if s.lower() in ["nan", "none", "null", "<na>"]:
        return ""
    return s


def clean_iata(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).replace("-", "").replace(" ", "").strip()


def safe_get_col(row, idx, default=""):
    if idx < len(row):
        return clean_str_strict(row.iloc[idx])
    return default


def safe_get_num(row, idx, default=0.0):
    if idx < len(row):
        return clean_num(row.iloc[idx])
    return default


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


def carregar_tratativas_e_logs_anteriores(out_file):
    """
    Busca o histórico de tratativas com a seguinte hierarquia:
    1. Supabase (PostgreSQL NUVEM) - Fonte Única da Verdade.
    2. Memória CSV local (Historico_Tratativas.csv) - Fallback offline.
    3. Consolidado Excel anterior (Dashboard_Revenue_Assurance_Consolidado.xlsx).
    """
    dict_historico = {}
    df_log_antigo = pd.DataFrame()

    def indexar_memoria(b_raw, dados):
        """Indexa os dados do bilhete aplicando chaves flexíveis e ignorando nulos."""
        if not b_raw or str(b_raw).strip().lower() in ["nan", "none", "", "-"]:
            return

        for k in extract_keys(b_raw):
            if k not in dict_historico:
                dict_historico[k] = {}
            for col, val in dados.items():
                val_str = str(val).strip() if pd.notna(val) else ""
                if val_str and val_str not in ["-", "nan", "none", "Sem tratativa na operação"]:
                    dict_historico[k][col] = val

    # 1º Passo: Carrega do Excel consolidado anterior (apenas o Dashboard principal)
    if os.path.exists(out_file):
        try:
            xls = pd.ExcelFile(out_file, engine="openpyxl")
            if "00_Log_Auditoria" in xls.sheet_names:
                df_log_antigo = pd.read_excel(xls, sheet_name="00_Log_Auditoria")

            abas_ordenadas = sorted(
                [a for a in xls.sheet_names if a.startswith("98_") or a.startswith("99_")],
                reverse=True
            )

            for aba in abas_ordenadas:
                df_temp = pd.read_excel(xls, sheet_name=aba)
                if "Bilhetes" in df_temp.columns:
                    for r in df_temp.to_dict("records"):
                        b_key = clean_str_strict(r.get("Bilhetes"))
                        if b_key:
                            indexar_memoria(b_key, {
                                "Status_Geral": r.get("Status_Geral"),
                                "Área Resp. Operação": r.get("Área Resp. Operação"),
                                "Obs. Operação": r.get("Obs. Operação"),
                                "Setor": r.get("Setor"),
                                "Obs_Auditoria_Replica": r.get("Obs_Auditoria_Replica"),
                                "Ponto de venda": r.get("Ponto de venda"),
                                "Código Iata": r.get("Código Iata"),
                                "Data_Modificacao": r.get("Data_Modificacao"),
                                "Usuario_Modificacao": r.get("Usuario_Modificacao"),
                                "Ultima_Alteracao": r.get("Ultima_Alteracao"),
                            })
        except Exception as e:
            print(f"⚠️ Aviso ao carregar histórico do Excel: {e}")

    # 2º Passo: Aplica a memória protegida em CSV local
    file_memoria = "Historico_Tratativas.csv"
    if os.path.exists(file_memoria) and os.path.getsize(file_memoria) > 0:
        try:
            df_mem = pd.read_csv(file_memoria, dtype=str)
            for r in df_mem.to_dict("records"):
                b_key = clean_str_strict(r.get("Bilhetes", ""))
                if b_key:
                    indexar_memoria(b_key, {
                        "Status_Geral": r.get("Status_Geral"),
                        "Área Resp. Operação": r.get("Área Resp. Operação"),
                        "Obs. Operação": r.get("Obs. Operação"),
                        "Setor": r.get("Setor"),
                        "Obs_Auditoria_Replica": r.get("Obs_Auditoria_Replica"),
                        "Ponto de venda": r.get("Ponto de venda"),
                        "Código Iata": r.get("Código Iata"),
                        "Data_Modificacao": r.get("Data_Modificacao"),
                        "Usuario_Modificacao": r.get("Usuario_Modificacao"),
                        "Ultima_Alteracao": r.get("Ultima_Alteracao"),
                    })
            print(f"🛡️ Memória protegida '{file_memoria}' sincronizada com sucesso.")
        except Exception as e:
            print(f"⚠️ Erro ao ler memória CSV: {e}")

    # 3º Passo: RESGATE DO SUPABASE (PRIORIDADE MÁXIMA PARA TRATATIVAS WEB)
    engine_sb = obter_engine_supabase()
    if engine_sb:
        try:
            df_sb = pd.read_sql("SELECT bilhete, status_geral, area_resp, obs_operacao FROM tratativas", engine_sb)
            if not df_sb.empty:
                for r in df_sb.to_dict("records"):
                    b_key = clean_str_strict(r.get("bilhete"))
                    if b_key:
                        indexar_memoria(b_key, {
                            "Status_Geral": r.get("status_geral"),
                            "Área Resp. Operação": r.get("area_resp"),
                            "Obs. Operação": r.get("obs_operacao"),
                        })
                print(f"☁️ Supabase conectado! {len(df_sb)} tratativas resgatadas da nuvem com sucesso.")
        except Exception as e:
            print(f"⚠️ Aviso ao conectar/carregar dados do Supabase: {e}")
    else:
        print("⚠️ Conexão com Supabase não estabelecida. Verifique se a pasta '.streamlit/secrets.toml' existe.")

    return dict_historico, df_log_antigo


def buscar_memoria(dict_historico, bilhete_raw):
    for k in extract_keys(bilhete_raw):
        if k in dict_historico:
            return dict_historico[k]
    return {}


def encontrar_arquivo(nomes_possiveis):
    for nome in nomes_possiveis:
        if os.path.exists(nome):
            return nome
    return None


def salvar_csv_atomico(df, caminho_csv):
    """Garante gravação segura protegida contra travamentos do OneDrive."""
    tmp_path = f"{caminho_csv}.tmp"
    df.to_csv(tmp_path, index=False, encoding="utf-8-sig")
    if os.path.exists(caminho_csv):
        os.remove(caminho_csv)
    os.rename(tmp_path, caminho_csv)


def sincronizar_supabase_fim(df_master):
    """Envia tratativas e histórico de auditoria em LOTE (Batch Executemany) para o Supabase."""
    engine_sb = obter_engine_supabase()
    if not engine_sb or df_master.empty:
        return

    try:
        from sqlalchemy import text
        df_trat = df_master[["Bilhetes", "Status_Geral", "Área Resp. Operação", "Obs. Operação"]].dropna(subset=["Bilhetes"]).drop_duplicates(subset=["Bilhetes"])

        upsert_sql = text("""
            INSERT INTO tratativas (bilhete, status_geral, area_resp, obs_operacao, usuario_modificacao, data_modificacao)
            VALUES (:bilhete, :status_geral, :area_resp, :obs_operacao, 'Motor_VSCode', NOW())
            ON CONFLICT (bilhete) DO UPDATE SET
                status_geral = EXCLUDED.status_geral,
                area_resp = EXCLUDED.area_resp,
                obs_operacao = EXCLUDED.obs_operacao,
                data_modificacao = NOW();
        """)

        # Monta a lista completa para inserção otimizada em lote único
        dados_lote = []
        for _, r in df_trat.iterrows():
            b_val = clean_str_strict(r["Bilhetes"])
            if b_val:
                dados_lote.append({
                    "bilhete": b_val,
                    "status_geral": str(r.get("Status_Geral", "")),
                    "area_resp": str(r.get("Área Resp. Operação", "")),
                    "obs_operacao": str(r.get("Obs. Operação", ""))
                })

        if dados_lote:
            with engine_sb.begin() as conn:
                conn.execute(upsert_sql, dados_lote)
            print(f"☁️ {len(dados_lote)} tratativas sincronizadas em lote no Supabase!")
    except Exception as e:
        print(f"⚠️ Aviso ao sincronizar com o Supabase: {e}")


def executar_auditoria():
    auto_converter_benner_parquet()

    out_file_model = "Dashboard_Revenue_Assurance_Consolidado.xlsx"

    dict_historico, df_log_acumulado = carregar_tratativas_e_logs_anteriores(out_file_model)

    if os.path.exists(out_file_model):
        dt_bkp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            shutil.copy2(out_file_model, f"Backup_Dashboard_{dt_bkp}.xlsx")
            print(f"📦 Backup de segurança gerado: Backup_Dashboard_{dt_bkp}.xlsx")
        except Exception as e:
            print(f"⚠️ Aviso ao criar backup: {e}")

    print("\n[1/5] Carregando Relação IATA...")
    caminho_iata = encontrar_arquivo(["Relacao_Iata_2.xlsx", "Relacao_Iata.xlsx", "Relacao_Iata_2.XLSX"])
    df_iata = pd.read_excel(caminho_iata) if caminho_iata else pd.DataFrame()
    iata_dict = {}

    if not df_iata.empty:
        for idx, r in tqdm(df_iata.iterrows(), total=len(df_iata), desc="Indexando IATAs", unit="reg"):
            k = clean_iata(r.get("Código IATA"))
            if k:
                iata_dict[k] = {
                    "Nome_IATA_Oficial": clean_str_strict(r.get("Nome IATA")),
                    "Gerente_Responsavel": clean_str_strict(r.get("Gerentes")),
                }

    print("\n[2/5] Indexando Extrato OBT Lemontech...")
    caminho_lemon = encontrar_arquivo(["Extrato_Bilhetes_lemontech.xlsx", "Extrato_Bilhetes_lemontech.XLSX"])
    lemon_index = {}

    if caminho_lemon:
        df_lemon = pd.read_excel(caminho_lemon, engine="openpyxl")
        for idx, r in tqdm(df_lemon.iterrows(), total=len(df_lemon), desc="Indexando Lemontech", unit="reg"):
            cc_desc = f"{clean_str_strict(r.get('Centro de Custo'))} - {clean_str_strict(r.get('Descrição Centro de Custo'))}".strip(" -")
            info_lemon = {
                "Lemon_Solicitante": clean_str_strict(r.get("Solicitante")) or "-",
                "Lemon_Passageiro": clean_str_strict(r.get("Passageiro")) or "-",
                "Lemon_Consultor": clean_str_strict(r.get("Consultor")) or "-",
                "Lemon_Emissor_Reserva": clean_str_strict(r.get("Emissor Reserva")) or "-",
                "Lemon_Centro_Custo": cc_desc if cc_desc else "-",
                "Lemon_Forma_Pagto": clean_str_strict(r.get("Forma de Pagamento")) or "-",
                "Lemon_Autorizacao_Cartao": clean_str_strict(r.get("Autorização do Cartão")) or "-",
                "Lemon_OnOff": clean_str_strict(r.get("On|Off")) or "-",
                "Lemon_Source": clean_str_strict(r.get("Source")) or "-",
            }
            for col_k in ["Bilhete", "Nº Pedido", "Solicitação"]:
                for k in extract_keys(r.get(col_k)):
                    if k not in lemon_index:
                        lemon_index[k] = info_lemon

    print("\n[3/5] Indexando ERP Benner...")
    if os.path.exists("Acumulado.parquet"):
        df_benner = pd.read_parquet("Acumulado.parquet")
    else:
        caminho_benner = encontrar_arquivo(["Acumulado.xlsx", "Acumulado.XLSX"])
        df_benner = pd.read_excel(caminho_benner, sheet_name="Planilha1") if caminho_benner else pd.DataFrame()

    benner_index = {}
    key_cols_benner = [
        "Bilhete", "Localizador", "Rloc Cia", "Rloc CIA", "Código Rloc",
        "Pedido", "Fatura/Documento", "Handle Accounting", "Accounting VM", "Apurações Fee"
    ]

    if not df_benner.empty:
        for idx, r in tqdm(df_benner.iterrows(), total=len(df_benner), desc="Indexando Benner", unit="linha"):
            rloc_cia_val = clean_str_strict(r.get("Rloc Cia") or r.get("Rloc CIA") or r.get("Código Rloc") or r.get("Localizador")) or "-"
            info = {
                "Benner_Index": idx,
                "Benner_Situação": clean_str_strict(r.get("Situação")) or "ATIVO",
                "Benner_Fornecedor": clean_str_strict(r.get("Apelido Fornecedor") or r.get("Fornecedor")) or "-",
                "Benner_Localizador": clean_str_strict(r.get("Localizador") or r.get("Código Rloc") or r.get("Rloc Cia")) or "-",
                "Benner_Rloc_Cia": rloc_cia_val,
                "Benner_Bilhete": clean_str_strict(r.get("Bilhete")) or "-",
                "Benner_Tarifa": clean_num(r.get("Tarifa")),
                "Benner_Taxas": clean_num(r.get("Taxas")) + clean_num(r.get("Taxa BR")),
                "Benner_Receita": clean_num(r.get("Comissão")) + clean_num(r.get("Taxa DU")) + clean_num(r.get("Incentivo")) + clean_num(r.get("Fee")) + clean_num(r.get("RAV")),
                "Benner_Cliente": clean_str_strict(r.get("Cliente") or r.get("Apelido Cliente")) or "-",
                "Benner_Emissor": clean_str_strict(r.get("Emissor") or r.get("Agente Criação")) or "-",
                "Benner_Passageiro": clean_str_strict(r.get("Passageiro")) or "-",
                "Benner_Sistema_Reserva": clean_str_strict(r.get("Sistema Reserva")) or "-",
            }
            for col in key_cols_benner:
                for k in extract_keys(r.get(col)):
                    if k not in benner_index:
                        benner_index[k] = info

    print("\n[4/5] Indexando Relatório Sabre...")
    caminho_sabre = encontrar_arquivo(["Relacao_Sabre_3.xlsx", "Relacao_Sabre.xlsx", "Relacao_Sabre_3.XLSX"])
    sabre_index = {}
    if caminho_sabre:
        lines = pd.read_excel(caminho_sabre).iloc[:, 0].dropna().astype(str).tolist()
        sabre_rows = []

        for line in lines:
            if "Total Amount" in line or "Total Commission" in line or "PCC,DATE,AL CODE" in line:
                continue
            parts = list(csv.reader([line]))[0]
            if len(parts) >= 15 and parts[0] != "PCC":
                cleaned = [p.replace('="', "").replace('"', "").strip() for p in parts]
                sabre_rows.append(cleaned[:15])

        df_sabre = pd.DataFrame(
            sabre_rows,
            columns=["PCC", "DATE", "AL_CODE", "TICKET_NUM", "PNR", "NAME", "LAST_NAME", "CUR", "OB", "COM_AMT", "TOTAL", "FOP", "AGT", "TIME", "STATUS"]
        )

        for idx, r in tqdm(df_sabre.iterrows(), total=len(df_sabre), desc="Indexando Sabre", unit="reg"):
            info = {
                "Sabre_PCC": r["PCC"],
                "Sabre_AGT": r["AGT"],
                "Sabre_Passageiro": f"{r['NAME']} {r['LAST_NAME']}".strip(),
            }
            for k in extract_keys(r["TICKET_NUM"]) + extract_keys(r["PNR"]):
                if k not in sabre_index:
                    sabre_index[k] = info

    print("\n[5/5] Auditando e Conciliando Emissões das Cias Aéreas...")
    fontes = [
        (["AZUL.XLSX", "AZUL.xlsx"], "Azul"),
        (["BSP.XLSX", "BSP.xlsx"], "BSP"),
        (["HOT.XLSX", "HOT.xlsx"], "HOT"),
        (["GOL.XLSX", "GOL.xlsx"], "GOL"),
    ]
    registros_conciliados = []

    for lista_nomes, nome_fonte in fontes:
        caminho_arq = encontrar_arquivo(lista_nomes)
        if not caminho_arq:
            continue

        df_raw = pd.read_excel(caminho_arq, header=None)
        curr_iata, curr_ponto_venda = "", ""

        for r_idx in tqdm(range(len(df_raw)), desc=f"Auditando {nome_fonte}", unit="linha"):
            row = df_raw.iloc[r_idx]
            col0 = safe_get_col(row, 0)
            col1 = safe_get_col(row, 1)
            col3 = safe_get_col(row, 3)
            col4 = safe_get_col(row, 4)
            col5 = safe_get_col(row, 5)
            col6 = safe_get_col(row, 6)

            if re.match(r"^\d{2}-\d{5}", col0) or re.match(r"^\d{7,8}$", col0):
                curr_iata = col0
                if col4:
                    curr_ponto_venda = col4
                continue

            if col3 != "" and col3 not in ["BILHETE\\RLOC", "BILHETE", "RLOC", "[]"] and col1 != "CIA":
                if col3.upper() in ["BILHETE\\RLOC", "BILHETE", "RLOC"] or col5.upper() in ["EMISSÃO", "EMISSAO"]:
                    continue

                bilhete_chave = clean_str_strict(col3)

                dt_emissao = ""
                pagto = ""
                if re.match(r"^\d{2}/\d{2}/\d{4}", col4):
                    dt_emissao = col4
                    pagto = col5
                elif re.match(r"^\d{2}/\d{2}/\d{4}", col5):
                    dt_emissao = col5
                    pagto = col6
                else:
                    if col4 and col4.upper() not in ["EMISSÃO", "EMISSAO", "CARTÃO", "FATURA", "OUTRAS"]:
                        dt_emissao = col4
                        pagto = col5
                    elif col5 and col5.upper() not in ["EMISSÃO", "EMISSAO"]:
                        dt_emissao = col5
                        pagto = col6

                if not dt_emissao or dt_emissao.upper() in ["EMISSÃO", "EMISSAO"]:
                    continue

                doc_val = safe_get_col(row, 22)

                hist_data = buscar_memoria(dict_historico, bilhete_chave)

                ponto_venda_final = curr_ponto_venda or hist_data.get("Ponto de venda") or "Não Mapeado"
                iata_final = curr_iata or hist_data.get("Código Iata") or "DIRETO"

                iata_clean = clean_iata(iata_final)
                m_iata = iata_dict.get(iata_clean, {})

                gerente_resp = hist_data.get("Área Resp. Operação") or m_iata.get("Gerente_Responsavel", "Não Mapeado")
                obs_op = hist_data.get("Obs. Operação") or "Sem tratativa na operação"
                obs_replica = hist_data.get("Obs_Auditoria_Replica") or "-"

                st_lower = str(hist_data.get("Status_Geral") or "").strip().lower()
                area_lower = str(gerente_resp).strip().lower()
                obs_lower = str(obs_op).strip().lower()
                setor_hist_lower = str(hist_data.get("Setor") or "").strip().lower()

                e_suporte_backoffice = (
                    "backoffice" in st_lower or "suporte" in st_lower or "encaminhado" in st_lower
                    or any(k in area_lower for k in ["backoffice", "suporte", "benner", "katia", "ti"])
                    or any(k in setor_hist_lower for k in ["backoffice", "suporte"])
                    or any(p in obs_lower for p in ["ticket", "chamado", "backoffice", "suporte", "benner", "erro de integração", "erro integração", "aguardando suporte"])
                )

                if hist_data.get("Setor"):
                    setor_final = hist_data["Setor"]
                elif e_suporte_backoffice:
                    setor_final = "Suporte backoffice"
                    gerente_resp = "Suporte Backoffice"
                elif "JAIME SCHNAIDER" in area_lower.upper() or "UNIQUE" in area_lower.upper():
                    setor_final = "Unique"
                elif "CENTRAL DE EVENTOS" in area_lower.upper() or "EVENTO" in area_lower.upper():
                    setor_final = "Central de Eventos"
                elif "FABIANO SOUZA" in area_lower.upper() or "LAZER" in area_lower.upper() or "CONCIERGE" in area_lower.upper():
                    setor_final = "Concierge/Lazer"
                elif "PRIVATE" in area_lower.upper() or "SILVANA CELANI" in area_lower.upper():
                    setor_final = "Private"
                else:
                    setor_final = "Operação"

                keys_emissao = set(extract_keys(col3) + extract_keys(doc_val))

                b_match = next((benner_index[ek] for ek in keys_emissao if ek in benner_index), None)
                s_match = next((sabre_index[ek] for ek in keys_emissao if ek in sabre_index), None)
                l_match = next((lemon_index[ek] for ek in keys_emissao if ek in lemon_index), None)

                a_vista = safe_get_num(row, 8)
                a_credito = safe_get_num(row, 10)
                tarifa_emitida = a_vista + a_credito
                taxa_emitida = safe_get_num(row, 11)
                comissao = safe_get_num(row, 13)
                taxa_du = safe_get_num(row, 14)
                desc = safe_get_num(row, 15)
                incentivo = safe_get_num(row, 18)
                receita_emitida = comissao + taxa_du + incentivo
                vl_liquido = safe_get_num(row, 20)

                if b_match:
                    status_sistema = b_match["Benner_Situação"]
                    fornec_sistema = b_match["Benner_Fornecedor"]
                    loc_sistema = b_match["Benner_Localizador"]
                    rloc_cia_sistema = b_match["Benner_Rloc_Cia"]
                    bilhete_sistema = b_match["Benner_Bilhete"]
                    tarifa_sistema = b_match["Benner_Tarifa"]
                    taxa_sistema = b_match["Benner_Taxas"]
                    receita_sistema = b_match["Benner_Receita"]
                    cliente_sistema = b_match["Benner_Cliente"]
                    emissor_sistema = b_match["Benner_Emissor"]
                    sist_reserva = b_match["Benner_Sistema_Reserva"]

                    sigla_cia = col1.strip().upper()
                    nome_esperado_cia = MAPA_CIAS.get(sigla_cia, sigla_cia)
                    status_cia = (
                        "OK" if (nome_esperado_cia in fornec_sistema.upper() or sigla_cia in fornec_sistema.upper())
                        else "Divergência de Cia Aérea"
                    )

                    dif_tarifa = tarifa_emitida - tarifa_sistema
                    dif_taxa = taxa_emitida - taxa_sistema
                    dif_receita = receita_emitida - receita_sistema

                    if nome_fonte == "HOT":
                        if status_cia != "OK":
                            status_divergencia = "Divergência de Cia Aérea"
                            aba_destino = "98_OK_Divergencia_Operacao"
                        else:
                            status_divergencia = "Valores Corretos"
                            aba_destino = "98_OK_Sem_Divergencia_Concil"
                    else:
                        divs = []
                        if round(abs(dif_tarifa), 2) >= 0.01:
                            divs.append("Tarifa")
                        if round(abs(dif_taxa), 2) >= 0.01:
                            divs.append("Taxa")
                        if round(abs(dif_receita), 2) >= 0.01:
                            divs.append("Receita")
                        if status_cia != "OK":
                            divs.append("Cia Aérea")

                        if divs:
                            status_divergencia = f"Divergência de {' e '.join(divs)}"
                            aba_destino = "98_OK_Divergencia_Operacao"
                        else:
                            status_divergencia = "Valores Corretos"
                            aba_destino = "98_OK_Sem_Divergencia_Concil"

                    status_geral = hist_data.get("Status_Geral") or "Emitido e Lançado"
                else:
                    status_sistema = "NAO_CONSTA"
                    fornec_sistema = "-"
                    loc_sistema = col3
                    rloc_cia_sistema = (
                        col3 if (len(col3) <= 8 and not col3.isdigit() and col3.upper() not in ["[]", "NAN"])
                        else "-"
                    )
                    bilhete_sistema = col3
                    tarifa_sistema, taxa_sistema, receita_sistema = 0.0, 0.0, 0.0
                    cliente_sistema = ponto_venda_final
                    emissor_sistema, sist_reserva = "-", "-"
                    status_cia = "Pendente"
                    dif_tarifa, dif_taxa, dif_receita = tarifa_emitida, taxa_emitida, receita_emitida
                    status_divergencia = "Pendente de Lançamento"
                    status_geral = hist_data.get("Status_Geral") or "Pendente de Lançamento (Não Consta)"
                    aba_destino = "99_Base_Divergencias_Geral"

                sg_str = str(status_geral).lower()
                if any(term in sg_str for term in ["já lançado", "conciliado", "regularizado", "sem divergência"]):
                    status_divergencia = "Valores Corretos"
                    aba_destino = "98_OK_Sem_Divergencia_Concil"

                rec = {
                    "Ponto de venda": ponto_venda_final,
                    "Código Iata": iata_final,
                    "Gerentes": gerente_resp,
                    "CIA": col1,
                    "Fornecedor_Sistema": fornec_sistema,
                    "Status_Cia": status_cia,
                    "Bilhetes": clean_str_strict(bilhete_chave),
                    "Localizador_Sistema": clean_str_strict(loc_sistema),
                    "Rloc_Cia": clean_str_strict(rloc_cia_sistema),
                    "Status_Sistema": status_sistema,
                    "Data Emissão": dt_emissao,
                    "Pagto": pagto,
                    "A vista": a_vista,
                    "A credito": a_credito,
                    "Tarifa_Sistema": tarifa_sistema,
                    "Dif_Tarifa": dif_tarifa,
                    "Taxa": taxa_emitida,
                    "Taxa_Sistema": taxa_sistema,
                    "Dif_Taxa": dif_taxa,
                    "Comissão": comissao,
                    "Taxa DU": taxa_du,
                    "Desc.": desc,
                    "Incentivo": incentivo,
                    "Receita_Sistema": receita_sistema,
                    "Dif_Receita": dif_receita,
                    "VL. Líquido": vl_liquido,
                    "Status_Divergencia": status_divergencia,
                    "Consultor": s_match["Sabre_AGT"] if s_match else (l_match["Lemon_Consultor"] if l_match else "-"),
                    "Status_Geral": status_geral,
                    "Área Resp. Operação": gerente_resp,
                    "Obs. Operação": obs_op,
                    "Obs_Auditoria_Replica": obs_replica,
                    "Data_Modificacao": hist_data.get("Data_Modificacao") or "-",
                    "Usuario_Modificacao": hist_data.get("Usuario_Modificacao") or "-",
                    "Ultima_Alteracao": hist_data.get("Ultima_Alteracao") or "-",
                    "Emissor": emissor_sistema if b_match else (l_match["Lemon_Emissor_Reserva"] if l_match else "-"),
                    "Sistema Reserva": sist_reserva if b_match else (l_match["Lemon_Source"] if l_match else "-"),
                    "Cliente": cliente_sistema,
                    "Setor": setor_final,
                    "Aba_Destino": aba_destino,
                    "Tipo_Emissao_Lemon": l_match["Lemon_OnOff"] if l_match else "-",
                    "Consultor_Lemon": l_match["Lemon_Consultor"] if l_match else "-",
                    "Emissor_Reserva_Lemon": l_match["Lemon_Emissor_Reserva"] if l_match else "-",
                    "Centro_Custo_Lemon": l_match["Lemon_Centro_Custo"] if l_match else "-",
                    "Forma_Pagto_Lemon": l_match["Lemon_Forma_Pagto"] if l_match else "-",
                    "Autorizacao_Cartao_Lemon": l_match["Lemon_Autorizacao_Cartao"] if l_match else "-",
                }
                registros_conciliados.append(rec)

    if not registros_conciliados:
        print("\n⚠️ NENHUM REGISTRO DE EMISSÃO FOI PROCESSADO!\n")
        return

    df_master = pd.DataFrame(registros_conciliados)

    a_vista_col = df_master["A vista"] if "A vista" in df_master.columns else pd.Series(0.0, index=df_master.index)
    a_credito_col = df_master["A credito"] if "A credito" in df_master.columns else pd.Series(0.0, index=df_master.index)
    df_master["Tarifa_Total"] = pd.to_numeric(a_vista_col, errors="coerce").fillna(0.0) + pd.to_numeric(a_credito_col, errors="coerce").fillna(0.0)

    inc_col = df_master["Incentivo"] if "Incentivo" in df_master.columns else pd.Series(0.0, index=df_master.index)
    com_col = df_master["Comissão"] if "Comissão" in df_master.columns else pd.Series(0.0, index=df_master.index)
    tdu_col = df_master["Taxa DU"] if "Taxa DU" in df_master.columns else pd.Series(0.0, index=df_master.index)
    df_master["Receita_Total"] = (
        pd.to_numeric(inc_col, errors="coerce").fillna(0.0)
        + pd.to_numeric(com_col, errors="coerce").fillna(0.0)
        + pd.to_numeric(tdu_col, errors="coerce").fillna(0.0)
    )

    pareto_df = (
        df_master.groupby(["Setor", "Ponto de venda"])
        .agg(
            Qtd_Bilhetes=("Bilhetes", "count"),
            Tarifa_Pendente_R_=("Tarifa_Total", "sum"),
            Taxa_Pendente_R_=("Taxa", "sum"),
            Receita_Pendente_R_=("Receita_Total", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "Ponto de venda": "Cliente / Ponto de Venda",
                "Tarifa_Pendente_R_": "Tarifa_Pendente_R$",
                "Taxa_Pendente_R_": "Taxa_Pendente_R$",
                "Receita_Pendente_R_": "Receita_Pendente_R$",
            }
        )
        .sort_values(by=["Qtd_Bilhetes", "Tarifa_Pendente_R$"], ascending=False)
    )

    cols_35 = [
        "Ponto de venda", "Código Iata", "Gerentes", "CIA", "Fornecedor_Sistema", "Status_Cia",
        "Bilhetes", "Localizador_Sistema", "Rloc_Cia", "Status_Sistema", "Data Emissão", "Pagto",
        "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa", "Taxa_Sistema", "Dif_Taxa",
        "Comissão", "Taxa DU", "Desc.", "Incentivo", "Receita_Sistema", "Dif_Receita", "VL. Líquido",
        "Status_Divergencia", "Consultor", "Status_Geral", "Área Resp. Operação", "Obs. Operação",
        "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao",
        "Emissor", "Sistema Reserva", "Cliente", "Setor",
    ]

    cols_30_furo = [
        "Ponto de venda", "Código Iata", "Gerentes", "CIA", "Bilhetes", "Localizador_Sistema",
        "Rloc_Cia", "Status_Sistema", "Data Emissão", "Pagto", "A vista", "A credito", "Taxa",
        "Comissão", "Taxa DU", "Desc.", "Incentivo", "VL. Líquido", "Status_Geral",
        "Área Resp. Operação", "Obs. Operação", "Data_Modificacao", "Usuario_Modificacao", "Ultima_Alteracao",
        "Tipo_Emissao_Lemon", "Consultor_Lemon", "Emissor_Reserva_Lemon", "Centro_Custo_Lemon",
        "Forma_Pagto_Lemon", "Autorizacao_Cartao_Lemon", "Sistema Reserva", "Cliente", "Setor",
    ]

    df_98_div = df_master[df_master["Aba_Destino"] == "98_OK_Divergencia_Operacao"][cols_35] if len(df_master[df_master["Aba_Destino"] == "98_OK_Divergencia_Operacao"]) > 0 else pd.DataFrame(columns=cols_35)
    df_98_ok = df_master[df_master["Aba_Destino"] == "98_OK_Sem_Divergencia_Concil"][cols_35] if len(df_master[df_master["Aba_Destino"] == "98_OK_Sem_Divergencia_Concil"]) > 0 else pd.DataFrame(columns=cols_35)
    df_99_gen = df_master[df_master["Aba_Destino"] == "99_Base_Divergencias_Geral"][cols_30_furo] if len(df_master[df_master["Aba_Destino"] == "99_Base_Divergencias_Geral"]) > 0 else pd.DataFrame(columns=cols_30_furo)

    df_99_suporte = df_99_gen[df_99_gen["Setor"] == "Suporte backoffice"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)
    df_99_eventos = df_99_gen[df_99_gen["Setor"] == "Central de Eventos"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)
    df_99_lazer = df_99_gen[df_99_gen["Setor"] == "Concierge/Lazer"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)
    df_99_unique = df_99_gen[df_99_gen["Setor"] == "Unique"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)
    df_99_private = df_99_gen[df_99_gen["Setor"] == "Private"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)
    df_99_operacao = df_99_gen[df_99_gen["Setor"] == "Operação"] if len(df_99_gen) > 0 else pd.DataFrame(columns=cols_30_furo)

    print("\n[6/6] Exportando Dashboard Formatado em Excel...")

    with pd.ExcelWriter(out_file_model, engine="openpyxl") as writer:
        pareto_df.to_excel(writer, sheet_name="01_Pareto_Cliente", index=False)
        df_98_div.to_excel(writer, sheet_name="98_OK_Divergencia_Operacao", index=False)
        df_98_ok.to_excel(writer, sheet_name="98_OK_Sem_Divergencia_Concil", index=False)
        df_99_gen.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
        df_99_suporte.to_excel(writer, sheet_name="99_Suporte backoffice", index=False)
        df_99_eventos.to_excel(writer, sheet_name="99_Central de Eventos", index=False)
        df_99_lazer.to_excel(writer, sheet_name="99_Concierge-Lazer", index=False)
        df_99_unique.to_excel(writer, sheet_name="99_Unique", index=False)
        df_99_private.to_excel(writer, sheet_name="99_Private", index=False)
        df_99_operacao.to_excel(writer, sheet_name="99_Operação", index=False)

        if not df_log_acumulado.empty:
            df_log_acumulado.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)
        else:
            pd.DataFrame(
                columns=["Data_Hora", "Bilhete", "Usuario_Acao", "Status_Anterior", "Novo_Status", "Area_Anterior", "Nova_Area", "Observacao", "Tipo_Interacao"]
            ).to_excel(writer, sheet_name="00_Log_Auditoria", index=False)

    wb = openpyxl.load_workbook(out_file_model)

    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
    )
    fill_alerta = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_alerta = Font(color="9C0006", bold=True)

    currency_cols = [
        "Tarifa_Pendente_R$", "Taxa_Pendente_R$", "Receita_Pendente_R$", "Receita_Risco",
        "A vista", "A credito", "Tarifa_Sistema", "Dif_Tarifa", "Taxa", "Taxa_Sistema",
        "Dif_Taxa", "Comissão", "Taxa DU", "Desc.", "Incentivo", "Receita_Sistema",
        "Dif_Receita", "VL. Líquido",
    ]

    for sheetname in tqdm(wb.sheetnames, desc="Formatando Abas do Excel", unit="aba"):
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True

        if sheetname not in ["01_Pareto_Cliente", "00_Log_Auditoria"]:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

        max_row = ws.max_row

        if max_row > 1 and sheetname not in ["01_Pareto_Cliente", "00_Log_Auditoria"]:
            ws.cell(row=max_row + 1, column=1, value="TOTAL").font = Font(bold=True)
            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                col_name = str(ws.cell(row=1, column=col_idx).value or "")
                if col_name in currency_cols or col_name in ["Qtd_Bilhetes"]:
                    c_tot = ws.cell(
                        row=max_row + 1,
                        column=col_idx,
                        value=f"=SUM({col_letter}2:{col_letter}{max_row})",
                    )
                    c_tot.font = Font(bold=True)

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            col_name = str(ws.cell(row=1, column=col_idx).value or "")

            len_vals = [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, ws.max_row + 1)]
            max_len = max(len_vals) if len_vals else 10
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            if col_name in currency_cols:
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col_idx)
                    if cell.value not in ["-", None]:
                        try:
                            cell.value = float(cell.value)
                            cell.number_format = "R$ #,##0.00"
                        except (ValueError, TypeError):
                            pass

            if col_name.startswith("Dif_"):
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col_idx)
                    try:
                        if round(abs(float(cell.value)), 2) >= 0.01:
                            cell.fill = fill_alerta
                            cell.font = font_alerta
                    except (ValueError, TypeError):
                        pass

    wb.save(out_file_model)

    # 1. Salva Memória Protegida CSV
    salvar_csv_atomico(
        df_master[["Bilhetes", "Status_Geral", "Área Resp. Operação", "Obs. Operação", "Setor", "Ponto de venda", "Código Iata"]].drop_duplicates(subset=["Bilhetes"]),
        "Historico_Tratativas.csv"
    )

    # 2. Salva SQLite Local
    conn = sqlite3.connect("revenue_assurance.db")
    try:
        with conn:
            df_99_gen.to_sql("tb_master", conn, if_exists="replace", index=False)
            df_98_div.to_sql("tb_div_op", conn, if_exists="replace", index=False)
            df_98_ok.to_sql("tb_sem_div", conn, if_exists="replace", index=False)
            df_99_suporte.to_sql("tb_backoffice", conn, if_exists="replace", index=False)
        print("⚡ Base SQLite (revenue_assurance.db) atualizada com sucesso!")
    except Exception as e:
        print(f"⚠️ Erro ao atualizar o banco SQLite: {e}")
    finally:
        conn.close()

    # 3. Salva/Sincroniza com o Supabase Nuvem em Lote Único
    sincronizar_supabase_fim(df_master)

    print("\n" + "=" * 75)
    print(f" AUDITORIA CONCLUÍDA COM SUCESSO! SALVO EM: {out_file_model} ".center(75, "="))
    print("=" * 75)


if __name__ == "__main__":
    executar_auditoria()