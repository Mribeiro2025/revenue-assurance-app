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
MAPA_CIAS = {
    "JJ": "LATAM", "LA": "LATAM", "PZ": "LATAM", "4C": "LATAM", "XL": "LATAM", "4M": "LATAM",
    "G3": "GOL", "AD": "AZUL", "TP": "TAP", "CM": "COPA", "AA": "AMERICAN", "UA": "UNITED",
    "DL": "DELTA", "AF": "AIR FRANCE", "KL": "KLM", "AR": "AEROLINEAS", "IB": "IBERIA",
    "UX": "EUROPA", "BA": "BRITISH", "AV": "AVIANCA", "O6": "AVIANCA", "AC": "CANADA",
    "AM": "AEROMEXICO", "QR": "QATAR", "EK": "EMIRATES", "TK": "TURKISH", "LH": "LUFTHANSA",
    "LX": "SWISS", "ET": "ETHIOPIAN", "AT": "MAROC", "SA": "SOUTH AFRICAN", "OB": "BOLIVIANA",
    "PY": "PARAGUAY", "HR": "HAHN", "AZ": "ITA AIRWAYS", "JA": "JETSMART",
}

def obter_engine_supabase():
    """Obtém a conexão SQLAlchemy com o Supabase."""
    try:
        from sqlalchemy import create_engine
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
                    if db_url.startswith("postgresql://"):
                        db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
                    return create_engine(db_url, pool_pre_ping=True)
    except Exception:
        pass
    return None

def criar_tabelas_supabase_se_nao_existirem(engine):
    """Garante a criação de todas as tabelas e views necessárias no Supabase."""
    if not engine:
        return
    from sqlalchemy import text
    ddl_script = """
    CREATE TABLE IF NOT EXISTS tb_conciliacao_base (
        bilhete VARCHAR(255) PRIMARY KEY,
        localizador VARCHAR(255),
        rloc_cia VARCHAR(255),
        cia VARCHAR(255),
        fornecedor_sistema VARCHAR(255),
        status_cia VARCHAR(255),
        status_sistema VARCHAR(255),
        data_emissao VARCHAR(255),
        pagto VARCHAR(255),
        sv VARCHAR(255),
        fop VARCHAR(255),
        a_vista NUMERIC(15,2) DEFAULT 0.00,
        a_credito NUMERIC(15,2) DEFAULT 0.00,
        tarifa_sistema NUMERIC(15,2) DEFAULT 0.00,
        dif_tarifa NUMERIC(15,2) DEFAULT 0.00,
        taxa_emitida NUMERIC(15,2) DEFAULT 0.00,
        taxa_sistema NUMERIC(15,2) DEFAULT 0.00,
        dif_taxa NUMERIC(15,2) DEFAULT 0.00,
        comissao NUMERIC(15,2) DEFAULT 0.00,
        taxa_du NUMERIC(15,2) DEFAULT 0.00,
        desconto NUMERIC(15,2) DEFAULT 0.00,
        incentivo NUMERIC(15,2) DEFAULT 0.00,
        receita_sistema NUMERIC(15,2) DEFAULT 0.00,
        dif_receita NUMERIC(15,2) DEFAULT 0.00,
        vl_liquido NUMERIC(15,2) DEFAULT 0.00,
        status_divergencia VARCHAR(255),
        consultor VARCHAR(255),
        emissor VARCHAR(255),
        sistema_reserva VARCHAR(255),
        cliente VARCHAR(255),
        ponto_venda VARCHAR(255),
        codigo_iata VARCHAR(255),
        gerente_sugerido VARCHAR(255),
        setor_sugerido VARCHAR(255),
        aba_destino VARCHAR(255),
        fonte_relatorio VARCHAR(255),
        tipo_emissao_lemon VARCHAR(255),
        consultor_lemon VARCHAR(255),
        emissor_reserva_lemon VARCHAR(255),
        centro_custo_lemon VARCHAR(255),
        forma_pagto_lemon VARCHAR(255),
        autorizacao_cartao_lemon VARCHAR(255),
        modalidade_lemon VARCHAR(255),
        emissao_online_lemon VARCHAR(255),
        source_lemon VARCHAR(255),
        ponto_venda_lemon VARCHAR(255),
        data_processamento TIMESTAMP DEFAULT NOW()
    );

    -- Garantia de adição das novas colunas em tabelas legadas no Supabase
    ALTER TABLE tb_conciliacao_base ADD COLUMN IF NOT EXISTS sv VARCHAR(255);
    ALTER TABLE tb_conciliacao_base ADD COLUMN IF NOT EXISTS fop VARCHAR(255);

    CREATE TABLE IF NOT EXISTS tratativas (
        bilhete VARCHAR(255) PRIMARY KEY,
        status_geral VARCHAR(255),
        area_resp VARCHAR(255),
        obs_operacao TEXT,
        setor VARCHAR(255),
        gerentes VARCHAR(255),
        sv VARCHAR(255),
        fop VARCHAR(255),
        rloc_cia VARCHAR(255),
        usuario_modificacao VARCHAR(255),
        data_modificacao TIMESTAMP DEFAULT NOW()
    );

    CREATE OR REPLACE VIEW vw_conciliacao_mestra AS
    SELECT 
        b.bilhete AS "Bilhetes",
        b.ponto_venda AS "Ponto de venda",
        b.codigo_iata AS "Código Iata",
        COALESCE(t.sv, b.sv, '-') AS "SV",
        COALESCE(t.rloc_cia, b.rloc_cia, '-') AS "Rloc_Cia",
        COALESCE(t.gerentes, t.area_resp, b.gerente_sugerido, 'Não Mapeado') AS "Gerentes",
        b.cia AS "CIA",
        b.fornecedor_sistema AS "Fornecedor_Sistema",
        b.status_cia AS "Status_Cia",
        b.localizador AS "Localizador_Sistema",
        b.status_sistema AS "Status_Sistema",
        b.data_emissao AS "Data Emissão",
        COALESCE(t.fop, b.fop, b.pagto, '-') AS "FOP",
        b.pagto AS "Pagto",
        b.a_vista AS "A vista",
        b.a_credito AS "A credito",
        b.tarifa_sistema AS "Tarifa_Sistema",
        b.dif_tarifa AS "Dif_Tarifa",
        b.taxa_emitida AS "Taxa",
        b.taxa_sistema AS "Taxa_Sistema",
        b.dif_taxa AS "Dif_Taxa",
        b.comissao AS "Comissão",
        b.taxa_du AS "Taxa DU",
        b.desconto AS "Desc.",
        b.incentivo AS "Incentivo",
        b.receita_sistema AS "Receita_Sistema",
        b.dif_receita AS "Dif_Receita",
        b.vl_liquido AS "VL. Líquido",
        b.status_divergencia AS "Status_Divergencia",
        b.consultor AS "Consultor",
        COALESCE(t.status_geral, b.status_divergencia, 'Pendente de Lançamento (Não Consta)') AS "Status_Geral",
        COALESCE(t.area_resp, b.gerente_sugerido, 'Operação') AS "Área Resp. Operação",
        COALESCE(t.obs_operacao, '') AS "Obs. Operação",
        t.data_modificacao AS "Data_Modificacao",
        t.usuario_modificacao AS "Usuario_Modificacao",
        t.data_modificacao AS "Ultima_Alteracao",
        b.emissor AS "Emissor",
        b.sistema_reserva AS "Sistema Reserva",
        b.cliente AS "Cliente",
        COALESCE(t.setor, b.setor_sugerido, 'Operação') AS "Setor",
        b.tipo_emissao_lemon AS "Tipo_Emissao_Lemon",
        b.consultor_lemon AS "Consultor_Lemon",
        b.emissor_reserva_lemon AS "Emissor_Reserva_Lemon",
        b.centro_custo_lemon AS "Centro_Custo_Lemon",
        b.forma_pagto_lemon AS "Forma_Pagto_Lemon",
        b.autorizacao_cartao_lemon AS "Autorizacao_Cartao_Lemon",
        b.modalidade_lemon AS "Modalidade_Lemon",
        b.emissao_online_lemon AS "Emissao_Online_Lemon",
        b.source_lemon AS "Source_Lemon",
        b.ponto_venda_lemon AS "Ponto_Venda_Lemon",
        b.aba_destino AS "Aba_Destino",
        b.fonte_relatorio AS "Fonte_Relatorio"
    FROM tb_conciliacao_base b
    LEFT JOIN tratativas t ON b.bilhete = t.bilhete;
    """
    try:
        with engine.begin() as conn:
            conn.execute(text(ddl_script))
        print("⚡ Estrutura do Supabase (Tabelas e Visão) validada com sucesso!")
    except Exception as e:
        print(f"⚠️ Aviso ao verificar estrutura no Supabase: {e}")

def auto_converter_benner_parquet():
    arq_xlsx = "Acumulado.xlsx"
    arq_parquet = "Acumulado.parquet"
    if not os.path.exists(arq_xlsx):
        return
    precisa_converter = False
    if not os.path.exists(arq_parquet):
        precisa_converter = True
    else:
        if os.path.getmtime(arq_xlsx) > os.path.getmtime(arq_parquet):
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
    s = str(val).strip()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s.upper()

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

def gerar_chave_composta(codigo, data_str):
    cod_limpo = clean_str_strict(codigo).upper()
    dt_limpa = clean_str_strict(data_str).replace("/", "").replace("-", "")
    if cod_limpo and dt_limpa:
        return f"{cod_limpo}_{dt_limpa}"
    return cod_limpo

def identificar_fonte_relatorio(nome_fonte, emissor_str=""):
    s_em = str(emissor_str).lower()
    if "lemontech" in s_em or "virtual" in s_em:
        return "🤖 Relatório Lemontech (Virtual)"
    if "GOL" in nome_fonte.upper():
        return "✈️ Relatório GOL Direct"
    if "AZUL" in nome_fonte.upper():
        return "✈️ Relatório AZUL Direct"
    if "BSP" in nome_fonte.upper() or "HOT" in nome_fonte.upper():
        return "🌐 Relatório BSP (IATA)"
    return "🏢 Relatório ERP Benner"

def carregar_tratativas_e_logs_anteriores(out_file):
    dict_historico = {}
    df_log_antigo = pd.DataFrame()
    
    def indexar_memoria(b_raw, dados):
        if not b_raw or str(b_raw).strip().lower() in ["nan", "none", "", "-"]:
            return
        for k in extract_keys(b_raw):
            if k not in dict_historico:
                dict_historico[k] = {}
            for col, val in dados.items():
                val_str = str(val).strip() if pd.notna(val) else ""
                if val_str and val_str not in ["-", "nan", "none", "Sem tratativa na operação"]:
                    dict_historico[k][col] = val

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
                            obs_limpa = str(r.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip()
                            indexar_memoria(b_key, {
                                "Status_Geral": r.get("Status_Geral"),
                                "Área Resp. Operação": r.get("Área Resp. Operação"),
                                "Obs. Operação": obs_limpa,
                                "Setor": r.get("Setor"),
                                "Gerentes": r.get("Gerentes"),
                                "Ponto de venda": r.get("Ponto de venda"),
                                "Código Iata": r.get("Código Iata"),
                                "SV": r.get("SV"),
                                "FOP": r.get("FOP"),
                                "Rloc_Cia": r.get("Rloc_Cia") or r.get("Rloc_CIA"),
                            })
        except Exception as e:
            print(f"⚠️ Aviso ao carregar histórico do Excel: {e}")

    engine_sb = obter_engine_supabase()
    if engine_sb:
        try:
            df_sb = pd.read_sql("SELECT bilhete, status_geral, area_resp, obs_operacao, setor, gerentes, sv, fop, rloc_cia FROM tratativas", engine_sb)
            if not df_sb.empty:
                for r in df_sb.to_dict("records"):
                    b_key = clean_str_strict(r.get("bilhete"))
                    if b_key:
                        obs_limpa = str(r.get("obs_operacao", "")).replace("Sem tratativa na operação", "").strip()
                        indexar_memoria(b_key, {
                            "Status_Geral": r.get("status_geral"),
                            "Área Resp. Operação": r.get("area_resp"),
                            "Obs. Operação": obs_limpa,
                            "Setor": r.get("setor"),
                            "Gerentes": r.get("gerentes"),
                            "SV": r.get("sv"),
                            "FOP": r.get("fop"),
                            "Rloc_Cia": r.get("rloc_cia"),
                        })
                print(f"☁️ Supabase conectado! {len(df_sb)} tratativas resgatadas com sucesso.")
        except Exception as e:
            print(f"⚠️ Aviso ao carregar dados do Supabase: {e}")

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
        caminho_envio = os.path.join(PASTA_ENVIO, nome)
        if os.path.exists(caminho_envio):
            return caminho_envio
    return None

def salvar_csv_atomico(df, caminho_csv):
    tmp_path = f"{caminho_csv}.tmp"
    df.to_csv(tmp_path, index=False, encoding="utf-8-sig")
    if os.path.exists(caminho_csv):
        os.remove(caminho_csv)
    os.rename(tmp_path, caminho_csv)

def sincronizar_supabase_tudo(df_master):
    """Envia todos os bilhetes auditados para a tb_conciliacao_base no Supabase em lote."""
    engine_sb = obter_engine_supabase()
    if not engine_sb or df_master.empty:
        return
    
    criar_tabelas_supabase_se_nao_existirem(engine_sb)
    from sqlalchemy import text
    
    upsert_base_sql = text("""
        INSERT INTO tb_conciliacao_base (
            bilhete, localizador, rloc_cia, cia, fornecedor_sistema, status_cia, status_sistema,
            data_emissao, pagto, sv, fop, a_vista, a_credito, tarifa_sistema, dif_tarifa, taxa_emitida,
            taxa_sistema, dif_taxa, comissao, taxa_du, desconto, incentivo, receita_sistema,
            dif_receita, vl_liquido, status_divergencia, consultor, emissor, sistema_reserva,
            cliente, ponto_venda, codigo_iata, gerente_sugerido, setor_sugerido, aba_destino,
            fonte_relatorio, tipo_emissao_lemon, consultor_lemon, emissor_reserva_lemon,
            centro_custo_lemon, forma_pagto_lemon, autorizacao_cartao_lemon, modalidade_lemon,
            emissao_online_lemon, source_lemon, ponto_venda_lemon, data_processamento
        ) VALUES (
            :bilhete, :localizador, :rloc_cia, :cia, :fornecedor_sistema, :status_cia, :status_sistema,
            :data_emissao, :pagto, :sv, :fop, :a_vista, :a_credito, :tarifa_sistema, :dif_tarifa, :taxa_emitida,
            :taxa_sistema, :dif_taxa, :comissao, :taxa_du, :desconto, :incentivo, :receita_sistema,
            :dif_receita, :vl_liquido, :status_divergencia, :consultor, :emissor, :sistema_reserva,
            :cliente, :ponto_venda, :codigo_iata, :gerente_sugerido, :setor_sugerido, :aba_destino,
            :fonte_relatorio, :tipo_emissao_lemon, :consultor_lemon, :emissor_reserva_lemon,
            :centro_custo_lemon, :forma_pagto_lemon, :autorizacao_cartao_lemon, :modalidade_lemon,
            :emissao_online_lemon, :source_lemon, :ponto_venda_lemon, NOW()
        ) ON CONFLICT (bilhete) DO UPDATE SET
            localizador = EXCLUDED.localizador,
            rloc_cia = EXCLUDED.rloc_cia,
            cia = EXCLUDED.cia,
            fornecedor_sistema = EXCLUDED.fornecedor_sistema,
            status_cia = EXCLUDED.status_cia,
            status_sistema = EXCLUDED.status_sistema,
            data_emissao = EXCLUDED.data_emissao,
            pagto = EXCLUDED.pagto,
            sv = EXCLUDED.sv,
            fop = EXCLUDED.fop,
            a_vista = EXCLUDED.a_vista,
            a_credito = EXCLUDED.a_credito,
            tarifa_sistema = EXCLUDED.tarifa_sistema,
            dif_tarifa = EXCLUDED.dif_tarifa,
            taxa_emitida = EXCLUDED.taxa_emitida,
            taxa_sistema = EXCLUDED.taxa_sistema,
            dif_taxa = EXCLUDED.dif_taxa,
            comissao = EXCLUDED.comissao,
            taxa_du = EXCLUDED.taxa_du,
            desconto = EXCLUDED.desconto,
            incentivo = EXCLUDED.incentivo,
            receita_sistema = EXCLUDED.receita_sistema,
            dif_receita = EXCLUDED.dif_receita,
            vl_liquido = EXCLUDED.vl_liquido,
            status_divergencia = EXCLUDED.status_divergencia,
            consultor = EXCLUDED.consultor,
            emissor = EXCLUDED.emissor,
            sistema_reserva = EXCLUDED.sistema_reserva,
            cliente = EXCLUDED.cliente,
            ponto_venda = EXCLUDED.ponto_venda,
            codigo_iata = EXCLUDED.codigo_iata,
            gerente_sugerido = EXCLUDED.gerente_sugerido,
            setor_sugerido = EXCLUDED.setor_sugerido,
            aba_destino = EXCLUDED.aba_destino,
            fonte_relatorio = EXCLUDED.fonte_relatorio,
            tipo_emissao_lemon = EXCLUDED.tipo_emissao_lemon,
            consultor_lemon = EXCLUDED.consultor_lemon,
            emissor_reserva_lemon = EXCLUDED.emissor_reserva_lemon,
            centro_custo_lemon = EXCLUDED.centro_custo_lemon,
            forma_pagto_lemon = EXCLUDED.forma_pagto_lemon,
            autorizacao_cartao_lemon = EXCLUDED.autorizacao_cartao_lemon,
            modalidade_lemon = EXCLUDED.modalidade_lemon,
            emissao_online_lemon = EXCLUDED.emissao_online_lemon,
            source_lemon = EXCLUDED.source_lemon,
            ponto_venda_lemon = EXCLUDED.ponto_venda_lemon,
            data_processamento = NOW();
    """)
    
    lote_base = []
    for _, r in df_master.iterrows():
        b_val = clean_str_strict(r.get("Bilhetes"))
        if b_val:
            lote_base.append({
                "bilhete": b_val,
                "localizador": clean_str_strict(r.get("Localizador_Sistema")),
                "rloc_cia": clean_str_strict(r.get("Rloc_Cia") or r.get("Rloc_CIA")),
                "cia": clean_str_strict(r.get("CIA")),
                "fornecedor_sistema": clean_str_strict(r.get("Fornecedor_Sistema")),
                "status_cia": clean_str_strict(r.get("Status_Cia")),
                "status_sistema": clean_str_strict(r.get("Status_Sistema")),
                "data_emissao": clean_str_strict(r.get("Data Emissão")),
                "pagto": clean_str_strict(r.get("Pagto")),
                "sv": clean_str_strict(r.get("SV")),
                "fop": clean_str_strict(r.get("FOP")),
                "a_vista": clean_num(r.get("A vista")),
                "a_credito": clean_num(r.get("A credito")),
                "tarifa_sistema": clean_num(r.get("Tarifa_Sistema")),
                "dif_tarifa": clean_num(r.get("Dif_Tarifa")),
                "taxa_emitida": clean_num(r.get("Taxa")),
                "taxa_sistema": clean_num(r.get("Taxa_Sistema")),
                "dif_taxa": clean_num(r.get("Dif_Taxa")),
                "comissao": clean_num(r.get("Comissão")),
                "taxa_du": clean_num(r.get("Taxa DU")),
                "desconto": clean_num(r.get("Desc.")),
                "incentivo": clean_num(r.get("Incentivo")),
                "receita_sistema": clean_num(r.get("Receita_Sistema")),
                "dif_receita": clean_num(r.get("Dif_Receita")),
                "vl_liquido": clean_num(r.get("VL. Líquido")),
                "status_divergencia": clean_str_strict(r.get("Status_Divergencia")),
                "consultor": clean_str_strict(r.get("Consultor")),
                "emissor": clean_str_strict(r.get("Emissor")),
                "sistema_reserva": clean_str_strict(r.get("Sistema Reserva")),
                "cliente": clean_str_strict(r.get("Cliente")),
                "ponto_venda": clean_str_strict(r.get("Ponto de venda")),
                "codigo_iata": clean_str_strict(r.get("Código Iata")),
                "gerente_sugerido": clean_str_strict(r.get("Área Resp. Operação") or r.get("Gerentes")),
                "setor_sugerido": clean_str_strict(r.get("Setor")),
                "aba_destino": clean_str_strict(r.get("Aba_Destino")),
                "fonte_relatorio": clean_str_strict(r.get("Fonte_Relatorio")),
                "tipo_emissao_lemon": clean_str_strict(r.get("Tipo_Emissao_Lemon")),
                "consultor_lemon": clean_str_strict(r.get("Consultor_Lemon")),
                "emissor_reserva_lemon": clean_str_strict(r.get("Emissor_Reserva_Lemon")),
                "centro_custo_lemon": clean_str_strict(r.get("Centro_Custo_Lemon")),
                "forma_pagto_lemon": clean_str_strict(r.get("Forma_Pagto_Lemon")),
                "autorizacao_cartao_lemon": clean_str_strict(r.get("Autorizacao_Cartao_Lemon")),
                "modalidade_lemon": clean_str_strict(r.get("Modalidade_Lemon")),
                "emissao_online_lemon": clean_str_strict(r.get("Emissao_Online_Lemon")),
                "source_lemon": clean_str_strict(r.get("Source_Lemon")),
                "ponto_venda_lemon": clean_str_strict(r.get("Ponto_Venda_Lemon"))
            })
            
    if lote_base:
        try:
            tam_lote = 500
            for i in range(0, len(lote_base), tam_lote):
                chunk = lote_base[i:i + tam_lote]
                with engine_sb.begin() as conn:
                    conn.execute(upsert_base_sql, chunk)
            print(f"☁️ {len(lote_base)} registros sincronizados na 'tb_conciliacao_base' do Supabase!")
        except Exception as e:
            print(f"⚠️ Erro ao sincronizar lote na 'tb_conciliacao_base': {e}")

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
        cols_map = {re.sub(r"[^a-z]", "", str(col).lower()): col for col in df_iata.columns}
        col_codigo = next((cols_map[k] for k in cols_map if "cod" in k or "iata" in k), df_iata.columns[0])
        col_nome = next((cols_map[k] for k in cols_map if "nome" in k or "razao" in k or "empresa" in k), df_iata.columns[1] if len(df_iata.columns) > 1 else df_iata.columns[0])
        col_gerente = next((cols_map[k] for k in cols_map if "gerente" in k or "resp" in k), df_iata.columns[2] if len(df_iata.columns) > 2 else df_iata.columns[0])
        
        for idx, r in tqdm(df_iata.iterrows(), total=len(df_iata), desc="Indexando IATAs", unit="reg"):
            k = clean_iata(r.get(col_codigo))
            if k:
                iata_dict[k] = {
                    "Nome_IATA_Oficial": clean_str_strict(r.get(col_nome)),
                    "Gerente_Responsavel": clean_str_strict(r.get(col_gerente)),
                }

    print("\n[2/5] Indexando Extrato OBT Lemontech (Lendo SV e dados do cliente)...")
    caminho_lemon = encontrar_arquivo(["Extrato_Bilhetes_lemontech.xlsx", "Extrato_Bilhetes_lemontech.XLSX"])
    lemon_index = {}
    if caminho_lemon:
        df_lemon = pd.read_excel(caminho_lemon, engine="openpyxl")
        for idx, r in tqdm(df_lemon.iterrows(), total=len(df_lemon), desc="Indexando Lemontech", unit="reg"):
            cc_desc = f"{clean_str_strict(r.get('Centro de Custo'))} - {clean_str_strict(r.get('Descrição Centro de Custo'))}".strip(" -")
            
            # CAPTURA DO CAMPO SV (Solicitação no relatório Lemontech)
            sv_lemon = clean_str_strict(r.get("Solicitação") or r.get("Solicitacao") or r.get("Nº Pedido"))
            
            info_lemon = {
                "Lemon_SV": sv_lemon if sv_lemon else "-",
                "Lemon_Solicitante": clean_str_strict(r.get("Solicitante")) or "-",
                "Lemon_Passageiro": clean_str_strict(r.get("Passageiro")) or "-",
                "Lemon_Consultor": clean_str_strict(r.get("Consultor")) or "-",
                "Lemon_Emissor_Reserva": clean_str_strict(r.get("Emissor Reserva")) or "-",
                "Lemon_Centro_Custo": cc_desc if cc_desc else "-",
                "Lemon_Forma_Pagto": clean_str_strict(r.get("Forma de Pagamento")) or "-",
                "Lemon_Autorizacao_Cartao": clean_str_strict(r.get("Autorização do Cartão")) or "-",
                "Lemon_OnOff": clean_str_strict(r.get("On|Off")) or "-",
                "Lemon_Modalidade": clean_str_strict(r.get("Modalidade")) or "-",
                "Lemon_Emissao_Online": clean_str_strict(r.get("Emissão Online")) or "-",
                "Lemon_Source": clean_str_strict(r.get("Source")) or "-",
                "Lemon_Ponto_Venda": clean_str_strict(r.get("Ponto de Venda")) or "-",
            }
            for col_k in ["Bilhete", "Nº Pedido", "Solicitação"]:
                for k in extract_keys(r.get(col_k)):
                    if k not in lemon_index:
                        lemon_index[k] = info_lemon

    print("\n[3/5] Indexando ERP Benner (Lendo FOP e Cliente oficial)...")
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
            dt_benner_raw = clean_str_strict(r.get("Data Emissão") or r.get("Data") or r.get("Data Emissao") or "")
            
            # CAPTURA DO CAMPO FOP (Forma de Recebimento no Acumulado.xlsx)
            fop_benner = clean_str_strict(r.get("Forma de Recebimento") or r.get("Bandeira") or r.get("Forma de Pagamento")) or "-"
            
            # CAPTURA DO CAMPO CLIENTE (Cliente / Apelido Cliente no Acumulado.xlsx)
            cliente_benner = clean_str_strict(r.get("Cliente") or r.get("Apelido Cliente")) or "-"
            
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
                "Benner_Cliente": cliente_benner,
                "Benner_FOP": fop_benner,
                "Benner_Emissor": clean_str_strict(r.get("Emissor") or r.get("Agente Criação")) or "-",
                "Benner_Passageiro": clean_str_strict(r.get("Passageiro")) or "-",
                "Benner_Sistema_Reserva": clean_str_strict(r.get("Sistema Reserva")) or "-",
            }
            for col in key_cols_benner:
                for k in extract_keys(r.get(col)):
                    if k not in benner_index:
                        benner_index[k] = info
                    if dt_benner_raw:
                        ch_comp = gerar_chave_composta(k, dt_benner_raw)
                        if ch_comp not in benner_index:
                            benner_index[ch_comp] = info

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
        (["AZUL.XLSX", "AZUL.xlsx"], "AZUL"),
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
        col_map = {}
        for r_idx in tqdm(range(len(df_raw)), desc=f"Auditando {nome_fonte}", unit="linha"):
            row = df_raw.iloc[r_idx]
            col0 = safe_get_col(row, 0)
            col1 = safe_get_col(row, 1)
            col3 = safe_get_col(row, 3)
            col4 = safe_get_col(row, 4)
            col5 = safe_get_col(row, 5)
            col6 = safe_get_col(row, 6)
            
            row_str = [str(x).strip().upper() for x in row.values]
            if any("BILHETE" in x or "A VISTA" in x or "A CREDITO" in x or "TARIFA" in x or "DOCUMENTO" in x for x in row_str):
                col_map = {}
                for c_idx, val in enumerate(row.values):
                    if pd.notna(val):
                        v_clean = str(val).strip().upper()
                        col_map[v_clean] = c_idx
                continue
                
            if re.match(r"^\s*(\d{2}-?\d{5}-?\d?|\d{7,8})\s*$", col0):
                curr_iata = col0
                if col4:
                    curr_ponto_venda = col4
                continue
            if col3 != "" and col3 not in ["BILHETE\\RLOC", "BILHETE", "RLOC", "[]"] and col1 != "CIA":
                if col3.upper() in ["BILHETE\\RLOC", "BILHETE", "RLOC"] or col5.upper() in ["EMISSÃO", "EMISSAO"]:
                    continue
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
                doc_col_idx = col_map.get("DOCUMENTO", 22)
                doc_val = safe_get_col(row, doc_col_idx)
                clean_doc = re.sub(r"[^0-9]", "", doc_val)
                if clean_doc and len(clean_doc) >= 8:
                    bilhete_chave = doc_val
                    loc_reserva = col3
                else:
                    bilhete_chave = col3
                    loc_reserva = doc_val
                
                hist_data = buscar_memoria(dict_historico, bilhete_chave)
                iata_final = curr_iata or hist_data.get("Código Iata") or "DIRETO"
                iata_clean = clean_iata(iata_final)
                m_iata = iata_dict.get(iata_clean, {})
                
                gerente_do_dict = clean_str_strict(m_iata.get("Gerente_Responsavel"))
                hist_gerente = clean_str_strict(hist_data.get("Área Resp. Operação") or hist_data.get("Gerentes"))
                
                if hist_gerente and hist_gerente.lower() not in ["não mapeado", "nao mapeado", "-", "sem mapeamento", "nan", "none"]:
                    gerente_resp = hist_gerente
                elif gerente_do_dict and gerente_do_dict.lower() not in ["não mapeado", "nao mapeado", "-", "nan", "none"]:
                    gerente_resp = gerente_do_dict
                else:
                    gerente_resp = "Não Mapeado"
                
                ponto_curr = clean_str_strict(curr_ponto_venda)
                ponto_hist = clean_str_strict(hist_data.get("Ponto de venda"))
                ponto_oficial = clean_str_strict(m_iata.get("Nome_IATA_Oficial"))
                
                if ponto_curr and ponto_curr.lower() not in ["não mapeado", "nao mapeado", "-", "nan", "none"]:
                    ponto_venda_final = ponto_curr
                elif ponto_hist and ponto_hist.lower() not in ["não mapeado", "nao mapeado", "-", "nan", "none"]:
                    ponto_venda_final = ponto_hist
                elif ponto_oficial and ponto_oficial.lower() not in ["não mapeado", "nao mapeado", "-", "nan", "none"]:
                    ponto_venda_final = ponto_oficial
                else:
                    ponto_venda_final = "Não Mapeado"
                
                obs_op = str(hist_data.get("Obs. Operação", "")).replace("Sem tratativa na operação", "").strip()
                st_lower = str(hist_data.get("Status_Geral") or "").strip().lower()
                area_lower = str(gerente_resp).strip().lower()
                obs_lower = str(obs_op).strip().lower()
                tem_gerente_bo = any(k in area_lower for k in ["katia martins", "kátia martins", "suporte backoffice", "backoffice"])
                tem_chamado_obs = any(p in obs_lower or p in st_lower for p in ["ticket", "chamado", "suporte"])
                e_suporte_backoffice = tem_gerente_bo and tem_chamado_obs
                
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
                
                keys_emissao = set(extract_keys(bilhete_chave) + extract_keys(col3) + extract_keys(doc_val))
                b_match = None
                for ek in keys_emissao:
                    ch_comp = gerar_chave_composta(ek, dt_emissao)
                    if ch_comp in benner_index:
                        b_match = benner_index[ch_comp]
                        break
                    elif ek in benner_index:
                        b_match = benner_index[ek]
                
                s_match = next((sabre_index[ek] for ek in keys_emissao if ek in sabre_index), None)
                l_match = next((lemon_index[ek] for ek in keys_emissao if ek in lemon_index), None)
                
                def get_val_dinamico(hdr_options):
                    for h in hdr_options:
                        h_up = h.upper()
                        if h_up in col_map:
                            return safe_get_num(row, col_map[h_up])
                    return 0.0

                a_vista = get_val_dinamico(["A vista", "A VISTA"])
                a_credito = get_val_dinamico(["A credito", "A CREDITO", "A Crédito"])
                tarifa_emitida = a_vista + a_credito
                taxa_emitida = get_val_dinamico(["Taxa", "TAXA"])
                comissao = get_val_dinamico(["Comissão", "COMISSÃO", "Comissao"])
                taxa_du = get_val_dinamico(["Taxa DU", "TAXA DU"])
                desc = get_val_dinamico(["Desc.", "DESC."])
                incentivo = get_val_dinamico(["Incentivo", "INCENTIVO"])
                receita_emitida = comissao + taxa_du + incentivo
                vl_liquido = get_val_dinamico(["VL. Líquido", "VL. LÍQUIDO", "VL. Liquido"])
                
                # DEFINIÇÃO DOS NOVOS CAMPOS EXCLUSIVOS
                sv_final = l_match["Lemon_SV"] if l_match and l_match.get("Lemon_SV") != "-" else (hist_data.get("SV") or "-")
                
                if b_match and b_match.get("Benner_FOP") and b_match["Benner_FOP"] != "-":
                    fop_final = b_match["Benner_FOP"]
                elif l_match and l_match.get("Lemon_Forma_Pagto") and l_match["Lemon_Forma_Pagto"] != "-":
                    fop_final = l_match["Lemon_Forma_Pagto"]
                elif pagto and pagto != "-":
                    fop_final = pagto
                else:
                    fop_final = hist_data.get("FOP") or "-"
                
                # CLIENTE OFICIAL
                if b_match and b_match.get("Benner_Cliente") and b_match["Benner_Cliente"] != "-":
                    cliente_final = b_match["Benner_Cliente"]
                elif l_match and l_match.get("Lemon_Solicitante") and l_match["Lemon_Solicitante"] != "-":
                    cliente_final = l_match["Lemon_Solicitante"]
                else:
                    cliente_final = ponto_venda_final
                
                if b_match:
                    status_sistema = b_match["Benner_Situação"]
                    fornec_sistema = b_match["Benner_Fornecedor"]
                    loc_sistema = b_match["Benner_Localizador"]
                    rloc_cia_sistema = b_match["Benner_Rloc_Cia"]
                    tarifa_sistema = b_match["Benner_Tarifa"]
                    taxa_sistema = b_match["Benner_Taxas"]
                    receita_sistema = b_match["Benner_Receita"]
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
                    loc_sistema = loc_reserva or col3
                    rloc_cia_sistema = (
                        col3 if (len(col3) <= 8 and not col3.isdigit() and col3.upper() not in ["[]", "NAN"])
                        else "-"
                    )
                    tarifa_sistema, taxa_sistema, receita_sistema = 0.0, 0.0, 0.0
                    emissor_sistema, sist_reserva = "-", "-"
                    status_cia = "Pendente"
                    dif_tarifa, dif_taxa, dif_receita = tarifa_emitida, taxa_emitida, receita_emitida
                    status_divergencia = "Pendente de Lançamento"
                    status_geral = hist_data.get("Status_Geral") or "Pendente de Lançamento (Não Consta)"
                    aba_destino = "99_Base_Divergencias_Geral"
                
                fonte_rel = identificar_fonte_relatorio(nome_fonte, emissor_sistema)
                
                rec = {
                    "Ponto de venda": ponto_venda_final,
                    "Código Iata": iata_final,
                    "SV": sv_final,
                    "FOP": fop_final,
                    "Gerentes": gerente_resp,
                    "CIA": col1,
                    "Fornecedor_Sistema": fornec_sistema,
                    "Status_Cia": status_cia,
                    "Bilhetes": clean_str_strict(bilhete_chave),
                    "Localizador_Sistema": clean_str_strict(loc_sistema),
                    "Rloc_Cia": clean_str_strict(rloc_cia_sistema),
                    "Rloc_CIA": clean_str_strict(rloc_cia_sistema),
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
                    "Emissor": emissor_sistema if b_match else (l_match["Lemon_Emissor_Reserva"] if l_match else "-"),
                    "Sistema Reserva": sist_reserva if b_match else (l_match["Lemon_Source"] if l_match else "-"),
                    "Cliente": cliente_final,
                    "Setor": setor_final,
                    "Aba_Destino": aba_destino,
                    "Fonte_Relatorio": fonte_rel,
                    "Tipo_Emissao_Lemon": l_match["Lemon_OnOff"] if l_match else "-",
                    "Consultor_Lemon": l_match["Lemon_Consultor"] if l_match else "-",
                    "Emissor_Reserva_Lemon": l_match["Lemon_Emissor_Reserva"] if l_match else "-",
                    "Centro_Custo_Lemon": l_match["Lemon_Centro_Custo"] if l_match else "-",
                    "Forma_Pagto_Lemon": l_match["Lemon_Forma_Pagto"] if l_match else "-",
                    "Autorizacao_Cartao_Lemon": l_match["Lemon_Autorizacao_Cartao"] if l_match else "-",
                    "Modalidade_Lemon": l_match["Lemon_Modalidade"] if l_match else "-",
                    "Emissao_Online_Lemon": l_match["Lemon_Emissao_Online"] if l_match else "-",
                    "Source_Lemon": l_match["Lemon_Source"] if l_match else "-",
                    "Ponto_Venda_Lemon": l_match["Lemon_Ponto_Venda"] if l_match else "-",
                }
                registros_conciliados.append(rec)

    if not registros_conciliados:
        print("\n⚠️ NENHUM REGISTRO DE EMISSÃO FOI PROCESSADO!\n")
        return

    df_master = pd.DataFrame(registros_conciliados)
    for col_str in ["Bilhetes", "Localizador_Sistema", "Data Emissão", "CIA", "Ponto de venda"]:
        if col_str in df_master.columns:
            df_master[col_str] = df_master[col_str].astype(str).str.strip()
            
    df_master = df_master.drop_duplicates(
        subset=["Bilhetes", "Data Emissão", "A vista", "A credito", "Taxa", "Comissão"],
        keep="first"
    )

    # Exportação e backup em Excel
    print("\n[6/6] Exportando Dashboard Formatado em Excel...")
    with pd.ExcelWriter(out_file_model, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="99_Base_Divergencias_Geral", index=False)
        if not df_log_acumulado.empty:
            df_log_acumulado.to_excel(writer, sheet_name="00_Log_Auditoria", index=False)

    salvar_csv_atomico(
        df_master[["Bilhetes", "Status_Geral", "Área Resp. Operação", "Obs. Operação", "Setor", "Ponto de venda", "Código Iata", "SV", "FOP", "Rloc_Cia"]].drop_duplicates(subset=["Bilhetes"]),
        "Historico_Tratativas.csv"
    )

    # Sincronização direta com a tb_conciliacao_base no Supabase
    sincronizar_supabase_tudo(df_master)

    print("\n" + "=" * 75)
    print(f" AUDITORIA E SINCRONIZAÇÃO SUPABASE CONCLUÍDAS COM SUCESSO! ".center(75, "="))
    print("=" * 75)

if __name__ == "__main__":
    executar_auditoria()