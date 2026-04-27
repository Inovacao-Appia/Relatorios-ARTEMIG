import os
from datetime import datetime, date, timedelta
from calendar import monthrange

import streamlit as st
import pandas as pd
import oracledb
from dotenv import load_dotenv

# ── Carrega variaveis de ambiente ─────────────────────────────────────────────
load_dotenv()

DB_HOST    = os.getenv("DB_HOST", "oracluster")
DB_PORT    = int(os.getenv("DB_PORT", "1521"))
DB_SERVICE = os.getenv("DB_SERVICE", "srv_bi")
DB_USER    = os.getenv("DB_USER")
DB_PASS    = os.getenv("DB_PASSWORD")

# ── Modo THICK ────────────────────────────────────────────────────────────────
try:
    oracledb.init_oracle_client()
except Exception:
    pass

# ── Constantes ────────────────────────────────────────────────────────────────
CONCESSIONARIA = "Via Nascentes"

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

COLUNAS_ESCOPO = [
    "CnpjConcessionaria",
    "IdPassagem",
    "CodigoPracaPedagio",
    "CodigoCabine",
    "DataHoraPassagem",
    "Faixa",
    "Sentido",
    "TipoVeiculo",
    "QuantidadeEixosVeiculo",
    "CategoriaVeiculo",
    "Rodagem",
    "EixoSuspenso",
    "QuantidadeEixosSuspensos",
    "Isento",
    "TipoCobrancaEfetuada",
    "Placa",
    "ValorDevido",
    "ValorArrecadado",
    "IdTag",
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def safe(val, default="") -> str:
    try:
        if pd.isna(val):
            return default
    except Exception:
        pass
    return str(val).strip()

def safe_float(val, default=0.0) -> float:
    try:
        if pd.isna(val):
            return default
        return float(val)
    except Exception:
        return default

def safe_int(val, default=0) -> int:
    try:
        if pd.isna(val):
            return default
        return int(float(val))
    except Exception:
        return default

def formatar_datahora(data_val, hora_val) -> str:
    try:
        if isinstance(data_val, datetime): dt = data_val.date()
        elif isinstance(data_val, date):   dt = data_val
        else: dt = datetime.strptime(str(data_val).strip()[:10], "%Y-%m-%d").date()
    except Exception:
        return ""
    try:
        h = str(int(hora_val)).zfill(6)
        hs = "{}:{}:{}".format(h[0:2], h[2:4], h[4:6])
    except Exception:
        hs = "00:00:00"
    return "{}-{}".format(dt.strftime("%d/%m/%Y"), hs)

def fmt_valor(val: float) -> str:
    try:
        return "{:,.2f}".format(val).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"

# ── Mapeamentos ───────────────────────────────────────────────────────────────
def map_tipo_veiculo(qtd_passeio, qtd_comercial, qtd_moto) -> str:
    if safe_int(qtd_moto)      > 0: return "Motocicleta"
    if safe_int(qtd_comercial) > 0: return "Comercial"
    if safe_int(qtd_passeio)   > 0: return "Passeio"
    return ""

def map_rodagem(dac_posder) -> str:
    return "Dupla" if safe_int(dac_posder) > 1 else "Simples"

def map_tipo_cobranca(subforma: str) -> str:
    s = str(subforma).strip().upper()
    if "TAG"   in s: return "Automatica TAG"
    if "OCR"   in s or "PLACA" in s: return "Automatica OCR/PLACA"
    if "MAN"   in s: return "Manual"
    return safe(subforma)

def map_sentido(dsc_sentido: str) -> str:
    s = str(dsc_sentido).strip().upper()
    if "CRESC" in s or "LESTE" in s or "NORTE" in s: return "Crescente"
    if "DECRESC" in s or "OESTE" in s or "SUL"  in s: return "Decrescente"
    return safe(dsc_sentido)

# ── Conexao Oracle ────────────────────────────────────────────────────────────
def get_connection():
    dsn = oracledb.makedsn(DB_HOST, DB_PORT, service_name=DB_SERVICE)
    return oracledb.connect(user=DB_USER, password=DB_PASS, dsn=dsn)

# ── Queries por dia ───────────────────────────────────────────────────────────
QUERY_TRAFEGO_DIA = """
    SELECT
        f.COD_ID_TRANSACTION,
        f.SK_VALIDADOR,
        f.NUM_HORA_PASSAGEM,
        f.DAC_POSDER,
        f.DAC_POSIZQ,
        f.QTD_PASSEIO,
        f.QTD_COMERCIAL,
        f.QTD_MOTO,
        f.DAC_ADICION,
        f.DAC_CLASSE,
        f.COD_CAT_VEIC,
        f.COD_CLASSE,
        f.NUM_OTICO_NUM_EIXOS_SUSP,
        f.DSC_SUBFORMA_PGTO,
        f.SK_FORMA_PAGAMENTO,
        f.VAL_VALOR_ARREC,
        d.DAT_REF_FORMATADA,
        p.COD_PRACA,
        p.DSC_PRACA,
        p.DSC_SENTIDO,
        p.COD_VIA
    FROM DW.FAT_ARRECADACAO_TRAFEGO f
    LEFT JOIN DW.DIM_DAT_PERIODO d
        ON f.SK_DAT_PERIODO = d.SK_DAT_PERIODO
    LEFT JOIN DW.DIM_EMPRESA_PRACA_VIA p
        ON f.SK_EMPRESA_PRACA_VIA = p.SK_EMPRESA_PRACA_VIA
    WHERE p.DSC_EMPRESA = 'NASCENTES DAS GERAIS'
      AND d.DAT_REF_FORMATADA >= TO_DATE(:dt_dia,      'YYYY-MM-DD')
      AND d.DAT_REF_FORMATADA <  TO_DATE(:dt_dia_prox, 'YYYY-MM-DD')
    ORDER BY f.NUM_HORA_PASSAGEM
"""

QUERY_EVASOES_DIA = """
    SELECT
        ev.IDTRANSACTION,
        ev.PLACA,
        f.SK_VALIDADOR,
        f.NUM_HORA_PASSAGEM,
        f.DAC_POSDER,
        f.DAC_POSIZQ,
        f.QTD_PASSEIO,
        f.QTD_COMERCIAL,
        f.QTD_MOTO,
        f.DAC_ADICION,
        f.DAC_CLASSE,
        f.COD_CAT_VEIC,
        f.COD_CLASSE,
        f.NUM_OTICO_NUM_EIXOS_SUSP,
        f.DSC_SUBFORMA_PGTO,
        f.SK_FORMA_PAGAMENTO,
        f.COD_ID_TRANSACTION,
        d.DAT_REF_FORMATADA,
        p.COD_PRACA,
        p.DSC_PRACA,
        p.DSC_SENTIDO
    FROM DW.FAT_EVASOES ev
    LEFT JOIN DW.FAT_ARRECADACAO_TRAFEGO f
        ON TO_CHAR(f.COD_ID_TRANSACTION) = TO_CHAR(ev.IDTRANSACTION)
    LEFT JOIN DW.DIM_DAT_PERIODO d
        ON f.SK_DAT_PERIODO = d.SK_DAT_PERIODO
    LEFT JOIN DW.DIM_EMPRESA_PRACA_VIA p
        ON f.SK_EMPRESA_PRACA_VIA = p.SK_EMPRESA_PRACA_VIA
    WHERE p.DSC_EMPRESA = 'NASCENTES DAS GERAIS'
      AND d.DAT_REF_FORMATADA >= TO_DATE(:dt_dia,      'YYYY-MM-DD')
      AND d.DAT_REF_FORMATADA <  TO_DATE(:dt_dia_prox, 'YYYY-MM-DD')
    ORDER BY f.NUM_HORA_PASSAGEM
"""

# ── Transformacao — trafego ───────────────────────────────────────────────────
def transformar_trafego(df: pd.DataFrame, tarifas: dict) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        val_arrec  = safe_float(r.get("VAL_VALOR_ARREC"))
        eixos_susp = safe_int(r.get("NUM_OTICO_NUM_EIXOS_SUSP"))
        eixos_tot  = safe_int(r.get("DAC_ADICION")) + safe_int(r.get("DAC_CLASSE"))
        rows.append({
            "CnpjConcessionaria":       CONCESSIONARIA,
            "IdPassagem":               safe(r.get("COD_ID_TRANSACTION")),
            "CodigoPracaPedagio":       safe(r.get("COD_PRACA")),
            "CodigoCabine":             safe(r.get("SK_VALIDADOR")),
            "DataHoraPassagem":         formatar_datahora(r.get("DAT_REF_FORMATADA"), r.get("NUM_HORA_PASSAGEM")),
            "Faixa":                    safe(r.get("DAC_POSDER")),
            "Sentido":                  map_sentido(safe(r.get("DSC_SENTIDO"))),
            "TipoVeiculo":              map_tipo_veiculo(r.get("QTD_PASSEIO"), r.get("QTD_COMERCIAL"), r.get("QTD_MOTO")),
            "QuantidadeEixosVeiculo":   eixos_tot if eixos_tot > 0 else "",
            "CategoriaVeiculo":         safe(r.get("COD_CAT_VEIC")),
            "Rodagem":                  map_rodagem(r.get("DAC_POSDER")),
            "EixoSuspenso":             "Sim" if eixos_susp > 0 else "Nao",
            "QuantidadeEixosSuspensos": eixos_susp if eixos_susp > 0 else "",
            "Isento":                   "Sim" if val_arrec == 0 else "Nao",
            "TipoCobrancaEfetuada":     map_tipo_cobranca(safe(r.get("DSC_SUBFORMA_PGTO"))),
            "Placa":                    "",
            "ValorDevido":              fmt_valor(val_arrec),
            "ValorArrecadado":          fmt_valor(val_arrec),
            "IdTag":                    safe(r.get("SK_FORMA_PAGAMENTO")),
        })
    return pd.DataFrame(rows, columns=COLUNAS_ESCOPO)

# ── Transformacao — evasoes ───────────────────────────────────────────────────
def transformar_evasoes(df: pd.DataFrame, tarifas: dict) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        cod_praca  = safe(r.get("COD_PRACA"))
        tarifa_ref = tarifas.get(cod_praca, 0.0)
        eixos_susp = safe_int(r.get("NUM_OTICO_NUM_EIXOS_SUSP"))
        eixos_tot  = safe_int(r.get("DAC_ADICION")) + safe_int(r.get("DAC_CLASSE"))
        rows.append({
            "CnpjConcessionaria":       CONCESSIONARIA,
            "IdPassagem":               safe(r.get("IDTRANSACTION")),
            "CodigoPracaPedagio":       cod_praca,
            "CodigoCabine":             safe(r.get("SK_VALIDADOR")),
            "DataHoraPassagem":         formatar_datahora(r.get("DAT_REF_FORMATADA"), r.get("NUM_HORA_PASSAGEM")),
            "Faixa":                    safe(r.get("DAC_POSDER")),
            "Sentido":                  map_sentido(safe(r.get("DSC_SENTIDO"))),
            "TipoVeiculo":              map_tipo_veiculo(r.get("QTD_PASSEIO"), r.get("QTD_COMERCIAL"), r.get("QTD_MOTO")),
            "QuantidadeEixosVeiculo":   eixos_tot if eixos_tot > 0 else "",
            "CategoriaVeiculo":         safe(r.get("COD_CAT_VEIC")),
            "Rodagem":                  map_rodagem(r.get("DAC_POSDER")),
            "EixoSuspenso":             "Sim" if eixos_susp > 0 else "Nao",
            "QuantidadeEixosSuspensos": eixos_susp if eixos_susp > 0 else "",
            "Isento":                   "Nao",
            "TipoCobrancaEfetuada":     "Evasao",
            "Placa":                    safe(r.get("PLACA")),
            "ValorDevido":              fmt_valor(tarifa_ref),
            "ValorArrecadado":          fmt_valor(0.0),
            "IdTag":                    "",
        })
    return pd.DataFrame(rows, columns=COLUNAS_ESCOPO)

# ── Caminho do CSV de saida ───────────────────────────────────────────────────
def caminho_csv(pasta_base: str, mes: int, ano: int) -> str:
    pasta = os.path.join(
        pasta_base, "A-OPERACAO", "trafego_praca_pedagio",
        str(ano), "{}_{}".format(str(mes).zfill(2), MESES_PT[mes])
    )
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, "{}_{}_Trafego_Praca_Pedagio.csv".format(ano, str(mes).zfill(2)))

# ── Processamento paginado por dia — escrita em CSV ───────────────────────────
def gerar_por_dia(mes: int, ano: int, pasta_base: str):
    ultimo_dia  = monthrange(ano, mes)[1]
    dias        = [date(ano, mes, d) for d in range(1, ultimo_dia + 1)]
    total_dias  = len(dias)
    caminho     = caminho_csv(pasta_base, mes, ano)

    tarifas    = {}
    total_pass = 0
    total_ev   = 0
    total_perd = 0.0
    dias_ok    = []
    primeiro   = True   # controla cabecalho do CSV

    # Limpa arquivo anterior se existir
    if os.path.exists(caminho):
        os.remove(caminho)

    conn = get_connection()

    barra        = st.progress(0, text="Iniciando...")
    status_tx    = st.empty()
    col1, col2, col3 = st.columns(3)
    m_pass  = col1.empty()
    m_ev    = col2.empty()
    m_perd  = col3.empty()

    dia_str = ""
    try:
        for idx, dia in enumerate(dias):
            dia_str      = dia.strftime("%Y-%m-%d")
            dia_prox_str = (dia + timedelta(days=1)).strftime("%Y-%m-%d")
            params       = {"dt_dia": dia_str, "dt_dia_prox": dia_prox_str}

            status_tx.info(
                "📅  Processando **{}/{}** ({}/{})...".format(
                    dia.strftime("%d/%m/%Y"), ano, idx + 1, total_dias
                )
            )

            # ── Trafego do dia ────────────────────────────────────────────
            df_t = pd.read_sql(QUERY_TRAFEGO_DIA, conn, params=params)

            # Atualiza tarifas de referencia
            df_pago = df_t[df_t["VAL_VALOR_ARREC"] > 0]
            for praca, grupo in df_pago.groupby("COD_PRACA"):
                chave = str(praca)
                if chave not in tarifas:
                    tarifas[chave] = round(grupo["VAL_VALOR_ARREC"].median(), 2)

            # ── Evasoes do dia ────────────────────────────────────────────
            df_e = pd.read_sql(QUERY_EVASOES_DIA, conn, params=params)

            # ── Transforma ────────────────────────────────────────────────
            frames = []
            if not df_t.empty:
                frames.append(transformar_trafego(df_t, tarifas))
                total_pass += len(df_t)
            if not df_e.empty:
                frames.append(transformar_evasoes(df_e, tarifas))
                perdido_dia = sum(
                    tarifas.get(safe(r.get("COD_PRACA")), 0.0)
                    for _, r in df_e.iterrows()
                )
                total_ev   += len(df_e)
                total_perd += perdido_dia

            # ── Salva no CSV (append) — instantaneo ───────────────────────
            if frames:
                df_dia = pd.concat(frames, ignore_index=True)
                df_dia.to_csv(
                    caminho,
                    mode="w" if primeiro else "a",
                    header=primeiro,
                    index=False,
                    sep=";",
                    encoding="utf-8-sig",  # compativel com Excel ao abrir
                )
                primeiro = False

            dias_ok.append(dia.strftime("%d/%m"))

            # ── Atualiza UI ───────────────────────────────────────────────
            pct = int(((idx + 1) / total_dias) * 100)
            barra.progress(pct, text="{}% — Dia {}/{} concluido".format(pct, idx + 1, total_dias))
            m_pass.metric("Passagens",     "{:,}".format(total_pass))
            m_ev.metric("Evasoes",         "{:,}".format(total_ev))
            m_perd.metric("Valor Perdido", "R$ {}".format(fmt_valor(total_perd)))

    except Exception as e:
        st.error("Erro no dia {}: {}".format(dia_str, e))
        if dias_ok:
            st.warning(
                "⚠️  Processamento interrompido. "
                "CSV parcial com os dias {} salvo em: {}".format(", ".join(dias_ok), caminho)
            )
    finally:
        conn.close()

    return caminho, total_pass, total_ev, total_perd

# ── Interface Streamlit ───────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Relatorios Nascentes - Trafego",
        page_icon="🛣️",
        layout="centered",
    )

    st.title("Relatorios Nascentes")
    st.subheader("A-OPERACAO - Trafego Praca de Pedagio")
    st.caption(
        "Extracao paginada por dia. Cada dia e salvo imediatamente no CSV. "
        "Se interrompido, o arquivo parcial ja esta disponivel."
    )
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox(
            "Mes do relatorio",
            options=list(MESES_PT.keys()),
            format_func=lambda m: MESES_PT[m],
            index=datetime.now().month - 1,
        )
    with col2:
        ano_atual = datetime.now().year
        ano = st.number_input(
            "Ano do relatorio",
            min_value=2020, max_value=ano_atual + 1,
            value=ano_atual, step=1,
        )

    st.divider()
    st.info(
        "ℹ️  Saida em CSV (separador ;) | Empresa: NASCENTES DAS GERAIS | "
        "Periodo: {}/{} | {} dias a processar.".format(
            str(mes).zfill(2), ano, monthrange(ano, mes)[1]
        )
    )
    st.divider()

    pasta_base = st.text_input(
        "Pasta de destino",
        value=os.path.join(os.path.expanduser("~"), "Relatorios Nascentes"),
        help="O CSV sera salvo e atualizado a cada dia processado nessa pasta."
    )
    st.divider()

    if st.button("Gerar Relatorio", type="primary", use_container_width=True):
        if not pasta_base:
            st.error("Informe a pasta de destino.")
            return

        caminho, total_pass, total_ev, total_perd = gerar_por_dia(mes, ano, pasta_base)

        if os.path.exists(caminho):
            st.success(
                "✅  Concluido! {} passagens | {} evasoes | R$ {} perdido.".format(
                    "{:,}".format(total_pass),
                    "{:,}".format(total_ev),
                    fmt_valor(total_perd),
                )
            )
            st.success("Arquivo salvo em: {}".format(caminho))

            with open(caminho, "rb") as f:
                st.download_button(
                    label="📥  Baixar CSV",
                    data=f.read(),
                    file_name="{}_{}_Trafego_Praca_Pedagio.csv".format(ano, str(mes).zfill(2)),
                    mime="text/csv",
                    use_container_width=True,
                )

if __name__ == "__main__":
    main()
