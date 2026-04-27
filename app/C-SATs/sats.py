import os
import calendar
import zipfile
from datetime import date, datetime
from io import BytesIO

import streamlit as st
import pandas as pd
import pyodbc
from dotenv import load_dotenv

# ── Carrega variáveis de ambiente ─────────────────────────────────────────────
load_dotenv()

SQL_HOST = os.getenv("SQL_HOST", "192.168.40.22")
SQL_PORT = os.getenv("SQL_PORT", "1433")
SQL_DB   = os.getenv("SQL_DB",   "AT3_db_2")
SQL_USER = os.getenv("SQL_USER", "sa")
SQL_PASS = os.getenv("SQL_PASS", "$4dmp4dr40@nasc")

CONNECTION_STRING = (
    "DRIVER={{SQL Server}};"
    "SERVER={host},{port};"
    "DATABASE={db};"
    "UID={user};"
    "PWD={pwd};"
).format(host=SQL_HOST, port=SQL_PORT, db=SQL_DB, user=SQL_USER, pwd=SQL_PASS)

# ── Constantes ────────────────────────────────────────────────────────────────
MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

TABELAS = [
    "AT3_1_Amostra",   "AT3_1_Ocorrencia",
    "AT3_2_Amostra",   "AT3_2_Ocorrencia",
    "AT3_3_Amostra",   "AT3_3_Ocorrencia",
    "AT3_4_Amostra",   "AT3_4_Ocorrencia",
    "AT3_5_Amostra",   "AT3_5_Ocorrencia",
    "AT3_6_Amostra",   "AT3_6_Ocorrencia",
    "AT3_7_Amostra",   "AT3_7_Ocorrencia",
    "AT3_8_Amostra",   "AT3_8_Ocorrencia",
    "AT3_9_Amostra",   "AT3_9_Ocorrencia",
    "AT3_10_Amostra",  "AT3_10_Ocorrencia",
    "AT3_11_Amostra",  "AT3_11_Ocorrencia",
    "AT3_12_Amostra",  "AT3_12_Ocorrencia",
    "AT3_13_Amostra",  "AT3_13_Ocorrencia",
    "AT3_14_Amostra",  "AT3_14_Ocorrencia",
    "AT3_15_Amostra",  "AT3_15_Ocorrencia",
    "AT3_16_Amostra",  "AT3_16_Ocorrencia",
    "AT3_17_Amostra",  "AT3_17_Ocorrencia",
    "AT3_18_Amostra",  "AT3_18_Ocorrencia",
    "AT3_19_Amostra",  "AT3_19_Ocorrencia",
    "AT3_20_Amostra",  "AT3_20_Ocorrencia",
    "AT3_37_Amostra",  "AT3_37_Ocorrencia",
    "AT3_99_Amostra",  "AT3_99_Ocorrencia",
    "AT3_231_Amostra", "AT3_231_Ocorrencia",
]

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_connection():
    return pyodbc.connect(CONNECTION_STRING)


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def build_zip(resultados: dict, mes: int, ano: int) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for nome, xlsx_bytes in resultados.items():
            pasta = "AMOSTRAS" if nome.endswith("_Amostra") else "OCORRENCIAS"
            zf.writestr("{}/{}.xlsx".format(pasta, nome), xlsx_bytes)
    return buf.getvalue()


# ── Processamento ─────────────────────────────────────────────────────────────
def gerar_sats(mes: int, ano: int) -> dict:
    """Consulta todas as tabelas AT3 e retorna dict {nome_tabela: bytes_xlsx}."""

    ultimo_dia  = calendar.monthrange(ano, mes)[1]
    data_inicio = date(ano, mes, 1)
    data_fim    = date(ano, mes, ultimo_dia)

    total     = len(TABELAS)
    barra     = st.progress(0, text="Iniciando...")
    status_tx = st.empty()

    col1, col2, col3 = st.columns(3)
    m_amostras    = col1.empty()
    m_ocorrencias = col2.empty()
    m_linhas      = col3.empty()

    qtd_amostras    = 0
    qtd_ocorrencias = 0
    total_linhas    = 0
    resultados      = {}

    try:
        conn = get_connection()

        for i, nome_tabela in enumerate(TABELAS):
            status_tx.info(
                "Processando **{}** ({}/{})...".format(nome_tabela, i + 1, total)
            )

            query = """
                SELECT *
                FROM AT3_db_1.dbo.{tabela}
                WHERE 1=1
                  AND Instante >= '{dt_ini}'
                  AND Instante <= '{dt_fim}'
                ORDER BY Instante DESC
            """.format(tabela=nome_tabela, dt_ini=data_inicio, dt_fim=data_fim)

            df = pd.read_sql(query, conn)
            total_linhas += len(df)

            if nome_tabela.endswith("_Amostra"):
                qtd_amostras += 1
            else:
                qtd_ocorrencias += 1

            resultados[nome_tabela] = df_to_xlsx_bytes(df)

            pct = int(((i + 1) / total) * 100)
            barra.progress(pct, text="{}% — {}/{} tabelas".format(pct, i + 1, total))
            m_amostras.metric("Amostras",      qtd_amostras)
            m_ocorrencias.metric("Ocorrencias", qtd_ocorrencias)
            m_linhas.metric("Linhas totais",   "{:,}".format(total_linhas))

        conn.close()
        status_tx.empty()

        st.success(
            "Concluido! {} amostras | {} ocorrencias | {:,} linhas totais.".format(
                qtd_amostras, qtd_ocorrencias, total_linhas
            )
        )

    except pyodbc.Error as e:
        st.error("Erro de conexao SQL Server: {}".format(e))
    except Exception as e:
        st.error("Erro geral: {}".format(e))

    return resultados


# ── Seção de downloads ────────────────────────────────────────────────────────
def render_downloads(resultados: dict, mes: int, ano: int):
    if not resultados:
        return

    prefixo = "{}_{}_{:04d}".format(str(mes).zfill(2), MESES_PT[mes], ano)

    # Botão ZIP com todos os relatórios
    zip_bytes = build_zip(resultados, mes, ano)
    st.download_button(
        label="Baixar TODOS os relatorios (ZIP)",
        data=zip_bytes,
        file_name="CSATS_{}.zip".format(prefixo),
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )

    st.divider()

    amostras    = {k: v for k, v in resultados.items() if k.endswith("_Amostra")}
    ocorrencias = {k: v for k, v in resultados.items() if not k.endswith("_Amostra")}

    col_a, col_o = st.columns(2)

    with col_a:
        st.markdown("### Amostras")
        for nome, xlsx_bytes in amostras.items():
            st.download_button(
                label="{}".format(nome),
                data=xlsx_bytes,
                file_name="{}_{}.xlsx".format(nome, prefixo),
                mime=MIME_XLSX,
                use_container_width=True,
                key="dl_{}".format(nome),
            )

    with col_o:
        st.markdown("### Ocorrencias")
        for nome, xlsx_bytes in ocorrencias.items():
            st.download_button(
                label="{}".format(nome),
                data=xlsx_bytes,
                file_name="{}_{}.xlsx".format(nome, prefixo),
                mime=MIME_XLSX,
                use_container_width=True,
                key="dl_{}".format(nome),
            )


# ── Interface Streamlit ───────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Relatorios Nascentes - C-SATS",
        page_icon="📊",
        layout="wide",
    )

    st.title("Relatorios Nascentes")
    st.subheader("C-SATS - Extracao de Dados AT3")
    st.caption(
        "Consulta todas as tabelas AT3 (Amostras e Ocorrencias) do SQL Server "
        "e disponibiliza cada arquivo .xlsx para download direto pelo navegador."
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
    st.markdown("### Fonte de dados")
    st.info(
        "**SQL Server** — {}:{} | Database: AT3_db_1 | "
        "{} tabelas a processar | Periodo: {}/{}.".format(
            SQL_HOST, SQL_PORT, len(TABELAS), str(mes).zfill(2), ano
        ),
        icon="ℹ️",
    )
    st.divider()

    if st.button("Gerar Relatorios SATS", type="primary", use_container_width=True):
        resultados = gerar_sats(mes, int(ano))
        if resultados:
            st.session_state["resultados"] = resultados
            st.session_state["mes"] = mes
            st.session_state["ano"] = int(ano)

    if "resultados" in st.session_state:
        st.divider()
        st.markdown("## Downloads")
        render_downloads(
            st.session_state["resultados"],
            st.session_state["mes"],
            st.session_state["ano"],
        )


if __name__ == "__main__":
    main()
