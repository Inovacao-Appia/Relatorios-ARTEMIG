import pandas as pd
from pathlib import Path


def calcular_hora_chegada(hora_ocorrencia, tempo_chegada_minutos):
    """
    Calcula a hora de chegada a partir da hora da ocorrência + tempo de chegada (em minutos).

    Args:
        hora_ocorrencia: inteiro ou string representando hora cheia (0 a 23)
        tempo_chegada_minutos: inteiro ou string representando minutos

    Returns:
        string no formato HH:MM (24h) ou None em caso de erro
    """
    try:
        hora_ocorrencia = int(hora_ocorrencia)
        tempo_chegada_minutos = int(tempo_chegada_minutos)

        # Converte para minutos totais e depois volta para HH:MM (ciclo de 24h)
        minutos_totais = hora_ocorrencia * 60 + tempo_chegada_minutos
        hora = (minutos_totais // 60) % 24
        minuto = minutos_totais % 60

        return f"{hora:02d}:{minuto:02d}"
    except Exception as e:
        print("Erro ao calcular hora de chegada:", e)
        return None


def gerar_relatorio(pasta_destino: str, mes_ref: str = None):
    """
    Gera relatório de COMBATE-A-INCENDIO.

    Args:
        pasta_destino: Caminho base onde salvar o relatório (ex: "Excel/2026-01")
        mes_ref: Mês de referência no formato YYYY-MM (opcional)

    Returns:
        Path do arquivo gerado
    """

    # Converter para Path
    pasta_destino = Path(pasta_destino)

    # Criar subpasta específica do relatório
    pasta_saida = pasta_destino / "A-OPERACAO" / "COMBATE-A-INCENDIO"
    pasta_saida.mkdir(parents=True, exist_ok=True)

    # Caminho do arquivo de entrada
    # Ajuste o caminho abaixo se o arquivo estiver em outra pasta (ex: "SQL/STG_ATENDIMENTOS_KCOR.xlsx")
    URL_ATENDIMENTOS = Path("SQL\\STG_ATENDIMENTOS_KCOR.xlsx")

    # Nome do arquivo de saída
    if mes_ref:
        OUTPUT_XLSX = pasta_saida / f"COMBATE-A-INCENDIO_{mes_ref}.xlsx"
    else:
        OUTPUT_XLSX = pasta_saida / "COMBATE-A-INCENDIO.xlsx"

    print("\n🔥 Gerando: COMBATE-A-INCENDIO")
    print(f"📂 Destino: {OUTPUT_XLSX}")

    # 1) Ler base de origem
    print("  → Lendo planilha de atendimentos...")
    df_origem = pd.read_excel(URL_ATENDIMENTOS)
    print(f"    • Registros originais: {len(df_origem)}")

    # 2) Filtrar por mês, se fornecido
    if mes_ref and "DATAOCORRENCIA" in df_origem.columns:
        df_origem["DATAOCORRENCIA"] = pd.to_datetime(
            df_origem["DATAOCORRENCIA"], errors="coerce"
        )
        df_origem["ano_mes"] = df_origem["DATAOCORRENCIA"].dt.strftime("%Y-%m")

        df_mes = df_origem[df_origem["ano_mes"] == mes_ref].copy()
        print(f"  → Filtrando para {mes_ref}: {len(df_mes)} registros após filtro de mês")

        if len(df_mes) > 0:
            df_origem = df_mes

    # 3) Filtros específicos de incêndio
    # Lista de tipos de ocorrência de incêndio (DESCROCORRENCIA)
    filtro_incendio = [
        "Incêndio de Grande Proporção",
        "Incêndio de Média Proporção",
        "Incêndio de Pequena Proporção",
    ]

    # Lista de tipos de atendimento a excluir (DSC_TIPO_ATENDIMENTO)
    tipos_atendimento_excluir = [
        "Cancelado(QTA)",
        "Não Localizado",
    ]

    # Aplica o filtro de incêndio
    df_filtrado = df_origem[
        df_origem["DESCROCORRENCIA"].isin(filtro_incendio)
    ].copy()

    # Exclui certos tipos de atendimento
    df_filtrado = df_filtrado[
        ~df_filtrado["DSC_TIPO_ATENDIMENTO"].isin(tipos_atendimento_excluir)
    ].copy()

    print(f"  → Registros após filtro de incêndio: {len(df_filtrado)}")

    if len(df_filtrado) == 0:
        print("  ⚠️ Nenhum registro de incêndio encontrado para os filtros aplicados.")
        # Ainda assim gera uma planilha vazia com colunas no padrão
        df_final_vazio = pd.DataFrame(
            columns=[
                "Ocorrência",
                "Data",
                "Rodovia",
                "Km Inicial",
                "HM Inicial",
                "Km Final",
                "HM Final",
                "Área atingida (ha)",
                "Hora de Acionamento",
                "Hora de Chegada",
                "Tempo de Chegada",
                "Recursos Mobilizados",
                "Causas Prováveis",
            ]
        )
        df_final_vazio.to_excel(OUTPUT_XLSX, index=False)
        print("  ✅ Arquivo vazio gerado (sem registros).")
        return OUTPUT_XLSX

    # 4) Monta o DataFrame final no padrão da planilha de combate a incêndio
    df_final = pd.DataFrame()

    # De-para conforme o código original
    df_final["Ocorrência"] = df_filtrado["KOCORRENCIA"]
    df_final["Data"] = df_filtrado["DATAOCORRENCIA"]
    df_final["Rodovia"] = df_filtrado["RODOVIA"]
    df_final["Km Inicial"] = df_filtrado["KM"]
    df_final["HM Inicial"] = df_filtrado["HORAOCORRENCIA"]

    # Km Final = mesmo valor do Km Inicial (replicado)
    df_final["Km Final"] = df_filtrado["KM"]

    # HM Final = HORAOCORRENCIA + TEMPOCHEGADA
    df_final["HM Final"] = df_filtrado.apply(
        lambda row: calcular_hora_chegada(
            row.get("HORAOCORRENCIA"), row.get("TEMPOCHEGADA")
        ),
        axis=1,
    )

    # Área atingida = null (não temos essa informação)
    df_final["Área atingida (ha)"] = None

    df_final["Hora de Acionamento"] = df_filtrado["HORAOCORRENCIA"]

    # Hora de Chegada = HORAOCORRENCIA + TEMPOCHEGADA
    df_final["Hora de Chegada"] = df_filtrado.apply(
        lambda row: calcular_hora_chegada(
            row.get("HORAOCORRENCIA"), row.get("TEMPOCHEGADA")
        ),
        axis=1,
    )

    df_final["Tempo de Chegada"] = df_filtrado["TEMPOCHEGADA"]
    df_final["Recursos Mobilizados"] = df_filtrado["TIPO"]

    # Causas Prováveis = null (não temos essa informação)
    df_final["Causas Prováveis"] = None

    # Garante índice sequencial
    df_final = df_final.reset_index(drop=True)

    # 5) Salvar arquivo
    df_final.to_excel(OUTPUT_XLSX, index=False)
    print(f"  ✅ Salvo: {len(df_final)} registros")

    return OUTPUT_XLSX


# Execução direta (teste local)
if __name__ == "__main__":
    import sys

    pasta = sys.argv[1] if len(sys.argv) > 1 else "Excel"
    mes = sys.argv[2] if len(sys.argv) > 2 else None
    gerar_relatorio(pasta, mes)
