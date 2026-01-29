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

        minutos_totais = hora_ocorrencia * 60 + tempo_chegada_minutos
        hora = (minutos_totais // 60) % 24
        minuto = minutos_totais % 60

        return f"{hora:02d}:{minuto:02d}"
    except Exception as e:
        print("Erro ao calcular hora de chegada:", e)
        return None


def gerar_relatorio(pasta_destino: str, mes_ref: str = None):
    """
    Gera relatório de APREENSAO-DE-ANIMAIS.

    Args:
        pasta_destino: Caminho base onde salvar o relatório (ex: "Excel/2026-01")
        mes_ref: Mês de referência no formato YYYY-MM (opcional)

    Returns:
        Path do arquivo gerado
    """

    # Converter para Path
    pasta_destino = Path(pasta_destino)

    # Criar subpasta específica do relatório
    pasta_saida = pasta_destino / "A-OPERACAO" / "APREENSAO-DE-ANIMAIS"
    pasta_saida.mkdir(parents=True, exist_ok=True)

    # Caminho do arquivo de entrada
    # Ajuste o caminho abaixo se o arquivo estiver em outra pasta (ex: "SQL/STG_ATENDIMENTOS_KCOR.xlsx")
    URL_ATENDIMENTOS = Path("SQL\\STG_ATENDIMENTOS_KCOR.xlsx")

    # Nome do arquivo de saída
    if mes_ref:
        OUTPUT_XLSX = pasta_saida / f"APREENSAO-DE-ANIMAIS_{mes_ref}.xlsx"
    else:
        OUTPUT_XLSX = pasta_saida / "APREENSAO-DE-ANIMAIS.xlsx"

    print("\n🐾 Gerando: APREENSAO-DE-ANIMAIS")
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

    # 3) Filtros específicos de ocorrências com animais
    filtro_animais = [
        "Animal Doméstico Morto",
        "Animal Silvestre Morto",
        "Animal Solto",
        "Animal Apreendido",
        "Animal de Grande Porte Morto",
        "Animal Resgatado",
    ]

    tipos_atendimento_excluir = [
        "Cancelado(QTA)",
        "Não Localizado",
    ]

    df_filtrado = df_origem[df_origem["DESCROCORRENCIA"].isin(filtro_animais)].copy()
    df_filtrado = df_filtrado[
        ~df_filtrado["DSC_TIPO_ATENDIMENTO"].isin(tipos_atendimento_excluir)
    ].copy()

    print(f"  → Registros após filtro de animais: {len(df_filtrado)}")

    if len(df_filtrado) == 0:
        print("  ⚠️ Nenhum registro de apreensão/ocorrência com animais encontrado para os filtros aplicados.")
        df_final_vazio = pd.DataFrame(
            columns=[
                "Ocorrência",
                "Data",
                "Rodovia",
                "KM",
                "HM",
                "Tipo de Animal",
                "Hora de Acionamento",
                "Hora de Chegada",
                "Tempo de Chegada",
                "Destino",
            ]
        )
        df_final_vazio.to_excel(OUTPUT_XLSX, index=False)
        print("  ✅ Arquivo vazio gerado (sem registros).")
        return OUTPUT_XLSX

    # 4) Monta o DataFrame final no padrão do relatório
    df_final = pd.DataFrame()

    df_final["Ocorrência"] = df_filtrado["NUMOCORRENCIA"]
    df_final["Data"] = df_filtrado["DATAOCORRENCIA"]
    df_final["Rodovia"] = df_filtrado["RODOVIA"]
    df_final["KM"] = df_filtrado["KM"]
    df_final["HM"] = df_filtrado["HORAOCORRENCIA"]

    # Não temos essa informação -> null
    df_final["Tipo de Animal"] = None

    df_final["Hora de Acionamento"] = df_filtrado["HORAOCORRENCIA"]

    df_final["Hora de Chegada"] = df_filtrado.apply(
        lambda row: calcular_hora_chegada(
            row.get("HORAOCORRENCIA"), row.get("TEMPOCHEGADA")
        ),
        axis=1,
    )

    df_final["Tempo de Chegada"] = df_filtrado["TEMPOCHEGADA"]

    # Não temos essa informação -> null
    df_final["Destino"] = None

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
