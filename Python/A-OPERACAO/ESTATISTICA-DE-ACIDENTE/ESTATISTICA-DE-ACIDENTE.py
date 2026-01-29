import pandas as pd
from pathlib import Path

def norm_text(s: pd.Series) -> pd.Series:
    """Normaliza texto para comparação (minúsculas, sem espaços)"""
    return (
        s.astype(str)
         .str.strip()
         .str.lower()
         .replace({"nan": ""})
    )

def gerar_relatorio(pasta_destino: str, mes_ref: str = None):
    """
    Gera relatório de estatísticas de acidentes
    
    Args:
        pasta_destino: Caminho onde salvar o relatório (ex: "Excel/2026-01")
        mes_ref: Mês de referência no formato YYYY-MM (opcional)
    
    Returns:
        Path do arquivo gerado
    """
    
    # Converter para Path
    pasta_destino = Path(pasta_destino)
    
    # Criar subpastas se necessário
    pasta_saida = pasta_destino / "A-OPERACAO" / "ESTATISTICA-DE-ACIDENTE"
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    # Caminhos dos arquivos de entrada (fixos)
    URL_ACIDENTES = Path("SQL/STG_KCOR_ACIDENTES2.xlsx")
    URL_VITIMAS = Path("SQL/STG_KCOR_ACIDENTES_VITIMAS.xlsx")
    
    # Nome do arquivo de saída
    if mes_ref:
        OUTPUT_XLSX = pasta_saida / f"ESTATISTICA-DE-ACIDENTE_{mes_ref}.xlsx"
    else:
        OUTPUT_XLSX = pasta_saida / "ESTATISTICA-DE-ACIDENTE.xlsx"
    
    print(f"\n📊 Gerando: ESTATISTICA-DE-ACIDENTE")
    print(f"📂 Destino: {OUTPUT_XLSX}")
    
    # 1) Ler as bases
    print("  → Lendo planilhas...")
    df_ac = pd.read_excel(URL_ACIDENTES, sheet_name="STG_KCOR_ACIDENTES2")
    df_vi = pd.read_excel(URL_VITIMAS, sheet_name="STG_KCOR_ACIDENTES_VITIMAS")
    
    print(f"    • Acidentes: {len(df_ac)} registros")
    print(f"    • Vítimas: {len(df_vi)} registros")
    
    # 2) Filtrar por mês se fornecido
    if mes_ref and "DATAACIDENTE" in df_ac.columns:
        df_ac["DATAACIDENTE"] = pd.to_datetime(df_ac["DATAACIDENTE"], errors='coerce')
        df_ac["ano_mes"] = df_ac["DATAACIDENTE"].dt.strftime("%Y-%m")
        
        df_ac_filtrado = df_ac[df_ac["ano_mes"] == mes_ref].copy()
        
        if len(df_ac_filtrado) > 0:
            print(f"  → Filtrando para {mes_ref}: {len(df_ac_filtrado)} acidentes")
            df_ac = df_ac_filtrado
            
            # Filtrar vítimas relacionadas
            kocorrencias = set(df_ac["KOCORRENCIA"].unique())
            df_vi = df_vi[df_vi["KOCORRENCIA"].isin(kocorrencias)].copy()
    
    # 3) Agregar vítimas por acidente
    print("  → Agregando vítimas...")
    cond = norm_text(df_vi["CONDICAO"])
    
    is_fatal = cond.eq("fatal") | cond.str.contains("fatal", na=False)
    is_grave = cond.eq("grave") | cond.str.contains("grave", na=False)
    is_leve = cond.eq("leve") | cond.str.contains("leve", na=False)
    
    agg_vitimas = (
        df_vi.assign(
            _fatal=is_fatal.astype(int),
            _grave=is_grave.astype(int),
            _leve=is_leve.astype(int),
        )
        .groupby("KOCORRENCIA", as_index=False)[["_fatal", "_grave", "_leve"]]
        .sum()
        .rename(columns={
            "_fatal": "numero_mortos",
            "_grave": "numero_feridos_graves",
            "_leve": "numero_feridos_leves",
        })
    )
    
    # 4) Construir localização
    sigla = df_ac["SIGLARODOVIA"].astype(str).str.strip()
    km = df_ac["KM"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    mts = df_ac["MTS"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    
    df_ac["localizacao_quilometro"] = sigla + " " + km + "+" + mts
    
    # 5) Juntar dados
    df_final = df_ac.merge(agg_vitimas, how="left", on="KOCORRENCIA")
    
    # 6) Preencher nulos
    for c in ["numero_mortos", "numero_feridos_graves", "numero_feridos_leves"]:
        if c in df_final.columns:
            df_final[c] = df_final[c].fillna(0).astype(int)
    
    # 7) Montar dataframe final
    df_out = pd.DataFrame({
        "numero_acidentes": df_final["KOCORRENCIA"],
        "tipo_de_acidente": df_final["TIPOACIDENTE"],
        "numero_mortos": df_final["numero_mortos"],
        "numero_feridos_graves": df_final["numero_feridos_graves"],
        "numero_feridos_leves": df_final["numero_feridos_leves"],
        "localizacao_quilometro": df_final["localizacao_quilometro"],
        "horario": df_final["HORAACIDENTE"],
        "causa_provavel": df_final["CAUSAPROVAVEL"],
    })
    
    # 8) Salvar
    df_out.to_excel(OUTPUT_XLSX, index=False)
    print(f"  ✅ Salvo: {len(df_out)} registros")
    
    return OUTPUT_XLSX

# Para permitir execução direta do arquivo (teste)
if __name__ == "__main__":
    import sys
    pasta = sys.argv[1] if len(sys.argv) > 1 else "Excel"
    mes = sys.argv[2] if len(sys.argv) > 2 else None
    gerar_relatorio(pasta, mes)
