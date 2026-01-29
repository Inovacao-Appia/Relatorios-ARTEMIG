import pandas as pd

URL_ATENDIMENTOS = "STG_ATENDIMENTOS_KCOR.xlsx"
OUTPUT_XLSX = "COMBATE_A_INCENDIO.xlsx"

# Lê a base de origem
df_origem = pd.read_excel(URL_ATENDIMENTOS)

# Lista de tipos de ocorrência de incêndio para o filtro (DESCROCORRENCIA)
filtro_incendio = [
    'Incêndio de Grande Proporção',
    'Incêndio de Média Proporção',
    'Incêndio de Pequena Proporção'
]

# Lista de tipos de atendimento que devemos excluir (DSC_TIPO_ATENDIMENTO)
tipos_atendimento_excluir = [
    'Cancelado(QTA)',
    'Não Localizado'
]

# Aplica o filtro na coluna DESCROCORRENCIA
df_filtrado = df_origem[df_origem['DESCROCORRENCIA'].isin(filtro_incendio)].copy()

# Aplica o filtro excluindo esses tipos na coluna DSC_TIPO_ATENDIMENTO
df_filtrado= df_filtrado[~df_filtrado['DSC_TIPO_ATENDIMENTO'].isin(tipos_atendimento_excluir)].copy()


# Função para calcular hora de chegada (HORAOCORRENCIA + TEMPOCHEGADA em minutos)
def calcular_hora_chegada(hora_ocorrencia, tempo_chegada_minutos):
    """
    hora_ocorrencia: inteiro de 0 a 23 (hora cheia)
    tempo_chegada_minutos: inteiro (minutos)
    retorna string no formato HH:MM
    """
    try:
        hora_ocorrencia = int(hora_ocorrencia)
        tempo_chegada_minutos = int(tempo_chegada_minutos)

        # converte para minutos totais e depois volta para HH:MM (ciclo de 24h)
        minutos_totais = hora_ocorrencia * 60 + tempo_chegada_minutos
        hora = (minutos_totais // 60) % 24
        minuto = minutos_totais % 60

        return f"{hora:02d}:{minuto:02d}"
    except Exception as e:
        print("Erro ao calcular hora de chegada:", e)
        return None

# Monta o DataFrame final no padrão da planilha de combate a incêndio
df_final = pd.DataFrame()

# De-para conforme solicitado
df_final['Ocorrência'] = df_filtrado['KOCORRENCIA']
df_final['Data'] = df_filtrado['DATAOCORRENCIA']
df_final['Rodovia'] = df_filtrado['RODOVIA']
df_final['Km Inicial'] = df_filtrado['KM']
df_final['HM Inicial'] = df_filtrado['HORAOCORRENCIA']

# Km Final = mesmo valor do Km Inicial (replicado)
df_final['Km Final'] = df_filtrado['KM']

# HM Final = HORAOCORRENCIA + TEMPOCHEGADA
df_final['HM Final'] = df_filtrado.apply(
    lambda row: calcular_hora_chegada(row['HORAOCORRENCIA'], row['TEMPOCHEGADA']),
    axis=1
)

# Área atingida = null (não temos essa informação)
df_final['Área atingida (ha)'] = None

df_final['Hora de Acionamento'] = df_filtrado['HORAOCORRENCIA']

# Hora de Chegada = HORAOCORRENCIA + TEMPOCHEGADA
df_final['Hora de Chegada'] = df_filtrado.apply(
    lambda row: calcular_hora_chegada(row['HORAOCORRENCIA'], row['TEMPOCHEGADA']),
    axis=1
)

df_final['Tempo de Chegada'] = df_filtrado['TEMPOCHEGADA']
df_final['Recursos Mobilizados'] = df_filtrado['TIPO']

# Causas Prováveis = null (não temos essa informação)
df_final['Causas Prováveis'] = None

# Garante índice sequencial
df_final = df_final.reset_index(drop=True)

# Salva no arquivo de saída
df_final.to_excel(OUTPUT_XLSX, index=False)

print("Registros originais:", len(df_origem))
print("Registros após filtro de incêndio:", len(df_filtrado))
print("Arquivo gerado:", OUTPUT_XLSX)