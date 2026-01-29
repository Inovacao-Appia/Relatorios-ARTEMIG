import pandas as pd

URL_ATENDIMENTOS = "STG_ATENDIMENTOS_KCOR.xlsx"
OUTPUT_XLSX = "ATENDIMENTO_MECANICO.xlsx"

# Lê a base de origem
df_origem = pd.read_excel(URL_ATENDIMENTOS)

# Lista de tipos de ocorrência mecânica para o filtro
filtro_mecanico = [
    'Pane Mecânica',
    'Pane na bomba de combustível',
    'Pane na suspensão',
    'Pane na transmissão',
    'Pane no motor',
    'Pane no sistema de refrigeração',
    'Pane nos freios',
    'Pane Seca'
]

# Lista de tipos de atendimento que devemos excluir (DSC_TIPO_ATENDIMENTO)
tipos_atendimento_excluir = [
    'Cancelado(QTA)',
    'Não Localizado'
]

# Aplica o filtro na coluna DESCROCORRENCIA
df_filtrado = df_origem[df_origem['DESCROCORRENCIA'].isin(filtro_mecanico)].copy()

# Aplica o filtro excluindo esses tipos na coluna DSC_TIPO_ATENDIMENTO
df_filtrado= df_filtrado[~df_filtrado['DSC_TIPO_ATENDIMENTO'].isin(tipos_atendimento_excluir)].copy()

# Função para calcular a Hora de Chegada (HORAOCORRENCIA + TEMPOCHEGADA em minutos)
def calcular_hora_chegada(hora_ocorrencia, tempo_chegada_minutos):
    """
    hora_ocorrencia: inteiro de 0 a 23 (hora cheia)
    tempo_chegada_minutos: inteiro (minutos)
    retorna string no formato HH:MM
    """
    try:
        # garante que são inteiros
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

# Monta o DataFrame final no padrão da fiscalizadora
df_final = pd.DataFrame()

df_final['Ocorrência'] = df_filtrado['NUMOCORRENCIA']
df_final['Data'] = df_filtrado['DATAOCORRENCIA']  # pode formatar depois se precisar
df_final['Rodovia'] = df_filtrado['RODOVIA']
df_final['Km'] = df_filtrado['KM']
df_final['HM'] = df_filtrado['HORAOCORRENCIA']
df_final['Tipo de Ocorrência'] = df_filtrado['DESCROCORRENCIA']
df_final['Tipo de Veículo'] = df_filtrado['TIPO']
df_final['Hora de Acionamento'] = df_filtrado['HORAOCORRENCIA']

df_final['Hora de Chegada'] = df_filtrado.apply(
    lambda row: calcular_hora_chegada(row['HORAOCORRENCIA'], row['TEMPOCHEGADA']),
    axis=1
)

df_final['Tempo de Chegada'] = df_filtrado['TEMPOCHEGADA']

# Garante índice sequencial
df_final = df_final.reset_index(drop=True)

# Salva no arquivo de saída
df_final.to_excel(OUTPUT_XLSX, index=False)

print("Registros originais:", len(df_origem))
print("Registros após filtro mecânico:", len(df_filtrado))
print("Arquivo gerado:", OUTPUT_XLSX)

