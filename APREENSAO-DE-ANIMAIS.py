import pandas as pd

URL_ATENDIMENTOS = "STG_ATENDIMENTOS_KCOR.xlsx"
OUTPUT_XLSX = "APREENSAO_DE_ANIMAIS.xlsx"

# Lê a base de origem
df_origem = pd.read_excel(URL_ATENDIMENTOS)

# Lista de tipos de ocorrência com animais para o filtro (DESCROCORRENCIA)
filtro_animais = [
    'Animal Doméstico Morto',
    'Animal Silvestre Morto',
    'Animal Solto',
    'Animal Apreendido',
    'Animal de Grande Porte Morto',
    'Animal Resgatado'
]

# Lista de tipos de atendimento que devemos excluir (DSC_TIPO_ATENDIMENTO)
tipos_atendimento_excluir = [
    'Cancelado(QTA)',
    'Não Localizado'
]
# Aplica o filtro na coluna DESCROCORRENCIA
df_filtrado = df_origem[df_origem['DESCROCORRENCIA'].isin(filtro_animais)].copy()

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

# Monta o DataFrame final no padrão da planilha de apreensão de animais
df_final = pd.DataFrame()

# De-para conforme solicitado
df_final['Ocorrência'] = df_filtrado['NUMOCORRENCIA']
df_final['Data'] = df_filtrado['DATAOCORRENCIA']
df_final['Rodovia'] = df_filtrado['RODOVIA']
df_final['KM'] = df_filtrado['KM']
df_final['HM'] = df_filtrado['HORAOCORRENCIA']

# não temos essa informação -> deixar null
df_final['Tipo de Animal'] = None

df_final['Hora de Acionamento'] = df_filtrado['HORAOCORRENCIA']

df_final['Hora de Chegada'] = df_filtrado.apply(
    lambda row: calcular_hora_chegada(row['HORAOCORRENCIA'], row['TEMPOCHEGADA']),
    axis=1
)

df_final['Tempo de Chegada'] = df_filtrado['TEMPOCHEGADA']

# não temos essa informação -> deixar null
df_final['Destino'] = None

# Garante índice sequencial
df_final = df_final.reset_index(drop=True)

# Salva no arquivo de saída
df_final.to_excel(OUTPUT_XLSX, index=False)

print("Registros originais:", len(df_origem))
print("Registros após filtro de animais:", len(df_filtrado))
print("Arquivo gerado:", OUTPUT_XLSX)
