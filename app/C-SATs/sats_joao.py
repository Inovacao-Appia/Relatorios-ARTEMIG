import pyodbc
import pandas as pd
from dotenv import load_dotenv
import os
import time
import calendar
from datetime import date
import time
# === CARREGA VARIÁVEIS DO .env ===
load_dotenv()

mes= 3
ano=2026
# calendar.monthrange retorna um tupla: (dia_da_semana_que_comeca, ultimo_dia_do_mes)
ultimo_dia = calendar.monthrange(ano, mes)[1]
    
data_inicio = date(ano, mes, 1)
data_fim = date(ano, mes, ultimo_dia)
    
    
# === CONFIGURAÇÃO DA CONEXÃO ===
# Você pode usar .env ou deixar direto no código (para testes)
server   = os.getenv("SQL_HOST",   "192.168.40.22")
port     = os.getenv("SQL_PORT",   "1433")
database = os.getenv("SQL_DB",     "AT3_db_2")   # banco onde estão as tabelas
username = os.getenv("SQL_USER",   "sa")
password = os.getenv("SQL_PASS",   "$4dmp4dr40@nasc")

# === STRING DE CONEXÃO ===
connection_string = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={server},{port};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
)
# Caminhos exatos das pastas que você já criou (conforme a imagem)
pasta_amostras = os.path.join("SQL/SATS", "AMOSTRAS")
pasta_ocorrencias = os.path.join("SQL/SATS", "OCORRENCIAS")

# ======================================================================
# 2. Lista Completa de Tabelas
# ======================================================================
tabelas = [
    "AT3_1_Amostra", "AT3_1_Ocorrencia",
    "AT3_2_Amostra", "AT3_2_Ocorrencia",
    "AT3_3_Amostra", "AT3_3_Ocorrencia",
    "AT3_4_Amostra", "AT3_4_Ocorrencia",
    "AT3_5_Amostra", "AT3_5_Ocorrencia",
    "AT3_6_Amostra", "AT3_6_Ocorrencia",
    "AT3_7_Amostra", "AT3_7_Ocorrencia",
    "AT3_8_Amostra", "AT3_8_Ocorrencia",
    "AT3_9_Amostra", "AT3_9_Ocorrencia",
    "AT3_10_Amostra", "AT3_10_Ocorrencia",
    "AT3_11_Amostra", "AT3_11_Ocorrencia",
    "AT3_12_Amostra", "AT3_12_Ocorrencia",
    "AT3_13_Amostra", "AT3_13_Ocorrencia",
    "AT3_14_Amostra", "AT3_14_Ocorrencia",
    "AT3_15_Amostra", "AT3_15_Ocorrencia",
    "AT3_16_Amostra", "AT3_16_Ocorrencia",
    "AT3_17_Amostra", "AT3_17_Ocorrencia",
    "AT3_18_Amostra", "AT3_18_Ocorrencia",
    "AT3_19_Amostra", "AT3_19_Ocorrencia",
    "AT3_20_Amostra", "AT3_20_Ocorrencia",
    "AT3_37_Amostra", "AT3_37_Ocorrencia",
    "AT3_99_Amostra", "AT3_99_Ocorrencia",
    "AT3_231_Amostra", "AT3_231_Ocorrencia"
]

# Listas do Python onde os DataFrames ficarão salvos na memória
amostras_dfs = []
ocorrencias_dfs = []

# ======================================================================
# 3. === CONECTA, EXECUTA E EXPORTA ===
# ======================================================================
try:
    with pyodbc.connect(connection_string) as conn:
        print("✓ Conectado ao SQL Server com sucesso!\n")
        
        for nome_tabela in tabelas:
            query = f"""
                    SELECT *
                    FROM AT3_db_1.dbo.{nome_tabela}
                    WHERE 
                    1=1
                    AND Instante >= '{data_inicio}'
                    AND Instante <= '{data_fim}'
                    ORDER BY Instante DESC
                    """
            print(f"→ Executando query: {nome_tabela}")
            
            # Carrega a tabela em um DataFrame
            df = pd.read_sql(query, conn)
            print(f"  ✓ {len(df)} linhas retornadas")
            print(df)
            Nome_do_Excel = f"{nome_tabela}.xlsx"
            
            # Avalia se a tabela é de Amostra ou de Ocorrência
            if nome_tabela.endswith("_Amostra"):
                amostras_dfs.append(df)
                caminho_salvar = os.path.join(pasta_amostras, Nome_do_Excel)
                
            elif nome_tabela.endswith("_Ocorrencia"):
                ocorrencias_dfs.append(df)
                caminho_salvar = os.path.join(pasta_ocorrencias, Nome_do_Excel)
            
            # Salva o Excel direto na pasta correspondente
            df.to_excel(caminho_salvar, index=False)
            print(f"  ✓ Arquivo salvo em: {caminho_salvar}\n")
            time.sleep(5) 

    print("-" * 50)
    print("✓ Processo finalizado com sucesso!")
    print(f"Arquivos gerados: {len(amostras_dfs)} nas AMOSTRAS e {len(ocorrencias_dfs)} nas OCORRENCIAS.")

except pyodbc.Error as e:
    print(f"✗ Erro de conexão/query: {e}")
except Exception as e:
    print(f"✗ Erro geral: {e}")