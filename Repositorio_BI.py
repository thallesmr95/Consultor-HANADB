import os
from hdbcli import dbapi
import pandas as pd

# Nome do arquivo de saída
output_file = "Repositorio_BI.xlsx"

# Excluir o arquivo antigo se existir
if os.path.exists(output_file):
    os.remove(output_file)

# Configuração da conexão
conn = dbapi.connect(
    address="HostName",
    port=30015,  # Porta padrão do HANA
    user="Usuario",
    password="Senha"
)

cursor = conn.cursor()

# Executando a consulta SQL
cursor.execute("""SELECT TOP 200 * FROM COREPRIME.REPOSITORIO_BI_4""")

# Obtendo os resultados
rows = cursor.fetchall()
headers = [desc[0] for desc in cursor.description]

# Criando um DataFrame Pandas
df = pd.DataFrame(rows, columns=headers)

# Corrigindo nomes das colunas (removendo espaços e caracteres especiais)
df.columns = [col.replace(" ", "_") for col in df.columns]

# Verificando os primeiros dados antes de salvar
print(df.head())

# Salvando no Excel com openpyxl
df.to_excel(output_file, index=False, engine="openpyxl")

print(f"Arquivo {output_file} criado com sucesso!")

# Fechando conexão
cursor.close()
conn.close()
