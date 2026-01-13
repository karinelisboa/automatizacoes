import pandas as pd
import os
import traceback
from datetime import datetime, timedelta
from google.cloud import bigquery
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / "config" / ".env"

load_dotenv(dotenv_path=ENV_PATH, override=True)

DATA_DIA = os.getenv("DATA_DIA")
MES = os.getenv('MES')
ANO = os.getenv('ANO')
key_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
USUARIO = os.getenv('USUARIO')
BIGQUERY_PROJECT = os.getenv('BIGQUERY_PROJECT')
BIGQUERY_DATASET = os.getenv('BIGQUERY_DATASET')
# Pastas
pasta_origem_venda = f"X:\\Dados\\Bases Operacionais\\RDO\\1. RIO CARD\\ARQUIVOS_RECEBIDOS\\VENDA_A_BORDO\\RECEBIDO_RCTI\\{ANO}_{MES:02d}"
pasta_origem_vt = f"X:\\Dados\\Bases Operacionais\\RDO\\1. RIO CARD\\ARQUIVOS_RECEBIDOS\\VT\\{ANO}_{MES:02d}"
pasta_destino = f"C:\Users\{USUARIO}\Desktop\teste_venda_a_bordo"
# ====================================

print("=" * 60)
print("PROCESSAMENTO DE ARQUIVOS - VENDA A BORDO E VT")
print("=" * 60)
print(f"\nMes/Ano: {MES:02d}/{ANO}")
print(f"Destino: {pasta_destino}")
print()

# Cria a pasta de destino se nao existir
print("Verificando pasta de destino...")
try:
    os.makedirs(pasta_destino, exist_ok=True)
    print("Pasta de destino pronta")
except Exception as e:
    print(f"Erro ao criar pasta: {e}")
    exit()

print("\n" + "=" * 60)
print("1. PROCESSANDO VENDA A BORDO")
print("=" * 60)

# Verifica se a pasta de origem existe
print("\nVerificando pasta de origem...")
if not os.path.exists(pasta_origem_venda):
    print(f"ERRO: Pasta de origem nao encontrada!")
    print(f"{pasta_origem_venda}")
    exit()
else:
    print(f"Pasta de origem encontrada")
    print(f"{pasta_origem_venda}")

# Procura arquivos Excel na pasta
print("\nProcurando arquivos Excel...")
try:
    todos_arquivos = os.listdir(pasta_origem_venda)
    arquivos_excel = []
    for f in todos_arquivos:
        if f.endswith('.xlsx') or f.endswith('.xls'):
            if not f.startswith('~'):
                arquivos_excel.append(f)
except Exception as e:
    print(f"Erro ao listar arquivos: {e}")
    exit()

if not arquivos_excel:
    print("ERRO: Nenhum arquivo Excel encontrado na pasta!")
    exit()

print(f"{len(arquivos_excel)} arquivo(s) Excel encontrado(s)")
for arq in arquivos_excel:
    print(f"  {arq}")

arquivo_escolhido = arquivos_excel[0]
arquivo_entrada = os.path.join(pasta_origem_venda, arquivo_escolhido)

print(f"\nCarregando arquivo...")
print(f"Arquivo: {arquivo_escolhido}")

try:
    df = pd.read_excel(arquivo_entrada)
    print(f"{len(df):,} linhas carregadas".replace(',', '.'))
except Exception as e:
    print(f"\nERRO ao carregar arquivo:")
    print(f"{str(e)}")
    traceback.print_exc()
    exit()

# Verifica se as colunas necessarias existem
print(f"\nVerificando colunas...")

mapeamento_colunas = {
    'CD_OPERAD': 'codigo_operadora',
    'NM_FANT': 'nome_fantasia',
    'CD_LINHA': 'codigo_linha',
    'NR_LINHA': 'servico',
    'NR_ESTAC_CARRO': 'numero_carro',
    'DT_TRANS': 'data_transacao',
    'QT_TRANS': 'qtd_transacao',
    'VL_TRANS': 'valor_transacao'
}

colunas_necessarias = list(mapeamento_colunas.keys())
colunas_faltando = [col for col in colunas_necessarias if col not in df.columns]

if colunas_faltando:
    print(f"\nERRO: Colunas nao encontradas: {', '.join(colunas_faltando)}")
    print(f"\nColunas disponiveis ({len(df.columns)}):")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
    exit()

print(f"Selecionando apenas colunas do BigQuery...")
df = df[colunas_necessarias]

print(f"Renomeando colunas para padrao BigQuery...")
df = df.rename(columns=mapeamento_colunas)

print(f"Colunas apos selecao e renomeacao:")
for col in df.columns:
    print(f"  {col}")

# Funcao para corrigir encoding
def corrigir_encoding(texto):
    if pd.isna(texto):
        return texto
    
    texto = str(texto)
    
    mapeamento = {
        'Ã§': 'ç', 'Ã£': 'ã', 'Ã¡': 'á', 'Ã©': 'é', 'Ã­': 'í',
        'Ã³': 'ó', 'Ãº': 'ú', 'Ã¢': 'â', 'Ãª': 'ê', 'Ã´': 'ô',
        'Ã ': 'à', 'Ã¨': 'è', 'Ã¬': 'ì', 'Ã²': 'ò', 'Ã¹': 'ù',
        'Ã': 'Á', 'Ã‰': 'É', 'Ã': 'Í', 'Ã"': 'Ó', 'Ãš': 'Ú',
        'Ã‚': 'Â', 'ÃŠ': 'Ê', 'Ã"': 'Ô', 'Ã€': 'À', 'Ã‡': 'Ç',
    }
    
    for errado, correto in mapeamento.items():
        texto = texto.replace(errado, correto)
    
    return texto

# Funcao para corrigir espacos extras
def corrigir_espacos(texto):
    if pd.isna(texto):
        return texto
    
    texto = str(texto)
    import re
    texto = re.sub(r'\s+', ' ', texto)
    texto = texto.strip()
    
    return texto

# Funcao para corrigir valor transacao
def corrigir_valor_transacao(valor):
    if pd.isna(valor):
        return None

    try:
        # Converte para string
        s = str(valor).strip()

        # Remove espaços
        s = s.replace(' ', '')

        # Caso tenha mais de uma vírgula → milhar + decimal
        if s.count(',') > 1:
            partes = s.split(',')
            decimal = partes[-1]
            inteiro = ''.join(partes[:-1])
            s = inteiro + '.' + decimal

        # Caso padrão brasileiro: 1.706,10
        elif ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')

        # Caso só vírgula: 1024,60
        elif ',' in s:
            s = s.replace(',', '.')

        # Converte para float
        return float(s)

    except Exception:
        return None


# Funcao para converter data
def converter_data_excel(valor):
    if pd.isna(valor):
        return valor
    try:
        if isinstance(valor, (int, float)):
            data_base = datetime(1899, 12, 30)
            data_convertida = data_base + timedelta(days=int(valor))
            return data_convertida.date()
        elif isinstance(valor, datetime):
            return valor.date()
        elif isinstance(valor, pd.Timestamp):
            return valor.date()
        return valor
    except:
        return valor

print(f"\nCorrigindo dados...")

registros_encoding = df['nome_fantasia'].astype(str).str.contains('Ã', na=False).sum()
registros_espacos = df['nome_fantasia'].astype(str).str.contains(r'\s{2,}', regex=True, na=False).sum()

print(f"Registros com problema de encoding: {registros_encoding}")
print(f"Registros com espacos extras: {registros_espacos}")

df['nome_fantasia'] = df['nome_fantasia'].apply(corrigir_encoding)
df['nome_fantasia'] = df['nome_fantasia'].apply(corrigir_espacos)

print(f"\nConvertendo datas...")
df['data_transacao'] = df['data_transacao'].apply(converter_data_excel)

print(f"\nCorrigindo valores de transacao...")
df['valor_transacao'] = df['valor_transacao'].apply(corrigir_valor_transacao)

print(f"Correcoes aplicadas")

print(f"\nExemplos de valores corrigidos:")
print(f"\nnome_fantasia:")
exemplos_unicos = df['nome_fantasia'].dropna().unique()[:3]
for exemplo in exemplos_unicos:
    print(f"  {exemplo}")

print(f"\ndata_transacao (primeiras 3 datas):")
exemplos_datas = df['data_transacao'].dropna().head(3)
for data in exemplos_datas:
    print(f"  {data}")

print(f"\nvalor_transacao (primeiros 3 valores):")
exemplos_valores = df['valor_transacao'].dropna().head(3)
for valor in exemplos_valores:
    print(f"  {valor}")

nome_saida_venda = f"{ANO}_{MES:02d}_VENDA_A_BORDO_PROCESSADO.xlsx"
arquivo_saida = os.path.join(pasta_destino, nome_saida_venda)

print(f"\nSalvando arquivo...")
print(f"Nome: {nome_saida_venda}")

try:
    df.to_excel(arquivo_saida, index=False)
    
    if os.path.exists(arquivo_saida):
        tamanho = os.path.getsize(arquivo_saida) / 1024 / 1024
        print(f"Arquivo salvo com sucesso! ({tamanho:.2f} MB)")
    else:
        print(f"Comando executado mas arquivo nao encontrado!")
        
except Exception as e:
    print(f"\nERRO ao salvar arquivo:")
    print(f"{str(e)}")
    traceback.print_exc()
    exit()

print(f"\nProcesso concluido!")
print(f"Arquivo salvo: {arquivo_saida}")
print(f"Total de registros: {len(df):,}".replace(',', '.'))

# ============================================================
# PROCESSA ARQUIVO VT
# ============================================================

print("\n" + "=" * 60)
print("2. PROCESSANDO VT")
print("=" * 60)

print("\nVerificando pasta de origem VT...")
if not os.path.exists(pasta_origem_vt):
    print(f"ERRO: Pasta de origem VT nao encontrada!")
    print(f"{pasta_origem_vt}")
    exit()
else:
    print(f"Pasta de origem encontrada")
    print(f"{pasta_origem_vt}")

print("\nProcurando arquivo Pasta1...")
arquivo_vt = None
for ext in ['.xlsx', '.xls', '.xlsm']:
    caminho_teste = os.path.join(pasta_origem_vt, f"Pasta1{ext}")
    if os.path.exists(caminho_teste):
        arquivo_vt = caminho_teste
        break

if not arquivo_vt:
    print("ERRO: Arquivo 'Pasta1' nao encontrado!")
    print("\nArquivos disponiveis na pasta:")
    try:
        arquivos_vt = os.listdir(pasta_origem_vt)
        for arq in arquivos_vt[:10]:
            print(f"  {arq}")
    except:
        pass
    exit()

print(f"Arquivo encontrado: {os.path.basename(arquivo_vt)}")

nome_aba = f"{ANO}_{MES:02d}"
print(f"\nCarregando arquivo VT da aba '{nome_aba}'...")
try:
    df_vt = pd.read_excel(arquivo_vt, sheet_name=nome_aba)
    print(f"{len(df_vt):,} linhas carregadas".replace(',', '.'))
except Exception as e:
    print(f"\nERRO ao carregar arquivo:")
    print(f"{str(e)}")
    
    try:
        print(f"\nAbas disponiveis no arquivo:")
        xls = pd.ExcelFile(arquivo_vt)
        for aba in xls.sheet_names:
            print(f"  {aba}")
    except:
        pass
    
    traceback.print_exc()
    exit()

if 'Data2' not in df_vt.columns:
    print(f"\nERRO: Coluna 'Data2' nao encontrada!")
    print(f"\nColunas disponiveis ({len(df_vt.columns)}):")
    for i, col in enumerate(df_vt.columns, 1):
        print(f"{i}. {col}")
    exit()

print(f"Coluna 'Data2' encontrada")

print(f"\nExcluindo coluna 'Campo28' se existir...")
if 'Campo28' in df_vt.columns:
    df_vt = df_vt.drop(columns=['Campo28'])
    print(f"Coluna 'Campo28' excluida")
else:
    print(f"Coluna 'Campo28' nao encontrada (pulando)")

print(f"\nConvertendo datas na coluna Data2 (removendo horas)...")
df_vt['Data2'] = df_vt['Data2'].apply(converter_data_excel)
print(f"Conversao concluida - apenas datas, sem horarios")

print(f"\nData2 (primeiras 3 datas):")
exemplos_datas_vt = df_vt['Data2'].dropna().head(3)
for data in exemplos_datas_vt:
    print(f"  {data}")

nome_saida_vt = f"{ANO}_{MES:02d}_VT_PROCESSADO.xlsx"
arquivo_saida_vt = os.path.join(pasta_destino, nome_saida_vt)

print(f"\nSalvando arquivo VT...")
print(f"Nome: {nome_saida_vt}")

try:
    df_vt.to_excel(arquivo_saida_vt, index=False)
    
    if os.path.exists(arquivo_saida_vt):
        tamanho = os.path.getsize(arquivo_saida_vt) / 1024 / 1024
        print(f"Arquivo salvo com sucesso! ({tamanho:.2f} MB)")
    else:
        print(f"Comando executado mas arquivo nao encontrado!")
        
except Exception as e:
    print(f"\nERRO ao salvar arquivo:")
    print(f"{str(e)}")
    traceback.print_exc()
    exit()

print(f"\nProcesso concluido!")
print(f"Arquivo salvo: {arquivo_saida_vt}")
print(f"Total de registros: {len(df_vt):,}".replace(',', '.'))

print("\n" + "=" * 60)
print("PROCESSAMENTO FINALIZADO")
print("=" * 60)
print(f"\nArquivos salvos em: {pasta_destino}")
print(f"  1. {nome_saida_venda}")
print(f"  2. {nome_saida_vt}")

print("\n" + "=" * 60)

# ============================================================
# UPLOAD PARA BIGQUERY - VENDA A BORDO
# ============================================================

print("\n" + "=" * 60)
print("3. UPLOAD PARA BIGQUERY - VENDA A BORDO")
print("=" * 60)

print("\nConectando ao BigQuery...")
try:
    client = bigquery.Client.from_service_account_json(key_path)
    print("Conexao estabelecida com sucesso")
except Exception as e:
    print(f"ERRO ao conectar ao BigQuery:")
    print(f"{str(e)}")
    traceback.print_exc()
    exit()

# Configuracoes principais
project_id = BIGQUERY_PROJECT
dataset_id = BIGQUERY_DATASET
table_id = "venda_a_bordo"
source_file = arquivo_saida

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

print(f"\nTabela de destino: {table_ref}")

# Schema da tabela
schema = [
    bigquery.SchemaField("codigo_operadora", "STRING"),
    bigquery.SchemaField("nome_fantasia", "STRING"),
    bigquery.SchemaField("codigo_linha", "STRING"),
    bigquery.SchemaField("servico", "STRING"),
    bigquery.SchemaField("numero_carro", "STRING"),
    bigquery.SchemaField("data_transacao", "DATE"),
    bigquery.SchemaField("qtd_transacao", "INTEGER"),
    bigquery.SchemaField("valor_transacao", "FLOAT64"),
]

# Configuracoes do job de carregamento
job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.CSV,
    skip_leading_rows=1,
    autodetect=False,
    field_delimiter=',',
    schema=schema,
    write_disposition=bigquery.WriteDisposition.WRITE_APPEND
)

print(f"\nCarregando arquivo para o BigQuery...")
print(f"Arquivo: {source_file}")

try:
    # Converte Excel para CSV temporario
    csv_temp = source_file.replace('.xlsx', '_temp.csv')
    df.to_csv(csv_temp, index=False, encoding='utf-8')
    
    # Carrega o arquivo CSV para o BigQuery
    with open(csv_temp, "rb") as file:
        job = client.load_table_from_file(file, table_ref, job_config=job_config)
    
    # Aguarda o job ser concluido
    job.result()
    
    # Remove arquivo temporario
    os.remove(csv_temp)
    
    print(f"Upload concluido com sucesso!")
    print(f"Total de linhas carregadas: {len(df):,}".replace(',', '.'))
    
except Exception as e:
    print(f"\nERRO ao fazer upload para o BigQuery:")
    print(f"{str(e)}")
    traceback.print_exc()

print("\n" + "=" * 60)
print("PROCESSO COMPLETO FINALIZADO")
print("=" * 60)

# ============================================================
# UPLOAD PARA BIGQUERY - Vt
# ============================================================

print("\n" + "=" * 60)
print("3. UPLOAD PARA BIGQUERY - vt")
print("=" * 60)

print("\nConectando ao BigQuery...")
try:
    client = bigquery.Client.from_service_account_json(key_path)
    print("Conexao estabelecida com sucesso")
except Exception as e:
    print(f"ERRO ao conectar ao BigQuery:")
    print(f"{str(e)}")
    traceback.print_exc()
    exit()

# Configuracoes principais
project_id = BIGQUERY_PROJECT
dataset_id = BIGQUERY_DATASET
table_id = "vt"
source_file = arquivo_saida_vt

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

print(f"\nTabela de destino: {table_ref}")

# Schema da tabela
schema = [
    bigquery.SchemaField("codigo_empresa", "STRING"),
    bigquery.SchemaField("servico", "STRING"),
    bigquery.SchemaField("anomesdia", "STRING"),

    bigquery.SchemaField("ido", "INTEGER"),
    bigquery.SchemaField("pne", "INTEGER"),
    bigquery.SchemaField("uni", "INTEGER"),
    bigquery.SchemaField("est", "INTEGER"),
    bigquery.SchemaField("mun", "INTEGER"),
    bigquery.SchemaField("emp", "INTEGER"),

    bigquery.SchemaField("qtd_buc1", "INTEGER"),
    bigquery.SchemaField("qtd_buc2", "INTEGER"),
    bigquery.SchemaField("valor_buc", "FLOAT64"),

    bigquery.SchemaField("qtd_bus1", "INTEGER"),
    bigquery.SchemaField("qtd_bus2", "INTEGER"),
    bigquery.SchemaField("valor_sv", "FLOAT64"),

    bigquery.SchemaField("qtd_vt", "INTEGER"),
    bigquery.SchemaField("receita_vt", "FLOAT64"),

    bigquery.SchemaField("qtd_vt_total", "INTEGER"),
    bigquery.SchemaField("receita_vt_total", "FLOAT64"),

    bigquery.SchemaField("qtd_gratuidade", "INTEGER"),

    bigquery.SchemaField("empresa", "STRING"),
    bigquery.SchemaField("data", "DATE"),

    bigquery.SchemaField("dia_semana", "INTEGER"),
    bigquery.SchemaField("dia", "INTEGER"),
    bigquery.SchemaField("mes", "INTEGER"),
    bigquery.SchemaField("ano", "INTEGER"),

    bigquery.SchemaField("ano_mes", "STRING"),
    bigquery.SchemaField("gtr_paslvruni", "INTEGER"),
]

# Configuracoes do job de carregamento
job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.CSV,
    skip_leading_rows=1,
    autodetect=False,
    field_delimiter=',',
    schema=schema,
    write_disposition=bigquery.WriteDisposition.WRITE_APPEND
)

print(f"\nCarregando arquivo para o BigQuery...")
print(f"Arquivo: {source_file}")

try:
    # Converte Excel para CSV temporario
    csv_temp = source_file.replace('.xlsx', '_temp.csv')
    df.to_csv(csv_temp, index=False, encoding='utf-8')
    
    # Carrega o arquivo CSV para o BigQuery
    with open(csv_temp, "rb") as file:
        job = client.load_table_from_file(file, table_ref, job_config=job_config)
    
    # Aguarda o job ser concluido
    job.result()
    
    # Remove arquivo temporario
    os.remove(csv_temp)
    
    print(f"Upload concluido com sucesso!")
    print(f"Total de linhas carregadas: {len(df):,}".replace(',', '.'))
    
except Exception as e:
    print(f"\nERRO ao fazer upload para o BigQuery:")
    print(f"{str(e)}")
    traceback.print_exc()

print("\n" + "=" * 60)
print("PROCESSO COMPLETO FINALIZADO")
print("=" * 60)