from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
import os
import pandas as pd
import re
from pathlib import Path
from selenium.webdriver.chrome.options import Options
from google.oauth2 import service_account
from google.cloud import bigquery
import shutil
import csv
import re
from dotenv import load_dotenv
import os
import yaml
from datetime import datetime
from pathlib import Path

load_dotenv()

def criar_cliente_bq():
    BQ_KEY_PATH = os.getenv("BQ_KEY_PATH")
    return bigquery.Client.from_service_account_json(BQ_KEY_PATH)

# DEFINICAO DE VARIAVEIS

data_atual = data_base.strftime("%d/%m/%Y")   # Power BI
data_bq    = data_base.strftime("%Y-%m-%d")   # BigQuery
data_ordem = data_base.strftime("%d-%m-%Y")   # Nome do arquivo

# Validar variável de ambiente DOWNLOADS_PATH antes de usar Path()
downloads_env = os.getenv("DOWNLOADS_PATH")
if not downloads_env:
    raise RuntimeError(
        "Environment variable DOWNLOADS_PATH is not set.\n"
        "Set DOWNLOADS_PATH (e.g. in a .env file) or export it in your shell before running the script."
    )

DOWNLOADS_PATH = Path(downloads_env)

# URL do Power BI
POWERBI_URL = os.getenv("POWERBI_URL")

# Caminho dos arquivos baixados renomeados
DESTINO_PATH   = Path(os.getenv("DESTINO_PATH"))
diretorio_destino = DESTINO_PATH

# Informando o caminho do programa do Chrome (chromedriver)
CHROMEDRIVER   = Path(os.getenv("CHROMEDRIVER_PATH"))
service = Service(CHROMEDRIVER)


# Caminho para sua chave de serviço JSON
BQ_KEY_PATH = os.getenv("BQ_KEY_PATH")

# Configura o cliente do BigQuery
client = criar_cliente_bq()



# CONFIGURACAO SELENIUM
chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Remove flag de automação
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  # Evita detecção
# chrome_options.add_argument("--headless=new")  # Usa o novo modo headless
chrome_options.add_argument("--disable-gpu")  # Evita problemas gráficos
chrome_options.add_argument("--start-maximized")  # Abre o navegador maximizado
chrome_options.add_argument("--no-sandbox")  # Evita problemas de permissão
chrome_options.add_argument("--disable-dev-shm-usage")  # Melhora o desempenho em sistemas com pouca RAM

# Evita bloqueio de download e verificação de vírus
prefs = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True
}
chrome_options.add_experimental_option("prefs", prefs)


# Criar o driver corretamente
driver = webdriver.Chrome(service=service, options=chrome_options)



# ACESSO AO DASHBOARD
# Acessa o link
driver.get(POWERBI_URL)

# Login no Power BI
# Preenche email
email_input = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.ID, "Email"))
)

POWERBI_EMAIL = os.getenv("POWERBI_EMAIL")
#email_input.send_keys(Keys.RETURN)
time.sleep(5)

# Preenche senha
senha_input = driver.find_element(By.ID, 'Password')
POWERBI_PASSWORD = os.getenv("POWERBI_PASSWORD")  # Senha

email_input.send_keys(POWERBI_EMAIL)
senha_input.send_keys(POWERBI_PASSWORD)
senha_input.send_keys(Keys.RETURN)

time.sleep(1)

# Clica no botão "Entrar"
botao_entrar = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.ID, "login-submit"))
)

# Tentativa direta
try:
    botao_entrar.click()
except:
    # Fallback via JS se o site bloquear o click do Selenium
    driver.execute_script("arguments[0].click();", botao_entrar)



#driver.switch_to.default_content()

# Esperar até o iframe com o src correto estar disponível
iframe = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'https://app.powerbi.com/reportEmbed?reportId=a6abefa3-5bc8')]"))
)

# Mudar para o iframe usando o src
driver.switch_to.frame(iframe)
time.sleep(10)


# FILTRO DE PERIODO
wait = WebDriverWait(driver, 60)

# Aguarda o visual do Indicador existir
indicador = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(@aria-label, 'Indicador') and contains(@aria-label, 'Clique aqui')]")
    )
)

# Garante que está visível
driver.execute_script(
    "arguments[0].scrollIntoView({block: 'center'});",
    indicador
)
time.sleep(1)

# Clique robusto
try:
    indicador.click()
except:
    driver.execute_script("arguments[0].click();", indicador)

time.sleep(4)


# Limpa todas as seleções do filtro
filtro = WebDriverWait(driver, 60).until(
    #EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Limpar todas as segmentações . Clique aqui para seguir link']"))
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Limpar todas as segmentações . Clique aqui para Seguir']"))
)
filtro.click()
time.sleep(5)

# Identifica os campos de data
campo_data_inicio = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, "//*[contains(@aria-label, 'Data de início')]"))
)
campo_data_fim = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, "//*[contains(@aria-label, 'Data de término')]"))
)

# Limpa os campos e insere a data do dia de hoje
campo_data_fim.clear()
campo_data_fim.send_keys(data_atual)
time.sleep(10)

campo_data_inicio.clear()
campo_data_inicio.send_keys(data_atual)


# Fecha o filtro clicando no ícone
wait = WebDriverWait(driver, 60)
indicador = wait.until(
    #EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']"))
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
)
indicador.click()



# DOWNLOAD
######################################### RESUMO
# Localiza o título da tabela
elemento_para_revelar = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, "(//div[@aria-label='Pagamento Operadora '])"))
)
                            
# Move o mouse até o elemento para aparecer o botão "Mais opções"
actions = ActionChains(driver)
actions.move_to_element(elemento_para_revelar).perform()

        
# Localiza e clica no botão "Mais opções"
botao_opcoes = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@class='vcMenuBtn' and @aria-label='Mais opções']"))
    )
botao_opcoes.click()

               
# Localiza e clica no botão "Exportar dados"
exportar_dados = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']"))
    )
exportar_dados.click()

exportar_botao = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Exportar']"))
    )
exportar_botao.click()
       

# Garante que a pasta de destino exista
os.makedirs(diretorio_destino, exist_ok=True)

# Espera o download finalizar
time.sleep(10)

# Lista todos os arquivos .xlsx da pasta de downloads
arquivos = [
    arquivo
    for arquivo in DOWNLOADS_PATH.iterdir()
    if arquivo.is_file() and arquivo.suffix.lower() == ".xlsx"
]


# Encontra o arquivo .xlsx mais recente (baseado na data de criação)
if arquivos:
    arquivo_mais_recente = max(arquivos, key=os.path.getctime)
    
    # Remove barras '/' do texto da linha (nome do arquivo)
    nome_arquivo_novo = f"{data_ordem} - Resumo.xlsx"
    
    # Define o caminho do novo arquivo na pasta de destino
    caminho_novo = os.path.join(diretorio_destino, nome_arquivo_novo)

    # Move o arquivo para a nova pasta
    shutil.move(arquivo_mais_recente, caminho_novo)

time.sleep(10)


#SUBINDO NA TABELA DO BIGQUERY
#data_ordem = '19-02-2025'
termo = 'Resumo'

        
# Lista para armazenar os DataFrames
lista_dataframes = []

# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    if arquivo.endswith('.xlsx') and termo in arquivo and str(data_ordem) in arquivo: 
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        df = pd.read_excel(caminho_completo)
        

        df = df.iloc[:-3]                                                                           #exclui as últimas 3 linhas do df
        lista_dataframes.append(df)  # Adicionar o DataFrame à lista

# Concatena todos os DataFrames em um único
resumo = pd.concat(lista_dataframes, ignore_index=True)


# Remove coluna duplicada de bloqueio judicial
#resumo = resumo.drop(columns=resumo.columns[[8]])

# Renomea colunas
resumo.columns = ['data','status_ordem','consorcio','operadora','ordem_pagamento','valor_bruto','valor_taxa','valor_bloqueio_judicial','valor_liquido','valor_debito','qtd_debito','valor_integracao','qtd_integracao','valor_rateio_credito','qtd_rateio_credito','valor_rateio_debito','qtd_rateio_debito','valor_venda_a_bordo','qtd_venda_a_bordo','valor_gratuidade','qtd_gratuidade','id']

# Padronização
resumo['data'] = pd.to_datetime(resumo['data'], format='%Y-%m-%d').dt.strftime('%Y-%m-%d')

# Garantindo que valores inteiros fiquem sem ".0"
colunas_inteiras = ['ordem_pagamento', 'id', 'qtd_debito', 'qtd_integracao','qtd_rateio_credito','qtd_rateio_debito','qtd_venda_a_bordo','qtd_gratuidade'] 

for coluna in colunas_inteiras:
    resumo[coluna] = resumo[coluna].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)

# Salvar no CSV sem converter inteiros para float
arquivo_resumo = DESTINO_PATH / f"{data_ordem} Resumo.csv"

resumo.to_csv(
    arquivo_resumo,
    index=False,
    sep=";",
    encoding="utf-8-sig",
    decimal="."
)



# Configurações principais
project_id = os.getenv("BQ_PROJECT")
dataset_id = os.getenv("BQ_DATASET")
table_id = os.getenv("BQ_TABLE_RESUMO")
source_file = arquivo_resumo

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_ordem", "DATE"),
    bigquery.SchemaField("status_ordem", "STRING"),
    bigquery.SchemaField("consorcio", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("ordem_pagamento", "STRING"),
    bigquery.SchemaField("valor_bruto", "FLOAT64"),
    bigquery.SchemaField("valor_taxa", "FLOAT64"),
    bigquery.SchemaField("valor_bloqueio_judicial", "FLOAT64"),
    bigquery.SchemaField("valor_liquido", "FLOAT64"),
    bigquery.SchemaField("valor_debito", "FLOAT64"),
    bigquery.SchemaField("qtd_debito", "INTEGER"),
    bigquery.SchemaField("valor_integracao", "FLOAT64"),
    bigquery.SchemaField("qtd_integracao", "INTEGER"),
    bigquery.SchemaField("valor_rateio_credito", "FLOAT64"),
    bigquery.SchemaField("qtd_rateio_credito", "INTEGER"),
    bigquery.SchemaField("valor_rateio_debito", "FLOAT64"),
    bigquery.SchemaField("qtd_rateio_debito", "INTEGER"),
    bigquery.SchemaField("valor_venda_a_bordo", "FLOAT64"),
    bigquery.SchemaField("qtd_venda_a_bordo", "INTEGER"),
    bigquery.SchemaField("valor_gratuidade", "FLOAT64"),
    bigquery.SchemaField("qtd_gratuidade", "INTEGER"),
    bigquery.SchemaField("id", "STRING")
]

job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.CSV,
    skip_leading_rows=1,
    autodetect=False,      # Desabilitar autodetecção de tipos
    field_delimiter=';',   # Garantir que o delimitador seja uma vírgula
    schema=schema  # Definir o esquema manualmente
)

# Carrega o arquivo local para o BigQuery
with open(source_file, "rb") as file:
    job = client.load_table_from_file(file, table_ref, job_config=job_config)

# Aguarda o job ser concluído
job.result()




# DOWNLOAD
######################################### RATEIO E TRANSACAO
# FUNCOES
# Função para realizar o download dos arquivos
def baixar_arquivos(tipo):
    """
    Função para baixar arquivos do Power BI.
    :param tipo: Tipo de arquivo a ser baixado ('Rateio' ou 'Transação').
    """
    try:
        # total de linhas
        linhas_tabela = len(WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@role='gridcell' and @column-index='0' and @aria-colindex='2' and not(contains(text(), 'Total'))]"))
        ))
        
        # DEBUG
        print(f"\n🔍 DEBUG: {consorcio_selecionado} - {tipo}")
        print(f"Total de linhas: {linhas_tabela}")
        
        # ALTERAÇÃO: Sempre processa TODAS as linhas, independente do consórcio
        linhas_para_processar = list(range(1, linhas_tabela + 1))
        print(f"\n📋 {consorcio_selecionado} - {tipo}: Processando TODAS as {linhas_tabela} linhas")
        print(f"DEBUG: Linhas que serão processadas: {linhas_para_processar}\n")
        
        # Itera sobre as linhas selecionadas
        for i in linhas_para_processar:
            try:
                print(f"\n{'=' * 60}")
                print(f"📍 Iniciando processamento da linha {i}")
                print(f"{'=' * 60}")
                
                # Localiza os elementos necessários na linha
                linha = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, f"(//div[@role='gridcell' and @column-index='3' and @aria-colindex='5'])[{i}]"))
                )
                texto_linha = linha.text
                print(f"✓ Texto da linha capturado: {texto_linha}")
                time.sleep(7)
                
                consorcio = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, f"(//div[@role='gridcell' and @column-index='2' and @aria-colindex='4'])[{i}]"))
                )
                texto_consorcio = consorcio.text
                print(f"✓ Consórcio capturado: {texto_consorcio}")
                time.sleep(7)
                
                data_linha = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, f"(//div[@role='gridcell' and @column-index='0' and @aria-colindex='2'])[{i}]"))
                )
                data_linha = data_linha.text.replace('/', '-')
                print(f"✓ Data capturada: {data_linha}")
                time.sleep(7)
                
                # Clica na linha
                print("⏳ Clicando na linha...")
                driver.execute_script("arguments[0].scrollIntoView(true);", linha)
                time.sleep(2)
                
                try:
                    linha.click()
                except Exception as click_error:
                    print(f"⚠️  Erro ao clicar normalmente, tentando com JavaScript...")
                    driver.execute_script("arguments[0].click();", linha)
                
                print("✓ Linha clicada com sucesso")
                time.sleep(7)
                
                # ALTERAÇÃO: Verifica download duplo SOMENTE para JABOUR (Santa Cruz) e REDENTOR (Transcarioca)
                precisa_duplo_download = False
                if tipo == "Transação":
                    if (texto_consorcio.upper() == "SANTA CRUZ" and "JABOUR" in texto_linha.upper()) or \
                       (texto_consorcio.upper() == "TRANSCARIOCA" and "REDENTOR" in texto_linha.upper()):
                        precisa_duplo_download = True
                        print("✓ Empresa identificada - download duplo necessário")
                        if "JABOUR" in texto_linha.upper():
                            print("   -> Empresa: JABOUR (Santa Cruz)")
                        if "REDENTOR" in texto_linha.upper():
                            print("   -> Empresa: REDENTOR (Transcarioca)")
                
                # Clica no botão de drill-through
                print(f"⏳ Procurando botão drill-through para {tipo}...")
                botao_drill = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, f"//*[@aria-label='Drill-through . Clique aqui para executar uma consulta drill-through em {tipo}']"))
                )
                
                try:
                    botao_drill.click()
                except Exception:
                    print("⚠️  Clique normal falhou, usando JavaScript...")
                    driver.execute_script("arguments[0].click();", botao_drill)
                
                print("✓ Botão drill-through clicado")
                time.sleep(10)
                
                # Move o mouse para revelar "Mais opções"
                print("⏳ Movendo mouse para revelar 'Mais opções'...")
                elemento_para_revelar = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, f"(//div[@title='{tipo}'])"))
                )
                actions = ActionChains(driver)
                actions.move_to_element(elemento_para_revelar).perform()
                time.sleep(7)
                
                # ==== PRIMEIRO DOWNLOAD (ORDENADO) ====
                if precisa_duplo_download:
                    print("\n" + "=" * 60)
                    print("📥 INICIANDO PRIMEIRO DOWNLOAD (ORDENADO)")
                    print("=" * 60)
                    
                    # 🔽 Ordenação SOMENTE para Transação
                    if tipo == "Transação":
                        print("⏳ Aplicando ordenação (somente Transação)...")
                        WebDriverWait(driver, 60).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[@class='powervisuals-glyph sort-icon caret-down ']"))
                            ).click()
                        time.sleep(7)

                    
                    # Captura estado atual da pasta antes do download
                    existentes = {f for f in DOWNLOADS_PATH.iterdir() if f.is_file()}
                    print(f"✓ Arquivos atuais na pasta: {len(existentes)}")
                    
                    # Clica em "Mais opções"
                    print("⏳ Clicando em 'Mais opções'...")
                    botao_mais_opcoes = WebDriverWait(driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@class='vcMenuBtn' and @aria-label='Mais opções']"))
                    )
                    
                    try:
                        botao_mais_opcoes.click()
                    except Exception:
                        print("⚠️  Clique normal falhou, usando JavaScript...")
                        driver.execute_script("arguments[0].click();", botao_mais_opcoes)
                    
                    print("✓ 'Mais opções' clicado")
                    time.sleep(7)
                    
                    
                    # Clica em "Exportar dados"
                    print("⏳ Clicando em 'Exportar dados'...")
                    WebDriverWait(driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']"))
                    ).click()
                    print("✓ 'Exportar dados' clicado")
                    time.sleep(7)
                    
                    # Clica no botão "Exportar"
                    print("⏳ Clicando no botão 'Exportar'...")
                    WebDriverWait(driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Exportar']"))
                    ).click()
                    print("✓ Botão 'Exportar' clicado")
                    time.sleep(7)
                    
                    # Espera o download terminar
                    print("⏳ Aguardando conclusão do download ordenado...")
                    arquivo_ordenado = aguardar_download(existentes)
                    
                    if arquivo_ordenado:
                        # Verificação: Conta as linhas do arquivo
                        print("\n" + "=" * 60)
                        print("🔍 VERIFICANDO NÚMERO DE LINHAS DO ARQUIVO")
                        print("=" * 60)
                        
                        try:
                            df_temp = pd.read_excel(arquivo_ordenado)
                            total_linhas = len(df_temp)
                            print(f"📊 Total de linhas no arquivo ordenado: {total_linhas}")
                            
                            # Define se precisa do segundo download baseado no número de linhas
                            if total_linhas < 150003:
                                print(f"✅ Arquivo tem MENOS de 150.003 linhas ({total_linhas})")
                                print("📌 NÃO será necessário segundo download - arquivo único será usado")
                                precisa_duplo_download = False
                                
                                # Salva como arquivo único (sem sufixo _ordenado)
                                DESTINO_PATH.mkdir(parents=True, exist_ok=True)
                                texto_linha_limpo = texto_linha.replace('/', '')
                                nome_unico = f"{data_linha} {texto_linha_limpo} - {texto_consorcio} - {tipo}.xlsx"
                                caminho_unico = os.path.join(diretorio_destino, nome_unico)
                                shutil.move(str(arquivo_ordenado), caminho_unico)
                                print(f"✅ Arquivo salvo como único: {nome_unico}")
                            else:
                                print(f"⚠️  Arquivo tem {total_linhas} linhas (≥ 150.003)")
                                print("📌 Segundo download (consolidado) SERÁ necessário")
                                
                                # Salva como arquivo ordenado
                                os.makedirs(diretorio_destino, exist_ok=True)
                                texto_linha_limpo = texto_linha.replace('/', '')
                                nome_ordenado = f"{data_linha} {texto_linha_limpo} - {texto_consorcio} - {tipo}_ordenado.xlsx"
                                caminho_ordenado = os.path.join(diretorio_destino, nome_ordenado)
                                shutil.move(str(arquivo_ordenado), caminho_ordenado)
                                print(f"✅ Arquivo ordenado salvo: {nome_ordenado}")
                        
                        except Exception as e:
                            print(f"❌ ERRO ao verificar linhas do arquivo: {e}")
                            print("⚠️  Continuando com download duplo por segurança...")
                            os.makedirs(diretorio_destino, exist_ok=True)
                            texto_linha_limpo = texto_linha.replace('/', '')
                            nome_ordenado = f"{data_linha} {texto_linha_limpo} - {texto_consorcio} - {tipo}_ordenado.xlsx"
                            caminho_ordenado = os.path.join(diretorio_destino, nome_ordenado)
                            shutil.move(str(arquivo_ordenado), caminho_ordenado)
                            print(f"✅ Arquivo ordenado salvo: {nome_ordenado}")
                    else:
                        print(f"❌ ERRO: Download ordenado não concluído para linha {i}")
                        continue

                

                # ==== SEGUNDO DOWNLOAD (CONSOLIDADO ou ÚNICO) ====
                sufixo = "_consolidado" if precisa_duplo_download else ""
                print("\n" + "=" * 60)
                print(f"📥 INICIANDO {'SEGUNDO' if precisa_duplo_download else ''} DOWNLOAD{' (CONSOLIDADO)' if precisa_duplo_download else ''}")
                print("=" * 60)
                
                # 🔽 Ordenação SOMENTE para Transação
                if tipo == "Transação":
                    print("⏳ Preparando para reordenar...")
                
                    # Mover o mouse novamente para garantir que os ícones apareçam
                    elemento_para_revelar = WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, f"(//div[@title='{tipo}'])"))
                    )
                    actions = ActionChains(driver)
                    actions.move_to_element(elemento_para_revelar).perform()
                    time.sleep(3)
                
                    print("⏳ Clicando na ordenação novamente...")
                    WebDriverWait(driver, 60).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'powervisuals-glyph') and contains(@class, 'sort-icon')]"))
                    ).click()
                    time.sleep(7)
          

                # Captura estado atual da pasta antes do download
                existentes = {f for f in DOWNLOADS_PATH.iterdir() if f.is_file()}
                print(f"✓ Arquivos atuais na pasta: {len(existentes)}")
                
                # Clica em "Mais opções" e "Exportar"
                print("⏳ Clicando em 'Mais opções'...")
                botao_mais_opcoes = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@class='vcMenuBtn' and @aria-label='Mais opções']"))
                )
                
                try:
                    botao_mais_opcoes.click()
                except Exception:
                    print("⚠️  Clique normal falhou, usando JavaScript...")
                    driver.execute_script("arguments[0].click();", botao_mais_opcoes)
                
                print("✓ 'Mais opções' clicado")
                time.sleep(7)
                
                print("⏳ Clicando em 'Exportar dados'...")
                WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']"))
                ).click()
                print("✓ 'Exportar dados' clicado")
                time.sleep(7)
                
                print("⏳ Clicando no botão 'Exportar'...")
                WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Exportar']"))
                ).click()
                print("✓ Botão 'Exportar' clicado")
                time.sleep(7)
                
                # Espera o download terminar
                print("⏳ Aguardando conclusão do download...")
                arquivo_baixado = aguardar_download(existentes)
                
                if arquivo_baixado:
                    # Renomeia e move o arquivo
                    os.makedirs(diretorio_destino, exist_ok=True)
                    texto_linha_limpo = texto_linha.replace('/', '')
                    nome_arquivo_novo = f"{data_linha} {texto_linha_limpo} - {texto_consorcio} - {tipo}{sufixo}.xlsx"
                    caminho_novo = os.path.join(diretorio_destino, nome_arquivo_novo)
                    shutil.move(str(arquivo_baixado), caminho_novo)
                    print(f"✅ Arquivo salvo: {nome_arquivo_novo}")
                else:
                    print(f"❌ ERRO: Download não concluído para linha {i}")
                    print("📁 Arquivos atuais na pasta de downloads:")
                    for f in DOWNLOADS_PATH(pasta).iterdir():
                        if f.is_file():
                            print(f"  - {f.name}")
                    break
                
                # Voltar à página anterior
                print("⏳ Voltando à página anterior...")
                WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Voltar . Clique aqui para voltar à página anterior neste relatório']"))
                ).click()
                print("✓ Voltou à página anterior")
                time.sleep(10)
                
                print(f"✅ Linha {i} processada com sucesso!")
            
            except Exception as e:
                print(f"\n❌ ERRO ao processar linha {i}: {e}")
                import traceback
                traceback.print_exc()
                
                # Tenta voltar à página anterior mesmo em caso de erro
                try:
                    print("⏳ Tentando voltar à página anterior após erro...")
                    WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Voltar . Clique aqui para voltar à página anterior neste relatório']"))
                    ).click()
                    time.sleep(10)
                except:
                    print("⚠️  Não foi possível voltar à página anterior")
                
                break
        
        # Validação final
        print("\n" + "=" * 60)
        print("📊 VALIDAÇÃO FINAL")
        print("=" * 60)
        
        arquivos_baixados = len(list(DESTINO_PATH.glob(f"{data_ordem}*{consorcio_selecionado} - {tipo}.xlsx")))
        arquivos_ordenados = len(list(DESTINO_PATH.glob(f"{data_ordem}*{consorcio_selecionado} - {tipo}_ordenado.xlsx")))
        arquivos_consolidados = len(list(DESTINO_PATH.glob(f"{data_ordem}*{consorcio_selecionado} - {tipo}_consolidado.xlsx")))

        
        total_arquivos = arquivos_baixados + arquivos_ordenados + arquivos_consolidados
        linhas_esperadas = len(linhas_para_processar)
        
        print(f"Arquivos normais: {arquivos_baixados}")
        print(f"Arquivos ordenados: {arquivos_ordenados}")
        print(f"Arquivos consolidados: {arquivos_consolidados}")
        print(f"Total de arquivos: {total_arquivos}")
        print(f"Linhas esperadas: {linhas_esperadas}")
        
        if total_arquivos < linhas_esperadas:
            print(f"\n⚠️  Download INCOMPLETO do consórcio {consorcio_selecionado} de {tipo}!")
            print(f"Esperado: {linhas_esperadas} | Obtido: {total_arquivos}")
        else:
            print(f"\n✅ Download completo do consórcio {consorcio_selecionado} de {tipo}.")
    
    except Exception as e:
        print(f"\n❌ ERRO FATAL na função baixar_arquivos({tipo}): {e}")
        import traceback
        traceback.print_exc()

def aguardar_download(existentes):
    """
    Aguarda o download ser concluído e retorna o arquivo baixado.
    """
    download_timeout = 520
    poll_interval = 1.0
    fim = time.time() + download_timeout
    arquivo_mais_recente = None

    while time.time() < fim:
        atuais = {f for f in DOWNLOADS_PATH(pasta).iterdir() if f.is_file()}
        novos = atuais - existentes
        if not novos:
            time.sleep(poll_interval)
            continue

        # ignora temporários
        if any(f.suffix in {".crdownload", ".part"} for f in novos):
            time.sleep(poll_interval)
            continue

        candidatos = [f for f in novos if f.suffix.lower() == ".xlsx"]
        if not candidatos:
            time.sleep(poll_interval)
            continue

        # espera estabilidade de tamanho
        estabilizado = False
        for candidato in candidatos:
            tamanho_anterior = -1
            stable_since = time.time()
            while time.time() < fim:
                try:
                    atual_size = candidato.stat().st_size
                except FileNotFoundError:
                    break
                if atual_size == tamanho_anterior:
                    if time.time() - stable_since >= poll_interval:
                        arquivo_mais_recente = candidato
                        estabilizado = True
                        break
                else:
                    tamanho_anterior = atual_size
                    stable_since = time.time()
                time.sleep(poll_interval)
            if estabilizado:
                break
        if estabilizado:
            break
        time.sleep(poll_interval)

    return arquivo_mais_recente

# Função para realizar o filtro por consorcio
def selecionar_consorcio(consorcio):
    """
    Seleciona um consórcio no filtro e baixa os arquivos de Rateio e Transação.
    :param consorcio: Nome do consórcio a ser filtrado.
    """
    
    global consorcio_selecionado
    consorcio_selecionado = consorcio
    
    wait = WebDriverWait(driver, 60)
    
    # Abre o filtro
    #filtro = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']")))
    filtro = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']")))
    filtro.click()
    time.sleep(7)
    
    # Seleciona o campo de consórcio
    campo_consorcio = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@class='slicer-dropdown-menu' and @aria-label='Nm_Consorcio']")))
    campo_consorcio.click()
    time.sleep(7)
    
    # Seleciona o consórcio desejado
    consorcio_escolhido = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[contains(@class, 'slicerText') and contains(text(), '{consorcio}')]")))
    consorcio_escolhido.click()
    time.sleep(7)
    
    # Fecha o filtro clicando novamente no ícone
    #indicador = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']")))
    indicador = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']")))
    indicador.click()
    
    # Chama função para baixar os arquivos
    baixar_arquivos("Rateio")
    baixar_arquivos("Transação")
    
    
# Chama função de seleção de consórcio
selecionar_consorcio("Internorte")
selecionar_consorcio("Intersul")
selecionar_consorcio("Santa Cruz")
#selecionar_consorcio("Transcarioca")





    
# Fecha o navegador
driver.close()  

#CONSOLIDANDO OS ARQUIVOS BAIXADOS EM UM ARQUIVO E SUBINDO NA TABELA DO BIGQUERY
#TRANSACAO
termo = 'Transação'

def extrair_id_do_excel(df, arquivo):
    # Procura o ID nas últimas 15 linhas da primeira coluna
    for i in range(1, 16):
        try:
            texto = str(df.iloc[-i, 0])
        except:
            continue

        match = re.search(r'\bid\s*é\s*(\d+)', texto, re.IGNORECASE)
        if match:
            return match.group(1)

    print(f"⚠️ ID NÃO encontrado no arquivo: {arquivo}")
    return None

# Lista para armazenar os DataFrames
lista_dataframes = []

# Dicionário para armazenar pares ordenado/consolidado
arquivos_para_consolidar = {}
bases_com_par = set()

# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    # IGNORA ARQUIVOS TEMPORÁRIOS DO EXCEL (começam com ~$)
    if arquivo.startswith('~$'):
        continue
    
    # IGNORA ARQUIVOS QUE CONTENHAM "RESUMO" NO NOME
    if 'resumo' in arquivo.lower():
        print(f"⏭️ Ignorando arquivo resumo: {arquivo}")
        continue
        
    if arquivo.endswith('.xlsx') and termo in arquivo and str(data_ordem) in arquivo:
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        
        # Verifica se é arquivo ordenado ou consolidado
        if 'ordenado' in arquivo or 'consolidado' in arquivo:
            # Extrai a chave base do arquivo (remove ordenado/consolidado do nome)
            chave_base = arquivo.replace('_ordenado', '').replace('_consolidado', '')
            
            # Lê o arquivo completo primeiro
            df_completo_temp = pd.read_excel(
                caminho_completo,
                dtype={
                    "Nr Linha": str,
                    "Ordem Ressarcimento": str,
                    "Prefixo Veículo": str
                }
            )
            
            print(f"Arquivo encontrado: {arquivo}")
            print(f"Total de linhas ANTES de remover últimas 3: {len(df_completo_temp)}")
            
            # Extrai o ID ANTES de remover as linhas
            id_extraido = extrair_id_do_excel(df_completo_temp, arquivo)

            
            # Agora remove as últimas 3 linhas
            df_temp = df_completo_temp.iloc[:-3]
            print(f"Total de linhas APÓS remover últimas 3: {len(df_temp)}")
            
            # Armazena no dicionário junto com o ID
            if chave_base not in arquivos_para_consolidar:
                arquivos_para_consolidar[chave_base] = {'id': id_extraido}
            
            if 'ordenado' in arquivo:
                arquivos_para_consolidar[chave_base]['ordenado'] = df_temp
            elif 'consolidado' in arquivo:
                arquivos_para_consolidar[chave_base]['consolidado'] = df_temp
        else:
            # Processa arquivos normais (sem ordenado/consolidado)
            chave_base = arquivo  # ou mesma lógica usada acima
        
            if chave_base in bases_com_par:
                print(f"⛔ Arquivo base ignorado (existe ordenado+consolidado): {arquivo}")
                continue
        
            df = pd.read_excel(
                caminho_completo,
                dtype={
                    "Nr Linha": str,
                    "Ordem Ressarcimento": str,
                    "Prefixo Veículo": str
                }
            )
        
            df['Data'] = data_ordem
            valor_ultima_linha = df.iloc[-1, 0]
            df['Id'] = re.search(r'Filtros aplicados:\nid é (\d+)', valor_ultima_linha).group(1)
            df = df.iloc[:-3]
        
            lista_dataframes.append(df)

# Processa arquivos ordenado/consolidado
# Processa arquivos ordenado/consolidado (ESPECIAIS)
for chave_base, arquivos in arquivos_para_consolidar.items():
    if 'ordenado' in arquivos and 'consolidado' in arquivos:
        print(f"\n=== CONSOLIDANDO ARQUIVOS ESPECIAIS: {chave_base} ===")
        print(f"Linhas no arquivo ordenado: {len(arquivos['ordenado'])}")
        print(f"Linhas no arquivo consolidado: {len(arquivos['consolidado'])}")
        
        # Junta SOMENTE ordenado + consolidado
        df_completo = pd.concat(
            [arquivos['ordenado'], arquivos['consolidado']],
            ignore_index=True
        )
        print(f"Total de linhas após junção: {len(df_completo)}")
        
        # Remove duplicatas ENTRE ELES
        df_completo = df_completo.drop_duplicates()
        print(f"Total de linhas após remover duplicatas: {len(df_completo)}")
        
        # Metadados
        df_completo['Data'] = data_ordem
        df_completo['Id'] = arquivos['id']

        # 🔹 SALVA ARQUIVO FINAL DO ESPECIAL
        nome_base = chave_base.replace('.xlsx', '')
        arquivo_especial = (
            DESTINO_PATH / f"{data_ordem}_{nome_base}_Transacao_ESPECIAL.csv"
        )

        df_completo.to_csv(
            arquivo_especial,
            index=False,
            sep=";",
            encoding="utf-8-sig",
            decimal="."
        )

        print(f"✅ Arquivo ESPECIAL criado: {arquivo_especial}")
        print(f"📊 Linhas no arquivo especial: {len(df_completo)}")

        # ✅ ADICIONA À LISTA PARA IR PRO BIGQUERY
        lista_dataframes.append(df_completo)
        print(f"✅ Arquivo ESPECIAL adicionado à lista principal!")

    else:
        print(f"\n⚠️ AVISO: Arquivo especial incompleto ignorado: {chave_base}")

# Concatena todos os DataFrames em um único
# Concatena todos os DataFrames em um único
transacao = pd.concat(lista_dataframes, ignore_index=True)

print(f"\n=== CONTAGEM ANTES DE REMOVER DUPLICATAS FINAIS ===")
print(f"Total de linhas antes: {len(transacao)}")

# Remove duplicatas do DataFrame final
transacao = transacao.drop_duplicates()

print(f"\n=== CONTAGEM APÓS REMOVER DUPLICATAS FINAIS ===")
print(f"Total de linhas SEM DUPLICATAS: {len(transacao)}")

# Organiza a ordem das colunas do DataFrame
# ✅ ORDEM CORRIGIDA (Data e Id estão no final nos arquivos especiais)
transacao = transacao[['Data Transação','Data Processamento','Ordem Ressarcimento','Consórcio','Operadora','Nr Linha','Modal','Linha','Prefixo Veículo','Validador','Tipo Transação','Tipo Usuário','Produto','Tipo Produto','Mídia','Transação','Qtd Transação','Valor Tarifa','Valor Transação','Data','Id']]

# Reordena para colocar Data e Id no início
transacao = transacao[['Data','Consórcio','Id','Data Transação','Data Processamento','Ordem Ressarcimento','Operadora','Nr Linha','Modal','Linha','Prefixo Veículo','Validador','Tipo Transação','Tipo Usuário','Produto','Tipo Produto','Mídia','Transação','Qtd Transação','Valor Tarifa','Valor Transação']]

# Renomea colunas
transacao.columns = ['data_ordem','consorcio','id','data_transacao','data_processamento','ordem_ressarcimento','operadora','servico','modal','linha','prefixo_veiculo','validador','tipo_transacao','tipo_usuario','produto','tipo_produto','midia','id_transacao','qtd_transacao','valor_tarifa','valor_transacao']

# Padronização
transacao['data_ordem'] = pd.to_datetime(transacao['data_ordem'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d')
transacao['data_transacao'] = pd.to_datetime(transacao['data_transacao'], format='%d/%m/%Y %H:%M:%S', errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')

# Garantindo que os valores monetários tenham casas decimais
colunas_float = ['valor_tarifa', 'valor_transacao']

for coluna in colunas_float:
    transacao[coluna] = transacao[coluna].astype(float)
    
# Garantindo que valores inteiros fiquem sem ".0"
colunas_inteiras = ['qtd_transacao']

for coluna in colunas_inteiras:
    transacao[coluna] = transacao[coluna].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)
    
# Garantindo que valores string sem ".0"
for coluna in ['ordem_ressarcimento', 'prefixo_veiculo']:
    transacao[coluna] = transacao[coluna].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0', '').isdigit() else str(x)
    )

# CRIA O ARQUIVO CONSOLIDADO NA PASTA BASES_RESSARCIMENTO
arquivo_consolidado = DESTINO_PATH / f"{data_ordem}_Transacao_Consolidado.csv"
transacao.to_csv(arquivo_consolidado, index=False, sep=";", encoding="utf-8-sig", decimal=".")

print(f"\n✅ Arquivo consolidado criado: {arquivo_consolidado}")
print(f"📊 Total de linhas no arquivo (sem duplicatas): {len(transacao)}")

# Configurações principais
project_id = "ro-areatecnica"
dataset_id = "ressarcimento_jae"
table_id = "transacao"
source_file = arquivo_consolidado  # USA O ARQUIVO CONSOLIDADO CRIADO

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_ordem", "DATE"),
    bigquery.SchemaField("consorcio", "STRING"),
    bigquery.SchemaField("id", "STRING"),
    bigquery.SchemaField("data_transacao", "DATETIME"),
    bigquery.SchemaField("data_processamento", "DATETIME"),
    bigquery.SchemaField("ordem_ressarcimento", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("servico", "STRING"),  # ← Mude de nr_linha para servico
    bigquery.SchemaField("modal", "STRING"),
    bigquery.SchemaField("linha", "STRING"),
    bigquery.SchemaField("prefixo_veiculo", "STRING"),
    bigquery.SchemaField("validador", "STRING"),
    bigquery.SchemaField("tipo_transacao", "STRING"),
    bigquery.SchemaField("tipo_usuario", "STRING"),
    bigquery.SchemaField("produto", "STRING"),
    bigquery.SchemaField("tipo_produto", "STRING"),
    bigquery.SchemaField("midia", "STRING"),
    bigquery.SchemaField("id_transacao", "STRING"),
    bigquery.SchemaField("qtd_transacao", "INTEGER"),
    bigquery.SchemaField("valor_tarifa", "FLOAT64"),
    bigquery.SchemaField("valor_transacao", "FLOAT64")
]

job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.CSV,
    skip_leading_rows=1,
    autodetect=False,
    field_delimiter=';',
    schema=schema
)

print(f"\n🚀 Enviando arquivo para o BigQuery: {source_file}")

# Carrega o arquivo local para o BigQuery
with open(source_file, "rb") as file:
    job = client.load_table_from_file(file, table_ref, job_config=job_config)

# Aguarda o job ser concluído
job.result()

# Verifique os erros detalhados
if job.errors:
    print(f"❌ Erros durante o carregamento: {job.errors}")
else:
    print(f"✅ Arquivo {source_file} carregado com sucesso para {table_ref}!")
    print(f"📊 Total de linhas enviadas: {len(transacao)}")

#RATEIO
#data = '18-02-2025'
termo = 'Rateio'

        
# Lista para armazenar os DataFrames
lista_dataframes = []

# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    # IGNORA ARQUIVOS TEMPORÁRIOS DO EXCEL
    if arquivo.startswith('~$'):
        continue
        
    if arquivo.endswith('.xlsx') and termo in arquivo and str(data_ordem) in arquivo: 
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        df = pd.read_excel(caminho_completo)
        df['Data'] = data_ordem
        
        if 'Internorte' in arquivo:
            df['Consorcio'] = 'Internorte'
        if 'Santa Cruz' in arquivo:
            df['Consorcio'] = 'Santa Cruz'
        if 'Intersul' in arquivo:
            df['Consorcio'] = 'Intersul'
        if 'Transcarioca' in arquivo:
            df['Consorcio'] = 'Transcarioca'
            
        df['Nome_Arquivo'] = arquivo  # Adiciona o nome do arquivo ao DataFrame
        
        # Expressão regular para extrair o nome da empresa (parte entre a data e o primeiro "-")
        match = re.search(r'\d{2}-\d{2}-\d{4} (.*?) -', arquivo)
        if match:
            df['Operadora'] = match.group(1)  # Adiciona a empresa ao DataFrame
            
        df["Operadora"] = df["Operadora"].str.replace(" SA", " S/A", regex=False)
            
        valor_ultima_linha = df.iloc[-1, 0]                                                         #seleciona o valor do primeiro campo da última linha
        df['Id'] = re.search(r'Filtros aplicados:\nid é (\d+)', valor_ultima_linha).group(1)        #mantem apenas o valor do id
        df = df.iloc[:-2]                                                                           #exclui as últimas 2 linhas do df
        lista_dataframes.append(df)  # Adicionar o DataFrame à lista


# Concatena todos os DataFrames em um único
rateio = pd.concat(lista_dataframes, ignore_index=True)

# Organiza a ordem das colunas do DatFrame
rateio = rateio[['Data','Consorcio', 'Operadora','Id','Data P1','Modal P1','Linha P1','Rateio P1','% P1','Transação P1','Data P2','Modal P2','Linha P2','Rateio P2','% P2','Transação P2','Data P3','Modal P3','Linha P3','Rateio P3','% P3','Transação P3','Data P4','Modal P4','Linha P4','Rateio P4','% P4','Transação P4','Data P5','Modal P5','Linha P5','Rateio P5','% P5','Transação P5']]

# Renomea colunas
rateio.columns =['data_ordem','consorcio', 'operadora','id','data_p1','modal_p1','linha_p1','rateio_p1','percentual_p1','id_transacao_p1','data_p2','modal_p2','linha_p2','rateio_p2','percentual_p2','id_transacao_p2','data_p3','modal_p3','linha_p3','rateio_p3','percentual_p3','id_transacao_p3','data_p4','modal_p4','linha_p4','rateio_p4','percentual_p4','id_transacao_p4','data_p5','modal_p5','linha_p5','rateio_p5','percentual_p5','id_transacao_p5']

# Padronização
rateio['data_ordem'] = pd.to_datetime(rateio['data_ordem'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d')
rateio['percentual_p1'] = rateio['percentual_p1'].astype(float)
rateio['rateio_p1'] = rateio['rateio_p1'].astype(float)
rateio['percentual_p2'] = rateio['percentual_p2'].astype(float)
rateio['rateio_p2'] = rateio['rateio_p2'].astype(float)
rateio['percentual_p3'] = rateio['percentual_p3'].astype(float)
rateio['rateio_p3'] = rateio['rateio_p3'].astype(float)
rateio['percentual_p4'] = rateio['percentual_p4'].astype(float)
rateio['rateio_p4'] = rateio['rateio_p4'].astype(float)
rateio['percentual_p5'] = rateio['percentual_p5'].astype(float)
rateio['rateio_p5'] = rateio['rateio_p5'].astype(float)


# Garantindo que os valores monetários tenham 2 casas decimais
colunas_float = ['rateio_p1', 'percentual_p1','rateio_p2', 'percentual_p2','rateio_p3', 'percentual_p3','rateio_p4', 'percentual_p4','rateio_p5', 'percentual_p5']

for coluna in colunas_float:
    rateio[coluna] = rateio[coluna].astype(float)

# Salvar no CSV sem converter inteiros para float
arquivo_rateio = DESTINO_PATH / f"{data_ordem} Rateio.csv"

rateio.to_csv(
    arquivo_rateio,
    index=False,
    sep=";",
    encoding="utf-8-sig",
    decimal="."
)

source_file = arquivo_rateio

# Configurações do BigQuery
project_id = "ro-areatecnica"
dataset_id = "ressarcimento_jae"
table_id = "rateio"


# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_ordem", "DATE"),
    bigquery.SchemaField("consorcio", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("id", "STRING"),
    bigquery.SchemaField("data_p1", "DATETIME"),
    bigquery.SchemaField("modal_p1", "STRING"),
    bigquery.SchemaField("linha_p1", "STRING"),
    bigquery.SchemaField("rateio_p1", "FLOAT64"),
    bigquery.SchemaField("percentual_p1", "FLOAT64"),
    bigquery.SchemaField("id_transacao_p1", "STRING"),
    bigquery.SchemaField("data_p2", "DATETIME"),
    bigquery.SchemaField("modal_p2", "STRING"),
    bigquery.SchemaField("linha_p2", "STRING"),
    bigquery.SchemaField("rateio_p2", "FLOAT64"),
    bigquery.SchemaField("percentual_p2", "FLOAT64"),
    bigquery.SchemaField("id_transacao_p2", "STRING"),
    bigquery.SchemaField("data_p3", "DATETIME"),
    bigquery.SchemaField("modal_p3", "STRING"),
    bigquery.SchemaField("linha_p3", "STRING"),
    bigquery.SchemaField("rateio_p3", "FLOAT64"),
    bigquery.SchemaField("percentual_p3", "FLOAT64"),
    bigquery.SchemaField("id_transacao_p3", "STRING"),
    bigquery.SchemaField("data_p4", "DATETIME"),
    bigquery.SchemaField("modal_p4", "STRING"),
    bigquery.SchemaField("linha_p4", "STRING"),
    bigquery.SchemaField("rateio_p4", "FLOAT64"),
    bigquery.SchemaField("percentual_p4", "FLOAT64"),
    bigquery.SchemaField("id_transacao_p4", "STRING"),
    bigquery.SchemaField("data_p5", "DATETIME"),
    bigquery.SchemaField("modal_p5", "STRING"),
    bigquery.SchemaField("linha_p5", "STRING"),
    bigquery.SchemaField("rateio_p5", "FLOAT64"),
    bigquery.SchemaField("percentual_p5", "FLOAT64"),
    bigquery.SchemaField("id_transacao_p5", "STRING")
]

job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.CSV,
    skip_leading_rows=1,
    autodetect=False,      # Desabilitar autodetecção de tipos
    field_delimiter=';',   # Garantir que o delimitador seja uma vírgula
    schema=schema  # Definir o esquema manualmente
)


# Carrega o arquivo local para o BigQuery
with open(source_file, "rb") as file:
    job = client.load_table_from_file(file, table_ref, job_config=job_config)

# Aguarda o job ser concluído
job.result()

# Verifique os erros detalhados
if job.errors:
    print(f"Erros durante o carregamento: {job.errors}")
else:
    print(f"Arquivo {source_file} carregado com sucesso para {table_ref}!")
    

# Insere o rateio baixado na taela de distintos
rateio_distinto = """
INSERT INTO `ro-areatecnica.ressarcimento_jae.rateio_distinto` (
   data_ordem, data_p1, modal_p1, linha_p1, rateio_p1, percentual_p1, id_transacao_p1,
   data_p2, modal_p2, linha_p2, rateio_p2, percentual_p2, id_transacao_p2,
   data_p3, modal_p3, linha_p3, rateio_p3, percentual_p3, id_transacao_p3,
   data_p4, modal_p4, linha_p4, rateio_p4, percentual_p4, id_transacao_p4,
   data_p5, modal_p5, linha_p5, rateio_p5, percentual_p5, id_transacao_p5
)
SELECT DISTINCT
   data_ordem, data_p1, modal_p1, linha_p1, rateio_p1, percentual_p1, id_transacao_p1,
   data_p2, modal_p2, linha_p2, rateio_p2, percentual_p2, id_transacao_p2,
   data_p3, modal_p3, linha_p3, rateio_p3, percentual_p3, id_transacao_p3,
   data_p4, modal_p4, linha_p4, rateio_p4, percentual_p4, id_transacao_p4,
   data_p5, modal_p5, linha_p5, rateio_p5, percentual_p5, id_transacao_p5
FROM `ro-areatecnica.ressarcimento_jae.rateio`
WHERE data_ordem = @data_bq
"""

# Configuração do parâmetro
job_config = bigquery.QueryJobConfig(
    query_parameters=[
        bigquery.ScalarQueryParameter("data_bq", "DATE", data_bq)
    ]
)


# Executa a consulta
rateio_distinto_job = client.query(rateio_distinto, job_config=job_config)
rateio_distinto_job.result()