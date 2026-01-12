from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
import os
import pandas as pd
import re
from pathlib import Path
from selenium.webdriver.chrome.options import Options
from google.cloud import bigquery
import shutil
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / "config" / ".env"

load_dotenv(dotenv_path=ENV_PATH, override=True)

def clicar_exportar(driver, tentativas=5):
    for tentativa in range(1, tentativas + 1):
        try:
            print(f"Tentativa {tentativa}/{tentativas} de encontrar 'Exportar'")

            # garante iframe correto
            entrar_no_iframe(driver)

            # move mouse para forçar render
            elemento = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "(//div[@role='columnheader'])[1]"
                ))
            )
            ActionChains(driver).move_to_element(elemento).perform()
            time.sleep(2)

            botao_exportar = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Exportar']"))
            )
            botao_exportar.click()
            print("Exportar clicado com sucesso")
            return

        except TimeoutException:
            print("Botão Exportar não apareceu")
            time.sleep(5)

    raise TimeoutException("Não foi possível clicar em Exportar após várias tentativas")

url = os.getenv('POWERBI_URL')
email = os.getenv('POWERBI_EMAIL')
senha = os.getenv('POWERBI_PASSWORD')
USUARIO = os.getenv('USUARIO')
key_path = os.getenv("KEY_PATH")


data_atual = os.getenv('DATA_ATUAL')
data_bq = os.getenv('DATA_BQ')
data_ordem = os.getenv('DATA_ORDEM')

pasta = f"C:/Users/{USUARIO}/Downloads"
diretorio_destino = f"C:/Users/{USUARIO}/Desktop/Bases_Ressarcimento_Linha"
driver_path = f"C:/Users/{USUARIO}/Downloads/chromedriver-win64/chromedriver.exe"
key_path = f"C:/Users/{USUARIO}/Documents/{key_path}"


caminho_pasta = Path(pasta)


pasta_final = Path(diretorio_destino)


service = Service(driver_path)

client = bigquery.Client.from_service_account_json(key_path)

chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Remove flag de automação
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  # Evita detecção
# chrome_options.add_argument("--headless=new")  # Usa o novo modo headless
chrome_options.add_argument("--disable-gpu")  # Evita problemas gráficos
chrome_options.add_argument("--start-maximized")  # Abre o navegador maximizado
chrome_options.add_argument("--no-sandbox")  # Evita problemas de permissão
chrome_options.add_argument("--disable-dev-shm-usage")  # Melhora o desempenho em sistemas com pouca RAM

prefs = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True
}
chrome_options.add_experimental_option("prefs", prefs)


driver = webdriver.Chrome(service=service, options=chrome_options)

# ACESSO AO DASHBOARD
# Acessa o link
driver.get(url)

# Login no Power BI
# Preenche email
email_input = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.ID, "Email"))
)

email_input.send_keys(email)
#email_input.send_keys(Keys.RETURN)
time.sleep(5)

# Preenche senha
senha_input = driver.find_element(By.ID, 'Password')
senha_input.send_keys(senha)  # Senha
senha_input.send_keys(Keys.RETURN)

#driver.switch_to.default_content()

# Esperar até o iframe com o src correto estar disponível
iframe = WebDriverWait(driver, 120).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'https://app.powerbi.com/reportEmbed?reportId=a6abefa3-5bc8')]"))
)

# Mudar para o iframe usando o src
driver.switch_to.frame(iframe)
time.sleep(10)

def entrar_no_iframe(driver):
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((
            By.XPATH,
            "//iframe[contains(@src, 'reportEmbed')]"
        ))
    )
    driver.switch_to.frame(iframe)


# FILTRO DE PERIODO
# Localiza e clica no elemento de filtro
entrar_no_iframe(driver)

max_tentativas = 3
for tentativa in range(max_tentativas):
    try:
        filtro = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
        )
        driver.execute_script("arguments[0].click();", filtro)
        filtro.click()
        break  # Se funcionou, sai do loop
    except StaleElementReferenceException:
        if tentativa < max_tentativas - 1:
            print(f"Elemento stale, tentando novamente... ({tentativa + 1}/{max_tentativas})")
            time.sleep(2)
            entrar_no_iframe(driver)  # Re-entra no iframe
        else:
            raise  # Se esgotou tentativas, levanta o erro
time.sleep(5)

# Limpa todas as seleções do filtro
filtro = WebDriverWait(driver, 30).until(
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
wait = WebDriverWait(driver, 30)
indicador = wait.until(
    #EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']"))
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
)
indicador.click()



# DOWNLOAD
######################################### RATEIO E TRANSACAO
# FUNCOES
# Função para realizar o download dos arquivos
def baixar_arquivos(tipo):
    """
    Função para baixar arquivos do Power BI.
    :param tipo: Tipo de arquivo a ser baixado ('Ordem Ressarcimento' ou 'Ordem Rateio').
    """
    # total de linha - só para contar
    linhas = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located((
            By.XPATH,
            "//div[@role='gridcell' and @column-index='3' and @aria-colindex='5']"
        ))
    )
    
    linhas_tabela = len(linhas)
    print(f"Total de linhas encontradas: {linhas_tabela}")

    
    i = 1
    empresas_processadas = 0
    
    while empresas_processadas < 4:
        try:
            print(f"\n=== Tentando processar índice {i} para {tipo} ===")
            
            consorcio = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, f"(//div[@role='gridcell' and @column-index='2' and @aria-colindex='4'])[{i}]"))
            )
            texto_consorcio = consorcio.text
            print(f" Consórcio encontrado: {texto_consorcio}")
            time.sleep(3)
            
            data_linha = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, f"(//div[@role='gridcell' and @column-index='0' and @aria-colindex='2'])[{i}]"))
            )
            data_linha = data_linha.text.replace('/', '-')
            print(f" Data encontrada: {data_linha}")
            time.sleep(3)
            
            linha = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, f"(//div[@role='gridcell' and @column-index='3' and @aria-colindex='5'])[{i}]"))
            )
            
            driver.execute_script("arguments[0].scrollIntoView(true);", linha)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", linha)
            print(f" Clicou na linha {i} da página principal")
            time.sleep(7)
            

            existentes = {f for f in Path(pasta).iterdir() if f.is_file()}
            print(f" Arquivos existentes capturados: {len(existentes)} arquivos")

            # Clica no botão de drill-through para ir à página de detalhes
            print(f"Tentando clicar no drill-through de '{tipo}'...")
            botao_drill = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, f"//*[@aria-label='Drill-through . Clique aqui para executar uma consulta drill-through em {tipo} Drill']"))
            )
            botao_drill.click()
            print(f" Clicou no drill-through de {tipo}")
            time.sleep(10)
            
            # AGORA ESTAMOS NA PÁGINA DE DRILL-THROUGH - NÃO CLICA EM NENHUMA LINHA
            # Move o mouse para revelar "Mais opções" DIRETAMENTE
            print("Tentando revelar 'Mais opções' na página de drill-through...")
            elemento_para_revelar = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "(//div[@role='columnheader' and @column-index='0' and @aria-colindex='2'])[1]"))
            )
            actions = ActionChains(driver)
            actions.move_to_element(elemento_para_revelar).perform()
            print(" Mouse movido para revelar menu")
            time.sleep(5)
            
            # Clica em "Mais opções"
            print("Tentando clicar em 'Mais opções'...")
            mais_opcoes = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@aria-label,'Mais opções')]"))
            )
            
            driver.execute_script("arguments[0].click();", mais_opcoes)
            print(" Clicou em 'Mais opções' via JS")
            time.sleep(3)

            print(" Clicou em 'Mais opções'")
            time.sleep(5)
            
            # Clica em "Exportar dados"
            print("Tentando clicar em 'Exportar dados'...")
            exportar_dados = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']"))
            )
            exportar_dados.click()
            print(" Clicou em 'Exportar dados'")
            time.sleep(5)
            
            # Clica no botão "Exportar" final
            print("Tentando clicar no botão 'Exportar' final...")
            clicar_exportar(driver, tentativas=5)
            time.sleep(5)
            print(" Clicou no botão 'Exportar' final - Download deve começar agora")
            time.sleep(5)
            
            # Voltar à página anterior (página principal)
            print("Voltando à página principal...")
            botao_voltar = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Voltar . Clique aqui para voltar à página anterior neste relatório']"))
            )
            botao_voltar.click()
            print(" Voltou à página principal")
            time.sleep(7)
            
            # Espera o download terminar
            print("Aguardando download do arquivo...")
            download_timeout = 520
            poll_interval = 1.0
            fim = time.time() + download_timeout
            arquivo_mais_recente = None
            tentativas = 0

            while time.time() < fim:
                tentativas += 1
                if tentativas % 10 == 0:  # Print a cada 10 segundos
                    print(f"  Aguardando... ({tentativas} segundos)")
                
                atuais = {f for f in Path(pasta).iterdir() if f.is_file()}
                novos = atuais - existentes
                
                if not novos:
                    time.sleep(poll_interval)
                    continue

                # ignora temporários
                temporarios = [f for f in novos if f.suffix in {".crdownload", ".part"}]
                if temporarios:
                    if tentativas % 5 == 0:
                        print(f"  Arquivo temporário detectado: {temporarios[0].name}")
                    time.sleep(poll_interval)
                    continue

                candidatos = [f for f in novos if f.suffix.lower() == ".xlsx"]
                if not candidatos:
                    time.sleep(poll_interval)
                    continue

                print(f"  Arquivo .xlsx detectado: {candidatos[0].name}")

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
                                print(f" Download estabilizado: {candidato.name} ({atual_size} bytes)")
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

            if not arquivo_mais_recente:
                print(f" Download não concluído para linha {i} de {tipo}. Timeout atingido.")
                i += 1
                continue

            # Garante que a pasta de destino exista
            os.makedirs(diretorio_destino, exist_ok=True)

            
            # Renomeia o arquivo no padrão
            nome_arquivo_novo = f"{data_linha} - {texto_consorcio} - {tipo}.xlsx"
            
            # Move para a pasta final
            caminho_novo = os.path.join(diretorio_destino, nome_arquivo_novo)
            shutil.move(str(arquivo_mais_recente), caminho_novo)
            print(f" Arquivo movido para: {caminho_novo}")
            
            empresas_processadas += 1
            print(f" Empresa {empresas_processadas}/4 processada com sucesso!")
            i += 1
            time.sleep(5)
        
        except Exception as e:
            print(f" Erro na linha {i} ({tipo}): {e}")
            import traceback
            traceback.print_exc()
            i += 1
            continue


baixar_arquivos("Ordem Ressarcimento")
print("==============================================")
print("FINALIZADO: Download da ORDEM RESSARCIMENTO")
print("Iniciando agora o download da ORDEM RATEIO")
print("==============================================")
baixar_arquivos("Ordem Rateio")


# Fecha o navegador
driver.quit()


#CONSOLIDANDO OS ARQUIVOS BAIXADOS EM UM ARQUIVO E SUBINDO NA TABELA DO BIGQUERY
#Ordem Ressarcimento
termo = 'Ordem Ressarcimento'
        
# Lista para armazenar os DataFrames
lista_dataframes = []


# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    if arquivo.endswith('.xlsx') and termo in arquivo and str(data_ordem) in arquivo: 
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        df = pd.read_excel(
            caminho_completo,
            dtype={
                "Nr Linha": str,
                "Ordem Ressarcimento": str,
            }
        )
        
        if 'Internorte' in arquivo:
            df['Consorcio'] = 'Internorte'
        elif 'Santa Cruz' in arquivo:
            df['Consorcio'] = 'Santa Cruz'
        elif 'Intersul' in arquivo:
            df['Consorcio'] = 'Intersul'
        elif 'Transcarioca' in arquivo:
            df['Consorcio'] = 'Transcarioca'

        valor_ultima_linha = df.iloc[-1, 0]

        match_consorcio = re.search(r'Filtros aplicados:\nid_ordem_pagamento_consorcio é (\d+)', valor_ultima_linha)
        df['id_ordem_pagamento_consorcio'] = match_consorcio.group(1) if match_consorcio else None

        match_pagamento = re.search(r'id_pagamento é (\d+)', valor_ultima_linha)
        df['id_pagamento'] = match_pagamento.group(1) if match_pagamento else None

        df = df.iloc[:-3]
        lista_dataframes.append(df)



# Concatena todos os DataFrames em um único
ordem_ressarcimento = pd.concat(lista_dataframes, ignore_index=True)

# Organiza a ordem das colunas do DatFrame
ordem_ressarcimento = ordem_ressarcimento[['Data Ordem Ressarcimento','Consorcio','Ordem Ressarcimento','id_ordem_pagamento_consorcio','id_pagamento','Status Ordem','Operadora','Nr Linha','Linha',
                                           'Valor Bruto','Valor Taxa','Valor Líquido','Valor Débito','Qtd Débito','Valor Integração','Qtd Integração','Valor Rateiro Crédito','Qtd Rateio Crédito',
                                           'Valor Rateio Débito','Qtd Rateio Débito','Valor Venda a Bordo','Qtd Venda a Bordo','Valor Gratuidade','Qtd Gratuidade']]

# Renomea colunas
ordem_ressarcimento.columns = ['data_ordem','consorcio','ordem_ressarcimento','id_ordem_pagamento_consorcio','id_pagamento','status_ordem','operadora','servico','linha',
                              'valor_bruto','valor_taxa','valor_liquido','valor_debito','qtd_debito','valor_integracao','qtd_integracao','valor_rateio_credito','qtd_rateio_credito',
                              'valor_rateio_debito','qtd_rateio_debito','valor_venda_a_bordo','qtd_venda_a_bordo','valor_gratuidade','qtd_gratuidade']


# Padronização
ordem_ressarcimento['data_ordem'] = pd.to_datetime(ordem_ressarcimento['data_ordem'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d')

# Garantindo que valores inteiros fiquem sem ".0"
colunas_inteiras = ['ordem_ressarcimento','id_ordem_pagamento_consorcio','id_pagamento','qtd_debito','qtd_integracao','qtd_rateio_credito','qtd_rateio_debito','qtd_venda_a_bordo','qtd_gratuidade'] 

for coluna in colunas_inteiras:
    ordem_ressarcimento[coluna] = ordem_ressarcimento[coluna].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)

# Salvar no CSV sem converter inteiros para float
ordem_ressarcimento.to_csv(f"C:/Users/{USUARIO}/Desktop/Bases_Ressarcimento_Linha/{data_ordem} Ordem Ressarcimento.csv", 
              index=False, sep=";", encoding="utf-8-sig", decimal=".")


# Configurações principais
project_id = os.getenv('BQ_PROJECT')
dataset_id = os.getenv('BQ_DATASET')
table_id = os.getenv('BQ_TABLE_RESUMO')
source_file = f"C:/Users/{USUARIO}/Desktop/Bases_Ressarcimento_Linha/{data_ordem} Ordem Ressarcimento.csv"

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_ordem", "DATE"),
    bigquery.SchemaField("consorcio", "STRING"),
    bigquery.SchemaField("ordem_ressarcimento", "STRING"),
    bigquery.SchemaField("id_ordem_pagamento_consorcio", "STRING"),
    bigquery.SchemaField("id_pagamento", "STRING"),
    bigquery.SchemaField("status_ordem", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("servico", "STRING"),
    bigquery.SchemaField("linha", "STRING"),
    bigquery.SchemaField("valor_bruto", "FLOAT64"),
    bigquery.SchemaField("valor_taxa", "FLOAT64"),
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
    bigquery.SchemaField("qtd_gratuidade", "INTEGER")  
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
client.close()



#Ordem Rateio
termo = 'Ordem Rateio'
        
# Lista para armazenar os DataFrames
lista_dataframes = []

# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    if arquivo.endswith('.xlsx') and termo in arquivo and str(data_ordem) in arquivo: 
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        df = pd.read_excel(
            caminho_completo,
            dtype={
                "Linha": str,
                "id_ordem_rateio": str,
            }
        )
        
        if 'Internorte' in arquivo:
            df['Consorcio'] = 'Internorte'
        elif 'Santa Cruz' in arquivo:
            df['Consorcio'] = 'Santa Cruz'
        elif 'Intersul' in arquivo:
            df['Consorcio'] = 'Intersul'
        elif 'Transcarioca' in arquivo:
            df['Consorcio'] = 'Transcarioca'

        valor_ultima_linha = df.iloc[-1, 0]

        match_consorcio = re.search(r'id_ordem_pagamento_consorcio é (\d+)', valor_ultima_linha)
        df['id_ordem_pagamento_consorcio'] = match_consorcio.group(1) if match_consorcio else None


        df = df.iloc[:-3]
        lista_dataframes.append(df)

# Concatena todos os DataFrames em um único
ordem_rateio = pd.concat(lista_dataframes, ignore_index=True)

# Organiza a ordem das colunas do DatFrame
ordem_rateio = ordem_rateio[['Data Ordem Rateio','Consorcio','id_ordem_pagamento_consorcio','id_ordem_rateio','Operadora','Linha',
                             'Qtd débito total','Valor débito total','Qtd crédito total','Valor crédito total']]

# Renomea colunas
ordem_rateio.columns = ['data_ordem','consorcio','id_ordem_pagamento_consorcio','id_ordem_rateio','operadora','linha',
                        'qtd_debito_total','valor_debito_total','qtd_credito_total','valor_credito_total']


# Padronização
ordem_rateio['data_ordem'] = pd.to_datetime(ordem_rateio['data_ordem'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d')

# Garantindo que valores inteiros fiquem sem ".0"
colunas_inteiras = ['id_ordem_pagamento_consorcio','id_ordem_rateio','qtd_debito_total','qtd_credito_total'] 

for coluna in colunas_inteiras:
    ordem_rateio[coluna] = ordem_rateio[coluna].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)

# Salvar no CSV sem converter inteiros para float
ordem_rateio.to_csv(f"C:/Users/{USUARIO}/Desktop/Bases_Ressarcimento_Linha/{data_ordem} Ordem Rateio.csv", 
              index=False, sep=";", encoding="utf-8-sig", decimal=".")


# Configurações principais
project_id = os.getenv('BQ_PROJECT')
dataset_id = os.getenv('BQ_DATASET')
table_id = os.getenv('BQ_TABLE_RATEIO')
source_file = f"C:/Users/{USUARIO}/Desktop/Bases_Ressarcimento_Linha/{data_ordem} Ordem Rateio.csv"

# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_ordem", "DATE"),
    bigquery.SchemaField("consorcio", "STRING"),
    bigquery.SchemaField("id_ordem_pagamento_consorcio", "STRING"),
    bigquery.SchemaField("id_ordem_rateio", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("linha", "STRING"),
    bigquery.SchemaField("qtd_debito_total", "INTEGER"),
    bigquery.SchemaField("valor_debito_total", "FLOAT64"),
    bigquery.SchemaField("qtd_credito_total", "INTEGER"),
    bigquery.SchemaField("valor_credito_total", "FLOAT64")
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

client.close()

print("SCRIPT FINALIZADO COM SUCESSO")