from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.action_chains import ActionChains
import os
import pandas as pd
from pathlib import Path
from selenium.webdriver.chrome.options import Options
from google.cloud import bigquery
import shutil
from datetime import datetime, timedelta
from dotenv import load_dotenv
import sys
sys.stdout.reconfigure(encoding='utf-8')

print("Olá, ação, índice, coração")

BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / "config" / ".env"

load_dotenv(dotenv_path=ENV_PATH, override=True)

DATA_DIA = os.getenv("DATA_DIA")
DATA_MES = os.getenv("DATA_MES")
DATA_ANO = os.getenv("DATA_ANO")
POWERBI_EMAIL = os.getenv("POWERBI_EMAIL")
POWERBI_SENHA = os.getenv("POWERBI_PASSWORD")
USUARIO = os.getenv("USUARIO")
URL = os.getenv("URL")
BQ_KEY_PATH= os.getenv("BQ_KEY_PATH")


# Caminho dos arquivos brutos baixados
pasta = f"C:/Users/{USUARIO}/Downloads"
caminho_pasta = Path(pasta)

# Caminho dos arquivos baixados renomeados
diretorio_destino = f"C:/Users/{USUARIO}/Desktop/Bases_Gratuidades"
pasta_final = Path(diretorio_destino)

# Informando o caminho do programa do Chrome (chromedriver)
driver_path = f"C:\\Users\\{USUARIO}\\Downloads\\chromedriver-win64\\chromedriver.exe"
service = Service(driver_path)

# Caminho para sua chave de serviço JSON
key_path = BQ_KEY_PATH

# Configura o cliente do BigQuery
client = bigquery.Client.from_service_account_json(key_path)


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
driver.get(URL)

# Login no Power BI
# Preenche email
email_input = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.ID, "Email"))
)

email_input.send_keys(POWERBI_EMAIL)
#email_input.send_keys(Keys.RETURN)
time.sleep(3)

# Preenche senha
senha_input = driver.find_element(By.ID, 'Password')
senha_input.send_keys(POWERBI_SENHA)  # Senha
senha_input.send_keys(Keys.RETURN)

#driver.switch_to.default_content()

# Esperar até o iframe com o src correto estar disponível
iframe = WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'https://app.powerbi.com/reportEmbed?reportId=553bdead-44d6-44b8-a987-52b1a0e10fd5')]"))
)

# Mudar para o iframe usando o src
driver.switch_to.frame(iframe)


# FILTRO DE PERIODO
# Localiza e clica no elemento de filtro
filtro = WebDriverWait(driver, 60).until(
    #EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']"))
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
)
filtro.click()


# Localiza o container onde o tem o ano
item_ano = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((
        By.XPATH,
        f"//div[@class='slicerItemContainer'][.//span[@class='slicerText' and text()='{DATA_ANO}']]"
    ))
)



'''
# Dentro dele, pega o botão de expandir
botao_expand = item_ano.find_element(By.CSS_SELECTOR, "div.expandButton")
botao_expand.click()
'''



# Localiza o bloco de scroll de data
bloco_scroll = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".scroll-content.scroll-scrolly_visible"))
)

# Rola até o elemento
driver.execute_script("arguments[0].scrollTop += 50;", bloco_scroll)


# Localiza o container onde o texto é o mês
item_mes = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((
        By.XPATH,
        f"//div[@class='slicerItemContainer'][.//span[@class='slicerText' and text()='{DATA_MES}']]"
    ))
)

# Dentro desse container, pega o expandButton
botao_expand = item_mes.find_element(By.CSS_SELECTOR, "div.expandButton")
botao_expand.click()



# Localiza o bloco de scroll de data
bloco_scroll = WebDriverWait(driver, 15).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".scroll-content.scroll-scrolly_visible"))
)


# Função para tentar localizar a data dentro do bloco
def encontrar_item(bloco, termo):
    try:
        return bloco.find_element(By.XPATH, f".//span[contains(@class,'slicerText') and text()='{DATA_DIA}']")
    except:
        return None

# Loop para rolar até encontrar a data ou atingir o final
scroll_max = 0
while True:
    item = encontrar_item(bloco_scroll, DATA_DIA)
    if item:
        # Item encontrado, clicar
        item.click()
        break
    else:
        # Rolar para baixo
        driver.execute_script("arguments[0].scrollTop += 50;", bloco_scroll)
        time.sleep(0.5)  # espera o scroll renderizar
        # Verifica se chegou ao final
        novo_scroll = driver.execute_script("return arguments[0].scrollTop;", bloco_scroll)
        if novo_scroll == scroll_max:
            print("Data não encontrada no slicer.")
            break
        scroll_max = novo_scroll
        
    
# Clica no filtro de tipo transação
transacao = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='ds_tipo_transacao']"))
)
transacao.click()

time.sleep(3)


# Digita "grat" para filtrar as gratuidades
acao = ActionChains(driver)
acao.send_keys("grat").perform()

time.sleep(3)


# Faz CTRL + clique em todos os itens
itens = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, "//div[@class='slicerItemContainer' and @role='option']"))
)

acao = ActionChains(driver)
for item in itens:
    acao.key_down(Keys.CONTROL).click(item).key_up(Keys.CONTROL)

acao.perform()



# Localiza e clica no elemento de filtro
filtro = WebDriverWait(driver, 60).until(
    #EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para seguir link']"))
    EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
)
filtro.click()




empresas = ['Alpha', 'Jabour', 'Palmares', 'Tijuca', 'Tres Amigos', 'Caprichosa', 'Braso', 'Ideal', 'Pegaso', 'Recreio', 'Gire', 'Real Auto', 'Matias L', 'SANCETUR',
            'Vila Isabel', 'Barra Ltda', 'Campo Grande', 'Futuro', 'Paranapuan', 'Transurb', 'MENDANHA', 'Normandy', 'Senhora das Gra', 'Senhora de Lourdes', 'Novacap',
            'Pavunense','Redentor', 'Verdun', 'VG', 'Vila Real']


quantidade = len(empresas)

i = 1

for empresa in empresas:

    try:
        # Localiza e clica no elemento de filtro
        # Clicar no filtro principal
        filtro = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
        )
        filtro.click()
        time.sleep(0.5)

        # Abrir filtro de operadora
        filtro_operadora = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='nm_operadora']"))
        )
        filtro_operadora.click()
        time.sleep(0.5)

        # Localizar input de pesquisa, limpar e digitar empresa
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.DELETE).perform()
        time.sleep(0.6)

        # Digitar nova empresa
        acao = ActionChains(driver)
        acao.send_keys(empresa).perform()

        time.sleep(3)  # espera a lista ser filtrada


        # Esperar e clicar na primeira empresa
        try:
            operadora = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//div[@class='slicerItemContainer' and contains(., '{empresa}')]")
                )
            )
            operadora.click()
            time.sleep(3)
        except:
            print(f"Empresa {empresa} não localizada no filtro, pulando...")
            # Fecha o filtro antes de continuar
            try:
                filtro.click()
            except:
                pass
            continue
        
        # Fecha o filtro
        filtro = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Indicador . Clique aqui para Seguir']"))
        )
        filtro.click()
        time.sleep(1)


        # Esperar o eixo X aparecer e clicar na coluna do gráfico
        elementos = driver.find_elements(
            By.XPATH,
            f"//*[name()='text' and contains(normalize-space(.), '{empresa}')]"
        )
        
        if elementos:
            elemento = elementos[0]
        
            # Centraliza o elemento na tela
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", elemento
            )
            time.sleep(1)
        
            clicou = False
        
            for tentativa in range(3):
                try:
                    elemento.click()
                    time.sleep(0.8)
                    clicou = True
                    break
                except WebDriverException:
                    try:
                        # fallback: clique via JavaScript
                        driver.execute_script("arguments[0].click();", elemento)
                        time.sleep(0.8)
                        clicou = True
                        break
                    except:
                        time.sleep(1)
        
            if not clicou:
                print(f"Não foi possível clicar na empresa {empresa} após várias tentativas.")
        
        else:
            print(f"Empresa {empresa} não encontrada no gráfico.")
            continue



        # Captura estado atual da pasta antes do download
        existentes = {f for f in Path(pasta).iterdir() if f.is_file()}


        # Clica na opção Detalhe
        botao_detalhe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, "//div[@role='button' and contains(@aria-label, 'Detalhe Transação')]"
            ))
        )

        botao_detalhe.click()


        # Move o mouse para revelar "Mais opções"
        elemento_para_revelar = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "(//div[@title='Detalhe de Transações'])"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(elemento_para_revelar).perform()
        time.sleep(7)
        
        # Clica em "Mais opções" e "Exportar"
        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@class='vcMenuBtn' and @aria-label='Mais opções']"))
        ).click()
        time.sleep(7)
        
        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Exportar dados']"))
        ).click()
        time.sleep(7)
        
        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Exportar']"))
        ).click()


        # Espera o download terminar (substitui sleep fixo)
        download_timeout = 520
        poll_interval = 1.0
        fim = time.time() + download_timeout
        arquivo_mais_recente = None

        while time.time() < fim:
            atuais = {f for f in Path(pasta).iterdir() if f.is_file()}
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

        if not arquivo_mais_recente:
            print(f"Download não concluído para empresa {i} - {empresa}.")
            break
        time.sleep(2)

        # garante que o arquivo existe mesmo
        if not os.path.exists(arquivo_mais_recente):
            print("Arquivo não encontrado:", arquivo_mais_recente)
        else:
            novo_nome = f"Detalhe de Transações - {empresa}.xlsx"
            destino = os.path.join(diretorio_destino, novo_nome)
        
            # renomeia e move
            shutil.move(arquivo_mais_recente, destino)
            print("Arquivo movido para:", destino)

        # Garante que a pasta de destino exista
        # os.makedirs(diretorio_destino, exist_ok=True)
        
        # Renomeia o arquivo no padrão
        # nome_arquivo_novo = f"Detalhe de Transações - {empresa}.xlsx"
        
        # Move para a pasta final
        # caminho_novo = os.path.join(diretorio_destino, nome_arquivo_novo)
        # shutil.move(str(arquivo_mais_recente), caminho_novo)
        
        # Voltar à página anterior
        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@aria-label='Voltar . Clique aqui para voltar à página anterior neste relatório']"))
        ).click()



    except Exception as e:
        print(f"Fim do loop ou erro ao processar empresa: {e}")


# Fecha o navegador
driver.close()  



# Consolidando arquivos
termo = 'Detalhe de Transações'

# Lista para armazenar os DataFrames
lista_dataframes = []

# Itera pelos arquivos na pasta
for arquivo in os.listdir(diretorio_destino):
    if arquivo.endswith('.xlsx') and termo in arquivo in arquivo: 
        caminho_completo = os.path.join(diretorio_destino, arquivo)
        df = pd.read_excel(
            caminho_completo,
            dtype={
                "Nr Linha": str,
                "Serviço": str,
                "Id Cliente": str,
                "Prefixo Veículo": str
            }
        )
        
        df = df.iloc[:-3]                                                                           #exclui as últimas 3 linhas do df
        lista_dataframes.append(df)  # Adicionar o DataFrame à lista

# Concatena todos os DataFrames em um único
gratuidades = pd.concat(lista_dataframes, ignore_index=True)

gratuidades = gratuidades.drop_duplicates()
total = gratuidades['Qtde Transação'].sum()

# Renomea colunas
gratuidades.columns = ['data_transacao','data_processamento','id_cliente','operadora','nr_linha','linha','validador','prefixo_veiculo','servico','tipo_produto','produto','tipo_midia','tipo_usuario','tipo_transacao','qtd_transacao','valor_transacao','id_transacao']

# Padronização
gratuidades['data_processamento'] = pd.to_datetime(gratuidades['data_processamento'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d')
gratuidades['data_transacao'] = pd.to_datetime(gratuidades['data_transacao'], format='%d-%m-%Y').dt.strftime('%Y-%m-%d %hh:%mm:%ss')

# Garantindo que os valores monetários tenham casas decimais
colunas_float = ['valor_transacao']

for coluna in colunas_float:
    gratuidades[coluna] = gratuidades[coluna].astype(float)
    
# Garantindo que valores inteiros fiquem sem ".0"
colunas_inteiras = ['qtd_transacao']

for coluna in colunas_inteiras:
    gratuidades[coluna] = gratuidades[coluna].apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x)
    
# Garantindo que valores string sem ".0"
for coluna in ['id_cliente','prefixo_veiculo']:
    gratuidades[coluna] = gratuidades[coluna].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0', '').isdigit() else str(x)
    )

# Salvar no CSV sem converter inteiros para float
gratuidades.to_csv(f"C:/Users/{USUARIO}/Desktop/Bases_Gratuidades/Gratuidades.csv", 
                 index=False, sep=";", encoding="utf-8-sig", decimal=".")


# Configurações principais
project_id = os.getenv("BQ_PROJECT")
dataset_id = os.getenv("BQ_DATASET")
table_id = os.getenv("BQ_TABLE")
source_file = f"C:/Users/{USUARIO}/Desktop/Bases_Gratuidades/Gratuidades.csv"


# Define a tabela de destino no formato completo
table_ref = f"{project_id}.{dataset_id}.{table_id}"

# Configurações do job de carregamento
schema = [
    bigquery.SchemaField("data_transacao", "DATETIME"),
    bigquery.SchemaField("data_processamento", "DATETIME"),
    bigquery.SchemaField("id_cliente", "STRING"),
    bigquery.SchemaField("operadora", "STRING"),
    bigquery.SchemaField("nr_linha", "STRING"),
    bigquery.SchemaField("linha", "STRING"),
    bigquery.SchemaField("validador", "STRING"),
    bigquery.SchemaField("prefixo_veiculo", "STRING"),
    bigquery.SchemaField("servico", "STRING"),
    bigquery.SchemaField("tipo_produto", "STRING"),
    bigquery.SchemaField("produto", "STRING"),
    bigquery.SchemaField("tipo_midia", "STRING"),
    bigquery.SchemaField("tipo_usuario", "STRING"),
    bigquery.SchemaField("tipo_transacao", "STRING"),
    bigquery.SchemaField("qtd_transacao", "INTEGER"),
    bigquery.SchemaField("valor_transacao", "FLOAT64"),
    bigquery.SchemaField("id_transacao", "STRING")
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



# Percorre todos os arquivos e remove os que forem .xlsx ou .csv
for arquivo in Path(pasta_final).iterdir():
    if arquivo.is_file() and arquivo.suffix.lower() in {'.xlsx', '.csv'}:
        try:
            os.remove(arquivo)
            print(f"Arquivo removido: {arquivo.name}")
        except Exception as e:
            print(f"Erro ao remover {arquivo.name}: {e}")

print("Limpeza concluída.")