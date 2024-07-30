import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime
from selenium.webdriver.common.keys import Keys
import openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json

meses = {
    'jan': 'jan',
    'feb': 'fev',
    'mar': 'mar',
    'apr': 'abr',
    'may': 'mai',
    'jun': 'jun',
    'jul': 'jul',
    'aug': 'ago',
    'sep': 'set',
    'oct': 'out',
    'nov': 'nov',
    'dec': 'dez'
}

# Carregar configurações do arquivo JSON
with open("config.json") as config_file:
    config = json.load(config_file)

jira_url = config["jira_url"]
username = config["login"]["username"]
password = config["login"]["password"]
planilha_path = config["planilha_path"]
url_chamado_prefix = config["url_chamado_prefix"]

def login_jira(driver):
    driver.get(jira_url)
    time.sleep(4)
    driver.find_element("xpath", '//*[@id="content"]/div/div/section/div/div/p[2]/a').click()
    time.sleep(4)
    driver.find_element(By.ID, "login-form-username").send_keys(username)
    driver.find_element(By.ID, "login-form-password").send_keys(password)
    driver.find_element(By.ID, "login-form-submit").click()

# Função para converter Timestamp para string no formato desejado
def formatar_data(data):
    data_formatada = data.strftime('%d/%b/%Y').lower()  # %b para abreviação do mês
    for mes_ingles, mes_portugues in meses.items():
        data_formatada = data_formatada.replace(mes_ingles, mes_portugues)
    return data_formatada

def transformar_tempo(hora):
    horas, minutos = map(int, hora.split(':'))
    return f"{horas * 60 + minutos}m"

def valida_campos(data, total, chamado, desc, apontado):
    if data is None or total == '0m' or '#' in total or desc is None or apontado == 'Sim' or apontado is None or chamado is None:
        return False
    else:
        return True

def apontar_horas(driver, data, total, chamado, desc):
    driver.get(url_chamado_prefix + chamado)
    time.sleep(4)

    body = driver.find_element(By.TAG_NAME, 'body')
    body.send_keys('w')
    time.sleep(3)
    
    elements = driver.find_elements(By.XPATH, '//*[@id="comment"]')
    # Filtrar o elemento que está visível
    element = None
    for el in elements:
        if el.is_displayed():
            element = el
            break
    
    if element is None:
        raise Exception("Nenhum elemento visível encontrado com o id 'comment'")

    driver.execute_script("arguments[0].scrollIntoView();", element)
    element.send_keys(desc)    
    driver.find_element(By.ID, "started").send_keys(Keys.CONTROL + "a")
    driver.find_element(By.ID, "started").send_keys(Keys.BACKSPACE)
    driver.find_element(By.ID, "started").send_keys(data)
    driver.find_element(By.XPATH, '//*[@id="worklogForm"]/span/div/div[2]/div/div').click()
    driver.find_element(By.ID, "timeSpentSeconds").send_keys(total)
    driver.find_element(By.XPATH, '//*[@id="worklogForm"]/footer/div[2]/button[1]').click()
    time.sleep(2)
    return True

# Configurações do Selenium
chrome_options = Options()
#chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
driver = webdriver.Chrome(options=chrome_options)

try:
    login_jira(driver)
    time.sleep(3)

    df = pd.read_excel(planilha_path)
    wb = openpyxl.load_workbook(planilha_path)
    ws = wb.active

    # Iterar sobre a planilha e apontar horas no Jira
    for index, row in df.iterrows():
        data = str(formatar_data(row['Data']))
        total = transformar_tempo(row['Total'].strftime('%H:%M'))
        chamado = row['Chamado']
        desc = row['Descrição']
        apontado = row['Apontado']

        if valida_campos(data, total, chamado, desc, apontado):
            apontar_horas(driver, data, total, chamado, desc)
            
            df.at[index, 'Apontado'] = 'Sim'
            ws.cell(row=index + 2, column=df.columns.get_loc('Apontado') + 1, value='Sim')
            print(f"Apontamento atualizado para 'Sim' para a linha {index + 2}")

    # Salvando as alterações no arquivo Excel
    wb.save(planilha_path)
    print("Alterações salvas no arquivo Excel.")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

finally:
    driver.quit()

#Made with <3 by @heloisacst.
