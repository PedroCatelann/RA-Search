from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
import time
import csv
import os
import pandas as pd
import sys
# Definir o nome do arquivo CSV a ser deletado
nome_arquivo_csv = "tabela_excel.xlsx"
print(os.path.dirname(sys.executable))
# Verificar se o arquivo existe antes de tentar deletá-lo
if os.path.exists(nome_arquivo_csv):
    # Deletar o arquivo
    os.remove(nome_arquivo_csv)
def document_initialised(driver):
    return driver.execute_script("return initialised")

driver_path =  'C:\\Users\\autbank\\Downloads\\chromedriver.exe'

service = Service(executable_path=driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# URL da página local hospedada no Tomcat
url = 'http://140.1.254.190:9090/regatendimento/faces/raindex.jsp'

driver.get(url)

window_after = driver.window_handles[1]

driver.switch_to.window(window_after)

name = driver.find_element(By.ID,"frmcaindex:NOME_USUARIO")
name.send_keys("PHFCATELAN")
password = driver.find_element(By.ID,"frmcaindex:SENHA_ATUAL")
password.send_keys("28phfktlan02")

button = driver.find_element(By.ID,'frmcaindex:btnOk')
button.click()

frame = driver.find_element(By.NAME,'main')
driver.switch_to.frame(frame)

button2 = driver.find_element(By.NAME,'page:frmSideMenu:_id44')
button2.click()

button3 = driver.find_element(By.ID,'frmTopLayoutDoubleMenu:topmenu:tm_3')
button3.click()

button4 = driver.find_element(By.ID,'frmTopLayoutDoubleMenu:page:t_3_1')
button4.click()


frame = driver.find_element(By.NAME,'conteudo')
driver.switch_to.frame(frame)


select_element = driver.find_element(By.ID,'page:frmra_consra_analista_r:codigosistema')
select = Select(select_element)
select.select_by_value('CC')

select_element2 = driver.find_element(By.ID,'page:frmra_consra_analista_r:situacaora')
select2 = Select(select_element2)
select2.select_by_value('99')

date = driver.find_element(By.ID,"page:frmra_consra_analista_r:datade")
date.clear()

button2 = driver.find_element(By.NAME,'page:frmra_consra_analista_r:_id79')
button2.click()

time.sleep(10)
headers = ['CLIENTE','RA','AMBIENTE','RESPONSAVEL','DATA ABERTURA','SITUAÇÃO']
table_csv = []

qntd = driver.find_element(By.CLASS_NAME,'pagerDeluxe_text')
print(qntd.text.split(" ")[3])

for i in range(int(qntd.text.split(" ")[3])):

    tabela = driver.find_element(By.ID,'page:frmra_consra_analista_r:ssBTORAGrid0')


    linhas = tabela.find_elements(By.TAG_NAME,'tr')

    valores_da_tabela = []

    for linha in linhas[5:]:
        colunas = linha.find_elements(By.TAG_NAME,'td')
        
        valores = [coluna.text for coluna in colunas]
        valores_da_tabela.append(valores)

    line_table_csv = []

    for j,valores in enumerate(valores_da_tabela):
        print(valores_da_tabela[j][0])
        print(valores_da_tabela[j][1])
        print(valores_da_tabela[j][3])
        print(valores_da_tabela[j][4])
        print(valores_da_tabela[j][6])
        print(valores_da_tabela[j][8])
        print("-"*100)
        c = [valores_da_tabela[j][0],valores_da_tabela[j][1],valores_da_tabela[j][3],valores_da_tabela[j][4],valores_da_tabela[j][6],valores_da_tabela[j][8]]
        table_csv.append(c)
    button3 = driver.find_element(By.ID,'page:frmra_consra_analista_r:ssBTORAGrid0:ssBTORAGrid0_scroll_3next')
    button3.click()
    

table_csv.insert(0,headers)
for a in table_csv:
    print(a)

df = pd.DataFrame(table_csv[1:], columns=table_csv[0])
nome_planilha = "Dados_RA"

nome_arquivo_excel = "tabela_excel.xlsx"

df.to_excel(nome_arquivo_excel, sheet_name=nome_planilha, index=False)

driver.close()

