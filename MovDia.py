

import os
import glob
import pandas as pd
from selenium import webdriver
import getpass
from datetime import datetime, timedelta
import shutil
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.by import By
import openpyxl
import xlwings as xw
from pandas.tseries.offsets import BDay


url = 'https://www.itaucustodia.com.br/Passivo/'
usuario = 
senha = 
username = getpass.getuser()       
source = r'C://Users\{}//Downloads'.format(username)
path = 'Z:/Outros/CAIXA ON/MOV DIA/'
consulta_movimentos_do_dia = 'https://www.itaucustodia.com.br/Passivo/abreFiltroConsultaMovimentosDiaGestor.do'
cod_gestor = '1406'

#wb = xw.Book("Z:\Outros\CAIXA ON\v13.xlsm")
#ws = wb.sheets['MENU']
#data1 = ws.range('C4').value
data1 = datetime.today()
data = data1.strftime('%Y_%m_%d')

driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get(url)

driver.find_element(By.XPATH,'//*[@id="combo"]/form/table/tbody/tr[1]/td[2]/input').send_keys(usuario)

driver.find_element(By.XPATH,'//*[@id="combo"]/form/table/tbody/tr[2]/td[2]/input').send_keys(senha)

driver.find_element(By.XPATH,'//*[@id="combo"]/form/table/tbody/tr[2]/td[3]/a[1]/img').click()

driver.get(consulta_movimentos_do_dia)

driver.find_element(By.NAME,'idGestor').send_keys(cod_gestor)

driver.find_element(By.XPATH,'//*[@id="conteudo"]/div/table/tbody/tr[10]/td/a/img').click()

driver.find_element(By.XPATH,'//*[@id="conteudo"]/div/table[1]/tbody/tr[6]/td[3]/a/img').click()

time.sleep(8)

driver.close()

print('file nos downloads')

files = [(file, os.path.getmtime(os.path.join(source, file))) for file in os.listdir(source)]
files.sort(key=lambda x: x[1], reverse=True)
most_recent_file = files[0][0]
print(most_recent_file)



input_path = os.path.join(source, most_recent_file)
df = pd.read_html(input_path, header = 1, thousands = ".", decimal = ",")[0]


path_file = os.path.join(path, most_recent_file)

shutil.move(input_path, path)

df.rename(columns ={'Unnamed: 0' : 'Column1'}, inplace = True)


df.to_excel(path_file.replace('.xls', '.xlsx'), sheet_name = 'Sheet1', index = False)
os.remove(path_file)


files = [(file, os.path.getmtime(os.path.join(path, file))) for file in os.listdir(path)]
files.sort(key=lambda x: x[1], reverse=True)
most_recent_file2 = files[0][0]
print(most_recent_file2)

filedir = os.path.join(path, most_recent_file2)

output_path = os.path.join(path, "MOVIMENTACOES - {}.xlsx").format(data)

os.rename(filedir, output_path)