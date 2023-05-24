# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 14:21:35 2022

@author: akligerman
"""

from datetime import datetime, timedelta
import getpass
import pandas as pd
import xlwings as xw
import tkinter
import time
import requests 
import io
import sys



username = getpass.getuser()


lista_fundos = ['INTMODAL63542','INTMODAL63957','INTNOVUS63553', 'INTNOVUS63566','INTNOVUS64893', 'INTNOVUS65376']

liquidez_path = r"Z:\Outros\LIQUIDEZ\Teste_Liquidez.xlsm"

#------------------------------PATHs---------------------------------------#
#Nessa etapa definimos os caminhos dos arquivos e as datas base pro código

workbook = pd.read_excel(liquidez_path)

username = getpass.getuser()

#delta = int(input('DIGITE O DELTA DE DIAS: ' ))

data = sys.argv[1]
delta = float(sys.argv[2])


#ALFA = float(input("QUAL ALFA VAMOS USAR (20 = 20%):  "))    
#ALFA = ALFA/100
ALFA = float(sys.argv[3])/100

print(data,delta, ALFA)


date_obj = datetime.strptime(data, '%d/%m/%Y')

menosum = datetime.today() - timedelta(days=1)
menostres = datetime.today() - timedelta(days=3)
menosdelta = date_obj - timedelta(days=delta)





def nome_data(menostres,menosum):
    x = datetime.today().weekday()
    if x == 0:
        data_hoje = menostres
    else:
        data_hoje = menosum 
    return data_hoje

data = date_obj.strftime('%Y_%m_%d')
ano = date_obj.strftime('%Y')
mes = date_obj.strftime('%m')
data_sheet_bbg_final = date_obj.strftime('%Y%m%d')
data_sheet_bbg_inicio = menosdelta.strftime('%Y%m%d')








path_destino_LIQUIDEZ = "Z:\Outros\LIQUIDEZ\LIQUIDEZ.txt"
path_carteira = "Z:\_ARQUIVOS INPUTS\INTRAG\CARTEIRAS\CARTEIRAS - {}.xls".format(data)
path_VOLUMES = "P:\DADOSBBG\{}.xlsx".format(username)
path_bbg = "P:\DADOSBBG\BloombergV4.2.xlsm"




#---------------------------DATAFRAMES BASES -----------------------------------------------------------#
#Nessa etapa definimos os DataFrames gerais das carteiras divididos por Tipo de Ativo e o Patrimonio do Fundo

#Acoes
df_car_acoes = pd.read_excel(path_carteira, sheet_name = "Acoes", index_col= False)

#Futuro
df_car_futuros = pd.read_excel(path_carteira, sheet_name = 'Futuros', index_col= False)

#RF
df_car_RF = pd.read_excel(path_carteira, sheet_name = 'Renda_Fixa', index_col= False)

#Volumes do BBG
df_vol = pd.read_excel(path_VOLUMES, sheet_name = "Sheet1", index_col= False)

#Patrimonio dos Fundos
df_patrimonio = pd.read_excel(path_carteira, sheet_name = "Patrimonio_Cotas", index_col= False)


#---------------------TOTAL DE AÇOES POR CARTEIRA-----------------------------------------------#
#Aqui calculamos o total financeiro investido por fundo

#Acoes
df_tot_acoes =  df_car_acoes.groupby(['Carteira/Fundo']).agg(TOT_EQ_FUNDO = ('Valor Mercado',sum))

#Futuro
df_tot_futuros =  df_car_futuros.groupby(['Carteira/Fundo']).agg(TOT_FUT_FUNDO = ('Valor de Mercado',sum))

#RF
df_tot_RF =  df_car_RF.groupby(['Carteira/Fundo']).agg(TOT_RF_FUNDO = ('Valor Bruto',sum))


# verificar pivot table

patrimonio = df_patrimonio[['Código da Carteira','Patrimônio']].copy()
patrimonio.rename(columns={'Código da Carteira': 'Carteira/Fundo'}, inplace = True)


#------------------------COLUNAS DA BASE DE DADOS--------------------------------------------#
#Aqui definimos as colunas que usaremos para a base de dados do Liquidez.
# Para os Futuros e Renda Fixa, temos um De-Para

#Acoes
carteira_acoes = df_car_acoes[['Carteira/Fundo', 'Código', 'Qtde Total','Qtde Bloqueada', 'Valor Mercado']].copy()
carteira_acoes.set_index('Carteira/Fundo', inplace = True)
carteira_acoes["CODIGO BBG"] = carteira_acoes["Código"]+" BZ EQUITY"


#Futuro
carteira_futuros = df_car_futuros[['Carteira/Fundo', 'Ativo','Vencimento', 'Quantidade','Valor de Mercado']].copy()
carteira_futuros.set_index('Carteira/Fundo', inplace = True)
lista_futuros = (carteira_futuros["Ativo"]+carteira_futuros["Vencimento"]).tolist()



#RF
carteira_renda_fixa = df_car_RF[['Carteira/Fundo', 'Nome','Vencimento', 'Quantidade ','Valor Bruto']].copy()
carteira_renda_fixa.set_index('Carteira/Fundo', inplace = True)
lista_RF = (carteira_renda_fixa["Nome"]).tolist()
lista_vencimento_RF = (carteira_renda_fixa["Vencimento"]).tolist()

#-----------------------DE-PARA COD BBG FUTUROS ----------------------------------------#

futuros_bbg = []
cadastrar = []
for el in lista_futuros:
    fut = el.split(' ')[1]
    if fut[0:3].strip() == "DDI":
        futuros_bbg.append("EV" + fut[3:6] + ' CURNCY')
    elif fut[0:3].strip() == "DI1":
        futuros_bbg.append("OD" + fut[3:6] + ' COMDTY')
    elif fut[0:3].strip() == "ISP":
        futuros_bbg.append("BSP" + fut[3]+fut[5] + ' INDEX')
    elif fut[0:3].strip() == "IND":
        futuros_bbg.append("BZ" + fut[3]+fut[5] + ' INDEX')
    elif fut[0:3].strip() == "DOL":
        futuros_bbg.append("UC" + fut[3:6] + ' CURNCY')    
    elif fut[0:3].strip() == "DAP":
        futuros_bbg.append("WL" + fut[3:6] + ' COMDTY')        
    elif fut[0:3].strip() == "CAN":
        futuros_bbg.append("CAI" + fut[3]+fut[5] + ' CURNCY')        
    elif fut[0:3].strip() == "AUS":
        futuros_bbg.append("AUO" + fut[3]+fut[5] + ' CURNCY')        
    elif fut[0:3].strip() == "EUP":
        futuros_bbg.append("URO" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "GBR":
        futuros_bbg.append("BPB" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "JAP":
        futuros_bbg.append("JPO" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "T10":
        futuros_bbg.append("OI" + fut[3]+fut[5] + ' COMDTY')  
    elif fut[0:3].strip() == "WIN":
        futuros_bbg.append("XB" + fut[3]+fut[5] + ' INDEX')     
    elif fut[0:3].strip() == "AFS":
        futuros_bbg.append("AZA" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "CNH":
        futuros_bbg.append("CNS" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "MEX":
        futuros_bbg.append("MEO" + fut[3]+fut[5] + ' CURNCY')      
    elif fut[0:3].strip() == "WDO":
        futuros_bbg.append("WDO" + fut[3]+fut[5] + ' INDEX')    
    elif fut[0:3].strip() == "WSP":
        futuros_bbg.append("WSP" + fut[3]+fut[5] + ' INDEX')     
 
    else:
        print( 'CADASTRAR:', fut)
        

        
carteira_futuros["CODIGO BBG"] = futuros_bbg
carteira_futuros["Código"] = lista_futuros



#-----------------------DE-PARA COD BBG RENDA FIXA ----------------------------------------#
# Esse de-para tem uma peculiaridade.
#O codigo que usamos para encontrar o ativo na BBG, não é o mesmo que encontramos na planilha dos Volumes.
# Desse modo criamos duas colunas com o Codigo do Ativo (um para o bbg, outro para a planilha)


nome_RF = []
Vencimento_RF = []
renda_fixa_bbg = []
Vencimento_RF_INDEX = []

for el in lista_RF:
    if el == 'LETRA  FIN.TES.NAC.':
        nome_RF.append('BLFT 0')
    elif el == 'LETRA TES. NACIONAL':
        nome_RF.append('BLTN 0')
    elif el == 'NTN-B':
        nome_RF.append('BNTNB 6')    
    elif el.split(' ')[1] == 'OVER':
        nome_RF.append('COMPROMISSADA')        

for el in lista_vencimento_RF:
    dia = el.split('/')[0]
    mes = el.split('/')[1]
    ano = el.split('/')[2][2:4]
    Vencimento_RF.append(mes+"/"+dia+'/'+ano)
    Vencimento_RF_INDEX.append(mes + " "+ dia + ' '+ ano)



carteira_renda_fixa["BBG INDEX"] = Vencimento_RF_INDEX
carteira_renda_fixa["Nome BBG"] = nome_RF
carteira_renda_fixa["Vencimento"] = Vencimento_RF
carteira_renda_fixa["CODIGO BBG"] = carteira_renda_fixa["Nome BBG"]+ " " + carteira_renda_fixa["Vencimento"] + " " + '@CBSM Govt'
carteira_renda_fixa["CODIGO"] = carteira_renda_fixa["Nome BBG"]+ " " + carteira_renda_fixa["BBG INDEX"] + " " + 'CBSM GOVT'
carteira_renda_fixa = carteira_renda_fixa[carteira_renda_fixa.CODIGO.str.contains("COMPROMISSADA") == False]



#----------------------OLHAR NA PLANILHA PARA VER OVO LUME MÉDIO ----------------------------------------#
#Essa etapa te joga pra planilha que pega os dados da BBG e já calcula a média dos ativos de acordo com o gap da data que for inputada

df_BBG = pd.DataFrame()
df_BBG = carteira_acoes["CODIGO BBG"].reset_index(drop = True).drop_duplicates()
df_BBG = df_BBG.append(carteira_futuros["CODIGO BBG"].reset_index(drop = True).drop_duplicates())
df_BBG = df_BBG.append(carteira_renda_fixa["CODIGO BBG"].reset_index(drop = True).drop_duplicates())
df_BBG = df_BBG[df_BBG.str.contains("COMPROMISSADA") == False]


lista = df_BBG.to_list()
lista2 = []
for el in lista:
    el = str(el).replace(' ', '%20')
    el = str(el).replace('/', '%2F')
    el = str(el).replace('@', '%40')
    lista2.append(el)
lista2 = ','.join(lista2)

a = "http://192.168.211.59:8050/bbg/bdh?tickers={}&start_date={}&end_date={}&fields=VOLUME&fill=zeros&datesfill=ACTIVE_DAYS_ONLY".format(lista2,data_sheet_bbg_inicio,data_sheet_bbg_final)
r = requests.get(a)


df_vol = pd.read_csv(io.StringIO(r.content.decode('utf-8')))

df_vol = pd.DataFrame(eval(r.content))



l= []
f = []
r = []
for el in carteira_acoes["CODIGO BBG"].values:
    l.append(df_vol["{} VOLUME".format(el)].mean())
    
for el in carteira_futuros["CODIGO BBG"].values:
    f.append(df_vol["{} VOLUME".format(el)].mean())    

for el in carteira_renda_fixa["CODIGO"].values:
    r.append(df_vol["{} VOLUME".format(el)].mean()) 

carteira_acoes["VOLUME MÉDIO"] = l
carteira_futuros["VOLUME MÉDIO"] = f
carteira_renda_fixa["VOLUME MÉDIO"] = r

#-----------------------AJUSTA A BASE COM AS INFOS ----------------------------------------#



# Calculando os dias para liquidar usando o Alfa definido lá em cima
carteira_acoes["DIAS PARA LIQUIDAR"] = (abs(carteira_acoes['Qtde Total'])/ (carteira_acoes["VOLUME MÉDIO"] * ALFA)+2)
carteira_futuros["DIAS PARA LIQUIDAR"] = carteira_futuros['Quantidade'] / (carteira_futuros["VOLUME MÉDIO"] * ALFA)
carteira_renda_fixa["DIAS PARA LIQUIDAR"] = carteira_renda_fixa['Quantidade '] / (carteira_renda_fixa["VOLUME MÉDIO"] * ALFA)

#Arrumando o Dataframe Final do Acoes e acrescentando as colunas de proporcao
merged_acoes = pd.merge(df_tot_acoes, carteira_acoes, how = "inner", on = 'Carteira/Fundo')
merged2_acoes = pd.merge(merged_acoes, patrimonio, how = "inner", on = 'Carteira/Fundo')
merged2_acoes['% NO FUNDO'] = merged2_acoes['Valor Mercado'] / merged2_acoes['Patrimônio']
merged2_acoes['% EM AÇÃO'] = merged2_acoes['Valor Mercado'] / merged2_acoes['TOT_EQ_FUNDO']
merged2_acoes['Valor de Mercado'] = merged2_acoes['Valor Mercado']
merged2_acoes['CLASSE'] = 'ACAO'

principal_acoes = merged2_acoes[['Carteira/Fundo', 'Código','CODIGO BBG', 'Qtde Total','Valor de Mercado', 'Patrimônio', 'DIAS PARA LIQUIDAR',"VOLUME MÉDIO",
        '% NO FUNDO', '% EM AÇÃO','CLASSE']].copy()

Dias_Para_liquidar_ACOES = principal_acoes.groupby(["Carteira/Fundo"]).agg(DIAS_PARA_LIQUIDAR_AÇOES = ('DIAS PARA LIQUIDAR',max))


#Arrumando o Dataframe Final do Futuros e acrescentando as colunas de proporcao
merged_futuros = pd.merge(df_tot_futuros, carteira_futuros, how = "inner", on = 'Carteira/Fundo')
merged2_futuros = pd.merge(merged_futuros, patrimonio, how = "inner", on = 'Carteira/Fundo')
merged2_futuros['% NO FUNDO'] = abs(merged2_futuros['Valor de Mercado']) / merged2_futuros['Patrimônio']
merged2_futuros['% EM FUTUROS'] = abs(merged2_futuros['Valor de Mercado']) / merged2_futuros['TOT_FUT_FUNDO']
merged2_futuros['Código'] = lista_futuros
merged2_futuros['Qtde Total'] = merged2_futuros['Quantidade']
merged2_futuros['CLASSE'] = 'FUTURO'

principal_futuros = merged2_futuros[['Carteira/Fundo', 'Código','CODIGO BBG', 'Qtde Total','Valor de Mercado', 'Patrimônio', 'DIAS PARA LIQUIDAR',"VOLUME MÉDIO",
        '% NO FUNDO', '% EM FUTUROS','CLASSE']].copy()

Dias_Para_liquidar_FUTUROS = principal_futuros.groupby(["Carteira/Fundo"]).agg(DIAS_PARA_LIQUIDAR_FUTUROS = ('DIAS PARA LIQUIDAR',max))



#Arrumando o Dataframe Final do RF e acrescentando as colunas de proporcao
merged_RF = pd.merge(df_tot_RF, carteira_renda_fixa, how = "inner", on = 'Carteira/Fundo')
merged2_RF = pd.merge(merged_RF, patrimonio, how = "inner", on = 'Carteira/Fundo')
merged2_RF['% NO FUNDO'] = merged2_RF['Valor Bruto'] / merged2_RF['Patrimônio']
merged2_RF['% EM RF'] = merged2_RF['Valor Bruto'] / merged2_RF['TOT_RF_FUNDO']
merged2_RF['Código'] = merged2_RF['Nome'] + ' ' + merged2_RF['Vencimento']
merged2_RF['Qtde Total'] = merged2_RF['Quantidade ']
merged2_RF['Valor de Mercado'] = merged2_RF['Valor Bruto']
merged2_RF['CLASSE'] = 'RF'

principal_RF = merged2_RF[['Carteira/Fundo', 'Código','CODIGO BBG', 'Qtde Total','Valor de Mercado', 'Patrimônio', 'DIAS PARA LIQUIDAR',"VOLUME MÉDIO", '% NO FUNDO', '% EM RF','CLASSE']].copy()

Dias_Para_liquidar_RF = principal_RF.groupby(["Carteira/Fundo"]).agg(DIAS_PARA_LIQUIDAR_RF = ('DIAS PARA LIQUIDAR',max))

#-----------------------EXPORTA DATAFRAME PARA SER A BASE DO RELATORIO DE LIQUIDEZ ----------------------------------------#


LIQUIDEZ = pd.concat([principal_futuros, principal_acoes, principal_RF], ignore_index=True)

data2 = nome_data(menostres,menosum).strftime('%d/%m/%Y')
LIQUIDEZ['DATA'] =  data2


open(path_destino_LIQUIDEZ,"w").close()

LIQUIDEZ.to_csv(path_destino_LIQUIDEZ,  index=None, sep='\t', mode='a')

print('VASCO PORRA')
time.sleep(2)












