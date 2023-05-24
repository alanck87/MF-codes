# -*- coding: utf-8 -*-
"""
Created on Mon Apr  3 17:59:29 2023

@author: akligerman
"""

#add opcao offshore
import pandas as pd
from datetime import datetime, timedelta
import getpass
from pandas.tseries.offsets import BDay
import os
#------------------PATHS---------------#


username = getpass.getuser()



data = datetime.today().strftime('%Y_%m_%d')
data2 = datetime.today().strftime('%Y%m%d')
menosum = (datetime.today() - BDay(1)).strftime('%Y_%m_%d')
menosum2 = (datetime.today() - BDay(1)).strftime('%Y%m%d')
menos1 = (datetime.today() - BDay(1)).strftime('%d%b%Y')
hoje = (datetime.today()).strftime('%d%b%Y')
ontem = (datetime.today() - BDay(1)).strftime('%d%b%Y')
#path do relatorio gerencial
path_gerencial = r"R:\PLANILHAS PRINCIPAIS\RELATÓRIO GERENCIAL.xlsb"

path_pricing = "Z:\_ARQUIVOS OUTPUTS\PRICING NOVA\PRICING - {}.xlsb".format(menosum)


path_resultado = "R:\ARQUIVOS OUTPUTS\BASE GERENCIAL\BASE GERENCIAL - {}.xlsb".format(menosum)


#path do arquivo de posicao D0
path_local_posicao = "Z:\_ARQUIVOS INPUTS\POSIÇÃO LOTE45\POSIÇÃO LOTE45 - {}.txt".format(data)

#path doa trades ainda no pending
path_PendingAllocation = "Z:\Alocação\ARQUIVOS BATIMENTO TRADE\LOTE 45\Trades LOTE45\FundsTrades-{}.txt".format(data2)

#path dos FPRs
path_FPRs = "R:\PLANILHAS PRINCIPAIS\PRICING - NOVUS.xlsm"

#paths de destinos dos arquivos gerados
path_destino_posiçao = "Z:\_ARQUIVOS OUTPUTS\BASE DAYTRADES\BASE BI\POSIÇÃO LOTE45.txt"
path_destino_pending = "Z:\_ARQUIVOS OUTPUTS\BASE DAYTRADES\BASE BI\Pending.txt"

path_patrimonio = "R:\Relatórios\Arquivos LOTE45\Funds Calculated Nav\CalculatedNAV_{}.txt".format(hoje)



path_var_produtos = "R:\Relatórios\Arquivos LOTE45\VaR e Stress - Fundos"

path_destino_base = "Z:\_ARQUIVOS OUTPUTS\BASE DAYTRADES\BASE BI\BASE_RESULT.txt"

path_destino_VAR = "Z:\_ARQUIVOS OUTPUTS\BASE DAYTRADES\BASE BI\BASE_VAR.txt"
#-----------------------dataframes--------------------------#

df_posicao = pd.read_csv(path_local_posicao, sep = '\t')


df_trades = pd.read_csv(path_PendingAllocation, sep = '\t')
df_trades_filtered = df_trades[(df_trades['Trading Desk'] == 'PENDING ALLOCATION') & (df_trades['Trading Desk'] == 'PENDING ALLOCATION OFFSHORE')]
df_trades_filtered = df_trades_filtered[['Trading Desk', 'Trade Date', 'Product', 'Amount']]



df_posicao.rename(columns={'Book': 'BOOK','TradingDesk' : 'Fund'}, inplace=True)

df_AUX = pd.read_excel(path_gerencial, sheet_name = 'TABELAS AUXILIARES', engine='pyxlsb', skiprows=3 , usecols = [1,2,3,4,5,6,7])

df_exposicao = pd.read_excel(path_pricing, sheet_name = 'POS FINAL', engine='pyxlsb')
df_exposicao = df_exposicao[['Product','Lote', 'Exchange Curncy','Delta Unitário', 'Notional Unitário']]
df_exposicao = df_exposicao.drop(index=df_exposicao[df_exposicao['Delta Unitário'] == '0x7'].index)


df_exposicao = df_exposicao.drop_duplicates(subset='Product', keep='first', inplace=False)


df_FPR = pd.read_excel(path_FPRs, sheet_name = 'Cadastro FPRs',  usecols = [1,3])
df_FPR.rename(columns={'FPR2': 'FPR', 'Ativo': 'Product'}, inplace=True)

df_patrimonio =  pd.read_csv(path_patrimonio, sep = '\t',usecols=[0, 1])
#------------------ADICIONANDO AS TABELAS AUXILIARES ------------------#

POSICAOmaisAUX = pd.merge(df_posicao, df_AUX, on = 'BOOK')
AUXmaisFPR = pd.merge(POSICAOmaisAUX, df_FPR, how = 'left',  on = 'Product')

AUXmaisFPR = AUXmaisFPR[['ValDate', 'Fund', 'Product', 'ProductClass', 'BOOK', 'Position', 'FinancialPU', 'Amount', 'YestAmount','TRADER',
       'MERCADO GLOBAL', 'PL', 'FPR']]


AUXmaisFPR['Traded'] = AUXmaisFPR['Amount'] - AUXmaisFPR['YestAmount']
l = []
for index, row in AUXmaisFPR.iterrows():
    if row['YestAmount'] == 0 and row['Amount'] != 0:
        l.append('Open')
        
    elif row['YestAmount'] != 0 and row['Amount'] == 0:  
        l.append('Close')
        
    elif row['YestAmount'] != row['Amount']:  
        l.append('Traded')

    else:
        l.append(" ")
 
AUXmaisFPR['Open/Close'] = l

l_onoff = []
for index, row in AUXmaisFPR.iterrows():
    if row['Fund'] in ['NOVUS SP','TOTAL RETURN SP','NOVUS PREV FI SP' ]:
        l_onoff.append('OFF')
    else:
        l_onoff.append('ON')
        
AUXmaisFPR['On/Off'] = l_onoff
#------------------Buscando Patrimonio por fundo ------------------#

# primeiramente devemos pegar o patrimonio de cada fundo

df_patrimonio.columns = ['Fund', 'Patrimonio']




#o dataframe "u" contem os dados de patrimonio do ultimo dia de cada fundo#


#------------------ Criando Coluna de Exposição ------------------#


AUXmaisFPR =  pd.merge(df_exposicao, AUXmaisFPR, on=['Product'], how='right')

df_merged = pd.merge(df_patrimonio, AUXmaisFPR, on=['Fund'], how='right')

lista_exp = [] 
l_type = []
for index, row in df_merged.iterrows():

    if 'Future' in str(row['ProductClass']):
        exp = (row['FinancialPU'] * row['Amount'] * row['Lote'] * row['Exchange Curncy']) / row['Patrimonio']
        lista_exp.append(exp)
        l_type.append('Futuro')
        
    elif row['ProductClass'] in [ 'Equity', 'US Equity'] :
        if row['Exchange Curncy'] > 0:
            exp = (row['FinancialPU'] * row['Amount']* row['Exchange Curncy']) / row['Patrimonio']
        else:
            exp = (row['FinancialPU'] * row['Amount']* 1) / row['Patrimonio']
        lista_exp.append(exp)

        l_type.append('Ação')
        
    elif 'Swap' in row['ProductClass']:
        exp = (row['Notional Unitário'] * row['Amount']) / row['Patrimonio']
        lista_exp.append(exp)  
        l_type.append('Swap')
        
    elif 'Option' in  row['ProductClass']:
        exp = (row['Delta Unitário'] * row['Amount'] ) / row['Patrimonio']
        lista_exp.append(exp) 
        l_type.append('Opção')
        
    elif row['ProductClass'] == 'ETF':
        exp = (row['FinancialPU'] * row['Amount']* row['Exchange Curncy']) / row['Patrimonio']
        lista_exp.append(exp)  
        l_type.append('Ação')    
    
    else:
        exp = (row['Amount'] * row['Delta Unitário']* row['Exchange Curncy']) / row['Patrimonio']
        lista_exp.append(exp)
        l_type.append('DI')



df_merged['% DAY'] = df_merged['PL'] / df_merged['Patrimonio']
df_merged['EXPOSIÇAO'] = lista_exp
df_merged['Type'] = l_type
df_merged = df_merged.drop('PL', axis=1)
df_merged = df_merged.drop('ValDate', axis=1)

#------------------Passando arquivo para a base do powerBI ------------------#

open(path_destino_pending,"w").close()

df_trades_filtered.to_csv(path_destino_pending,  index=None, sep='\t', mode='a')

open(path_destino_posiçao,"w").close()

df_merged.to_csv(path_destino_posiçao,  index=None, sep='\t', mode='a')

