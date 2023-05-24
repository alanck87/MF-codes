# -*- coding: utf-8 -*-
"""
Created on Tue Apr 18 13:46:27 2023

@author: akligerman
"""

import pandas as pd
from datetime import datetime, timedelta
import getpass
import shutil
import os
from pandas.tseries.offsets import BDay
import calendar

#---------------AUX---------------------#

date_format = '%m/%d/%Y'


dolar = 5.15

def str_to_zero(x):
    if isinstance(x, str):
        return 0
    else:
        return x


def ecxeldate_to_normal(df, column):
    l = list(df[column].to_numpy())
    lista_datas = []
    for el in l:
        if el >0 :
            excel_date = int(el)
            dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_date - 2).strftime('%m/%d/%Y')
            lista_datas.append(dt)
        else:
            lista_datas.append(el)
    return  lista_datas


def get_dates_in_month(month, year):
    
    
    days_in_month = calendar.monthrange(year, month)[1]
    dates = []
    for day in range(1, days_in_month + 1):
        date_str = f"{month:02}/{day:02}/{year}"
        dates.append(date_str)
    return dates
#----------Paths utilizado--------------#

class BASE_RESULTADO():
    def __init__(self):
        self.data = '04/28/2023'
        self.data_br ='2023_04_28'
        self.data_us = '20230428'
        self.data_anterior = '2023_04_27'
        self.data_anterior2 = '04/27/2023'
        self.data_trades = '20230428-20230428'
        self.paths = {
            'FPRs': "R:\PLANILHAS PRINCIPAIS\PRICING - NOVUS.xlsm",
            'cotas': "P:\DADOS DIÁRIOS\SÉRIE HISTÓRICA - PATRIMÔNIO.xlsx",
            'resultado': "Z:\_ARQUIVOS INPUTS\RESULTADO LOTE45\RESULTADO - {}.txt".format(self.data_br),
            'rel_gerencial': r"R:\PLANILHAS PRINCIPAIS\RELATÓRIO GERENCIAL.xlsb",
            'gerencial_cortado': "R:\ARQUIVOS OUTPUTS\BASE GERENCIAL\BASE GERENCIAL - {}.xlsb".format(self.data_anterior),
            'destino_base_gerencial': r"R:\ARQUIVOS OUTPUTS\BASE GERENCIAL\BASE GERENCIAL - TESTE.xlsx",
            'corretagem_BMF' : "Z:\_ARQUIVOS INPUTS\SAG\CORRETAGEM BMF - {}.xlsx".format(self.data_br),
            'corretagem_BOVESPA' : "Z:\_ARQUIVOS INPUTS\SAG\CORRETAGEM BOVESPA - {}.xlsx".format(self.data_br),
            'corretagem_gs' : r"Z:\_ARQUIVOS INPUTS\GOLDMAN\CORRETAGEM_GS_{}.xls".format(self.data_us), 
            'corretagem_equity' : "Z:\_ARQUIVOS INPUTS\CONTROLE DE CORRETAGEM\Offshore\BTG CAYMAN.xls",
            'trades_d-1' : "Z:\_ARQUIVOS INPUTS\TRADES LOTE45\FundsTrades-{}.txt".format(self.data_trades)           
        }
        
        
        

    def read_dataframes(self):
        
        
        
        df_FPR = pd.read_excel(self.paths['FPRs'], sheet_name='Cadastro FPRs', usecols=[1, 3])
        df_FPR = df_FPR.rename(columns={'FPR2': 'FPR', 'Ativo': 'Product'})
        
        df_serie_CDI = pd.read_excel(self.paths['rel_gerencial'], sheet_name='SÉRIE - % CDI', engine='pyxlsb')
        df_serie_CDI['DATA'] = ecxeldate_to_normal(df_serie_CDI, 'DATA.BASE')
    
        df_AUX = pd.read_excel(self.paths['rel_gerencial'], sheet_name='TABELAS AUXILIARES', engine='pyxlsb', skiprows=3, usecols=[1,2,3,4,5,6,7])
        
        df_base_result = pd.read_excel(self.paths['gerencial_cortado'], sheet_name='BASE - RESULTADO', engine='pyxlsb')
        df_base_result = df_base_result.drop(columns=['SEMANAL','MENSAL','SEMESTRAL','ANUAL','TRIMESTRE', 'IS_STOCK_MACRO','nome', 'len', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23'])
        df_base_result['DATA'] = ecxeldate_to_normal(df_base_result, "DATA")
        
        df_resultado_data_base = pd.read_csv(self.paths['resultado'], sep="\t", header=0)
        
        df_séries_cotas = pd.read_excel(self.paths['rel_gerencial'], sheet_name='SÉRIE - COTAS', engine='pyxlsb', parse_dates=True)
        
        df_qntd_cotas =  pd.read_excel(self.paths['rel_gerencial'], sheet_name='SÉRIE - QTD_COTAS', engine='pyxlsb', parse_dates=True)
        
        return [df_FPR, df_serie_CDI, df_AUX, df_base_result, df_séries_cotas,df_resultado_data_base, df_qntd_cotas]
        
 
         
        
    def corretagem(self):   
                
        dic_corretoras = {'SAFDIE' : 'MODAL','BTG' : 'BTG PACTUAL CM', 'CONCORDIA S/A C' : 'CONCÓRDIA S/A C.V.M.C.C.', 'ITAÚ' : 'ITAÚ CORRETORA DE VALORES S/A',
                          'TERRA' : 'Terra', 'XP' : 'XP INVESTIMENTOS CCTVM LTDA.',  'ORAMA DTVM S A' : 'ORAMA'}
        
        df_Corretagem = pd.DataFrame()
        
        df_bmf = pd.read_excel(self.paths['corretagem_BMF'],sheet_name='Custos Operacionais', skiprows = 4)
        df_bovespa = pd.read_excel(self.paths['corretagem_BOVESPA'],sheet_name='Custos Operacionais', skiprows = 4)
        html = pd.read_html(self.paths['corretagem_gs'], skiprows = 2)
        df_gs = html[0]
        df_gs = df_gs.set_axis(df_gs.iloc[0], axis=1)
        df_gs = df_gs[1:]
        
        df_equity = pd.read_excel(self.paths['corretagem_equity'],sheet_name='History',engine='xlrd', skiprows = 2)
        df_trades = pd.read_csv(self.paths['trades_d-1'], sep = '\t')
        
        df_bmf["Nome / Razão Social"] = df_bmf["Nome / Razão Social"].apply(lambda x: x.split('INT')[0])
        #df_bmf = df_bmf.drop_duplicates(subset=["Nome / Razão Social", 'Ativo','Corretora' ])
        
        
        
        df_bovespa["Nome / Razão Social"] = df_bovespa["Nome / Razão Social"].apply(lambda x: x.split('INT')[0])

        
        #df_bovespa = df_bovespa.drop_duplicates(subset=["Nome / Razão Social", 'Ativo','Corretora' ])

        


        df_bmf['BMF RATIO'] =  df_bmf['Corretagem Execução'] + df_bmf['Corretagem Clearing'] + df_bmf['Emolumentos'] + df_bmf['Outros Custos'] + df_bmf['Taxa de registro'] + df_bmf['Taxa de permanência']
        df_bmf['Ativo'] =  df_bmf['Ativo'] +  df_bmf['Série']
        df_bmf = df_bmf.rename(columns={"Corretora": 'Dealer'})
        df_bmf['Dealer'] = df_bmf['Dealer'].replace(dic_corretoras)
        result = df_bmf.groupby(["Nome / Razão Social",'Ativo','Dealer' ]).agg({'Qtde': 'sum', 'BMF RATIO' :'sum' })
        
        
        
        
        df_bovespa['BVSP RATIO'] =  (df_bovespa['Corretagem Execução'] + df_bovespa['Corretagem Clearing'] + df_bovespa['Emolumentos'] + df_bovespa['Taxa de registro'])
        df_bovespa = df_bovespa.rename(columns={"Corretora": 'Dealer'})
        df_bovespa['Dealer'] = df_bovespa['Dealer'].replace(dic_corretoras)
        result2 = df_bovespa.groupby(["Nome / Razão Social",'Ativo','Dealer' ]).agg({'Qtde': 'sum', 'BVSP RATIO' :'sum', 'Preço' : 'mean' })
        
        
        
        #df_gs2 = df_gs.replace('.', ',')
        #df_gs2[['Clearing + Execution Commission (local)','Total Fees (local)', 'Quantity']] = df_gs[['Clearing + Execution Commission (local)','Total Fees (local)', 'Quantity']].applymap(float)
        #df_gs['RATIO FUT OFF'] = df_gs['Clearing + Execution Commission (local)'] + df_gs['Total Fees (local)']
        #df_gs = df_gs.rename(columns={'Executing Broker': 'Dealer'})
        #result3 = df_gs.groupby(["Account",'Bloomberg Code','Dealer']).agg({'Quantity': 'sum', 'RATIO FUT OFF' :'sum'})
        
        
        #df_bmf = df_bmf[[ "Ativo", 'Corretora', 'Qtde', 'BMF RATIO']]
        #df_bovespa = df_bovespa[[ "Ativo", 'Corretora', 'Qtde', 'BVSP RATIO']]
        #df_bmf.dropna(inplace=True)
        #df_bovespa.dropna(inplace=True)
        
        
        df_trades = df_trades.rename(columns={'Product' : 'Ativo'})
        


        
        l =[]
        
        df_Corretagem = pd.merge(df_trades,result, on = ['Ativo', 'Dealer'],  how='left' ) 
        df_Corretagem = pd.merge(df_Corretagem, result2, on = ['Ativo', 'Dealer'],  how='left' ) 

        '''
        for index, row in df_Corretagem.iterrows():
            if row['ProductClass'] == 'Equity':
                l.append(row['BVSP RATIO'] * row['Price'] * row['Amount'])

            else: 
                l.append(row['BMF RATIO'] * row['Amount'])
                
        
        df_Corretagem['CORRETAGEM'] = l
        
        df_Corretagem = df_Corretagem[['Trading Desk', 'Ativo',  'Book', 'CORRETAGEM' ]] 
        '''
        
        
        
        return  [result, result2, df_trades, df_Corretagem]
           
 
    
 
    def arruma_série_cotas(self,df,date_format):
       self.lista_datas = ecxeldate_to_normal(df,'DATA BASE')
       df['DATA BASE'] = self.lista_datas
       df.set_index('DATA BASE', inplace = True)
       headers =['COTA - INSTITUCIONAL MASTER FIM',
        'INSTITUCIONAL MASTER FIM',
        'COTA - INSTITUCIONAL FIC FIM',
        'INSTITUCIONAL FIC FIM',
        'COTA - NOVUS CAPITAL MASTER FIM',
        'NOVUS CAPITAL MASTER FIM',
        'COTA - NOVUS CAPITAL FIC FIM',
        'NOVUS CAPITAL FIC FIM',
        'COTA - NC RENDA FIXA EXCLUSIVE FIC',
        'NC RENDA FIXA EXCLUSIVE FIC', #add o fic
        'COTA - PREV ADVISORY FIC FIM',
        'PREV ADVISORY FIC FIM',
        'COTA - PREV ADVISORY',
        'PREV ADVISORY',
        'COTA - NOVUS FIXED INCOME',
        'NOVUS FIXED INCOME',
        'COTA - RAPTOR FIRF',
        'RAPTOR FIRF',
        'COTA - FLAG MASTER FIM',
        'FLAG MASTER FIM',
        'COTA - NOVUS MACRO D5 FIC FIM',
        'NOVUS MACRO D5 FIC FIM',
        'COTA - NOVUS ACOES INSTITUCIONAL',
        'NOVUS ACOES INSTITUCIONAL',
        'COTA - NOVUS ACOES INSTITUCIONAL FIC FIA',
        'NOVUS ACOES INSTITUCIONAL FIC FIA',
        'COTA - NOVUS MACRO II FIC FIM',
        'NOVUS MACRO II FIC FIM',
        'COTA - TOTAL RETURN SP',
        'TOTAL RETURN SP',
        'COTA - NOVUS RETORNO ABSOLUTO',
        'NOVUS RETORNO ABSOLUTO',
        'COTA - NOVUS RETORNO ABSOLUTO FIC FIM',
        'NOVUS RETORNO ABSOLUTO FIC FIM',
        'COTA - NOVUS SP',
        'NOVUS SP',
        'COTA - NOVUS PREV RENDA FIXA',
        'NOVUS PREV RENDA FIXA',
        'COTA - NOVUS PREV RENDA FIXA FIC FIM',
        'NOVUS PREV RENDA FIXA FIC FIM',
        'COTA - NOVUS MACRO I FIC FIM',
        'NOVUS MACRO I FIC FIM',
        'COTA - NOVUS MACRO RELIANCE FIC FIM',
        'NOVUS MACRO RELIANCE FIC FIM',
        'COTA - PATRIMÔNIO - NOVUS RETORNO ABSOLUTO ITAU',
        'NOVUS RETORNO ABSOLUTO ITAU',
        'COTA - PATRIMÔNIO - RETORNO ABSOLUTO FIC FIA',
        'RETORNO ABSOLUTO FIC FIA',
        'COTA - NC VALOR ACOES',
        'NC VALOR ACOES',
        'COTA - NOVUS PREV FI SP',
        'NOVUS PREV FI SP',
        'COTA - NOVUS PREV RETORNO ABSOLUTO',
        'NOVUS PREV RETORNO ABSOLUTO',
        'COTA - NOVUS PREV RETORNO ABSOLUTO II FIC FIA',
        'NOVUS PREV RETORNO ABSOLUTO II FIC FIA',
        'COTA - NOVUS PREV RETORNO ABSOLUTO XP SEGUROS',
        'NOVUS PREV RETORNO ABSOLUTO XP SEGUROS',
        'COTA - PREV II INST',
        'PREV II INST',
        'COTA - NOVUS MACRO FOF 2',
        'NOVUS MACRO FOF 2',
        'COTA - NOVUS MACRO A PREV 2',
        'NOVUS MACRO A PREV 2',
        'COTA - NOVUS MACRO PREV IQ',
        'NOVUS MACRO PREV IQ',
        'COTA - NOVUS MACRO FOF FIC',
        'NOVUS MACRO FOF FIC',
        'COTA - NOVUS ABSOLUTO A',
        'NOVUS ABSOLUTO A',
        'COTA - NOVUS RED RENDA FIXA',
        'NOVUS RED RENDA FIXA',
        'COTA - NOVUS PREV INST XP SEGUROS',
        'NOVUS PREV INST XP SEGUROS',
        'COTA - NC PETROS FIM',
        'NC PETROS FIM', 
        'COTA - NOVUS RENDA FIXA ITAU PREV FIE I FIC FI LP',
        'NOVUS RENDA FIXA ITAU PREV FIE I FIC FI LP',
        'COTA - NOVUS PREV INSTITUCIONAL',
        'NOVUS PREV INSTITUCIONAL',
        'COTA - NOVUS RENDA FIXA',
        'NOVUS RENDA FIXA']
       df.columns = headers
       u= df.stack()
       u = u.reset_index()
       u.rename(columns={'level_1': 'FUNDO', 'DATA BASE': 'ValDate', 0 : 'COTAS' }, inplace=True)

       return u   
   
    def qntd_cotas(self, df):
        self.lista_datas = ecxeldate_to_normal(df,'DATA BASE')
        df['DATA BASE'] = self.lista_datas
        df.set_index('DATA BASE', inplace = True)
        u= df.stack()
        u = u.reset_index()
        u.rename(columns={'level_1': 'FUNDO', 'DATA BASE': 'ValDate', 0 : 'COTAS' }, inplace=True)
        #print(u)
        datas_dia = list(self.tabela_data_dia_semana(date_format))
        data_mes_ano = list(self.tabela_data_dia_mes_ano(date_format))
        data_mes_ano.append(datas_dia[1])
        fund = []
        rat = []
        dat = []
        n = 0
        dif = ['mes', 'ano', 'semana']
        df = pd.DataFrame()
        for el in data_mes_ano:
            fund = []
            rat = []
            dat = []
            filtered_df = u[u['ValDate'] == el]
            filtered_df.drop(columns=['ValDate'])
            filtered_df.rename(columns={'COTAS': 'COT'}, inplace=True)
            index_cotas = pd.merge(u, filtered_df, on=['FUNDO'],  how='left' )

            for index, row in index_cotas.iterrows():
                if type(row['COTAS']) == float and type(row['COT']) == float:
                    ratio = float(row['COTAS']) / float(row['COT'])
                    #print(row['COTAS'], row['COT'], row['FUNDO'],row['ValDate_x'] )
                    rat.append(ratio)
                else: 
                    rat.append(0)
                fundo = row['FUNDO']#.split(' - ')[1]
                data = str(row['ValDate_x'])
                fund.append(fundo)
                dat.append(data)

            df['FUNDO'] = fund
            df['RATIO {}'.format(dif[n])] = rat
            df['ValDate'] = dat

            n+=1
        df.dropna()  
        return df   
        
    def tabela_data_dia_semana(self, date_format):  
        date = datetime.strptime(self.data, date_format)
        weekday = date.weekday()
        dia = self.data
        first_semana = (date - timedelta(days=weekday)).strftime(date_format)
        last_semana = (date - timedelta(days=weekday+3)).strftime(date_format)
        return [dia,first_semana, last_semana]

    def tabela_data_dia_mes_ano(self, date_format): 
        data_mod = datetime.strptime(self.data, date_format)     
        this_year = (data_mod.year)
        this_month = (data_mod.month)
        filtered_month = [date for date in self.lista_datas if datetime.strptime(date, date_format).month == this_month and datetime.strptime(date, date_format).year == this_year]
        filtered_year = [date for date in self.lista_datas if datetime.strptime(date, date_format).year == this_year]
        return [min(filtered_month), min(filtered_year)]
    
    def sheet_posicao(self,df_result, df_AUX,df_FPR):
        df_result.rename(columns={'Book': 'BOOK', 'TradingDesk' : 'FUNDO', 'PositionBenchmarkPL': 'CARRY COST', 'PositionExBenchmarkPLTotal' : 'RESULTADO EX-BENCHMARK', 'ValDate': 'DATA','PositionExBenchmarkPL Pct':'RESULTADO %'}, inplace=True)
        df_result = df_result[["FUNDO",'BOOK','Product','DATA','CARRY COST','RESULTADO %','RESULTADO EX-BENCHMARK']]
        df_result['MÊS'] = self.data.split('/')[0]
        df_result['RESULTADO %'] = df_result['RESULTADO %'].str.replace(',', '.')
        df_result['RESULTADO EX-BENCHMARK'] = df_result['RESULTADO EX-BENCHMARK'].str.replace(',', '.')
        df_result['CARRY COST'] = df_result['CARRY COST'].str.replace(',', '.')
        df_resultado_AUX = pd.merge(df_result, df_AUX, on = 'BOOK')
        Base_Atualizada = pd.merge(df_resultado_AUX, df_FPR, how = 'left',  on = 'Product')
        Base_Atualizada.rename(columns={'Product': 'ATIVO'}, inplace=True)
        
        return Base_Atualizada

    def coluna_Resultado_posicao(self, sheet_posicao,serie_cota):
        l = []
        for index, row in pd.DataFrame(sheet_posicao).iterrows():
            patrimonio = serie_cota.loc[(serie_cota['FUNDO'] == row['FUNDO']) & (serie_cota['ValDate'] == self.data_anterior2), 'COTAS'].iloc[0]            
            if row['FPR'] in ['CAIXA', 'CUSTOS']:
                 resultado = (float(row['CARRY COST']) + float(row['RESULTADO EX-BENCHMARK'])) / patrimonio
                 l.append(resultado)
            else: 
                resultado = (float(row['RESULTADO EX-BENCHMARK'])) / patrimonio
                l.append(resultado)
        sheet_posicao = sheet_posicao.drop(columns=['RESULTADO %'])
        sheet_posicao['RESULTADO %'] =  l      
        return sheet_posicao

    def semanal(self, datas_dia_semana, serie_cota,base_resultado_gerada):
        filtered_df = serie_cota[serie_cota['ValDate'] == datas_dia_semana[0]]
        filtered_df = filtered_df.drop(columns=['ValDate'])

        monday = datetime.strptime(datas_dia_semana[1], '%m/%d/%Y')  # Convert to datetime object
        dates = []
        for i in range(5):  # 5 days from Monday to Friday
            day = (monday + timedelta(days=i)).strftime('%m/%d/%Y')

            dates.append(day)

        filtered_df2 = serie_cota[serie_cota['ValDate'].isin(dates)]
        filtered_df2 = filtered_df2.drop(columns=['COTAS'])
        
        filtered_avec_cotas = pd.merge(filtered_df2, filtered_df, on=['FUNDO'],  how='left' )
        filtered_avec_cotas.rename(columns={'ValDate' : "DATA"}, inplace=True)    

        base_resultado_gerada = pd.merge(base_resultado_gerada, filtered_avec_cotas, on=['DATA', 'FUNDO'],  how='left')
        
        a = base_resultado_gerada['COTAS'].tolist()

        return a

    def mensal(self ,datas_mes_ano,  serie_cota,base_resultado_gerada):
        filtered_df = serie_cota[serie_cota['ValDate'] == datas_mes_ano[0]]
        filtered_df = filtered_df.drop(columns=['ValDate'])

        dates = get_dates_in_month(int(self.data.split('/')[0]), int(self.data.split('/')[2]))
        
        filtered_df2 = serie_cota[serie_cota['ValDate'].isin(dates)]
        filtered_df2 = filtered_df2.drop(columns=['COTAS'])
        filtered_avec_cotas = pd.merge(filtered_df2, filtered_df, on=['FUNDO'],  how='left' )
        filtered_avec_cotas.rename(columns={'ValDate' : "DATA"}, inplace=True) 
        
        base_resultado_gerada = pd.merge(base_resultado_gerada, filtered_avec_cotas, on=['DATA', 'FUNDO'],  how='left' )
        a = base_resultado_gerada['COTAS'].tolist()
        return a
    
    def ano(self, datas_mes_ano, serie_cota,base_resultado_gerada):
        filtered_df = serie_cota[serie_cota['ValDate'] == datas_mes_ano[1]]
        filtered_df.rename(columns={'ValDate' : "DATA"}, inplace=True)    
        base_resultado_gerada = pd.merge(base_resultado_gerada, filtered_df, on=['FUNDO'],  how='left' )
        a = base_resultado_gerada['COTAS'].tolist()
        return a


    def CDI(self, base_resultado_gerada, df_serie_CDI):

        base_resultado_gerada = pd.merge(base_resultado_gerada, df_serie_CDI, on=['DATA'],  how='left' )
        filtered_df = df_serie_CDI[df_serie_CDI['DATA'] == self.data]
        base_resultado_gerada['INDEX DIA']  = list(filtered_df['INDEX'].to_numpy())[0]
        base_resultado_gerada['Index Ratio'] = base_resultado_gerada['INDEX DIA'] / base_resultado_gerada['INDEX']
        base_resultado_gerada = base_resultado_gerada.drop(columns=['INDEX DIA', 'INDEX', 'CDI', 'DATA.BASE', 'CDI.DIARIO'])
        return base_resultado_gerada

    

    def colunas_janelas(self, base_resultado_gerada,df_serie_CDI,datas_mes_ano,datas_dia_semana, serie_cota):
                
        base_resultado_gerada = self.CDI(base_resultado_gerada, df_serie_CDI)
        print(1)

        base_resultado_gerada['PAT SEMANAL'] = self.semanal(datas_dia_semana, serie_cota, base_resultado_gerada)
        
        base_resultado_gerada['PAT MENSAL'] =self.mensal(datas_mes_ano, serie_cota,base_resultado_gerada)
        
        base_resultado_gerada['PAT ANUAL'] = self.ano(datas_mes_ano, serie_cota,base_resultado_gerada)
        
        
        base_resultado_gerada.fillna(0, inplace=True)

        l_semana = []
        l_mes = []
        l_ano = []
        for index, row in base_resultado_gerada.iterrows():

            if row['PAT SEMANAL'] == 0 or row['RATIO semana'] == 0:
                l_semana.append(0)
                
            elif row['PAT SEMANAL'] != 0 and row['RATIO semana'] != 0:
                l_semana.append((float(row['RESULTADO EX-BENCHMARK']) / float(row['RATIO semana'])) * (row['Index Ratio']))

                
            if row['PAT MENSAL'] == 0 or row['RATIO mes'] == 0:
                l_mes.append(0)
                
            elif row['PAT MENSAL'] != 0 and row['RATIO mes'] != 0:
                l_mes.append((float(row['RESULTADO EX-BENCHMARK']) / float(row['RATIO mes'])) * (row['Index Ratio']))
    
            if row['PAT ANUAL'] == 0 or row['RATIO ano'] == 0:
                l_ano.append(0)
                
            elif row['PAT ANUAL'] != 0 and row['RATIO ano'] != 0:
    
                l_ano.append((float(row['RESULTADO EX-BENCHMARK']) / float(row['RATIO ano'])) * (row['Index Ratio']))

        base_resultado_gerada['SEMANAL'] = l_semana
        base_resultado_gerada['MENSAL'] = l_mes  
        base_resultado_gerada['SEMESTRAL'] = l_ano
        base_resultado_gerada['ANUAL'] = l_ano
    
        
        base_resultado_gerada = base_resultado_gerada.drop(columns=['PAT SEMANAL','RATIO mes','RATIO ano','RATIO semana', 'PAT MENSAL', 'PAT ANUAL', 'Index Ratio'])
        
        
        return base_resultado_gerada



paths = BASE_RESULTADO()

df_corretagem = paths.corretagem()

#dfs = paths.read_dataframes()



#serie_cota = paths.arruma_série_cotas(dfs[4],date_format)
#datas_dia_semana = paths.tabela_data_dia_semana(date_format)
#datas_mes_ano = paths.tabela_data_dia_mes_ano(date_format)

#sheet_posicao = paths.sheet_posicao(dfs[5], dfs[2], dfs[0])
#sheet_qntd_cotas = paths.qntd_cotas(dfs[6])

#sheet_qntd_cotas.rename(columns={'ValDate' : "DATA"}, inplace=True)  

#POSICAO = paths.coluna_Resultado_posicao(sheet_posicao, serie_cota)
#base_resultado_gerada =  pd.merge(pd.concat([dfs[3], POSICAO], ignore_index=True),sheet_qntd_cotas, on=['DATA', 'FUNDO'],  how='left') 
#BASE_RESULTDO = paths.colunas_janelas(base_resultado_gerada,dfs[1],datas_mes_ano,datas_dia_semana, serie_cota)




#BASE_RESULTDO = BASE_RESULTDO.reset_index(drop = True)

#BASE_RESULTDO.to_excel(r"R:\ARQUIVOS OUTPUTS\BASE GERENCIAL\BASE GERENCIAL - TESTE.xlsx")
