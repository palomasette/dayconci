import pandas as pd
import openpyxl
import datetime
from credentials import USER_ATIVA, PASS_ATIVA
import xlsxwriter
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import numpy as np
from utils import convert_string_to_float
from sinqiaconn import Connect

class Dados:
    def __init__(self) -> None:
        self.data = datetime.date.today().strftime("%d_%m_%Y")
        self.path = "historico\\CONCI_DAYC_ATV_{data}.xlsx"
    
    def daycoval_data(self, file_path):
        # definindo a ultima coluna como header
        df1 = pd.read_csv(file_path, header=None, on_bad_lines='skip',
                                    sep=';', encoding='latin-1')
        df1.columns = df1.iloc[-1]
        df1 = df1[:-1]

        # filtrando 'SALDO FINAL' 
        df1 = df1[df1['Hist'] == 'SALDO FINAL']

        # colocando em ordem alfabetica
        df1 = df1.sort_values(by='NomeCrt')

        # considerando apenas as colunas necessarias
        df_daycoval = df1[['NomeCli', 'NomeCrt', 'ValBrt', 'CNPJCrt']]
        df_daycoval = df_daycoval.reset_index(drop=True)

        # convertendo de string para float
        df_daycoval['ValBrt'] = df_daycoval['ValBrt'].astype(str)
        df_daycoval['ValBrt'] = df_daycoval['ValBrt'].apply(convert_string_to_float)
        # eliminando os valores nulos
        df_daycoval = df_daycoval[df_daycoval['ValBrt'] != 0.00]
        
        df_daycoval['NomeCli'] = df_daycoval['NomeCli'].str.replace('ATIVA ', '')
        
        df_daycoval = df_daycoval.groupby(['NomeCli', 'CNPJCrt', 'NomeCrt'])['ValBrt'].sum().reset_index()
        df_daycoval = df_daycoval.sort_values(by='NomeCrt')
        df_daycoval['CNPJCrt'] = df_daycoval['CNPJCrt'].str.replace('[\./\-]', '', regex=True)

        df_daycoval.rename(columns={
                    'NomeCli': 'Cliente',
                    'CNPJCrt': 'CNPJ_Fundo',
                    'NomeCrt': 'Nome_Fundo', 
                    'ValBrt': 'Posição_Dayc'
                }, inplace=True)
        return df_daycoval
    
    def ativa_data(self):
        file_45 = Connect().get45(USER_ATIVA, PASS_ATIVA)

        df_45 = pd.read_csv(file_45, header=None, on_bad_lines='skip',
                                            sep=';', encoding='latin-1')
        # definindo a ultima coluna como header
        df_45.columns = df_45.iloc[-1]
        df_45 = df_45[:-1]

        # considerando apenas as colunas necessárias
        df_45 = df_45[['Cliente', 'NomeLCrt', 'ValorAtu','CPF_CNPJ']]

        # convertendo de string para float
        df_45['ValorAtu'] = df_45['ValorAtu'].astype(str)
        df_45['ValorAtu'] = df_45['ValorAtu'].apply(convert_string_to_float)
        # eliminando os valores nulos
        df_45 = df_45[df_45['ValorAtu'] != 0.00]
        #ajustando a coluna Cliente
        df_45['Cliente'] = df_45['Cliente'].astype(str).str.replace(',0', '')
        #organizando em ordem alfabetica os fundos
        df_45 = df_45.sort_values(by='NomeLCrt')

        df_45 = df_45.groupby(['Cliente', 'NomeLCrt', 'CPF_CNPJ'])['ValorAtu'].sum().reset_index()
        df_45 = df_45.sort_values(by='NomeLCrt')
        
        df_45.rename(columns={
                    'CPF_CNPJ': 'CNPJ_Fundo',
                    'NomeLCrt': 'Nome_Fundo', 
                    'ValorAtu': 'Posição_Ativa'
                }, inplace=True)
        
        return df_45
    

