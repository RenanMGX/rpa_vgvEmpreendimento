import pandas as pd
from calculations import financial_utils
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import numpy as nb
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class TratarDados:
    @staticmethod
    def get_files(folder_path:str):
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"o caminho '{folder_path}' não foi encontrado")
        if not os.path.isdir(folder_path):
            raise TypeError(f"o caminho '{folder_path}' não é de uma pasta")
        
        files_path:dict = {}
        for file in os.listdir(folder_path):
            file = os.path.join(folder_path, file)
            if "Empreendimentos_" in os.path.basename(file):
                files_path["Empreendimentos"] = file
            elif "Vendas_" in os.path.basename(file):
                files_path["Vendas"] = file
        
        return files_path
    
    
    @property
    def df_empre(self) -> pd.DataFrame:
        return self.__df_empre
    
    @property
    def df_vendas(self) -> pd.DataFrame:
        return self.__df_vendas
    
    @property
    def df_incc(self) -> pd.DataFrame|pd.Series:
        return self.__df_incc
    
    @property
    def df_ic(self) -> pd.DataFrame:
        return self.__df_ic
    
    @property
    def inccMenos2(self) -> pd.Series:
        result = self.df_incc[pd.to_datetime(self.df_incc['Data']) == (self.date.replace(day=1) - relativedelta(months=2))]
        if not result.empty:
            return result.iloc[0]
        raise Exception("não foi possivel encontrar o Incc menos 2")
    
    def __init__(self, *, paths:dict, incc_path:str, ic_path:str, date:datetime):
        if not os.path.exists(paths['Empreendimentos']):
            raise FileNotFoundError(f"A planilha '{paths['Empreendimentos']=}' não foi encontrada!")
        if not os.path.exists(paths['Vendas']):
            raise FileNotFoundError(f"A planilha '{paths['Vendas']=}' não foi encontrada!")
        if not os.path.exists(incc_path):
            raise FileNotFoundError(f"A planilha '{incc_path=}' não foi encontrada!")
        if not os.path.exists(ic_path):
            raise FileNotFoundError(f"A planilha '{ic_path=}' não foi encontrada!")
        
        self.date = datetime(date.year,date.month,date.day)
        self.__df_empre = pd.read_excel(paths['Empreendimentos'])
        
        self.__df_vendas = pd.read_excel(paths['Vendas'])
        self.__df_vendas = self.__df_vendas[
            pd.to_datetime(self.__df_vendas['Data Base Proposta']) <= self.date
        ]
        
        self.__df_incc = pd.read_json(incc_path)
        self.__df_incc = self.__df_incc[self.__df_incc['Indice'] == 'INCC'][['Data', 'Valor']]
        self.__df_incc['Data'] = self.__df_incc['Data'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d'))
        
        self.__df_ic = pd.read_excel(ic_path, sheet_name='Base Estoque')
        self.__df_ic = self.__df_ic[
            pd.to_datetime(self.__df_ic['Referência Tabela']) == self.date.replace(day=1)
        ]
        
    def start(self) -> pd.DataFrame:
        df = self.df_empre[['Código Empreendimento', 'Nome Do Empreendimento', 'Nome Do Bloco', 'Código Da Unidade', 'Área Privativa']]
        
        df['PV MAIOR S/ EXTRA'] = df.apply(financial_utils.findPVMaior, axis=1, args=(self.df_vendas,))
        df['DATA BASE PROPOSTA'] = df.apply(financial_utils.findDataProposta, axis=1, args=(self.df_vendas,))
        df['INCC -2 PROPOSTA'] = df.apply(financial_utils.findINCC, axis=1, args=(self.df_incc,))
        df['REAJUSTE BASE VIABILIDADE'] = df.apply(financial_utils.findReajusteViabidade, axis=1, args=(self.inccMenos2,))
        df['PV MAIOR REAJUSTADO'] = df.apply(financial_utils.findPVMaiorReajustado, axis=1)
        df['PREÇO ESTOQUE'] =  df.apply(financial_utils.findPrecoEstoque, axis=1, args=(self.df_ic,))
        df['TIPO'] =  df.apply(financial_utils.findTipo, axis=1, args=(self.df_ic,))
        df['create_date'] = self.date.replace(day=1)        
        
        
        return df

if __name__ == "__main__":
    pass
