"""
Este módulo contém funções para tratar dados de relatórios de VGV de empreendimentos.
"""
import pandas as pd
from calculations import financial_utils
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import numpy as nb
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class TratarDados:
    """
    Classe para tratar dados de relatórios de VGV de empreendimentos.
    """
    
    @staticmethod
    def get_files(folder_path: str):
        """
        Obtém os caminhos dos arquivos de empreendimentos e vendas em uma pasta.

        Args:
            folder_path (str): Caminho da pasta contendo os arquivos.

        Returns:
            dict: Dicionário com os caminhos dos arquivos de empreendimentos e vendas.

        Raises:
            FileNotFoundError: Se o caminho da pasta não for encontrado.
            TypeError: Se o caminho fornecido não for uma pasta.
        """
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"o caminho '{folder_path}' não foi encontrado")
        if not os.path.isdir(folder_path):
            raise TypeError(f"o caminho '{folder_path}' não é de uma pasta")
        
        files_path: dict = {}
        for file in os.listdir(folder_path):
            file = os.path.join(folder_path, file)
            if "Empreendimentos_" in os.path.basename(file):
                files_path["Empreendimentos"] = file
            elif "Vendas_" in os.path.basename(file):
                files_path["Vendas"] = file
        
        return files_path
    
    @property
    def df_empre(self) -> pd.DataFrame:
        """
        DataFrame contendo os dados dos empreendimentos.

        Returns:
            pd.DataFrame: DataFrame dos empreendimentos.
        """
        return self.__df_empre
    
    @property
    def df_vendas(self) -> pd.DataFrame:
        """
        DataFrame contendo os dados das vendas.

        Returns:
            pd.DataFrame: DataFrame das vendas.
        """
        return self.__df_vendas
    
    @property
    def df_incc(self) -> pd.DataFrame | pd.Series:
        """
        DataFrame ou Series contendo os dados do INCC.

        Returns:
            pd.DataFrame | pd.Series: DataFrame ou Series do INCC.
        """
        return self.__df_incc
    
    @property
    def df_ic(self) -> pd.DataFrame:
        """
        DataFrame contendo os dados do estoque.

        Returns:
            pd.DataFrame: DataFrame do estoque.
        """
        return self.__df_ic
    
    @property
    def inccMenos2(self) -> pd.Series:
        """
        Obtém o valor do INCC de dois meses atrás.

        Returns:
            pd.Series: Série contendo o valor do INCC de dois meses atrás.

        Raises:
            Exception: Se não for possível encontrar o INCC de dois meses atrás.
        """
        result = self.df_incc[pd.to_datetime(self.df_incc['Data']) == (self.date.replace(day=1) - relativedelta(months=2))]
        if not result.empty:
            return result.iloc[0]
        raise Exception("não foi possivel encontrar o Incc menos 2")
    
    def __init__(self, *, paths: dict, incc: pd.DataFrame, ic: pd.DataFrame, date: datetime):
        """
        Inicializa a classe TratarDados.

        Args:
            paths (dict): Dicionário contendo os caminhos dos arquivos de empreendimentos e vendas.
            incc_path (str): Caminho do arquivo de dados do INCC.
            ic_path (str): Caminho do arquivo de dados do estoque.
            date (datetime): Data de referência para os dados.

        Raises:
            FileNotFoundError: Se algum dos arquivos não for encontrado.
        """
        if not os.path.exists(paths['Empreendimentos']):
            raise FileNotFoundError(f"A planilha '{paths['Empreendimentos']=}' não foi encontrada!")
        if not os.path.exists(paths['Vendas']):
            raise FileNotFoundError(f"A planilha '{paths['Vendas']=}' não foi encontrada!")
        
        self.date = datetime(date.year, date.month, date.day)
        self.__df_empre = pd.read_excel(paths['Empreendimentos'])
        
        self.__df_vendas = pd.read_excel(paths['Vendas'])
        self.__df_vendas = self.__df_vendas[
            pd.to_datetime(self.__df_vendas['Data Base Proposta']) <= self.date
        ]
        
        self.__df_incc = incc.copy()
        self.__df_incc = self.__df_incc[self.__df_incc['Indice'] == 'INCC'][['Data', 'Valor']]
        self.__df_incc['Data'] = self.__df_incc['Data'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d'))
        
        self.__df_ic = ic.copy()
        self.__df_ic = self.__df_ic[
            pd.to_datetime(self.__df_ic['Referência Tabela']) == self.date.replace(day=1)
        ]
        
    def start(self) -> pd.DataFrame:
        """
        Inicia o tratamento dos dados e retorna um DataFrame com os resultados.

        Returns:
            pd.DataFrame: DataFrame contendo os dados tratados.
        """
        df = self.df_empre[['Código Empreendimento', 'Nome Do Empreendimento', 'Nome Do Bloco', 'Código Da Unidade', 'Área Privativa', 'PEP Unidade']].copy()
        
        df_vendas = self.df_vendas
        df_incc = self.df_incc
        df_ic = self.df_ic
        incc_menos2 = self.inccMenos2
                
        
        df['PV MAIOR S/ EXTRA'] = df.apply(financial_utils.findPVMaior, axis=1, args=(df_vendas,))
        df['DATA BASE PROPOSTA'] = df.apply(financial_utils.findDataProposta, axis=1, args=(df_vendas,))
        df['INCC -2 PROPOSTA'] = df.apply(financial_utils.findINCC, axis=1, args=(df_incc,))
        df['REAJUSTE BASE VIABILIDADE'] = df.apply(financial_utils.findReajusteViabidade, axis=1, args=(incc_menos2,))
        df['PV MAIOR REAJUSTADO'] = df.apply(financial_utils.findPVMaiorReajustado, axis=1)
        df['PREÇO ESTOQUE'] =  df.apply(financial_utils.findPrecoEstoque, axis=1, args=(df_ic,))
        df['TIPO'] =  df.apply(financial_utils.findTipo, axis=1, args=(df_ic,))
        df['create_date'] = self.date.replace(day=1)        
        
        return df

if __name__ == "__main__":
    pass
