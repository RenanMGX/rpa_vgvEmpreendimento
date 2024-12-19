import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import numpy as nb

def findPVMaior(row:pd.Series, df:pd.DataFrame):
    bloco = row['Nome Do Bloco']
    unidade = row['Código Da Unidade']
    
    result:pd.Series = df[
        (df['Bloco'] == bloco) &
        (df['Unidade'] == unidade) &
        (~df['Status Unidade'].isin(['Bloqueada', 'Disponível', 'Análise de Crédito / Risco']))
    ]['PV Maior s/ Extra']
    
    return result.iloc[-1] if not result.empty else ""

def findDataProposta(row:pd.Series, df:pd.DataFrame):
    bloco = row['Nome Do Bloco']
    unidade = row['Código Da Unidade']
    
    result:pd.Series = df[
        (df['Bloco'] == bloco) &
        (df['Unidade'] == unidade) &
        (~df['Status Unidade'].isin(['Bloqueada', 'Disponível', 'Análise de Crédito / Risco']))
    ]['Data Base Proposta']
    
    return result.iloc[-1] if not result.empty else ""


def findINCC(row:pd.Series, df:pd.DataFrame):
    data:datetime = row['DATA BASE PROPOSTA']
    try:
        data = (data.replace(day=1)) - relativedelta(months=2)
        
        result:pd.Series = df[
           pd.to_datetime(df['Data']) == data
        ]['Valor']

        return result.iloc[0] if not result.empty else ""
    except:
        return ""
    
def findReajusteViabidade(row:pd.Series, incc:pd.Series):
    try:
        inncProposta = row['INCC -2 PROPOSTA']
        incc = incc['Valor']
        return incc/inncProposta
        
    except:
        return ""
    
def findPVMaiorReajustado(row:pd.Series):
    try:
        pmMaior = row['PV MAIOR S/ EXTRA']
        reajuste = row['REAJUSTE BASE VIABILIDADE']
        #print(pmMaior, reajuste, pmMaior * reajuste)
        return round(pmMaior * reajuste, 2)
        
    except:
        return ""
    
    
def findPrecoEstoque(row:pd.Series, df:pd.DataFrame):
    bloco = row['Nome Do Bloco']
    try:
        unidade = str(int(row['Código Da Unidade']))
    except:
        unidade = row['Código Da Unidade']
    
    result:pd.Series = df[
        (df['Torre/Bloco'].astype(str) == bloco) &
        (df['Unidade/Casa'].astype(str) == unidade)
    ]['Preço - Margem de Desconto']
    
    return result.iloc[0] if not result.empty else ""

def findTipo(row:pd.Series, df:pd.DataFrame):
    bloco = row['Nome Do Bloco']
    try:
        unidade = str(int(row['Código Da Unidade']))
    except:
        unidade = row['Código Da Unidade']
    
    result:pd.Series = df[
        (df['Torre/Bloco'].astype(str) == bloco) &
        (df['Unidade/Casa'].astype(str) == unidade)
    ]['Tipo']
    
    return result.iloc[0] if not result.empty else ""    

if __name__ == "__main__":
    pass
