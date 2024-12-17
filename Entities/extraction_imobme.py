import os
from selenium import webdriver
from selenium.webdriver.webkitgtk.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep
from datetime import datetime
from typing import Dict
from dateutil.relativedelta import relativedelta
from dependencies.credenciais import Credential
from dependencies.logs import Logs
from functools import wraps
from dependencies.config import Config
from dependencies.functions import P
import traceback

def _verific_login(*, nav:webdriver.Chrome) -> None:
    try:
        nav.find_element(By.XPATH, '//*[@id="login"]')
        nav.maximize_window()
        print(P("Fazendo Login", color='cyan'))
            
        crd = Credential(Config()['credential']['crd']).load()
        nav.find_element(By.XPATH, '//*[@id="login"]').send_keys(crd['login'])
        nav.find_element(By.XPATH, '//*[@id="password"]').send_keys(crd['password'])
        nav.find_element(By.XPATH, '//*[@id="password"]').send_keys(Keys.RETURN)
            
        if nav.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div/ul/li').text == 'Login não encontrado.':
            raise PermissionError("Login não encontrado.")
            
        if 'Senha Inválida.' in (return_error:=nav.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div/ul/li').text):
            raise PermissionError(return_error)
            
        nav.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/button[1]/span').click()
    except:
        pass

def verific_login(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if (b:=kwargs.get('browser')):
            _verific_login(nav=b)
        return f(*args, **kwargs)
    return wrap

@verific_login
def _find_element(*, browser: webdriver.Chrome, mod, target:str, timeout:int=30, can_pass:bool=False, wait:float=.2):
    if wait > 0:
        sleep(wait)
    for _ in range(timeout*4):
        try:
            obj = browser.find_element(mod, target)
            #print(target)
            return obj
        except:
            sleep(0.25)
    
    if can_pass:
        #print(f"{can_pass=}")
        return browser.find_element(By.TAG_NAME, 'html')
    
    raise Exception(f"não foi possivel encontrar o {target=} pelo {mod=}")


class BotExtractionImobme():
    def __init__(self, user:str, password:str, download_path:str=f"{os.getcwd()}\\downloads\\") -> None:
        self.__user:str = user
        self.__password:str = password
        self.download_path:str = download_path
        
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        else:
            for file in os.listdir(self.download_path):
                try:
                    os.unlink(self.download_path + file)
                except PermissionError:
                    os.rmdir(self.download_path + file)
                    
        prefs: dict = {"download.default_directory" : self.download_path}
        chrome_options: Options = Options()
        chrome_options.add_experimental_option("prefs", prefs)     
        
        self.navegador: webdriver.Chrome = webdriver.Chrome(options=chrome_options)
        self.navegador.get("http://patrimarengenharia.imobme.com/Autenticacao/Login")
        self.navegador.maximize_window()
    
    
    def _login(self) -> None:
        _find_element(browser=self.navegador, mod=By.TAG_NAME, target='html')
        return
        self.navegador.get("http://patrimarengenharia.imobme.com/Autenticacao/Login")
        self.navegador.maximize_window()
        
        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="login"]').send_keys(self.__user)
        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="password"]').send_keys(self.__password)
        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="password"]').send_keys(Keys.RETURN)
        
        if _find_element(browser=self.navegador, mod=By.XPATH, target='/html/body/div[1]/div/div/div/div[2]/form/div/ul/li', timeout=1, can_pass=True).text == 'Login não encontrado.':
            Logs().register(status='Report', description="Login não encontrado")
            raise PermissionError("Login não encontrado.")
        
        if 'Senha Inválida.' in (return_error:=_find_element(browser=self.navegador, mod=By.XPATH, target='/html/body/div[1]/div/div/div/div[2]/form/div/ul/li', timeout=1, can_pass=True).text):
            Logs().register(status='Report', description=return_error)
            raise PermissionError(return_error)
        
        _find_element(browser=self.navegador, mod=By.XPATH, target='/html/body/div[2]/div[3]/div/button[1]/span', timeout=2, can_pass=True).click()
    
    def _extrair_relatorio(self, relatories:list) -> list:
        NUM_TENTATIVAS:int = 2
        TEMPO_ESPERA_TENTATIVA:float = .5
        
        if not isinstance(relatories, list):
            Logs().register(status='Report', description=f"para extrair relatorios apenas listas são permitidas, {relatories=}")
            raise TypeError(f"para extrair relatorios apenas listas são permitidas, {relatories=}")
        if not relatories:
            Logs().register(status='Report', description=f"a lista '{relatories=}' não pode estar vazia")
            raise ValueError(f"a lista '{relatories=}' não pode estar vazia")
        print(P(relatories))
        
        self.relatories_id: dict = {}
        
        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
        self.navegador.maximize_window()
        
        for rel in relatories:
            if (relatorie:="imobme_empreendimento") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_8"]').click() # clique em IMOBME - Empreendimento
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em selecionar Emprendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em selecionar todos os empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em selecionar Emprendimentos novamente para sair
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break
                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

                
            elif (relatorie:="imobme_controle_vendas") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_6"]').click() # clique em IMOBME - Contre de Vendas
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="TipoDataSelecionada_chzn"]/a').click() # clique em Tipo Data
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="TipoDataSelecionada_chzn_o_0"]').click() # clique em Data Lançamento Venda
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break
                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")
                
            elif (relatorie:="imobme_controle_vendas_90_dias") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_6"]').click() # clique em IMOBME - Contre de Vendas
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys((datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")) # escreve a data de inicio com um range de 90 dias
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="TipoDataSelecionada_chzn"]/a').click() # clique em Tipo Data
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="TipoDataSelecionada_chzn_o_0"]').click() # clique em Data Lançamento Venda
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clique em Empreendimentos
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break
                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")
                
            elif (relatorie:="imobme_contratos_rescindidos") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_4"]').click()  # clique em IMOBME - Contratos Rescindicos
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click()  # clique em Tipo de Contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click() # clique em Tipo de Contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                    
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break
                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")
                
            elif (relatorie:="imobme_contratos_rescindidos_90_dias") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_4"]').click()  # clique em IMOBME - Contratos Rescindicos
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys((datetime.now() - relativedelta(days=90)).strftime("%d%m%Y")) # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click()  # clique em Tipo de Contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div/div[3]/div/button').click() # clique em Tipo de Contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clique em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break
                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_dados_contrato") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_7"]').click() # clique em IMOBME - Dados de Contrato
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[1]/div').click() # clica fora
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[3]/div[1]/div/button').click() # clica em tipos de contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[3]/div[1]/div/ul/li[2]/a/label/input').click() # clica em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[1]/div').click() # clica fora
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_previsao_receita") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_9"]').click() # clique em IMOBME - Previsão de Receita
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys((datetime.now() + relativedelta(years=25)).strftime("%d%m%Y")) # escreve a data de fim padrao com a data atual mais 25 anos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        sleep(1)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button').click() # clica em Tipo Parcela
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[4]/div/div[2]/div/ul/li[2]/a/label/input').click() # clica em Todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[4]/div/div[2]/div/button').click() # clica em Tipo Parcela
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_relacao_clientes") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(2):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_10"]').click() # clique em IMOBME - Relação de Clientes
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_relacao_clientes_x_clientes") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(2):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_10"]').click() # clique em IMOBME - Relação de Clientes
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="tipoReportCliente_chzn"]').click() # clica em tipo de relatorio
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="tipoReportCliente_chzn_o_1"]').click() # clica em clientes x contratos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvTipoContrato"]/div/button').click() # clica em tipo de contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvTipoContrato"]/div/ul/li[2]/a/label/input').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvTipoContrato"]/div/button').click() # clica em tipo de contrato
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012015") # escreve a data de inicio padrao 01/01/2015
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_cadastro_datas") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_3"]').click()  # clique em IMOBME - Cadastro de Datas
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="recebimentos_compensados") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_16"]').click()  # clique em Recebimentos Compensados
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataInicio"]').send_keys("01012020") # escreve a data de inicio padrao 01/01/2020
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataFim"]').send_keys(datetime.now().strftime("%d%m%Y")) # escreve a data de hoje
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/ul/li[2]/a/label').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="dvEmpreendimento"]/div[1]/div/div/button').click() # clica em Empreendimentos

                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

            elif (relatorie:="imobme_controle_estoque") == rel:
                finalizou:bool = False
                for num in range(5):
                    try:
                        for _ in range(NUM_TENTATIVAS):
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Content"]').location_once_scrolled_into_view
                            _find_element(browser=self.navegador, mod=By.ID, target='Relatorios_chzn').click() # clique em selecionar Relatorios
                            _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Relatorios_chzn_o_5"]').click() # clique em IMOBME - Controle de Estoque
                            sleep(TEMPO_ESPERA_TENTATIVA)
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button').click() # clica em Empreendimentos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div[1]/div/div/ul/li[2]/a/label/input').click() # clica em todos
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="parametrosReport"]/div[2]/div[1]/div/div/button').click() # clica em Empreendimentos novamente para sair
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="DataBase"]').send_keys(datetime.now().strftime("%d%m%Y"))
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="Header"]/div[1]/img[1]').click() #<-------------------
                        
                        _find_element(browser=self.navegador, mod=By.XPATH, target='//*[@id="GerarRelatorio"]').click() # clica em gerar relatorio
                        sleep(7)
                        self.relatories_id[relatorie] = self.navegador.find_element(By.XPATH, '//*[@id="result-table"]/tbody/tr[1]/td[1]').text
                        finalizou = True
                        print(P(f"relatorio {relatorie} foi gerado!"))
                        break

                    except Exception as error:
                        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
                        Logs().register(status='Report', description=f"erro {num} ao gerar {relatorie=}", exception=traceback.format_exc())
                        sleep(1)
                if not finalizou:
                    raise Exception(f"erro ao gerar {relatorie=}")

        if len(self.relatories_id) >= 1:
            return list(self.relatories_id.values())
        else:
            Logs().register(status='Report', description="nenhum relatório foi gerado")
            raise FileNotFoundError("nenhum relatório foi gerado")
        
    def start(self, relatories:list) -> None:
        self._login()
        
        print(P("iniciando geração de relatorios"))
        sleep(1)
        relatories_id:list = self._extrair_relatorio(relatories=relatories)
        #relatories_id = ['25347']
        
        self.navegador.get("https://patrimarengenharia.imobme.com/Relatorio/")
        #verificar itens para download
        print(P("Iniciando verificação de download", color='cyan'))
        cont_final: int = 0
        while True:
            if cont_final >= 1080:
                print(P("saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas"))
                Logs().register(status='Report', description=f"saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas")
                raise TimeoutError("saida emergencia acionada a espera da geração dos relatorios superou as 1,5 horas")
            else:
                cont_final += 1
            if not relatories_id:
                break
            
            try:
                table = _find_element(browser=self.navegador, mod=By.ID, target='result-table')
                tbody = table.find_element(By.TAG_NAME, 'tbody')
                for tr in tbody.find_elements(By.TAG_NAME, 'tr'):
                    for id in relatories_id:
                        if id == tr.find_elements(By.TAG_NAME, 'td')[0].text:
                            for tag_a in tr.find_elements(By.TAG_NAME, 'a'):
                                if tag_a.get_attribute('title') == 'Download':
                                    print(P(f"o {relatories_id=} foi baixado!", color='green'))
                                    tag_a.send_keys(Keys.ENTER)
                                    relatories_id.pop(relatories_id.index(id))
            except:
                sleep(5)
                continue
            
            _find_element(browser=self.navegador, mod=By.ID, target='btnProximaDefinicao').click()
            sleep(5)
            print(P("Atualizando Pagina"))
        
        print(P("verificando pasta de download"))
        for _ in range(10*60):
            isnot_excel = False 
            for file in os.listdir(self.download_path):
                if not file.endswith(".xlsx"):
                    isnot_excel = True
            if not isnot_excel:
                sleep(2)
                break
            else:
                sleep(1)
        self.navegador.find_element(By.TAG_NAME, 'html').location
        print(P("extração de relatorios no imobme concluida!"))
        
        
if __name__ == "__main__":
    pass

    # from credential.carregar_credenciais import Credential # type: ignore
    
    # inicio_tempo: datetime = datetime.now()
    # creden: dict = Credential().credencial()
    
    # if (creden['usuario'] == None) or (creden['senha'] == None):
    #     raise PermissionError("Credenciais Invalidas")


    # bot:BotExtractionImobme = BotExtractionImobme(user=creden['usuario'],password=creden['senha'],download_path=f"{os.getcwd()}\\downloads_samba\\")
    # bot.start(["recebimentos_compensados"])
    # bot.navegador.close()
    
    #arquivos = bot.obter_relatorios(["imobme_relacao_clientes_x_clientes"])
    #     #arquivos = bot.obter_relatorios(["imobme_controle_vendas", "imobme_contratos_rescindidos"])
    #     #arquivos = bot.obter_relatorios(["imobme_empreendimento"])
    #     bot.obter_relatorios(["imobme_contratos_rescindidos"])

    # except Exception as error:
    #     print(f"{type(error)} ---> {error.with_traceback()}")
    
    # print(datetime.now() - inicio_tempo)

