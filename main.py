from Entities.extraction_imobme import ExtractionImobme
from Entities.tratar_dados import TratarDados, datetime, pd, relativedelta
from Entities.dependencies.folder_register import FolderRegister
import os
from Entities.dependencies.arguments import Arguments
from typing import List

class Execute:
    @staticmethod
    def start(date = datetime.now(), extract_imobme:bool=True, save:bool=True):
        
        
        path_incc = os.path.join(FolderRegister(r"PATRIMAR ENGENHARIA S A\RPA - Documentos\RPA - Dados\Indices").value, "indices_financeiros.json")
        if not os.path.exists(path_incc):
            raise FileNotFoundError(f"A planilha '{path_incc=}' não foi encontrada!")
        df_incc = pd.read_json(path_incc)
        
        path_ic = os.path.join(os.getcwd(), r'IC_BASE\Base Estoque.xlsx')
        if not os.path.exists(path_ic):
            raise FileNotFoundError(f"A planilha '{path_ic=}' não foi encontrada!")
        df_ic = pd.read_excel(path_ic, sheet_name='Base Estoque')
        
        imobme = ExtractionImobme()
        if extract_imobme:
            imobme.start([
                "imobme_empreendimento",
                "imobme_controle_vendas"
            ])
        
        files_path = TratarDados.get_files(imobme.download_path)
                
        df_mes_selecionado = TratarDados(paths=files_path, incc=df_incc, ic=df_ic, date=date).start()
        
        path_target_file = os.path.join(FolderRegister(r'PATRIMAR ENGENHARIA S A\RPA - Documentos\RPA - Dados\Relatorios\Relatorio VGV Empreendimentos').value, 'base.json')
        
        if os.path.exists(path_target_file):
            df_target = pd.read_json(path_target_file)
        else:
            df_target = pd.DataFrame()
            
        # import pdb; pdb.set_trace()
        # df_target[df_target['create_date'] >= (date - relativedelta(years=2)).replace(day=1).strftime("%Y-%m-%dT00:00:00.000")]['create_date'].unique()
        
        if not df_target.empty:
            df_target = df_target[
                pd.to_datetime(df_target['create_date']) != date.replace(day=1).strftime("%Y-%m-%dT00:00:00.000")
            ]  
            
        
        df = pd.concat([df_target, df_mes_selecionado], ignore_index=True)
        
        if save:
            df.to_json(path_target_file, orient='records', date_format='iso')
        
    @staticmethod
    def start_all():
        date_now = datetime.now().replace(day=1,hour=0,minute=0,second=0,microsecond=0)
        dates:List[datetime] = []
        for _ in range(30):
            dates.append(date_now)
            date_now = date_now - relativedelta(months=1)
        
        for date in dates:
            msg = f"Iniciado {date} "
            print(msg, end='\r')
            Execute.start(date=date, extract_imobme=False)
            print(f"{msg} - Finalizado!")
    
    @staticmethod     
    def test():
        Execute.start(date=datetime.now(), extract_imobme=False, save=False)
              
if __name__ == "__main__":
    Arguments({
        'start': Execute.start,
        'start_all': Execute.start_all,
        'teste': Execute.test
    })
    