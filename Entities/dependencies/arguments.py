import sys
from typing import Dict, List
from Entities.dependencies.logs import Logs, traceback
from typing import Literal

class Arguments:
    def __init__(self, valid_arguments:Dict[str, object], log_status:Literal['Error', 'Concluido', 'Report', 'Test']='Concluido') -> None:
        self.__valid_arguments = valid_arguments
        self.__argv:list = sys.argv
        self.__status:Literal['Error', 'Concluido', 'Report', 'Test'] = log_status
        
        self.__start()
    
    
    def __start(self):
        if len(self.__argv) > 1:
            selected_argv = self.__argv[1]
            if selected_argv in self.__valid_arguments:
                try:
                    if len(self.__argv) == 3:
                        self.__valid_arguments[selected_argv](self.__argv[2]) #type: ignore
                    elif len(self.__argv) > 3:
                        self.__valid_arguments[selected_argv](self.__argv[2:]) #type: ignore
                    else:
                        self.__valid_arguments[selected_argv]() #type: ignore
                    Logs().register(status=self.__status, description="Automação concluida com sucesso!")
                except Exception as error:
                    print(type(error), str(error))
                    Logs().register(status='Error', description=str(error), exception=traceback.format_exc())
            else:
                print("argumento não existe!")
                self.__listar_argvs()
        else:
            self.__listar_argvs()
            
    def __listar_argvs(self):
        print("são permitido apenas os seguintes argumentos: ")
        for key, value in self.__valid_arguments.items():
            print(key)

def teste(args):
    print(args)
    
if __name__ == "__main__":
    Arguments(valid_arguments={
        "teste": teste
    })