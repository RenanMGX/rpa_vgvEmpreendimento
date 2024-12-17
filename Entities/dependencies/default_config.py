from typing import Dict
from getpass import getuser

default:Dict[str, Dict[str,object]] = {
    'credential': {
        'navegador': 'KPMG'
    },
    'log': {
        'hostname': 'Patrimar-RPA',
        'port': '80',
        'token': ''
    },
    'url': {
        "default": "https://krast.kpmg.com.br"
    },
    'path': {
        'sharepoint': 'RPA - Dados\\Relatorios Auditoria\\KPMG'
    }
}