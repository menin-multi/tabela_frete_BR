import os
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
MAIN_DIR = Path(os.getenv('MAIN_DIR'))

start_time = time.time()
mes_ano_atual = datetime.now().strftime("%m-%Y")

print("Criando grupos 1 e 2")
subprocess.run([sys.executable, MAIN_DIR / 'separacao_intelipost' / 'intelipost_add_value' / 'main.py'], check=True)
print("+ Grupos separados +")

print("Criando grupo 3")
subprocess.run([sys.executable, MAIN_DIR / 'separacao_intelipost' / 'atualizar_tabela.py'], check=True)
print("+ Grupo 3 separado +")