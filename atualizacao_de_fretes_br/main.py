import os
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from logging_utils import configure_logging

load_dotenv()
MAIN_DIR = Path(os.getenv('MAIN_DIR'))

logger = configure_logging(__name__)
print = logger.info


start_time = time.time()
mes_ano_atual = datetime.now().strftime("%m-%Y")

print("Rodando cotação de fretes BR")
subprocess.run([sys.executable, MAIN_DIR / 'api' / 'main.py'], check=True)
print("+ BR COTADO +")

######################################################################################################################
print("criando a tabela de fretes 1 (sem fórmulas)")
subprocess.run([sys.executable, MAIN_DIR / 'create_sheet' / 'create.py'], check=True)
print("+ TABELAS DE FRETE CRIADAS +")


print("separando arquivos intelipost")
subprocess.run([sys.executable, MAIN_DIR / "separacao_intelipost" / "main.py"], check=True)
print("+ TABELAS DE INTELIPOST SEPARADAS +")


print("criando a tabela de fretes 2 (com fórmulas)")
subprocess.run([sys.executable, MAIN_DIR / 'create_sheet2' / 'main.py'], check=True)
print("+ TABELAS DE FRETE (FÓRMULAS) CRIADAS +")


print("Criando tabela de fretes VTEX")
subprocess.run([sys.executable, MAIN_DIR / 'atualizacao_vtex' / 'main.py'], check=True)
print("+ TABELAS DE FRETE VTEX CRIADAS +")
######################################################################################################################

print("Enviando arquivos por email")
subprocess.run([sys.executable, MAIN_DIR / "email_sender.py"], check=True)
print("+ TABELAS DE EMAIL ENVIADAS +")

elapsed_time = time.time() - start_time
hours = int(elapsed_time // 3600)
minutes = int((elapsed_time % 3600) // 60)
seconds = int(elapsed_time % 60)

print(f"Tempo total de execução: {hours:02d}h{minutes}m{seconds}s")