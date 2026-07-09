import os
import subprocess
import sys
import time
from pathlib import Path
from datetime import datetime

MAIN_DIR = Path(__file__).resolve().parents[0]



start_time = time.time()
mes_ano_atual = datetime.now().strftime("%m-%Y")

print("RODANDO GRUPO 1")
subprocess.run([sys.executable, MAIN_DIR / 'grupo1' / 'grupo1.py'], check=True)

print("RODANDO SEPARAÇÃO GRUPO 1")
subprocess.run([sys.executable, MAIN_DIR / 'grupo1' / 'separador1.py'], check=True)
print("+ GRUPO 1 CONCLUÍDO +")

############################################################################################################

print("RODANDO GRUPO 2")
subprocess.run([sys.executable, MAIN_DIR / 'grupo2' / 'grupo2.py'], check=True)

print("RODANDO SEPARAÇÃO GRUPO 2")
subprocess.run([sys.executable, MAIN_DIR / 'grupo2' / 'separador2.py'], check=True)
print("+ GRUPO 2 CONCLUÍDO +")

elapsed_time = time.time() - start_time
hours = int(elapsed_time // 3600)
minutes = int((elapsed_time % 3600) // 60)
seconds = int(elapsed_time % 60)

print(f"Tempo total de execução: {hours:02d}h{minutes}m{seconds}s")