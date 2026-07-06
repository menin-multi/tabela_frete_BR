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

subprocess.run([sys.executable, MAIN_DIR / 'atualizacao_vtex' / 'cotacao' /'main.py'], check=True)

subprocess.run([sys.executable, MAIN_DIR / 'atualizacao_vtex' / 'create.py'], check=True)

subprocess.run([sys.executable, MAIN_DIR / 'atualizacao_vtex' / 'formatacao.py'], check=True)

elapsed_time = time.time() - start_time

hours = int(elapsed_time // 3600)
minutes = int((elapsed_time % 3600) // 60)
seconds = int(elapsed_time % 60)


print(f"Tempo total de execução de cotacao, criaçao de planilha e formatação da vtex: {hours:02d}h{minutes}m e {seconds}s")