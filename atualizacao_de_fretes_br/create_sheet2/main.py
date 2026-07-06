import os
import subprocess
import sys
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
MAIN_DIR = Path(os.getenv('MAIN_DIR'))


subprocess.run([sys.executable, MAIN_DIR / 'create_sheet2' / 'create.py'])

subprocess.run([sys.executable, MAIN_DIR / 'create_sheet2' / 'formatacoes.py'])