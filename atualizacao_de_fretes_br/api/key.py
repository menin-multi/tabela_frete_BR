from dotenv import load_dotenv
from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parent
env_path = BASE_DIR / "env 1"

load_dotenv(env_path)

def quote_connection():
    base_url = os.getenv('base_url')
    api_token = os.getenv('api_token')

    return base_url, api_token