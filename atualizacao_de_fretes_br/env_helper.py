from pathlib import Path
import os
from dotenv import load_dotenv

load_dotenv()

def get_main_dir() -> Path:
    """Return MAIN_DIR from .env as a Path. Falls back to repository root if unset."""
    main = os.getenv('MAIN_DIR')
    if main:
        return Path(main)
    return Path(__file__).resolve().parent
