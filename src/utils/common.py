from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
CACHE_DIR = BASE_DIR / 'cache'

CACHE_DIR.mkdir(exist_ok=True)