import os
from dotenv import load_dotenv

load_dotenv()

DATUM_USER = os.getenv('DATUM_USERNAME')
DATUM_PASSWORD = os.getenv('DATUM_PASSWORD')
DATUM_BASE_URL = os.getenv('DATUM_BASE')

AREAS_BLACKLIST = ["pruebas"]