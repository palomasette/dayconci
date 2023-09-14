from dotenv import load_dotenv
import os

load_dotenv()

USER_SQL_SINQIA = str(os.getenv('USER_SQL_SINQIA'))
PASS_SQL_SINQIA = str(os.getenv('PASS_SQL_SINQIA'))
USER_ATIVA = str(os.getenv('USER_ATIVA'))
PASS_ATIVA = str(os.getenv('PASS_ATIVA'))