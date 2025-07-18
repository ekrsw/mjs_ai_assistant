import os
import logging
from dotenv import load_dotenv

# .envファイルを読み込み
load_dotenv()

class Settings:
    HEADLESS_MODE = True
    RETRY_COUNT = 3
    FILE_PATH = "files"
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    LOG_FILE = os.path.join(BASE_DIR, "log", "app.log")
    FILE = os.path.join(BASE_DIR, FILE_PATH, "data.xlsx")
    
    # ログレベル設定
    LOG_LEVEL = logging.INFO
    LOG_FORMAT = '%(asctime)s:%(levelname)s:%(name)s:%(message)s'
    
    # Selenium関連設定
    SELENIUM_WAIT_TIME = 5
    PAGE_LOAD_WAIT = 3
    
    # プロキシ設定（.envファイルから読み込み）
    PROXY_HOST = os.getenv('PROXY_HOST')
    PROXY_PORT = int(os.getenv('PROXY_PORT', '80'))
    
    # sampleプログラム互換のプロキシ設定
    HTTP_PROXY = os.getenv('HTTP_PROXY')
    HTTPS_PROXY = os.getenv('HTTPS_PROXY')
    NO_PROXY = os.getenv('NO_PROXY')
    
    # 初期URL設定  
    INITIAL_URL = os.getenv('INITIAL_URL', "http://sv-vw-ejap:5555/SupportCenter/main.aspx?area=nav_answers&etc=127&page=CS&pageType=EntityList&web=true")


def setup_logging():
    """アプリケーション全体で使用するロガーの設定を行います"""
    logging.basicConfig(
        level=settings.LOG_LEVEL,
        format=settings.LOG_FORMAT,
        handlers=[
            logging.FileHandler(settings.LOG_FILE, mode='a', encoding='utf-8'),
            logging.StreamHandler()  # コンソールへ出力
        ]
    )
    return logging.getLogger(__name__)


settings = Settings()
