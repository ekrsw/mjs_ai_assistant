import os
import logging

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
    
    # プロキシ設定
    PROXY_HOST = "mjsproxy.mjs.co.jp"  # 例: "proxy.company.com"
    PROXY_PORT = 80  # 例: 8080
    
    # sampleプログラム互換のプロキシ設定
    HTTP_PROXY = "http://@mjsproxy.mjs.co.jp:80"
    HTTPS_PROXY = "http://@mjsproxy.mjs.co.jp:80"
    NO_PROXY = "localhost,127.0.0.1,sv-vw-ejap,*.local,192.168.*,10.*,172.16.*,172.17.*,172.18.*,172.19.*,172.20.*,172.21.*,172.22.*,172.23.*,172.24.*,172.25.*,172.26.*,172.27.*,172.28.*,172.29.*,172.30.*,172.31.*"
    
    # 初期URL設定  
    INITIAL_URL = "http://sv-vw-ejap:5555/SupportCenter/main.aspx?area=nav_answers&etc=127&page=CS&pageType=EntityList&web=true"


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
