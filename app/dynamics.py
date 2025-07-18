import logging
import time
import traceback
import os

# Webスクレイピング関係ライブラリ
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

from app.config import settings

logger = logging.getLogger(__name__)


class PublishButtonNotActiveError(Exception):
    """承認ボタンが非アクティブで最大リトライ回数に達した場合の例外"""
    pass


class Dynamics:
    """Dynamics CRMのナレッジ記事を更新するためのクラス
    
    このクラスはSeleniumを使用してDynamics CRMのナレッジ記事を
    自動的に更新します。
    """
    def __init__(self, headless_mode=settings.HEADLESS_MODE):
        self.headless_mode = headless_mode
        self.driver = None
        self._initialized = False
    
    def initialize_driver(self):
        """WebDriverを初期化（初回のみ初期URLにアクセス）"""
        if self._initialized and self.driver:
            return True
            
        try:
            # 既存のドライバーがあれば閉じる
            if getattr(self, 'driver', None):
                self.driver.close()
                self.driver = None

            # sampleプログラム互換：環境変数設定
            if hasattr(settings, 'HTTP_PROXY') and settings.HTTP_PROXY:
                os.environ['HTTP_PROXY'] = settings.HTTP_PROXY
                os.environ['http_proxy'] = settings.HTTP_PROXY
                
            if hasattr(settings, 'HTTPS_PROXY') and settings.HTTPS_PROXY:
                os.environ['HTTPS_PROXY'] = settings.HTTPS_PROXY
                os.environ['https_proxy'] = settings.HTTPS_PROXY
                
            if hasattr(settings, 'NO_PROXY') and settings.NO_PROXY:
                os.environ['NO_PROXY'] = settings.NO_PROXY
                os.environ['no_proxy'] = settings.NO_PROXY

            options = Options()

            # ヘッドレスモードの設定
            if self.headless_mode:
                options.add_argument('--headless')
                
            # sampleプログラム互換：安定性のための追加Chromeオプション
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            
            # プロキシ設定
            if hasattr(settings, 'NO_PROXY') and settings.NO_PROXY:
                options.add_argument(f'--proxy-bypass-list={settings.NO_PROXY}')
            
            if hasattr(settings, 'HTTP_PROXY') and settings.HTTP_PROXY:
                options.add_argument(f'--proxy-server={settings.HTTP_PROXY}')
            elif settings.PROXY_HOST and settings.PROXY_PORT:
                proxy_server = f"{settings.PROXY_HOST}:{settings.PROXY_PORT}"
                options.add_argument(f'--proxy-server={proxy_server}')
                logger.debug(f"プロキシを設定しました: {proxy_server}")
            else:
                # プロキシを使用しない設定
                options.add_argument('--no-proxy-server')
                
            options.add_argument('--disable-blink-features=AutomationControlled')
            
            # コンソールログを抑制
            options.add_argument('--disable-logging')
            options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            # webdriver-managerを使用してChromeDriverを自動管理
            logger.info("ChromeDriverを初期化中...")
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.implicitly_wait(settings.SELENIUM_WAIT_TIME)
            
            # 初回のみ：セッション確立のため初期URLにアクセス
            if hasattr(settings, 'INITIAL_URL') and settings.INITIAL_URL:
                logger.debug(f"初期URLにアクセス: {settings.INITIAL_URL}")
                self.driver.get(settings.INITIAL_URL)
                time.sleep(2)  # ページの読み込みを待機
            
            self._initialized = True
            logger.info("WebDriverの初期化が正常に完了しました")
            return True
        except WebDriverException as e:
            logger.error(f"ドライバーの初期化に失敗しました: {str(e)}")
            self._initialized = False
            return False
    
    def close_driver(self):
        """WebDriverを安全に閉じる"""
        try:
            if getattr(self, 'driver', None):
                self.driver.quit()
                self.driver = None
                self._initialized = False
        except Exception as e:
            logger.warning(f"ドライバーのクローズ中にエラーが発生しました: {str(e)}")
        
    def update_dynamics(self, kba, url, target=None):
        """指定されたKBA番号のナレッジベース記事を更新します
        
        Args:
            kba: KBA番号
            url: 更新対象のURL
            target: 対象区分（社内向け/社外向け/該当なし）
            
        Returns:
            bool: 更新が成功したかどうか
        """
        retry_count = 0
        while retry_count <= settings.RETRY_COUNT:
            try:
                # 初期化に失敗した場合はリトライ
                if not self.initialize_driver():
                    retry_count += 1
                    logger.warning(f"ドライバー初期化に失敗しました。リトライ {retry_count}/{settings.RETRY_COUNT}")
                    time.sleep(1)
                    continue
                
                # 更新処理の実行
                self._update(url, target)
                logger.info(f"KBA: {kba} : URL: {url} : 更新成功")
                return True

            except Exception as e:
                retry_count += 1
                error_details = traceback.format_exc()
                
                # リトライ回数を超えた場合はエラーログを記録して次の記事へ
                if retry_count > settings.RETRY_COUNT:
                    logger.error(f"KBA: {kba} : URL: {url} : 更新失敗 ({retry_count}回目): {str(e)}")
                    logger.debug(f"エラー詳細: {error_details}")
                    return False
                
                # リトライ
                logger.warning(f"KBA: {kba} : URL: {url} : エラー発生、リトライします ({retry_count}/{settings.RETRY_COUNT}): {str(e)}")
                time.sleep(1)
        
        # 最終的に失敗した場合
        return False
    
    def _update(self, url, target=None):
        """記事の実際の更新処理を行います
        
        Args:
            url: 更新対象のURL
            target: 対象区分（社内向け/社外向け/該当なし）
            
        Raises:
            各種例外: 処理中に発生した例外
        """
        try:
            # URLにアクセス
            self.driver.get(url)
            time.sleep(settings.PAGE_LOAD_WAIT)

            # 公開の取り下げ
            try:
                unpublish_button = self.driver.find_element(By.ID, 'kbarticle|NoRelationship|Form|Mscrm.Form.kbarticle.Unpublish-Medium')
                
                # ボタンがアクティブかどうかを確認
                if unpublish_button.is_enabled() and unpublish_button.get_attribute('disabled') != 'true':
                    unpublish_button.click()
                    logger.debug("公開取り下げボタンをクリックしました")
                else:
                    logger.info("公開取り下げボタンが非アクティブのため、クリックをスキップしました")
            except NoSuchElementException:
                logger.warning("公開取り下げボタンが見つかりませんでした")

            # iframe切り替え
            self.driver.switch_to.frame(self.driver.find_element(By.ID, 'contentIFrame'))
            logger.debug("iframeに切り替えました")
            
            # タイトル更新
            try:
                title_el = self.driver.find_element(By.ID, 'title')
                current_value = title_el.get_attribute('value')
                
                if current_value and not current_value.startswith('【メンテ済】'):
                    new_value = f"【メンテ済】{current_value}"
                    title_el.clear()
                    title_el.send_keys(new_value)
                    logger.debug(f"タイトルを更新しました: {new_value}")
            except NoSuchElementException:
                logger.warning("タイトル要素が見つかりませんでした")
            
            # 社外/社内区分の更新
            if target is not None:
                try:
                    target_el = self.driver.find_element(By.ID, 'mjs_target')
                    select = Select(target_el)

                    if target == "社内向け":
                        select.select_by_value("1")
                    elif target == "社外向け":
                        select.select_by_value("2")
                    elif target == "該当なし":
                        select.select_by_value("3")
                    else:
                        logger.warning(f"不正な対象区分です: {target}")
                        raise ValueError(f"不正な対象区分です: {target}")
                    
                    logger.debug(f"対象区分を更新しました: {target}")
                except NoSuchElementException:
                    logger.warning("対象区分要素が見つかりませんでした")

            # iframe切り替え
            self.driver.switch_to.default_content()
            logger.debug("デフォルトコンテンツに戻りました")

            # 承認（リトライロジック付き）
            publish_retry_count = 0
            max_publish_retries = 3
            
            while publish_retry_count < max_publish_retries:
                try:
                    publish_button = self.driver.find_element(By.ID, 'kbarticle|NoRelationship|Form|Mscrm.Form.kbarticle.Publish-Medium')
                    
                    # ボタンがアクティブかどうかを確認
                    if publish_button.is_enabled() and publish_button.get_attribute('disabled') != 'true':
                        publish_button.click()
                        logger.debug("承認ボタンをクリックしました")
                        break  # 成功したらループを抜ける
                    else:
                        publish_retry_count += 1
                        if publish_retry_count < max_publish_retries:
                            logger.info(f"承認ボタンが非アクティブです。2秒待機してリトライします ({publish_retry_count}/{max_publish_retries})")
                            time.sleep(2)
                        else:
                            logger.error("承認ボタンが非アクティブのため、最大リトライ回数に達しました。更新処理を失敗とします")
                            raise PublishButtonNotActiveError("承認ボタンが非アクティブで最大リトライ回数に達しました")
                            
                except NoSuchElementException:
                    logger.error("承認ボタンが見つかりませんでした")
                    raise
            
        except NoSuchElementException as e:
            logger.error(f"要素が見つかりません: {str(e)}")
            raise
        except TimeoutException as e:
            logger.error(f"ページ読み込みがタイムアウトしました: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"更新処理中に予期せぬエラーが発生しました: {str(e)}")
            raise
