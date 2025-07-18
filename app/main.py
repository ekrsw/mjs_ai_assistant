import sys
import logging
from app.controller import update_dynamics
from app.config import settings, setup_logging

# ロガーの設定
logger = setup_logging()

def main():
    """
    メインエントリーポイント
    エクセルファイルからDynamics CRMの記事を更新します
    """
    logger.info("アプリケーションを開始します")
    
    try:
        # 更新処理の実行
        result = update_dynamics()
        
        # 処理結果に応じた終了コードの設定
        if result:
            logger.info("アプリケーションは正常に終了しました")
            return 0
        else:
            logger.warning("アプリケーションはエラーで終了しました")
            return 1
            
    except Exception as e:
        # 予期せぬエラーの処理
        logger.error(f"予期せぬエラーが発生しました: {str(e)}", exc_info=True)
        print(f"エラーが発生しました: {str(e)}")
        return 1

if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
