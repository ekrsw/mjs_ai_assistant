import os
import pandas as pd
import logging
import traceback
from app.dynamics import Dynamics

from app.config import settings

logger = logging.getLogger(__name__)

def update_dynamics():
    """
    エクセルファイルからデータを読み取り、Dynamics CRMの記事を更新します。
    エラーが発生しても可能な限り処理を継続します。
    """
    logger.info("更新処理を開始します")
    
    # ファイルの存在確認
    if not os.path.exists(settings.FILE):
        error_msg = f'ファイルが見つかりません: {settings.FILE}'
        logger.error(error_msg)
        print(error_msg)
        return False
    
    try:
        # エクセルファイルの読み込み
        logger.info(f"エクセルファイルを読み込みます: {settings.FILE}")
        try:
            df = pd.read_excel(settings.FILE)
            logger.info(f"エクセルファイルの読み込みに成功しました。レコード数: {len(df)}")
        except Exception as e:
            logger.error(f"エクセルファイルの読み込みに失敗しました: {str(e)}")
            print(f"エクセルファイルの読み込みに失敗しました: {str(e)}")
            return False
        
        # Dynamicsインスタンスの作成
        dynamics = Dynamics()
        
        # 成功・失敗カウンター
        success_count = 0
        failure_count = 0
        
        # 各行の処理
        total_rows = len(df)
        for index, row in df.iterrows():
            try:
                # 必須項目の確認
                if '記事' not in row or pd.isna(row['記事']) or not row['記事']:
                    logger.warning(f"行 {index+1}: 主キーが指定されていません。スキップします。")
                    failure_count += 1
                    continue
                
                # データの取得
                kba = row.get('番号', f"不明_{index+1}")
                url = create_url(row['記事'])
                target = row.get('対象')
                
                logger.info(f"処理中 ({index+1}/{total_rows}): KBA: {kba}")
                
                # 更新処理の実行
                result = dynamics.update_dynamics(kba, url, target)
                
                # 結果の記録
                if result:
                    success_count += 1
                    print(f'記事番号: {kba} 更新成功 ({index+1}/{total_rows})')
                else:
                    failure_count += 1
                    print(f'記事番号: {kba} 更新失敗 ({index+1}/{total_rows})')
                
            except Exception as e:
                # 個別の行の処理中にエラーが発生しても続行
                error_details = traceback.format_exc()
                logger.error(f"行 {index+1} の処理中にエラーが発生しました: {str(e)}")
                logger.debug(f"エラー詳細: {error_details}")
                failure_count += 1
                print(f'記事番号: {kba if "kba" in locals() else f"不明_{index+1}"} 処理エラー: {str(e)}')
                continue
        
        # 処理結果のサマリーを出力
        summary = f"処理完了: 合計 {total_rows}件、成功 {success_count}件、失敗 {failure_count}件"
        logger.info(summary)
        print(summary)
        return True
        
    except Exception as e:
        # 予期せぬエラーの処理
        error_details = traceback.format_exc()
        logger.error(f"更新処理中に予期せぬエラーが発生しました: {str(e)}")
        logger.debug(f"エラー詳細: {error_details}")
        print(f"エラーが発生しました: {str(e)}")
        return False

def create_url(article_str):
    """
    記事IDからURLを生成します
    
    Args:
        article_str: 記事ID
        
    Returns:
        str: 生成されたURL
    """
    return f'http://sv-vw-ejap:5555/SupportCenter/main.aspx?etc=127&extraqs=%3fetc%3d127%26id%3d%257b{article_str}%257d&newWindow=true&pagetype=entityrecord'
