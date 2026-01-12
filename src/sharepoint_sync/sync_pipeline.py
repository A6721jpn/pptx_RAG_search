"""
SharePoint同期とバッチ処理パイプライン
大規模展開のためのオーケストレーター
"""

import asyncio
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime
import logging
import hashlib
import yaml
from dataclasses import dataclass

from tenacity import retry, stop_after_attempt, wait_exponential

from sharepoint_client import SharePointClient
import sys
sys.path.append(str(Path(__file__).parent.parent))
from utils.db_manager import ProcessedFilesDB

logger = logging.getLogger(__name__)


@dataclass
class SyncConfig:
    """同期設定"""
    tenant_id: str
    client_id: str
    client_secret: str
    site_urls: List[str]
    batch_size: int = 50
    parallel_downloads: int = 10
    retry_attempts: int = 3
    temp_dir: Path = Path("data/pptx_temp")
    db_path: Path = Path("data/processed_files.db")


class SharePointSyncPipeline:
    """SharePoint同期とバッチ処理パイプライン"""

    def __init__(self, config: SyncConfig):
        """
        Args:
            config: 同期設定
        """
        self.config = config
        self.db = ProcessedFilesDB(config.db_path)
        self.client: Optional[SharePointClient] = None

        # 一時ディレクトリ作成
        self.config.temp_dir.mkdir(parents=True, exist_ok=True)

    async def initialize(self):
        """パイプラインの初期化"""
        self.client = SharePointClient(
            tenant_id=self.config.tenant_id,
            client_id=self.config.client_id,
            client_secret=self.config.client_secret
        )
        logger.info("SharePointクライアント初期化完了")

    async def discover_files(self, incremental: bool = True) -> List[Dict]:
        """
        SharePointからPPTXファイルを検出

        Args:
            incremental: True の場合、変更されたファイルのみ処理

        Returns:
            処理対象ファイルのリスト
        """
        all_files = []

        for site_url in self.config.site_urls:
            logger.info(f"サイトをスキャン中: {site_url}")

            try:
                # サイトIDとドライブIDを取得
                site_id = await self.client.get_site_id(site_url)
                drive_id = await self.client.get_drive_id(site_id)

                # PPTXファイルを取得
                files = await self.client.list_pptx_files(
                    site_id=site_id,
                    drive_id=drive_id
                )

                all_files.extend(files)
                logger.info(f"検出: {len(files)}件 ({site_url})")

            except Exception as e:
                logger.error(f"サイトスキャン失敗 ({site_url}): {e}")
                continue

        logger.info(f"総検出ファイル数: {len(all_files)}")

        # データベースに登録し、処理が必要なファイルをフィルタ
        files_to_process = []
        for file_info in all_files:
            needs_processing = self.db.add_or_update_file(file_info)

            if needs_processing or not incremental:
                files_to_process.append(file_info)

        logger.info(f"処理対象ファイル数: {len(files_to_process)}")
        return files_to_process

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=60),
        reraise=True
    )
    async def download_file_with_retry(self, file_info: Dict) -> Path:
        """
        ファイルをダウンロード（リトライ付き）

        Args:
            file_info: ファイル情報

        Returns:
            ダウンロードしたファイルのパス
        """
        local_path = self.config.temp_dir / f"{file_info['id']}.pptx"

        try:
            await self.client.download_file(
                file_info['download_url'],
                local_path
            )
            self.db.add_log(file_info['id'], 'download', f"ダウンロード成功: {local_path}")
            return local_path

        except Exception as e:
            logger.error(f"ダウンロードエラー ({file_info['name']}): {e}")
            self.db.add_log(file_info['id'], 'download', f"ダウンロード失敗: {e}")
            raise

    async def download_batch(self, files: List[Dict]) -> List[Dict]:
        """
        ファイルを並列ダウンロード

        Args:
            files: ファイル情報のリスト

        Returns:
            ダウンロード成功したファイルの情報リスト
        """
        semaphore = asyncio.Semaphore(self.config.parallel_downloads)

        async def download_with_semaphore(file_info):
            async with semaphore:
                try:
                    local_path = await self.download_file_with_retry(file_info)
                    return {**file_info, 'local_path': local_path}
                except Exception as e:
                    logger.error(f"ダウンロード最終失敗 ({file_info['name']}): {e}")
                    self.db.update_status(file_info['id'], 'failed', str(e))
                    return None

        # 並列ダウンロード実行
        download_tasks = [download_with_semaphore(f) for f in files]
        results = await asyncio.gather(*download_tasks)

        # 成功したファイルのみ返す
        successful = [r for r in results if r is not None]
        logger.info(f"ダウンロード完了: {len(successful)}/{len(files)}")

        return successful

    def compute_doc_id(self, file_path: Path) -> str:
        """
        ファイル内容からドキュメントIDを生成

        Args:
            file_path: ファイルパス

        Returns:
            ドキュメントID (SHA1ハッシュの最初の12文字)
        """
        sha1 = hashlib.sha1()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                sha1.update(chunk)

        return sha1.hexdigest()[:12]

    def process_single_file(self, file_info: Dict) -> Dict:
        """
        単一ファイルを処理（テキスト抽出、レンダリング、埋め込み、インデックス）

        Args:
            file_info: ファイル情報（local_pathを含む）

        Returns:
            処理結果
        """
        file_id = file_info['id']
        local_path = file_info['local_path']

        start_time = datetime.now()

        try:
            self.db.update_status(file_id, 'processing')

            # ドキュメントID生成
            doc_id = self.compute_doc_id(local_path)
            logger.info(f"処理開始: {file_info['name']} (doc_id: {doc_id})")

            # ========== ここに実際の処理を実装 ==========
            #
            # 1. テキスト抽出 (python-pptx)
            #    from ingest.pptx_extract import extract_text
            #    extracted = extract_text(local_path)
            #
            # 2. スライドレンダリング (PowerPoint COM)
            #    from ingest.pptx_render_com import render_slides
            #    rendered_dir = render_slides(local_path, doc_id)
            #
            # 3. 埋め込み計算
            #    from ingest.build_index import compute_embeddings
            #    embeddings = compute_embeddings(extracted, rendered_dir)
            #
            # 4. Qdrantインデックス更新
            #    from ingest.build_index import update_index
            #    update_index(doc_id, embeddings)
            #
            # ===========================================

            # 仮の処理（実装時は上記に置き換え）
            slide_count = 10  # 実際の処理結果に置き換え

            # 処理完了
            duration = (datetime.now() - start_time).total_seconds()
            self.db.update_status(
                file_id,
                'success',
                doc_id=doc_id,
                slide_count=slide_count,
                duration=duration
            )

            logger.info(f"処理完了: {file_info['name']} ({duration:.2f}秒)")

            return {
                'file_id': file_id,
                'doc_id': doc_id,
                'slide_count': slide_count,
                'duration': duration,
                'status': 'success'
            }

        except Exception as e:
            duration = (datetime.now() - start_time).total_seconds()
            logger.error(f"処理エラー ({file_info['name']}): {e}")
            self.db.update_status(file_id, 'failed', str(e), duration=duration)

            return {
                'file_id': file_id,
                'status': 'failed',
                'error': str(e)
            }

        finally:
            # 一時ファイルを削除
            if local_path.exists():
                local_path.unlink()

    async def process_batch(self, files: List[Dict]) -> List[Dict]:
        """
        ファイルのバッチ処理

        Args:
            files: ダウンロード済みファイル情報のリスト

        Returns:
            処理結果のリスト
        """
        # COM制約により、レンダリングは順次処理
        # テキスト抽出と埋め込みは並列化可能（実装次第）

        results = []
        for i, file_info in enumerate(files, 1):
            logger.info(f"処理中: {i}/{len(files)} - {file_info['name']}")
            result = self.process_single_file(file_info)
            results.append(result)

        return results

    async def run(self, incremental: bool = True) -> Dict:
        """
        パイプライン全体を実行

        Args:
            incremental: 増分処理モード

        Returns:
            処理結果の統計
        """
        pipeline_start = datetime.now()

        try:
            # 初期化
            await self.initialize()

            # ファイル検出
            logger.info("=== ファイル検出フェーズ ===")
            files_to_process = await self.discover_files(incremental=incremental)

            if not files_to_process:
                logger.info("処理対象ファイルなし")
                return {
                    'status': 'success',
                    'files_processed': 0,
                    'duration_seconds': 0
                }

            # バッチ処理
            total_processed = 0
            total_failed = 0
            batch_size = self.config.batch_size

            for batch_idx in range(0, len(files_to_process), batch_size):
                batch = files_to_process[batch_idx:batch_idx + batch_size]
                batch_num = (batch_idx // batch_size) + 1
                total_batches = (len(files_to_process) + batch_size - 1) // batch_size

                logger.info(f"\n=== バッチ {batch_num}/{total_batches} ({len(batch)}ファイル) ===")

                # ダウンロード
                logger.info("ダウンロードフェーズ")
                downloaded = await self.download_batch(batch)

                # 処理
                logger.info("処理フェーズ")
                results = await self.process_batch(downloaded)

                # 統計更新
                for result in results:
                    if result['status'] == 'success':
                        total_processed += 1
                    else:
                        total_failed += 1

                logger.info(f"バッチ完了: 成功={total_processed}, 失敗={total_failed}")

            # パイプライン完了
            duration = (datetime.now() - pipeline_start).total_seconds()

            # 最終統計
            stats = self.db.get_statistics()

            logger.info(f"\n=== パイプライン完了 ===")
            logger.info(f"処理時間: {duration:.2f}秒")
            logger.info(f"成功: {total_processed}")
            logger.info(f"失敗: {total_failed}")
            logger.info(f"総スライド数: {stats.get('total_slides', 0)}")

            return {
                'status': 'success',
                'files_processed': total_processed,
                'files_failed': total_failed,
                'duration_seconds': duration,
                'statistics': stats
            }

        except Exception as e:
            logger.error(f"パイプラインエラー: {e}")
            return {
                'status': 'failed',
                'error': str(e)
            }

        finally:
            # クリーンアップ
            if self.client:
                await self.client.close()
            self.db.close()


# ========== CLI エントリーポイント ==========

async def main_cli():
    """CLIエントリーポイント"""
    import argparse

    parser = argparse.ArgumentParser(description="SharePoint PPTX同期パイプライン")
    parser.add_argument(
        '--config',
        type=Path,
        default=Path('configs/sharepoint_prod.yaml'),
        help='設定ファイルパス'
    )
    parser.add_argument(
        '--incremental',
        action='store_true',
        help='増分処理モード（変更されたファイルのみ）'
    )
    parser.add_argument(
        '--full',
        action='store_true',
        help='フルスキャンモード（すべてのファイルを再処理）'
    )

    args = parser.parse_args()

    # 設定ファイル読み込み
    with open(args.config) as f:
        config_dict = yaml.safe_load(f)

    # 設定オブジェクト作成
    config = SyncConfig(
        tenant_id=config_dict['sharepoint']['tenant_id'],
        client_id=config_dict['sharepoint']['client_id'],
        client_secret=config_dict['sharepoint']['client_secret'],
        site_urls=config_dict['sharepoint']['site_urls'],
        batch_size=config_dict.get('processing', {}).get('batch_size', 50),
        parallel_downloads=config_dict.get('processing', {}).get('parallel_downloads', 10)
    )

    # パイプライン実行
    pipeline = SharePointSyncPipeline(config)
    result = await pipeline.run(incremental=not args.full)

    # 結果出力
    print("\n" + "="*50)
    print("処理結果:")
    print(f"  ステータス: {result['status']}")
    if result['status'] == 'success':
        print(f"  処理ファイル数: {result['files_processed']}")
        print(f"  失敗ファイル数: {result['files_failed']}")
        print(f"  処理時間: {result['duration_seconds']:.2f}秒")
    else:
        print(f"  エラー: {result.get('error')}")
    print("="*50)


if __name__ == "__main__":
    # ロギング設定
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('data/logs/sync_pipeline.log'),
            logging.StreamHandler()
        ]
    )

    # 実行
    asyncio.run(main_cli())
