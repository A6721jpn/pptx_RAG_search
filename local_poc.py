"""
OneDrive同期フォルダを使ったローカルPOC
Azure ADアプリ登録不要でSharePointファイルを処理
"""

import asyncio
from pathlib import Path
from datetime import datetime
import hashlib
import logging
import sys

# パスを追加
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from utils.db_manager import ProcessedFilesDB

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data/logs/local_poc.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class LocalPPTXScanner:
    """ローカルフォルダからPPTXファイルをスキャン"""

    def __init__(self, source_dir: Path, db_path: Path):
        """
        Args:
            source_dir: OneDrive同期フォルダのパス
            db_path: 処理状態データベースのパス
        """
        self.source_dir = source_dir
        self.db = ProcessedFilesDB(db_path)

    def scan_pptx_files(self) -> list:
        """
        PPTXファイルを再帰的にスキャン

        Returns:
            ファイル情報のリスト
        """
        logger.info(f"スキャン開始: {self.source_dir}")

        pptx_files = []

        # .pptx と .ppt ファイルを再帰的に検索
        for pattern in ['**/*.pptx', '**/*.ppt']:
            for file_path in self.source_dir.glob(pattern):
                # 一時ファイルやロックファイルをスキップ
                if file_path.name.startswith('~$'):
                    continue

                # ファイル情報を取得
                stat = file_path.stat()

                file_info = {
                    'id': self.compute_file_id(file_path),
                    'name': file_path.name,
                    'path': str(file_path.relative_to(self.source_dir)),
                    'full_path': str(file_path),
                    'modified': datetime.fromtimestamp(stat.st_mtime),
                    'size': stat.st_size
                }

                pptx_files.append(file_info)

        logger.info(f"検出ファイル数: {len(pptx_files)}")
        return pptx_files

    def compute_file_id(self, file_path: Path) -> str:
        """
        ファイルパスからIDを生成

        Args:
            file_path: ファイルパス

        Returns:
            ファイルID
        """
        # パスの正規化されたハッシュを使用
        normalized_path = str(file_path.resolve())
        return hashlib.sha1(normalized_path.encode()).hexdigest()[:12]

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

    def get_files_to_process(self, incremental: bool = True) -> list:
        """
        処理が必要なファイルを取得

        Args:
            incremental: True の場合、変更されたファイルのみ

        Returns:
            処理対象ファイルのリスト
        """
        all_files = self.scan_pptx_files()

        if not incremental:
            logger.info("フルスキャンモード: すべてのファイルを処理対象にします")
            return all_files

        # 増分モード: 変更されたファイルのみ
        files_to_process = []
        for file_info in all_files:
            needs_processing = self.db.add_or_update_file(file_info)
            if needs_processing:
                files_to_process.append(file_info)

        logger.info(f"処理対象ファイル数: {len(files_to_process)} / {len(all_files)}")
        return files_to_process

    def process_file(self, file_info: dict) -> dict:
        """
        単一ファイルを処理

        Args:
            file_info: ファイル情報

        Returns:
            処理結果
        """
        file_id = file_info['id']
        file_path = Path(file_info['full_path'])

        start_time = datetime.now()

        try:
            self.db.update_status(file_id, 'processing')

            # ドキュメントID生成
            doc_id = self.compute_doc_id(file_path)
            logger.info(f"処理開始: {file_info['name']} (doc_id: {doc_id})")

            # ========== ここに実際の処理を実装 ==========
            #
            # 1. テキスト抽出 (python-pptx)
            #    from ingest.pptx_extract import extract_text
            #    extracted = extract_text(file_path)
            #
            # 2. スライドレンダリング (PowerPoint COM)
            #    from ingest.pptx_render_com import render_slides
            #    rendered_dir = render_slides(file_path, doc_id)
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

    def run(self, incremental: bool = True) -> dict:
        """
        POC全体を実行

        Args:
            incremental: 増分処理モード

        Returns:
            処理結果の統計
        """
        pipeline_start = datetime.now()

        logger.info("=== ローカルPOC開始 ===")
        logger.info(f"ソースディレクトリ: {self.source_dir}")

        # ファイル検出
        files_to_process = self.get_files_to_process(incremental=incremental)

        if not files_to_process:
            logger.info("処理対象ファイルなし")
            return {
                'status': 'success',
                'files_processed': 0,
                'duration_seconds': 0
            }

        # 処理実行
        total_processed = 0
        total_failed = 0

        for i, file_info in enumerate(files_to_process, 1):
            logger.info(f"\n処理中: {i}/{len(files_to_process)} - {file_info['name']}")
            result = self.process_file(file_info)

            if result['status'] == 'success':
                total_processed += 1
            else:
                total_failed += 1

        # パイプライン完了
        duration = (datetime.now() - pipeline_start).total_seconds()

        # 最終統計
        stats = self.db.get_statistics()

        logger.info(f"\n=== POC完了 ===")
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


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(description="OneDrive同期フォルダからPOC実行")
    parser.add_argument(
        '--source',
        type=Path,
        required=True,
        help='OneDrive同期フォルダのパス（例: C:\\Users\\YourName\\OneDrive - Company\\Site - Documents）'
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

    # ソースディレクトリの確認
    if not args.source.exists():
        print(f"❌ エラー: ディレクトリが見つかりません: {args.source}")
        print("\nOneDrive同期フォルダのパスを確認してください。")
        print("例: C:\\Users\\YourName\\OneDrive - Company\\Site - Documents")
        sys.exit(1)

    if not args.source.is_dir():
        print(f"❌ エラー: パスがディレクトリではありません: {args.source}")
        sys.exit(1)

    # データディレクトリ作成
    Path('data/logs').mkdir(parents=True, exist_ok=True)

    # スキャナー作成
    scanner = LocalPPTXScanner(
        source_dir=args.source,
        db_path=Path('data/processed_files.db')
    )

    # POC実行
    try:
        result = scanner.run(incremental=not args.full)

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

    except KeyboardInterrupt:
        print("\n\n処理が中断されました")
        sys.exit(1)
    except Exception as e:
        logger.error(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
