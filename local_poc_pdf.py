"""
PDF版ローカルPOC
Power AutomateでPDF化されたファイルをOneDrive同期フォルダから処理
PowerPoint COM不要で軽量・高速
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
        logging.FileHandler('data/logs/local_poc_pdf.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class PDFScanner:
    """PDFフォルダからファイルをスキャン"""

    def __init__(self, source_dir: Path, db_path: Path):
        """
        Args:
            source_dir: PDF同期フォルダのパス
            db_path: 処理状態データベースのパス
        """
        self.source_dir = source_dir
        self.db = ProcessedFilesDB(db_path)

    def scan_pdf_files(self) -> list:
        """
        PDFファイルを再帰的にスキャン

        Returns:
            ファイル情報のリスト
        """
        logger.info(f"スキャン開始: {self.source_dir}")

        pdf_files = []

        # .pdf ファイルを再帰的に検索
        for file_path in self.source_dir.glob('**/*.pdf'):
            # 一時ファイルをスキップ
            if file_path.name.startswith('.'):
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

            pdf_files.append(file_info)

        logger.info(f"検出ファイル数: {len(pdf_files)}")
        return pdf_files

    def compute_file_id(self, file_path: Path) -> str:
        """ファイルパスからIDを生成"""
        normalized_path = str(file_path.resolve())
        return hashlib.sha1(normalized_path.encode()).hexdigest()[:12]

    def compute_doc_id(self, file_path: Path) -> str:
        """ファイル内容からドキュメントIDを生成"""
        sha1 = hashlib.sha1()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                sha1.update(chunk)
        return sha1.hexdigest()[:12]

    def extract_text_from_pdf(self, pdf_path: Path) -> list:
        """
        PDFからテキストを抽出

        Args:
            pdf_path: PDFファイルパス

        Returns:
            ページごとのテキストリスト
        """
        try:
            import pdfplumber
        except ImportError:
            logger.error("pdfplumberがインストールされていません: pip install pdfplumber")
            raise

        pages_text = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                pages_text.append({
                    'page_num': page_num,
                    'text': text.strip()
                })
                logger.debug(f"  ページ{page_num}: {len(text)}文字")

        logger.info(f"テキスト抽出完了: {len(pages_text)}ページ")
        return pages_text

    def render_pdf_pages(self, pdf_path: Path, output_dir: Path) -> list:
        """
        PDFページを画像にレンダリング

        Args:
            pdf_path: PDFファイルパス
            output_dir: 出力ディレクトリ

        Returns:
            生成された画像ファイルパスのリスト
        """
        try:
            from pdf2image import convert_from_path
        except ImportError:
            logger.error("pdf2imageがインストールされていません: pip install pdf2image")
            logger.error("また、popplerが必要です（Windows: https://github.com/oschwartz10612/poppler-windows/releases/）")
            raise

        output_dir.mkdir(parents=True, exist_ok=True)

        # PDFを画像に変換
        images = convert_from_path(
            pdf_path,
            dpi=150,  # 解像度（高いほど品質向上、処理時間増）
            fmt='png'
        )

        image_paths = []

        for i, image in enumerate(images, 1):
            image_path = output_dir / f"page_{i:04d}.png"
            image.save(image_path, 'PNG')
            image_paths.append(image_path)
            logger.debug(f"  ページ{i}レンダリング完了: {image_path}")

        logger.info(f"レンダリング完了: {len(image_paths)}ページ")
        return image_paths

    def get_files_to_process(self, incremental: bool = True) -> list:
        """処理が必要なファイルを取得"""
        all_files = self.scan_pdf_files()

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
        単一PDFファイルを処理

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

            # 1. テキスト抽出
            logger.info("  [1/4] テキスト抽出中...")
            pages_text = self.extract_text_from_pdf(file_path)

            # 2. ページ画像のレンダリング
            logger.info("  [2/4] ページレンダリング中...")
            rendered_dir = Path('data/rendered') / doc_id
            image_paths = self.render_pdf_pages(file_path, rendered_dir)

            page_count = len(pages_text)

            # ========== ここに実際の埋め込み・インデックス処理を実装 ==========
            #
            # 3. 埋め込み計算
            #    logger.info("  [3/4] 埋め込み計算中...")
            #    from ingest.build_index import compute_embeddings
            #    embeddings = compute_embeddings(pages_text, image_paths)
            #
            # 4. Qdrantインデックス更新
            #    logger.info("  [4/4] インデックス更新中...")
            #    from ingest.build_index import update_index
            #    update_index(doc_id, embeddings)
            #
            # ===============================================================

            logger.info("  [3/4] 埋め込み計算中... (スキップ - 未実装)")
            logger.info("  [4/4] インデックス更新中... (スキップ - 未実装)")

            # 処理完了
            duration = (datetime.now() - start_time).total_seconds()
            self.db.update_status(
                file_id,
                'success',
                doc_id=doc_id,
                slide_count=page_count,
                duration=duration
            )

            logger.info(f"処理完了: {file_info['name']} ({duration:.2f}秒, {page_count}ページ)")

            return {
                'file_id': file_id,
                'doc_id': doc_id,
                'page_count': page_count,
                'duration': duration,
                'status': 'success'
            }

        except Exception as e:
            duration = (datetime.now() - start_time).total_seconds()
            logger.error(f"処理エラー ({file_info['name']}): {e}")
            import traceback
            traceback.print_exc()
            self.db.update_status(file_id, 'failed', str(e), duration=duration)

            return {
                'file_id': file_id,
                'status': 'failed',
                'error': str(e)
            }

    def run(self, incremental: bool = True) -> dict:
        """POC全体を実行"""
        pipeline_start = datetime.now()

        logger.info("=== PDF版ローカルPOC開始 ===")
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
        total_pages = 0

        for i, file_info in enumerate(files_to_process, 1):
            logger.info(f"\n処理中: {i}/{len(files_to_process)} - {file_info['name']}")
            result = self.process_file(file_info)

            if result['status'] == 'success':
                total_processed += 1
                total_pages += result.get('page_count', 0)
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
        logger.info(f"総ページ数: {total_pages}")
        logger.info(f"平均処理時間: {stats.get('avg_processing_seconds', 0):.2f}秒/ファイル")

        return {
            'status': 'success',
            'files_processed': total_processed,
            'files_failed': total_failed,
            'total_pages': total_pages,
            'duration_seconds': duration,
            'statistics': stats
        }


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(description="PDF同期フォルダからPOC実行")
    parser.add_argument(
        '--source',
        type=Path,
        required=True,
        help='PDF同期フォルダのパス（例: C:\\Users\\YourName\\OneDrive - Company\\Site - Documents\\PDF_Converted）'
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
        print("\nPDF同期フォルダのパスを確認してください。")
        print("例: C:\\Users\\YourName\\OneDrive - Company\\Site - Documents\\PDF_Converted")
        sys.exit(1)

    if not args.source.is_dir():
        print(f"❌ エラー: パスがディレクトリではありません: {args.source}")
        sys.exit(1)

    # 必要なライブラリの確認
    try:
        import pdfplumber
        import pdf2image
    except ImportError as e:
        print(f"❌ エラー: 必要なライブラリがインストールされていません")
        print("\n以下のコマンドでインストールしてください:")
        print("  pip install pdfplumber pdf2image")
        print("\nまた、pdf2imageにはpopplerが必要です:")
        print("  Windows: https://github.com/oschwartz10612/poppler-windows/releases/")
        print("  ダウンロード後、PATHに追加してください")
        sys.exit(1)

    # データディレクトリ作成
    Path('data/logs').mkdir(parents=True, exist_ok=True)
    Path('data/rendered').mkdir(parents=True, exist_ok=True)

    # スキャナー作成
    scanner = PDFScanner(
        source_dir=args.source,
        db_path=Path('data/processed_files_pdf.db')
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
            print(f"  総ページ数: {result['total_pages']}")
            print(f"  処理時間: {result['duration_seconds']:.2f}秒")

            if result['files_processed'] > 0:
                avg_time = result['duration_seconds'] / result['files_processed']
                print(f"  平均処理時間: {avg_time:.2f}秒/ファイル")
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
