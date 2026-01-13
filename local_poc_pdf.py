"""
PDF版ローカルPOC
PDFからテキスト抽出、埋め込み計算、Qdrantインデックス化まで一貫処理
"""

from pathlib import Path
from datetime import datetime
import hashlib
import logging
import sys

# パスを追加
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from utils.db_manager import ProcessedFilesDB
from ingest.embeddings import TextEmbedder
from ingest.indexer import QdrantIndexer

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


class PDFPOCProcessor:
    """PDF処理からインデックス化までの統合プロセッサ"""

    def __init__(self, source_dir: Path, db_path: Path, qdrant_path: str):
        """
        Args:
            source_dir: PDFフォルダのパス
            db_path: 処理状態データベースのパス
            qdrant_path: Qdrantストレージパス
        """
        self.source_dir = source_dir
        self.db = ProcessedFilesDB(db_path)

        # 埋め込みとインデックス初期化
        self.embedder = TextEmbedder(model_name="intfloat/e5-base-v2", device="cpu")
        self.indexer = QdrantIndexer(storage_path=qdrant_path)

        # 初回のみQdrant初期化
        self._qdrant_initialized = False

    def _ensure_qdrant_initialized(self):
        """Qdrant初期化（初回のみ）"""
        if not self._qdrant_initialized:
            vector_dim = self.embedder.get_dimension()
            self.indexer.initialize(vector_dimension=vector_dim)
            self._qdrant_initialized = True

    def scan_pdf_files(self) -> list:
        """PDFファイルをスキャン"""
        logger.info(f"スキャン開始: {self.source_dir}")

        pdf_files = []

        for file_path in self.source_dir.glob('**/*.pdf'):
            if file_path.name.startswith('.'):
                continue

            stat = file_path.stat()
            file_info = {
                'id': self._compute_file_id(file_path),
                'name': file_path.name,
                'path': str(file_path.relative_to(self.source_dir)),
                'full_path': str(file_path),
                'modified': datetime.fromtimestamp(stat.st_mtime),
                'size': stat.st_size
            }

            pdf_files.append(file_info)

        logger.info(f"検出ファイル数: {len(pdf_files)}")
        return pdf_files

    def _compute_file_id(self, file_path: Path) -> str:
        """ファイルパスからIDを生成"""
        return hashlib.sha1(str(file_path.resolve()).encode()).hexdigest()[:12]

    def _compute_doc_id(self, file_path: Path) -> str:
        """ファイル内容からドキュメントIDを生成"""
        sha1 = hashlib.sha1()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                sha1.update(chunk)
        return sha1.hexdigest()[:12]

    def extract_text_from_pdf(self, pdf_path: Path) -> list:
        """PDFからテキストを抽出"""
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

        logger.info(f"テキスト抽出完了: {len(pages_text)}ページ")
        return pages_text

    def render_pdf_pages(self, pdf_path: Path, output_dir: Path) -> list:
        """PDFページを画像にレンダリング (PyMuPDF使用)"""
        try:
            import fitz
        except ImportError:
            logger.error("pymupdfがインストールされていません: pip install pymupdf")
            raise

        output_dir.mkdir(parents=True, exist_ok=True)
        
        doc = fitz.open(pdf_path)
        image_paths = []
        
        # 150 DPI相当 (72 * 2.08)
        zoom = 150 / 72
        mat = fitz.Matrix(zoom, zoom)

        for i, page in enumerate(doc, 1):
            image_path = output_dir / f"page_{i:04d}.png"
            pix = page.get_pixmap(matrix=mat)
            pix.save(str(image_path))
            image_paths.append(image_path)

        logger.info(f"レンダリング完了: {len(image_paths)}ページ")
        return image_paths

    def get_files_to_process(self, incremental: bool = True) -> list:
        """処理が必要なファイルを取得"""
        all_files = self.scan_pdf_files()

        if not incremental:
            logger.info("フルスキャンモード: すべてのファイルを処理")
            return all_files

        files_to_process = []
        for file_info in all_files:
            if self.db.add_or_update_file(file_info):
                files_to_process.append(file_info)

        logger.info(f"処理対象: {len(files_to_process)} / {len(all_files)}")
        return files_to_process

    def process_file(self, file_info: dict) -> dict:
        """単一PDFファイルを処理"""
        file_id = file_info['id']
        file_path = Path(file_info['full_path'])

        start_time = datetime.now()

        try:
            self.db.update_status(file_id, 'processing')

            doc_id = self._compute_doc_id(file_path)
            logger.info(f"処理開始: {file_info['name']} (doc_id: {doc_id})")

            # 1. テキスト抽出
            logger.info("  [1/4] テキスト抽出中...")
            pages = self.extract_text_from_pdf(file_path)

            # 2. ページレンダリング
            logger.info("  [2/4] ページレンダリング中...")
            rendered_dir = Path('data/rendered') / doc_id
            image_paths = self.render_pdf_pages(file_path, rendered_dir)

            # ページ情報に画像パスを追加
            for page_info, img_path in zip(pages, image_paths):
                page_info['image_path'] = img_path

            # 3. 埋め込み計算
            logger.info("  [3/4] 埋め込み計算中...")
            texts = [p['text'] for p in pages]
            embeddings = self.embedder.embed_texts(texts, batch_size=16)

            # 4. Qdrantインデックス更新
            logger.info("  [4/4] インデックス更新中...")
            self._ensure_qdrant_initialized()
            self.indexer.index_pages(
                doc_id=doc_id,
                file_name=file_info['name'],
                pages=pages,
                embeddings=embeddings
            )

            # 処理完了
            duration = (datetime.now() - start_time).total_seconds()
            self.db.update_status(
                file_id,
                'success',
                doc_id=doc_id,
                slide_count=len(pages),
                duration=duration
            )

            logger.info(f"処理完了: {file_info['name']} ({duration:.2f}秒, {len(pages)}ページ)")

            return {
                'file_id': file_id,
                'doc_id': doc_id,
                'page_count': len(pages),
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

        # 完了
        duration = (datetime.now() - pipeline_start).total_seconds()
        # stats = self.db.get_statistics()

        # Qdrant情報
        qdrant_info = {}
        if self._qdrant_initialized:
            try:
                qdrant_info = self.indexer.get_collection_info()
            except Exception as e:
                logger.warning(f"Qdrant情報取得失敗: {e}")

        logger.info(f"\n=== POC完了 ===")
        logger.info(f"処理時間: {duration:.2f}秒")
        logger.info(f"成功: {total_processed}")
        logger.info(f"失敗: {total_failed}")
        logger.info(f"総ページ数: {total_pages}")
        if qdrant_info:
            logger.info(f"Qdrantポイント数: {qdrant_info.get('points_count', 'N/A')}")
        else:
            logger.info("Qdrantポイント数: N/A")

        return {
            'status': 'success',
            'files_processed': total_processed,
            'files_failed': total_failed,
            'total_pages': total_pages,
            'duration_seconds': duration,
            # 'statistics': stats,
            'qdrant_info': qdrant_info
        }


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(description="PDF版ローカルPOC")
    parser.add_argument(
        '--source',
        type=Path,
        required=True,
        help='PDFフォルダのパス'
    )
    parser.add_argument(
        '--incremental',
        action='store_true',
        help='増分処理モード'
    )
    parser.add_argument(
        '--full',
        action='store_true',
        help='フルスキャンモード'
    )
    parser.add_argument(
        '--qdrant',
        type=str,
        default='index/qdrant_storage',
        help='Qdrantストレージパス'
    )

    args = parser.parse_args()

    # ソースディレクトリ確認
    if not args.source.exists():
        print(f"❌ エラー: ディレクトリが見つかりません: {args.source}")
        sys.exit(1)

    # データディレクトリ作成
    Path('data/logs').mkdir(parents=True, exist_ok=True)
    Path('data/rendered').mkdir(parents=True, exist_ok=True)

    # プロセッサ作成
    processor = PDFPOCProcessor(
        source_dir=args.source,
        db_path=Path('data/processed_files_pdf.db'),
        qdrant_path=args.qdrant
    )

    # POC実行
    try:
        result = processor.run(incremental=not args.full)

        # 結果出力
        print("\n" + "="*50)
        print("処理結果:")
        print(f"  ステータス: {result['status']}")
        print(f"  処理ファイル数: {result['files_processed']}")
        print(f"  失敗ファイル数: {result['files_failed']}")
        print(f"  総ページ数: {result['total_pages']}")
        print(f"  処理時間: {result['duration_seconds']:.2f}秒")

        if result.get('qdrant_info'):
            print(f"  Qdrantポイント数: {result['qdrant_info']['points_count']}")

        print("="*50)

        sys.exit(0 if result['files_failed'] == 0 else 1)

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
