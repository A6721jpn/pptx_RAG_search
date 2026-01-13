"""
ローカルPPTX→PDF変換スクリプト
PowerPoint COMを使用してPPTXファイルをPDFに変換
起動時に増分ファイルのみを処理
"""

import sys
from pathlib import Path
from datetime import datetime
import hashlib
import logging
import yaml
import sqlite3

# Windows専用
try:
    import win32com.client
    import pythoncom
except ImportError:
    print("❌ エラー: pywin32がインストールされていません")
    print("以下のコマンドでインストールしてください:")
    print("  pip install pywin32")
    sys.exit(1)

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data/logs/pdf_conversion.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ConversionDB:
    """PDF変換履歴データベース"""

    def __init__(self, db_path: Path):
        """
        Args:
            db_path: データベースファイルのパス
        """
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = None
        self._initialize_db()

    def _initialize_db(self):
        """データベース初期化"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row

        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS conversions (
                file_id TEXT PRIMARY KEY,
                source_path TEXT NOT NULL,
                pdf_path TEXT NOT NULL,
                source_modified TIMESTAMP,
                converted_at TIMESTAMP,
                status TEXT DEFAULT 'success',
                error_message TEXT
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_source_path
            ON conversions(source_path)
        """)

        self.conn.commit()
        logger.info(f"データベース初期化完了: {self.db_path}")

    def needs_conversion(self, source_path: Path, source_modified: datetime) -> bool:
        """
        変換が必要か確認

        Args:
            source_path: ソースファイルパス
            source_modified: ソースファイルの更新日時

        Returns:
            True: 変換が必要, False: 変換不要
        """
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT source_modified, status FROM conversions WHERE source_path = ?",
            (str(source_path),)
        )
        row = cursor.fetchone()

        # 新規ファイル
        if row is None:
            return True

        # 失敗したファイルは再変換
        if row['status'] == 'failed':
            return True

        # 更新されたファイル
        existing_modified = datetime.fromisoformat(row['source_modified'])
        if source_modified > existing_modified:
            logger.info(f"ファイル更新検出: {source_path.name}")
            return True

        return False

    def record_conversion(
        self,
        file_id: str,
        source_path: Path,
        pdf_path: Path,
        source_modified: datetime,
        status: str = 'success',
        error_message: str = None
    ):
        """
        変換履歴を記録

        Args:
            file_id: ファイルID
            source_path: ソースファイルパス
            pdf_path: PDF出力パス
            source_modified: ソースファイルの更新日時
            status: 変換ステータス
            error_message: エラーメッセージ
        """
        cursor = self.conn.cursor()

        cursor.execute("""
            INSERT OR REPLACE INTO conversions
            (file_id, source_path, pdf_path, source_modified, converted_at, status, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            file_id,
            str(source_path),
            str(pdf_path),
            source_modified.isoformat(),
            datetime.now().isoformat(),
            status,
            error_message
        ))

        self.conn.commit()

    def get_statistics(self) -> dict:
        """変換統計を取得"""
        cursor = self.conn.cursor()

        stats = {}

        # 総変換数
        cursor.execute("SELECT COUNT(*) as total FROM conversions")
        stats['total'] = cursor.fetchone()['total']

        # ステータス別
        cursor.execute("""
            SELECT status, COUNT(*) as count
            FROM conversions
            GROUP BY status
        """)
        stats['by_status'] = {row['status']: row['count'] for row in cursor.fetchall()}

        # 最終変換日時
        cursor.execute("SELECT MAX(converted_at) as last FROM conversions")
        stats['last_conversion'] = cursor.fetchone()['last']

        return stats

    def close(self):
        """データベース接続をクローズ"""
        if self.conn:
            self.conn.close()


class PPTXToPDFConverter:
    """PowerPoint COMを使ったPDF変換"""

    def __init__(self, config: dict):
        """
        Args:
            config: 設定辞書
        """
        self.config = config
        self.source_folder = Path(config['source']['pptx_folder'])
        self.output_folder = Path(config['output']['pdf_folder'])
        self.recursive = config['source'].get('recursive', True)
        self.preserve_structure = config['output'].get('preserve_structure', True)

        # データベース
        db_path = Path(config['conversion']['db_path'])
        self.db = ConversionDB(db_path)

        # PowerPoint設定
        self.powerpoint = None

    def initialize_powerpoint(self):
        """PowerPointアプリケーションを初期化"""
        try:
            pythoncom.CoInitialize()
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")

            # ウィンドウを非表示（一部のPPTバージョンでは設定できない場合がある）
            visible = self.config['conversion']['powerpoint'].get('visible', False)
            try:
                if hasattr(self.powerpoint, 'Visible'):
                    self.powerpoint.Visible = visible
            except Exception as e:
                logger.warning(f"PowerPoint Visible設定をスキップしました: {e}")
                # 続行可能

            logger.info("PowerPointアプリケーション初期化完了")

        except Exception as e:
            logger.error(f"PowerPoint初期化エラー: {e}")
            raise

    def close_powerpoint(self):
        """PowerPointアプリケーションを終了"""
        if self.powerpoint:
            try:
                self.powerpoint.Quit()
                self.powerpoint = None
                pythoncom.CoUninitialize()
                logger.info("PowerPointアプリケーション終了")
            except Exception as e:
                logger.warning(f"PowerPoint終了時の警告: {e}")

    def scan_pptx_files(self) -> list:
        """
        PPTXファイルをスキャン

        Returns:
            ファイル情報のリスト
        """
        logger.info(f"スキャン開始: {self.source_folder}")

        if not self.source_folder.exists():
            raise FileNotFoundError(f"ソースフォルダが存在しません: {self.source_folder}")

        pptx_files = []

        # 検索パターン
        if self.recursive:
            patterns = ['**/*.pptx', '**/*.ppt']
        else:
            patterns = ['*.pptx', '*.ppt']

        for pattern in patterns:
            for file_path in self.source_folder.glob(pattern):
                # 一時ファイルをスキップ
                if file_path.name.startswith('~$'):
                    continue

                stat = file_path.stat()
                file_info = {
                    'path': file_path,
                    'name': file_path.name,
                    'modified': datetime.fromtimestamp(stat.st_mtime),
                    'size': stat.st_size
                }

                pptx_files.append(file_info)

        logger.info(f"検出ファイル数: {len(pptx_files)}")
        return pptx_files

    def get_output_path(self, source_path: Path) -> Path:
        """
        PDF出力パスを生成

        Args:
            source_path: ソースファイルパス

        Returns:
            PDF出力パス
        """
        # 相対パスを計算
        if self.preserve_structure:
            relative_path = source_path.relative_to(self.source_folder)
            pdf_path = self.output_folder / relative_path.parent / f"{relative_path.stem}.pdf"
        else:
            pdf_path = self.output_folder / f"{source_path.stem}.pdf"

        # 出力ディレクトリを作成
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        return pdf_path

    def convert_to_pdf(self, source_path: Path, pdf_path: Path) -> bool:
        """
        PPTXをPDFに変換

        Args:
            source_path: ソースファイルパス
            pdf_path: PDF出力パス

        Returns:
            True: 成功, False: 失敗
        """
        presentation = None

        try:
            logger.info(f"変換開始: {source_path.name}")

            # プレゼンテーションを開く
            presentation = self.powerpoint.Presentations.Open(
                str(source_path.resolve()),
                ReadOnly=True,
                WithWindow=False
            )

            # PDF品質設定
            # ppFixedFormatIntentScreen = 1
            # ppFixedFormatIntentPrint = 2
            quality_map = {
                'standard': 1,
                'high': 2,
                'minimum': 1
            }
            quality = quality_map.get(
                self.config['conversion'].get('pdf_quality', 'standard'),
                1
            )

            # PDFとして保存 (ExportAsFixedFormatを使用)
            # ppFixedFormatTypePDF = 2
            # PrintRange=None is required to avoid COM object error
            presentation.ExportAsFixedFormat(
                str(pdf_path.resolve()),
                2,  # ppFixedFormatTypePDF
                PrintRange=None
            )

            logger.info(f"変換成功: {pdf_path.name}")
            return True

        except Exception as e:
            logger.error(f"変換エラー ({source_path.name}): {e}")
            return False

        finally:
            # プレゼンテーションを閉じる
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass

    def process_file(self, file_info: dict) -> dict:
        """
        単一ファイルを処理

        Args:
            file_info: ファイル情報

        Returns:
            処理結果
        """
        source_path = file_info['path']
        source_modified = file_info['modified']

        # ファイルID生成
        file_id = hashlib.sha1(str(source_path.resolve()).encode()).hexdigest()[:12]

        # 変換が必要か確認
        if not self.db.needs_conversion(source_path, source_modified):
            logger.debug(f"変換不要（変更なし）: {source_path.name}")
            return {'status': 'skipped', 'file_id': file_id}

        # PDF出力パス
        pdf_path = self.get_output_path(source_path)

        # 変換実行
        start_time = datetime.now()

        try:
            success = self.convert_to_pdf(source_path, pdf_path)

            duration = (datetime.now() - start_time).total_seconds()

            if success:
                # 成功を記録
                self.db.record_conversion(
                    file_id=file_id,
                    source_path=source_path,
                    pdf_path=pdf_path,
                    source_modified=source_modified,
                    status='success'
                )

                return {
                    'status': 'success',
                    'file_id': file_id,
                    'source': str(source_path),
                    'pdf': str(pdf_path),
                    'duration': duration
                }
            else:
                # 失敗を記録
                self.db.record_conversion(
                    file_id=file_id,
                    source_path=source_path,
                    pdf_path=pdf_path,
                    source_modified=source_modified,
                    status='failed',
                    error_message='変換失敗'
                )

                return {
                    'status': 'failed',
                    'file_id': file_id,
                    'error': '変換失敗'
                }

        except Exception as e:
            # エラーを記録
            self.db.record_conversion(
                file_id=file_id,
                source_path=source_path,
                pdf_path=pdf_path,
                source_modified=source_modified,
                status='failed',
                error_message=str(e)
            )

            return {
                'status': 'failed',
                'file_id': file_id,
                'error': str(e)
            }

    def run(self) -> dict:
        """
        変換処理を実行

        Returns:
            処理結果の統計
        """
        start_time = datetime.now()

        logger.info("=== PPTX→PDF変換開始 ===")

        try:
            # PowerPoint初期化
            self.initialize_powerpoint()

            # ファイルスキャン
            pptx_files = self.scan_pptx_files()

            if not pptx_files:
                logger.info("処理対象ファイルなし")
                return {
                    'status': 'success',
                    'converted': 0,
                    'skipped': 0,
                    'failed': 0,
                    'duration_seconds': 0
                }

            # 処理実行
            converted = 0
            skipped = 0
            failed = 0

            for i, file_info in enumerate(pptx_files, 1):
                logger.info(f"\n処理中: {i}/{len(pptx_files)} - {file_info['name']}")

                result = self.process_file(file_info)

                if result['status'] == 'success':
                    converted += 1
                elif result['status'] == 'skipped':
                    skipped += 1
                else:
                    failed += 1

            # 完了
            duration = (datetime.now() - start_time).total_seconds()

            # 統計
            stats = self.db.get_statistics()

            logger.info(f"\n=== 変換完了 ===")
            logger.info(f"処理時間: {duration:.2f}秒")
            logger.info(f"変換: {converted}")
            logger.info(f"スキップ: {skipped}")
            logger.info(f"失敗: {failed}")

            return {
                'status': 'success',
                'converted': converted,
                'skipped': skipped,
                'failed': failed,
                'duration_seconds': duration,
                'statistics': stats
            }

        finally:
            # PowerPoint終了
            self.close_powerpoint()
            self.db.close()


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(description="ローカルPPTX→PDF変換")
    parser.add_argument(
        '--config',
        type=Path,
        default=Path('configs/local_convert.yaml'),
        help='設定ファイルのパス'
    )

    args = parser.parse_args()

    # 設定ファイル読み込み
    if not args.config.exists():
        print(f"❌ エラー: 設定ファイルが見つかりません: {args.config}")
        print("\n以下のコマンドでテンプレートをコピーしてください:")
        print("  cp configs/local_convert.yaml.template configs/local_convert.yaml")
        sys.exit(1)

    with open(args.config, encoding='utf-8') as f:
        config = yaml.safe_load(f)

    # ログディレクトリ作成
    Path('data/logs').mkdir(parents=True, exist_ok=True)

    # 変換実行
    converter = PPTXToPDFConverter(config)

    try:
        result = converter.run()

        # 結果出力
        print("\n" + "="*50)
        print("処理結果:")
        print(f"  ステータス: {result['status']}")
        print(f"  変換: {result['converted']}")
        print(f"  スキップ: {result['skipped']}")
        print(f"  失敗: {result['failed']}")
        print(f"  処理時間: {result['duration_seconds']:.2f}秒")
        print("="*50)

        sys.exit(0 if result['failed'] == 0 else 1)

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
