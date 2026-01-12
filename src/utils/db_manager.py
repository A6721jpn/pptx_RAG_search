"""
処理状態管理データベース
SQLiteを使用してSharePointファイルの処理状態を追跡
"""

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict
import logging

logger = logging.getLogger(__name__)


class ProcessedFilesDB:
    """処理済みファイルの状態管理データベース"""

    def __init__(self, db_path: Path):
        """
        Args:
            db_path: SQLiteデータベースファイルのパス
        """
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn: Optional[sqlite3.Connection] = None
        self._initialize_db()

    def _initialize_db(self):
        """データベーステーブルを初期化"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row  # 辞書形式で結果を取得

        cursor = self.conn.cursor()

        # 処理済みファイルテーブル
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS processed_files (
                file_id TEXT PRIMARY KEY,
                file_name TEXT NOT NULL,
                sharepoint_url TEXT,
                sharepoint_path TEXT,
                site_id TEXT,
                drive_id TEXT,
                modified_date TIMESTAMP,
                file_size INTEGER,
                doc_id TEXT,
                processed_at TIMESTAMP,
                status TEXT DEFAULT 'pending',
                error_message TEXT,
                slide_count INTEGER,
                processing_duration_seconds REAL
            )
        """)

        # インデックス作成
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_modified
            ON processed_files(modified_date)
        """)
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_status
            ON processed_files(status)
        """)
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_doc_id
            ON processed_files(doc_id)
        """)

        # 処理ログテーブル
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS processing_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id TEXT NOT NULL,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                event_type TEXT,
                message TEXT,
                FOREIGN KEY (file_id) REFERENCES processed_files(file_id)
            )
        """)

        self.conn.commit()
        logger.info(f"データベース初期化完了: {self.db_path}")

    def add_or_update_file(self, file_info: Dict) -> bool:
        """
        ファイル情報を追加または更新

        Args:
            file_info: SharePointファイル情報

        Returns:
            True: 新規追加または更新が必要, False: 既存で変更なし
        """
        cursor = self.conn.cursor()

        # 既存レコードを確認
        cursor.execute(
            "SELECT modified_date, status FROM processed_files WHERE file_id = ?",
            (file_info['id'],)
        )
        row = cursor.fetchone()

        # 新規ファイル
        if row is None:
            cursor.execute("""
                INSERT INTO processed_files (
                    file_id, file_name, sharepoint_url, sharepoint_path,
                    site_id, drive_id, modified_date, file_size, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending')
            """, (
                file_info['id'],
                file_info['name'],
                file_info.get('web_url'),
                file_info.get('path'),
                file_info.get('site_id'),
                file_info.get('drive_id'),
                file_info['modified'],
                file_info['size']
            ))
            self.conn.commit()
            logger.info(f"新規ファイル追加: {file_info['name']}")
            return True

        # 既存ファイルが更新されている場合
        existing_modified = datetime.fromisoformat(row['modified_date'].replace('Z', '+00:00'))
        new_modified = file_info['modified']

        if new_modified > existing_modified:
            cursor.execute("""
                UPDATE processed_files
                SET modified_date = ?,
                    file_size = ?,
                    status = 'pending',
                    error_message = NULL
                WHERE file_id = ?
            """, (
                file_info['modified'],
                file_info['size'],
                file_info['id']
            ))
            self.conn.commit()
            logger.info(f"ファイル更新検出: {file_info['name']}")
            return True

        # 処理失敗したファイルは再処理
        if row['status'] == 'failed':
            logger.info(f"失敗ファイルを再処理対象に: {file_info['name']}")
            return True

        # 変更なし
        return False

    def get_pending_files(self, limit: Optional[int] = None) -> List[Dict]:
        """
        処理待ちファイルを取得

        Args:
            limit: 取得件数制限

        Returns:
            処理待ちファイルのリスト
        """
        cursor = self.conn.cursor()

        query = """
            SELECT * FROM processed_files
            WHERE status = 'pending'
            ORDER BY modified_date DESC
        """
        if limit:
            query += f" LIMIT {limit}"

        cursor.execute(query)
        return [dict(row) for row in cursor.fetchall()]

    def update_status(
        self,
        file_id: str,
        status: str,
        error_message: Optional[str] = None,
        doc_id: Optional[str] = None,
        slide_count: Optional[int] = None,
        duration: Optional[float] = None
    ):
        """
        ファイルの処理状態を更新

        Args:
            file_id: ファイルID
            status: 状態 ('processing', 'success', 'failed')
            error_message: エラーメッセージ（失敗時）
            doc_id: ドキュメントID（成功時）
            slide_count: スライド数（成功時）
            duration: 処理時間（秒）
        """
        cursor = self.conn.cursor()

        update_fields = ["status = ?"]
        params = [status]

        if error_message:
            update_fields.append("error_message = ?")
            params.append(error_message)

        if doc_id:
            update_fields.append("doc_id = ?")
            params.append(doc_id)

        if slide_count is not None:
            update_fields.append("slide_count = ?")
            params.append(slide_count)

        if duration is not None:
            update_fields.append("processing_duration_seconds = ?")
            params.append(duration)

        if status in ('success', 'failed'):
            update_fields.append("processed_at = ?")
            params.append(datetime.utcnow().isoformat())

        params.append(file_id)

        query = f"""
            UPDATE processed_files
            SET {', '.join(update_fields)}
            WHERE file_id = ?
        """

        cursor.execute(query, params)
        self.conn.commit()

        logger.info(f"状態更新: {file_id} -> {status}")

    def add_log(self, file_id: str, event_type: str, message: str):
        """
        処理ログを追加

        Args:
            file_id: ファイルID
            event_type: イベントタイプ ('download', 'extract', 'render', 'embed', 'index')
            message: ログメッセージ
        """
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO processing_logs (file_id, event_type, message)
            VALUES (?, ?, ?)
        """, (file_id, event_type, message))
        self.conn.commit()

    def get_statistics(self) -> Dict:
        """
        処理統計を取得

        Returns:
            統計情報の辞書
        """
        cursor = self.conn.cursor()

        stats = {}

        # 状態別件数
        cursor.execute("""
            SELECT status, COUNT(*) as count
            FROM processed_files
            GROUP BY status
        """)
        stats['by_status'] = {row['status']: row['count'] for row in cursor.fetchall()}

        # 総ファイル数
        cursor.execute("SELECT COUNT(*) as total FROM processed_files")
        stats['total_files'] = cursor.fetchone()['total']

        # 総スライド数
        cursor.execute("SELECT SUM(slide_count) as total FROM processed_files WHERE status = 'success'")
        result = cursor.fetchone()
        stats['total_slides'] = result['total'] if result['total'] else 0

        # 平均処理時間
        cursor.execute("""
            SELECT AVG(processing_duration_seconds) as avg_duration
            FROM processed_files
            WHERE status = 'success' AND processing_duration_seconds IS NOT NULL
        """)
        result = cursor.fetchone()
        stats['avg_processing_seconds'] = result['avg_duration'] if result['avg_duration'] else 0

        # 最終処理日時
        cursor.execute("""
            SELECT MAX(processed_at) as last_processed
            FROM processed_files
            WHERE status = 'success'
        """)
        result = cursor.fetchone()
        stats['last_processed'] = result['last_processed']

        return stats

    def get_failed_files(self) -> List[Dict]:
        """失敗したファイルのリストを取得"""
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT file_id, file_name, error_message, processed_at
            FROM processed_files
            WHERE status = 'failed'
            ORDER BY processed_at DESC
        """)
        return [dict(row) for row in cursor.fetchall()]

    def reset_failed_files(self):
        """失敗したファイルを再処理対象に設定"""
        cursor = self.conn.cursor()
        cursor.execute("""
            UPDATE processed_files
            SET status = 'pending', error_message = NULL
            WHERE status = 'failed'
        """)
        affected = cursor.rowcount
        self.conn.commit()
        logger.info(f"{affected}件の失敗ファイルを再処理対象に設定")
        return affected

    def close(self):
        """データベース接続をクローズ"""
        if self.conn:
            self.conn.close()
            logger.info("データベース接続クローズ")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# ========== 使用例 ==========

if __name__ == "__main__":
    # ロギング設定
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # データベース初期化
    db = ProcessedFilesDB(Path("data/processed_files.db"))

    # テストデータ追加
    test_file = {
        'id': 'test-file-001',
        'name': 'test_presentation.pptx',
        'web_url': 'https://example.sharepoint.com/test.pptx',
        'path': '/Design Guides/test.pptx',
        'modified': datetime.utcnow(),
        'size': 1024000
    }

    # ファイル追加
    is_new = db.add_or_update_file(test_file)
    print(f"新規追加: {is_new}")

    # 処理中に更新
    db.update_status('test-file-001', 'processing')

    # ログ追加
    db.add_log('test-file-001', 'download', 'ダウンロード開始')

    # 処理成功
    db.update_status(
        'test-file-001',
        'success',
        doc_id='abc123def456',
        slide_count=25,
        duration=45.2
    )

    # 統計取得
    stats = db.get_statistics()
    print("\n統計情報:")
    for key, value in stats.items():
        print(f"  {key}: {value}")

    # クローズ
    db.close()
