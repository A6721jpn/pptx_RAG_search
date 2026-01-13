"""
Local Processed Files Database
Tracks the processing status of local PowerPoint files.
"""

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class ProcessedFilesDB:
    """Database for tracking local file processing status."""

    def __init__(self, db_path: Path):
        """
        Args:
            db_path: Path to the SQLite database file
        """
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn: Optional[sqlite3.Connection] = None
        self._initialize_db()

    def _initialize_db(self):
        """Initialize database tables."""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row

        cursor = self.conn.cursor()

        # Processed files table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS processed_files (
                file_path TEXT PRIMARY KEY,
                file_hash TEXT,
                file_size INTEGER,
                modified_time REAL,
                status TEXT DEFAULT 'pending',
                error_message TEXT,
                doc_id TEXT,
                slide_count INTEGER,
                processed_at TIMESTAMP,
                processing_duration_seconds REAL
            )
        """)

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_status ON processed_files(status)")
        
        # Processing logs table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS processing_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                event_type TEXT,
                message TEXT,
                FOREIGN KEY (file_path) REFERENCES processed_files(file_path)
            )
        """)
        
        self.conn.commit()

    def should_process(self, file_path: str, file_hash: str, modified_time: float) -> bool:
        """
        Check if a file needs to be processed.
        Returns True if new or modified.
        """
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT file_hash, modified_time, status FROM processed_files WHERE file_path = ?",
            (file_path,)
        )
        row = cursor.fetchone()

        if row is None:
            return True  # New file

        # Check if modified (by hash or mtime, hash is safer but mtime is faster)
        # Here we assume caller provides hash if they want hash-based check, or consistent mtime.
        # If hash is provided and different, reprocess.
        if file_hash and row['file_hash'] != file_hash:
            return True
        
        # Fallback to modified time if hash matches or not computed yet but mtime changed?
        # Ideally rely on hash. If mtime implies change, we might want to re-check hash.
        # For simplicity, if modified_time is strictly newer, we process.
        if modified_time > row['modified_time']:
            return True

        if row['status'] == 'failed':
            return True # Retry failed

        return False

    def register_file(self, file_path: str, file_hash: str, file_size: int, modified_time: float):
        """Register or update a file as pending processing."""
        cursor = self.conn.cursor()
        
        # Check if exists
        cursor.execute("SELECT 1 FROM processed_files WHERE file_path = ?", (file_path,))
        exists = cursor.fetchone() is not None

        if exists:
            cursor.execute("""
                UPDATE processed_files
                SET file_hash = ?, file_size = ?, modified_time = ?, status = 'pending', error_message = NULL
                WHERE file_path = ?
            """, (file_hash, file_size, modified_time, file_path))
        else:
            cursor.execute("""
                INSERT INTO processed_files (file_path, file_hash, file_size, modified_time, status)
                VALUES (?, ?, ?, ?, 'pending')
            """, (file_path, file_hash, file_size, modified_time))
        
        self.conn.commit()

    def update_status(self, file_path: str, status: str, error_message: str = None, 
                      doc_id: str = None, slide_count: int = None, duration: float = None):
        """Update processing status."""
        cursor = self.conn.cursor()
        
        fields = ["status = ?"]
        params = [status]

        if error_message is not None:
            fields.append("error_message = ?")
            params.append(error_message)
        
        if doc_id is not None:
            fields.append("doc_id = ?")
            params.append(doc_id)
            
        if slide_count is not None:
            fields.append("slide_count = ?")
            params.append(slide_count)
            
        if duration is not None:
            fields.append("processing_duration_seconds = ?")
            params.append(duration)
            
        if status in ('success', 'failed'):
            fields.append("processed_at = ?")
            params.append(datetime.utcnow().isoformat())

        params.append(file_path)

        cursor.execute(f"UPDATE processed_files SET {', '.join(fields)} WHERE file_path = ?", params)
        self.conn.commit()

    def add_or_update_file(self, file_info: Dict[str, Any]) -> bool:
        """
        Check if file needs processing and register it.
        Compatible with local_poc_pdf.py which expects file_info dict.
        """
        file_path = str(file_info['path'])
        
        # Calculate hash if not present (or use modified time logic)
        # For simplicity in this alias, we assume file_info has what we need
        # local_poc_pdf.py passes file_info with 'path', 'id' (as hash), 'size', 'modified'
        
        file_hash = file_info.get('id', '')
        file_size = file_info.get('size', 0)
        modified_time = file_info.get('modified').timestamp() if isinstance(file_info.get('modified'), datetime) else 0.0

        if self.should_process(file_path, file_hash, modified_time):
            self.register_file(file_path, file_hash, file_size, modified_time)
            return True
        return False

    def close(self):
        if self.conn:
            self.conn.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
