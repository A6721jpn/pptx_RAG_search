"""
SharePoint同期モジュール
Microsoft Graph APIを使用してSharePointからPPTXファイルを取得し、
バッチ処理パイプラインで大規模にインデックス化
"""

from .sharepoint_client import SharePointClient
from .sync_pipeline import SharePointSyncPipeline, SyncConfig

__all__ = [
    'SharePointClient',
    'SharePointSyncPipeline',
    'SyncConfig'
]

__version__ = '0.1.0'
