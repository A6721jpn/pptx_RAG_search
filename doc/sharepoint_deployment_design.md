# SharePoint中規模POC設計書

**著者:** Claude
**日付:** 2026-01-12
**ステータス:** Draft v2 - ローカルPC対応
**ベースドキュメント:** pptx_rag_local_poc_design.md

---

## 1. 概要

ローカルPOCで検証したPPTX RAGシステムを、SharePoint上の実際のPPTXファイル（約100ファイル規模）に対して実行し、実用性と処理速度を検証するための中規模POC設計。

### 1.1 主な変更点

- **データソース**: ローカルファイル → SharePoint Online/Server
- **スケール**: 数十ファイル → 約100ファイル（中規模POC）
- **処理モデル**: 手動実行 → バッチ処理 + 増分更新
- **認証**: なし → OAuth 2.0 / Entra ID
- **実行環境**: ローカルWindows PC

---

## 2. アーキテクチャ

### 2.1 システム構成（ローカルPC）
```
SharePoint Online/Server
    ↓ (Microsoft Graph API)
ローカルWindows PC
    - SharePoint統合モジュール
    - PPTX処理パイプライン (PowerPoint COM使用)
    - Qdrantローカルインスタンス
    - 手動実行またはスケジューラー（オプション）
```

### 2.2 POCの目的と特徴

**検証項目:**
- SharePoint接続と認証の動作確認
- 実際のファイルに対する処理速度の測定
- 100ファイル規模での実用性評価
- エラーハンドリングの妥当性確認

**技術的特徴:**
- PowerPoint COM による安定したレンダリング
- ローカルQdrantで低レイテンシ
- 増分更新による効率的な再処理
- 処理状態の永続化とリトライ機能

**制約事項:**
- COM処理はシングルスレッド（順次処理）
- Windows環境必須（PowerPoint要インストール）
- ローカルPC実行のため、大規模展開時は別途検討必要

**POC規模:**
- 対象: 約100ファイル
- 想定処理時間: 初回50分程度（1ファイル30秒想定）
- 増分更新: 5-10分程度（変更分のみ）

---

## 3. SharePoint統合設計

### 3.1 認証方法

**Microsoft Graph API使用 (推奨):**

```python
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient

# Entra ID (Azure AD) アプリ登録が必要
credential = ClientSecretCredential(
    tenant_id="YOUR_TENANT_ID",
    client_id="YOUR_CLIENT_ID",
    client_secret="YOUR_CLIENT_SECRET"
)

graph_client = GraphServiceClient(credentials=credential)
```

**必要な権限:**
- `Sites.Read.All` (SharePointサイト読み取り)
- `Files.Read.All` (ファイル読み取り)

### 3.2 SharePointファイル探索

**パターン1: 特定ドキュメントライブラリをスキャン**

```python
async def list_pptx_files(site_id, drive_id):
    """SharePointドライブからすべてのPPTXファイルを取得"""
    items = await graph_client.sites.by_site_id(site_id)\
        .drives.by_drive_id(drive_id)\
        .root.children.get()

    pptx_files = []
    for item in items.value:
        if item.file and item.name.endswith(('.pptx', '.ppt')):
            pptx_files.append({
                'id': item.id,
                'name': item.name,
                'web_url': item.web_url,
                'modified': item.last_modified_date_time,
                'size': item.size,
                'download_url': item.microsoft_graph_download_url
            })

    return pptx_files
```

**パターン2: 検索クエリで全サイトからPPTXを取得**

```python
async def search_all_pptx(query="*.pptx"):
    """全SharePointサイトからPPTXファイルを検索"""
    search_results = await graph_client.search.query.post({
        "requests": [{
            "entityTypes": ["driveItem"],
            "query": {
                "queryString": "fileExtension:pptx OR fileExtension:ppt"
            },
            "size": 1000
        }]
    })
    return search_results
```

### 3.3 ファイルダウンロード戦略

**ストリーミングダウンロード (大容量対応):**

```python
import aiohttp
import aiofiles

async def download_pptx(download_url, local_path):
    """SharePointからPPTXをダウンロード"""
    async with aiohttp.ClientSession() as session:
        # Graph API tokenを使用
        headers = {'Authorization': f'Bearer {access_token}'}

        async with session.get(download_url, headers=headers) as resp:
            async with aiofiles.open(local_path, 'wb') as f:
                async for chunk in resp.content.iter_chunked(8192):
                    await f.write(chunk)
```

**増分ダウンロード (変更分のみ):**

```python
def needs_update(file_metadata, local_cache_db):
    """ファイルが更新されたか確認"""
    cached = local_cache_db.get(file_metadata['id'])

    if not cached:
        return True  # 新規ファイル

    # 変更日時を比較
    if file_metadata['modified'] > cached['modified']:
        return True

    return False
```

---

## 4. バッチ処理パイプライン

### 4.1 全体フロー

```
[1] SharePointファイル一覧取得
    ↓
[2] 変更検知 (増分処理)
    ↓
[3] ダウンロードキュー作成
    ↓
[4] 並列ダウンロード (10並列)
    ↓
[5] PPTX処理キュー
    ↓
[6] 順次処理 (COM制約によりシングルスレッド)
    - テキスト抽出
    - スライドレンダリング
    - 埋め込み計算
    - Qdrantインデックス更新
    ↓
[7] 処理完了・メタデータ更新
```

### 4.2 状態管理データベース

**SQLiteで処理状態を管理:**

```sql
CREATE TABLE processed_files (
    file_id TEXT PRIMARY KEY,
    file_name TEXT,
    sharepoint_url TEXT,
    modified_date TIMESTAMP,
    doc_id TEXT,  -- content hash
    processed_at TIMESTAMP,
    status TEXT,  -- 'success', 'failed', 'processing'
    error_message TEXT,
    slide_count INTEGER
);

CREATE INDEX idx_modified ON processed_files(modified_date);
CREATE INDEX idx_status ON processed_files(status);
```

### 4.3 エラーハンドリング

**再試行ロジック (tenacity使用):**

```python
from tenacity import retry, stop_after_attempt, wait_exponential

@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    reraise=True
)
async def download_with_retry(file_metadata):
    """ダウンロードを最大3回リトライ"""
    return await download_pptx(
        file_metadata['download_url'],
        f"data/pptx_temp/{file_metadata['id']}.pptx"
    )
```

**失敗ファイルの隔離:**

```python
def handle_processing_error(file_id, error):
    """処理失敗時のハンドリング"""
    db.update_status(file_id, 'failed', str(error))

    # 管理者通知
    notify_admin(f"Processing failed for {file_id}: {error}")

    # 失敗ログ記録
    with open('data/logs/failed_files.jsonl', 'a') as f:
        json.dump({
            'file_id': file_id,
            'timestamp': datetime.now().isoformat(),
            'error': str(error)
        }, f)
        f.write('\n')
```

---

## 5. スケーラビリティ最適化

### 5.1 並列化戦略

**段階的並列化:**

1. **ダウンロード**: 10-20並列 (I/O bound)
2. **テキスト抽出**: 5並列 (python-pptxは軽量)
3. **スライドレンダリング**: 1並列 (COM制約)
4. **埋め込み計算**: 4並列 (CPU bound)

**実装例 (asyncio + ProcessPoolExecutor):**

```python
import asyncio
from concurrent.futures import ProcessPoolExecutor

async def process_batch(file_list):
    # ダウンロード (並列)
    download_tasks = [download_pptx(f) for f in file_list]
    local_files = await asyncio.gather(*download_tasks)

    # テキスト抽出 (並列)
    with ProcessPoolExecutor(max_workers=5) as executor:
        loop = asyncio.get_event_loop()
        extract_tasks = [
            loop.run_in_executor(executor, extract_text, f)
            for f in local_files
        ]
        extracted = await asyncio.gather(*extract_tasks)

    # レンダリング (順次 - COM制約)
    for file in local_files:
        render_slides_com(file)

    # 埋め込み (並列)
    with ProcessPoolExecutor(max_workers=4) as executor:
        loop = asyncio.get_event_loop()
        embed_tasks = [
            loop.run_in_executor(executor, compute_embeddings, ex)
            for ex in extracted
        ]
        embeddings = await asyncio.gather(*embed_tasks)
```

### 5.2 メモリ管理

**大規模処理時のメモリ対策:**

```python
def process_in_batches(file_list, batch_size=50):
    """バッチサイズを制限してメモリ消費を抑制"""
    for i in range(0, len(file_list), batch_size):
        batch = file_list[i:i+batch_size]

        # バッチ処理
        process_batch(batch)

        # メモリクリーンアップ
        gc.collect()

        # 進捗ログ
        logger.info(f"Processed {i+len(batch)}/{len(file_list)} files")
```

### 5.3 増分インデックス更新

**Qdrantでの効率的な更新:**

```python
def update_index_incremental(doc_id, slide_records):
    """既存ドキュメントのスライドを削除してから追加"""
    from qdrant_client import QdrantClient
    from qdrant_client.models import Filter, FieldCondition, MatchValue

    client = QdrantClient(path="index/qdrant_storage")

    # 既存の同一ドキュメントのポイントを削除
    client.delete(
        collection_name="pptx_slides",
        points_selector=Filter(
            must=[
                FieldCondition(
                    key="doc_id",
                    match=MatchValue(value=doc_id)
                )
            ]
        )
    )

    # 新しいポイントを一括追加
    client.upload_points(
        collection_name="pptx_slides",
        points=slide_records
    )
```

---

## 6. スケジューリング

### 6.1 Windowsタスクスケジューラー (シンプル)

**バッチスクリプト例:**

```batch
@echo off
REM 毎日午前2時に実行
cd C:\pptx_rag_deployment
python src\sharepoint_sync\sync_and_index.py --incremental
```

**タスクスケジューラー設定:**
- トリガー: 毎日午前2時
- アクション: 上記バッチファイル実行
- 条件: AC電源接続時のみ
- 通知: 失敗時にメール送信

### 6.2 Celery + Redis (スケーラブル)

**構成:**

```python
# celery_config.py
from celery import Celery
from celery.schedules import crontab

app = Celery('pptx_rag', broker='redis://localhost:6379/0')

app.conf.beat_schedule = {
    'daily-sharepoint-sync': {
        'task': 'tasks.sync_and_index',
        'schedule': crontab(hour=2, minute=0),  # 毎日午前2時
    },
}

# tasks.py
@app.task
def sync_and_index():
    """SharePoint同期とインデックス更新"""
    from sharepoint_sync import SharePointSyncPipeline

    pipeline = SharePointSyncPipeline()
    results = pipeline.run(incremental=True)

    return {
        'processed': results['success_count'],
        'failed': results['failed_count'],
        'duration': results['duration_seconds']
    }
```

---

## 7. 監視とログ

### 7.1 ロギング戦略

**構造化ログ (JSON形式):**

```python
import logging
import json
from datetime import datetime

class JSONFormatter(logging.Formatter):
    def format(self, record):
        log_obj = {
            'timestamp': datetime.utcnow().isoformat(),
            'level': record.levelname,
            'component': record.name,
            'message': record.getMessage(),
        }
        if record.exc_info:
            log_obj['exception'] = self.formatException(record.exc_info)
        return json.dumps(log_obj)

# 設定
logger = logging.getLogger('pptx_rag')
handler = logging.FileHandler('logs/pptx_rag.jsonl')
handler.setFormatter(JSONFormatter())
logger.addHandler(handler)
```

### 7.2 メトリクス収集

**処理メトリクス:**

```python
metrics = {
    'batch_start': datetime.now(),
    'files_discovered': 0,
    'files_downloaded': 0,
    'files_processed': 0,
    'files_failed': 0,
    'total_slides': 0,
    'avg_processing_time_per_file': 0,
    'errors': []
}

# Prometheusエクスポート (オプション)
from prometheus_client import Counter, Histogram, Gauge

files_processed = Counter('pptx_files_processed_total', 'Total files processed')
processing_duration = Histogram('pptx_processing_seconds', 'Time to process file')
active_files = Gauge('pptx_active_processing', 'Files currently being processed')
```

### 7.3 アラート

**失敗率が閾値を超えたら通知:**

```python
def check_and_alert(metrics):
    failure_rate = metrics['files_failed'] / max(metrics['files_discovered'], 1)

    if failure_rate > 0.1:  # 10%以上失敗
        send_alert(
            severity='HIGH',
            message=f"Batch processing failure rate: {failure_rate:.1%}",
            details=metrics
        )

def send_alert(severity, message, details):
    """メール or Teams通知"""
    # 実装: SMTPまたはMicrosoft Teams Webhook
    pass
```

---

## 8. 中規模POC実行手順

### ステップ1: 環境セットアップ（30分）

1. **前提条件確認**
   - Windows 10/11 PC
   - Microsoft PowerPoint インストール済み
   - Python 3.9+ インストール済み
   - SharePointアクセス権限あり

2. **リポジトリクローン**
   ```bash
   git clone https://github.com/A6721jpn/pptx_RAG_search.git
   cd pptx_RAG_search
   ```

3. **Python環境構築**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

### ステップ2: Azure AD設定（30分）

1. **Azure Portalでアプリ登録**
   - [Azure Portal](https://portal.azure.com) → Azure Active Directory
   - アプリの登録 → 新規登録
   - アプリ名: "PPTX-RAG-POC"
   - サポートされているアカウントの種類: 「この組織ディレクトリのみ」

2. **API権限追加**
   - APIのアクセス許可 → Microsoft Graph
   - アプリケーション権限:
     - `Sites.Read.All`
     - `Files.Read.All`
   - 「管理者の同意を付与」をクリック

3. **クライアントシークレット作成**
   - 証明書とシークレット → 新しいクライアントシークレット
   - 説明: "POC用"、有効期限: 6ヶ月
   - **値をコピー**（後で使用）

4. **情報をメモ**
   - テナントID（ディレクトリID）
   - アプリケーションID（クライアントID）
   - クライアントシークレット（値）

### ステップ3: 設定ファイル作成（10分）

```bash
# テンプレートをコピー
cp configs/sharepoint_template.yaml configs/sharepoint_poc.yaml
```

**`configs/sharepoint_poc.yaml` を編集:**
```yaml
sharepoint:
  tenant_id: "YOUR_TENANT_ID"        # ステップ2でメモしたテナントID
  client_id: "YOUR_CLIENT_ID"        # アプリケーションID
  client_secret: "YOUR_SECRET"       # クライアントシークレット
  site_urls:
    - "https://yourcompany.sharepoint.com/sites/YourSite"  # 実際のSharePointサイトURL

processing:
  batch_size: 20         # 小さめに設定（ローカルPC用）
  parallel_downloads: 5  # ローカルPC用に控えめ
  retry_attempts: 3

qdrant:
  path: "index/qdrant_storage"
  collection_name: "pptx_slides_poc"
```

### ステップ4: 接続テスト（10分）

**SharePoint接続確認スクリプト実行:**
```python
# test_connection.py
import asyncio
from pathlib import Path
import yaml
from src.sharepoint_sync.sharepoint_client import SharePointClient

async def test():
    # 設定読み込み
    with open('configs/sharepoint_poc.yaml') as f:
        config = yaml.safe_load(f)

    # クライアント作成
    client = SharePointClient(
        tenant_id=config['sharepoint']['tenant_id'],
        client_id=config['sharepoint']['client_id'],
        client_secret=config['sharepoint']['client_secret']
    )

    try:
        # サイトID取得
        site_url = config['sharepoint']['site_urls'][0]
        site_id = await client.get_site_id(site_url)
        print(f"✅ SharePoint接続成功: {site_id}")

        # ドライブID取得
        drive_id = await client.get_drive_id(site_id)
        print(f"✅ ドライブアクセス成功: {drive_id}")

        # ファイル一覧取得（最初の10件）
        files = await client.list_pptx_files(site_id, drive_id)
        print(f"✅ PPTXファイル検出: {len(files)}件")
        for i, f in enumerate(files[:10], 1):
            print(f"  {i}. {f['name']} ({f['size']/1024:.1f} KB)")

    finally:
        await client.close()

if __name__ == "__main__":
    asyncio.run(test())
```

実行:
```bash
python test_connection.py
```

### ステップ5: 中規模POC実行（1-2時間）

**初回フルスキャン:**
```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_poc.yaml --full
```

**処理の進捗確認:**
- ターミナルに進捗ログが表示されます
- 処理状態は `data/processed_files.db` に保存されます
- ログは `data/logs/sync_pipeline.log` に記録されます

### ステップ6: 結果確認（10分）

**処理統計の確認:**
```python
from pathlib import Path
from src.utils.db_manager import ProcessedFilesDB

db = ProcessedFilesDB(Path('data/processed_files.db'))
stats = db.get_statistics()

print("=== POC処理結果 ===")
print(f"総ファイル数: {stats['total_files']}")
print(f"成功: {stats['by_status'].get('success', 0)}")
print(f"失敗: {stats['by_status'].get('failed', 0)}")
print(f"総スライド数: {stats['total_slides']}")
print(f"平均処理時間: {stats['avg_processing_seconds']:.2f}秒/ファイル")
print(f"最終処理日時: {stats['last_processed']}")

# 失敗ファイルの確認
failed = db.get_failed_files()
if failed:
    print(f"\n失敗ファイル ({len(failed)}件):")
    for f in failed:
        print(f"  - {f['file_name']}: {f['error_message']}")
```

### ステップ7: 増分更新テスト（オプション）

SharePoint上でファイルを更新後:
```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_poc.yaml --incremental
```

変更されたファイルのみが再処理されることを確認します。

---

## 9. POC評価観点

### 9.1 性能評価

**測定項目:**
- ファイルあたりの平均処理時間
- 全体の処理時間（100ファイル想定: 50分程度）
- メモリ使用量のピーク値
- ダウンロード速度
- エラー発生率

**評価方法:**
```python
# 処理統計の取得
from src.utils.db_manager import ProcessedFilesDB
from pathlib import Path

db = ProcessedFilesDB(Path('data/processed_files.db'))
stats = db.get_statistics()

print(f"平均処理時間: {stats['avg_processing_seconds']:.2f}秒")
print(f"成功率: {stats['by_status'].get('success', 0) / stats['total_files'] * 100:.1f}%")
```

### 9.2 機能評価

**確認項目:**
- ✅ SharePoint認証の安定性
- ✅ 増分更新の正確性（変更ファイルのみ処理）
- ✅ エラーハンドリング（リトライ機能）
- ✅ 処理状態の永続化
- ✅ ログの有用性

### 9.3 セキュリティ（POC段階）

**基本的な対策:**
- 設定ファイル（`sharepoint_poc.yaml`）を `.gitignore` に追加
- クライアントシークレットは環境変数での管理も検討
- ログに機密情報を含めない

### 9.4 データバックアップ（POC段階）

**簡易バックアップ:**
```bash
# Windowsの場合
mkdir data\backups
copy data\processed_files.db data\backups\processed_files_%date:~0,4%%date:~5,2%%date:~8,2%.db

# Qdrantインデックスのコピー
xcopy /E /I index\qdrant_storage data\backups\qdrant_backup_%date:~0,4%%date:~5,2%%date:~8,2%
```

---

## 10. POC処理時間見積もり（100ファイル規模）

### 想定処理時間

| フェーズ | 時間 | 備考 |
|---------|------|------|
| SharePoint接続・ファイル一覧取得 | 1-2分 | ネットワーク速度に依存 |
| ファイルダウンロード（並列5） | 5-10分 | ファイルサイズに依存 |
| PPTX処理（テキスト抽出・レンダリング） | 40-50分 | 1ファイル30秒×100 |
| 埋め込み計算・インデックス更新 | 5-10分 | CPU性能に依存 |
| **合計（初回フルスキャン）** | **約50-70分** | |

**増分更新（5ファイル変更想定）:**
- 約3-5分

### リソース要件（ローカルPC）

- **CPU**: 4コア以上推奨
- **メモリ**: 8GB以上（16GB推奨）
- **ディスク**: 10GB以上の空き容量
- **ネットワーク**: 安定したインターネット接続

---

## 11. POC完了後の次のステップ

### 評価結果に基づく判断

**POC成功の場合:**
1. 処理速度が実用的か評価
2. エラー率が許容範囲か確認（<5%推奨）
3. 大規模展開の検討（専用サーバー or クラウド）

**改善が必要な場合:**
1. ボトルネックの特定（ダウンロード/COM処理/埋め込み）
2. パラメータチューニング（並列数、バッチサイズ）
3. エラー原因の分析と対策

### 大規模展開への移行

POCで実用性が確認できた場合、以下を検討:
- 専用Windowsサーバーへの移行
- 自動スケジューリング設定
- 監視・アラート機能の追加
- 数百〜数千ファイルへのスケールアップ

---

## 付録: 実装済みコンポーネント

すべての実装コードは `src/sharepoint_sync/` ディレクトリに配置されています：

- **sharepoint_client.py**: Microsoft Graph API クライアント
- **sync_pipeline.py**: バッチ処理パイプライン
- **db_manager.py**: 処理状態管理 (`src/utils/`)

設定テンプレートは `configs/sharepoint_template.yaml` を参照してください。

---

End of document.
