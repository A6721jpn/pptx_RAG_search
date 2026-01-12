# SharePoint大規模展開設計書

**著者:** Claude
**日付:** 2026-01-12
**ステータス:** Draft v1
**ベースドキュメント:** pptx_rag_local_poc_design.md

---

## 1. 概要

ローカルPOCで検証したPPTX RAGシステムを、SharePoint上の大量のPPTXファイルに対して展開するための設計。

### 1.1 主な変更点

- **データソース**: ローカルファイル → SharePoint Online/Server
- **スケール**: 数十ファイル → 数百〜数千ファイル
- **処理モデル**: 手動実行 → 自動バッチ処理 + 増分更新
- **認証**: なし → OAuth 2.0 / Entra ID
- **デプロイ**: ローカルマシン → 専用Windowsサーバー

---

## 2. アーキテクチャ

### 2.1 システム構成
```
SharePoint Online/Server
    ↓ (Microsoft Graph API or SharePoint REST API)
専用Windowsサーバー
    - SharePoint統合モジュール
    - PPTX処理パイプライン (COM使用)
    - Qdrantローカルインスタンス
    - スケジューラー (タスクスケジューラー or Celery)
```

### 2.2 主な特徴

**メリット:**
- 既存のPowerPoint COM実装を再利用可能
- 安定したレンダリング品質
- ローカルQdrantで低レイテンシ
- 実装が比較的シンプル

**制約事項:**
- COM処理はシングルスレッド制約あり（順次処理必須）
- Windowsサーバー環境が必要

**適用規模:**
- 数百〜数千ファイル規模
- 1日1回のバッチ更新
- 既存のオンプレミスインフラ活用可能

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

## 8. デプロイメント手順

### Phase 1: 開発環境セットアップ (1-2日)

1. **Azure ADアプリ登録**
   - Azure Portalでアプリ登録
   - 必要な権限付与 (`Sites.Read.All`, `Files.Read.All`)
   - クライアントシークレット生成

2. **開発環境構築**
   ```bash
   git clone <repository>
   cd pptx_RAG_search
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **設定ファイル作成**
   ```yaml
   # configs/sharepoint_prod.yaml
   sharepoint:
     tenant_id: "your-tenant-id"
     client_id: "your-client-id"
     client_secret: "your-client-secret"
     site_urls:
       - "https://company.sharepoint.com/sites/Engineering"
       - "https://company.sharepoint.com/sites/DesignGuides"

   processing:
     batch_size: 50
     parallel_downloads: 10
     retry_attempts: 3

   qdrant:
     path: "index/qdrant_storage"
     collection_name: "pptx_slides"
   ```

### Phase 2: POC実装 (3-5日)

1. SharePoint統合モジュール実装
2. バッチ処理パイプライン実装
3. 小規模テスト (10-50ファイル)
4. エラーハンドリング確認

### Phase 3: スケールテスト (2-3日)

1. 中規模テスト (100-500ファイル)
2. パフォーマンスチューニング
3. メモリ使用量最適化
4. ログ・監視確認

### Phase 4: 本番展開 (1-2日)

1. 本番サーバーセットアップ
2. フルスキャン実行
3. スケジューラー設定
4. 運用手順書作成

---

## 9. 運用考慮事項

### 9.1 セキュリティ

- **シークレット管理**: Azure Key Vault使用推奨
- **アクセス制御**: 最小権限の原則
- **ログ保護**: 個人情報を含まないよう注意
- **ネットワーク**: ファイアウォール設定

### 9.2 バックアップ

```bash
# Qdrantインデックスのバックアップ
tar -czf qdrant_backup_$(date +%Y%m%d).tar.gz index/qdrant_storage/

# 処理状態DBのバックアップ
sqlite3 data/processed_files.db ".backup 'data/backups/processed_files_$(date +%Y%m%d).db'"
```

### 9.3 ディザスタリカバリ

- **定期バックアップ**: 週次フルバックアップ
- **復旧手順**: ドキュメント化
- **テスト**: 四半期ごとに復旧テスト

---

## 10. コスト見積もり (1000ファイル想定)

### システム運用コスト

| 項目 | コスト (月額) |
|------|---------------|
| Windows Server (Azure VM Standard_D4s_v3) | $140 |
| ストレージ (500GB Premium SSD) | $70 |
| アウトバウンド通信 (100GB) | $9 |
| **合計** | **$219/月** |

**注記:** オンプレミスサーバーを使用する場合はクラウドコストは不要

### 処理時間見積もり

- 1ファイルあたり平均処理時間: 30秒
- 1000ファイル: 約8.3時間 (初回フルスキャン)
- 増分更新 (5%変更): 約25分/日

---

## 11. 次のステップ

1. **要件確認**
   - SharePoint環境 (Online/Server)
   - ファイル数・総容量
   - 更新頻度

2. **Azure AD準備**
   - アプリ登録権限の取得
   - テストサイトの確保

3. **実装開始**
   - SharePoint統合モジュール（実装済み）
   - バッチ処理パイプライン（実装済み）

---

## 付録: 実装済みコンポーネント

すべての実装コードは `src/sharepoint_sync/` ディレクトリに配置されています：

- **sharepoint_client.py**: Microsoft Graph API クライアント
- **sync_pipeline.py**: バッチ処理パイプライン
- **db_manager.py**: 処理状態管理 (`src/utils/`)

設定テンプレートは `configs/sharepoint_template.yaml` を参照してください。

---

End of document.
