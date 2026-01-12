# PPTX RAG Search - SharePoint中規模POC

SharePoint上のPowerPointファイルに対する検索可能なRAG（Retrieval-Augmented Generation）システムの中規模POC

## 概要

このプロジェクトは、SharePoint上に保存された約100ファイル規模のPowerPointドキュメントライブラリを対象に、テキストとビジュアルの両方を活用した検索システムの実用性と処理速度を検証する中規模POCです。

### POCの目的

- **技術検証**: SharePoint統合の動作確認
- **性能測定**: 実際のファイルに対する処理速度の計測
- **実用性評価**: 100ファイル規模での運用可能性の判断
- **課題発見**: ボトルネックやエラーの特定

### 主な機能

- **SharePoint統合**: Microsoft Graph APIを使用したシームレスな連携
- **マルチモーダル検索**: テキストと画像の両方から検索（設計済み）
- **増分更新**: 変更されたファイルのみを効率的に処理
- **エラーハンドリング**: リトライ機能と処理状態の永続化

### アーキテクチャ

```
SharePoint Online/Server
    ↓ (Microsoft Graph API)
ローカルWindows PC
    - SharePoint同期モジュール
    - PPTX処理パイプライン (PowerPoint COM)
    - Qdrantローカルインスタンス
```

## セットアップ

### 前提条件（ローカルPC）

- **OS**: Windows 10/11
- **PowerPoint**: Microsoft PowerPoint インストール済み（Office 365 または 2019+）
- **Python**: 3.9以上
- **メモリ**: 8GB以上（16GB推奨）
- **ディスク**: 10GB以上の空き容量
- **権限**: Azure AD アプリ登録権限、SharePointアクセス権限

### 1. Azure AD アプリ登録

1. [Azure Portal](https://portal.azure.com) にアクセス
2. **Azure Active Directory** > **アプリの登録** > **新規登録**
3. アプリ名を入力（例: "PPTX-RAG-Sync"）
4. **API のアクセス許可** > **アクセス許可の追加** > **Microsoft Graph**:
   - `Sites.Read.All` (アプリケーション)
   - `Files.Read.All` (アプリケーション)
5. **証明書とシークレット** > **新しいクライアントシークレット** を作成
6. 以下をメモ:
   - テナントID（ディレクトリID）
   - アプリケーション（クライアント）ID
   - クライアントシークレット

### 2. Python環境構築

```bash
# リポジトリクローン
git clone <repository-url>
cd pptx_RAG_search

# 仮想環境作成
python -m venv venv
venv\Scripts\activate  # Windows

# 依存パッケージインストール
pip install -r requirements.txt
```

### 3. 設定ファイル作成

```bash
# テンプレートをコピー
cp configs/sharepoint_template.yaml configs/sharepoint_poc.yaml

# エディタで編集
notepad configs/sharepoint_poc.yaml
```

**`configs/sharepoint_poc.yaml`** を編集:

```yaml
sharepoint:
  tenant_id: "あなたのテナントID"
  client_id: "あなたのクライアントID"
  client_secret: "あなたのクライアントシークレット"
  site_urls:
    - "https://yourcompany.sharepoint.com/sites/YourSite"  # 実際のサイトURL

processing:
  batch_size: 20         # ローカルPC用に小さめ
  parallel_downloads: 5  # ローカルPC用に控えめ

qdrant:
  collection_name: "pptx_slides_poc"
```

### 4. Qdrant起動

**オプション1: Dockerで起動（推奨）**

```bash
docker run -d -p 6333:6333 -v $(pwd)/index/qdrant_storage:/qdrant/storage qdrant/qdrant
```

**オプション2: ローカルモード**

Qdrantはファイルベースでも動作します（設定不要）

## 使い方（POC実行）

### ステップ1: SharePoint接続テスト

まず接続が正常に動作するか確認:

```bash
python test_connection.py
```

正常に動作すると、以下のように表示されます:
```
✅ SharePoint接続成功: [サイトID]
✅ ドライブアクセス成功: [ドライブID]
✅ PPTXファイル検出: 100件
  1. design_guide_01.pptx (2.3 MB)
  2. mechanical_standard.pptx (1.5 MB)
  ...
```

### ステップ2: 中規模POC実行（初回フルスキャン）

約100ファイルを処理（想定時間: 50-70分）:

```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_poc.yaml --full
```

### ステップ3: 増分更新テスト

SharePoint上でファイルを変更後、増分更新を実行:

```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_poc.yaml --incremental
```

### ステップ4: POC結果確認

```bash
python -c "
from pathlib import Path
from src.utils.db_manager import ProcessedFilesDB

db = ProcessedFilesDB(Path('data/processed_files.db'))
stats = db.get_statistics()

print('=== POC処理結果 ===')
print(f\"総ファイル数: {stats['total_files']}\")
print(f\"成功: {stats['by_status'].get('success', 0)}\")
print(f\"失敗: {stats['by_status'].get('failed', 0)}\")
print(f\"総スライド数: {stats['total_slides']}\")
print(f\"平均処理時間: {stats['avg_processing_seconds']:.2f}秒/ファイル\")
print(f\"成功率: {stats['by_status'].get('success', 0) / stats['total_files'] * 100:.1f}%\")

# 失敗ファイル確認
failed = db.get_failed_files()
if failed:
    print(f\"\\n失敗ファイル ({len(failed)}件):\")
    for f in failed[:5]:
        print(f\"  - {f['file_name']}: {f['error_message']}\")
"
```

### 失敗ファイルの再試行

```bash
python -c "
from pathlib import Path
from src.utils.db_manager import ProcessedFilesDB

db = ProcessedFilesDB(Path('data/processed_files.db'))
count = db.reset_failed_files()
print(f'{count}件のファイルを再処理対象に設定しました')
"
```

## スケジューリング

### Windowsタスクスケジューラーで自動実行

1. **タスクスケジューラー** を起動
2. **基本タスクの作成**
3. **トリガー**: 毎日午前2時
4. **操作**: プログラムの開始
   - プログラム: `C:\path\to\venv\Scripts\python.exe`
   - 引数: `src\sharepoint_sync\sync_pipeline.py --config configs\sharepoint_prod.yaml --incremental`
   - 開始: `C:\path\to\pptx_RAG_search`

### バッチファイル作成（推奨）

**`scripts/daily_sync.bat`**:

```batch
@echo off
cd C:\path\to\pptx_RAG_search
call venv\Scripts\activate
python src\sharepoint_sync\sync_pipeline.py --config configs\sharepoint_prod.yaml --incremental

REM エラー時にメール通知（オプション）
if %ERRORLEVEL% NEQ 0 (
    echo "Sync failed" | mail -s "PPTX RAG Sync Failed" admin@company.com
)
```

## プロジェクト構造

```
pptx_RAG_search/
├── configs/
│   ├── sharepoint_template.yaml    # 設定テンプレート
│   └── sharepoint_prod.yaml        # 本番設定（作成必要）
├── data/
│   ├── pptx_temp/                  # 一時ダウンロードファイル
│   ├── logs/                       # ログファイル
│   ├── rendered/                   # レンダリング済みPNG
│   └── processed_files.db          # 処理状態データベース
├── doc/
│   ├── pptx_rag_local_poc_design.md        # ローカルPOC設計書
│   └── sharepoint_deployment_design.md     # SharePoint展開設計書
├── index/
│   └── qdrant_storage/             # Qdrantベクトルインデックス
├── src/
│   ├── sharepoint_sync/
│   │   ├── sharepoint_client.py    # SharePoint APIクライアント
│   │   └── sync_pipeline.py        # バッチ処理パイプライン
│   ├── utils/
│   │   └── db_manager.py           # 処理状態DB管理
│   └── ingest/                     # PPTX処理モジュール（今後実装）
├── requirements.txt                # Python依存パッケージ
└── README.md                       # 本ファイル
```

## トラブルシューティング

### エラー: "認証に失敗しました"

- Azure ADアプリの権限が正しく設定されているか確認
- テナント管理者による同意が必要な場合があります
- クライアントシークレットが有効か確認

### エラー: "PowerPoint COMエラー"

- PowerPointがインストールされているか確認
- PowerPointプロセスが残っていないか確認:
  ```bash
  taskkill /F /IM POWERPNT.EXE
  ```

### 処理が遅い

- `parallel_downloads` を増やす（デフォルト: 10）
- `batch_size` を調整（デフォルト: 50）
- CPUコア数に応じて並列処理数を調整

### メモリ不足

- `batch_size` を小さくする（例: 20）
- 大規模ファイルを除外するフィルタを追加

## 運用のベストプラクティス

### 1. 定期バックアップ

**Qdrantインデックス**:
```bash
tar -czf qdrant_backup_$(date +%Y%m%d).tar.gz index/qdrant_storage/
```

**処理状態データベース**:
```bash
sqlite3 data/processed_files.db ".backup 'data/backups/processed_files_$(date +%Y%m%d).db'"
```

### 2. ログ監視

```bash
# 最新のエラーを確認
grep ERROR data/logs/sync_pipeline.log | tail -20

# 失敗ファイル一覧
python -c "
from pathlib import Path
from src.utils.db_manager import ProcessedFilesDB
db = ProcessedFilesDB(Path('data/processed_files.db'))
for f in db.get_failed_files():
    print(f'{f[\"file_name\"]}: {f[\"error_message\"]}')
"
```

### 3. パフォーマンスチューニング

処理時間の統計を確認:
```bash
sqlite3 data/processed_files.db "
SELECT
    AVG(processing_duration_seconds) as avg,
    MIN(processing_duration_seconds) as min,
    MAX(processing_duration_seconds) as max
FROM processed_files
WHERE status = 'success'
"
```

## 今後の拡張

- [ ] リアルタイム検索WebUI
- [ ] 外部LLMによる回答生成
- [ ] マルチテナント対応
- [ ] 権限ベースのアクセス制御
- [ ] 並列処理の最適化（テキスト抽出・埋め込み計算）

## ライセンス

[ライセンス情報を記載]

## サポート

問題が発生した場合は、以下の情報とともにIssueを作成してください:

- エラーメッセージ
- `data/logs/sync_pipeline.log` の関連部分
- SharePoint環境情報（Online/Server、ファイル数規模）
- 実行したコマンド

---

**設計ドキュメント**:
- [ローカルPOC設計書](doc/pptx_rag_local_poc_design.md)
- [SharePoint展開設計書](doc/sharepoint_deployment_design.md)
