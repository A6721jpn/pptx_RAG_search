# PPTX RAG Search - SharePoint大規模展開

SharePoint上の大量のPowerPointファイルに対する検索可能なRAG（Retrieval-Augmented Generation）システム

## 概要

このプロジェクトは、SharePoint上に保存された数百〜数千のPowerPointファイルを対象に、テキストとビジュアルの両方を活用した高度な検索システムを提供します。

### 主な機能

- **SharePoint統合**: Microsoft Graph APIを使用したシームレスな連携
- **マルチモーダル検索**: テキストと画像の両方から検索
- **増分更新**: 変更されたファイルのみを効率的に処理
- **大規模対応**: バッチ処理と並列化による高スループット
- **自動スケジューリング**: 定期的な自動同期

### アーキテクチャ

```
SharePoint Online
    ↓ (Microsoft Graph API)
専用Windowsサーバー
    - SharePoint同期モジュール
    - PPTX処理パイプライン (COM)
    - Qdrantローカルインスタンス
    - スケジューラー
```

## セットアップ

### 前提条件

- **OS**: Windows 10/11 または Windows Server 2019+
- **PowerPoint**: Microsoft PowerPoint（Office 365 または 2019+）
- **Python**: 3.9+
- **権限**: Azure AD アプリ登録権限

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
cp configs/sharepoint_template.yaml configs/sharepoint_prod.yaml

# エディタで編集
notepad configs/sharepoint_prod.yaml
```

**`configs/sharepoint_prod.yaml`** を編集:

```yaml
sharepoint:
  tenant_id: "あなたのテナントID"
  client_id: "あなたのクライアントID"
  client_secret: "あなたのクライアントシークレット"
  site_urls:
    - "https://yourcompany.sharepoint.com/sites/Engineering"
```

### 4. Qdrant起動

**オプション1: Dockerで起動（推奨）**

```bash
docker run -d -p 6333:6333 -v $(pwd)/index/qdrant_storage:/qdrant/storage qdrant/qdrant
```

**オプション2: ローカルモード**

Qdrantはファイルベースでも動作します（設定不要）

## 使い方

### 初回フルスキャン

```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_prod.yaml --full
```

### 増分更新（変更されたファイルのみ）

```bash
python src/sharepoint_sync/sync_pipeline.py --config configs/sharepoint_prod.yaml --incremental
```

### 処理状態確認

```bash
python -c "
from pathlib import Path
from src.utils.db_manager import ProcessedFilesDB

db = ProcessedFilesDB(Path('data/processed_files.db'))
stats = db.get_statistics()

print('総ファイル数:', stats['total_files'])
print('総スライド数:', stats['total_slides'])
print('状態別件数:', stats['by_status'])
print('平均処理時間:', f\"{stats['avg_processing_seconds']:.2f}秒\")
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
