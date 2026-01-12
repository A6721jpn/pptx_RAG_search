# Power Automate: PPTX→PDF自動変換フロー設計

**目的**: SharePoint上のPPTXファイルを自動的にPDFに変換し、OneDrive同期フォルダに保存

---

## 前提条件

- Power Automate Premium ライセンス（または会社のライセンス）
- SharePointサイトへのアクセス権限
- OneDriveアカウント

---

## フロー設計

### トリガー: ファイルが作成または変更されたとき

**設定:**
```
コネクタ: SharePoint
トリガー: ファイルが作成または変更されたとき (プロパティのみ)
サイトのアドレス: [SharePointサイトURL]
ライブラリ名: Documents
フォルダー: [対象フォルダを指定、または空で全体]
```

**フィルタ条件（高度な設定）:**
```
ファイル拡張子が次の値に等しい: .pptx
または
ファイル拡張子が次の値に等しい: .ppt
```

---

### アクション1: ファイルコンテンツの取得

**設定:**
```
コネクタ: SharePoint
アクション: ファイル コンテンツの取得
サイトのアドレス: [SharePointサイトURL]
ファイル識別子: [トリガーからの動的コンテンツ: ファイル識別子]
```

---

### アクション2: PDFへの変換

**方法A: PowerPoint Online コネクタ（推奨）**

```
コネクタ: PowerPoint Online (Business)
アクション: プレゼンテーションをPDFに変換
ファイル: [アクション1からの動的コンテンツ: ファイル コンテンツ]
```

**方法B: OneDrive for Business（代替）**

```
1. OneDrive for Businessにファイルを作成
   - ファイル名: [ファイル名（拡張子なし）].pptx
   - ファイルコンテンツ: [アクション1からのファイルコンテンツ]

2. OneDriveファイルをPDFに変換
   - ファイル: [前のステップで作成したファイルID]

3. 元のPPTXファイルを削除（オプション）
```

---

### アクション3: PDFファイルの保存

**保存先A: SharePoint（同じサイトの別フォルダ）**

```
コネクタ: SharePoint
アクション: ファイルの作成
サイトのアドレス: [SharePointサイトURL]
フォルダーのパス: /PDF_Converted  （新規フォルダを作成）
ファイル名: [トリガーからの動的コンテンツ: ファイル名（拡張子なし）].pdf
ファイル コンテンツ: [アクション2からの動的コンテンツ: PDF コンテンツ]
```

**保存先B: OneDrive for Business**

```
コネクタ: OneDrive for Business
アクション: ファイルの作成
フォルダーのパス: /PDF_Converted
ファイル名: [トリガーからの動的コンテンツ: ファイル名（拡張子なし）].pdf
ファイル コンテンツ: [アクション2からの動的コンテンツ: PDF コンテンツ]
```

---

### アクション4: メタデータの記録（オプション）

**SharePointリストに変換履歴を記録:**

```
コネクタ: SharePoint
アクション: 項目の作成
サイトのアドレス: [SharePointサイトURL]
リスト名: PDF_Conversion_Log  （事前に作成）

フィールド:
- 元ファイル名: [トリガー: ファイル名]
- 元ファイルURL: [トリガー: ファイルのリンク]
- PDFファイル名: [生成されたPDFファイル名]
- 変換日時: [utcNow()]
- ステータス: 成功
```

---

## フロー全体図

```
[SharePointにPPTX追加]
    ↓
[トリガー: ファイルが作成/変更]
    ↓
[条件: 拡張子が.pptx/.ppt?]
    ↓ Yes
[ファイルコンテンツ取得]
    ↓
[PowerPoint Online: PDFに変換]
    ↓
[SharePoint: PDF保存 (/PDF_Converted)]
    ↓
[SharePointリスト: ログ記録]
    ↓
[完了]
```

---

## エラーハンドリング

### 並列分岐の追加

```
メインフロー
    ↓
[スコープ: 変換処理]
    ├─ 成功時 → ログ記録
    └─ 失敗時 → エラー通知
```

**エラー時の通知:**

```
コネクタ: Office 365 Outlook（またはTeams）
アクション: メールの送信
宛先: [管理者メールアドレス]
件名: PDF変換失敗: @{triggerOutputs()?['body/Name']}
本文:
  ファイル名: @{triggerOutputs()?['body/Name']}
  エラー: @{body('PDFへの変換')?['error']?['message']}
  時刻: @{utcNow()}
```

---

## 既存ファイルの一括変換

### 手動フロー（既存PPTXを一括変換）

```
トリガー: 手動でフローをトリガー（ボタン）
    ↓
[SharePoint: 複数のファイルを取得]
    - サイト: [SharePointサイトURL]
    - ライブラリ: Documents
    - フィルタークエリ: FileLeafRef like '%.pptx' or FileLeafRef like '%.ppt'
    ↓
[Apply to each: 各ファイルに対して]
    ↓ （各アイテムで以下を実行）
    [ファイルコンテンツ取得]
        ↓
    [PDFに変換]
        ↓
    [PDFを保存]
```

**注意**: 大量ファイルの場合、Power Automateの実行制限に注意（1日のアクション数など）

---

## OneDrive同期の設定

PDFフォルダをOneDriveで同期：

1. SharePointの `/PDF_Converted` フォルダを開く
2. 上部の「**同期**」ボタンをクリック
3. OneDriveクライアントで同期開始
4. 同期先パスをメモ：
   ```
   C:\Users\[ユーザー名]\OneDrive - [会社名]\[サイト名] - Documents\PDF_Converted
   ```

---

## POCスクリプトでの利用

同期されたPDFフォルダを指定：

```bash
python local_poc_pdf.py --source "C:\Users\...\OneDrive\...\PDF_Converted" --full
```

---

## コスト・制限事項

### Power Automate制限

- **無料プラン**: 750実行/月、15分タイムアウト
- **プレミアムプラン**: 無制限実行、処理時間制限あり

### 推奨設定

- バッチ処理: 大量ファイルは分割して処理
- スケジュール: 夜間や週末に一括変換
- 監視: 失敗通知の設定

---

## トラブルシューティング

### エラー: "変換に失敗しました"

**原因:**
- PPTXファイルが破損している
- ファイルサイズが大きすぎる（100MB超）
- 特殊なフォントや埋め込みオブジェクト

**対処:**
- 手動でPowerPointを開いて修復
- ファイルサイズを縮小
- 該当ファイルをスキップ

### エラー: "権限がありません"

**原因:**
- SharePointサイトへのアクセス権限不足
- OneDriveフォルダへの書き込み権限不足

**対処:**
- SharePoint管理者に権限付与を依頼
- フロー接続の再認証

---

## 次のステップ

1. Power Automateにサインイン: https://make.powerautomate.com
2. 「作成」→「自動化したクラウドフロー」を選択
3. 上記の設計に従ってフローを構築
4. テスト実行（1ファイルで確認）
5. 本番運用開始

---

End of document.
