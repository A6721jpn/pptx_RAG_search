"""
SharePoint接続テストスクリプト
POC実行前にSharePoint接続が正常に動作するか確認
"""

import asyncio
from pathlib import Path
import yaml
import sys

# パスを追加
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from sharepoint_sync.sharepoint_client import SharePointClient


async def test_connection():
    """SharePoint接続テスト"""

    # 設定ファイルのパスを確認
    config_path = Path('configs/sharepoint_poc.yaml')

    if not config_path.exists():
        print("❌ 設定ファイルが見つかりません: configs/sharepoint_poc.yaml")
        print("\n以下のコマンドで設定ファイルを作成してください:")
        print("  cp configs/sharepoint_template.yaml configs/sharepoint_poc.yaml")
        print("  notepad configs/sharepoint_poc.yaml")
        return False

    # 設定読み込み
    with open(config_path) as f:
        config = yaml.safe_load(f)

    print("=== SharePoint接続テスト ===\n")

    # クライアント作成
    client = SharePointClient(
        tenant_id=config['sharepoint']['tenant_id'],
        client_id=config['sharepoint']['client_id'],
        client_secret=config['sharepoint']['client_secret']
    )

    try:
        # サイトURL取得
        site_urls = config['sharepoint']['site_urls']
        if not site_urls:
            print("❌ site_urlsが設定されていません")
            return False

        site_url = site_urls[0]
        print(f"接続先: {site_url}\n")

        # 1. サイトID取得
        print("1. SharePointサイトへの接続...")
        try:
            site_id = await client.get_site_id(site_url)
            print(f"   ✅ 接続成功")
            print(f"   サイトID: {site_id}\n")
        except Exception as e:
            print(f"   ❌ 接続失敗: {e}")
            print("\n確認事項:")
            print("  - tenant_id, client_id, client_secretが正しいか")
            print("  - Azure ADアプリに権限が付与されているか")
            print("  - 管理者の同意が完了しているか")
            return False

        # 2. ドライブID取得
        print("2. ドキュメントライブラリへのアクセス...")
        try:
            drive_id = await client.get_drive_id(site_id)
            print(f"   ✅ アクセス成功")
            print(f"   ドライブID: {drive_id}\n")
        except Exception as e:
            print(f"   ❌ アクセス失敗: {e}")
            return False

        # 3. PPTXファイル一覧取得
        print("3. PPTXファイルの検出...")
        try:
            files = await client.list_pptx_files(site_id, drive_id)
            print(f"   ✅ 検出成功: {len(files)}件のPPTXファイル\n")

            if len(files) == 0:
                print("   ⚠️  PPTXファイルが見つかりませんでした")
                print("   SharePointサイトにPPTXファイルをアップロードしてください\n")
            else:
                print("   検出されたファイル（最初の10件）:")
                for i, f in enumerate(files[:10], 1):
                    size_mb = f['size'] / (1024 * 1024)
                    print(f"     {i}. {f['name']} ({size_mb:.2f} MB)")

                if len(files) > 10:
                    print(f"     ... 他 {len(files) - 10}件\n")
                else:
                    print()

                # 統計情報
                total_size = sum(f['size'] for f in files)
                avg_size = total_size / len(files)
                print(f"   統計:")
                print(f"     総ファイル数: {len(files)}")
                print(f"     総サイズ: {total_size / (1024 * 1024):.2f} MB")
                print(f"     平均ファイルサイズ: {avg_size / (1024 * 1024):.2f} MB\n")

                # POC想定時間の見積もり
                estimated_time = len(files) * 30  # 1ファイル30秒想定
                print(f"   POC想定処理時間:")
                print(f"     初回フルスキャン: 約{estimated_time // 60}分")
                print(f"     （1ファイルあたり30秒で計算）\n")

        except Exception as e:
            print(f"   ❌ 検出失敗: {e}")
            return False

        print("="*50)
        print("✅ すべてのテストが成功しました！")
        print("\n次のコマンドでPOCを実行できます:")
        print(f"  python src/sharepoint_sync/sync_pipeline.py --config {config_path} --full")
        print("="*50)

        return True

    finally:
        await client.close()


def main():
    """メイン関数"""
    try:
        success = asyncio.run(test_connection())
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\nテストが中断されました")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
