"""
SharePoint Online統合クライアント
Microsoft Graph APIを使用してSharePointからPPTXファイルを取得
"""

import asyncio
import aiohttp
import aiofiles
from typing import List, Dict, Optional
from datetime import datetime
from pathlib import Path
import logging

from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

logger = logging.getLogger(__name__)


class SharePointClient:
    """SharePoint Onlineクライアント"""

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        scopes: Optional[List[str]] = None
    ):
        """
        Args:
            tenant_id: Azure AD テナントID
            client_id: アプリケーション（クライアント）ID
            client_secret: クライアントシークレット
            scopes: アクセススコープ（デフォルト: ['https://graph.microsoft.com/.default']）
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scopes = scopes or ['https://graph.microsoft.com/.default']

        # 認証情報
        self.credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )

        # Graph クライアント
        self.graph_client = GraphServiceClient(
            credentials=self.credential,
            scopes=self.scopes
        )

    async def get_site_id(self, site_url: str) -> str:
        """
        SharePointサイトURLからサイトIDを取得

        Args:
            site_url: SharePointサイトURL
                例: "https://company.sharepoint.com/sites/Engineering"

        Returns:
            サイトID
        """
        try:
            # URLからホスト名とサイトパスを抽出
            from urllib.parse import urlparse
            parsed = urlparse(site_url)
            hostname = parsed.netloc
            site_path = parsed.path

            # Graph API: /sites/{hostname}:{site_path}
            site = await self.graph_client.sites.by_site_id(
                f"{hostname}:{site_path}"
            ).get()

            logger.info(f"Site ID取得成功: {site.id}")
            return site.id

        except ODataError as e:
            logger.error(f"Site ID取得失敗: {e.error.message}")
            raise

    async def get_drive_id(self, site_id: str, drive_name: str = "Documents") -> str:
        """
        SharePointサイトのドキュメントライブラリIDを取得

        Args:
            site_id: サイトID
            drive_name: ドキュメントライブラリ名（デフォルト: "Documents"）

        Returns:
            ドライブID
        """
        try:
            drives = await self.graph_client.sites.by_site_id(site_id).drives.get()

            for drive in drives.value:
                if drive.name == drive_name:
                    logger.info(f"Drive ID取得成功: {drive.id} (名前: {drive_name})")
                    return drive.id

            raise ValueError(f"ドライブ '{drive_name}' が見つかりません")

        except ODataError as e:
            logger.error(f"Drive ID取得失敗: {e.error.message}")
            raise

    async def list_pptx_files(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = ""
    ) -> List[Dict]:
        """
        SharePointドライブからすべてのPPTXファイルを再帰的に取得

        Args:
            site_id: サイトID
            drive_id: ドライブID
            folder_path: フォルダパス（ルートからの相対パス、例: "Design Guides"）

        Returns:
            PPTXファイル情報のリスト
        """
        pptx_files = []

        try:
            if folder_path:
                # 特定フォルダ内のアイテムを取得
                items = await self.graph_client.sites.by_site_id(site_id)\
                    .drives.by_drive_id(drive_id)\
                    .root.item_with_path(folder_path)\
                    .children.get()
            else:
                # ルートフォルダのアイテムを取得
                items = await self.graph_client.sites.by_site_id(site_id)\
                    .drives.by_drive_id(drive_id)\
                    .root.children.get()

            for item in items.value:
                # フォルダの場合は再帰的に探索
                if item.folder:
                    subfolder_path = f"{folder_path}/{item.name}" if folder_path else item.name
                    sub_files = await self.list_pptx_files(
                        site_id, drive_id, subfolder_path
                    )
                    pptx_files.extend(sub_files)

                # PPTXファイルの場合は情報を保存
                elif item.file and item.name.lower().endswith(('.pptx', '.ppt')):
                    file_info = {
                        'id': item.id,
                        'name': item.name,
                        'web_url': item.web_url,
                        'download_url': item.additional_data.get('@microsoft.graph.downloadUrl'),
                        'modified': item.last_modified_date_time,
                        'size': item.size,
                        'path': f"{folder_path}/{item.name}" if folder_path else item.name,
                        'site_id': site_id,
                        'drive_id': drive_id
                    }
                    pptx_files.append(file_info)

            logger.info(f"検出ファイル数: {len(pptx_files)} (フォルダ: {folder_path or 'root'})")
            return pptx_files

        except ODataError as e:
            logger.error(f"ファイル一覧取得失敗: {e.error.message}")
            raise

    async def search_pptx_files(self, site_id: str, query: str = "*.pptx") -> List[Dict]:
        """
        SharePointサイト全体からPPTXファイルを検索

        Args:
            site_id: サイトID
            query: 検索クエリ（デフォルト: "*.pptx"）

        Returns:
            PPTXファイル情報のリスト
        """
        try:
            from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody
            from msgraph.generated.models.search_request import SearchRequest
            from msgraph.generated.models.search_query import SearchQuery

            # 検索リクエスト構築
            request_body = QueryPostRequestBody(
                requests=[
                    SearchRequest(
                        entity_types=["driveItem"],
                        query=SearchQuery(
                            query_string=f"(fileExtension:pptx OR fileExtension:ppt) AND path:{site_id}"
                        ),
                        size=1000
                    )
                ]
            )

            # 検索実行
            results = await self.graph_client.search.query.post(request_body)

            pptx_files = []
            for result_set in results.value:
                for hit in result_set.hits_containers[0].hits:
                    resource = hit.resource
                    if hasattr(resource, 'name'):
                        file_info = {
                            'id': resource.id,
                            'name': resource.name,
                            'web_url': resource.web_url,
                            'download_url': resource.additional_data.get('@microsoft.graph.downloadUrl'),
                            'modified': resource.last_modified_date_time,
                            'size': resource.size,
                            'path': resource.additional_data.get('path', ''),
                            'site_id': site_id
                        }
                        pptx_files.append(file_info)

            logger.info(f"検索結果: {len(pptx_files)}件")
            return pptx_files

        except ODataError as e:
            logger.error(f"検索失敗: {e.error.message}")
            raise

    async def download_file(
        self,
        download_url: str,
        local_path: Path,
        chunk_size: int = 8192
    ) -> Path:
        """
        SharePointからファイルをダウンロード

        Args:
            download_url: ダウンロードURL
            local_path: 保存先ローカルパス
            chunk_size: ダウンロードチャンクサイズ（バイト）

        Returns:
            保存されたファイルのパス
        """
        try:
            # 親ディレクトリを作成
            local_path.parent.mkdir(parents=True, exist_ok=True)

            # ダウンロード
            async with aiohttp.ClientSession() as session:
                async with session.get(download_url) as resp:
                    if resp.status != 200:
                        raise Exception(f"ダウンロード失敗: HTTP {resp.status}")

                    async with aiofiles.open(local_path, 'wb') as f:
                        async for chunk in resp.content.iter_chunked(chunk_size):
                            await f.write(chunk)

            logger.info(f"ダウンロード完了: {local_path}")
            return local_path

        except Exception as e:
            logger.error(f"ダウンロードエラー ({local_path}): {e}")
            raise

    async def close(self):
        """リソースのクリーンアップ"""
        await self.credential.close()


# ========== 使用例 ==========

async def main_example():
    """使用例"""
    # SharePointクライアント初期化
    client = SharePointClient(
        tenant_id="YOUR_TENANT_ID",
        client_id="YOUR_CLIENT_ID",
        client_secret="YOUR_CLIENT_SECRET"
    )

    try:
        # サイトIDを取得
        site_url = "https://company.sharepoint.com/sites/Engineering"
        site_id = await client.get_site_id(site_url)

        # ドライブIDを取得
        drive_id = await client.get_drive_id(site_id, drive_name="Documents")

        # PPTXファイル一覧を取得
        pptx_files = await client.list_pptx_files(
            site_id=site_id,
            drive_id=drive_id,
            folder_path="Design Guides"
        )

        print(f"検出ファイル数: {len(pptx_files)}")

        # 最初の5ファイルをダウンロード
        for file_info in pptx_files[:5]:
            local_path = Path(f"data/pptx_temp/{file_info['id']}.pptx")
            await client.download_file(file_info['download_url'], local_path)

    finally:
        await client.close()


if __name__ == "__main__":
    # ロギング設定
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # 実行
    asyncio.run(main_example())
