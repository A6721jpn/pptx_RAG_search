"""
Qdrantインデックス化
ローカルQdrantインスタンスへのベクトル保存
"""

import logging
from typing import List, Dict
from pathlib import Path
import numpy as np
from datetime import datetime

logger = logging.getLogger(__name__)


class QdrantIndexer:
    """Qdrantインデックス管理"""

    def __init__(
        self,
        storage_path: str = "index/qdrant_storage",
        collection_name: str = "pptx_pages"
    ):
        """
        Args:
            storage_path: Qdrantストレージパス
            collection_name: コレクション名
        """
        self.storage_path = Path(storage_path)
        self.collection_name = collection_name
        self.client = None
        self.vector_dimension = None

        logger.info(f"Qdrantストレージ: {storage_path}, コレクション: {collection_name}")

    def initialize(self, vector_dimension: int):
        """
        Qdrantクライアントとコレクションを初期化

        Args:
            vector_dimension: ベクトル次元数
        """
        try:
            from qdrant_client import QdrantClient
            from qdrant_client.models import Distance, VectorParams, PointStruct

            # ストレージディレクトリ作成
            self.storage_path.mkdir(parents=True, exist_ok=True)

            # クライアント初期化
            logger.info("Qdrantクライアント初期化中...")
            self.client = QdrantClient(path=str(self.storage_path))

            self.vector_dimension = vector_dimension

            # コレクション存在確認
            collections = self.client.get_collections().collections
            collection_exists = any(c.name == self.collection_name for c in collections)

            if not collection_exists:
                # コレクション作成
                logger.info(f"コレクション作成中: {self.collection_name}")
                self.client.create_collection(
                    collection_name=self.collection_name,
                    vectors_config=VectorParams(
                        size=vector_dimension,
                        distance=Distance.COSINE
                    )
                )
                logger.info("コレクション作成完了")
            else:
                logger.info(f"既存コレクションを使用: {self.collection_name}")

        except ImportError:
            logger.error("qdrant-clientがインストールされていません")
            logger.error("pip install qdrant-client")
            raise
        except Exception as e:
            logger.error(f"Qdrant初期化エラー: {e}")
            raise

    def delete_document(self, doc_id: str):
        """
        ドキュメントIDに紐づくすべてのポイントを削除

        Args:
            doc_id: ドキュメントID
        """
        from qdrant_client.models import Filter, FieldCondition, MatchValue

        logger.info(f"既存ドキュメント削除: {doc_id}")

        self.client.delete(
            collection_name=self.collection_name,
            points_selector=Filter(
                must=[
                    FieldCondition(
                        key="doc_id",
                        match=MatchValue(value=doc_id)
                    )
                ]
            )
        )

    def index_pages(
        self,
        doc_id: str,
        file_name: str,
        pages: List[Dict],
        embeddings: np.ndarray
    ):
        """
        ページをインデックス化

        Args:
            doc_id: ドキュメントID
            file_name: ファイル名
            pages: ページ情報のリスト [{page_num, text, image_path}, ...]
            embeddings: 埋め込みベクトル配列 (N, D)
        """
        from qdrant_client.models import PointStruct

        if len(pages) != len(embeddings):
            raise ValueError(f"ページ数({len(pages)})と埋め込み数({len(embeddings)})が一致しません")

        # 既存ドキュメントを削除
        self.delete_document(doc_id)

        # ポイント作成
        points = []
        for i, (page_info, embedding) in enumerate(zip(pages, embeddings)):
            point_id = f"{doc_id}_{page_info['page_num']}"

            payload = {
                "doc_id": doc_id,
                "file_name": file_name,
                "page_num": page_info['page_num'],
                "text": page_info['text'],
                "image_path": str(page_info.get('image_path', '')),
                "indexed_at": datetime.now().isoformat()
            }

            point = PointStruct(
                id=point_id,
                vector=embedding.tolist(),
                payload=payload
            )
            points.append(point)

        # バッチアップロード
        logger.info(f"インデックス化中: {len(points)}ページ")
        self.client.upsert(
            collection_name=self.collection_name,
            points=points
        )
        logger.info(f"インデックス化完了: {doc_id}")

    def search(
        self,
        query_vector: np.ndarray,
        top_k: int = 5,
        score_threshold: float = 0.0
    ) -> List[Dict]:
        """
        ベクトル検索

        Args:
            query_vector: クエリベクトル
            top_k: 取得件数
            score_threshold: スコア閾値

        Returns:
            検索結果のリスト
        """
        logger.info(f"検索実行: top_k={top_k}")

        results = self.client.search(
            collection_name=self.collection_name,
            query_vector=query_vector.tolist(),
            limit=top_k,
            score_threshold=score_threshold
        )

        # 結果を整形
        formatted_results = []
        for hit in results:
            formatted_results.append({
                "id": hit.id,
                "score": hit.score,
                "doc_id": hit.payload.get("doc_id"),
                "file_name": hit.payload.get("file_name"),
                "page_num": hit.payload.get("page_num"),
                "text": hit.payload.get("text"),
                "image_path": hit.payload.get("image_path")
            })

        logger.info(f"検索結果: {len(formatted_results)}件")
        return formatted_results

    def get_collection_info(self) -> Dict:
        """コレクション情報を取得"""
        info = self.client.get_collection(self.collection_name)

        # Qdrant APIバージョンによって属性が異なるため安全に取得
        result = {
            "collection_name": self.collection_name,
            "points_count": info.points_count if hasattr(info, 'points_count') else 0
        }

        # オプショナルな属性を安全に取得
        if hasattr(info, 'vectors_count'):
            result["vectors_count"] = info.vectors_count
        if hasattr(info, 'indexed_vectors_count'):
            result["indexed_vectors_count"] = info.indexed_vectors_count

        return result


# 使用例
if __name__ == "__main__":
    # ロギング設定
    logging.basicConfig(level=logging.INFO)

    # インデクサー初期化
    indexer = QdrantIndexer()
    indexer.initialize(vector_dimension=768)

    # ダミーデータでテスト
    import numpy as np

    pages = [
        {"page_num": 1, "text": "Page 1 content", "image_path": "page_1.png"},
        {"page_num": 2, "text": "Page 2 content", "image_path": "page_2.png"}
    ]

    embeddings = np.random.rand(2, 768)  # ダミー埋め込み

    indexer.index_pages(
        doc_id="test_doc_001",
        file_name="test.pdf",
        pages=pages,
        embeddings=embeddings
    )

    # 検索テスト
    query_vector = np.random.rand(768)
    results = indexer.search(query_vector, top_k=5)

    for result in results:
        print(f"Score: {result['score']:.4f}, Page: {result['page_num']}, Text: {result['text'][:50]}")

    # コレクション情報
    info = indexer.get_collection_info()
    print(f"\nCollection info: {info}")
