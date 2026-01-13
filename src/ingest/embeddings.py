"""
テキスト埋め込み処理
sentence-transformersを使用したCPU実行
"""

import logging
from typing import List
import numpy as np

logger = logging.getLogger(__name__)


class TextEmbedder:
    """テキスト埋め込み生成器"""

    def __init__(self, model_name: str = "intfloat/e5-base-v2", device: str = "cpu"):
        """
        Args:
            model_name: 使用するモデル名
            device: 実行デバイス（"cpu" or "cuda"）
        """
        self.model_name = model_name
        self.device = device
        self.model = None

        logger.info(f"テキスト埋め込みモデル: {model_name}, デバイス: {device}")

    def load_model(self):
        """モデルをロード"""
        if self.model is not None:
            return

        try:
            from sentence_transformers import SentenceTransformer

            logger.info(f"モデルロード中: {self.model_name}")
            self.model = SentenceTransformer(self.model_name, device=self.device)
            logger.info("モデルロード完了")

        except ImportError:
            logger.error("sentence-transformersがインストールされていません")
            logger.error("pip install sentence-transformers")
            raise
        except Exception as e:
            logger.error(f"モデルロードエラー: {e}")
            raise

    def embed_texts(self, texts: List[str], batch_size: int = 32) -> np.ndarray:
        """
        テキストリストを埋め込みベクトルに変換

        Args:
            texts: テキストのリスト
            batch_size: バッチサイズ

        Returns:
            埋め込みベクトル配列 (N, D)
        """
        if not texts:
            return np.array([])

        # モデルロード（初回のみ）
        self.load_model()

        # e5モデルの場合、クエリプレフィックスを追加
        if "e5" in self.model_name.lower():
            texts = [f"passage: {text}" for text in texts]

        # 埋め込み計算
        logger.info(f"埋め込み計算中: {len(texts)}件のテキスト")

        embeddings = self.model.encode(
            texts,
            batch_size=batch_size,
            show_progress_bar=True,
            convert_to_numpy=True,
            normalize_embeddings=True  # コサイン類似度用に正規化
        )

        logger.info(f"埋め込み完了: shape={embeddings.shape}")
        return embeddings

    def embed_single(self, text: str) -> np.ndarray:
        """
        単一テキストを埋め込みベクトルに変換

        Args:
            text: テキスト

        Returns:
            埋め込みベクトル (D,)
        """
        embeddings = self.embed_texts([text], batch_size=1)
        return embeddings[0]

    def get_dimension(self) -> int:
        """埋め込みベクトルの次元数を取得"""
        self.load_model()
        return self.model.get_sentence_embedding_dimension()


# 使用例
if __name__ == "__main__":
    # ロギング設定
    logging.basicConfig(level=logging.INFO)

    # エンベッダー初期化
    embedder = TextEmbedder()

    # テキスト埋め込み
    texts = [
        "This is a mechanical design guide for hinges.",
        "Tolerance stack-up analysis for assembly clearance.",
    ]

    embeddings = embedder.embed_texts(texts)
    print(f"Embeddings shape: {embeddings.shape}")
    print(f"Dimension: {embedder.get_dimension()}")
