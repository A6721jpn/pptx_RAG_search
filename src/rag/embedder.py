"""
Text Embedder using sentence-transformers.
"""
import logging
from typing import List
from sentence_transformers import SentenceTransformer

logger = logging.getLogger(__name__)

class TextEmbedder:
    """Generates text embeddings using sentence-transformers."""
    
    def __init__(self, model_name: str = "intfloat/e5-base-v2"):
        """
        Args:
            model_name: HuggingFace model name for embeddings.
        """
        logger.info(f"Loading embedding model: {model_name}")
        self.model = SentenceTransformer(model_name)
        self.dimension = self.model.get_sentence_embedding_dimension()
        logger.info(f"Model loaded. Embedding dimension: {self.dimension}")
    
    def embed(self, texts: List[str]) -> List[List[float]]:
        """
        Embed a list of texts.
        
        Args:
            texts: List of text strings to embed.
            
        Returns:
            List of embedding vectors.
        """
        # e5 models expect "query: " or "passage: " prefix
        # For indexing, use "passage: " prefix
        prefixed = [f"passage: {t}" for t in texts]
        embeddings = self.model.encode(prefixed, convert_to_numpy=True)
        return embeddings.tolist()
    
    def embed_query(self, query: str) -> List[float]:
        """
        Embed a single query text.
        
        Args:
            query: Query string.
            
        Returns:
            Embedding vector.
        """
        # e5 models expect "query: " prefix for queries
        prefixed = f"query: {query}"
        embedding = self.model.encode([prefixed], convert_to_numpy=True)
        return embedding[0].tolist()
