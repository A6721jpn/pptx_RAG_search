"""
Retriever for CLI search.
"""
import logging
from typing import List, Dict, Any
from rag.embedder import TextEmbedder
from rag.indexer import QdrantIndexer

logger = logging.getLogger(__name__)

class Retriever:
    """Retrieves slides matching a query."""
    
    def __init__(self, qdrant_path: str = "./index/qdrant_storage", collection_name: str = "pptx_slides"):
        self.embedder = TextEmbedder()
        self.indexer = QdrantIndexer(path=qdrant_path, collection_name=collection_name)
    
    def search(self, query: str, top_k: int = 5, score_threshold: float = 0.0) -> List[Dict[str, Any]]:
        """
        Search for slides matching query.
        
        Args:
            query: Search query text.
            top_k: Number of results.
            score_threshold: Minimum score to include.
            
        Returns:
            List of matching slide results.
        """
        logger.info(f"Searching for: '{query}'")
        
        query_vector = self.embedder.embed_query(query)
        results = self.indexer.search(query_vector, top_k=top_k)
        
        # Filter by score threshold
        filtered = [r for r in results if r["score"] >= score_threshold]
        
        logger.info(f"Found {len(filtered)} results above threshold {score_threshold}")
        return filtered
