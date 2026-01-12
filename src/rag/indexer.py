"""
Qdrant Indexer for storing slide embeddings.
"""
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams, PointStruct

logger = logging.getLogger(__name__)

class QdrantIndexer:
    """Manages Qdrant collection for slide embeddings."""
    
    def __init__(self, path: str = "./index/qdrant_storage", collection_name: str = "pptx_slides"):
        """
        Args:
            path: Path to Qdrant storage (local mode).
            collection_name: Name of the collection.
        """
        self.path = Path(path)
        self.path.mkdir(parents=True, exist_ok=True)
        self.collection_name = collection_name
        
        logger.info(f"Initializing Qdrant client at {self.path}")
        self.client = QdrantClient(path=str(self.path))
    
    def ensure_collection(self, vector_size: int):
        """Create collection if it doesn't exist."""
        collections = self.client.get_collections().collections
        exists = any(c.name == self.collection_name for c in collections)
        
        if not exists:
            logger.info(f"Creating collection '{self.collection_name}' with vector size {vector_size}")
            self.client.create_collection(
                collection_name=self.collection_name,
                vectors_config=VectorParams(size=vector_size, distance=Distance.COSINE)
            )
        else:
            logger.info(f"Collection '{self.collection_name}' already exists.")
    
    def index_slides(self, slides: List[Dict[str, Any]], embeddings: List[List[float]]):
        """
        Index slides with their embeddings.
        
        Args:
            slides: List of slide data dicts (file_path, slide_no, text_raw, notes_raw, thumb_path).
            embeddings: List of embedding vectors corresponding to slides.
        """
        if len(slides) != len(embeddings):
            raise ValueError(f"Mismatch: {len(slides)} slides vs {len(embeddings)} embeddings")
        
        points = []
        for i, (slide, embedding) in enumerate(zip(slides, embeddings)):
            # Generate unique ID from file_path and slide_no
            point_id = hash(f"{slide['file_path']}:{slide['slide_no']}") & 0x7FFFFFFFFFFFFFFF
            
            payload = {
                "file_path": slide["file_path"],
                "file_name": Path(slide["file_path"]).name,
                "slide_no": slide["slide_no"],
                "text_raw": slide.get("text_raw", ""),
                "notes_raw": slide.get("notes_raw", ""),
                "thumb_path": slide.get("thumb_path", "")
            }
            
            points.append(PointStruct(id=point_id, vector=embedding, payload=payload))
        
        logger.info(f"Indexing {len(points)} slides to Qdrant")
        self.client.upsert(collection_name=self.collection_name, points=points)
        logger.info("Indexing complete.")
    
    def search(self, query_vector: List[float], top_k: int = 5) -> List[Dict[str, Any]]:
        """
        Search for similar slides.
        
        Args:
            query_vector: Query embedding vector.
            top_k: Number of results to return.
            
        Returns:
            List of result dicts with score and payload.
        """
        results = self.client.query_points(
            collection_name=self.collection_name,
            query=query_vector,
            limit=top_k
        ).points
        
        return [
            {
                "score": hit.score,
                "file_name": hit.payload.get("file_name"),
                "file_path": hit.payload.get("file_path"),
                "slide_no": hit.payload.get("slide_no"),
                "text_raw": hit.payload.get("text_raw"),
                "notes_raw": hit.payload.get("notes_raw"),
                "thumb_path": hit.payload.get("thumb_path")
            }
            for hit in results
        ]
