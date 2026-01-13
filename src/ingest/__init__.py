"""
Ingest module for PPTX RAG system
"""

from .embeddings import TextEmbedder
from .indexer import QdrantIndexer

__all__ = ['TextEmbedder', 'QdrantIndexer']
