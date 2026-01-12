import logging
from pathlib import Path
from typing import Dict, Any, List
from utils.db_manager import ProcessedFilesDB
from ingest.pptx_extract import PptxExtractor
from ingest.pptx_render import PptxRenderer
from rag.embedder import TextEmbedder
from rag.indexer import QdrantIndexer

logger = logging.getLogger(__name__)

class IngestPipeline:
    """
    Main pipeline for ingesting local PPTX files.
    """
    def __init__(self, config: Dict[str, Any]):
        """
        Args:
            config: Configuration dictionary. 
                    Must contain 'data_dir', 'output_dir', 'db_path'.
        """
        self.data_dir = Path(config.get('data_dir', './data/pptx'))
        self.output_dir = Path(config.get('output_dir', './data'))
        self.db_path = Path(config.get('db_path', './data/processed_files.db'))
        qdrant_path = config.get('qdrant_path', './index/qdrant_storage')
        collection_name = config.get('collection_name', 'pptx_slides')
        
        self.db = ProcessedFilesDB(self.db_path)
        self.extractor = PptxExtractor()
        self.renderer = PptxRenderer()
        self.embedder = TextEmbedder()
        self.indexer = QdrantIndexer(path=qdrant_path, collection_name=collection_name)
        
        # Ensure collection exists
        self.indexer.ensure_collection(self.embedder.dimension)
        
    def run(self):
        """Run the ingestion pipeline on the configured directory."""
        logger.info(f"Starting ingestion from {self.data_dir}")
        self.data_dir.mkdir(parents=True, exist_ok=True)
        
        count = 0
        # Recursive glob for .pptx
        for file_path in self.data_dir.rglob("*.pptx"):
            # Skip temp files (starting with ~$)
            if file_path.name.startswith("~$"):
                continue
                
            self.process_file(file_path)
            count += 1
            
        logger.info(f"Ingestion scan finished. Scanned {count} files.")

    def process_file(self, file_path: Path):
        """
        Process a single PPTX file: Extract -> Render -> Embed -> Index.
        """
        str_path = str(file_path.resolve())
        
        try:
            stat = file_path.stat()
            file_size = stat.st_size
            mtime = stat.st_mtime
            
            # Simple pseudo-hash from metadata for speed
            file_hash = f"{file_size}-{mtime}"
            
            # Check if processing is needed
            if not self.db.should_process(str_path, file_hash, mtime):
                logger.debug(f"Skipping unchanged file: {file_path.name}")
                return

            logger.info(f"Processing file: {file_path.name}")
            
            # Register as pending
            self.db.register_file(str_path, file_hash, file_size, mtime)
            self.db.update_status(str_path, "processing")
            
            # 1. Extract Text
            logger.info(f"Extracting text: {file_path.name}")
            slides_data = self.extractor.extract(file_path)
            
            # 2. Render Slides
            render_dir = self.output_dir / "rendered" / file_path.stem
            logger.info(f"Rendering slides to: {render_dir}")
            image_paths = self.renderer.render(file_path, render_dir)
            
            # Verify counts match
            if len(slides_data) != len(image_paths):
                logger.warning(f"Mismatch in count: {len(slides_data)} slides text, {len(image_paths)} images.")
            
            # 3. Prepare slides for indexing
            slides_for_index = []
            for i, slide in enumerate(slides_data):
                thumb_path = image_paths[i] if i < len(image_paths) else ""
                slides_for_index.append({
                    "file_path": str_path,
                    "slide_no": slide["slide_no"],
                    "text_raw": slide["text_raw"],
                    "notes_raw": slide["notes_raw"],
                    "thumb_path": thumb_path
                })
            
            # 4. Compute Embeddings
            logger.info(f"Computing embeddings for {len(slides_for_index)} slides")
            texts_to_embed = [
                f"{s['text_raw']}\n\n{s['notes_raw']}" for s in slides_for_index
            ]
            embeddings = self.embedder.embed(texts_to_embed)
            
            # 5. Index to Qdrant
            logger.info(f"Indexing slides to Qdrant")
            self.indexer.index_slides(slides_for_index, embeddings)
            
            self.db.update_status(
                str_path, 
                "success", 
                slide_count=len(slides_data),
                duration=0.0 # TODO: Measure time
            )
            logger.info(f"Successfully processed {file_path.name}")
            
        except Exception as e:
            logger.error(f"Failed to process {file_path}: {e}", exc_info=True)
            self.db.update_status(str_path, "failed", error_message=str(e))
