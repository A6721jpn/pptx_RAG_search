import argparse
import logging
import sys
from pathlib import Path

# Add project root to path for direct script execution
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
# Also add src directory
sys.path.insert(0, str(Path(__file__).resolve().parent))

import yaml
from ingest.ingest_pipeline import IngestPipeline
from rag.retriever import Retriever

def load_config(config_path: str):
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def setup_logging(config):
    log_config = config.get('logging', {})
    log_file = Path(log_config.get('file', 'data/logs/app.log'))
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    logging.basicConfig(
        level=log_config.get('level', 'INFO'),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding='utf-8')
        ]
    )

def run_search(config: dict, query: str, top_k: int = 5, threshold: float = 0.0):
    """Run search and display results."""
    ingest_config = config.get('ingest', {})
    qdrant_path = ingest_config.get('qdrant_path', './index/qdrant_storage')
    collection_name = ingest_config.get('collection_name', 'pptx_slides')
    
    retriever = Retriever(qdrant_path=qdrant_path, collection_name=collection_name)
    results = retriever.search(query, top_k=top_k, score_threshold=threshold)
    
    if not results:
        print("\nNo matching slides found.")
        return
    
    print(f"\n{'='*60}")
    print(f"Search Results for: \"{query}\"")
    print(f"{'='*60}\n")
    
    for i, r in enumerate(results, 1):
        print(f"--- Result {i} (Score: {r['score']:.4f}) ---")
        print(f"File: {r['file_name']}")
        print(f"Slide: {r['slide_no']}")
        print(f"Image: {r['thumb_path']}")
        print(f"Text:\n{r['text_raw'][:300]}...")
        if r['notes_raw']:
            print(f"Notes:\n{r['notes_raw'][:200]}...")
        print()

def main():
    parser = argparse.ArgumentParser(description="PPTX RAG Local PoC")
    parser.add_argument("--config", default="configs/local_config.yaml", help="Path to config file")
    parser.add_argument("--mode", choices=["ingest", "search"], default="ingest", help="Operation mode")
    parser.add_argument("--query", type=str, help="Search query (required for search mode)")
    parser.add_argument("--top-k", type=int, default=5, help="Number of results to return")
    parser.add_argument("--threshold", type=float, default=0.0, help="Minimum score threshold")
    args = parser.parse_args()
    
    try:
        config = load_config(args.config)
        setup_logging(config)
        
        logger = logging.getLogger(__name__)
        logger.info(f"Starting application in {args.mode} mode")
        
        if args.mode == "ingest":
            ingest_config = config.get('ingest', {})
            pipeline = IngestPipeline(ingest_config)
            pipeline.run()
        elif args.mode == "search":
            if not args.query:
                print("Error: --query is required for search mode.")
                sys.exit(1)
            run_search(config, args.query, args.top_k, args.threshold)
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
