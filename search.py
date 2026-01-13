"""
高速検索スクリプト
インデックス化済みのQdrantから瞬時に検索
"""

import sys
from pathlib import Path
import logging
import time

# パスを追加
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from ingest.embeddings import TextEmbedder
from ingest.indexer import QdrantIndexer

# ロギング設定
logging.basicConfig(
    level=logging.WARNING,  # 検索時は警告のみ表示
    format='%(message)s'
)

# 埋め込みとインデクサーのログを抑制
logging.getLogger('ingest.embeddings').setLevel(logging.WARNING)
logging.getLogger('ingest.indexer').setLevel(logging.WARNING)


class FastSearch:
    """高速検索"""

    def __init__(self, qdrant_path: str = "index/qdrant_storage"):
        """
        Args:
            qdrant_path: Qdrantストレージパス
        """
        print("🔧 検索エンジン初期化中...", end='', flush=True)
        start = time.time()

        # 埋め込みモデル初期化
        self.embedder = TextEmbedder(
            model_name="intfloat/e5-base-v2",
            device="cpu"
        )

        # Qdrantクライアント初期化
        self.indexer = QdrantIndexer(storage_path=qdrant_path)
        vector_dim = self.embedder.get_dimension()
        self.indexer.initialize(vector_dimension=vector_dim)

        elapsed = time.time() - start
        print(f" 完了 ({elapsed:.2f}秒)")

        # コレクション情報表示
        info = self.indexer.get_collection_info()
        print(f"📚 インデックス済みページ数: {info['points_count']}\n")

    def search(self, query: str, top_k: int = 5, show_text: bool = True):
        """
        検索実行

        Args:
            query: 検索クエリ
            top_k: 取得件数
            show_text: テキストを表示するか
        """
        print(f"🔍 検索クエリ: \"{query}\"")
        start = time.time()

        # クエリ埋め込み計算
        query_vector = self.embedder.embed_texts([query])[0]

        # Qdrant検索
        results = self.indexer.search(
            query_vector=query_vector,
            top_k=top_k,
            score_threshold=0.0
        )

        elapsed = time.time() - start
        print(f"⏱️  検索時間: {elapsed:.3f}秒")
        print(f"📊 検索結果: {len(results)}件\n")

        # 結果表示
        if not results:
            print("❌ 該当する結果が見つかりませんでした")
            return

        for i, result in enumerate(results, 1):
            score = result['score']
            file_name = result['file_name']
            page_num = result['page_num']
            text = result['text']
            image_path = result['image_path']

            # スコアバー表示
            bar_length = int(score * 20)
            bar = "█" * bar_length + "░" * (20 - bar_length)

            print(f"{'='*70}")
            print(f"🏆 結果 #{i}")
            print(f"   スコア: {bar} {score:.4f}")
            print(f"   ファイル: {file_name}")
            print(f"   ページ: {page_num}")
            print(f"   画像: {image_path}")

            if show_text and text:
                # テキスト要約（最初の300文字）
                text_preview = text[:300]
                if len(text) > 300:
                    text_preview += "..."
                print(f"\n   📝 テキスト抜粋:")
                print(f"   {text_preview}\n")

        print(f"{'='*70}\n")


def interactive_mode(searcher: FastSearch):
    """対話モード"""
    print("=" * 70)
    print("🚀 対話モード開始")
    print("   - 検索クエリを入力してください")
    print("   - 終了: 'exit', 'quit', 'q'")
    print("=" * 70)
    print()

    while True:
        try:
            query = input("🔍 検索 > ").strip()

            if not query:
                continue

            if query.lower() in ['exit', 'quit', 'q']:
                print("\n👋 検索を終了します")
                break

            searcher.search(query, top_k=5, show_text=True)

        except KeyboardInterrupt:
            print("\n\n👋 検索を終了します")
            break
        except Exception as e:
            print(f"❌ エラー: {e}")
            import traceback
            traceback.print_exc()


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(
        description="高速検索CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 対話モード
  python search.py

  # ワンショット検索
  python search.py --query "hinge tolerance"

  # 結果件数指定
  python search.py --query "mechanical design" --top-k 10

  # テキスト非表示（ファイル名とページ番号のみ）
  python search.py --query "assembly" --no-text
        """
    )

    parser.add_argument(
        '--query', '-q',
        type=str,
        help='検索クエリ（指定しない場合は対話モード）'
    )
    parser.add_argument(
        '--top-k', '-k',
        type=int,
        default=5,
        help='取得件数（デフォルト: 5）'
    )
    parser.add_argument(
        '--no-text',
        action='store_true',
        help='テキストを表示しない'
    )
    parser.add_argument(
        '--qdrant',
        type=str,
        default='index/qdrant_storage',
        help='Qdrantストレージパス'
    )

    args = parser.parse_args()

    # Qdrant存在確認
    if not Path(args.qdrant).exists():
        print(f"❌ エラー: Qdrantストレージが見つかりません: {args.qdrant}")
        print("\n💡 ヒント: 先にインデックス化を実行してください:")
        print("   python local_poc_pdf.py --source <PDFフォルダ> --full")
        sys.exit(1)

    # 検索エンジン初期化
    try:
        searcher = FastSearch(qdrant_path=args.qdrant)
    except Exception as e:
        print(f"❌ 初期化エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # クエリモード or 対話モード
    if args.query:
        # ワンショット検索
        searcher.search(
            query=args.query,
            top_k=args.top_k,
            show_text=not args.no_text
        )
    else:
        # 対話モード
        interactive_mode(searcher)


if __name__ == "__main__":
    main()
