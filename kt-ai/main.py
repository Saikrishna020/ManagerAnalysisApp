"""
KT AI - Main Entry Point

CLI interface for repository analysis, documentation generation,
and the KT AI assistant.

Usage:
    python main.py analyze <repo_url> [--output-dir generated_docs]
    python main.py analyze-local <repo_path> [--output-dir generated_docs]
    python main.py chat
    python main.py serve [--host 0.0.0.0] [--port 8000]
"""

import argparse
import logging
import sys
import os

from scanner.repo_scanner import RepoScanner
from parser.file_parser import FileParser
from extraction.knowledge_extractor import KnowledgeExtractor
from docs.doc_generator import DocGenerator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("kt_ai")


def cmd_analyze(args):
    """Analyze a remote git repository."""
    repo_url = args.repo_url
    output_dir = args.output_dir
    target_dir = args.target_dir

    print(f"\n🔍 Analyzing repository: {repo_url}\n")

    # Step 1: Clone and scan
    print("📂 Step 1: Cloning and scanning repository...")
    scanner = RepoScanner()
    scan_result = scanner.clone_and_scan(repo_url, target_dir)
    print(f"   Found {scan_result.total_files} relevant files\n")

    # Step 2: Parse files
    print("📄 Step 2: Parsing files...")
    parser = FileParser()
    parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)
    print(f"   Parsed {parse_result.total_parsed} files\n")

    # Step 3: Extract knowledge
    print("🧠 Step 3: Extracting knowledge...")
    extractor = KnowledgeExtractor()
    knowledge = extractor.extract(parse_result)
    print(f"   Services: {len(knowledge.services)}")
    print(f"   Dependencies: {len(knowledge.dependencies)}")
    print(f"   APIs: {len(knowledge.apis)}\n")

    # Step 4: Generate documentation
    print("📝 Step 4: Generating documentation...")
    doc_gen = DocGenerator()
    docs = doc_gen.generate(knowledge, output_dir)
    print(f"   Output directory: {docs.output_dir}\n")

    # Step 5: Index into vector store (optional)
    if not args.no_index:
        print("🔢 Step 5: Creating vector index...")
        try:
            from rag.vector_store import VectorStore
            store = VectorStore()
            store.index_knowledge(knowledge, parse_result)
            print("   Vector index created successfully\n")
        except ImportError:
            print("   ⚠️  Vector store dependencies not available. Skipping indexing.\n")
        except Exception as e:
            print(f"   ⚠️  Vector indexing failed: {e}. Skipping.\n")

    print("✅ Analysis complete!")
    print(f"\nSummary:\n{knowledge.summary}\n")
    print(f"Generated documentation in: {output_dir}/")
    print("  - README.md")
    print("  - architecture.md")
    print("  - services.md")
    print("  - deployment.md")


def cmd_analyze_local(args):
    """Analyze a local repository directory."""
    repo_path = args.repo_path
    output_dir = args.output_dir

    if not os.path.isdir(repo_path):
        print(f"❌ Error: Path does not exist: {repo_path}")
        sys.exit(1)

    print(f"\n🔍 Analyzing local repository: {repo_path}\n")

    # Step 1: Scan
    print("📂 Step 1: Scanning repository...")
    scanner = RepoScanner()
    scan_result = scanner.scan(repo_path)
    print(f"   Found {scan_result.total_files} relevant files\n")

    # Step 2: Parse files
    print("📄 Step 2: Parsing files...")
    parser = FileParser()
    parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)
    print(f"   Parsed {parse_result.total_parsed} files\n")

    # Step 3: Extract knowledge
    print("🧠 Step 3: Extracting knowledge...")
    extractor = KnowledgeExtractor()
    knowledge = extractor.extract(parse_result)
    print(f"   Services: {len(knowledge.services)}")
    print(f"   Dependencies: {len(knowledge.dependencies)}")
    print(f"   APIs: {len(knowledge.apis)}\n")

    # Step 4: Generate documentation
    print("📝 Step 4: Generating documentation...")
    doc_gen = DocGenerator()
    docs = doc_gen.generate(knowledge, output_dir)
    print(f"   Output directory: {docs.output_dir}\n")

    # Step 5: Index into vector store (optional)
    if not args.no_index:
        print("🔢 Step 5: Creating vector index...")
        try:
            from rag.vector_store import VectorStore
            store = VectorStore()
            store.index_knowledge(knowledge, parse_result)
            print("   Vector index created successfully\n")
        except ImportError:
            print("   ⚠️  Vector store dependencies not available. Skipping indexing.\n")
        except Exception as e:
            print(f"   ⚠️  Vector indexing failed: {e}. Skipping.\n")

    print("✅ Analysis complete!")
    print(f"\nSummary:\n{knowledge.summary}\n")
    print(f"Generated documentation in: {output_dir}/")
    print("  - README.md")
    print("  - architecture.md")
    print("  - services.md")
    print("  - deployment.md")


def cmd_chat(args):
    """Start an interactive chat session."""
    try:
        from rag.vector_store import VectorStore
        from rag.assistant import KTAssistant

        store = VectorStore()
        assistant = KTAssistant(vector_store=store)
        assistant.chat()
    except ImportError:
        print("❌ Chat requires vector store dependencies (chromadb, langchain).")
        print("Install with: pip install chromadb langchain langchain-chroma")
        sys.exit(1)


def cmd_serve(args):
    """Start the FastAPI server."""
    try:
        import uvicorn
        print(f"\n🚀 Starting KT AI server on {args.host}:{args.port}\n")
        uvicorn.run("api.server:app", host=args.host, port=args.port, reload=args.reload)
    except ImportError:
        print("❌ Server requires uvicorn. Install with: pip install uvicorn")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="KT AI - AI-powered repository analysis and knowledge transfer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # analyze command
    analyze_parser = subparsers.add_parser("analyze", help="Analyze a remote git repository")
    analyze_parser.add_argument("repo_url", help="Git repository URL")
    analyze_parser.add_argument("--target-dir", help="Directory to clone into")
    analyze_parser.add_argument("--output-dir", default="generated_docs", help="Output directory for docs")
    analyze_parser.add_argument("--no-index", action="store_true", help="Skip vector indexing")

    # analyze-local command
    local_parser = subparsers.add_parser("analyze-local", help="Analyze a local repository")
    local_parser.add_argument("repo_path", help="Path to local repository")
    local_parser.add_argument("--output-dir", default="generated_docs", help="Output directory for docs")
    local_parser.add_argument("--no-index", action="store_true", help="Skip vector indexing")

    # chat command
    subparsers.add_parser("chat", help="Start interactive chat")

    # serve command
    serve_parser = subparsers.add_parser("serve", help="Start the API server")
    serve_parser.add_argument("--host", default="127.0.0.1", help="Server host")
    serve_parser.add_argument("--port", type=int, default=8000, help="Server port")
    serve_parser.add_argument("--reload", action="store_true", help="Enable auto-reload")

    args = parser.parse_args()

    if args.command == "analyze":
        cmd_analyze(args)
    elif args.command == "analyze-local":
        cmd_analyze_local(args)
    elif args.command == "chat":
        cmd_chat(args)
    elif args.command == "serve":
        cmd_serve(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
