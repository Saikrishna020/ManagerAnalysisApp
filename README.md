# Manager Case Analysis App

This project contains a Streamlit application that analyzes manager assignments from an uploaded Excel file.

## KT AI – Knowledge Transfer AI

KT AI is an AI system that analyzes software repositories and automatically generates documentation and knowledge to assist developer onboarding and knowledge transfer.

### Features

- **Repository Scanning**: Clone and scan git repositories, identifying relevant source code, configuration, and documentation files
- **File Parsing**: Extract metadata from source files including service names, imports, API endpoints, Docker images, Kubernetes resources, and environment variables
- **Knowledge Extraction**: Detect services, dependencies, infrastructure, and APIs from parsed file metadata
- **Documentation Generation**: Generate structured documentation (README, architecture, services, deployment) using templates or LLM
- **Vector Knowledge Index**: Embed repository knowledge into a Chroma vector database for semantic search
- **KT AI Assistant**: RAG-based conversational assistant to answer developer questions about the repository

### Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Analyze a remote repository
python main.py analyze https://github.com/GoogleCloudPlatform/microservices-demo

# Analyze a local repository
python main.py analyze-local /path/to/repo

# Start interactive chat (requires prior analysis with vector indexing)
python main.py chat

# Start the API server
python main.py serve
```

### Project Structure

```
├── scanner/
│   └── repo_scanner.py          # Repository cloning and file scanning
├── parser/
│   └── file_parser.py           # File parsing and metadata extraction
├── extraction/
│   └── knowledge_extractor.py   # Knowledge extraction engine
├── docs/
│   └── doc_generator.py         # Documentation generation
├── rag/
│   ├── vector_store.py          # Chroma vector database integration
│   └── assistant.py             # RAG-based Q&A assistant
├── api/
│   └── server.py                # FastAPI REST API server
├── tests/                       # Unit and integration tests
├── main.py                      # CLI entry point
└── requirements.txt             # Python dependencies
```

### API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Health check |
| POST | `/analyze` | Analyze a remote git repository |
| POST | `/analyze/local` | Analyze a local repository |
| POST | `/chat` | Ask a question about the analyzed repository |
| GET | `/knowledge` | Get the current repository knowledge model |

### Running Tests

```bash
python -m pytest tests/ -v
```


