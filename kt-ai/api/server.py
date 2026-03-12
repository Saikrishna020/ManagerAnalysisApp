"""
FastAPI Server for KT AI

Provides REST API endpoints for repository analysis,
documentation generation, and the KT AI assistant.
"""

import os
import logging
from typing import Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from scanner.repo_scanner import RepoScanner
from parser.file_parser import FileParser
from extraction.knowledge_extractor import KnowledgeExtractor
from docs.doc_generator import DocGenerator
from rag.vector_store import VectorStore
from rag.assistant import KTAssistant

logger = logging.getLogger(__name__)

app = FastAPI(
    title="KT AI",
    description="AI-powered repository analysis and knowledge transfer assistant",
    version="1.0.0",
)

# Global state for the current analysis session
_state = {
    "knowledge": None,
    "parse_result": None,
    "vector_store": None,
    "assistant": None,
}


class AnalyzeRequest(BaseModel):
    repo_url: str = Field(..., description="Git repository URL to analyze")
    target_dir: Optional[str] = Field(None, description="Local directory to clone into")
    output_dir: str = Field("generated_docs", description="Output directory for generated docs")


class AnalyzeLocalRequest(BaseModel):
    repo_path: str = Field(..., description="Local path to repository")
    output_dir: str = Field("generated_docs", description="Output directory for generated docs")


class ChatRequest(BaseModel):
    question: str = Field(..., description="Question about the repository")


class AnalyzeResponse(BaseModel):
    repo_name: str
    services_count: int
    dependencies_count: int
    apis_count: int
    languages: list
    docs_generated: list
    summary: str


class ChatResponse(BaseModel):
    answer: str
    sources: list


@app.get("/health")
def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "service": "kt-ai"}


@app.post("/analyze", response_model=AnalyzeResponse)
def analyze_repo(request: AnalyzeRequest):
    """Analyze a git repository and generate documentation."""
    try:
        scanner = RepoScanner()
        parser = FileParser()
        extractor = KnowledgeExtractor()
        doc_gen = DocGenerator()

        # Clone and scan
        scan_result = scanner.clone_and_scan(request.repo_url, request.target_dir)

        # Parse files
        parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)

        # Extract knowledge
        knowledge = extractor.extract(parse_result)

        # Generate documentation
        docs = doc_gen.generate(knowledge, request.output_dir)

        # Index into vector store
        vector_store = VectorStore()
        vector_store.index_knowledge(knowledge, parse_result)

        # Store state
        _state["knowledge"] = knowledge
        _state["parse_result"] = parse_result
        _state["vector_store"] = vector_store
        _state["assistant"] = KTAssistant(vector_store=vector_store)

        return AnalyzeResponse(
            repo_name=knowledge.repo_name,
            services_count=len(knowledge.services),
            dependencies_count=len(knowledge.dependencies),
            apis_count=len(knowledge.apis),
            languages=knowledge.languages,
            docs_generated=["README.md", "architecture.md", "services.md", "deployment.md"],
            summary=knowledge.summary,
        )
    except Exception as e:
        logger.exception("Analysis failed")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/analyze/local", response_model=AnalyzeResponse)
def analyze_local_repo(request: AnalyzeLocalRequest):
    """Analyze a local repository directory."""
    if not os.path.isdir(request.repo_path):
        raise HTTPException(status_code=400, detail=f"Path does not exist: {request.repo_path}")

    try:
        scanner = RepoScanner()
        parser = FileParser()
        extractor = KnowledgeExtractor()
        doc_gen = DocGenerator()

        # Scan local directory
        scan_result = scanner.scan(request.repo_path)

        # Parse files
        parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)

        # Extract knowledge
        knowledge = extractor.extract(parse_result)

        # Generate documentation
        docs = doc_gen.generate(knowledge, request.output_dir)

        # Index into vector store
        vector_store = VectorStore()
        vector_store.index_knowledge(knowledge, parse_result)

        # Store state
        _state["knowledge"] = knowledge
        _state["parse_result"] = parse_result
        _state["vector_store"] = vector_store
        _state["assistant"] = KTAssistant(vector_store=vector_store)

        return AnalyzeResponse(
            repo_name=knowledge.repo_name,
            services_count=len(knowledge.services),
            dependencies_count=len(knowledge.dependencies),
            apis_count=len(knowledge.apis),
            languages=knowledge.languages,
            docs_generated=["README.md", "architecture.md", "services.md", "deployment.md"],
            summary=knowledge.summary,
        )
    except Exception as e:
        logger.exception("Analysis failed")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/chat", response_model=ChatResponse)
def chat(request: ChatRequest):
    """Ask a question about the analyzed repository."""
    if _state["assistant"] is None:
        raise HTTPException(
            status_code=400,
            detail="No repository has been analyzed yet. Please run /analyze first.",
        )

    result = _state["assistant"].ask(request.question)
    return ChatResponse(
        answer=result["answer"],
        sources=result["sources"],
    )


@app.get("/knowledge")
def get_knowledge():
    """Get the current repository knowledge model."""
    if _state["knowledge"] is None:
        raise HTTPException(
            status_code=400,
            detail="No repository has been analyzed yet.",
        )

    knowledge = _state["knowledge"]
    return {
        "repo_name": knowledge.repo_name,
        "services": [
            {
                "name": s.name,
                "language": s.language,
                "path": s.path,
                "docker_image": s.docker_image,
                "api_endpoints": s.api_endpoints,
                "env_variables": s.env_variables,
            }
            for s in knowledge.services
        ],
        "dependencies": [
            {
                "source": d.source,
                "target": d.target,
                "relation_type": d.relation_type,
            }
            for d in knowledge.dependencies
        ],
        "infrastructure": {
            "docker_images": knowledge.infrastructure.docker_images,
            "k8s_resources": knowledge.infrastructure.k8s_resources,
            "cicd_files": knowledge.infrastructure.cicd_files,
        },
        "apis": knowledge.apis,
        "languages": knowledge.languages,
        "summary": knowledge.summary,
    }
