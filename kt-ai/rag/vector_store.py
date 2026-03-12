"""
Vector Knowledge Index Module

Converts repository documents into embeddings and stores them
in a Chroma vector database for semantic retrieval.
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)


class VectorStore:
    """Manages a Chroma vector store for repository knowledge."""

    def __init__(
        self,
        persist_directory: str = "./chroma_db",
        collection_name: str = "kt_ai",
        embedding_function=None,
    ):
        """
        Initialize the vector store.

        Args:
            persist_directory: Directory to persist the Chroma database.
            collection_name: Name of the Chroma collection.
            embedding_function: Optional LangChain embedding function.
                               If None, uses default Chroma embedding.
        """
        self.persist_directory = persist_directory
        self.collection_name = collection_name
        self._embedding_function = embedding_function
        self._store = None

    def _get_store(self):
        """Lazily initialize the Chroma store."""
        if self._store is None:
            try:
                from langchain_chroma import Chroma
            except ImportError:
                from langchain_community.vectorstores import Chroma

            kwargs = {
                "collection_name": self.collection_name,
                "persist_directory": self.persist_directory,
            }
            if self._embedding_function:
                kwargs["embedding_function"] = self._embedding_function

            self._store = Chroma(**kwargs)
            logger.info(
                "Initialized Chroma store at %s (collection: %s)",
                self.persist_directory, self.collection_name,
            )
        return self._store

    def index_documents(self, documents: list, metadatas: Optional[list] = None, ids: Optional[list] = None):
        """
        Add documents to the vector store.

        Args:
            documents: List of text strings to embed and store.
            metadatas: Optional list of metadata dicts for each document.
            ids: Optional list of unique IDs for each document.
        """
        if not documents:
            logger.warning("No documents to index")
            return

        store = self._get_store()

        if ids is None:
            ids = [f"doc_{i}" for i in range(len(documents))]

        store.add_texts(texts=documents, metadatas=metadatas, ids=ids)
        logger.info("Indexed %d documents into vector store", len(documents))

    def index_knowledge(self, knowledge, parse_result=None):
        """
        Index repository knowledge and optionally raw file content.

        Args:
            knowledge: RepositoryKnowledge instance.
            parse_result: Optional ParseResult with raw file contents.
        """
        documents = []
        metadatas = []
        ids = []
        doc_id = 0

        # Index the summary
        if knowledge.summary:
            documents.append(knowledge.summary)
            metadatas.append({"type": "summary", "source": "knowledge"})
            ids.append(f"summary_{doc_id}")
            doc_id += 1

        # Index service information
        for svc in knowledge.services:
            text = (
                f"Service: {svc.name}\n"
                f"Language: {svc.language}\n"
                f"Path: {svc.path}\n"
                f"Docker Image: {svc.docker_image}\n"
            )
            if svc.api_endpoints:
                text += f"API Endpoints: {', '.join(svc.api_endpoints)}\n"
            if svc.env_variables:
                text += f"Environment Variables: {', '.join(svc.env_variables)}\n"

            documents.append(text)
            metadatas.append({"type": "service", "name": svc.name, "source": "knowledge"})
            ids.append(f"service_{svc.name}_{doc_id}")
            doc_id += 1

        # Index dependencies
        for dep in knowledge.dependencies:
            text = f"Dependency: {dep.source} depends on {dep.target} ({dep.relation_type})"
            documents.append(text)
            metadatas.append({"type": "dependency", "source_svc": dep.source, "target": dep.target})
            ids.append(f"dep_{doc_id}")
            doc_id += 1

        # Index infrastructure
        if knowledge.infrastructure.k8s_resources:
            k8s_text = "Kubernetes Resources:\n"
            for res in knowledge.infrastructure.k8s_resources:
                k8s_text += f"- {res['kind']}: {res['name']}\n"
            documents.append(k8s_text)
            metadatas.append({"type": "infrastructure", "source": "kubernetes"})
            ids.append(f"k8s_{doc_id}")
            doc_id += 1

        if knowledge.infrastructure.docker_images:
            docker_text = "Docker Images:\n"
            for img in knowledge.infrastructure.docker_images:
                docker_text += f"- {img}\n"
            documents.append(docker_text)
            metadatas.append({"type": "infrastructure", "source": "docker"})
            ids.append(f"docker_{doc_id}")
            doc_id += 1

        # Index raw file contents (chunked)
        if parse_result:
            for file_meta in parse_result.files:
                if file_meta.content and len(file_meta.content) < 10000:
                    documents.append(f"File: {file_meta.path}\n\n{file_meta.content}")
                    metadatas.append({
                        "type": "source",
                        "path": file_meta.path,
                        "language": file_meta.language,
                    })
                    ids.append(f"file_{doc_id}")
                    doc_id += 1

        self.index_documents(documents, metadatas, ids)

    def search(self, query: str, k: int = 5) -> list:
        """
        Search the vector store for relevant documents.

        Args:
            query: Search query string.
            k: Number of results to return.

        Returns:
            List of (document, metadata, score) tuples.
        """
        store = self._get_store()
        results = store.similarity_search_with_score(query, k=k)
        return [(doc.page_content, doc.metadata, score) for doc, score in results]

    def as_retriever(self, **kwargs):
        """Return a LangChain retriever from this vector store."""
        store = self._get_store()
        return store.as_retriever(**kwargs)
