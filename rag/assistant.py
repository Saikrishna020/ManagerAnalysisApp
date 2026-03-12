"""
KT AI Assistant Module

Provides a RAG-based conversational assistant that answers questions
about a repository using the vector knowledge index.
"""

import logging
from typing import Optional

logger = logging.getLogger(__name__)


class KTAssistant:
    """RAG-based assistant for answering repository questions."""

    def __init__(self, vector_store=None, llm=None):
        """
        Initialize the KT AI Assistant.

        Args:
            vector_store: VectorStore instance with indexed knowledge.
            llm: Optional LangChain LLM instance. If None, returns context only.
        """
        self.vector_store = vector_store
        self.llm = llm
        self._chain = None

    def _build_chain(self):
        """Build a LangChain RAG chain if LLM is available."""
        if self.llm is None or self.vector_store is None:
            return None

        try:
            from langchain.chains import RetrievalQA
            from langchain.prompts import PromptTemplate

            prompt_template = PromptTemplate(
                input_variables=["context", "question"],
                template=(
                    "You are KT AI, a knowledgeable assistant that helps developers "
                    "understand software repositories.\n\n"
                    "Use the following context to answer the question. If you don't know "
                    "the answer, say so honestly.\n\n"
                    "Context:\n{context}\n\n"
                    "Question: {question}\n\n"
                    "Answer:"
                ),
            )

            retriever = self.vector_store.as_retriever(search_kwargs={"k": 5})
            chain = RetrievalQA.from_chain_type(
                llm=self.llm,
                chain_type="stuff",
                retriever=retriever,
                chain_type_kwargs={"prompt": prompt_template},
                return_source_documents=True,
            )
            return chain
        except Exception as e:
            logger.error("Failed to build RAG chain: %s", e)
            return None

    def ask(self, question: str) -> dict:
        """
        Ask a question about the repository.

        Args:
            question: The developer's question.

        Returns:
            Dict with 'answer', 'sources', and 'context' keys.
        """
        if self.vector_store is None:
            return {
                "answer": "No repository has been analyzed yet. Please run analysis first.",
                "sources": [],
                "context": [],
            }

        # Retrieve relevant context
        results = self.vector_store.search(question, k=5)
        context_docs = [r[0] for r in results]
        context_metadata = [r[1] for r in results]

        # If LLM is available, use RAG chain
        if self.llm:
            if self._chain is None:
                self._chain = self._build_chain()

            if self._chain:
                try:
                    response = self._chain.invoke({"query": question})
                    return {
                        "answer": response.get("result", ""),
                        "sources": context_metadata,
                        "context": context_docs,
                    }
                except Exception as e:
                    logger.error("RAG chain error: %s", e)

        # Fallback: return context without LLM-generated answer
        answer = self._format_context_answer(question, context_docs, context_metadata)
        return {
            "answer": answer,
            "sources": context_metadata,
            "context": context_docs,
        }

    def _format_context_answer(self, question: str, context_docs: list, metadata: list) -> str:
        """Format an answer from raw context (no LLM)."""
        if not context_docs:
            return "No relevant information found for your question."

        lines = [f"Based on the repository analysis, here is relevant information:\n"]
        for i, (doc, meta) in enumerate(zip(context_docs, metadata), 1):
            source = meta.get("type", "unknown")
            name = meta.get("name", meta.get("path", ""))
            lines.append(f"### Source {i} ({source}: {name})\n")
            lines.append(doc[:500])
            lines.append("")

        return "\n".join(lines)

    def chat(self):
        """Start an interactive chat session (CLI mode)."""
        print("\n🤖 KT AI Assistant")
        print("Ask questions about the repository. Type 'quit' to exit.\n")

        while True:
            try:
                question = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not question or question.lower() in ("quit", "exit", "q"):
                print("Goodbye!")
                break

            result = self.ask(question)
            print(f"\nKT AI: {result['answer']}\n")
