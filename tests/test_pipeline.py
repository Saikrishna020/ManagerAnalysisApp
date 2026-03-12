"""Tests for the end-to-end analysis pipeline."""

import os
import pytest

from scanner.repo_scanner import RepoScanner
from parser.file_parser import FileParser
from extraction.knowledge_extractor import KnowledgeExtractor
from docs.doc_generator import DocGenerator


@pytest.fixture
def microservices_repo(tmp_path):
    """Create a mock microservices repository structure."""
    # Frontend service
    frontend = tmp_path / "frontend"
    frontend.mkdir()
    (frontend / "index.js").write_text(
        "const express = require('express');\n"
        "const app = express();\n"
        "const port = process.env.PORT || 3000;\n\n"
        "app.get('/', (req, res) => res.send('Hello'));\n"
        "app.get('/health', (req, res) => res.send('ok'));\n"
    )
    (frontend / "Dockerfile").write_text(
        "FROM node:18-alpine\n"
        "WORKDIR /app\n"
        "COPY . .\n"
        "ENV PORT=3000\n"
        "CMD [\"node\", \"index.js\"]\n"
    )

    # Checkout service
    checkout = tmp_path / "checkoutservice"
    checkout.mkdir()
    (checkout / "main.go").write_text(
        'package main\n\n'
        'import (\n    "net/http"\n    "os"\n)\n\n'
        'func main() {\n'
        '    port := os.Getenv("PORT")\n'
        '    http.HandleFunc("/checkout", checkoutHandler)\n'
        '    // calls paymentservice for payment processing\n'
        '    // uses redis for cart storage\n'
        '}\n'
    )
    (checkout / "Dockerfile").write_text(
        "FROM golang:1.19-alpine\n"
        "WORKDIR /app\n"
        "COPY . .\n"
        "CMD [\"go\", \"run\", \"main.go\"]\n"
    )

    # Payment service
    payment = tmp_path / "paymentservice"
    payment.mkdir()
    (payment / "main.py").write_text(
        "import os\n"
        "from flask import Flask\n\n"
        "app = Flask(__name__)\n\n"
        "@app.route('/pay')\n"
        "def process_payment():\n"
        "    return 'payment processed'\n"
    )
    (payment / "Dockerfile").write_text(
        "FROM python:3.11-slim\n"
        "COPY . /app\n"
        "CMD [\"python\", \"main.py\"]\n"
    )

    # Kubernetes manifests
    k8s = tmp_path / "kubernetes"
    k8s.mkdir()
    (k8s / "checkout-deployment.yaml").write_text(
        "apiVersion: apps/v1\n"
        "kind: Deployment\n"
        "metadata:\n"
        "  name: checkoutservice\n"
        "spec:\n"
        "  template:\n"
        "    spec:\n"
        "      containers:\n"
        "        - name: checkout\n"
        "          image: gcr.io/checkoutservice:v1\n"
        "          env:\n"
        "            - name: PAYMENT_SVC_ADDR\n"
        "              value: paymentservice:50051\n"
    )
    (k8s / "payment-deployment.yaml").write_text(
        "apiVersion: apps/v1\n"
        "kind: Deployment\n"
        "metadata:\n"
        "  name: paymentservice\n"
        "spec:\n"
        "  template:\n"
        "    spec:\n"
        "      containers:\n"
        "        - name: payment\n"
        "          image: gcr.io/paymentservice:v1\n"
    )

    # Docker compose
    (tmp_path / "docker-compose.yaml").write_text(
        "version: '3'\n"
        "services:\n"
        "  frontend:\n"
        "    build: ./frontend\n"
        "    image: frontend:latest\n"
        "  checkout:\n"
        "    build: ./checkoutservice\n"
        "  payment:\n"
        "    build: ./paymentservice\n"
        "  redis:\n"
        "    image: redis:7\n"
    )

    # Documentation
    (tmp_path / "README.md").write_text(
        "# Microservices Demo\n\n"
        "A sample microservices application.\n"
    )

    return tmp_path


class TestEndToEndPipeline:
    def test_full_pipeline(self, microservices_repo, tmp_path):
        """Test the full analysis pipeline: scan → parse → extract → generate docs."""
        output_dir = str(tmp_path / "output_docs")

        # Step 1: Scan
        scanner = RepoScanner()
        scan_result = scanner.scan(str(microservices_repo))
        assert scan_result.total_files > 0

        # Step 2: Parse
        parser = FileParser()
        parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)
        assert parse_result.total_parsed > 0

        # Step 3: Extract knowledge
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        # Verify services detected
        service_names = [s.name for s in knowledge.services]
        assert "checkoutservice" in service_names
        assert "paymentservice" in service_names
        assert "frontend" in service_names

        # Verify languages detected
        assert "go" in knowledge.languages
        assert "python" in knowledge.languages
        assert "javascript" in knowledge.languages

        # Verify infrastructure
        assert len(knowledge.infrastructure.docker_images) > 0
        assert len(knowledge.infrastructure.k8s_resources) > 0

        # Step 4: Generate documentation
        doc_gen = DocGenerator()
        docs = doc_gen.generate(knowledge, output_dir)

        # Verify files generated
        assert os.path.isfile(os.path.join(output_dir, "README.md"))
        assert os.path.isfile(os.path.join(output_dir, "architecture.md"))
        assert os.path.isfile(os.path.join(output_dir, "services.md"))
        assert os.path.isfile(os.path.join(output_dir, "deployment.md"))

        # Verify content quality
        assert "checkoutservice" in docs.services
        assert "paymentservice" in docs.services
        assert "frontend" in docs.services

    def test_pipeline_with_empty_repo(self, tmp_path):
        """Test pipeline with empty repository."""
        empty_repo = tmp_path / "empty"
        empty_repo.mkdir()
        output_dir = str(tmp_path / "output")

        scanner = RepoScanner()
        scan_result = scanner.scan(str(empty_repo))
        assert scan_result.total_files == 0

        parser = FileParser()
        parse_result = parser.parse_files(scan_result.repo_path, scan_result.files)

        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)
        assert len(knowledge.services) == 0

        doc_gen = DocGenerator()
        docs = doc_gen.generate(knowledge, output_dir)
        assert os.path.isfile(os.path.join(output_dir, "README.md"))
