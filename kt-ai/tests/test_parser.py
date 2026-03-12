"""Tests for the File Parser module."""

import os
import pytest

from parser.file_parser import FileParser, FileMetadata, ParseResult


@pytest.fixture
def sample_repo(tmp_path):
    """Create a sample repo with various file types."""
    # Dockerfile
    (tmp_path / "Dockerfile").write_text(
        "FROM python:3.11-slim\n"
        "ENV APP_PORT=8080\n"
        "COPY . /app\n"
        "CMD [\"python\", \"main.py\"]\n"
    )

    # Python file
    svc = tmp_path / "checkoutservice"
    svc.mkdir()
    (svc / "main.py").write_text(
        "import os\n"
        "from flask import Flask\n\n"
        "app = Flask(__name__)\n\n"
        "@app.route('/checkout')\n"
        "def checkout():\n"
        "    port = os.environ['PORT']\n"
        "    return 'OK'\n"
    )

    # Go file
    (svc / "handler.go").write_text(
        'package main\n\n'
        'import (\n    "net/http"\n    "os"\n)\n\n'
        'func main() {\n'
        '    host := os.Getenv("HOST")\n'
        '    http.HandleFunc("/api/cart", cartHandler)\n'
        '}\n'
    )

    # Kubernetes YAML
    k8s = tmp_path / "kubernetes"
    k8s.mkdir()
    (k8s / "deployment.yaml").write_text(
        "apiVersion: apps/v1\n"
        "kind: Deployment\n"
        "metadata:\n"
        "  name: checkout-deployment\n"
        "spec:\n"
        "  template:\n"
        "    spec:\n"
        "      containers:\n"
        "        - name: checkout\n"
        "          image: gcr.io/checkout:v1\n"
        "          env:\n"
        "            - name: REDIS_HOST\n"
        "              value: redis\n"
    )

    # Docker Compose
    (tmp_path / "docker-compose.yml").write_text(
        "version: '3'\n"
        "services:\n"
        "  web:\n"
        "    image: nginx:latest\n"
        "    environment:\n"
        "      - API_KEY=secret\n"
        "  redis:\n"
        "    image: redis:7\n"
    )

    # Markdown
    (tmp_path / "README.md").write_text("# Test Project\n\nA sample project.\n")

    # JavaScript file
    js_dir = tmp_path / "frontend"
    js_dir.mkdir()
    (js_dir / "server.js").write_text(
        "const express = require('express');\n"
        "const app = express();\n"
        "const port = process.env.PORT;\n\n"
        "app.get('/health', (req, res) => res.send('ok'));\n"
    )

    return tmp_path


class TestFileParser:
    def test_parse_files_basic(self, sample_repo):
        files = [
            "Dockerfile",
            os.path.join("checkoutservice", "main.py"),
            "README.md",
        ]
        parser = FileParser()
        result = parser.parse_files(str(sample_repo), files)

        assert isinstance(result, ParseResult)
        assert result.total_parsed == 3

    def test_parse_dockerfile(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(str(sample_repo), ["Dockerfile"])

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.file_type == "dockerfile"
        assert "python:3.11-slim" in meta.docker_images
        assert "APP_PORT" in meta.env_variables

    def test_parse_python_file(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(
            str(sample_repo),
            [os.path.join("checkoutservice", "main.py")],
        )

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.language == "python"
        assert "flask" in [i.lower() for i in meta.imports]
        assert "/checkout" in meta.api_endpoints
        assert "PORT" in meta.env_variables

    def test_parse_go_file(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(
            str(sample_repo),
            [os.path.join("checkoutservice", "handler.go")],
        )

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.language == "go"
        assert "HOST" in meta.env_variables
        assert "/api/cart" in meta.api_endpoints

    def test_parse_kubernetes_yaml(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(
            str(sample_repo),
            [os.path.join("kubernetes", "deployment.yaml")],
        )

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.file_type == "kubernetes"
        assert len(meta.k8s_resources) == 1
        assert meta.k8s_resources[0]["kind"] == "Deployment"
        assert meta.k8s_resources[0]["name"] == "checkout-deployment"
        assert "gcr.io/checkout:v1" in meta.docker_images
        assert "REDIS_HOST" in meta.env_variables

    def test_parse_docker_compose(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(str(sample_repo), ["docker-compose.yml"])

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.file_type == "docker-compose"
        assert "web" in meta.service_names
        assert "redis" in meta.service_names
        assert "nginx:latest" in meta.docker_images

    def test_parse_javascript_file(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(
            str(sample_repo),
            [os.path.join("frontend", "server.js")],
        )

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.language == "javascript"
        assert "/health" in meta.api_endpoints
        assert "PORT" in meta.env_variables

    def test_parse_markdown_file(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(str(sample_repo), ["README.md"])

        assert result.total_parsed == 1
        meta = result.files[0]
        assert meta.file_type == "documentation"
        assert meta.language == "markdown"

    def test_skip_missing_files(self, sample_repo):
        parser = FileParser()
        result = parser.parse_files(str(sample_repo), ["nonexistent.py"])

        assert result.total_parsed == 0

    def test_skip_large_files(self, sample_repo):
        large_file = sample_repo / "big.py"
        large_file.write_text("x" * 2_000_000)

        parser = FileParser(max_file_size=1_000_000)
        result = parser.parse_files(str(sample_repo), ["big.py"])

        assert result.total_parsed == 0

    def test_detect_file_type(self):
        parser = FileParser()
        assert parser._detect_file_type("Dockerfile", "Dockerfile") == "dockerfile"
        assert parser._detect_file_type("docker-compose.yml", "docker-compose.yml") == "docker-compose"
        assert parser._detect_file_type("kubernetes/deploy.yaml", "deploy.yaml") == "kubernetes"
        assert parser._detect_file_type("README.md", "README.md") == "documentation"
        assert parser._detect_file_type("src/main.py", "main.py") == "source"
        assert parser._detect_file_type(".github/workflows/ci.yml", "ci.yml") == "cicd"
