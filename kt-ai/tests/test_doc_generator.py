"""Tests for the Documentation Generator module."""

import os
import pytest

from docs.doc_generator import DocGenerator, GeneratedDocs
from extraction.knowledge_extractor import (
    RepositoryKnowledge,
    ServiceInfo,
    DependencyRelation,
    InfrastructureInfo,
)


@pytest.fixture
def sample_knowledge():
    """Create sample repository knowledge for testing."""
    return RepositoryKnowledge(
        repo_path="/repo/microservices-demo",
        repo_name="microservices-demo",
        services=[
            ServiceInfo(
                name="checkoutservice",
                language="go",
                path="src/checkoutservice/main.go",
                docker_image="gcr.io/checkout:v1",
                api_endpoints=["/checkout", "/cart"],
                env_variables=["PORT", "REDIS_HOST"],
            ),
            ServiceInfo(
                name="paymentservice",
                language="python",
                path="src/paymentservice/main.py",
                docker_image="gcr.io/payment:v1",
                api_endpoints=["/pay"],
                env_variables=["PORT", "STRIPE_KEY"],
            ),
        ],
        infrastructure=InfrastructureInfo(
            docker_images=["gcr.io/checkout:v1", "gcr.io/payment:v1"],
            k8s_resources=[
                {"kind": "Deployment", "name": "checkout"},
                {"kind": "Deployment", "name": "payment"},
                {"kind": "Service", "name": "checkout-svc"},
            ],
            cicd_files=[".github/workflows/ci.yml"],
            config_files=["config.yaml"],
        ),
        dependencies=[
            DependencyRelation(
                source="checkoutservice",
                target="paymentservice",
                relation_type="service-to-service",
            ),
            DependencyRelation(
                source="checkoutservice",
                target="redis",
                relation_type="service-to-cache",
            ),
        ],
        apis=[
            {"path": "/checkout", "source_file": "src/checkoutservice/main.go"},
            {"path": "/pay", "source_file": "src/paymentservice/main.py"},
        ],
        languages=["go", "python"],
        summary="Repository: microservices-demo\nLanguages: go, python\nServices: 2 detected",
    )


class TestDocGenerator:
    def test_generate_creates_files(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert os.path.isfile(os.path.join(output_dir, "README.md"))
        assert os.path.isfile(os.path.join(output_dir, "architecture.md"))
        assert os.path.isfile(os.path.join(output_dir, "services.md"))
        assert os.path.isfile(os.path.join(output_dir, "deployment.md"))

    def test_readme_content(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert "microservices-demo" in docs.readme
        assert "checkoutservice" in docs.readme
        assert "paymentservice" in docs.readme

    def test_architecture_content(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert "Architecture" in docs.architecture
        assert "go" in docs.architecture
        assert "python" in docs.architecture
        assert "checkoutservice" in docs.architecture

    def test_services_content(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert "checkoutservice" in docs.services
        assert "paymentservice" in docs.services
        assert "/checkout" in docs.services
        assert "/pay" in docs.services
        assert "PORT" in docs.services

    def test_deployment_content(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert "Docker" in docs.deployment
        assert "Kubernetes" in docs.deployment
        assert "gcr.io/checkout:v1" in docs.deployment
        assert "checkout" in docs.deployment

    def test_empty_knowledge(self, tmp_path):
        output_dir = str(tmp_path / "docs")
        knowledge = RepositoryKnowledge(
            repo_path="/repo/empty",
            repo_name="empty",
        )
        gen = DocGenerator()
        docs = gen.generate(knowledge, output_dir)

        assert "empty" in docs.readme
        assert os.path.isfile(os.path.join(output_dir, "README.md"))

    def test_output_dir_created(self, sample_knowledge, tmp_path):
        output_dir = str(tmp_path / "new" / "nested" / "docs")
        gen = DocGenerator()
        docs = gen.generate(sample_knowledge, output_dir)

        assert os.path.isdir(output_dir)
        assert docs.output_dir == output_dir
