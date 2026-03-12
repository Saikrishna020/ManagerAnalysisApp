"""Tests for the Repository Scanner module."""

import os
import tempfile
import pytest

from scanner.repo_scanner import RepoScanner, ScanResult


@pytest.fixture
def sample_repo(tmp_path):
    """Create a sample repository directory structure."""
    # Create service directories
    svc1 = tmp_path / "checkoutservice"
    svc1.mkdir()
    (svc1 / "main.go").write_text("package main\nfunc main() {}")
    (svc1 / "Dockerfile").write_text("FROM golang:1.19\nCOPY . .")

    svc2 = tmp_path / "paymentservice"
    svc2.mkdir()
    (svc2 / "main.py").write_text("from flask import Flask\napp = Flask(__name__)")
    (svc2 / "requirements.txt").write_text("flask\n")
    (svc2 / "Dockerfile").write_text("FROM python:3.11\nCOPY . .")

    # Create k8s manifests
    k8s = tmp_path / "kubernetes"
    k8s.mkdir()
    (k8s / "deployment.yaml").write_text(
        "apiVersion: apps/v1\nkind: Deployment\nmetadata:\n  name: checkout\n"
    )

    # Create a config file
    (tmp_path / "docker-compose.yaml").write_text(
        "version: '3'\nservices:\n  web:\n    image: nginx\n"
    )

    # Create documentation
    (tmp_path / "README.md").write_text("# Sample Repo")

    # Create ignored directories
    node_modules = tmp_path / "node_modules"
    node_modules.mkdir()
    (node_modules / "lodash.js").write_text("module.exports = {}")

    git_dir = tmp_path / ".git"
    git_dir.mkdir()
    (git_dir / "config").write_text("[core]")

    return tmp_path


class TestRepoScanner:
    def test_scan_finds_relevant_files(self, sample_repo):
        scanner = RepoScanner()
        result = scanner.scan(str(sample_repo))

        assert isinstance(result, ScanResult)
        assert result.total_files > 0

        # Should find Go, Python, Dockerfile, YAML, MD files
        extensions = {os.path.splitext(f)[1] for f in result.files}
        filenames = {os.path.basename(f) for f in result.files}

        assert ".go" in extensions
        assert ".py" in extensions
        assert ".yaml" in extensions
        assert ".md" in extensions
        assert "Dockerfile" in filenames

    def test_scan_ignores_node_modules(self, sample_repo):
        scanner = RepoScanner()
        result = scanner.scan(str(sample_repo))

        for f in result.files:
            assert "node_modules" not in f

    def test_scan_ignores_git_directory(self, sample_repo):
        scanner = RepoScanner()
        result = scanner.scan(str(sample_repo))

        for f in result.files:
            assert ".git" not in f.split(os.sep)

    def test_scan_builds_folder_structure(self, sample_repo):
        scanner = RepoScanner()
        result = scanner.scan(str(sample_repo))

        assert isinstance(result.folder_structure, dict)
        assert len(result.folder_structure) > 0

    def test_scan_nonexistent_path_raises(self):
        scanner = RepoScanner()
        with pytest.raises(ValueError, match="does not exist"):
            scanner.scan("/nonexistent/path")

    def test_should_include_by_extension(self):
        scanner = RepoScanner()
        assert scanner._should_include("main.py")
        assert scanner._should_include("server.go")
        assert scanner._should_include("app.js")
        assert scanner._should_include("config.yaml")
        assert scanner._should_include("README.md")
        assert not scanner._should_include("image.png")
        assert not scanner._should_include("data.csv")

    def test_should_include_by_filename(self):
        scanner = RepoScanner()
        assert scanner._should_include("Dockerfile")
        assert scanner._should_include("Makefile")

    def test_custom_extensions(self):
        scanner = RepoScanner(include_extensions={".rs", ".toml"})
        assert scanner._should_include("main.rs")
        assert scanner._should_include("Cargo.toml")
        assert not scanner._should_include("main.py")
