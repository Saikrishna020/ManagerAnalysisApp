"""Tests for the Knowledge Extraction Engine."""

import os
import pytest
from dataclasses import dataclass, field

from extraction.knowledge_extractor import (
    KnowledgeExtractor,
    RepositoryKnowledge,
    ServiceInfo,
    DependencyRelation,
)
from parser.file_parser import FileMetadata, ParseResult


def _make_parse_result(repo_path, files):
    """Helper to create a ParseResult from FileMetadata list."""
    return ParseResult(repo_path=repo_path, files=files, total_parsed=len(files))


class TestKnowledgeExtractor:
    def test_extract_services_from_dockerfile(self):
        files = [
            FileMetadata(
                path=os.path.join("checkoutservice", "Dockerfile"),
                content="FROM golang:1.19\nCOPY . .",
                file_type="dockerfile",
                language="",
                service_names=["checkoutservice"],
                docker_images=["golang:1.19"],
            ),
            FileMetadata(
                path=os.path.join("paymentservice", "Dockerfile"),
                content="FROM python:3.11\nCOPY . .",
                file_type="dockerfile",
                language="",
                service_names=["paymentservice"],
                docker_images=["python:3.11"],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert len(knowledge.services) >= 2
        names = [s.name for s in knowledge.services]
        assert "checkoutservice" in names
        assert "paymentservice" in names

    def test_extract_services_from_kubernetes(self):
        files = [
            FileMetadata(
                path=os.path.join("k8s", "deployment.yaml"),
                content="apiVersion: apps/v1\nkind: Deployment\nmetadata:\n  name: frontend",
                file_type="kubernetes",
                language="yaml",
                service_names=["frontend"],
                k8s_resources=[{"kind": "Deployment", "name": "frontend"}],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        names = [s.name for s in knowledge.services]
        assert "frontend" in names
        assert len(knowledge.infrastructure.k8s_resources) == 1

    def test_extract_services_from_docker_compose(self):
        files = [
            FileMetadata(
                path="docker-compose.yml",
                content="version: '3'\nservices:\n  web:\n    image: nginx\n  redis:\n    image: redis:7",
                file_type="docker-compose",
                language="yaml",
                service_names=["web", "redis"],
                docker_images=["nginx", "redis:7"],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        names = [s.name for s in knowledge.services]
        assert "web" in names
        assert "redis" in names

    def test_extract_dependencies(self):
        files = [
            FileMetadata(
                path=os.path.join("checkoutservice", "main.py"),
                content="import requests\nresponse = requests.get('http://paymentservice/pay')\nredis_conn = connect('redis')",
                file_type="source",
                language="python",
                service_names=[],
            ),
            FileMetadata(
                path=os.path.join("checkoutservice", "Dockerfile"),
                content="FROM python:3.11",
                file_type="dockerfile",
                language="",
                service_names=["checkoutservice"],
            ),
            FileMetadata(
                path=os.path.join("paymentservice", "Dockerfile"),
                content="FROM python:3.11",
                file_type="dockerfile",
                language="",
                service_names=["paymentservice"],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        # Should detect checkout → payment dependency
        dep_pairs = [(d.source, d.target) for d in knowledge.dependencies]
        assert ("checkoutservice", "paymentservice") in dep_pairs

        # Should detect checkout → redis dependency
        assert any(
            d.source == "checkoutservice" and d.target == "redis"
            for d in knowledge.dependencies
        )

    def test_extract_languages(self):
        files = [
            FileMetadata(
                path="main.py",
                content="print('hello')",
                file_type="source",
                language="python",
            ),
            FileMetadata(
                path="main.go",
                content="package main",
                file_type="source",
                language="go",
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert "python" in knowledge.languages
        assert "go" in knowledge.languages

    def test_extract_apis(self):
        files = [
            FileMetadata(
                path="server.py",
                content="@app.route('/api/items')",
                file_type="source",
                language="python",
                api_endpoints=["/api/items"],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert len(knowledge.apis) >= 1
        assert any(a["path"] == "/api/items" for a in knowledge.apis)

    def test_extract_documentation_files(self):
        files = [
            FileMetadata(
                path="README.md",
                content="# Project",
                file_type="documentation",
                language="markdown",
            ),
            FileMetadata(
                path="docs/guide.md",
                content="# Guide",
                file_type="documentation",
                language="markdown",
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert "README.md" in knowledge.documentation_files
        assert "docs/guide.md" in knowledge.documentation_files

    def test_generate_summary(self):
        files = [
            FileMetadata(
                path=os.path.join("svc", "Dockerfile"),
                content="FROM node:18",
                file_type="dockerfile",
                language="",
                service_names=["svc"],
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert "repo" in knowledge.summary.lower()
        assert "Services" in knowledge.summary

    def test_empty_repository(self):
        parse_result = _make_parse_result("/repo", [])
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        assert knowledge.repo_name == "repo"
        assert len(knowledge.services) == 0
        assert len(knowledge.dependencies) == 0

    def test_detect_services_from_folder_structure(self):
        files = [
            FileMetadata(
                path=os.path.join("frontend", "index.js"),
                content="const express = require('express');",
                file_type="source",
                language="javascript",
            ),
            FileMetadata(
                path=os.path.join("backend", "main.py"),
                content="from flask import Flask",
                file_type="source",
                language="python",
            ),
        ]
        parse_result = _make_parse_result("/repo", files)
        extractor = KnowledgeExtractor()
        knowledge = extractor.extract(parse_result)

        names = [s.name for s in knowledge.services]
        assert "frontend" in names
        assert "backend" in names
