"""
Knowledge Extraction Engine

Converts parsed file metadata into a structured repository knowledge model,
detecting services, dependencies, infrastructure, and APIs.
"""

import os
import re
import logging
from dataclasses import dataclass, field
from typing import Optional

logger = logging.getLogger(__name__)


@dataclass
class ServiceInfo:
    """Information about a detected service."""
    name: str
    source: str = ""
    language: str = ""
    path: str = ""
    dependencies: list = field(default_factory=list)
    api_endpoints: list = field(default_factory=list)
    docker_image: str = ""
    env_variables: list = field(default_factory=list)
    description: str = ""


@dataclass
class InfrastructureInfo:
    """Information about infrastructure components."""
    docker_images: list = field(default_factory=list)
    k8s_resources: list = field(default_factory=list)
    cicd_files: list = field(default_factory=list)
    config_files: list = field(default_factory=list)


@dataclass
class DependencyRelation:
    """A dependency relationship between two components."""
    source: str
    target: str
    relation_type: str = ""


@dataclass
class RepositoryKnowledge:
    """Complete knowledge model for a repository."""
    repo_path: str
    repo_name: str = ""
    services: list = field(default_factory=list)
    infrastructure: InfrastructureInfo = field(default_factory=InfrastructureInfo)
    dependencies: list = field(default_factory=list)
    apis: list = field(default_factory=list)
    documentation_files: list = field(default_factory=list)
    languages: list = field(default_factory=list)
    summary: str = ""


class KnowledgeExtractor:
    """Extracts structured knowledge from parsed file metadata."""

    def extract(self, parse_result) -> RepositoryKnowledge:
        """Extract knowledge from a ParseResult."""
        repo_name = os.path.basename(parse_result.repo_path.rstrip("/"))

        knowledge = RepositoryKnowledge(
            repo_path=parse_result.repo_path,
            repo_name=repo_name,
        )

        # Collect raw data from files
        all_service_names = set()
        language_set = set()
        all_endpoints = []
        docker_images = set()
        k8s_resources = []
        cicd_files = []
        config_files = []
        doc_files = []

        service_path_map = {}

        for file_meta in parse_result.files:
            # Collect languages
            if file_meta.language and file_meta.language not in ("yaml", "markdown", "json"):
                language_set.add(file_meta.language)

            # Collect services
            for svc_name in file_meta.service_names:
                all_service_names.add(svc_name)
                if svc_name not in service_path_map:
                    service_path_map[svc_name] = file_meta

            # Collect Docker images
            for img in file_meta.docker_images:
                docker_images.add(img)

            # Collect K8s resources
            for res in file_meta.k8s_resources:
                k8s_resources.append(res)

            # Collect API endpoints
            for endpoint in file_meta.api_endpoints:
                all_endpoints.append({
                    "path": endpoint,
                    "source_file": file_meta.path,
                })

            # Classify files
            if file_meta.file_type == "cicd":
                cicd_files.append(file_meta.path)
            elif file_meta.file_type == "config":
                config_files.append(file_meta.path)
            elif file_meta.file_type == "documentation":
                doc_files.append(file_meta.path)

        # Also detect services from folder structure
        self._detect_services_from_structure(parse_result, all_service_names, service_path_map)

        # Build service objects
        services = []
        for svc_name in sorted(all_service_names):
            source_file = service_path_map.get(svc_name)
            svc = ServiceInfo(
                name=svc_name,
                path=source_file.path if source_file else "",
                language=source_file.language if source_file else "",
                source=source_file.file_type if source_file else "folder",
            )

            # Find endpoints for this service
            for ep in all_endpoints:
                if svc_name.lower() in ep["source_file"].lower():
                    svc.api_endpoints.append(ep["path"])

            # Find docker image
            for img in docker_images:
                if svc_name.lower() in img.lower():
                    svc.docker_image = img
                    break

            # Find env variables
            if source_file:
                svc.env_variables = list(set(source_file.env_variables))

            services.append(svc)

        # Detect dependencies between services
        dependencies = self._detect_dependencies(parse_result, all_service_names)

        # Build infrastructure info
        infrastructure = InfrastructureInfo(
            docker_images=sorted(docker_images),
            k8s_resources=k8s_resources,
            cicd_files=cicd_files,
            config_files=config_files,
        )

        knowledge.services = services
        knowledge.infrastructure = infrastructure
        knowledge.dependencies = dependencies
        knowledge.apis = all_endpoints
        knowledge.documentation_files = doc_files
        knowledge.languages = sorted(language_set)
        knowledge.summary = self._generate_summary(knowledge)

        logger.info(
            "Knowledge extraction complete: %d services, %d dependencies, %d APIs",
            len(services), len(dependencies), len(all_endpoints),
        )
        return knowledge

    def _detect_services_from_structure(self, parse_result, service_names, service_path_map):
        """Detect services from directory structure patterns."""
        service_indicators = {"Dockerfile", "main.py", "main.go", "index.js", "index.ts", "pom.xml", "build.gradle"}
        dir_services = {}

        for file_meta in parse_result.files:
            parts = file_meta.path.split(os.sep)
            if len(parts) >= 2:
                top_dir = parts[0]
                filename = parts[-1]
                if filename in service_indicators:
                    if top_dir not in dir_services:
                        dir_services[top_dir] = file_meta

        for dir_name, file_meta in dir_services.items():
            cleaned = dir_name.lower().replace("-", "").replace("_", "")
            already_found = any(
                s.lower().replace("-", "").replace("_", "") == cleaned
                for s in service_names
            )
            if not already_found:
                service_names.add(dir_name)
                if dir_name not in service_path_map:
                    service_path_map[dir_name] = file_meta

    def _detect_dependencies(self, parse_result, service_names) -> list:
        """Detect dependency relationships between services."""
        dependencies = []
        service_names_lower = {s.lower(): s for s in service_names}

        for file_meta in parse_result.files:
            content_lower = file_meta.content.lower()
            file_service = None

            # Determine which service this file belongs to
            for svc_lower, svc_name in service_names_lower.items():
                if svc_lower in file_meta.path.lower():
                    file_service = svc_name
                    break

            if not file_service:
                continue

            # Look for references to other services
            for svc_lower, svc_name in service_names_lower.items():
                if svc_name == file_service:
                    continue
                if svc_lower in content_lower:
                    dep = DependencyRelation(
                        source=file_service,
                        target=svc_name,
                        relation_type="service-to-service",
                    )
                    if not any(
                        d.source == dep.source and d.target == dep.target
                        for d in dependencies
                    ):
                        dependencies.append(dep)

            # Detect database/queue dependencies
            infra_patterns = {
                "redis": "cache",
                "postgres": "database",
                "mysql": "database",
                "mongodb": "database",
                "rabbitmq": "queue",
                "kafka": "queue",
                "memcached": "cache",
                "elasticsearch": "search",
            }
            for pattern, dep_type in infra_patterns.items():
                if pattern in content_lower:
                    dep = DependencyRelation(
                        source=file_service,
                        target=pattern,
                        relation_type=f"service-to-{dep_type}",
                    )
                    if not any(
                        d.source == dep.source and d.target == dep.target
                        for d in dependencies
                    ):
                        dependencies.append(dep)

        return dependencies

    def _generate_summary(self, knowledge: RepositoryKnowledge) -> str:
        """Generate a brief summary of the repository."""
        parts = [f"Repository: {knowledge.repo_name}"]

        if knowledge.languages:
            parts.append(f"Languages: {', '.join(knowledge.languages)}")

        if knowledge.services:
            parts.append(f"Services: {len(knowledge.services)} detected")
            svc_names = [s.name for s in knowledge.services[:10]]
            parts.append(f"  - {', '.join(svc_names)}")

        if knowledge.infrastructure.k8s_resources:
            parts.append(f"Kubernetes resources: {len(knowledge.infrastructure.k8s_resources)}")

        if knowledge.infrastructure.docker_images:
            parts.append(f"Docker images: {len(knowledge.infrastructure.docker_images)}")

        if knowledge.apis:
            parts.append(f"API endpoints: {len(knowledge.apis)}")

        if knowledge.dependencies:
            parts.append(f"Dependencies: {len(knowledge.dependencies)}")

        return "\n".join(parts)
