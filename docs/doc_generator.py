"""
Documentation Generator Module

Generates structured documentation files from the repository knowledge model.
Supports optional LLM-enhanced generation using OpenAI or other providers.
"""

import os
import logging
from dataclasses import dataclass, field
from typing import Optional

logger = logging.getLogger(__name__)


@dataclass
class GeneratedDocs:
    """Collection of generated documentation files."""
    readme: str = ""
    architecture: str = ""
    services: str = ""
    deployment: str = ""
    output_dir: str = ""


class DocGenerator:
    """Generates documentation from repository knowledge."""

    def __init__(self, llm_client=None):
        """
        Initialize the documentation generator.

        Args:
            llm_client: Optional LangChain LLM instance for enhanced generation.
                        If None, uses template-based generation.
        """
        self.llm_client = llm_client

    def generate(self, knowledge, output_dir: str = "generated_docs") -> GeneratedDocs:
        """Generate documentation from a RepositoryKnowledge model."""
        os.makedirs(output_dir, exist_ok=True)

        if self.llm_client:
            docs = self._generate_with_llm(knowledge)
        else:
            docs = self._generate_from_templates(knowledge)

        docs.output_dir = output_dir

        # Write files
        file_map = {
            "README.md": docs.readme,
            "architecture.md": docs.architecture,
            "services.md": docs.services,
            "deployment.md": docs.deployment,
        }

        for filename, content in file_map.items():
            filepath = os.path.join(output_dir, filename)
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(content)
            logger.info("Generated %s", filepath)

        return docs

    def _generate_with_llm(self, knowledge) -> GeneratedDocs:
        """Generate documentation using an LLM."""
        context = self._build_context(knowledge)

        docs = GeneratedDocs()
        docs.readme = self._llm_generate_section(
            context,
            "Generate a comprehensive README.md for this repository. Include: "
            "project overview, list of services, quick start guide, and architecture summary.",
        )
        docs.architecture = self._llm_generate_section(
            context,
            "Generate an architecture.md document. Include: "
            "system architecture overview, component interactions, data flow, "
            "and technology stack.",
        )
        docs.services = self._llm_generate_section(
            context,
            "Generate a services.md document. Include: "
            "detailed description of each service, its responsibilities, "
            "API endpoints, dependencies, and configuration.",
        )
        docs.deployment = self._llm_generate_section(
            context,
            "Generate a deployment.md document. Include: "
            "deployment instructions, infrastructure requirements, "
            "Docker and Kubernetes configuration, and environment variables.",
        )
        return docs

    def _llm_generate_section(self, context: str, instruction: str) -> str:
        """Generate a documentation section using the LLM."""
        prompt = (
            "You are a software architect generating documentation.\n\n"
            f"Repository Knowledge:\n{context}\n\n"
            f"Task: {instruction}\n\n"
            "Generate well-structured Markdown documentation."
        )
        try:
            response = self.llm_client.invoke(prompt)
            if hasattr(response, "content"):
                return response.content
            return str(response)
        except Exception as e:
            logger.error("LLM generation failed: %s. Falling back to templates.", e)
            return ""

    def _generate_from_templates(self, knowledge) -> GeneratedDocs:
        """Generate documentation using templates (no LLM required)."""
        docs = GeneratedDocs()
        docs.readme = self._template_readme(knowledge)
        docs.architecture = self._template_architecture(knowledge)
        docs.services = self._template_services(knowledge)
        docs.deployment = self._template_deployment(knowledge)
        return docs

    def _build_context(self, knowledge) -> str:
        """Build a context string from the knowledge model for LLM prompts."""
        lines = [f"# Repository: {knowledge.repo_name}\n"]

        if knowledge.summary:
            lines.append(f"## Summary\n{knowledge.summary}\n")

        if knowledge.services:
            lines.append("## Services")
            for svc in knowledge.services:
                lines.append(f"- **{svc.name}**: language={svc.language}, path={svc.path}")
                if svc.api_endpoints:
                    lines.append(f"  - Endpoints: {', '.join(svc.api_endpoints)}")
                if svc.dependencies:
                    lines.append(f"  - Dependencies: {', '.join(svc.dependencies)}")
            lines.append("")

        if knowledge.dependencies:
            lines.append("## Dependencies")
            for dep in knowledge.dependencies:
                lines.append(f"- {dep.source} → {dep.target} ({dep.relation_type})")
            lines.append("")

        if knowledge.infrastructure.docker_images:
            lines.append("## Docker Images")
            for img in knowledge.infrastructure.docker_images:
                lines.append(f"- {img}")
            lines.append("")

        if knowledge.infrastructure.k8s_resources:
            lines.append("## Kubernetes Resources")
            for res in knowledge.infrastructure.k8s_resources:
                lines.append(f"- {res['kind']}: {res['name']}")
            lines.append("")

        if knowledge.apis:
            lines.append("## API Endpoints")
            for api in knowledge.apis:
                lines.append(f"- {api['path']} (from {api['source_file']})")
            lines.append("")

        return "\n".join(lines)

    def _template_readme(self, knowledge) -> str:
        """Generate README.md from template."""
        lines = [
            f"# {knowledge.repo_name}\n",
            "## Overview\n",
            f"This repository contains {len(knowledge.services)} service(s) "
            f"built with {', '.join(knowledge.languages) if knowledge.languages else 'various technologies'}.\n",
        ]

        if knowledge.services:
            lines.append("## Services\n")
            lines.append("| Service | Language | Path |")
            lines.append("|---------|----------|------|")
            for svc in knowledge.services:
                lines.append(f"| {svc.name} | {svc.language} | {svc.path} |")
            lines.append("")

        if knowledge.dependencies:
            lines.append("## Dependencies\n")
            for dep in knowledge.dependencies:
                lines.append(f"- {dep.source} → {dep.target}")
            lines.append("")

        lines.extend([
            "## Quick Start\n",
            "Refer to [deployment.md](deployment.md) for setup instructions.\n",
            "## Documentation\n",
            "- [Architecture](architecture.md)",
            "- [Services](services.md)",
            "- [Deployment](deployment.md)",
            "",
        ])
        return "\n".join(lines)

    def _template_architecture(self, knowledge) -> str:
        """Generate architecture.md from template."""
        lines = [
            f"# Architecture: {knowledge.repo_name}\n",
            "## System Overview\n",
            f"This system consists of {len(knowledge.services)} service(s).\n",
        ]

        if knowledge.languages:
            lines.append("## Technology Stack\n")
            for lang in knowledge.languages:
                lines.append(f"- {lang}")
            lines.append("")

        if knowledge.services:
            lines.append("## Component Diagram\n")
            lines.append("```")
            for svc in knowledge.services:
                lines.append(f"  [{svc.name}]")
            lines.append("```\n")

        if knowledge.dependencies:
            lines.append("## Component Interactions\n")
            for dep in knowledge.dependencies:
                lines.append(f"- **{dep.source}** → **{dep.target}** ({dep.relation_type})")
            lines.append("")

        if knowledge.infrastructure.docker_images:
            lines.append("## Docker Images\n")
            for img in knowledge.infrastructure.docker_images:
                lines.append(f"- `{img}`")
            lines.append("")

        return "\n".join(lines)

    def _template_services(self, knowledge) -> str:
        """Generate services.md from template."""
        lines = [
            f"# Services: {knowledge.repo_name}\n",
            f"Total services detected: {len(knowledge.services)}\n",
        ]

        for svc in knowledge.services:
            lines.append(f"## {svc.name}\n")
            lines.append(f"- **Language**: {svc.language or 'N/A'}")
            lines.append(f"- **Path**: `{svc.path or 'N/A'}`")
            lines.append(f"- **Source**: {svc.source or 'N/A'}")

            if svc.docker_image:
                lines.append(f"- **Docker Image**: `{svc.docker_image}`")

            if svc.api_endpoints:
                lines.append("\n### API Endpoints\n")
                for ep in svc.api_endpoints:
                    lines.append(f"- `{ep}`")

            if svc.env_variables:
                lines.append("\n### Environment Variables\n")
                for var in svc.env_variables:
                    lines.append(f"- `{var}`")

            lines.append("")

        return "\n".join(lines)

    def _template_deployment(self, knowledge) -> str:
        """Generate deployment.md from template."""
        lines = [
            f"# Deployment: {knowledge.repo_name}\n",
        ]

        if knowledge.infrastructure.docker_images:
            lines.append("## Docker\n")
            lines.append("The following Docker images are used:\n")
            for img in knowledge.infrastructure.docker_images:
                lines.append(f"- `{img}`")
            lines.append("")

        if knowledge.infrastructure.k8s_resources:
            lines.append("## Kubernetes\n")
            lines.append("The following Kubernetes resources are defined:\n")
            lines.append("| Kind | Name |")
            lines.append("|------|------|")
            for res in knowledge.infrastructure.k8s_resources:
                lines.append(f"| {res['kind']} | {res['name']} |")
            lines.append("")

        if knowledge.infrastructure.cicd_files:
            lines.append("## CI/CD\n")
            lines.append("CI/CD configuration files:\n")
            for f in knowledge.infrastructure.cicd_files:
                lines.append(f"- `{f}`")
            lines.append("")

        if knowledge.infrastructure.config_files:
            lines.append("## Configuration Files\n")
            for f in knowledge.infrastructure.config_files:
                lines.append(f"- `{f}`")
            lines.append("")

        # Collect all env vars across services
        all_env = set()
        for svc in knowledge.services:
            all_env.update(svc.env_variables)
        if all_env:
            lines.append("## Environment Variables\n")
            lines.append("| Variable | Description |")
            lines.append("|----------|-------------|")
            for var in sorted(all_env):
                lines.append(f"| `{var}` | - |")
            lines.append("")

        return "\n".join(lines)
