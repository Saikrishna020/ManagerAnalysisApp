"""
File Parser Module

Reads and parses relevant files from a scanned repository,
extracting structured metadata such as service names, imports,
API endpoints, Docker images, Kubernetes deployments, and
environment variables.
"""

import os
import re
import logging
from dataclasses import dataclass, field
from typing import Optional

import yaml

logger = logging.getLogger(__name__)


@dataclass
class FileMetadata:
    """Metadata extracted from a single file."""
    path: str
    content: str = ""
    file_type: str = ""
    language: str = ""
    service_names: list = field(default_factory=list)
    imports: list = field(default_factory=list)
    api_endpoints: list = field(default_factory=list)
    docker_images: list = field(default_factory=list)
    k8s_resources: list = field(default_factory=list)
    env_variables: list = field(default_factory=list)


@dataclass
class ParseResult:
    """Result of parsing all files in a repository."""
    repo_path: str
    files: list = field(default_factory=list)
    total_parsed: int = 0


LANGUAGE_MAP = {
    ".py": "python",
    ".go": "go",
    ".js": "javascript",
    ".ts": "typescript",
    ".java": "java",
    ".yaml": "yaml",
    ".yml": "yaml",
    ".md": "markdown",
    ".json": "json",
}


class FileParser:
    """Parses repository files and extracts metadata."""

    def __init__(self, max_file_size: int = 1_000_000):
        self.max_file_size = max_file_size

    def parse_files(self, repo_path: str, file_list: list) -> ParseResult:
        """Parse a list of files from a repository."""
        parsed_files = []

        for rel_path in file_list:
            full_path = os.path.join(repo_path, rel_path)
            if not os.path.isfile(full_path):
                logger.warning("File not found: %s", full_path)
                continue

            file_size = os.path.getsize(full_path)
            if file_size > self.max_file_size:
                logger.warning("Skipping large file (%d bytes): %s", file_size, rel_path)
                continue

            metadata = self._parse_file(full_path, rel_path)
            if metadata:
                parsed_files.append(metadata)

        result = ParseResult(
            repo_path=repo_path,
            files=parsed_files,
            total_parsed=len(parsed_files),
        )
        logger.info("Parsed %d files from %s", result.total_parsed, repo_path)
        return result

    def _parse_file(self, full_path: str, rel_path: str) -> Optional[FileMetadata]:
        """Parse a single file and extract metadata."""
        try:
            with open(full_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
        except Exception as e:
            logger.error("Error reading file %s: %s", rel_path, e)
            return None

        filename = os.path.basename(rel_path)
        _, ext = os.path.splitext(filename)
        language = LANGUAGE_MAP.get(ext.lower(), "")

        metadata = FileMetadata(
            path=rel_path,
            content=content,
            file_type=self._detect_file_type(rel_path, filename),
            language=language,
        )

        # Extract metadata based on file type
        if filename == "Dockerfile" or filename.startswith("Dockerfile"):
            self._parse_dockerfile(metadata)
        elif ext.lower() in (".yaml", ".yml"):
            self._parse_yaml(metadata)
        elif ext.lower() == ".py":
            self._parse_python(metadata)
        elif ext.lower() in (".go",):
            self._parse_go(metadata)
        elif ext.lower() in (".js", ".ts"):
            self._parse_javascript(metadata)
        elif ext.lower() == ".java":
            self._parse_java(metadata)

        return metadata

    def _detect_file_type(self, rel_path: str, filename: str) -> str:
        """Detect the type of a file based on its path and name."""
        lower_path = rel_path.lower()

        if filename == "Dockerfile" or filename.startswith("Dockerfile"):
            return "dockerfile"
        if "docker-compose" in filename.lower():
            return "docker-compose"
        if "kubernetes" in lower_path or "k8s" in lower_path:
            return "kubernetes"
        if any(d in lower_path for d in (".github/workflows", "ci", "pipeline", "jenkinsfile")):
            return "cicd"
        if filename.lower().endswith(".md"):
            return "documentation"
        if filename.lower() in ("requirements.txt", "package.json", "go.mod", "pom.xml", "build.gradle"):
            return "dependency"
        if filename.lower() in (".env", ".env.example"):
            return "config"

        _, ext = os.path.splitext(filename)
        if ext.lower() in (".yaml", ".yml"):
            return "config"
        return "source"

    def _parse_dockerfile(self, metadata: FileMetadata):
        """Extract metadata from Dockerfiles."""
        content = metadata.content
        images = re.findall(r"FROM\s+(\S+)", content)
        metadata.docker_images.extend(images)

        env_vars = re.findall(r"ENV\s+(\w+)", content)
        metadata.env_variables.extend(env_vars)

        # Try to detect service name from path
        parent_dir = os.path.basename(os.path.dirname(metadata.path))
        if parent_dir and parent_dir != ".":
            metadata.service_names.append(parent_dir)

    def _parse_yaml(self, metadata: FileMetadata):
        """Extract metadata from YAML files."""
        try:
            docs = list(yaml.safe_load_all(metadata.content))
        except yaml.YAMLError:
            logger.debug("Could not parse YAML: %s", metadata.path)
            return

        for doc in docs:
            if not isinstance(doc, dict):
                continue

            kind = doc.get("kind", "")
            name = ""

            # Kubernetes resources
            if kind in ("Deployment", "Service", "StatefulSet", "DaemonSet", "Job", "CronJob", "Pod"):
                metadata.file_type = "kubernetes"
                name = doc.get("metadata", {}).get("name", "")
                if name:
                    metadata.k8s_resources.append({"kind": kind, "name": name})

                # Extract container images
                spec = doc.get("spec", {})
                template_spec = spec.get("template", {}).get("spec", {}) if isinstance(spec, dict) else {}
                containers = template_spec.get("containers", []) if isinstance(template_spec, dict) else []
                for container in containers:
                    if isinstance(container, dict):
                        img = container.get("image", "")
                        if img:
                            metadata.docker_images.append(img)

                        # Extract env variables
                        for env in container.get("env", []):
                            if isinstance(env, dict) and "name" in env:
                                metadata.env_variables.append(env["name"])

                # Extract service names from Deployments
                if kind == "Deployment" and name:
                    metadata.service_names.append(name)

            # Docker Compose services
            if "services" in doc and isinstance(doc["services"], dict):
                metadata.file_type = "docker-compose"
                for svc_name, svc_config in doc["services"].items():
                    metadata.service_names.append(svc_name)
                    if isinstance(svc_config, dict):
                        img = svc_config.get("image", "")
                        if img:
                            metadata.docker_images.append(img)
                        for env in svc_config.get("environment", []):
                            if isinstance(env, str) and "=" in env:
                                metadata.env_variables.append(env.split("=")[0])

    def _parse_python(self, metadata: FileMetadata):
        """Extract metadata from Python files."""
        content = metadata.content

        imports = re.findall(r"^(?:from\s+(\S+)\s+import|import\s+(\S+))", content, re.MULTILINE)
        for imp in imports:
            module = imp[0] or imp[1]
            metadata.imports.append(module)

        # Detect Flask/FastAPI endpoints
        endpoints = re.findall(
            r'@\w+\.(?:route|get|post|put|delete|patch)\s*\(\s*["\']([^"\']+)',
            content,
        )
        metadata.api_endpoints.extend(endpoints)

        # Detect environment variables
        env_vars = re.findall(r'os\.(?:environ|getenv)\s*[\[\(]\s*["\'](\w+)', content)
        metadata.env_variables.extend(env_vars)

    def _parse_go(self, metadata: FileMetadata):
        """Extract metadata from Go files."""
        content = metadata.content

        imports = re.findall(r'"([^"]+)"', content)
        metadata.imports.extend(imports)

        # Detect HTTP endpoints
        endpoints = re.findall(r'(?:HandleFunc|Handle|GET|POST|PUT|DELETE)\s*\(\s*"([^"]+)"', content)
        metadata.api_endpoints.extend(endpoints)

        # Detect environment variables
        env_vars = re.findall(r'os\.Getenv\s*\(\s*"(\w+)"', content)
        metadata.env_variables.extend(env_vars)

    def _parse_javascript(self, metadata: FileMetadata):
        """Extract metadata from JavaScript/TypeScript files."""
        content = metadata.content

        imports = re.findall(r"(?:require|from)\s*\(\s*['\"]([^'\"]+)", content)
        imports += re.findall(r"from\s+['\"]([^'\"]+)", content)
        metadata.imports.extend(list(set(imports)))

        # Detect Express/REST endpoints
        endpoints = re.findall(r'\.(?:get|post|put|delete|patch)\s*\(\s*["\']([^"\']+)', content)
        metadata.api_endpoints.extend(endpoints)

        # Detect environment variables
        env_vars = re.findall(r'process\.env\.(\w+)', content)
        metadata.env_variables.extend(env_vars)

    def _parse_java(self, metadata: FileMetadata):
        """Extract metadata from Java files."""
        content = metadata.content

        imports = re.findall(r"^import\s+([\w.]+);", content, re.MULTILINE)
        metadata.imports.extend(imports)

        # Detect Spring endpoints
        endpoints = re.findall(
            r'@(?:GetMapping|PostMapping|PutMapping|DeleteMapping|RequestMapping)\s*\(\s*(?:value\s*=\s*)?["\']([^"\']+)',
            content,
        )
        metadata.api_endpoints.extend(endpoints)
