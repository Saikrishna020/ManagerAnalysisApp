"""
Repository Scanner Module

Clones and scans a git repository, identifying relevant files
and building a folder structure map.
"""

import os
import logging
import shutil
import tempfile
from dataclasses import dataclass, field
from typing import Optional

import git

logger = logging.getLogger(__name__)

INCLUDE_EXTENSIONS = {
    ".py", ".go", ".js", ".ts", ".java",
    ".yaml", ".yml", ".md", ".json",
}

INCLUDE_FILENAMES = {
    "Dockerfile", "Makefile", "Procfile",
    ".env.example", "docker-compose.yml", "docker-compose.yaml",
}

IGNORE_DIRS = {
    "node_modules", "dist", "build", ".git",
    "__pycache__", ".tox", ".mypy_cache",
    "vendor", ".gradle", "target",
}


@dataclass
class ScanResult:
    """Result of scanning a repository."""
    repo_path: str
    files: list = field(default_factory=list)
    folder_structure: dict = field(default_factory=dict)
    total_files: int = 0


class RepoScanner:
    """Scans a git repository and identifies relevant files."""

    def __init__(
        self,
        include_extensions: Optional[set] = None,
        include_filenames: Optional[set] = None,
        ignore_dirs: Optional[set] = None,
    ):
        self.include_extensions = include_extensions or INCLUDE_EXTENSIONS
        self.include_filenames = include_filenames or INCLUDE_FILENAMES
        self.ignore_dirs = ignore_dirs or IGNORE_DIRS

    def clone_repo(self, repo_url: str, target_dir: Optional[str] = None) -> str:
        """Clone a git repository to a local directory."""
        if target_dir is None:
            target_dir = tempfile.mkdtemp(prefix="kt_ai_")

        if os.path.exists(target_dir) and os.listdir(target_dir):
            logger.info("Directory %s already exists, removing first", target_dir)
            shutil.rmtree(target_dir)

        logger.info("Cloning repository %s to %s", repo_url, target_dir)
        git.Repo.clone_from(repo_url, target_dir, depth=1)
        logger.info("Repository cloned successfully")
        return target_dir

    def scan(self, repo_path: str) -> ScanResult:
        """Scan a local repository directory for relevant files."""
        if not os.path.isdir(repo_path):
            raise ValueError(f"Repository path does not exist: {repo_path}")

        logger.info("Scanning repository at %s", repo_path)
        files = []
        folder_structure = {}

        for root, dirs, filenames in os.walk(repo_path):
            # Filter out ignored directories in-place
            dirs[:] = [d for d in dirs if d not in self.ignore_dirs]

            rel_root = os.path.relpath(root, repo_path)
            if rel_root == ".":
                rel_root = ""

            current_files = []
            for filename in filenames:
                if self._should_include(filename):
                    rel_path = os.path.join(rel_root, filename) if rel_root else filename
                    files.append(rel_path)
                    current_files.append(filename)

            if current_files:
                key = rel_root if rel_root else "."
                folder_structure[key] = current_files

        result = ScanResult(
            repo_path=repo_path,
            files=sorted(files),
            folder_structure=folder_structure,
            total_files=len(files),
        )
        logger.info("Scan complete: found %d relevant files", result.total_files)
        return result

    def _should_include(self, filename: str) -> bool:
        """Check if a file should be included based on extension or name."""
        if filename in self.include_filenames:
            return True
        _, ext = os.path.splitext(filename)
        return ext.lower() in self.include_extensions

    def clone_and_scan(self, repo_url: str, target_dir: Optional[str] = None) -> ScanResult:
        """Clone a repository and scan it in one step."""
        local_path = self.clone_repo(repo_url, target_dir)
        return self.scan(local_path)
