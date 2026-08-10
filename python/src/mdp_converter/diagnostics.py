"""Parse diagnostics collection."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ParseDiagnostics:
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def error(self, message: str) -> None:
        if message not in self.errors:
            self.errors.append(message)

    def save(self, path: str | Path) -> None:
        data = {"warnings": self.warnings, "errors": self.errors}
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
