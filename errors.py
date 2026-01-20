"""Common application exceptions."""
from __future__ import annotations


class MissingDependencyError(RuntimeError):
    """Raised when a required runtime dependency is missing."""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message

    def __str__(self) -> str:  # pragma: no cover - mirrors base behaviour
        return self.message


__all__ = ["MissingDependencyError"]
