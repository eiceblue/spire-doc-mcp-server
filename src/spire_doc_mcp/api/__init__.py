"""
SpireDoc MCP API Package.

This package provides API endpoints for document operations:
- Document creation and manipulation
- Paragraph formatting and management
- Table operations and styling
- Document format conversion
"""

from .base import get_doc_path

__all__ = [
    # Base utilities
    "get_doc_path",
]
