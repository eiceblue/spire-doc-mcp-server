"""
Base API module.

This module provides base functionality and utilities for all API modules:
- Document path handling
"""

import os
from typing import Dict, Any
from ..utils.exceptions import DocumentError


def get_doc_path(filename: str, doc_files_path: str = None) -> str:
    """Get the full path to a document file.

    Args:
        filename: Name of the document file
        doc_files_path: Optional custom documents directory path

    Returns:
        Full path to the document file

    Raises:
        DocumentError: If the document path is invalid
    """
    # Use provided path or default to documents directory
    if doc_files_path is None:
        doc_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "documents")
    else:
        doc_dir = doc_files_path

    # Create documents directory if it doesn't exist
    if not os.path.exists(doc_dir):
        os.makedirs(doc_dir)

    # Construct full path
    doc_path = os.path.join(doc_dir, filename)

    # Validate path
    if not os.path.dirname(doc_path) == doc_dir:
        raise DocumentError("Invalid document path")

    return doc_path
