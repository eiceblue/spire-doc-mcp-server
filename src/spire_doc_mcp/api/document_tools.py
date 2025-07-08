"""
Document API module for SpireDoc MCP Server.

This module provides document-related APIs for the MCP server, including:
- Document creation and management
- Document protection and security
- Document content operations
- Document search and replace
- Document merge
- Document watermark
"""

from typing import Dict, Any, Optional
from .base import get_doc_path
from ..core.document import DocumentHandler
from ..utils.exceptions import DocumentError
from spire.doc import WatermarkLayout


async def create_document(document_name: str, template_path: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Create a new document.

    Args:
        document_name: Name of the document to create
        template_path: Optional path to a template document
        doc_files_path: Path to the documents directory

    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with document details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = DocumentHandler(doc_path, create_if_not_exists=True)
        result = handler.create_document(template_path)

        return {"success": True, "message": "Document created successfully", "data": result}

    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "DOC_CREATE_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the document name is valid and you have write permissions.",
            },
        }


async def set_document_protection(
    document_name: str, protection_level: str, password: Optional[str] = None, doc_files_path: str = "./word_files"
) -> Dict[str, Any]:
    """Set document protection level.

    Args:
        document_name: Name of the document
        protection_level: Level of protection to apply
        password: Optional password for protection
        doc_files_path: Path to the documents directory

    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with protection details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = DocumentHandler(doc_path)
        result = handler.protect_document(protection_level, password)

        return {"success": True, "message": "Document protection set successfully", "data": result}

    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "DOC_PROTECT_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the protection level is valid and you have the necessary permissions.",
            },
        }


async def find_and_replace(
    document_name: str, find_text: str, replace_text: str, match_case: bool = False, match_whole_word: bool = False, doc_files_path: str = "./word_files"
) -> Dict[str, Any]:
    """Find and replace text in the document.

    Args:
        document_name: Name of the document
        find_text: Text to find
        replace_text: Text to replace with
        match_case: Whether to match case
        match_whole_word: Whether to match whole words only
        doc_files_path: Path to the documents directory

    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with replacement details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = DocumentHandler(doc_path)
        result = handler.find_and_replace(
            find_text=find_text, replace_text=replace_text, match_case=match_case, match_whole_word=match_whole_word
        )

        return {"success": True, "message": "Text replacement completed successfully", "data": result}

    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "DOC_REPLACE_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the document exists and the search parameters are valid.",
            },
        }


async def merge_documents(original_document: str, merge_document: str, output_document: str, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Merge another document into the current document.

    Args:
        original_document: Name of the original document
        merge_document: Name of the document to merge
        output_document: Name of the output merged document
        doc_files_path: Path to the documents directory

    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with merge details
    """
    try:
        original_path = get_doc_path(original_document, doc_files_path)
        merge_path = get_doc_path(merge_document, doc_files_path)
        output_path = get_doc_path(output_document, doc_files_path)

        handler = DocumentHandler(original_path)
        result = handler.merge_documents(merge_path=merge_path, output_path=output_path)

        return {"success": True, "message": "Documents merged successfully", "data": result}

    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "DOC_MERGE_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if both documents exist and are accessible.",
            },
        }


async def add_text_watermark(
    document_name: str,
    text: str,
    font_size: int = 65,
    color: str = "Red",
    layout: Any = WatermarkLayout.Diagonal,
    doc_files_path: str = "./word_files"
) -> Dict[str, Any]:
    """Add text watermark to the document.

    Args:
        document_name: Name of the document
        text: Text to display as watermark
        font_size: Font size of the watermark text
        color: Color of the watermark text
        layout: Layout of the watermark
        doc_files_path: Path to the documents directory

    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with watermark details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = DocumentHandler(doc_path)
        result = handler.add_text_watermark(
            text=text,
            font_size=font_size,
            color=color,
            layout=layout
        )

        return {"success": True, "message": "Text watermark added successfully", "data": result}

    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "DOC_WATERMARK_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the document exists and the watermark parameters are valid.",
            },
        }
