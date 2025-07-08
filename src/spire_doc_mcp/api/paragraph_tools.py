"""
Paragraph API module for SpireDoc MCP Server.

This module provides paragraph-related APIs for the MCP server, including:
- Paragraph creation and management
- Paragraph content operations
- Paragraph information retrieval
"""

from typing import Dict, Any, Optional
from .base import get_doc_path
from ..core.paragraph import ParagraphHandler
from ..utils.exceptions import ParagraphError



async def get_paragraph_text(document_name: str, paragraph_index: int, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get the text content of a paragraph.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with paragraph text information
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ParagraphHandler(doc_path)
        result = handler.get_paragraph_text(paragraph_index, section_index)
        return {
            "success": True,
            "message": "Paragraph text retrieved successfully",
            "data": result
        }
    except ParagraphError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "PARAGRAPH_TEXT_ERROR",
                "type": "ParagraphError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index and section index are valid."
            }
        }

async def delete_paragraph(document_name: str, paragraph_index: int, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Delete a paragraph from the document.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph to delete
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with deletion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ParagraphHandler(doc_path)
        result = handler.delete_paragraph(paragraph_index, section_index)
        return {
            "success": True,
            "message": "Paragraph deleted successfully",
            "data": result
        }
    except ParagraphError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "PARAGRAPH_DELETE_ERROR",
                "type": "ParagraphError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index and section index are valid."
            }
        }

async def get_paragraph_info(document_name: str, paragraph_index: int, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get detailed information about a paragraph.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with paragraph information
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ParagraphHandler(doc_path)
        result = handler.get_paragraphinfo(paragraph_index, section_index)
        return {
            "success": True,
            "message": "Paragraph information retrieved successfully",
            "data": result
        }
    except ParagraphError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "PARAGRAPH_INFO_ERROR",
                "type": "ParagraphError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index and section index are valid."
            }
        }

async def add_paragraph(document_name: str, text: str, section_index: int = 0, paragraph_index: Optional[int] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Add a new paragraph to the document.
    Args:
        document_name: Name of the document
        text: Text content for the new paragraph
        section_index: Index of the section (default: 0)
        paragraph_index: Index where to insert the paragraph (optional, adds to end if not specified)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with paragraph creation details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ParagraphHandler(doc_path)
        result = handler.add_paragraph(text, section_index, paragraph_index)
        return {
            "success": True,
            "message": "Paragraph added successfully",
            "data": result
        }
    except ParagraphError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "PARAGRAPH_ADD_ERROR",
                "type": "ParagraphError",
                "details": str(e),
                "suggestion": "Please check if the section index and paragraph index are valid."
            }
        }

async def update_paragraph_text(document_name: str, paragraph_index: int, new_text: str, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Update the text content of a paragraph.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph to update
        new_text: New text content
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with update details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ParagraphHandler(doc_path)
        result = handler.update_paragraph_text(paragraph_index, new_text, section_index)
        return {
            "success": True,
            "message": "Paragraph text updated successfully",
            "data": result
        }
    except ParagraphError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "PARAGRAPH_UPDATE_ERROR",
                "type": "ParagraphError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index and section index are valid."
            }
        } 