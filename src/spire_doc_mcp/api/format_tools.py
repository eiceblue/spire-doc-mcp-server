"""
Format API module for SpireDoc MCP Server.

This module provides formatting-related APIs for the MCP server, including:
- Paragraph formatting
- Format information retrieval
"""

from typing import Dict, Any, Optional
from .base import get_doc_path
from ..core.format import FormatHandler
from ..utils.exceptions import DocumentError


async def format_paragraph(
    document_name: str,
    paragraph_index: int,
    alignment: Optional[str] = None,
    first_line_indent: Optional[float] = None,
    left_indent: Optional[float] = None,
    right_indent: Optional[float] = None,
    line_spacing: Optional[float] = None,
    line_spacing_rule: Optional[str] = None,
    before_spacing: Optional[float] = None,
    after_spacing: Optional[float] = None,
    doc_files_path: str = "./word_files"
) -> Dict[str, Any]:
    """Apply formatting to a paragraph.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph to format
        alignment: Alignment type ('left', 'center', 'right', 'justify')
        first_line_indent: First line indent in points
        left_indent: Left indent in points
        right_indent: Right indent in points
        line_spacing: Line spacing value
        line_spacing_rule: Line spacing rule ('at_least', 'exactly', 'multiple')
        before_spacing: Space before paragraph in points
        after_spacing: Space after paragraph in points
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with formatting details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = FormatHandler(doc_path)
        result = handler.format_paragraph(
            paragraph_index=paragraph_index,
            alignment=alignment,
            first_line_indent=first_line_indent,
            left_indent=left_indent,
            right_indent=right_indent,
            line_spacing=line_spacing,
            line_spacing_rule=line_spacing_rule,
            before_spacing=before_spacing,
            after_spacing=after_spacing
        )
        return {
            "success": True,
            "message": "Paragraph formatting applied successfully",
            "data": result
        }
    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "FORMAT_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index is valid and the formatting parameters are correct."
            }
        }

async def get_paragraph_format(document_name: str, paragraph_index: int, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get paragraph formatting information.
    Args:
        document_name: Name of the document
        paragraph_index: Index of the paragraph
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with paragraph formatting details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = FormatHandler(doc_path)
        result = handler.get_paragraph_format(paragraph_index)
        return {
            "success": True,
            "message": "Paragraph formatting retrieved successfully",
            "data": result
        }
    except DocumentError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "FORMAT_INFO_ERROR",
                "type": "DocumentError",
                "details": str(e),
                "suggestion": "Please check if the paragraph index is valid."
            }
        } 