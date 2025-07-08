"""
Conversion API module for SpireDoc MCP Server.

This module provides document conversion APIs for the MCP server, including:
- Document format conversion
- Conversion status tracking
- Conversion history
"""

from typing import Dict, Any, Optional
from .base import get_doc_path
from ..core.conversion import ConversionHandler
from ..utils.exceptions import ConversionError


async def convert_document(document_name: str, target_format: str, output_path: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert the document to a different format.
    Args:
        document_name: Name of the document to convert
        target_format: Target format for conversion
        output_path: Optional output path for the converted document
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_document(target_format, output_path)
        return {
            "success": True,
            "message": "Document converted successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and the target format is supported."
            }
        }

async def get_conversion_status(document_name: str, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get the current conversion status.
    Args:
        document_name: Name of the document
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion status information
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.get_conversion_status()
        return {
            "success": True,
            "message": "Conversion status retrieved successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_STATUS_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists."
            }
        }

async def get_conversion_history(document_name: str, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get the conversion history.
    Args:
        document_name: Name of the document
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion history
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.get_conversion_history()
        return {
            "success": True,
            "message": "Conversion history retrieved successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_HISTORY_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists."
            }
        }

async def convert_to_pdf(document_name: str, output_name: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert a Word document to PDF format.
    Args:
        document_name: Name of the document to convert
        output_name: Optional name for the output PDF file
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_to_pdf(output_name)
        return {
            "success": True,
            "message": "Document converted to PDF successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_PDF_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and is accessible."
            }
        }

async def convert_to_html(document_name: str, output_name: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert a Word document to HTML format.
    Args:
        document_name: Name of the document to convert
        output_name: Optional name for the output HTML file
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_to_html(output_name)
        return {
            "success": True,
            "message": "Document converted to HTML successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_HTML_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and is accessible."
            }
        }

async def convert_to_rtf(document_name: str, output_name: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert a Word document to RTF format.
    Args:
        document_name: Name of the document to convert
        output_name: Optional name for the output RTF file
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_to_rtf(output_name)
        return {
            "success": True,
            "message": "Document converted to RTF successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_RTF_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and is accessible."
            }
        }

async def convert_to_txt(document_name: str, output_name: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert a Word document to plain text format.
    Args:
        document_name: Name of the document to convert
        output_name: Optional name for the output TXT file
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_to_txt(output_name)
        return {
            "success": True,
            "message": "Document converted to TXT successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_TXT_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and is accessible."
            }
        }

async def convert_to_odt(document_name: str, output_name: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Convert a Word document to ODT format.
    Args:
        document_name: Name of the document to convert
        output_name: Optional name for the output ODT file
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with conversion details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = ConversionHandler(doc_path)
        result = handler.convert_to_odt(output_name)
        return {
            "success": True,
            "message": "Document converted to ODT successfully",
            "data": result
        }
    except ConversionError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "CONVERSION_ODT_ERROR",
                "type": "ConversionError",
                "details": str(e),
                "suggestion": "Please check if the document exists and is accessible."
            }
        } 