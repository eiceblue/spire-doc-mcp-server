"""
Table API module for SpireDoc MCP Server.

This module provides table-related APIs for the MCP server, including:
- Table creation and management
- Table content operations
- Table formatting and styling
"""

from typing import Dict, Any, Optional
from .base import get_doc_path
from ..core.table import TableHandler
from ..utils.exceptions import TableError


async def create_table(document_name: str, rows: int, columns: int, section_index: int = 0,
                      paragraph_index: Optional[int] = None, style: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Create a new table in the document.
    Args:
        document_name: Name of the document
        rows: Number of rows in the table
        columns: Number of columns in the table
        section_index: Index of the section to add table to (default: 0)
        paragraph_index: Index of the paragraph to insert table after (optional)
        style: Optional table style to apply
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with table creation details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = TableHandler(doc_path)
        result = handler.create_table(rows, columns, section_index, paragraph_index, style)
        return {
            "success": True,
            "message": "Table created successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_CREATE_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table dimensions and section/paragraph indices are valid."
            }
        }

async def add_table_after_paragraph(document_name: str, rows: int, columns: int, section_index: int = 0,
                                   paragraph_index: int = 0, style: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Add a table after a specific paragraph.
    Args:
        document_name: Name of the document
        rows: Number of rows in the table
        columns: Number of columns in the table
        section_index: Index of the section (default: 0)
        paragraph_index: Index of the paragraph to insert table after
        style: Optional table style to apply
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with table creation details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = TableHandler(doc_path)
        result = handler.add_table_after_paragraph(rows, columns, section_index, paragraph_index, style)
        return {
            "success": True,
            "message": "Table added after paragraph successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_ADD_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table dimensions and paragraph index are valid."
            }
        }

async def add_table_to_section(document_name: str, rows: int, columns: int, section_index: int = 0,
                              style: Optional[str] = None, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Add a table to the end of a specific section.
    Args:
        document_name: Name of the document
        rows: Number of rows in the table
        columns: Number of columns in the table
        section_index: Index of the section (default: 0)
        style: Optional table style to apply
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with table creation details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = TableHandler(doc_path)
        result = handler.add_table_to_section(rows, columns, section_index, style)
        return {
            "success": True,
            "message": "Table added to section successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_ADD_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table dimensions and section index are valid."
            }
        }

async def get_table_info(document_name: str, table_index: int, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Get information about a table.
    Args:
        document_name: Name of the document
        table_index: Index of the table
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with table information
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = TableHandler(doc_path)
        result = handler.get_table_info(table_index, section_index)
        return {
            "success": True,
            "message": "Table information retrieved successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_INFO_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table index and section index are valid."
            }
        }

async def delete_table(document_name: str, table_index: int, section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Delete a table from the document.
    Args:
        document_name: Name of the document
        table_index: Index of the table to delete
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
        handler = TableHandler(doc_path)
        result = handler.delete_table(table_index, section_index)
        return {
            "success": True,
            "message": "Table deleted successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_DELETE_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table index and section index are valid."
            }
        }

async def set_cell_text(document_name: str, table_index: int, row: int, column: int, text: str,
                       section_index: int = 0, doc_files_path: str = "./word_files") -> Dict[str, Any]:
    """Set text content in a table cell.
    Args:
        document_name: Name of the document
        table_index: Index of the table
        row: Row index (0-based)
        column: Column index (0-based)
        text: Text content to set
        section_index: Index of the section (default: 0)
        doc_files_path: Path to the documents directory
    Returns:
        Dict containing:
            - success: bool
            - message: str
            - data: Dict with cell update details
    """
    try:
        doc_path = get_doc_path(document_name, doc_files_path)
        handler = TableHandler(doc_path)
        result = handler.set_cell_text(table_index, row, column, text, section_index)
        return {
            "success": True,
            "message": "Cell text set successfully",
            "data": result
        }
    except TableError as e:
        return {
            "success": False,
            "message": str(e),
            "error": {
                "code": "TABLE_CELL_ERROR",
                "type": "TableError",
                "details": str(e),
                "suggestion": "Please check if the table index, row, column, and section index are valid."
            }
        } 