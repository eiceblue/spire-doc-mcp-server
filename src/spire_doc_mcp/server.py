"""
SpireDoc MCP Server.

This module provides the main server implementation for the SpireDoc MCP service.
It handles API registration and server startup.
"""

import os
import sys
import logging
from mcp.server.fastmcp import FastMCP

# Import API modules
from spire_doc_mcp.api import (
    document_tools,
    format_tools,
    paragraph_tools,
    table_tools,
    conversion_tools
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("spire-doc-mcp.log")
    ],
    force=True
)
logger = logging.getLogger("spire-doc-mcp")

# Get Word files path from environment or use default
WORD_FILES_PATH = os.environ.get("WORD_FILES_PATH", "./word_files")

# # Create the directory if it doesn't exist
# os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

# Initialize FastMCP server
mcp = FastMCP(
    "spire-doc-mcp",
    version="1.0.0",
    description="Spire.Doc MCP Server for manipulating Word files",
    dependencies=["Spire.Doc.Free>=12.12.0"],
    env_vars={
        "WORD_FILES_PATH": {
            "description": "Path to Word files directory",
            "required": False,
            "default": WORD_FILES_PATH
        }
    }
)

def register_tools():
    """Register all API tools with the MCP server."""
    
    # Document tools
    @mcp.tool()
    async def create_document(document_name: str, template_path: str = None):
        return await document_tools.create_document(document_name, template_path, WORD_FILES_PATH)
    
    @mcp.tool()
    async def set_document_protection(document_name: str, protection_level: str, password: str = None):
        return await document_tools.set_document_protection(document_name, protection_level, password, WORD_FILES_PATH)
    
    @mcp.tool()
    async def find_and_replace(document_name: str, find_text: str, replace_text: str, match_case: bool = False, match_whole_word: bool = False):
        return await document_tools.find_and_replace(document_name, find_text, replace_text, match_case, match_whole_word, WORD_FILES_PATH)
    
    @mcp.tool()
    async def merge_documents(original_document: str, merge_document: str, output_document: str):
        return await document_tools.merge_documents(original_document, merge_document, output_document, WORD_FILES_PATH)
    
    @mcp.tool()
    async def add_text_watermark(document_name: str, text: str, font_size: int = 65, color: str = "Red"):
        return await document_tools.add_text_watermark(document_name, text, font_size, color, doc_files_path=WORD_FILES_PATH)
    
    # Paragraph tools
    @mcp.tool()
    async def get_paragraph_text(document_name: str, paragraph_index: int, section_index: int = 0):
        return await paragraph_tools.get_paragraph_text(document_name, paragraph_index, section_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def delete_paragraph(document_name: str, paragraph_index: int, section_index: int = 0):
        return await paragraph_tools.delete_paragraph(document_name, paragraph_index, section_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def get_paragraph_info(document_name: str, paragraph_index: int, section_index: int = 0):
        return await paragraph_tools.get_paragraph_info(document_name, paragraph_index, section_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def add_paragraph(document_name: str, text: str, section_index: int = 0, paragraph_index: int = None):
        return await paragraph_tools.add_paragraph(document_name, text, section_index, paragraph_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def update_paragraph_text(document_name: str, paragraph_index: int, new_text: str, section_index: int = 0):
        return await paragraph_tools.update_paragraph_text(document_name, paragraph_index, new_text, section_index, WORD_FILES_PATH)
    
    # Table tools
    @mcp.tool()
    async def create_table(document_name: str, rows: int, columns: int, section_index: int = 0, paragraph_index: int = None, style: str = None):
        return await table_tools.create_table(document_name, rows, columns, section_index, paragraph_index, style, WORD_FILES_PATH)
    
    @mcp.tool()
    async def add_table_after_paragraph(document_name: str, rows: int, columns: int, section_index: int = 0, paragraph_index: int = 0, style: str = None):
        return await table_tools.add_table_after_paragraph(document_name, rows, columns, section_index, paragraph_index, style, WORD_FILES_PATH)
    
    @mcp.tool()
    async def add_table_to_section(document_name: str, rows: int, columns: int, section_index: int = 0, style: str = None):
        return await table_tools.add_table_to_section(document_name, rows, columns, section_index, style, WORD_FILES_PATH)
    
    @mcp.tool()
    async def get_table_info(document_name: str, table_index: int, section_index: int = 0):
        return await table_tools.get_table_info(document_name, table_index, section_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def delete_table(document_name: str, table_index: int, section_index: int = 0):
        return await table_tools.delete_table(document_name, table_index, section_index, WORD_FILES_PATH)
    
    @mcp.tool()
    async def set_cell_text(document_name: str, table_index: int, row: int, column: int, text: str, section_index: int = 0):
        return await table_tools.set_cell_text(document_name, table_index, row, column, text, section_index, WORD_FILES_PATH)
    
    # Format tools
    @mcp.tool()
    async def format_paragraph(document_name: str, paragraph_index: int, alignment: str = None, first_line_indent: float = None, left_indent: float = None, right_indent: float = None, line_spacing: float = None, line_spacing_rule: str = None, before_spacing: float = None, after_spacing: float = None):
        return await format_tools.format_paragraph(document_name, paragraph_index, alignment, first_line_indent, left_indent, right_indent, line_spacing, line_spacing_rule, before_spacing, after_spacing, WORD_FILES_PATH)
    
    @mcp.tool()
    async def get_paragraph_format(document_name: str, paragraph_index: int):
        return await format_tools.get_paragraph_format(document_name, paragraph_index, WORD_FILES_PATH)
    
    # Conversion tools
    @mcp.tool()
    async def convert_document(document_name: str, target_format: str, output_path: str = None):
        return await conversion_tools.convert_document(document_name, target_format, output_path, WORD_FILES_PATH)
    
    @mcp.tool()
    async def get_conversion_status(document_name: str):
        return await conversion_tools.get_conversion_status(document_name, WORD_FILES_PATH)
    
    @mcp.tool()
    async def get_conversion_history(document_name: str):
        return await conversion_tools.get_conversion_history(document_name, WORD_FILES_PATH)

async def run_server():
    """Run the Spire.Doc MCP Server."""
    # Register all tools
    logger.info("Registering API modules...")
    register_tools()
    logger.info("API modules registered successfully")
    
    try:
        logger.info(f"Starting Spire.Doc MCP Server (files directory: {WORD_FILES_PATH})")
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        await mcp.shutdown()
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")