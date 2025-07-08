"""
Table core module for SpireDoc MCP Server.

This module provides table-level operations including:
- Table creation and management
- Table content operations
- Table formatting and styling
"""

import os
from typing import Dict, Any, Optional
from spire.doc import Document, Table, BorderStyle, HorizontalAlignment
from .base import BaseHandler
from ..utils.exceptions import TableError
from ..utils.constants import DEFAULT_VALUES, VALIDATION_RULES

class TableHandler(BaseHandler):
    """Handler for table operations."""
    
    def get_table_info(self, table_index: int, section_index: int = 0) -> Dict[str, Any]:
        """Get information about a table.
        
        Args:
            table_index: Index of the table
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing table information
            
        Raises:
            TableError: If the table cannot be found
        """
        try:
            self._validate_section_index(section_index)
            section = self.document.Sections[section_index]
            
            if table_index < 0 or table_index >= len(section.Tables):
                raise TableError(f"Invalid table index: {table_index}", "TABLE_INFO_ERROR")
            
            table = section.Tables[table_index]
            
            return {
                "table_index": table_index,
                "section_index": section_index,
                "rows": table.Rows.Count,
                "columns": table.Rows[0].Cells.Count if table.Rows.Count > 0 else 0,
                "total_cells": table.Rows.Count * (table.Rows[0].Cells.Count if table.Rows.Count > 0 else 0)
            }
            
        except Exception as e:
            if isinstance(e, TableError):
                raise e
            else:
                raise TableError(f"Failed to get table info: {str(e)}", "TABLE_INFO_ERROR")
    
    def create_table(self, rows: int, columns: int, section_index: int = 0, 
                    paragraph_index: Optional[int] = None, style: Optional[str] = None) -> Dict[str, Any]:
        """Create a new table in the document.
        
        Args:
            rows: Number of rows in the table
            columns: Number of columns in the table
            section_index: Index of the section to add table to (default: 0)
            paragraph_index: Index of the paragraph to insert table after (optional)
            style: Optional table style to apply
            
        Returns:
            Dict containing table creation details
            
        Raises:
            TableError: If table creation fails
        """
        try:
            # Validate section index using base class method
            self._validate_section_index(section_index)
            
            section = self.document.Sections[section_index]
            
            # Validate paragraph index if provided
            if paragraph_index is not None:
                self._validate_paragraph_index(paragraph_index, section_index)
            
            # Validate table dimensions
            if rows <= 0 or columns <= 0:
                raise TableError("Table dimensions must be positive", "TABLE_CREATE_ERROR")
            
            if rows > 1000 or columns > 100:
                raise TableError("Table dimensions too large (max: 1000 rows, 100 columns)", "TABLE_CREATE_ERROR")
            
            # Create the table using the correct Spire.Doc API
            if paragraph_index is not None:
                # Insert table after a specific paragraph
                table = section.AddTable(True)  # Create object only, don't insert
                table.ResetCells(rows, columns)
                insert_index = min(paragraph_index + 1, len(section.Tables))
                section.Tables.Insert(insert_index, table)
            else:
                # Add table to the end of the section
                table = section.AddTable()
                table.ResetCells(rows, columns)
            
            # Apply style if specified
            if style:
                try:
                    table.ApplyStyle(style)
                except Exception as e:
                    # If style application fails, continue without style
                    pass
            
            # Get the table index
            table_index = len(section.Tables) - 1
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "success": True,
                "message": f"Table created successfully",
                "table_index": table_index,
                "section_index": section_index,
                "paragraph_index": paragraph_index,
                "dimensions": {
                    "rows": rows,
                    "columns": columns
                },
                "style": style,
                "total_tables_in_section": len(section.Tables)
            }
            
        except Exception as e:
            if isinstance(e, TableError):
                raise e
            else:
                raise TableError(f"Failed to create table: {str(e)}", "TABLE_CREATE_ERROR")
    
    def add_table_after_paragraph(self, rows: int, columns: int, section_index: int = 0,
                                 paragraph_index: int = 0, style: Optional[str] = None) -> Dict[str, Any]:
        """Add a table after a specific paragraph.
        
        Args:
            rows: Number of rows in the table
            columns: Number of columns in the table
            section_index: Index of the section (default: 0)
            paragraph_index: Index of the paragraph to insert table after
            style: Optional table style to apply
            
        Returns:
            Dict containing table creation details
            
        Raises:
            TableError: If table creation fails
        """
        return self.create_table(rows, columns, section_index, paragraph_index, style)
    
    def add_table_to_section(self, rows: int, columns: int, section_index: int = 0,
                            style: Optional[str] = None) -> Dict[str, Any]:
        """Add a table to the end of a specific section.
        
        Args:
            rows: Number of rows in the table
            columns: Number of columns in the table
            section_index: Index of the section (default: 0)
            style: Optional table style to apply
            
        Returns:
            Dict containing table creation details
            
        Raises:
            TableError: If table creation fails
        """
        return self.create_table(rows, columns, section_index, None, style)
    
    def delete_table(self, table_index: int, section_index: int = 0) -> Dict[str, Any]:
        """Delete a table from the document.
        
        Args:
            table_index: Index of the table to delete
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing deletion details
            
        Raises:
            TableError: If the table cannot be deleted
        """
        try:
            # Get table info before deletion
            info = self.get_table_info(table_index, section_index)
            
            section = self.document.Sections[section_index]
            
            # Remove the table from the section
            section.Tables.RemoveAt(table_index)
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "section_index": section_index,
                "table_index": table_index,
                "deleted_dimensions": {
                    "rows": info["rows"],
                    "columns": info["columns"]
                },
                "total_tables_remaining": len(section.Tables)
            }
            
        except Exception as e:
            if isinstance(e, TableError):
                raise e
            else:
                raise TableError(f"Failed to delete table: {str(e)}", "TABLE_DELETE_ERROR")
    
    def set_cell_text(self, table_index: int, row: int, column: int, text: str, 
                     section_index: int = 0) -> Dict[str, Any]:
        """Set text content in a table cell.
        
        Args:
            table_index: Index of the table
            row: Row index (0-based)
            column: Column index (0-based)
            text: Text content to set
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing cell update details
            
        Raises:
            TableError: If the cell cannot be updated
        """
        try:
            # Get table info to validate dimensions
            table_info = self.get_table_info(table_index, section_index)
            
            # Validate cell coordinates
            if row < 0 or row >= table_info["rows"]:
                raise TableError(f"Invalid row index: {row}", "TABLE_CELL_ERROR")
            
            if column < 0 or column >= table_info["columns"]:
                raise TableError(f"Invalid column index: {column}", "TABLE_CELL_ERROR")
            
            section = self.document.Sections[section_index]
            table = section.Tables[table_index]
            
            # Get the cell and set text using the correct API
            row_obj = table.Rows[row]
            cell = row_obj.Cells[column]
            
            # Get the first paragraph of the cell and set its text
            if cell.Paragraphs.Count > 0:
                paragraph = cell.Paragraphs[0]
                original_text = paragraph.Text
                paragraph.Text = text
            else:
                # If no paragraph exists, add one
                paragraph = cell.AddParagraph()
                paragraph.Text = text
                original_text = ""
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "section_index": section_index,
                "table_index": table_index,
                "row": row,
                "column": column,
                "original_text": original_text,
                "new_text": text
            }
            
        except Exception as e:
            if isinstance(e, TableError):
                raise e
            else:
                raise TableError(f"Failed to set cell text: {str(e)}", "TABLE_CELL_ERROR")
    
    
    
    
    
    
    
    
