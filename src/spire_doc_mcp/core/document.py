"""
Document core module for SpireDoc MCP Server.

This module provides document-level operations including:
- Document creation and loading
- Document protection and security
- Document content operations
- Document search and replace
- Document merge
- Document watermark
"""

import os
from typing import Dict, Any, Optional
from spire.doc import Document, FileFormat, TextWatermark, WatermarkLayout, Color
from .base import BaseHandler
from ..utils.exceptions import DocumentError
from ..utils.constants import SUPPORTED_FORMATS, PROTECTION_TYPES

class DocumentHandler(BaseHandler):
    """Handler for document operations."""
    
    def __init__(self, doc_path: str, create_if_not_exists: bool = False):
        """Initialize the document handler.
        Args:
            doc_path: Path to the document file
            create_if_not_exists: If True, allow initialization even if file does not exist (for creation)
        """
        if create_if_not_exists and not os.path.isfile(doc_path):
            # Only validate path and extension, don't check if file exists
            if not doc_path:
                raise DocumentError("Document path cannot be empty")
            self.doc_path = os.path.abspath(doc_path)
            if not self.doc_path.lower().endswith((
                '.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm')):
                raise DocumentError(f"Unsupported document format: {os.path.splitext(self.doc_path)[1]}")
            self.document = None
        else:
            super().__init__(doc_path)
    
    def create_document(self, template_path: Optional[str] = None) -> Dict[str, Any]:
        """Create a new document.
        
        Args:
            template_path: Optional path to a template document
            
        Returns:
            Dict containing document details
            
        Raises:
            DocumentError: If document creation fails
        """
        try:
            if template_path and os.path.exists(template_path):
                self.document = Document()
                self.document.LoadFromFile(template_path)
            else:
                self.document = Document()
            
            self.document.SaveToFile(self.doc_path, FileFormat.Docx)
            
            return {
                "path": self.doc_path,
                "name": os.path.basename(self.doc_path),
                "size": os.path.getsize(self.doc_path),
                "created": os.path.getctime(self.doc_path),
                "modified": os.path.getmtime(self.doc_path)
            }
            
        except Exception as e:
            raise DocumentError(f"Failed to create document: {str(e)}")
    
    def protect_document(self, protection_level: str, password: Optional[str] = None) -> Dict[str, Any]:
        """Set document protection level.
        
        Args:
            protection_level: Level of protection to apply
            password: Optional password for protection
            
        Returns:
            Dict containing protection details
            
        Raises:
            DocumentError: If protection cannot be set
        """
        try:
            if not self.document:
                raise DocumentError("Document not loaded")
            
            if protection_level not in PROTECTION_TYPES:
                raise DocumentError(f"Invalid protection level: {protection_level}")
            
            protection_type = PROTECTION_TYPES[protection_level]
            self.document.Protect(protection_type, password)
            self.document.SaveToFile(self.doc_path, FileFormat.Docx)
            
            return {
                "protection_level": protection_level,
                "is_protected": True,
                "has_password": bool(password)
            }
            
        except Exception as e:
            raise DocumentError(f"Failed to protect document: {str(e)}")
    
    def find_and_replace(self, find_text: str, replace_text: str, 
                        match_case: bool = False, match_whole_word: bool = False) -> Dict[str, Any]:
        """Find and replace text in the document.
        
        Args:
            find_text: Text to find
            replace_text: Text to replace with
            match_case: Whether to match case
            match_whole_word: Whether to match whole words only
            
        Returns:
            Dict containing:
                - success: bool
                - message: str
                - replacements: Number of replacements made
                - search_params: Search parameters used
                
        Raises:
            DocumentError: If the replacement fails
        """
        try:
            if not self.document:
                raise DocumentError("Document not loaded")
            
            if not find_text:
                raise DocumentError("Find text cannot be empty")
            
            # Perform the replacement using Spire.Doc's Replace method
            self.document.Replace(find_text, replace_text, match_case, match_whole_word)
            
            # Save the document
            self.document.SaveToFile(self.doc_path, FileFormat.Docx2016)
            
            return {
                "success": True,
                "message": "Text replacement completed successfully",
                "replacements": 1,  # Spire.Doc's Replace method replaces all occurrences
                "search_params": {
                    "find_text": find_text,
                    "replace_text": replace_text,
                    "match_case": match_case,
                    "match_whole_word": match_whole_word
                }
            }
            
        except Exception as e:
            raise DocumentError(f"Failed to replace text: {str(e)}")
    
    
    def merge_documents(self, merge_path: str, output_path: str) -> Dict[str, Any]:
        """Merge another document into the current document using Spire.Doc's insertTextFromFile method.
        
        Args:
            merge_path: Path to the document to merge
            output_path: Path to save the merged document
            
        Returns:
            Dict containing:
                - success: bool
                - message: str
                - output_path: Path to the merged document
                - merge_stats: Statistics about the merge operation
                
        Raises:
            DocumentError: If the merge fails
        """
        try:
            if not os.path.exists(merge_path):
                raise DocumentError(f"Document to merge not found: {merge_path}")
            
            # Check if current document is loaded
            if not self.document:
                raise DocumentError("Original document not loaded")
            
            # Create a copy of the original document for merging
            merged_doc = Document()
            merged_doc.LoadFromFile(self.doc_path)
            
            # Use Spire.Doc's built-in method to insert content from another document
            # This method will insert the content from merge_path into merged_doc
            # The content will start from a new page
            merged_doc.InsertTextFromFile(merge_path, FileFormat.Auto)
            
            # Save the merged document
            merged_doc.SaveToFile(output_path, FileFormat.Docx2016)
            
            # Dispose resources
            merged_doc.Dispose()
            
            return {
                "success": True,
                "message": "Documents merged successfully using insertTextFromFile method",
                "output_path": output_path,
                "merge_stats": {
                    "method": "insertTextFromFile",
                    "merge_document": merge_path,
                    "output_document": output_path
                }
            }
            
        except Exception as e:
            raise DocumentError(f"Failed to merge documents: {str(e)}")
    
    def add_text_watermark(self, text: str, font_size: int = 65, color: str = "Red", 
                          layout: WatermarkLayout = WatermarkLayout.Diagonal) -> Dict[str, Any]:
        """Add text watermark to the document.
        
        Args:
            text: Text to display as watermark
            font_size: Font size of the watermark text
            color: Color of the watermark text (e.g., "Red", "Blue", "Green")
            layout: Layout of the watermark (Diagonal, Horizontal, Vertical)
            
        Returns:
            Dict containing:
                - success: bool
                - message: str
                - watermark_info: Information about the added watermark
                
        Raises:
            DocumentError: If adding watermark fails
        """
        try:
            if not self.document:
                raise DocumentError("Document not loaded")
            
            # Create a TextWatermark object
            txt_watermark = TextWatermark()
            
            # Set the format of the text watermark
            txt_watermark.Text = text
            txt_watermark.FontSize = font_size
            txt_watermark.Color = getattr(Color, f"get_{color}")()
            txt_watermark.Layout = layout
            
            # Add the text watermark to document
            self.document.Watermark = txt_watermark
            
            # Save the document
            self.document.SaveToFile(self.doc_path, FileFormat.Docx2016)
            
            return {
                "success": True,
                "message": "Text watermark added successfully",
                "watermark_info": {
                    "text": text,
                    "font_size": font_size,
                    "color": color,
                    "layout": str(layout)
                }
            }
            
        except Exception as e:
            raise DocumentError(f"Failed to add text watermark: {str(e)}")
    
    
    
    
    