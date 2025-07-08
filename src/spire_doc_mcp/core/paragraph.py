"""
Paragraph core module for SpireDoc MCP Server.

This module provides paragraph-level operations including:
- Paragraph creation and management
- Paragraph content operations
- Paragraph formatting and styling
"""

import os
from typing import Dict, Any, Optional
from spire.doc import Document, Paragraph, ParagraphStyle, HorizontalAlignment, VerticalAlignment
from .base import BaseHandler
from ..utils.exceptions import ParagraphError
from ..utils.constants import DEFAULT_VALUES, VALIDATION_RULES

class ParagraphHandler(BaseHandler):
    """Handler for paragraph operations in a document."""
    
    def get_paragraph_text(self, paragraph_index: int, section_index: int = 0) -> Dict[str, Any]:
        """Get the text content of a paragraph.
        
        Args:
            paragraph_index: Index of the paragraph
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing paragraph text information
            
        Raises:
            ParagraphError: If the paragraph cannot be accessed
        """
        try:
            # Validate indices using base class methods
            self._validate_paragraph_index(paragraph_index, section_index)
            
            section = self.document.Sections[section_index]
            paragraph = section.Paragraphs[paragraph_index]
            text = paragraph.Text
            
            return {
                "section_index": section_index,
                "paragraph_index": paragraph_index,
                "text": text,
                "text_length": len(text),
                "has_text": bool(text.strip()),
                "word_count": len(text.split()) if text.strip() else 0,
                "line_count": len(text.split('\n')) if text else 0
            }
            
        except Exception as e:
            raise ParagraphError(f"Failed to get paragraph text: {str(e)}")
    
    def delete_paragraph(self, paragraph_index: int, section_index: int = 0) -> Dict[str, Any]:
        """Delete a paragraph from the document.
        
        Args:
            paragraph_index: Index of the paragraph to delete
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing deletion details
            
        Raises:
            ParagraphError: If the paragraph cannot be deleted
        """
        try:
            # Validate indices using base class methods
            self._validate_paragraph_index(paragraph_index, section_index)
            
            section = self.document.Sections[section_index]
            
            # Get paragraph info before deletion
            info = self.get_paragraphinfo(paragraph_index, section_index)
            
            # Delete the paragraph
            section.Paragraphs.RemoveAt(paragraph_index)
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "section_index": section_index,
                "paragraph_index": paragraph_index,
                "deleted_text": info["text"],
            }
            
        except Exception as e:
            raise ParagraphError(f"Failed to delete paragraph: {str(e)}")
    
    def get_paragraphinfo(self, paragraph_index: int, section_index: int = 0) -> Dict[str, Any]:
        """Get detailed information about a paragraph.
        
        Args:
            paragraph_index: Index of the paragraph
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing paragraph information
            
        Raises:
            ParagraphError: If the paragraph cannot be accessed
        """
        try:
            # Validate indices using base class methods
            self._validate_paragraph_index(paragraph_index, section_index)
            
            section = self.document.Sections[section_index]
            paragraph = section.Paragraphs[paragraph_index]
            
            return {
                "section_index": section_index,
                "paragraph_index": paragraph_index,
                "text": paragraph.Text,
                "text_length": len(paragraph.Text),
                "has_text": bool(paragraph.Text.strip()),
                "alignment": str(paragraph.Format.HorizontalAlignment),
                "style_name": paragraph.StyleName if hasattr(paragraph, 'StyleName') else None
            }
            
        except Exception as e:
            raise ParagraphError(f"Failed to get paragraph info: {str(e)}")
    
    def add_paragraph(self, text: str, section_index: int = 0, paragraph_index: Optional[int] = None) -> Dict[str, Any]:
        """Add a new paragraph to the document.
        
        Args:
            text: Text content for the new paragraph
            section_index: Index of the section (default: 0)
            paragraph_index: Index where to insert the paragraph (optional, adds to end if not specified)
            
        Returns:
            Dict containing paragraph creation details
            
        Raises:
            ParagraphError: If the paragraph cannot be added
        """
        try:
            # Validate section index using base class method
            self._validate_section_index(section_index)
            
            section = self.document.Sections[section_index]
            
            # Validate paragraph index if provided
            if paragraph_index is not None:
                if paragraph_index < 0 or paragraph_index > len(section.Paragraphs):
                    raise ParagraphError(f"Invalid paragraph index: {paragraph_index}")
            
            # Create new paragraph
            if paragraph_index is not None:
                # Create paragraph at the end first
                paragraph = section.AddParagraph()
                # Insert at specified position
                section.Paragraphs.Insert(paragraph_index, paragraph)
                actual_index = paragraph_index
            else:
                paragraph = section.AddParagraph()
                actual_index = len(section.Paragraphs) - 1
            
            # Set text content
            paragraph.Text = text
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "section_index": section_index,
                "paragraph_index": actual_index,
                "text": text,
                "text_length": len(text),
                "word_count": len(text.split()) if text.strip() else 0,
                "total_paragraphs_in_section": len(section.Paragraphs)
            }
            
        except Exception as e:
            raise ParagraphError(f"Failed to add paragraph: {str(e)}")
    
    def update_paragraph_text(self, paragraph_index: int, new_text: str, section_index: int = 0) -> Dict[str, Any]:
        """Update the text content of a paragraph.
        
        Args:
            paragraph_index: Index of the paragraph to update
            new_text: New text content
            section_index: Index of the section (default: 0)
            
        Returns:
            Dict containing update details
            
        Raises:
            ParagraphError: If the paragraph cannot be updated
        """
        try:
            # Validate indices using base class methods
            self._validate_paragraph_index(paragraph_index, section_index)
            
            section = self.document.Sections[section_index]
            paragraph = section.Paragraphs[paragraph_index]
            
            # Store original text
            original_text = paragraph.Text
            
            # Update text
            paragraph.Text = new_text
            
            # Save the document
            self.document.SaveToFile(self.doc_path)
            
            return {
                "section_index": section_index,
                "paragraph_index": paragraph_index,
                "original_text": original_text,
                "new_text": new_text,
                "text_length": len(new_text),
                "word_count": len(new_text.split()) if new_text.strip() else 0
            }
            
        except Exception as e:
            raise ParagraphError(f"Failed to update paragraph text: {str(e)}")
   
