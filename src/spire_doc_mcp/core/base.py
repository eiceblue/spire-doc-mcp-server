"""
Base core module for SpireDoc MCP Server.

This module provides a base handler class that contains common functionality
used by all core modules, including document loading and initialization.
"""

import os
from typing import Optional
from spire.doc import Document
from ..utils.exceptions import DocumentError

class BaseHandler:
    """Base handler class for document operations."""
    
    def __init__(self, doc_path: str):
        """Initialize the base handler.
        
        Args:
            doc_path: Path to the document file
            
        Raises:
            DocumentError: If document path is invalid or document cannot be loaded
        """
        if not doc_path:
            raise DocumentError("Document path cannot be empty")
        
        # Normalize and validate the document path
        self.doc_path = os.path.abspath(doc_path)
        
        # Check if the path is a valid file
        if not os.path.isfile(self.doc_path):
            raise DocumentError(f"Document file not found: {self.doc_path}")
        
        # Check file extension
        if not self.doc_path.lower().endswith(('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm')):
            raise DocumentError(f"Unsupported document format: {os.path.splitext(self.doc_path)[1]}")
        
        self.document = None
        self._load_document()
    
    def _load_document(self) -> None:
        """Load the document file.
        
        Raises:
            DocumentError: If document cannot be loaded
        """
        try:
            # Check file permissions
            if not os.access(self.doc_path, os.R_OK):
                raise DocumentError(f"No read permission for document: {self.doc_path}")
            
            if not os.access(self.doc_path, os.W_OK):
                raise DocumentError(f"No write permission for document: {self.doc_path}")
            
            # Load the document
            self.document = Document()
            self.document.LoadFromFile(self.doc_path)
            
            # Verify document was loaded successfully
            if not self.document:
                raise DocumentError("Failed to create document object")
            
            # Check if document has at least one section
            if self.document.Sections.Count == 0:
                raise DocumentError("Document has no sections")
                
        except Exception as e:
            # Clean up on error
            if self.document:
                try:
                    self.document.Close()
                except:
                    pass
                self.document = None
            
            if isinstance(e, DocumentError):
                raise e
            else:
                raise DocumentError(f"Failed to load document: {str(e)}")
    
    def _validate_section_index(self, section_index: int) -> None:
        """Validate section index.
        
        Args:
            section_index: Index of the section to validate
            
        Raises:
            DocumentError: If section index is invalid
        """
        if not isinstance(section_index, int):
            raise DocumentError(f"Section index must be an integer, got {type(section_index)}")
        
        if section_index < 0:
            raise DocumentError(f"Section index cannot be negative: {section_index}")
        
        if not self.document:
            raise DocumentError("Document not loaded")
        
        if section_index >= self.document.Sections.Count:
            raise DocumentError(f"Section index {section_index} out of range (document has {self.document.Sections.Count} sections)")
    
    def _validate_paragraph_index(self, paragraph_index: int, section_index: int = 0) -> None:
        """Validate paragraph index.
        
        Args:
            paragraph_index: Index of the paragraph to validate
            section_index: Index of the section (default: 0)
            
        Raises:
            DocumentError: If paragraph index is invalid
        """
        self._validate_section_index(section_index)
        
        if not isinstance(paragraph_index, int):
            raise DocumentError(f"Paragraph index must be an integer, got {type(paragraph_index)}")
        
        if paragraph_index < 0:
            raise DocumentError(f"Paragraph index cannot be negative: {paragraph_index}")
        
        section = self.document.Sections[section_index]
        if paragraph_index >= section.Paragraphs.Count:
            raise DocumentError(f"Paragraph index {paragraph_index} out of range (section {section_index} has {section.Paragraphs.Count} paragraphs)")
    
    def __del__(self):
        """Clean up resources when the handler is destroyed."""
        if hasattr(self, 'document') and self.document:
            try:
                self.document.Close()
            except:
                pass 