"""
Conversion core functionality module.

This module provides core document conversion functionality, including:
- Document format conversion
- Conversion status tracking
- Conversion history management
"""

import os
from typing import Dict, Any, Optional
from datetime import datetime
from spire.doc import Document, FileFormat
from ..utils.exceptions import DocumentError, ConversionError
from ..utils.constants import SUPPORTED_FORMATS

class ConversionHandler:
    """Handler for document conversion operations."""
    
    def __init__(self, doc_path: Optional[str] = None):
        """Initialize the conversion handler.
        
        Args:
            doc_path: Optional path to the document file
        """
        self.doc_path = doc_path
        self.conversion_history = []
        self.conversion_status = {}
    
    def _load_document(self) -> Document:
        """Load the document file.
        
        Returns:
            Document: The loaded document
            
        Raises:
            DocumentError: If the document cannot be loaded
        """
        if not self.doc_path:
            raise DocumentError("No document path provided")
            
        try:
            doc = Document()
            doc.LoadFromFile(self.doc_path)
            return doc
        except Exception as e:
            raise DocumentError(f"Failed to load document: {str(e)}")
    
    def convert_document(self, target_format: str, output_path: Optional[str] = None) -> Dict[str, Any]:
        """Convert the document to a different format.
        
        Args:
            target_format: Target format for conversion
            output_path: Optional output path for the converted document
            
        Returns:
            Dict containing conversion details
            
        Raises:
            ConversionError: If the conversion fails
        """
        try:
            if target_format.lower() not in SUPPORTED_FORMATS:
                raise ConversionError(f"Unsupported format: {target_format}")
            
            doc = self._load_document()
            
            # Generate output path if not provided
            if output_path is None:
                output_path = self.doc_path.rsplit('.', 1)[0] + '.' + target_format.lower()
            else:
                # Ensure output path has correct extension
                if not output_path.lower().endswith('.' + target_format.lower()):
                    output_path = output_path.rsplit('.', 1)[0] + '.' + target_format.lower()
            
            # Check if output directory exists, create if not
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # Convert document
            doc.SaveToFile(output_path, SUPPORTED_FORMATS[target_format.lower()])
            
            # Update conversion history
            conversion_info = {
                "timestamp": datetime.now().isoformat(),
                "source_format": self.doc_path.split('.')[-1].lower(),
                "target_format": target_format.lower(),
                "output_path": output_path,
                "status": "success"
            }
            self.conversion_history.append(conversion_info)
            self.conversion_status = conversion_info
            
            return {
                "document": os.path.basename(self.doc_path),
                "source_format": conversion_info["source_format"],
                "target_format": conversion_info["target_format"],
                "output_path": output_path,
                "timestamp": conversion_info["timestamp"]
            }
            
        except Exception as e:
            error_info = {
                "timestamp": datetime.now().isoformat(),
                "source_format": self.doc_path.split('.')[-1].lower() if self.doc_path else None,
                "target_format": target_format.lower(),
                "error": str(e),
                "status": "failed"
            }
            self.conversion_history.append(error_info)
            self.conversion_status = error_info
            raise ConversionError(f"Failed to convert document: {str(e)}")
    
    def get_conversion_status(self) -> Dict[str, Any]:
        """Get the current conversion status.
        
        Returns:
            Dict containing conversion status information
        """
        return self.conversion_status
    
    def get_conversion_history(self) -> Dict[str, Any]:
        """Get the conversion history.
        
        Returns:
            Dict containing conversion history
        """
        return {
            "history": self.conversion_history,
            "total_conversions": len(self.conversion_history)
        }
    
