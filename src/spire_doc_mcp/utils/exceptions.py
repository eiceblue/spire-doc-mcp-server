"""
Exceptions module for SpireDoc MCP Server.

This module provides custom exception classes and error handling utilities.
"""

from typing import Dict, Any, Optional
from .constants import ERROR_CODES

class SpireDocError(Exception):
    """Base exception class for SpireDoc MCP errors."""
    
    def __init__(self, message: str, code: str = "SYS_UNKNOWN_ERROR", details: Optional[Dict[str, Any]] = None):
        """Initialize the exception.
        
        Args:
            message: Error message
            code: Error code from ERROR_CODES
            details: Optional additional error details
        """
        self.message = message
        self.code = code
        self.details = details or {}
        super().__init__(self.message)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert exception to dictionary format.
        
        Returns:
            Dict containing error information
        """
        return {
            "success": False,
            "message": self.message,
            "error": {
                "code": self.code,
                "type": self.__class__.__name__,
                "details": self.details,
                "suggestion": ERROR_CODES.get(self.code, "Unknown error")
            }
        }

class DocumentError(SpireDocError):
    """Exception raised for document-related errors."""
    
    def __init__(self, message: str, code: str = "DOC_LOAD_ERROR", details: Optional[Dict[str, Any]] = None):
        """Initialize the exception.
        
        Args:
            message: Error message
            code: Error code from ERROR_CODES
            details: Optional additional error details
        """
        super().__init__(message, code, details)

class TableError(SpireDocError):
    """Exception raised for table-related errors."""
    
    def __init__(self, message: str, code: str = "TABLE_CREATE_ERROR", details: Optional[Dict[str, Any]] = None):
        """Initialize the exception.
        
        Args:
            message: Error message
            code: Error code from ERROR_CODES
            details: Optional additional error details
        """
        super().__init__(message, code, details)

class ParagraphError(SpireDocError):
    """Exception raised for paragraph-related errors."""
    
    def __init__(self, message: str, code: str = "PARA_CREATE_ERROR", details: Optional[Dict[str, Any]] = None):
        """Initialize the exception.
        
        Args:
            message: Error message
            code: Error code from ERROR_CODES
            details: Optional additional error details
        """
        super().__init__(message, code, details)

class ConversionError(SpireDocError):
    """Exception raised for document conversion errors."""
    
    def __init__(self, message: str, code: str = "CONV_PROCESS_ERROR", details: Optional[Dict[str, Any]] = None):
        """Initialize the exception.
        
        Args:
            message: Error message
            code: Error code from ERROR_CODES
            details: Optional additional error details
        """
        super().__init__(message, code, details)

__all__ = [
    'SpireDocError',
    'DocumentError',
    'ParagraphError',
    'TableError',
    'ConversionError'
] 