"""
Utility package for SpireDoc MCP.

This package provides:
- Custom exceptions for error handling
- Helper functions for common operations
- Global constants and configuration
"""

from .exceptions import (
    SpireDocError,
    DocumentError,
    ParagraphError,
    TableError,
    ConversionError
)

from .constants import (
    SUPPORTED_FORMATS,
    PROTECTION_TYPES,
    ERROR_CODES,
    DEFAULT_VALUES,
    VALIDATION_RULES
)

__all__ = [
    # Exceptions
    'SpireDocError',
    'DocumentError',
    'ParagraphError',
    'TableError',
    'ConversionError',
    
    # Constants
    'SUPPORTED_FORMATS',
    'PROTECTION_TYPES',
    'ERROR_CODES',
    'DEFAULT_VALUES',
    'VALIDATION_RULES'
]