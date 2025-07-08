"""
Constants module for SpireDoc MCP Server.

This module provides global constants and configuration values used throughout the application.
"""

from spire.doc import FileFormat, ProtectionType

# Document formats
SUPPORTED_FORMATS = {
    "docx": FileFormat.Docx,
    "doc": FileFormat.Doc,
    "pdf": FileFormat.PDF,
    "rtf": FileFormat.Rtf,
    "txt": FileFormat.Txt,
    "html": FileFormat.Html,
    "epub": FileFormat.EPub,
    "odt": FileFormat.Odt,
    "xml": FileFormat.Xml,
    "md": FileFormat.Markdown
}

# Document protection levels
PROTECTION_TYPES = {
    "none": ProtectionType.NoProtection,
    "read_only": ProtectionType.AllowOnlyReading,
    "form_filling": ProtectionType.AllowOnlyFormFields,
    "comments": ProtectionType.AllowOnlyComments,
    "revisions": ProtectionType.AllowOnlyRevisions
}

# Error codes
ERROR_CODES = {
    # Document errors
    "DOC_CREATE_ERROR": "Failed to create document",
    "DOC_LOAD_ERROR": "Failed to load document",
    "DOC_SAVE_ERROR": "Failed to save document",
    "DOC_DELETE_ERROR": "Failed to delete document",
    "DOC_PROTECT_ERROR": "Failed to protect document",
    "DOC_INFO_ERROR": "Failed to get document info",
    
    # Table errors
    "TABLE_CREATE_ERROR": "Failed to create table",
    "TABLE_UPDATE_ERROR": "Failed to update table",
    "TABLE_DELETE_ERROR": "Failed to delete table",
    "TABLE_INFO_ERROR": "Failed to get table info",
    
    # Paragraph errors
    "PARA_CREATE_ERROR": "Failed to create paragraph",
    "PARA_UPDATE_ERROR": "Failed to update paragraph",
    "PARA_DELETE_ERROR": "Failed to delete paragraph",
    "PARA_INFO_ERROR": "Failed to get paragraph info",
    
    # Conversion errors
    "CONV_INPUT_ERROR": "Invalid input format",
    "CONV_OUTPUT_ERROR": "Invalid output format",
    "CONV_PROCESS_ERROR": "Failed to convert document",
    
    # Validation errors
    "VAL_REQUIRED": "Required field missing",
    "VAL_INVALID": "Invalid value",
    "VAL_FORMAT": "Invalid format",
    "VAL_RANGE": "Value out of range",
    
    # System errors
    "SYS_CONFIG_ERROR": "Configuration error",
    "SYS_PERMISSION_ERROR": "Permission denied",
    "SYS_RESOURCE_ERROR": "Resource not available",
    "SYS_UNKNOWN_ERROR": "Unknown error"
}

# Default values
DEFAULT_VALUES = {
    "document": {
        "format": "docx",
        "protection": "none",
        "template": None
    },
    "table": {
        "rows": 1,
        "columns": 1,
        "style": "Table Grid"
    },
    "paragraph": {
        "style": "Normal",
        "alignment": "left",
        "spacing": 1.0
    },
    "conversion": {
        "format": "docx",
        "quality": "high"
    }
}

# Validation rules
VALIDATION_RULES = {
    "document": {
        "name": {
            "pattern": r"^[a-zA-Z0-9_-]+\.(docx|doc|pdf|rtf|txt|html|epub|odt|xml|md)$",
            "message": "Invalid document name format"
        },
        "protection": {
            "values": list(PROTECTION_TYPES.keys()),
            "message": "Invalid protection level"
        }
    },
    "table": {
        "rows": {
            "min": 1,
            "max": 100,
            "message": "Number of rows must be between 1 and 100"
        },
        "columns": {
            "min": 1,
            "max": 20,
            "message": "Number of columns must be between 1 and 20"
        }
    },
    "paragraph": {
        "style": {
            "values": ["Normal", "Heading 1", "Heading 2", "Heading 3", "Title", "Subtitle"],
            "message": "Invalid paragraph style"
        },
        "alignment": {
            "values": ["left", "center", "right", "justify"],
            "message": "Invalid alignment value"
        }
    }
} 