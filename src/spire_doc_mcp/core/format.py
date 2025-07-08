"""
Format core module for SpireDoc MCP Server.

This module provides core formatting functionality including:
- Paragraph alignment and indentation
- Line spacing and paragraph spacing
- Text formatting and styling
"""

import os
from typing import Dict, Any, Optional, Union
from spire.doc import Document, HorizontalAlignment, LineSpacingRule
from ..utils.exceptions import DocumentError
from ..utils.constants import DEFAULT_VALUES
from .base import BaseHandler

class FormatHandler(BaseHandler):
    """Handler for document formatting operations."""
    
    def _validate_formatting_parameters(self, alignment: Optional[str] = None,
                                      line_spacing_rule: Optional[str] = None,
                                      **kwargs) -> None:
        """Validate formatting parameters.
        
        Args:
            alignment: Alignment type to validate
            line_spacing_rule: Line spacing rule to validate
            **kwargs: Other parameters to validate
            
        Raises:
            DocumentError: If any parameter is invalid
        """
        # Validate alignment
        if alignment is not None:
            valid_alignments = ['left', 'center', 'right', 'justify']
            if alignment.lower() not in valid_alignments:
                raise DocumentError(f"Invalid alignment: {alignment}. Valid values: {', '.join(valid_alignments)}")
        
        # Validate line spacing rule
        if line_spacing_rule is not None:
            valid_rules = ['at_least', 'exactly', 'multiple']
            if line_spacing_rule.lower() not in valid_rules:
                raise DocumentError(f"Invalid line spacing rule: {line_spacing_rule}. Valid values: {', '.join(valid_rules)}")
        
        # Validate numeric parameters
        numeric_params = ['first_line_indent', 'left_indent', 'right_indent', 'line_spacing', 'before_spacing', 'after_spacing']
        for param in numeric_params:
            value = kwargs.get(param)
            if value is not None:
                if not isinstance(value, (int, float)):
                    raise DocumentError(f"{param} must be a number, got {type(value)}")
                if value < 0:
                    raise DocumentError(f"{param} cannot be negative: {value}")
    
    def format_paragraph(self, paragraph_index: int,
                        alignment: Optional[str] = None,
                        first_line_indent: Optional[float] = None,
                        left_indent: Optional[float] = None,
                        right_indent: Optional[float] = None,
                        line_spacing: Optional[float] = None,
                        line_spacing_rule: Optional[str] = None,
                        before_spacing: Optional[float] = None,
                        after_spacing: Optional[float] = None) -> Dict[str, Any]:
        """Apply formatting to a paragraph.
        
        Args:
            paragraph_index: Index of the paragraph to format
            alignment: Alignment type ('left', 'center', 'right', 'justify')
            first_line_indent: First line indent in points
            left_indent: Left indent in points
            right_indent: Right indent in points
            line_spacing: Line spacing value
            line_spacing_rule: Line spacing rule ('at_least', 'exactly', 'multiple')
            before_spacing: Space before paragraph in points
            after_spacing: Space after paragraph in points
            
        Returns:
            Dict containing formatting details
            
        Raises:
            DocumentError: If formatting fails
        """
        try:
            # Validate parameters using base class method
            self._validate_paragraph_index(paragraph_index)
            self._validate_formatting_parameters(
                alignment=alignment,
                line_spacing_rule=line_spacing_rule,
                first_line_indent=first_line_indent,
                left_indent=left_indent,
                right_indent=right_indent,
                line_spacing=line_spacing,
                before_spacing=before_spacing,
                after_spacing=after_spacing
            )
            
            # Get the first section and paragraph
            section = self.document.Sections.get_Item(0)
            paragraph = section.Paragraphs.get_Item(paragraph_index)
            
            # Store original formatting for comparison
            original_formatting = {
                "alignment": str(paragraph.Format.HorizontalAlignment),
                "first_line_indent": paragraph.Format.FirstLineIndent,
                "left_indent": paragraph.Format.LeftIndent,
                "right_indent": paragraph.Format.RightIndent,
                "line_spacing": paragraph.Format.LineSpacing,
                "line_spacing_rule": str(paragraph.Format.LineSpacingRule),
                "before_spacing": paragraph.Format.BeforeSpacing,
                "after_spacing": paragraph.Format.AfterSpacing
            }
            
            # Apply alignment
            if alignment is not None:
                alignment_map = {
                    'left': HorizontalAlignment.Left,
                    'center': HorizontalAlignment.Center,
                    'right': HorizontalAlignment.Right,
                    'justify': HorizontalAlignment.Justify
                }
                paragraph.Format.HorizontalAlignment = alignment_map[alignment.lower()]
            
            # Apply indentation
            if first_line_indent is not None:
                paragraph.Format.FirstLineIndent = first_line_indent
            if left_indent is not None:
                paragraph.Format.LeftIndent = left_indent
            if right_indent is not None:
                paragraph.Format.RightIndent = right_indent
            
            # Apply line spacing
            if line_spacing is not None:
                paragraph.Format.LineSpacing = line_spacing
                if line_spacing_rule is not None:
                    rule_map = {
                        'at_least': LineSpacingRule.AtLeast,
                        'exactly': LineSpacingRule.Exactly,
                        'multiple': LineSpacingRule.Multiple
                    }
                    paragraph.Format.LineSpacingRule = rule_map[line_spacing_rule.lower()]
            
            # Apply paragraph spacing
            if before_spacing is not None:
                paragraph.Format.BeforeSpacing = before_spacing
            if after_spacing is not None:
                paragraph.Format.AfterSpacing = after_spacing
            
            # Save the document
            try:
                self.document.SaveToFile(self.doc_path)
            except Exception as e:
                raise DocumentError(f"Failed to save document: {str(e)}")
            
            # Get current formatting
            current_formatting = {
                "alignment": str(paragraph.Format.HorizontalAlignment),
                "first_line_indent": paragraph.Format.FirstLineIndent,
                "left_indent": paragraph.Format.LeftIndent,
                "right_indent": paragraph.Format.RightIndent,
                "line_spacing": paragraph.Format.LineSpacing,
                "line_spacing_rule": str(paragraph.Format.LineSpacingRule),
                "before_spacing": paragraph.Format.BeforeSpacing,
                "after_spacing": paragraph.Format.AfterSpacing
            }
            
            return {
                "success": True,
                "message": f"Formatting applied to paragraph {paragraph_index}",
                "paragraph_index": paragraph_index,
                "document_path": self.doc_path,
                "original_formatting": original_formatting,
                "current_formatting": current_formatting,
                "changes_applied": {
                    "alignment": alignment is not None,
                    "first_line_indent": first_line_indent is not None,
                    "left_indent": left_indent is not None,
                    "right_indent": right_indent is not None,
                    "line_spacing": line_spacing is not None,
                    "line_spacing_rule": line_spacing_rule is not None,
                    "before_spacing": before_spacing is not None,
                    "after_spacing": after_spacing is not None
                }
            }
            
        except Exception as e:
            if isinstance(e, DocumentError):
                raise e
            else:
                raise DocumentError(f"Failed to format paragraph: {str(e)}")
    
    def get_paragraph_format(self, paragraph_index: int) -> Dict[str, Any]:
        """Get paragraph formatting information.
        
        Args:
            paragraph_index: Index of the paragraph
            
        Returns:
            Dict containing paragraph formatting details
            
        Raises:
            DocumentError: If operation fails
        """
        try:
            # Validate paragraph index
            self._validate_paragraph_index(paragraph_index)
            
            # Get the first section and paragraph
            section = self.document.Sections.get_Item(0)
            paragraph = section.Paragraphs.get_Item(paragraph_index)
            
            return {
                "success": True,
                "message": f"Retrieved formatting for paragraph {paragraph_index}",
                "paragraph_index": paragraph_index,
                "document_path": self.doc_path,
                "formatting": {
                    "alignment": str(paragraph.Format.HorizontalAlignment),
                    "first_line_indent": paragraph.Format.FirstLineIndent,
                    "left_indent": paragraph.Format.LeftIndent,
                    "right_indent": paragraph.Format.RightIndent,
                    "line_spacing": paragraph.Format.LineSpacing,
                    "line_spacing_rule": str(paragraph.Format.LineSpacingRule),
                    "before_spacing": paragraph.Format.BeforeSpacing,
                    "after_spacing": paragraph.Format.AfterSpacing
                }
            }
            
        except Exception as e:
            if isinstance(e, DocumentError):
                raise e
            else:
                raise DocumentError(f"Failed to get paragraph format: {str(e)}")