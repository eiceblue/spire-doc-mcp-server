# Spire.Doc MCP Server Tools

This document provides comprehensive documentation for all available tools in the Spire.Doc MCP Server.

## Overview

The Spire.Doc MCP Server is a Model Context Protocol (MCP) server that provides comprehensive Word document manipulation capabilities through a set of asynchronous tools. Built on top of the Spire.Doc library, it offers document creation, editing, formatting, and conversion functionalities.

## Server Information

- **Server Name**: `spire-doc-mcp`
- **Version**: `0.1.1`
- **Description**: Spire.Doc MCP Server for manipulating Word files
- **Dependencies**: Spire.Doc.Free>=12.12.0

## Response Format

All tools return a consistent response format:

```python
{
    "success": bool,
    "message": str,
    "data": dict,  # Tool-specific data
    "error": {     # Only present if success is False
        "code": str,
        "type": str,
        "details": str,
        "suggestion": str
    }
}
```

## Available Tools

The server provides **18 tools** organized into 5 categories:

### Document Operations (6 tools)
- `create_document` - Create new Word documents
- `set_document_protection` - Set document protection levels
- `find_and_replace` - Find and replace text in documents
- `compare_documents` - Compare two document versions
- `merge_documents` - Merge documents together
- `add_text_watermark` - Add text watermarks to documents

### Paragraph Operations (5 tools)
- `get_paragraph_text` - Retrieve paragraph text content
- `delete_paragraph` - Delete paragraphs from documents
- `get_paragraph_info` - Get detailed paragraph information
- `add_paragraph` - Add new paragraphs to documents
- `update_paragraph_text` - Update paragraph text content

### Table Operations (6 tools)
- `create_table` - Create new tables in documents
- `add_table_after_paragraph` - Add tables after specific paragraphs
- `add_table_to_section` - Add tables to section ends
- `get_table_info` - Get table information
- `delete_table` - Delete tables from documents
- `set_cell_text` - Set text content in table cells

### Formatting Operations (2 tools)
- `format_paragraph` - Apply formatting to paragraphs
- `get_paragraph_format` - Get paragraph formatting information

### Conversion Operations (3 tools)
- `convert_document` - Convert documents to different formats
- `get_conversion_status` - Get conversion status
- `get_conversion_history` - Get conversion history

## Detailed Tool Documentation

### Document Operations

#### create_document

Creates a new Word document.

```python
create_document(document_name: str, template_path: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document to create (must include extension)
- `template_path` (str, optional): Path to a template document

**Returns:** Dictionary with document creation details

**Example Response:**
```python
{
    "success": True,
    "message": "Document created successfully",
    "data": {
        "path": "./word_files/example.docx",
        "name": "example.docx",
        "size": 12345,
        "created": 1640995200.0,
        "modified": 1640995200.0
    }
}
```

#### set_document_protection

Sets document protection level.

```python
set_document_protection(document_name: str, protection_level: str, password: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `protection_level` (str): Protection level ('none', 'read_only', 'form_filling', 'comments', 'revisions')
- `password` (str, optional): Password for protection

**Returns:** Dictionary with protection details

#### find_and_replace

Finds and replaces text in the document.

```python
find_and_replace(document_name: str, find_text: str, replace_text: str, match_case: bool = False, match_whole_word: bool = False) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `find_text` (str): Text to find
- `replace_text` (str): Text to replace with
- `match_case` (bool): Whether to match case (default: False)
- `match_whole_word` (bool): Whether to match whole words only (default: False)

**Returns:** Dictionary with replacement details

#### compare_documents

# ... compare_documents 相关内容全部删除 ...

#### merge_documents

Merges another document into the current document.

```python
merge_documents(original_document: str, merge_document: str, output_document: str) -> dict
```

**Parameters:**
- `original_document` (str): Name of the original document
- `merge_document` (str): Name of the document to merge
- `output_document` (str): Name of the output merged document

**Returns:** Dictionary with merge details

#### add_text_watermark

Adds text watermark to the document.

```python
add_text_watermark(document_name: str, text: str, font_size: int = 65, color: str = "Red") -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `text` (str): Text to display as watermark
- `font_size` (int): Font size of the watermark text (default: 65)
- `color` (str): Color of the watermark text (default: "Red")

**Returns:** Dictionary with watermark details

### Paragraph Operations

#### get_paragraph_text

Gets the text content of a paragraph.

```python
get_paragraph_text(document_name: str, paragraph_index: int, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph (0-based)
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with paragraph text information including text, length, word count, and line count

#### delete_paragraph

Deletes a paragraph from the document.

```python
delete_paragraph(document_name: str, paragraph_index: int, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph to delete
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with deletion details

#### get_paragraph_info

Gets detailed information about a paragraph.

```python
get_paragraph_info(document_name: str, paragraph_index: int, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with paragraph information including text, alignment, and style

#### add_paragraph

Adds a new paragraph to the document.

```python
add_paragraph(document_name: str, text: str, section_index: int = 0, paragraph_index: int = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `text` (str): Text content for the new paragraph
- `section_index` (int): Index of the section (default: 0)
- `paragraph_index` (int, optional): Index where to insert the paragraph (adds to end if not specified)

**Returns:** Dictionary with paragraph creation details

#### update_paragraph_text

Updates the text content of a paragraph.

```python
update_paragraph_text(document_name: str, paragraph_index: int, new_text: str, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph to update
- `new_text` (str): New text content
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with update details

### Table Operations

#### create_table

Creates a new table in the document.

```python
create_table(document_name: str, rows: int, columns: int, section_index: int = 0, paragraph_index: int = None, style: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `rows` (int): Number of rows in the table (1-100)
- `columns` (int): Number of columns in the table (1-20)
- `section_index` (int): Index of the section to add table to (default: 0)
- `paragraph_index` (int, optional): Index of the paragraph to insert table after
- `style` (str, optional): Table style to apply

**Returns:** Dictionary with table creation details

#### add_table_after_paragraph

Adds a table after a specific paragraph.

```python
add_table_after_paragraph(document_name: str, rows: int, columns: int, section_index: int = 0, paragraph_index: int = 0, style: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `rows` (int): Number of rows in the table
- `columns` (int): Number of columns in the table
- `section_index` (int): Index of the section (default: 0)
- `paragraph_index` (int): Index of the paragraph to insert table after (default: 0)
- `style` (str, optional): Table style to apply

**Returns:** Dictionary with table creation details

#### add_table_to_section

Adds a table to the end of a specific section.

```python
add_table_to_section(document_name: str, rows: int, columns: int, section_index: int = 0, style: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `rows` (int): Number of rows in the table
- `columns` (int): Number of columns in the table
- `section_index` (int): Index of the section (default: 0)
- `style` (str, optional): Table style to apply

**Returns:** Dictionary with table creation details

#### get_table_info

Gets information about a table.

```python
get_table_info(document_name: str, table_index: int, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `table_index` (int): Index of the table
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with table information including dimensions

#### delete_table

Deletes a table from the document.

```python
delete_table(document_name: str, table_index: int, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `table_index` (int): Index of the table to delete
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with deletion details

#### set_cell_text

Sets text content in a table cell.

```python
set_cell_text(document_name: str, table_index: int, row: int, column: int, text: str, section_index: int = 0) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `table_index` (int): Index of the table
- `row` (int): Row index (0-based)
- `column` (int): Column index (0-based)
- `text` (str): Text content to set
- `section_index` (int): Index of the section (default: 0)

**Returns:** Dictionary with cell update details

### Formatting Operations

#### format_paragraph

Applies formatting to a paragraph.

```python
format_paragraph(document_name: str, paragraph_index: int, alignment: str = None, first_line_indent: float = None, left_indent: float = None, right_indent: float = None, line_spacing: float = None, line_spacing_rule: str = None, before_spacing: float = None, after_spacing: float = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph to format
- `alignment` (str, optional): Alignment type ('left', 'center', 'right', 'justify')
- `first_line_indent` (float, optional): First line indent in points
- `left_indent` (float, optional): Left indent in points
- `right_indent` (float, optional): Right indent in points
- `line_spacing` (float, optional): Line spacing value
- `line_spacing_rule` (str, optional): Line spacing rule ('at_least', 'exactly', 'multiple')
- `before_spacing` (float, optional): Space before paragraph in points
- `after_spacing` (float, optional): Space after paragraph in points

**Returns:** Dictionary with formatting details

#### get_paragraph_format

Gets paragraph formatting information.

```python
get_paragraph_format(document_name: str, paragraph_index: int) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document
- `paragraph_index` (int): Index of the paragraph

**Returns:** Dictionary with paragraph formatting details

### Conversion Operations

#### convert_document

Converts the document to a different format.

```python
convert_document(document_name: str, target_format: str, output_path: str = None) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document to convert
- `target_format` (str): Target format for conversion
- `output_path` (str, optional): Output path for the converted document

**Returns:** Dictionary with conversion details

#### get_conversion_status

Gets the current conversion status.

```python
get_conversion_status(document_name: str) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document

**Returns:** Dictionary with conversion status information

#### get_conversion_history

Gets the conversion history.

```python
get_conversion_history(document_name: str) -> dict
```

**Parameters:**
- `document_name` (str): Name of the document

**Returns:** Dictionary with conversion history

## Error Handling

All tools include comprehensive error handling with:
- Descriptive error messages
- Error codes for programmatic handling
- Suggestions for resolving issues
- Consistent error format across all tools

### Common Error Codes

- `DOC_CREATE_ERROR`: Failed to create document
- `DOC_LOAD_ERROR`: Failed to load document
- `DOC_SAVE_ERROR`: Failed to save document
- `DOC_PROTECT_ERROR`: Failed to protect document
- `DOC_REPLACE_ERROR`: Failed to replace text
- `DOC_COMPARE_ERROR`: Failed to compare documents
- `TABLE_CREATE_ERROR`: Failed to create table
- `TABLE_ADD_ERROR`: Failed to add table
- `TABLE_DELETE_ERROR`: Failed to delete table
- `TABLE_INFO_ERROR`: Failed to get table info
- `TABLE_CELL_ERROR`: Failed to update cell
- `PARAGRAPH_TEXT_ERROR`: Failed to get paragraph text
- `PARAGRAPH_DELETE_ERROR`: Failed to delete paragraph
- `PARAGRAPH_INFO_ERROR`: Failed to get paragraph info
- `PARAGRAPH_ADD_ERROR`: Failed to add paragraph
- `PARAGRAPH_UPDATE_ERROR`: Failed to update paragraph
- `CONVERSION_ERROR`: Failed to convert document
- `CONVERSION_STATUS_ERROR`: Failed to get conversion status
- `CONVERSION_HISTORY_ERROR`: Failed to get conversion history
- `FORMAT_ERROR`: Failed to apply formatting
- `FORMAT_INFO_ERROR`: Failed to get formatting info

## Supported Formats

### Input Formats
- **DOC**: Microsoft Word 97-2003 Document
- **DOCX**: Microsoft Word 2007+ Document
- **DOCM**: Microsoft Word 2007+ Macro-Enabled Document
- **DOT**: Microsoft Word 97-2003 Template
- **DOTX**: Microsoft Word 2007+ Template
- **DOTM**: Microsoft Word 2007+ Macro-Enabled Template

### Output Formats
- **DOC**: Microsoft Word 97-2003 Document
- **DOCX**: Microsoft Word 2007+ Document
- **PDF**: Portable Document Format
- **RTF**: Rich Text Format
- **HTML**: HyperText Markup Language
- **TXT**: Plain Text
- **EPUB**: Electronic Publication
- **ODT**: OpenDocument Text
- **XML**: Extensible Markup Language
- **Markdown**: Markdown Format

## Configuration

The server uses the following environment variables:

- `WORD_FILES_PATH`: Path to Word files directory (default: "./word_files")

## Usage Guidelines

### 1. Document Naming
- All document names must include the file extension (e.g., "document.docx")
- Use descriptive names that reflect the document content
- Avoid special characters in filenames

### 2. Indexing System
- All indices are 0-based
- Section index defaults to 0 (first section)
- Paragraph and table indices are relative to their section

### 3. File Management
- Documents are stored in the configured `WORD_FILES_PATH` directory
- The server automatically creates the directory if it doesn't exist
- All file operations are relative to this directory

### 4. Async Operations
- All tools are asynchronous and must be awaited
- Handle responses properly by checking the `success` field
- Implement proper error handling for failed operations

### 5. Error Handling Best Practices
- Always check the `success` field before processing data
- Use the provided error codes for programmatic error handling
- Follow the suggestions in error responses for troubleshooting

## Examples

### Creating and Populating a Document

```python
# Create a new document
result = await create_document("report.docx")
if not result["success"]:
    print(f"Error: {result['error']['details']}")
    return

# Add a title paragraph
result = await add_paragraph("report.docx", "Annual Report 2024")
if result["success"]:
    print(f"Title added at index: {result['data']['paragraph_index']}")

# Create a table
result = await create_table("report.docx", 3, 4, style="Table Grid")
if result["success"]:
    print(f"Table created with {result['data']['dimensions']['rows']} rows")

# Add content to table cells
result = await set_cell_text("report.docx", 0, 0, 0, "Quarter")
result = await set_cell_text("report.docx", 0, 0, 1, "Revenue")
result = await set_cell_text("report.docx", 0, 0, 2, "Expenses")
result = await set_cell_text("report.docx", 0, 0, 3, "Profit")
```

### Document Conversion

```python
# Convert to PDF
result = await convert_document("report.docx", "pdf", "report.pdf")
if result["success"]:
    print(f"Document converted successfully: {result['data']['output_path']}")

# Convert to HTML
result = await convert_document("report.docx", "html", "report.html")
if result["success"]:
    print(f"Document converted to HTML: {result['data']['output_path']}")
```

### Document Comparison

# ... compare_documents 相关内容全部删除 ...

## Performance Considerations

1. **Large Documents**: For documents with many paragraphs or tables, consider processing in batches
2. **Memory Usage**: The server loads entire documents into memory, so very large documents may require more resources
3. **Concurrent Operations**: Multiple operations on the same document should be serialized to avoid conflicts
4. **File I/O**: Document operations involve file I/O, so consider this for performance-critical applications

## Troubleshooting

### Common Issues

1. **Document Not Found**: Ensure the document exists in the `WORD_FILES_PATH` directory
2. **Invalid Index**: Check that paragraph/table indices are within valid ranges
3. **Permission Errors**: Verify write permissions for the `WORD_FILES_PATH` directory
4. **Format Errors**: Ensure document names include valid file extensions

### Debug Information

- Check the server logs for detailed error information
- Use the `get_paragraph_info` and `get_table_info` tools to verify document structure
- Validate indices before performing operations

## API Versioning

The current API version is 0.1.1. Future versions will maintain backward compatibility for stable releases.
