# What's Spire.Doc MCP Server ?

Spire.Doc MCP Server is a Model Context Protocol (MCP) server that enables you to manipulate Word documents without requiring Microsoft Word. With it, you can [create](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Create-Word-Document.html), [read](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/read-word-doc-or-docx-files-in-python.html), and [modify](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Edit-or-Modify-a-Word-Document.html) Word documents directly using your AI agent.


## Key Features

- [Create Word documents from Scratch](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Create-Word-Document.html)
- [Modify existing Word files](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Edit-or-Modify-a-Word-Document.html)
- Convert [Word to PDF](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-PDF.html), [Word to Images](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-Images.html), and more with high fidelity
- [Read](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/read-word-doc-or-docx-files-in-python.html) and [write document content](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Create-Word-Document.html/)
- Apply formatting and styles
- Analyze document content
- [Table operations and management](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Table/Python-Create-Tables-in-a-Word-Document.html)
- [Paragraph manipulation and formatting](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Create-Word-Document.html#4)
- [Document protection and security](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Security/Python-Encrypt-or-Decrypt-Word-Documents.html)
- [Document comparison](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Compare-Two-Versions-of-a-Word-Document.html) and [merging](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Document-Operation/Python-Merge-Word-Documents.html)
- [Text watermark](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Watermark/Python-Insert-Watermarks-in-Word.html) support

## Quick Start

### Prerequisites

- Python 3.10 or higher

### Installation

1. Clone the repository:

```bash
git clone https://github.com/eiceblue/spire-doc-mcp-server.git
cd spire-doc-mcp-server
```

2. Install using uv:

```bash
uv pip install -e .
```

### Running the Server

Start the server (default port 8000):

```bash
uv run spire-doc-mcp-server
```

Custom port (e.g., 8080):

```bash
# Bash/Linux/macOS
export SPIRE_DOC_MCP_PORT=8080 && uv run spire-doc-mcp-server

# Windows PowerShell
$env:SPIRE_DOC_MCP_PORT = "8080"; uv run spire-doc-mcp-server
```

## Integration with AI Tools

### Cursor IDE

1. Add this configuration to Cursor:

```json
{
  "mcpServers": {
    "word": {
      "url": "http://localhost:8000/sse",
      "env": {
        "WORD_FILES_PATH": "/path/to/word/files"
      }
    }
  }
}
```

2. The Word tools will be available through your AI assistant.

## Environment Variables

| Variable | Description | Default |
|--------|------|--------|
| `WORD_FILES_PATH` | Directory for Word files | `./word_files` |

## Available Tools

The server provides **18 tools** organized into 5 categories:

### Document Operations (6 tools)

- **create_document**: Create new Word documents with optional templates
- **set_document_protection**: Apply password protection to documents
- **find_and_replace**: Search and replace text in documents
- **merge_documents**: Merge multiple documents into one
- **add_text_watermark**: Add text watermarks to documents

### Paragraph Operations (5 tools)

- **get_paragraph_text**: Retrieve text content from paragraphs
- **delete_paragraph**: Remove paragraphs from documents
- **get_paragraph_info**: Get detailed paragraph information
- **add_paragraph**: Add new paragraphs to documents
- **update_paragraph_text**: Modify paragraph text content

### Table Operations (6 tools)

- **create_table**: Create new tables with specified dimensions
- **add_table_after_paragraph**: Insert tables after specific paragraphs
- **add_table_to_section**: Add tables to the end of sections
- **get_table_info**: Get table information and structure
- **delete_table**: Remove tables from documents
- **set_cell_text**: Set text content in table cells

### Formatting Operations (2 tools)

- **format_paragraph**: Apply formatting to paragraphs (alignment, indentation, spacing)
- **get_paragraph_format**: Retrieve paragraph formatting information

### Conversion Operations (3 tools)

- **convert_document**: Convert documents between different formats
- **get_conversion_status**: Check conversion operation status
- **get_conversion_history**: View conversion history

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
- **[PDF](https://https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-PDF.html)**: Portable Document Format
- **[RTF](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-RTF-and-Vice-Versa.html)**: Rich Text Format
- **[HTML](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-HTML.html)**: HyperText Markup Language
- **[TXT](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Text-to-Word-or-Word-to-Text.html)**: Plain Text
- **[EPUB](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-EPUB.html)**: Electronic Publication
- **ODT**: OpenDocument Text
- **[XML](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-XML-Word-XML.html)**: Extensible Markup Language
- **[Markdown](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Markdown-to-Word-or-Word-to-Markdown.html)**: Markdown Format

## Development

### Project Structure

```
spire-doc-mcp-server/
├── src/spire_doc_mcp/
│   ├── api/                 # API layer (18 tools)
│   │   ├── document_tools.py
│   │   ├── paragraph_tools.py
│   │   ├── table_tools.py
│   │   ├── format_tools.py
│   │   └── conversion_tools.py
│   ├── core/               # Core business logic
│   │   ├── document.py
│   │   ├── paragraph.py
│   │   ├── table.py
│   │   ├── format.py
│   │   ├── conversion.py
│   │   └── base.py
│   ├── utils/              # Utilities and constants
│   │   ├── exceptions.py
│   │   └── constants.py
│   ├── server.py           # MCP server implementation
│   └── __main__.py         # Entry point
├── pyproject.toml          # Project configuration
├── README.md              # This file
└── TOOLS.md               # Detailed tool documentation
```

### Running Tests

```bash
uv run pytest
```

### Code Quality

The project follows Python best practices:

- Type hints throughout
- Comprehensive error handling
- Consistent code formatting
- Detailed documentation
- Modular architecture

## Documentation

- **[TOOLS.md](TOOLS.md)**: Comprehensive tool documentation with examples
- **[API Reference](TOOLS.md#detailed-tool-documentation)**: Detailed API documentation
- **[Error Handling](TOOLS.md#error-handling)**: Error codes and troubleshooting

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For support and questions:

- Check the [TOOLS.md](TOOLS.md) documentation
- Review the error handling section
- Open an issue on GitHub

## Version History

- **v1.0.0**: Initial stable release with 18 tools
- Complete Word document manipulation capabilities
- Comprehensive error handling
- Full format conversion support

