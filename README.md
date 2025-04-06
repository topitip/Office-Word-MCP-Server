# Office-Word-MCP-Server

A Model Context Protocol (MCP) server for creating, reading, and manipulating Microsoft Word documents. This server enables AI assistants to work with Word documents through a standardized interface, providing rich document editing capabilities.

[![Office Word Server MCP server](https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server/badge)](https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server)

![MCP Server](https://badge.mcpx.dev?type=server)

## Overview

Office-Word-MCP-Server implements the [Model Context Protocol](https://modelcontextprotocol.io/) to expose Word document operations as tools and resources. It serves as a bridge between AI assistants and Microsoft Word documents, allowing for document creation, content addition, formatting, and analysis.

### Recent Updates

- **2024-04-06**: Improved support for optional parameters. Now formatting parameters, including alignment, can be passed directly to functions like `add_heading`, `add_paragraph` and others without the need for additional methods.

### Example

#### Pormpt

![image](https://github.com/user-attachments/assets/f49b0bcc-88b2-4509-bf50-995b9a40038c)

#### Output

![image](https://github.com/user-attachments/assets/ff64385d-3822-4160-8cdf-f8a484ccc01a)

## Features

### Document Management

- Create new Word documents with metadata
- Extract text and analyze document structure
- View document properties and statistics
- List available documents in a directory
- Create copies of existing documents

### Content Creation

- Add headings with different levels
- Insert paragraphs with optional styling
- Create tables with custom data
- Add images with proportional scaling
- Insert page breaks

### Text Formatting

- Format specific text sections (bold, italic, underline)
- Change text color and font properties
- Apply custom styles to text elements
- Search and replace text throughout documents
- Control paragraph alignment (left, center, right, justified)

### Table Formatting

- Format tables with borders and styles
- Create header rows with distinct formatting
- Apply cell shading and custom borders
- Structure tables for better readability

### Advanced Document Manipulation

- Delete paragraphs
- Create custom document styles
- Apply consistent formatting throughout documents
- Format specific ranges of text with detailed control

## Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Basic Installation

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install dependencies
pip install -r requirements.txt
```

### Using the Setup Script

Alternatively, you can use the provided setup script which handles:

- Checking prerequisites
- Setting up a virtual environment
- Installing dependencies
- Generating MCP configuration

```bash
python setup_mcp.py
```

## Usage with Claude for Desktop

### Configuration

#### Method 1: After Local Installation

1. After installation, add the server to your Claude for Desktop configuration file:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": [
        "/path/to/word_server.py"
      ]
    }
  }
}
```

#### Method 2: Without Installation (Using uvx)

1. You can also configure Claude for Desktop to use the server without local installation by using the uvx package manager:

    ```json
    {
      "mcpServers": {
        "word-document-server": {
          "command": "uvx",
          "args": [
            "--from", "office-word-mcp-server", "word_mcp_server"
          ]
        }
      }
    }
    ```

2. Configuration file locations:
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

3. Restart Claude for Desktop to load the configuration.

### Example Operations

Once configured, you can ask Claude to perform operations like:

- "Create a new document called 'report.docx' with a title page"
- "Add a heading and three paragraphs to my document"
- "Add a centered heading to my document"
- "Insert a 4x4 table with sales data"
- "Format the word 'important' in paragraph 2 to be bold and red"
- "Search and replace all instances of 'old term' with 'new term'"
- "Create a custom style for section headings"
- "Apply formatting to the table in my document"
- "Add a right-aligned paragraph with summary information"
- "Create a document with justified text paragraphs"

## API Reference

### Document Creation and Properties

```python
create_document(filename, title=None, author=None)
get_document_info(filename)
get_document_text(filename)
get_document_outline(filename)
list_available_documents(directory=".")
copy_document(source_filename, destination_filename=None)
```

### Content Addition

```python
add_heading(filename, text, level=1, alignment=None)
add_paragraph(filename, text, style=None, alignment=None)
add_table(filename, rows, cols, data=None)
add_picture(filename, image_path, width=None)
add_page_break(filename)
```

### Text Formatting

```python
format_text(filename, paragraph_index, start_pos, end_pos, bold=None, 
            italic=None, underline=None, color=None, font_size=None, font_name=None)
search_and_replace(filename, find_text, replace_text)
delete_paragraph(filename, paragraph_index)
create_custom_style(filename, style_name, bold=None, italic=None, 
                    font_size=None, font_name=None, color=None, base_style=None)
set_paragraph_alignment(filename, paragraph_index, alignment)
```

### Table Formatting

```python
format_table(filename, table_index, has_header_row=None, 
             border_style=None, shading=None)
```

### Paragraph Alignment Notes

There are two approaches to align paragraphs (left, center, right, justify):

#### Method 1: Direct alignment specification when creating
```python
# Creating a paragraph or heading with alignment
add_paragraph("document.docx", "This paragraph is centered", alignment="center")
add_heading("document.docx", "This heading is centered", level=1, alignment="center")
```

#### Method 2: Setting alignment for an existing paragraph
```python
# Creating a paragraph
add_paragraph("document.docx", "This is a paragraph")

# Setting alignment for the paragraph
set_paragraph_alignment("document.docx", 0, "center")  # 0 = index of the first paragraph
```

Both approaches support the following alignment values: 'left', 'center', 'right', 'justify'.

## Troubleshooting

### Common Issues

1. **Missing Styles**
   - Some documents may lack required styles for heading and table operations
   - The server will attempt to create missing styles or use direct formatting
   - For best results, use templates with standard Word styles

2. **Permission Issues**
   - Ensure the server has permission to read/write to the document paths
   - Use the `copy_document` function to create editable copies of locked documents
   - Check file ownership and permissions if operations fail

3. **Image Insertion Problems**
   - Use absolute paths for image files
   - Verify image format compatibility (JPEG, PNG recommended)
   - Check image file size and permissions

### Known Limitations

1. **Complex Document Structures**
   - Some complex document structures (e.g., text boxes, charts) are not fully supported
   - Focus on core document elements like paragraphs, tables, and images

### Debugging

Enable detailed logging by setting the environment variable:

```bash
export MCP_DEBUG=1  # Linux/macOS
set MCP_DEBUG=1     # Windows
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) for the protocol specification
- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [FastMCP](https://github.com/modelcontextprotocol/python-sdk) for the Python MCP implementation

## Docker Support

### Using Docker

This MCP server can be run in a Docker container for easier deployment and isolation.

```bash
# Build the Docker image
docker build -t word-mcp-server:latest .

# Run the container
docker run -i --rm --name word-mcp-server -v $(pwd)/documents:/app/documents word-mcp-server:latest
```

### Using Docker Compose

```bash
# Start the service
docker-compose up -d

# View logs
docker-compose logs -f

# Stop the service
docker-compose down
```

### Connection Configuration

To connect to the Docker containerized MCP server, add the following configuration to your MCP client:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "-v",
        "$(pwd)/documents:/app/documents",
        "word-mcp-server:latest"
      ]
    }
  }
}
```

For Claude for Desktop or other MCP clients, add this configuration to their respective config files.

---

*Note: This server interacts with document files on your system. Always verify that requested operations are appropriate before confirming them in Claude for Desktop or other MCP clients.*
