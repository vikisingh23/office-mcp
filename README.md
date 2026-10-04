# @neuraforge/office-mcp

MCP server for creating, reading, and editing Office documents. **32 tools** for PowerPoint, Word, Excel, and CSV.

Works with Claude Code, Cursor, Gemini CLI, Kiro, and any MCP client.

## Requirements

- Node.js 18+
- **Python 3.8+ with `python-pptx` and `python-docx`** — required for every tool
  tagged "Uses python-pptx" / "Uses python-docx" below (in-place edits that
  preserve existing formatting: find/replace, insert, delete, duplicate,
  table-cell updates, structure reads). The PowerPoint/Word *creation* tools
  (`create_presentation`, `create_document`, etc.) are pure JS and don't need
  Python.

  ```bash
  pip install -r requirements.txt
  # or: pip install python-pptx python-docx
  ```

  If Python or these packages aren't available, the Python-backed tools
  return a clear `{"error": "Missing Python dependency: ..."}` instead of
  silently failing.

## Install

```json
{
  "mcpServers": {
    "office-pptx": { "command": "npx", "args": ["-y", "@neuraforge/office-mcp", "--pptx"] },
    "office-docx": { "command": "npx", "args": ["-y", "@neuraforge/office-mcp", "--docx"] },
    "office-xlsx": { "command": "npx", "args": ["-y", "@neuraforge/office-mcp", "--xlsx"] }
  }
}
```

Or run individually:
```bash
npx @neuraforge/office-mcp --pptx   # PowerPoint only
npx @neuraforge/office-mcp --docx   # Word only
npx @neuraforge/office-mcp --xlsx   # Excel/CSV only
```

## Tools

### PowerPoint (13 tools)
| Tool | Description |
|------|-------------|
| `create_presentation` | Create PPTX from structured slide data |
| `read_presentation` | Extract text and structure |
| `list_slides` | Quick summary of all slides |
| `add_slides` | Append slides to existing PPTX, preserving existing slides (uses python-pptx) |
| `add_slide_to_presentation` | Add single slide with layout (uses python-pptx) |
| `delete_slides` | Delete slides by number (uses python-pptx) |
| `delete_slide_from_presentation` | Delete single slide (uses python-pptx) |
| `modify_slide` | Replace slide content (uses python-pptx) |
| `modify_slide_text_in_presentation` | Find/replace text on slide (uses python-pptx) |
| `replace_text_in_presentation` | Find/replace across all slides (uses python-pptx) |
| `update_table_cell_in_presentation` | Update specific table cell (uses python-pptx) |
| `duplicate_slide_in_presentation` | Deep copy a slide (uses python-pptx) |
| `read_presentation_structure` | Detailed shape/text/table info (uses python-pptx) |

### Word (9 tools)
| Tool | Description |
|------|-------------|
| `create_document` | Create DOCX with headings, paragraphs, tables, images |
| `read_document` | Extract text and HTML |
| `append_to_document` | Add sections to existing document |
| `document_to_text` | Plain text extraction |
| `read_document_structure` | Detailed paragraph/style info (uses python-docx) |
| `replace_text_in_document` | Find/replace preserving formatting (uses python-docx) |
| `insert_after_text` | Insert content after specific text (uses python-docx) |
| `append_table_to_document` | Add table to existing document (uses python-docx) |
| `delete_paragraph_from_document` | Remove paragraphs by text match (uses python-docx) |

### Excel/CSV (10 tools)
| Tool | Description |
|------|-------------|
| `read_xlsx` | Read XLSX data (specific or all sheets) |
| `write_xlsx` | Write data to XLSX (create/append/replace) |
| `list_sheets` | List sheet names with row counts |
| `add_sheet` | Add new sheet to existing file |
| `delete_sheet` | Delete a sheet |
| `rename_sheet` | Rename a sheet |
| `update_cells` | Update specific cells |
| `read_csv` | Read CSV file |
| `write_csv` | Write CSV file |
| `filter_data` | Filter with conditions (equals, contains, gt, lt, regex) |

## License

Apache 2.0

Part of [NeuraForge AI](https://github.com/vikisingh23/neuraforge-ai).
