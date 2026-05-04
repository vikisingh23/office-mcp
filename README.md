# @neuraforge/office-mcp

MCP server for creating, reading, and editing Office documents. **32 tools** for PowerPoint, Word, Excel, and CSV.

Works with Claude Code, Cursor, Gemini CLI, Kiro, and any MCP client.

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
| `add_slides` | Append slides to existing PPTX |
| `add_slide_to_presentation` | Add single slide with layout |
| `delete_slides` | Delete slides by number |
| `delete_slide_from_presentation` | Delete single slide |
| `modify_slide` | Replace slide content |
| `modify_slide_text_in_presentation` | Find/replace text on slide |
| `replace_text_in_presentation` | Find/replace across all slides |
| `update_table_cell_in_presentation` | Update specific table cell |
| `duplicate_slide_in_presentation` | Deep copy a slide |
| `read_presentation_structure` | Detailed shape/text/table info |

### Word (9 tools)
| Tool | Description |
|------|-------------|
| `create_document` | Create DOCX with headings, paragraphs, tables, images |
| `read_document` | Extract text and HTML |
| `append_to_document` | Add sections to existing document |
| `document_to_text` | Plain text extraction |
| `read_document_structure` | Detailed paragraph/style info |
| `replace_text_in_document` | Find/replace preserving formatting |
| `insert_after_text` | Insert content after specific text |
| `append_table_to_document` | Add table to existing document |
| `delete_paragraph_from_document` | Remove paragraphs by text match |

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
