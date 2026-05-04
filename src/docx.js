#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  AlignmentType, ImageRun, PageBreak, Header, Footer,
  TableOfContents, ExternalHyperlink,
} from 'docx';
import mammoth from 'mammoth';
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';

const BRIDGE = path.join(path.dirname(new URL(import.meta.url).pathname), 'docx_bridge.py');

function callBridge(cmd, args) {
  const result = execFileSync('python3', [BRIDGE, cmd, JSON.stringify(args)], { encoding: 'utf-8', maxBuffer: 50 * 1024 * 1024 });
  return JSON.parse(result.trim());
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function buildParagraph(item) {
  if (typeof item === 'string') return new Paragraph({ text: item });

  const opts = {};

  // Heading
  if (item.heading) {
    const level = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4 };
    opts.heading = level[item.heading] || HeadingLevel.HEADING_1;
  }

  // Alignment
  if (item.align) {
    const map = { center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED };
    opts.alignment = map[item.align] || AlignmentType.LEFT;
  }

  // Bullet / numbered list
  if (item.bullet) opts.bullet = { level: item.indent || 0 };
  if (item.numbered) opts.numbering = { reference: 'default-numbering', level: item.indent || 0 };

  // Spacing
  if (item.spacingAfter) opts.spacing = { ...(opts.spacing || {}), after: item.spacingAfter };
  if (item.spacingBefore) opts.spacing = { ...(opts.spacing || {}), before: item.spacingBefore };

  // Page break
  if (item.pageBreak) opts.pageBreakBefore = true;

  // Text runs
  if (item.runs) {
    opts.children = item.runs.map(r => new TextRun({
      text: r.text,
      bold: r.bold, italic: r.italic, underline: r.underline ? {} : undefined,
      size: r.size, font: r.font, color: r.color,
      break: r.break ? 1 : undefined,
    }));
  } else if (item.text) {
    opts.children = [new TextRun({
      text: item.text,
      bold: item.bold, italic: item.italic, underline: item.underline ? {} : undefined,
      size: item.size, font: item.font, color: item.color,
    })];
  }

  // Hyperlink
  if (item.link) {
    opts.children = [new ExternalHyperlink({
      children: [new TextRun({ text: item.text || item.link, style: 'Hyperlink' })],
      link: item.link,
    })];
  }

  // Image
  if (item.image) {
    const imgData = fs.readFileSync(item.image);
    opts.children = [new ImageRun({
      data: imgData,
      transformation: { width: item.width || 400, height: item.height || 300 },
    })];
  }

  return new Paragraph(opts);
}

function buildTable(tableData) {
  const rows = tableData.rows.map((row, ri) =>
    new TableRow({
      children: row.map(cell => {
        const text = typeof cell === 'string' ? cell : cell.text || '';
        const isHeader = ri === 0 && tableData.headerRow !== false;
        return new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text, bold: isHeader, size: isHeader ? 22 : 20, font: 'Arial' })],
            alignment: AlignmentType.LEFT,
          })],
          shading: isHeader ? { fill: 'EDF1FF' } : undefined,
          width: cell.width ? { size: cell.width, type: WidthType.PERCENTAGE } : undefined,
        });
      }),
    })
  );

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
    },
  });
}

function buildSection(section) {
  const children = [];
  for (const item of section.content || []) {
    if (item.table) children.push(buildTable(item.table));
    else children.push(buildParagraph(item));
  }

  const sectionOpts = { children };

  if (section.header) {
    sectionOpts.headers = {
      default: new Header({ children: [new Paragraph({ text: section.header, alignment: AlignmentType.RIGHT })] }),
    };
  }
  if (section.footer) {
    sectionOpts.footers = {
      default: new Footer({ children: [new Paragraph({ text: section.footer, alignment: AlignmentType.CENTER })] }),
    };
  }

  return sectionOpts;
}

// ── MCP Server ───────────────────────────────────────────────────────────────

const server = new Server(
  { name: 'docx-crud-server', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

const TOOLS = [
  {
    name: 'create_document',
    description: 'Create a new Word document (.docx) from structured content. Supports headings, paragraphs, bold/italic/underline text, bullet lists, numbered lists, tables, images, hyperlinks, headers, footers, and page breaks.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Output .docx file path' },
        title: { type: 'string', description: 'Document title metadata' },
        author: { type: 'string', description: 'Author metadata' },
        sections: {
          type: 'array',
          description: 'Array of section objects. Each section has: content (array of paragraph/table items), header (string), footer (string). Content items can be: string, {text, bold, italic, heading:1-4, bullet:true, numbered:true, align, size, color, font}, {table:{rows:[][], headerRow:true}}, {image:"/path", width, height}, {link:"url", text}',
          items: { type: 'object' },
        },
      },
      required: ['filePath', 'sections'],
    },
  },
  {
    name: 'read_document',
    description: 'Read and extract content from an existing Word document. Returns both raw text and HTML representation.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to the .docx file' },
        format: { type: 'string', enum: ['text', 'html', 'both'], description: 'Output format (default: both)' },
      },
      required: ['filePath'],
    },
  },
  {
    name: 'append_to_document',
    description: 'Read an existing Word document, then create a new document with the original content plus appended sections. Note: formatting of original content is preserved as text.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to existing .docx file' },
        outputPath: { type: 'string', description: 'Output path (defaults to overwrite input)' },
        sections: {
          type: 'array',
          description: 'Sections to append (same format as create_document)',
          items: { type: 'object' },
        },
      },
      required: ['filePath', 'sections'],
    },
  },
  {
    name: 'document_to_text',
    description: 'Extract plain text from a Word document — useful for analysis, summarization, or feeding into other tools',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to the .docx file' },
      },
      required: ['filePath'],
    },
  },
  {
    name: 'read_document_structure',
    description: 'Read detailed structure of a Word document — paragraphs with styles/formatting, tables with data. Preserves formatting info. Uses python-docx.',
    inputSchema: {
      type: 'object',
      properties: { filePath: { type: 'string' } },
      required: ['filePath'],
    },
  },
  {
    name: 'replace_text_in_document',
    description: 'Find and replace text in a Word document IN-PLACE. Preserves all formatting (bold, italic, fonts, styles, images, tables). Uses python-docx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to the .docx file' },
        outputPath: { type: 'string', description: 'Output path (defaults to overwrite input)' },
        replacements: { type: 'object', description: 'Key-value pairs: {"old text": "new text", ...}' },
      },
      required: ['filePath', 'replacements'],
    },
  },
  {
    name: 'insert_after_text',
    description: 'Insert paragraphs after a specific text in a Word document. Preserves existing formatting. Uses python-docx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        afterText: { type: 'string', description: 'Text to find — new content inserted after this paragraph' },
        paragraphs: { type: 'array', description: 'Array of {text, heading:1-4, bold, italic, bullet, size, color}', items: { type: 'object' } },
      },
      required: ['filePath', 'afterText', 'paragraphs'],
    },
  },
  {
    name: 'append_table_to_document',
    description: 'Append a table to an existing Word document. Preserves all existing content. Uses python-docx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        headers: { type: 'array', items: { type: 'string' }, description: 'Column headers' },
        rows: { type: 'array', description: '2D array of cell values', items: { type: 'array' } },
      },
      required: ['filePath', 'headers', 'rows'],
    },
  },
  {
    name: 'delete_paragraph_from_document',
    description: 'Delete paragraphs containing specific text from a Word document. Preserves everything else. Uses python-docx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        containingText: { type: 'string', description: 'Delete paragraphs containing this text' },
      },
      required: ['filePath', 'containingText'],
    },
  },
];

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {

      case 'create_document': {
        const sections = args.sections.map(buildSection);
        const doc = new Document({
          title: args.title,
          creator: args.author || 'AMC OneView',
          numbering: {
            config: [{ reference: 'default-numbering', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT }] }],
          },
          sections,
        });

        const buf = await Packer.toBuffer(doc);
        const outPath = path.resolve(args.filePath);
        const dir = path.dirname(outPath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(outPath, buf);

        const sectionCount = args.sections.length;
        const contentCount = args.sections.reduce((sum, s) => sum + (s.content?.length || 0), 0);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, sections: sectionCount, contentBlocks: contentCount }) }] };
      }

      case 'read_document': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        const buf = fs.readFileSync(fp);
        const format = args.format || 'both';
        const result = { filePath: fp };

        if (format === 'text' || format === 'both') {
          const textResult = await mammoth.extractRawText({ buffer: buf });
          result.text = textResult.value;
          result.textLength = textResult.value.length;
        }
        if (format === 'html' || format === 'both') {
          const htmlResult = await mammoth.convertToHtml({ buffer: buf });
          result.html = htmlResult.value;
          result.warnings = htmlResult.messages.filter(m => m.type === 'warning').map(m => m.message);
        }

        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'append_to_document': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        // Read existing content as text
        const buf = fs.readFileSync(fp);
        const existing = await mammoth.extractRawText({ buffer: buf });
        const existingLines = existing.value.split('\n').filter(l => l.trim());

        // Build new doc: existing text + new sections
        const existingSection = {
          content: existingLines.map(line => ({ text: line })),
        };
        const allSections = [existingSection, ...args.sections].map(buildSection);

        const doc = new Document({
          creator: 'AMC OneView',
          numbering: {
            config: [{ reference: 'default-numbering', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT }] }],
          },
          sections: allSections,
        });

        const newBuf = await Packer.toBuffer(doc);
        const outPath = path.resolve(args.outputPath || args.filePath);
        fs.writeFileSync(outPath, newBuf);

        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, existingLines: existingLines.length, appendedSections: args.sections.length }) }] };
      }

      case 'document_to_text': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        const buf = fs.readFileSync(fp);
        const result = await mammoth.extractRawText({ buffer: buf });
        return { content: [{ type: 'text', text: result.value }] };
      }

      case 'read_document_structure': {
        const result = callBridge('read_structure', { path: path.resolve(args.filePath) });
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'replace_text_in_document': {
        const result = callBridge('replace_text', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), replacements: args.replacements });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'insert_after_text': {
        const result = callBridge('insert_after', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), afterText: args.afterText, paragraphs: args.paragraphs });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'append_table_to_document': {
        const result = callBridge('append_table', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), headers: args.headers, rows: args.rows });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'delete_paragraph_from_document': {
        const result = callBridge('delete_paragraph', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), containingText: args.containingText });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      default:
        return { content: [{ type: 'text', text: JSON.stringify({ error: `Unknown tool: ${name}` }) }] };
    }
  } catch (err) {
    return { content: [{ type: 'text', text: JSON.stringify({ error: err.message, stack: err.stack?.split('\n').slice(0, 3) }) }] };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(console.error);
