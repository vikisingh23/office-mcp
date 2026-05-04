#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';

const BRIDGE = path.join(path.dirname(new URL(import.meta.url).pathname), 'pptx_bridge.py');

function callBridge(cmd, args) {
  const result = execFileSync('python3', [BRIDGE, cmd, JSON.stringify(args)], { encoding: 'utf-8', maxBuffer: 50 * 1024 * 1024 });
  return JSON.parse(result.trim());
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function stripXmlTags(xml) {
  return xml.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}

async function readPptx(filePath) {
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);
  const slides = [];

  // Sort slide files numerically
  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1]);
      const nb = parseInt(b.match(/slide(\d+)/)[1]);
      return na - nb;
    });

  for (const sf of slideFiles) {
    const xml = await zip.files[sf].async('text');
    const texts = [];
    // Extract all <a:t> text runs
    const matches = xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g);
    for (const m of matches) texts.push(m[1]);

    // Detect images
    const images = (xml.match(/<a:blip/g) || []).length;

    // Detect tables
    const tables = (xml.match(/<a:tbl>/g) || []).length;

    // Detect charts (via relationships)
    const charts = (xml.match(/<c:chart/g) || []).length;

    slides.push({
      slideNumber: slides.length + 1,
      file: sf,
      texts,
      imageCount: images,
      tableCount: tables,
      chartCount: charts,
    });
  }

  // Read slide dimensions from presentation.xml
  let width = 10, height = 5.625;
  if (zip.files['ppt/presentation.xml']) {
    const presXml = await zip.files['ppt/presentation.xml'].async('text');
    const szMatch = presXml.match(/cy="(\d+)".*?cx="(\d+)"/s) || presXml.match(/cx="(\d+)".*?cy="(\d+)"/s);
    if (szMatch) {
      // EMU to inches
      width = parseInt(szMatch[1]) / 914400;
      height = parseInt(szMatch[2]) / 914400;
    }
  }

  return { filePath, slideCount: slides.length, width, height, slides };
}

function applySlideContent(slide, content, pptx) {
  // Title
  if (content.title) {
    slide.addText(content.title, {
      x: content.titleOpts?.x ?? 0.5,
      y: content.titleOpts?.y ?? 0.3,
      w: content.titleOpts?.w ?? '90%',
      fontSize: content.titleOpts?.fontSize ?? 28,
      fontFace: content.titleOpts?.fontFace ?? 'Arial',
      bold: content.titleOpts?.bold !== false,
      color: content.titleOpts?.color ?? '2E2A94',
      ...(content.titleOpts || {}),
    });
  }

  // Body text
  if (content.body) {
    const bodyText = Array.isArray(content.body) ? content.body : [content.body];
    slide.addText(
      bodyText.map(t => typeof t === 'string' ? { text: t, options: { breakLine: true } } : t),
      {
        x: content.bodyOpts?.x ?? 0.5,
        y: content.bodyOpts?.y ?? 1.5,
        w: content.bodyOpts?.w ?? '90%',
        h: content.bodyOpts?.h ?? '60%',
        fontSize: content.bodyOpts?.fontSize ?? 16,
        fontFace: content.bodyOpts?.fontFace ?? 'Arial',
        color: content.bodyOpts?.color ?? '333333',
        valign: 'top',
        ...(content.bodyOpts || {}),
      }
    );
  }

  // Bullet points
  if (content.bullets) {
    slide.addText(
      content.bullets.map(b => ({
        text: typeof b === 'string' ? b : b.text,
        options: { bullet: true, breakLine: true, indentLevel: b.indent ?? 0 },
      })),
      {
        x: content.bulletsOpts?.x ?? 0.5,
        y: content.bulletsOpts?.y ?? 1.5,
        w: content.bulletsOpts?.w ?? '90%',
        h: content.bulletsOpts?.h ?? '65%',
        fontSize: content.bulletsOpts?.fontSize ?? 16,
        fontFace: content.bulletsOpts?.fontFace ?? 'Arial',
        color: content.bulletsOpts?.color ?? '333333',
        valign: 'top',
        ...(content.bulletsOpts || {}),
      }
    );
  }

  // Table
  if (content.table) {
    const rows = content.table.map(row =>
      row.map(cell => typeof cell === 'string' ? { text: cell } : cell)
    );
    slide.addTable(rows, {
      x: content.tableOpts?.x ?? 0.5,
      y: content.tableOpts?.y ?? 1.5,
      w: content.tableOpts?.w ?? 9,
      fontSize: content.tableOpts?.fontSize ?? 12,
      border: { pt: 1, color: 'CFCFCF' },
      colW: content.tableOpts?.colW,
      autoPage: true,
      ...(content.tableOpts || {}),
    });
  }

  // Image
  if (content.image) {
    const imgOpts = {
      x: content.image.x ?? 1,
      y: content.image.y ?? 1.5,
      w: content.image.w ?? 4,
      h: content.image.h ?? 3,
    };
    if (content.image.path) imgOpts.path = content.image.path;
    else if (content.image.data) imgOpts.data = content.image.data;
    slide.addImage(imgOpts);
  }

  // Shape / rectangle
  if (content.shapes) {
    for (const s of content.shapes) {
      slide.addShape(pptx.ShapeType[s.type] || pptx.ShapeType.rect, {
        x: s.x ?? 0, y: s.y ?? 0, w: s.w ?? 2, h: s.h ?? 1,
        fill: { color: s.fill ?? 'EFEFEF' },
        line: s.line ? { color: s.line, width: s.lineWidth ?? 1 } : undefined,
      });
    }
  }

  // Background
  if (content.background) {
    if (typeof content.background === 'string') {
      slide.background = { fill: content.background };
    } else {
      slide.background = content.background;
    }
  }

  // Notes
  if (content.notes) {
    slide.addNotes(content.notes);
  }
}

async function createPptxFromSlides(slides, opts = {}) {
  const pptx = new PptxGenJS();
  pptx.title = opts.title || 'Presentation';
  pptx.author = opts.author || 'AMC OneView';
  if (opts.layout) pptx.layout = opts.layout;

  for (const content of slides) {
    const slide = pptx.addSlide();
    applySlideContent(slide, content, pptx);
  }

  return pptx;
}

// ── MCP Server ───────────────────────────────────────────────────────────────

const server = new Server(
  { name: 'pptx-crud-server', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

const TOOLS = [
  {
    name: 'create_presentation',
    description: 'Create a new PowerPoint presentation from structured slide data. Each slide can have title, body, bullets, table, image, shapes, background, and notes.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Output .pptx file path' },
        title: { type: 'string', description: 'Presentation title metadata' },
        author: { type: 'string', description: 'Author metadata' },
        layout: { type: 'string', description: 'Slide layout: LAYOUT_WIDE (default), LAYOUT_16x9, LAYOUT_4x3' },
        slides: {
          type: 'array',
          description: 'Array of slide objects. Each can have: title, body (string or array), bullets (array of strings), table (2D array), image ({path,x,y,w,h}), shapes, background, notes',
          items: { type: 'object' },
        },
      },
      required: ['filePath', 'slides'],
    },
  },
  {
    name: 'read_presentation',
    description: 'Read and extract all text content, structure, and metadata from an existing PPTX file',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to the .pptx file' },
      },
      required: ['filePath'],
    },
  },
  {
    name: 'list_slides',
    description: 'Quick summary of all slides in a PPTX — slide numbers, first text line, element counts',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to the .pptx file' },
      },
      required: ['filePath'],
    },
  },
  {
    name: 'add_slides',
    description: 'Add new slides to an existing PPTX. Reads the file, appends slides, writes back.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to existing .pptx file' },
        outputPath: { type: 'string', description: 'Output path (defaults to overwrite input)' },
        slides: {
          type: 'array',
          description: 'Slides to append (same format as create_presentation)',
          items: { type: 'object' },
        },
      },
      required: ['filePath', 'slides'],
    },
  },
  {
    name: 'delete_slides',
    description: 'Delete slides by index (1-based) from a PPTX. Rebuilds the presentation without those slides.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to existing .pptx file' },
        outputPath: { type: 'string', description: 'Output path (defaults to overwrite input)' },
        slideNumbers: {
          type: 'array',
          description: 'Slide numbers to delete (1-based)',
          items: { type: 'number' },
        },
      },
      required: ['filePath', 'slideNumbers'],
    },
  },
  {
    name: 'modify_slide',
    description: 'Replace content of a specific slide (by index) in an existing PPTX. Rebuilds the presentation with the updated slide.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: 'Path to existing .pptx file' },
        outputPath: { type: 'string', description: 'Output path (defaults to overwrite input)' },
        slideNumber: { type: 'number', description: 'Slide number to replace (1-based)' },
        content: { type: 'object', description: 'New slide content (same format as create_presentation slides)' },
      },
      required: ['filePath', 'slideNumber', 'content'],
    },
  },
  {
    name: 'read_presentation_structure',
    description: 'Read detailed structure of a PPTX — shapes, text frames, tables, formatting per slide. Preserves all info. Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: { filePath: { type: 'string' } },
      required: ['filePath'],
    },
  },
  {
    name: 'replace_text_in_presentation',
    description: 'Find and replace text across all slides IN-PLACE. Preserves formatting (fonts, colors, sizes, layouts, images). Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        replacements: { type: 'object', description: '{"old text": "new text", ...}' },
      },
      required: ['filePath', 'replacements'],
    },
  },
  {
    name: 'add_slide_to_presentation',
    description: 'Add a new slide to an existing PPTX using its own layouts. Preserves all existing slides with formatting. Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        layoutIndex: { type: 'number', description: 'Slide layout index (0=title, 1=title+content, 5=blank, etc.)' },
        content: { type: 'object', description: '{title, body (string or array), notes}' },
      },
      required: ['filePath', 'content'],
    },
  },
  {
    name: 'delete_slide_from_presentation',
    description: 'Delete a slide by number from a PPTX. Preserves all other slides with formatting. Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        slideNumber: { type: 'number', description: '1-based slide number to delete' },
      },
      required: ['filePath', 'slideNumber'],
    },
  },
  {
    name: 'modify_slide_text_in_presentation',
    description: 'Replace text on a specific slide IN-PLACE. Preserves formatting. Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        slideNumber: { type: 'number' },
        replacements: { type: 'object', description: '{"old text": "new text"}' },
      },
      required: ['filePath', 'slideNumber', 'replacements'],
    },
  },
  {
    name: 'update_table_cell_in_presentation',
    description: 'Update a specific table cell on a slide. Preserves all formatting. Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        slideNumber: { type: 'number' },
        tableIndex: { type: 'number', description: '0-based table index on the slide (default 0)' },
        row: { type: 'number' },
        col: { type: 'number' },
        value: { type: 'string' },
      },
      required: ['filePath', 'slideNumber', 'row', 'col', 'value'],
    },
  },
  {
    name: 'duplicate_slide_in_presentation',
    description: 'Duplicate a slide (deep copy with all formatting, shapes, images). Uses python-pptx.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: { type: 'string' },
        outputPath: { type: 'string' },
        slideNumber: { type: 'number', description: 'Slide to duplicate (1-based)' },
      },
      required: ['filePath', 'slideNumber'],
    },
  },
];

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {

      case 'create_presentation': {
        const pptx = await createPptxFromSlides(args.slides, {
          title: args.title,
          author: args.author,
          layout: args.layout,
        });
        const outPath = path.resolve(args.filePath);
        const dir = path.dirname(outPath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        await pptx.writeFile({ fileName: outPath });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, slideCount: args.slides.length }) }] };
      }

      case 'read_presentation': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };
        const data = await readPptx(fp);
        return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
      }

      case 'list_slides': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };
        const data = await readPptx(fp);
        const summary = data.slides.map(s => ({
          slide: s.slideNumber,
          preview: s.texts[0]?.substring(0, 80) || '(no text)',
          textBlocks: s.texts.length,
          images: s.imageCount,
          tables: s.tableCount,
        }));
        return { content: [{ type: 'text', text: JSON.stringify({ slideCount: data.slideCount, slides: summary }, null, 2) }] };
      }

      case 'add_slides': {
        // Read existing, extract text per slide, rebuild with new slides appended
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        const existing = await readPptx(fp);
        // Rebuild: existing slides as text-only + new slides
        const rebuiltSlides = existing.slides.map(s => ({
          title: s.texts[0] || '',
          body: s.texts.slice(1),
        }));
        const allSlides = [...rebuiltSlides, ...args.slides];
        const pptx = await createPptxFromSlides(allSlides, { title: 'Updated Presentation' });
        const outPath = path.resolve(args.outputPath || args.filePath);
        await pptx.writeFile({ fileName: outPath });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, previousSlides: existing.slideCount, addedSlides: args.slides.length, totalSlides: allSlides.length }) }] };
      }

      case 'delete_slides': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        const existing = await readPptx(fp);
        const toDelete = new Set(args.slideNumbers);
        const keptSlides = existing.slides
          .filter(s => !toDelete.has(s.slideNumber))
          .map(s => ({ title: s.texts[0] || '', body: s.texts.slice(1) }));

        const pptx = await createPptxFromSlides(keptSlides, { title: 'Updated Presentation' });
        const outPath = path.resolve(args.outputPath || args.filePath);
        await pptx.writeFile({ fileName: outPath });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, deletedSlides: args.slideNumbers, remainingSlides: keptSlides.length }) }] };
      }

      case 'modify_slide': {
        const fp = path.resolve(args.filePath);
        if (!fs.existsSync(fp)) return { content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${fp}` }) }] };

        const existing = await readPptx(fp);
        if (args.slideNumber < 1 || args.slideNumber > existing.slideCount) {
          return { content: [{ type: 'text', text: JSON.stringify({ error: `Slide ${args.slideNumber} out of range (1-${existing.slideCount})` }) }] };
        }

        const rebuiltSlides = existing.slides.map(s => {
          if (s.slideNumber === args.slideNumber) return args.content;
          return { title: s.texts[0] || '', body: s.texts.slice(1) };
        });

        const pptx = await createPptxFromSlides(rebuiltSlides, { title: 'Updated Presentation' });
        const outPath = path.resolve(args.outputPath || args.filePath);
        await pptx.writeFile({ fileName: outPath });
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, filePath: outPath, modifiedSlide: args.slideNumber, totalSlides: rebuiltSlides.length }) }] };
      }

      // === python-pptx powered tools (preserve formatting) ===

      case 'read_presentation_structure': {
        const result = callBridge('read_structure', { path: path.resolve(args.filePath) });
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'replace_text_in_presentation': {
        const result = callBridge('replace_text', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), replacements: args.replacements });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'add_slide_to_presentation': {
        const result = callBridge('add_slide', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), layoutIndex: args.layoutIndex || 1, content: args.content });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'delete_slide_from_presentation': {
        const result = callBridge('delete_slide', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), slideNumber: args.slideNumber });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'modify_slide_text_in_presentation': {
        const result = callBridge('modify_slide_text', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), slideNumber: args.slideNumber, replacements: args.replacements });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'update_table_cell_in_presentation': {
        const result = callBridge('update_table_cell', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), slideNumber: args.slideNumber, tableIndex: args.tableIndex || 0, row: args.row, col: args.col, value: args.value });
        return { content: [{ type: 'text', text: JSON.stringify(result) }] };
      }

      case 'duplicate_slide_in_presentation': {
        const result = callBridge('duplicate_slide', { path: path.resolve(args.filePath), output: path.resolve(args.outputPath || args.filePath), slideNumber: args.slideNumber });
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
