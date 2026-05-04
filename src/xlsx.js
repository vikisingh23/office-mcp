#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import XLSX from 'xlsx';
import fs from 'fs';
import csv from 'csv-parser';
import createCsvWriter from 'csv-writer';

class XlsxCsvCrudServer {
  constructor() {
    this.server = new Server(
      { name: 'xlsx-csv-crud-server', version: '2.0.0' },
      { capabilities: { tools: {} } }
    );
    this.setupToolHandlers();
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'read_xlsx',
          description: 'Read data from an XLSX file. Can read specific sheet or all sheets.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              sheetName: { type: 'string', description: 'Sheet name (optional, defaults to first sheet). Use "*" to read all sheets.' }
            },
            required: ['filePath']
          }
        },
        {
          name: 'write_xlsx',
          description: 'Write data to an XLSX file. Creates new file or adds/replaces a sheet in an existing file. Preserves other sheets.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              data: { type: 'array', description: 'Array of objects to write' },
              sheetName: { type: 'string', description: 'Sheet name (default: "Sheet1")' },
              mode: { type: 'string', enum: ['create', 'append', 'replace'], description: 'create: new file (default), append: add rows to existing sheet, replace: replace sheet in existing file' }
            },
            required: ['filePath', 'data']
          }
        },
        {
          name: 'list_sheets',
          description: 'List all sheet names in an XLSX file with row counts.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' }
            },
            required: ['filePath']
          }
        },
        {
          name: 'add_sheet',
          description: 'Add a new sheet to an existing XLSX file. Fails if sheet already exists.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              sheetName: { type: 'string', description: 'Name for the new sheet' },
              data: { type: 'array', description: 'Array of objects for the new sheet' }
            },
            required: ['filePath', 'sheetName', 'data']
          }
        },
        {
          name: 'delete_sheet',
          description: 'Delete a sheet from an XLSX file. Cannot delete the last remaining sheet.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              sheetName: { type: 'string', description: 'Sheet name to delete' }
            },
            required: ['filePath', 'sheetName']
          }
        },
        {
          name: 'rename_sheet',
          description: 'Rename a sheet in an XLSX file.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              oldName: { type: 'string', description: 'Current sheet name' },
              newName: { type: 'string', description: 'New sheet name' }
            },
            required: ['filePath', 'oldName', 'newName']
          }
        },
        {
          name: 'update_cells',
          description: 'Update specific cells in a sheet. Preserves all other data.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the XLSX file' },
              sheetName: { type: 'string', description: 'Sheet name' },
              updates: { type: 'array', description: 'Array of {cell: "A1", value: "new value"} or {row: 2, col: "Name", value: "new value"}', items: { type: 'object' } }
            },
            required: ['filePath', 'sheetName', 'updates']
          }
        },
        {
          name: 'read_csv',
          description: 'Read data from a CSV file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the CSV file' }
            },
            required: ['filePath']
          }
        },
        {
          name: 'write_csv',
          description: 'Write data to a CSV file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to save the CSV file' },
              data: { type: 'array', description: 'Array of objects to write' }
            },
            required: ['filePath', 'data']
          }
        },
        {
          name: 'filter_data',
          description: 'Filter data based on conditions (equals, contains, gt, lt, gte, lte, ne, in, regex)',
          inputSchema: {
            type: 'object',
            properties: {
              data: { type: 'array', description: 'Array of objects to filter' },
              conditions: { type: 'object', description: 'Filter conditions' }
            },
            required: ['data', 'conditions']
          }
        }
      ]
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      try {
        switch (name) {
          case 'read_xlsx': return await this.readXlsx(args.filePath, args.sheetName);
          case 'write_xlsx': return await this.writeXlsx(args.filePath, args.data, args.sheetName, args.mode);
          case 'list_sheets': return await this.listSheets(args.filePath);
          case 'add_sheet': return await this.addSheet(args.filePath, args.sheetName, args.data);
          case 'delete_sheet': return await this.deleteSheet(args.filePath, args.sheetName);
          case 'rename_sheet': return await this.renameSheet(args.filePath, args.oldName, args.newName);
          case 'update_cells': return await this.updateCells(args.filePath, args.sheetName, args.updates);
          case 'read_csv': return await this.readCsv(args.filePath);
          case 'write_csv': return await this.writeCsv(args.filePath, args.data);
          case 'filter_data': return await this.filterData(args.data, args.conditions);
          default: throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        return { content: [{ type: 'text', text: `Error: ${error.message}` }] };
      }
    });
  }

  // === XLSX Operations ===

  _loadWorkbook(filePath) {
    if (!fs.existsSync(filePath)) throw new Error(`File not found: ${filePath}`);
    return XLSX.readFile(filePath);
  }

  _saveWorkbook(workbook, filePath) {
    XLSX.writeFile(workbook, filePath);
  }

  async readXlsx(filePath, sheetName) {
    const workbook = this._loadWorkbook(filePath);

    if (sheetName === '*') {
      const allData = {};
      for (const name of workbook.SheetNames) {
        allData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
      }
      return {
        content: [
          { type: 'text', text: `Read ${workbook.SheetNames.length} sheets from ${filePath}: ${workbook.SheetNames.join(', ')}` },
          { type: 'text', text: JSON.stringify(allData, null, 2) }
        ]
      };
    }

    const targetSheet = sheetName || workbook.SheetNames[0];
    if (!workbook.Sheets[targetSheet]) throw new Error(`Sheet "${targetSheet}" not found. Available: ${workbook.SheetNames.join(', ')}`);

    const data = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet]);
    return {
      content: [
        { type: 'text', text: `Read ${data.length} rows from "${targetSheet}" in ${filePath}` },
        { type: 'text', text: JSON.stringify(data, null, 2) }
      ]
    };
  }

  async writeXlsx(filePath, data, sheetName = 'Sheet1', mode = 'create') {
    const newSheet = XLSX.utils.json_to_sheet(data);

    if (mode === 'create' || !fs.existsSync(filePath)) {
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, newSheet, sheetName);
      this._saveWorkbook(workbook, filePath);
      return { content: [{ type: 'text', text: `Created ${filePath} with ${data.length} rows in "${sheetName}"` }] };
    }

    const workbook = this._loadWorkbook(filePath);

    if (mode === 'append') {
      if (!workbook.Sheets[sheetName]) throw new Error(`Sheet "${sheetName}" not found for append. Available: ${workbook.SheetNames.join(', ')}`);
      const existing = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      const combined = [...existing, ...data];
      workbook.Sheets[sheetName] = XLSX.utils.json_to_sheet(combined);
      this._saveWorkbook(workbook, filePath);
      return { content: [{ type: 'text', text: `Appended ${data.length} rows to "${sheetName}" (${existing.length} → ${combined.length} total) in ${filePath}` }] };
    }

    if (mode === 'replace') {
      if (workbook.Sheets[sheetName]) {
        workbook.Sheets[sheetName] = newSheet;
      } else {
        XLSX.utils.book_append_sheet(workbook, newSheet, sheetName);
      }
      this._saveWorkbook(workbook, filePath);
      return { content: [{ type: 'text', text: `Replaced "${sheetName}" with ${data.length} rows in ${filePath}. Sheets: ${workbook.SheetNames.join(', ')}` }] };
    }

    throw new Error(`Invalid mode: ${mode}. Use: create, append, replace`);
  }

  async listSheets(filePath) {
    const workbook = this._loadWorkbook(filePath);
    const sheets = workbook.SheetNames.map(name => {
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
      const range = workbook.Sheets[name]['!ref'] || 'empty';
      return { name, rows: data.length, range };
    });
    return {
      content: [
        { type: 'text', text: `${filePath}: ${sheets.length} sheet(s)` },
        { type: 'text', text: JSON.stringify(sheets, null, 2) }
      ]
    };
  }

  async addSheet(filePath, sheetName, data) {
    const workbook = this._loadWorkbook(filePath);
    if (workbook.SheetNames.includes(sheetName)) throw new Error(`Sheet "${sheetName}" already exists. Use write_xlsx with mode "replace" to overwrite.`);
    const sheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    this._saveWorkbook(workbook, filePath);
    return { content: [{ type: 'text', text: `Added sheet "${sheetName}" with ${data.length} rows. Sheets: ${workbook.SheetNames.join(', ')}` }] };
  }

  async deleteSheet(filePath, sheetName) {
    const workbook = this._loadWorkbook(filePath);
    if (!workbook.SheetNames.includes(sheetName)) throw new Error(`Sheet "${sheetName}" not found. Available: ${workbook.SheetNames.join(', ')}`);
    if (workbook.SheetNames.length === 1) throw new Error('Cannot delete the last sheet.');
    const idx = workbook.SheetNames.indexOf(sheetName);
    workbook.SheetNames.splice(idx, 1);
    delete workbook.Sheets[sheetName];
    this._saveWorkbook(workbook, filePath);
    return { content: [{ type: 'text', text: `Deleted sheet "${sheetName}". Remaining: ${workbook.SheetNames.join(', ')}` }] };
  }

  async renameSheet(filePath, oldName, newName) {
    const workbook = this._loadWorkbook(filePath);
    if (!workbook.SheetNames.includes(oldName)) throw new Error(`Sheet "${oldName}" not found.`);
    if (workbook.SheetNames.includes(newName)) throw new Error(`Sheet "${newName}" already exists.`);
    const idx = workbook.SheetNames.indexOf(oldName);
    workbook.SheetNames[idx] = newName;
    workbook.Sheets[newName] = workbook.Sheets[oldName];
    delete workbook.Sheets[oldName];
    this._saveWorkbook(workbook, filePath);
    return { content: [{ type: 'text', text: `Renamed "${oldName}" → "${newName}". Sheets: ${workbook.SheetNames.join(', ')}` }] };
  }

  async updateCells(filePath, sheetName, updates) {
    const workbook = this._loadWorkbook(filePath);
    if (!workbook.Sheets[sheetName]) throw new Error(`Sheet "${sheetName}" not found.`);
    const sheet = workbook.Sheets[sheetName];
    let updated = 0;

    for (const update of updates) {
      if (update.cell) {
        // Direct cell reference: {cell: "A1", value: "hello"}
        sheet[update.cell] = { t: typeof update.value === 'number' ? 'n' : 's', v: update.value };
        updated++;
      } else if (update.row !== undefined && update.col) {
        // Row+column name: {row: 2, col: "Name", value: "John"}
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const headers = data[0] || [];
        const colIdx = headers.indexOf(update.col);
        if (colIdx === -1) throw new Error(`Column "${update.col}" not found. Available: ${headers.join(', ')}`);
        const cellRef = XLSX.utils.encode_cell({ r: update.row, c: colIdx });
        sheet[cellRef] = { t: typeof update.value === 'number' ? 'n' : 's', v: update.value };
        updated++;
      }
    }

    this._saveWorkbook(workbook, filePath);
    return { content: [{ type: 'text', text: `Updated ${updated} cell(s) in "${sheetName}"` }] };
  }

  // === CSV Operations ===

  async readCsv(filePath) {
    return new Promise((resolve, reject) => {
      const results = [];
      fs.createReadStream(filePath)
        .pipe(csv())
        .on('data', (data) => results.push(data))
        .on('end', () => {
          resolve({
            content: [
              { type: 'text', text: `Read ${results.length} rows from ${filePath}` },
              { type: 'text', text: JSON.stringify(results, null, 2) }
            ]
          });
        })
        .on('error', reject);
    });
  }

  async writeCsv(filePath, data) {
    if (!data.length) throw new Error('No data to write');
    const headers = Object.keys(data[0]).map(key => ({ id: key, title: key }));
    const csvWriter = createCsvWriter.createObjectCsvWriter({ path: filePath, header: headers });
    await csvWriter.writeRecords(data);
    return { content: [{ type: 'text', text: `Wrote ${data.length} rows to ${filePath}` }] };
  }

  // === Filter ===

  async filterData(data, conditions) {
    const filtered = data.filter(row => {
      return Object.entries(conditions).every(([key, value]) => {
        if (typeof value === 'object' && value !== null && value.operator) {
          const v = value.value;
          switch (value.operator) {
            case 'equals': return row[key] === v;
            case 'ne': return row[key] !== v;
            case 'contains': return String(row[key]).toLowerCase().includes(String(v).toLowerCase());
            case 'gt': return Number(row[key]) > Number(v);
            case 'lt': return Number(row[key]) < Number(v);
            case 'gte': return Number(row[key]) >= Number(v);
            case 'lte': return Number(row[key]) <= Number(v);
            case 'in': return Array.isArray(v) && v.includes(row[key]);
            case 'regex': return new RegExp(v).test(String(row[key]));
            default: return row[key] === v;
          }
        }
        return row[key] === value;
      });
    });
    return {
      content: [
        { type: 'text', text: `Filtered ${data.length} → ${filtered.length} rows` },
        { type: 'text', text: JSON.stringify(filtered, null, 2) }
      ]
    };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('XLSX/CSV CRUD MCP server v2.0 running on stdio');
  }
}

const server = new XlsxCsvCrudServer();
server.run().catch(console.error);
