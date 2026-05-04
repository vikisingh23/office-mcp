#!/usr/bin/env node

const arg = process.argv[2];

if (arg === 'pptx' || arg === '--pptx') {
  await import('./src/pptx.js');
} else if (arg === 'docx' || arg === '--docx') {
  await import('./src/docx.js');
} else if (arg === 'xlsx' || arg === '--xlsx') {
  await import('./src/xlsx.js');
} else {
  // Default: run all three as separate info
  console.error('@neuraforge/office-mcp — 32 tools for PPTX, DOCX, XLSX, CSV');
  console.error('');
  console.error('Usage:');
  console.error('  npx @neuraforge/office-mcp pptx   # PowerPoint (13 tools)');
  console.error('  npx @neuraforge/office-mcp docx   # Word (9 tools)');
  console.error('  npx @neuraforge/office-mcp xlsx   # Excel/CSV (10 tools)');
  process.exit(1);
}
