#!/usr/bin/env python3
"""
docx_bridge.py — python-docx backed operations for office-mcp.

Invoked by src/docx.js as: python3 docx_bridge.py <command> <json-args>
Always prints exactly one JSON object to stdout. On failure, prints
{"error": "..."} and exits 1.

These operate directly on the existing .docx via python-docx, so content
NOT targeted by the command (formatting, images, other paragraphs/tables)
is left untouched — unlike the pure-JS tools in docx.js, which are explicit
in their own descriptions about only working with plain re-generated text.
"""
import sys
import json
import copy

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError as e:
    print(json.dumps({
        "error": f"Missing Python dependency: {e}. Run: pip install python-docx",
    }))
    sys.exit(1)


def cmd_read_structure(args):
    doc = Document(args['path'])
    paragraphs = [{'text': p.text, 'style': (p.style.name if p.style else None)} for p in doc.paragraphs]
    tables = [{'rows': [[c.text for c in row.cells] for row in t.rows]} for t in doc.tables]
    return {'paragraphCount': len(paragraphs), 'paragraphs': paragraphs,
            'tableCount': len(tables), 'tables': tables}


def cmd_replace_text(args):
    doc = Document(args['path'])
    replacements = args['replacements']
    count = 0

    def replace_in_paragraphs(paragraphs):
        nonlocal count
        for p in paragraphs:
            for run in p.runs:
                for old, new in replacements.items():
                    if old in run.text:
                        run.text = run.text.replace(old, new)
                        count += 1

    replace_in_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_paragraphs(cell.paragraphs)

    doc.save(args['output'])
    return {'success': True, 'replacements': count}


def _insert_paragraph_after(paragraph, text=None, style=None):
    """python-docx only exposes insert_paragraph_before(); build the after-variant
    by inserting a new <w:p> element directly after this paragraph's XML node."""
    new_p = copy.deepcopy(paragraph._p)
    for child in list(new_p):
        new_p.remove(child)
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    if style:
        try:
            new_para.style = style
        except KeyError:
            pass
    return new_para


def cmd_insert_after(args):
    doc = Document(args['path'])
    target = next((p for p in doc.paragraphs if args['afterText'] in p.text), None)
    if target is None:
        return {'error': f"Text not found: {args['afterText']}"}

    anchor = target
    for item in args['paragraphs']:
        text = item if isinstance(item, str) else item.get('text', '')
        heading = item.get('heading') if isinstance(item, dict) else None
        style = f'Heading {heading}' if heading else None
        anchor = _insert_paragraph_after(anchor, text, style)
        if isinstance(item, dict):
            run = anchor.runs[0] if anchor.runs else anchor.add_run('')
            run.bold = item.get('bold', False)
            run.italic = item.get('italic', False)

    doc.save(args['output'])
    return {'success': True, 'inserted': len(args['paragraphs'])}


def _set_table_borders(table, color='CCCCCC'):
    """Apply plain borders directly via OXML instead of a named table style —
    named styles ('Table Grid', 'Light Grid Accent 1', etc.) are only present
    if the document's styles.xml defines them, which a minimal docx (e.g. one
    built by this package's own JS create_document) typically does not."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), color)
        borders.append(el)
    tblPr.append(borders)


def cmd_append_table(args):
    doc = Document(args['path'])
    headers = args.get('headers', [])
    rows = args.get('rows', [])
    n_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    n_rows = len(rows) + (1 if headers else 0)
    if n_cols == 0 or n_rows == 0:
        return {'error': 'headers/rows produced an empty table'}

    table = doc.add_table(rows=n_rows, cols=n_cols)
    _set_table_borders(table)
    r = 0
    if headers:
        for c, h in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = str(h)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True
        r = 1
    for row in rows:
        for c, val in enumerate(row):
            table.cell(r, c).text = str(val)
        r += 1

    doc.save(args['output'])
    return {'success': True, 'rows': n_rows, 'cols': n_cols}


def cmd_delete_paragraph(args):
    doc = Document(args['path'])
    containing = args['containingText']
    to_remove = [p for p in doc.paragraphs if containing in p.text]
    for p in to_remove:
        p._element.getparent().remove(p._element)
    doc.save(args['output'])
    return {'success': True, 'deleted': len(to_remove)}


COMMANDS = {
    'read_structure': cmd_read_structure,
    'replace_text': cmd_replace_text,
    'insert_after': cmd_insert_after,
    'append_table': cmd_append_table,
    'delete_paragraph': cmd_delete_paragraph,
}


def main():
    if len(sys.argv) < 3:
        print(json.dumps({'error': 'Usage: docx_bridge.py <command> <json-args>'}))
        sys.exit(1)
    command, raw_args = sys.argv[1], sys.argv[2]
    if command not in COMMANDS:
        print(json.dumps({'error': f'Unknown command: {command}'}))
        sys.exit(1)
    try:
        args = json.loads(raw_args)
        result = COMMANDS[command](args)
        print(json.dumps(result))
    except Exception as e:
        print(json.dumps({'error': str(e), 'type': type(e).__name__}))
        sys.exit(1)


if __name__ == '__main__':
    main()
