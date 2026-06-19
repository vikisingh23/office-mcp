#!/usr/bin/env python3
"""Bridge for office-pptx MCP — python-pptx formatting-preserving operations."""

import sys, json, copy
from pptx import Presentation
from pptx.util import Inches


def read_structure(args):
    prs = Presentation(args['path'])
    slides = []
    for i, slide in enumerate(prs.slides):
        sd = {'slideNumber': i + 1, 'shapes': []}
        for shape in slide.shapes:
            si = {
                'name': shape.name, 'type': str(shape.shape_type),
                'left': shape.left, 'top': shape.top,
                'width': shape.width, 'height': shape.height,
            }
            if shape.has_text_frame:
                si['text'] = shape.text_frame.text
                si['paragraphs'] = []
                for p in shape.text_frame.paragraphs:
                    para = {'text': p.text, 'level': p.level, 'runs': []}
                    for r in p.runs:
                        ri = {'text': r.text}
                        if r.font.name: ri['fontName'] = r.font.name
                        if r.font.size: ri['fontSize'] = r.font.size.pt
                        if r.font.bold is not None: ri['bold'] = r.font.bold
                        if r.font.italic is not None: ri['italic'] = r.font.italic
                        try:
                            if r.font.color and r.font.color.rgb: ri['color'] = str(r.font.color.rgb)
                        except Exception:
                            pass
                        para['runs'].append(ri)
                    si['paragraphs'].append(para)
            if shape.has_table:
                si['table'] = [[cell.text for cell in row.cells] for row in shape.table.rows]
            sd['shapes'].append(si)
        slides.append(sd)
    return {'filePath': args['path'], 'slideCount': len(prs.slides), 'slides': slides}


def _replace_in_shape(shape, replacements):
    count = 0
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for old, new in replacements.items():
                    if old in run.text:
                        run.text = run.text.replace(old, new)
                        count += 1
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        for old, new in replacements.items():
                            if old in run.text:
                                run.text = run.text.replace(old, new)
                                count += 1
    return count


def replace_text(args):
    prs = Presentation(args['path'])
    total = sum(_replace_in_shape(s, args['replacements']) for slide in prs.slides for s in slide.shapes)
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'replacements': total, 'output': out}


def add_slide(args):
    prs = Presentation(args['path'])
    layouts = prs.slide_layouts
    li = args.get('layoutIndex', 1)
    if li >= len(layouts): li = 0
    slide = prs.slides.add_slide(layouts[li])
    c = args.get('content', {})
    if c.get('title'):
        try:
            slide.shapes.title.text = c['title']
        except Exception:
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
            tb.text_frame.text = c['title']
    if c.get('body'):
        body = '\n'.join(str(b) for b in c['body']) if isinstance(c['body'], list) else c['body']
        placed = False
        for s in slide.placeholders:
            if s.placeholder_format.idx == 1:
                s.text_frame.text = body; placed = True; break
        if not placed:
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
            tb.text_frame.text = body
    if c.get('notes'):
        slide.notes_slide.notes_text_frame.text = c['notes']
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'slideCount': len(prs.slides), 'output': out}


def delete_slide(args):
    prs = Presentation(args['path'])
    n = args['slideNumber']
    if n < 1 or n > len(prs.slides):
        return {'error': f'Slide {n} out of range (1-{len(prs.slides)})'}
    rId = prs.slides._sldIdLst[n - 1].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[n - 1])
    for rel in list(prs.part.rels.values()):
        if rel.rId == rId:
            prs.part.drop_rel(rId); break
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'slideCount': len(prs.slides), 'output': out, 'deletedSlide': n}


def modify_slide_text(args):
    prs = Presentation(args['path'])
    n = args['slideNumber']
    if n < 1 or n > len(prs.slides):
        return {'error': f'Slide {n} out of range (1-{len(prs.slides)})'}
    total = sum(_replace_in_shape(s, args['replacements']) for s in prs.slides[n - 1].shapes)
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'replacements': total, 'output': out}


def update_table_cell(args):
    prs = Presentation(args['path'])
    n = args['slideNumber']
    if n < 1 or n > len(prs.slides):
        return {'error': f'Slide {n} out of range'}
    tables = [s for s in prs.slides[n - 1].shapes if s.has_table]
    ti = args.get('tableIndex', 0)
    if ti >= len(tables):
        return {'error': f'Table {ti} not found (has {len(tables)})'}
    t = tables[ti].table
    r, c = args['row'], args['col']
    if r >= len(t.rows) or c >= len(t.columns):
        return {'error': f'Cell r{r}c{c} out of range'}
    t.cell(r, c).text = args['value']
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'output': out, 'cell': f'r{r}c{c}', 'value': args['value']}


def duplicate_slide(args):
    prs = Presentation(args['path'])
    n = args['slideNumber']
    if n < 1 or n > len(prs.slides):
        return {'error': f'Slide {n} out of range'}
    src = prs.slides[n - 1]
    new_slide = prs.slides.add_slide(src.slide_layout)
    for shape in src.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shape._element))
    out = args.get('output', args['path'])
    prs.save(out)
    return {'success': True, 'slideCount': len(prs.slides), 'output': out, 'sourceSlide': n}


CMDS = {
    'read_structure': read_structure, 'replace_text': replace_text,
    'add_slide': add_slide, 'delete_slide': delete_slide,
    'modify_slide_text': modify_slide_text, 'update_table_cell': update_table_cell,
    'duplicate_slide': duplicate_slide,
}

if __name__ == '__main__':
    cmd, args = sys.argv[1], json.loads(sys.argv[2])
    if cmd not in CMDS:
        print(json.dumps({'error': f'Unknown: {cmd}'})); sys.exit(1)
    try:
        print(json.dumps(CMDS[cmd](args), default=str, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({'error': str(e)})); sys.exit(1)
