#!/usr/bin/env python3
"""
pptx_bridge.py — python-pptx backed operations for office-mcp.

Invoked by src/pptx.js as: python3 pptx_bridge.py <command> <json-args>
Always prints exactly one JSON object to stdout and exits 0 on success.
On failure, prints {"error": "..."} to stdout and exits 1 — the Node side
treats a non-JSON stdout or non-zero exit as a command failure either way,
but a clean JSON error is friendlier to surface to the calling agent.

Commands operate directly on the existing .pptx file via python-pptx, so
every slide/shape/table/image NOT targeted by the command is left byte-for-
byte as it was — this is what the JS-only (PptxGenJS) path could not do,
since PptxGenJS can only build a presentation from scratch.
"""
import sys
import json
import copy
import base64
import tempfile
import os

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
except ImportError as e:
    print(json.dumps({
        "error": f"Missing Python dependency: {e}. Run: pip install python-pptx",
    }))
    sys.exit(1)


def _hex_to_rgb(hex_color):
    hex_color = (hex_color or "FFFFFF").lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def _set_bullet(paragraph, char='•'):
    """Add a real XML bullet character to a paragraph (python-pptx has no high-level API for this)."""
    pPr = paragraph._p.get_or_add_pPr()
    buChar = pPr.makeelement(qn('a:buChar'), {'char': char})
    pPr.append(buChar)


def _render_content(slide, content, prs):
    """Shared renderer for add_rich_slide and replace_slide_content — mirrors the
    option keys applySlideContent() in pptx.js accepts (title, body, bullets,
    table, image, shapes, background, notes) so behavior is consistent whether a
    slide was built by PptxGenJS (create_presentation) or python-pptx (this file)."""
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    if content.get('title'):
        box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), slide_w - Inches(1), Inches(0.8))
        tf = box.text_frame
        tf.text = content['title']
        run = tf.paragraphs[0].runs[0]
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = _hex_to_rgb(content.get('titleColor', '2E2A94'))

    if content.get('body'):
        body = content['body']
        lines = body if isinstance(body, list) else [body]
        box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), slide_w - Inches(1), Inches(3.5))
        tf = box.text_frame
        tf.word_wrap = True
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line if isinstance(line, str) else line.get('text', '')
            for r in p.runs:
                r.font.size = Pt(16)

    if content.get('bullets'):
        box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), slide_w - Inches(1), Inches(3.5))
        tf = box.text_frame
        tf.word_wrap = True
        for i, b in enumerate(content['bullets']):
            text = b if isinstance(b, str) else b.get('text', '')
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = text
            p.level = (b.get('indent', 0) if isinstance(b, dict) else 0)
            _set_bullet(p)
            for r in p.runs:
                r.font.size = Pt(16)

    if content.get('table'):
        rows_data = content['table']
        n_rows = len(rows_data)
        n_cols = max(len(r) for r in rows_data) if rows_data else 0
        if n_rows and n_cols:
            gfx = slide.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(1.5), slide_w - Inches(1), Inches(0.4 * n_rows))
            table = gfx.table
            for ri, row in enumerate(rows_data):
                for ci, cell in enumerate(row):
                    text = cell if isinstance(cell, str) else cell.get('text', '')
                    table.cell(ri, ci).text = str(text)
                    if ri == 0:
                        for p in table.cell(ri, ci).text_frame.paragraphs:
                            for r in p.runs:
                                r.font.bold = True

    if content.get('image'):
        img = content['image']
        tmp_path = None
        try:
            if img.get('path'):
                src = img['path']
            elif img.get('data'):
                raw = img['data'].split(',', 1)[-1]  # tolerate data: URLs
                tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                tmp.write(base64.b64decode(raw))
                tmp.close()
                tmp_path = tmp.name
                src = tmp_path
            else:
                src = None
            if src:
                slide.shapes.add_picture(
                    src,
                    Inches(img.get('x', 1)), Inches(img.get('y', 1.5)),
                    Inches(img.get('w', 4)), Inches(img.get('h', 3)),
                )
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

    if content.get('shapes'):
        shape_map = {'rect': MSO_SHAPE.RECTANGLE, 'roundRect': MSO_SHAPE.ROUNDED_RECTANGLE, 'oval': MSO_SHAPE.OVAL}
        for s in content['shapes']:
            shp = slide.shapes.add_shape(
                shape_map.get(s.get('type'), MSO_SHAPE.RECTANGLE),
                Inches(s.get('x', 0)), Inches(s.get('y', 0)), Inches(s.get('w', 2)), Inches(s.get('h', 1)),
            )
            shp.fill.solid()
            shp.fill.fore_color.rgb = _hex_to_rgb(s.get('fill', 'EFEFEF'))
            if s.get('line'):
                shp.line.color.rgb = _hex_to_rgb(s['line'])
                shp.line.width = Pt(s.get('lineWidth', 1))

    if content.get('background'):
        bg = content['background']
        color = bg if isinstance(bg, str) else bg.get('fill', 'FFFFFF')
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _hex_to_rgb(color)

    if content.get('notes'):
        slide.notes_slide.notes_text_frame.text = content['notes']


def _clear_slide(slide):
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)


def _blank_layout(prs):
    # Prefer a layout with no placeholders ("Blank" is index 6 in the default
    # template) so _render_content fully controls what appears on the slide.
    for layout in prs.slide_layouts:
        if len(layout.placeholders) == 0:
            return layout
    return prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]


def cmd_read_structure(args):
    prs = Presentation(args['path'])
    slides_out = []
    for i, slide in enumerate(prs.slides):
        shapes_out = []
        for shape in slide.shapes:
            s = {'type': str(shape.shape_type), 'name': shape.name,
                 'left': shape.left, 'top': shape.top, 'width': shape.width, 'height': shape.height}
            if shape.has_text_frame:
                s['text'] = shape.text_frame.text
            if shape.has_table:
                t = shape.table
                s['table'] = [[c.text for c in row.cells] for row in t.rows]
            if shape.shape_type == 13:  # PICTURE
                s['isPicture'] = True
            shapes_out.append(s)
        slides_out.append({'slideNumber': i + 1, 'layout': slide.slide_layout.name, 'shapes': shapes_out})
    return {'slideCount': len(prs.slides), 'slides': slides_out}


def cmd_replace_text(args):
    prs = Presentation(args['path'])
    replacements = args['replacements']
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    for old, new in replacements.items():
                        if old in run.text:
                            run.text = run.text.replace(old, new)
                            count += 1
    prs.save(args['output'])
    return {'success': True, 'replacements': count}


def cmd_add_slide(args):
    """Simple title/body/notes slide using the deck's own layouts (original contract).

    Falls back to plain textboxes when the chosen layout has no title/body
    placeholder — true for any file this same package created via
    create_presentation (PptxGenJS doesn't define standard-idx placeholders),
    so relying on slide.shapes.title / placeholder idx 1 alone silently wrote
    nothing for that very common case."""
    prs = Presentation(args['path'])
    layout_idx = args.get('layoutIndex', 1)
    layout = prs.slide_layouts[min(layout_idx, len(prs.slide_layouts) - 1)]
    slide = prs.slides.add_slide(layout)
    content = args.get('content', {}) or {}

    if content.get('title'):
        if slide.shapes.title:
            slide.shapes.title.text = content['title']
        else:
            box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), prs.slide_width - Inches(1), Inches(0.8))
            box.text_frame.text = content['title']
            run = box.text_frame.paragraphs[0].runs[0]
            run.font.size = Pt(28)
            run.font.bold = True

    if content.get('body'):
        lines = content['body'] if isinstance(content['body'], list) else [content['body']]
        body_ph = next((p for p in slide.placeholders if p.placeholder_format.idx == 1), None)
        tf = body_ph.text_frame if body_ph else slide.shapes.add_textbox(
            Inches(0.5), Inches(1.5), prs.slide_width - Inches(1), Inches(3.5)).text_frame
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line if isinstance(line, str) else line.get('text', '')

    if content.get('notes'):
        slide.notes_slide.notes_text_frame.text = content['notes']

    prs.save(args['output'])
    return {'success': True, 'slideCount': len(prs.slides)}


def cmd_add_rich_slide(args):
    """Append one slide built from the full rich content schema (title, body,
    bullets, table, image, shapes, background, notes) — same schema create_presentation
    and modify_slide accept — without touching any existing slide."""
    prs = Presentation(args['path'])
    slide = prs.slides.add_slide(_blank_layout(prs))
    _clear_slide(slide)  # layout may still inject default placeholders
    _render_content(slide, args.get('content', {}) or {}, prs)
    prs.save(args['output'])
    return {'success': True, 'slideCount': len(prs.slides)}


def cmd_replace_slide_content(args):
    """Replace one existing slide's content in place (all other slides untouched)."""
    prs = Presentation(args['path'])
    idx = args['slideNumber'] - 1
    if idx < 0 or idx >= len(prs.slides):
        return {'error': f"Slide {args['slideNumber']} out of range (1-{len(prs.slides)})"}
    slide = prs.slides[idx]
    _clear_slide(slide)
    _render_content(slide, args.get('content', {}) or {}, prs)
    prs.save(args['output'])
    return {'success': True, 'modifiedSlide': args['slideNumber'], 'slideCount': len(prs.slides)}


def _delete_slide_by_index(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[idx])


def cmd_delete_slide(args):
    prs = Presentation(args['path'])
    idx = args['slideNumber'] - 1
    if idx < 0 or idx >= len(prs.slides):
        return {'error': f"Slide {args['slideNumber']} out of range (1-{len(prs.slides)})"}
    _delete_slide_by_index(prs, idx)
    prs.save(args['output'])
    return {'success': True, 'deletedSlide': args['slideNumber'], 'slideCount': len(prs.slides)}


def cmd_modify_slide_text(args):
    prs = Presentation(args['path'])
    idx = args['slideNumber'] - 1
    if idx < 0 or idx >= len(prs.slides):
        return {'error': f"Slide {args['slideNumber']} out of range (1-{len(prs.slides)})"}
    slide = prs.slides[idx]
    replacements = args['replacements']
    count = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                for old, new in replacements.items():
                    if old in run.text:
                        run.text = run.text.replace(old, new)
                        count += 1
    prs.save(args['output'])
    return {'success': True, 'replacements': count}


def cmd_update_table_cell(args):
    prs = Presentation(args['path'])
    idx = args['slideNumber'] - 1
    if idx < 0 or idx >= len(prs.slides):
        return {'error': f"Slide {args['slideNumber']} out of range (1-{len(prs.slides)})"}
    slide = prs.slides[idx]
    tables = [s for s in slide.shapes if s.has_table]
    table_idx = args.get('tableIndex', 0)
    if table_idx >= len(tables):
        return {'error': f"No table at index {table_idx} on slide {args['slideNumber']} ({len(tables)} table(s) found)"}
    table = tables[table_idx].table
    table.cell(args['row'], args['col']).text = str(args['value'])
    prs.save(args['output'])
    return {'success': True}


def cmd_duplicate_slide(args):
    prs = Presentation(args['path'])
    idx = args['slideNumber'] - 1
    if idx < 0 or idx >= len(prs.slides):
        return {'error': f"Slide {args['slideNumber']} out of range (1-{len(prs.slides)})"}
    source = prs.slides[idx]
    dest = prs.slides.add_slide(source.slide_layout)
    _clear_slide(dest)

    # Deep-copy every shape's XML, re-pointing any image relationships so
    # pictures on the duplicated slide still resolve.
    for shape in source.shapes:
        new_el = copy.deepcopy(shape._element)
        dest.shapes._spTree.append(new_el)
        for blip in new_el.findall('.//' + qn('a:blip')):
            r_embed = blip.get(qn('r:embed'))
            if r_embed and r_embed in source.part.rels:
                image_part = source.part.rels[r_embed].target_part
                new_rId = dest.part.relate_to(image_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
                blip.set(qn('r:embed'), new_rId)

    # Move the new slide to sit right after the source instead of at the end.
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    new_sld = slides[-1]
    xml_slides.remove(new_sld)
    xml_slides.insert(idx + 1, new_sld)

    prs.save(args['output'])
    return {'success': True, 'duplicatedFrom': args['slideNumber'], 'slideCount': len(prs.slides)}


COMMANDS = {
    'read_structure': cmd_read_structure,
    'replace_text': cmd_replace_text,
    'add_slide': cmd_add_slide,
    'add_rich_slide': cmd_add_rich_slide,
    'replace_slide_content': cmd_replace_slide_content,
    'delete_slide': cmd_delete_slide,
    'modify_slide_text': cmd_modify_slide_text,
    'update_table_cell': cmd_update_table_cell,
    'duplicate_slide': cmd_duplicate_slide,
}


def main():
    if len(sys.argv) < 3:
        print(json.dumps({'error': 'Usage: pptx_bridge.py <command> <json-args>'}))
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
