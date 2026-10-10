"""Rough page-count estimate for a docx: convert to HTML with the document's page size, margins, font sizes, spacing and
images, print with headless Chromium, count pages.  Usage: estimate.py file.docx [single_line_em]"""
import sys, os, re, zipfile, html, subprocess, shutil
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
src = sys.argv[1]
SINGLE = float(sys.argv[2]) if len(sys.argv) > 2 else 1.45      # single line height of the font, in em
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'est')
shutil.rmtree(out, ignore_errors=True)
os.makedirs(out)
z = zipfile.ZipFile(src)
x = etree.fromstring(z.read('word/document.xml'))
rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
rmap = {r.get('Id'): r.get('Target') for r in rels}
for n in z.namelist():
    if n.startswith('word/media/') and not n.endswith('/'):
        open(os.path.join(out, os.path.basename(n)), 'wb').write(z.read(n))
body = x.find(W + 'body')
sect = body.find(W + 'sectPr')
pg = sect.find(W + 'pgSz'); mar = sect.find(W + 'pgMar')
tw = lambda v: int(v) / 20                                   # twips -> pt
PW, PH = tw(pg.get(W + 'w')), tw(pg.get(W + 'h'))
MT, MB, MR, ML = (tw(mar.get(W + k)) for k in ('top', 'bottom', 'right', 'left'))


def rpr_size(r):
    v = r.find(W + 'rPr/' + W + 'szCs')
    if v is None:
        v = r.find(W + 'rPr/' + W + 'sz')
    return int(v.get(W + 'val')) / 2 if v is not None else 12


def para(p, in_cell=False):
    ppr = p.find(W + 'pPr')
    st = []
    sizes = []
    inner = []
    for r in p.iter(W + 'r'):
        for ch in r:
            if ch.tag == W + 't' and ch.text:
                b = r.find(W + 'rPr/' + W + 'b') is not None
                s = rpr_size(r)
                sizes.append(s)
                inner.append(f'<span style="font-size:{s}pt;{"font-weight:bold;" if b else ""}">{html.escape(ch.text)}</span>')
            elif ch.tag == W + 'br' and ch.get(W + 'type') != 'page':
                inner.append('<br/>')
            elif ch.tag == W + 'drawing':
                ext = ch.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
                blip = ch.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if ext is not None and blip is not None:
                    wpt, hpt = int(ext.get('cx')) / 12700, int(ext.get('cy')) / 12700
                    inner.append(f'<img src="{os.path.basename(rmap.get(blip.get(R + "embed"), ""))}" style="width:{wpt}pt;height:{hpt}pt"/>')
    size = max(sizes) if sizes else 12
    before = after = 0
    mult = 1.0
    exact = None
    if ppr is not None:
        sp = ppr.find(W + 'spacing')
        if sp is not None:
            before = int(sp.get(W + 'before') or 0) / 20
            after = int(sp.get(W + 'after') or 0) / 20
            if sp.get(W + 'line'):
                rule = sp.get(W + 'lineRule') or 'auto'
                if rule == 'auto':
                    mult = int(sp.get(W + 'line')) / 240
                else:
                    exact = int(sp.get(W + 'line')) / 20
        if ppr.find(W + 'pageBreakBefore') is not None:
            st.append('break-before:page')
        jc = ppr.find(W + 'jc')
        if jc is not None and jc.get(W + 'val') == 'center':
            st.append('text-align:center')
    lh = f'{exact}pt' if exact else f'{mult * SINGLE:.3f}'
    st.append(f'margin:{before}pt 0 {after}pt 0;line-height:{lh};font-size:{size}pt')
    if not inner:
        inner = ['&nbsp;']
    return f'<p style="{";".join(st)}">{"".join(inner)}</p>'


def table(t):
    rows = []
    for tr in t.findall(W + 'tr'):
        cells = []
        for tc in tr.findall(W + 'tc'):
            span = tc.find(W + 'tcPr/' + W + 'gridSpan')
            cs = f' colspan="{span.get(W + "val")}"' if span is not None else ''
            cells.append(f'<td{cs}>' + ''.join(para(p, True) for p in tc.findall(W + 'p')) + '</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')
    return '<table>' + ''.join(rows) + '</table>'


parts = []
for e in body:
    if e.tag == W + 'p':
        parts.append(para(e))
    elif e.tag == W + 'tbl':
        parts.append(table(e))
font = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'fonts', 'Amiri-1.000', 'Amiri-Regular.ttf'))
fontb = font.replace('Regular', 'Bold')
doc = f'''<!doctype html><html dir="rtl"><head><meta charset="utf-8"><style>
@font-face {{ font-family: T; src: url("file://{font}"); }}
@font-face {{ font-family: T; font-weight: bold; src: url("file://{fontb}"); }}
@page {{ size: {PW}pt {PH}pt; margin: {MT}pt {MR}pt {MB}pt {ML}pt; }}
body {{ font-family: T; margin: 0; text-align: justify; }}
p {{ orphans: 2; widows: 2; }}
table {{ border-collapse: collapse; width: 100%; margin: 4pt 0; }}
td {{ border: 0.75pt solid #888; padding: 2pt 4pt; vertical-align: top; }}
</style></head><body>{"".join(parts)}</body></html>'''
open(os.path.join(out, 'doc.html'), 'w', encoding='utf-8').write(doc)
pdf = os.path.join(out, 'doc.pdf')
subprocess.run(['/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
                '--headless', '--no-sandbox', '--disable-gpu', '--allow-file-access-from-files', f'--print-to-pdf={pdf}',
                '--no-pdf-header-footer', 'file://' + os.path.join(out, 'doc.html')], capture_output=True, timeout=600)
data = open(pdf, 'rb').read()
print('pages', len(re.findall(rb'/Type\s*/Page[^s]', data)))
