"""Print layout for Book Two at about 250 pages: font sizes, line and paragraph spacing, margins, tables and figures.
Usage: layout.py in.docx out.docx"""
import sys, os, zipfile, re
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
q = lambda t: W + t
src, dst = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(src)
x = etree.fromstring(z.read('word/document.xml'))
body = x.find(q('body'))
sect = body.find(q('sectPr'))
mar = sect.find(q('pgMar'))
for k, v in (('top', 900), ('bottom', 900), ('right', 1000), ('left', 1000), ('header', 500), ('footer', 500)):
    mar.set(q(k), str(v))
# run sizes (half-points): body 14 -> 13, tables 12 -> 11, headings a step smaller
SIZE = {'28': '25', '24': '22', '26': '24', '30': '28', '34': '30', '36': '32', '40': '34', '44': '36', '48': '40', '22': '20'}
def in_tbl(e):
    return any(a.tag == q('tbl') for a in e.iterancestors())
TSIZE = {'24': '21', '22': '21', '26': '22', '28': '22', '20': '19'}
for r in body.iter(q('rPr')):
    m = TSIZE if in_tbl(r) else SIZE
    for tag in ('sz', 'szCs'):
        v = r.find(q(tag))
        if v is not None and v.get(q('val')) in m:
            v.set(q('val'), m[v.get(q('val'))])
# spacing: auto line spacing capped at 1.1, paragraph spacing reduced
for sp in body.iter(q('spacing')):
    if sp.getparent().tag != q('pPr'):
        continue
    line = sp.get(q('line'))
    if in_tbl(sp):
        sp.set(q('line'), '240')
        sp.set(q('lineRule'), 'auto')
    elif line and (sp.get(q('lineRule')) in (None, 'auto')) and int(line) > 240:
        sp.set(q('line'), '240')
    for a in ('before', 'after'):
        v = sp.get(q(a))
        if v:
            n = int(v)
            cell = in_tbl(sp)
            sp.set(q(a), str(0 if cell and n <= 40 else (min(n, 20) if cell else int(n * 0.55))))
# tables: tighter cell margins
for mar_ in body.iter(q('tcMar')):
    for side in mar_:
        if side.get(q('w')) and int(side.get(q('w'))) > 40 and etree.QName(side).localname in ('top', 'bottom'):
            side.set(q('w'), '30')
# figures: 75 % of their size
for ext in x.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    for k in ('cx', 'cy'):
        ext.set(k, str(int(int(ext.get(k)) * 0.65)))
for ext in x.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
    if ext.get('cx'):
        for k in ('cx', 'cy'):
            ext.set(k, str(int(int(ext.get(k)) * 0.65)))
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True) if it.filename == 'word/document.xml'
                   else z.read(it.filename))
print('wrote', dst)
