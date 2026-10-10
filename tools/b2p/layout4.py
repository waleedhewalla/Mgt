"""Readable print layout for Book Two: body 14 pt, tables 12 pt, uniform line spacing, moderate margins and figures.
Usage: layout4.py in.docx out.docx [body_line] [para_factor] [fig_scale] [margin_tb] [margin_lr] [tbl_size]"""
import sys, zipfile
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
q = lambda t: W + t
a = sys.argv
src, dst = a[1], a[2]
LINE = int(a[3]) if len(a) > 3 else 288
PF = float(a[4]) if len(a) > 4 else 0.7
FIG = float(a[5]) if len(a) > 5 else 0.8
MTB = int(a[6]) if len(a) > 6 else 1000
MLR = int(a[7]) if len(a) > 7 else 1080
TSZ = a[8] if len(a) > 8 else '24'
z = zipfile.ZipFile(src)
x = etree.fromstring(z.read('word/document.xml'))
body = x.find(q('body'))
mar = body.find(q('sectPr')).find(q('pgMar'))
for k, v in (('top', MTB), ('bottom', MTB), ('right', MLR), ('left', MLR), ('header', 560), ('footer', 560)):
    mar.set(q(k), str(v))
in_tbl = lambda e: any(p.tag == q('tbl') for p in e.iterancestors())
for r in body.iter(q('rPr')):
    if in_tbl(r):
        for tag in ('sz', 'szCs'):
            v = r.find(q(tag))
            if v is not None and int(v.get(q('val'))) >= 22:
                v.set(q('val'), TSZ)
for p in body.iter(q('p')):
    ppr = p.find(q('pPr'))
    if ppr is None:
        continue
    sp = ppr.find(q('spacing'))
    if sp is None:
        continue
    if in_tbl(p):
        sp.set(q('line'), '240'); sp.set(q('lineRule'), 'auto')
        sp.set(q('before'), '0'); sp.set(q('after'), '0')
        continue
    line = sp.get(q('line'))
    if line and sp.get(q('lineRule')) in (None, 'auto') and int(line) > LINE:
        sp.set(q('line'), str(LINE))
    for k in ('before', 'after'):
        v = sp.get(q(k))
        if v:
            sp.set(q(k), str(int(int(v) * PF)))
for m in body.iter(q('tcMar')):
    for side in m:
        if etree.QName(side).localname in ('top', 'bottom'):
            side.set(q('w'), '20')
for ext in x.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    for k in ('cx', 'cy'):
        ext.set(k, str(int(int(ext.get(k)) * FIG)))
for ext in x.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
    if ext.get('cx'):
        for k in ('cx', 'cy'):
            ext.set(k, str(int(int(ext.get(k)) * FIG)))
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
                   if it.filename == 'word/document.xml' else z.read(it.filename))
