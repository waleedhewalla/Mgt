"""Guide formatting after fmt5: long centred text in table bodies and in the running text goes right-aligned (bidi default);
optional page margins.  Usage: fmt_g.py in.docx out.docx [margin_lr margin_tb]"""
import sys, zipfile
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
q = lambda t: W + t
src, dst = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(src)
x = etree.fromstring(z.read('word/document.xml'))
body = x.find(q('body'))
if len(sys.argv) > 4:
    mar = body.find(q('sectPr')).find(q('pgMar'))
    for k in ('left', 'right'):
        mar.set(q(k), sys.argv[3])
    for k in ('top', 'bottom'):
        mar.set(q(k), sys.argv[4])
txt = lambda p: ''.join(t.text or '' for t in p.iter(q('t'))).strip()
n_cell = n_top = 0
for tbl in body.iter(q('tbl')):
    for r, tr in enumerate(tbl.findall(q('tr'))):
        if r == 0 and len(tbl.findall(q('tr'))) > 1:
            continue
        for p in tr.iter(q('p')):
            jc = p.find(q('pPr') + '/' + q('jc'))
            if jc is not None and jc.get(q('val')) == 'center' and len(txt(p)) > 30:
                jc.getparent().remove(jc)
                n_cell += 1
for p in body.findall(q('p')):
    jc = p.find(q('pPr') + '/' + q('jc'))
    t = txt(p)
    if jc is not None and jc.get(q('val')) == 'center' and len(t) > 120 and '﴿' not in t:
        jc.getparent().remove(jc)
        n_top += 1
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
                   if it.filename == 'word/document.xml' else z.read(it.filename))
print('cells right-aligned', n_cell, '| paragraphs right-aligned', n_top)
