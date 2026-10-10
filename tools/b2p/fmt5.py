"""Alignment and table width for Book Two: Arabic paragraphs right-aligned the way Word writes it (bidi with no jc),
English references left-to-right, every table stretched to the full text width with fixed layout.
Usage: fmt5.py in.docx out.docx"""
import sys, zipfile, re
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
q = lambda t: W + t
src, dst = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(src)
x = etree.fromstring(z.read('word/document.xml'))
body = x.find(q('body'))
sect = body.find(q('sectPr'))
pg, mar = sect.find(q('pgSz')), sect.find(q('pgMar'))
TW = int(pg.get(q('w'))) - int(mar.get(q('left'))) - int(mar.get(q('right')))
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers',
             'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi']
latin = re.compile(r'^[\s\d.]*[A-Za-z]')
n_right = n_ltr = 0
for p in body.iter(q('p')):
    txt = ''.join(t.text or '' for t in p.iter(q('t')))
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.SubElement(p, q('pPr')); p.insert(0, ppr)
    jc = ppr.find(q('jc'))
    bidi = ppr.find(q('bidi'))
    ltr = jc is not None and jc.get(q('val')) == 'left' and latin.match(txt) and not re.search('[؀-ۿ]', txt)
    if ltr:
        if bidi is not None:
            ppr.remove(bidi)
        for r in p.iter(q('rtl')):
            r.getparent().remove(r)
        n_ltr += 1
        continue
    if bidi is None:
        bidi = etree.Element(q('bidi'))
        pos = 0
        for i, ch in enumerate(ppr):
            name = etree.QName(ch).localname
            if name in PPR_ORDER or name == 'rPr' and False:
                pos = i + 1
        ppr.insert(pos, bidi)
    if jc is not None and jc.get(q('val')) in ('right', 'both', 'start'):
        ppr.remove(jc)
        n_right += 1
TBL_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
             'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook']
def put(parent, tag, order, **attrs):
    el = parent.find(q(tag))
    if el is None:
        el = etree.Element(q(tag))
        idx = order.index(tag)
        pos = 0
        for i, ch in enumerate(parent):
            if etree.QName(ch).localname in order and order.index(etree.QName(ch).localname) < idx:
                pos = i + 1
        parent.insert(pos, el)
    for k, v in attrs.items():
        el.set(q(k), v)
    return el
n_tbl = 0
for t in body.iter(q('tbl')):
    grid = t.find(q('tblGrid'))
    cols = grid.findall(q('gridCol'))
    total = sum(int(c.get(q('w'))) for c in cols)
    if not total:
        continue
    f = TW / total
    acc = 0
    for i, c in enumerate(cols):
        w = TW - acc if i == len(cols) - 1 else round(int(c.get(q('w'))) * f)
        c.set(q('w'), str(w)); acc += w
    for tcw in t.iter(q('tcW')):
        if tcw.get(q('type')) == 'dxa':
            tcw.set(q('w'), str(round(int(tcw.get(q('w'))) * f)))
    tp = t.find(q('tblPr'))
    put(tp, 'tblW', TBL_ORDER, type='dxa', w=str(TW))
    jc = tp.find(q('jc'))
    if jc is not None:
        tp.remove(jc)
    put(tp, 'tblInd', TBL_ORDER, type='dxa', w='0')
    put(tp, 'tblLayout', TBL_ORDER, type='fixed')
    n_tbl += 1
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
                   if it.filename == 'word/document.xml' else z.read(it.filename))
print('text width', TW, '| right-aligned', n_right, '| LTR refs', n_ltr, '| tables', n_tbl)
