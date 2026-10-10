"""Reorder pPr / rPr / sectPr children into schema order and fill settings zoom; in place on a .docx."""
import sys, zipfile, shutil, os
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
       'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
       'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
       'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
       'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
       'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
SECT = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
        'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings',
        'sectPrChange']
TBL = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing',
       'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
TC = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']
def reorder(el, order):
    kids = list(el)
    def key(c):
        n = etree.QName(c).localname
        return order.index(n) if n in order else len(order)
    kids.sort(key=key)
    for c in kids:
        el.remove(c)
    seen = set()
    for c in kids:
        n = etree.QName(c).localname
        if n in seen and n not in ('headerReference', 'footerReference'):
            continue
        seen.add(n)
        el.append(c)
src = sys.argv[1]
tmp = src + '.tmp'
zin = zipfile.ZipFile(src)
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename.startswith('word/') and it.filename.endswith('.xml'):
            x = etree.fromstring(data)
            for e in x.iter(W + 'pPr'):
                reorder(e, PPR)
            for e in x.iter(W + 'rPr'):
                reorder(e, RPR)
            for e in x.iter(W + 'tblPr'):
                reorder(e, TBL)
            for e in x.iter(W + 'tcPr'):
                reorder(e, TC)
            for e in x.iter(W + 'sectPr'):
                reorder(e, SECT)
            for z in x.iter(W + 'zoom'):
                if z.get(W + 'percent') is None:
                    z.set(W + 'percent', '100')
            data = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(it, data)
shutil.move(tmp, src)
print('fixed', src)
