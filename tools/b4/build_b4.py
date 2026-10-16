"""Book Four builder: markup sources (src/*.md) -> docx in the house layout of Book Three.

Markup (one element per line; blank lines ignored):
  @title  line1 | line2 | ...            title page lines (series, book no., title, subtitle, author)
  @heading TEXT                          front/back heading on a new page (TOC level 1)
  @part TEXT                             part title on a new page (TOC level 1)
  @chapter الفصل الأول | العنوان          chapter number + title (TOC level 2)
  @verse ﴿...﴾ | السورة: رقم              verse block + reference line
  ### 1.1  TEXT                          numbered section (TC level 3)
  #### TEXT                              sub-heading
  @fig {json}                            diagram spec for b2/diagrams.py (+ "caption")
  @table  ... rows "a | b | c" ... @end   first row = header
  @note TEXT                             small centred note (caption style)
  anything else                          body paragraph
"""
import sys, os, re, json, copy, shutil, subprocess, zipfile, glob
HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.dirname(HERE)
sys.path.insert(0, S)
sys.path.insert(0, os.path.join(S, 'b2'))
from lxml import etree
import diagrams
from PIL import Image

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def q(t): return f'{{{W}}}{t}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'

REF = os.path.join(HERE, 'b3ref.docx')
OUT = os.path.join(HERE, 'build')
DST = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'b4_text.docx')
ORDER = [os.path.join(HERE, 'src', f) for f in (sys.argv[2].split(',') if len(sys.argv) > 2 else
         sorted(os.listdir(os.path.join(HERE, 'src'))))]

shutil.rmtree(OUT, ignore_errors=True)
with zipfile.ZipFile(REF) as z:
    z.extractall(OUT)
tree = etree.parse(f'{OUT}/word/document.xml')
body = tree.getroot().find(q('body'))
sect = body.find(q('sectPr'))
kids = [e for e in body if e is not sect]


def text(e): return ''.join(t.text or '' for t in e.iter(q('t'))).strip()


def find(pred, start=0):
    return next(e for e in kids[start:] if pred(e))


def strip_fields(p):
    for r in list(p.findall(q('r'))):
        if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
            p.remove(r)


def tmpl(e):
    e = copy.deepcopy(e)
    strip_fields(e)
    for a in list(e.attrib):
        if a.startswith(W14):
            del e.attrib[a]
    return e


P = lambda e: e.tag == q('p')
T = {
    'title_series': tmpl(kids[0]), 'title_no': tmpl(kids[1]), 'title_main': tmpl(kids[2]), 'title_sub': tmpl(kids[3]),
    'title_author': tmpl(kids[4]), 'toc_title': kids[5],
    'heading': tmpl(find(lambda e: P(e) and text(e) == 'تقديم الكتاب الثالث', 30)),
    'verse': tmpl(find(lambda e: P(e) and text(e).startswith('﴿وَعَلَّمَ'), 30)),
    'vref': tmpl(find(lambda e: P(e) and text(e) == 'البقرة: ٣١', 30)),
    'body': tmpl(find(lambda e: P(e) and text(e).startswith('هذا ثالث كتب'), 30)),
    'part': tmpl(find(lambda e: P(e) and text(e).startswith('الباب التمهيدي:'), 40)),
    'chno': tmpl(find(lambda e: P(e) and text(e) == 'الفصل الأول', 40)),
    'chtitle': tmpl(find(lambda e: P(e) and text(e) == 'QMF في البيئة القطاعية', 40)),
    'section': tmpl(find(lambda e: P(e) and text(e).startswith('1.1  '), 40)),
    'sub': tmpl(find(lambda e: P(e) and text(e) == 'الألفاظ والصيغ', 40)),
    'caption': tmpl(find(lambda e: P(e) and text(e).startswith('شكل 1:'), 40)),
    'table': copy.deepcopy(find(lambda e: e.tag == q('tbl'), 40)),
}
TOC_L0 = tmpl(next(e for e in kids if P(e) and text(e).startswith('الباب التمهيدي') and 'TOC' not in ''.join(e.itertext())))
TOC_L1 = tmpl(next(e for e in kids[6:40] if P(e) and text(e).startswith('الفصل الأول')))
# drop everything but the title page anchor, the TOC heading and the section properties
for e in kids:
    body.remove(e)


def para(kind, s):
    p = copy.deepcopy(T[kind])
    runs = p.findall(q('r'))
    rpr = next((r.find(q('rPr')) for r in runs if r.find(q('t')) is not None), None)
    for r in runs:
        p.remove(r)
    r = etree.SubElement(p, q('r'))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, q('t'))
    t.text = s
    t.set(XS, 'preserve')
    return p


def tc(p, label, level):
    for kind in ('begin', None, 'end'):
        r = etree.SubElement(p, q('r'))
        if kind:
            etree.SubElement(r, q('fldChar')).set(q('fldCharType'), kind)
        else:
            it = etree.SubElement(r, q('instrText'))
            it.set(XS, 'preserve')
            it.text = f' TC "{label}" \\l {level} '


def table(rows):
    t = copy.deepcopy(T['table'])
    trs = t.findall(q('tr'))
    head, alt = trs[0], trs[1:3] if len(trs) > 2 else trs[1:2]
    for tr in trs:
        t.remove(tr)
    n = max(len(r) for r in rows)
    grid = t.find(q('tblGrid'))
    for g in list(grid):
        grid.remove(g)
    wtot = 9648
    widths = [wtot // n] * n
    for w in widths:
        etree.SubElement(grid, q('gridCol')).set(q('w'), str(w))
    for i, row in enumerate(rows):
        src = head if i == 0 else alt[(i - 1) % len(alt)]
        tr = copy.deepcopy(src)
        tcs = tr.findall(q('tc'))
        proto = tcs[0]
        for c in tcs:
            tr.remove(c)
        for j in range(n):
            c = copy.deepcopy(proto)
            c.find(q('tcPr')).find(q('tcW')).set(q('w'), str(widths[j]))
            p = c.find(q('p'))
            runs = p.findall(q('r'))
            rpr = runs[0].find(q('rPr')) if runs else None
            for r in runs:
                p.remove(r)
            for k, line in enumerate((row[j] if j < len(row) else '').split(' / ')):
                if k:
                    p = copy.deepcopy(p)
                    for r in p.findall(q('r')):
                        p.remove(r)
                    c.append(p)
                r = etree.SubElement(p, q('r'))
                if rpr is not None:
                    r.append(copy.deepcopy(rpr))
                tt = etree.SubElement(r, q('t'))
                tt.text = line.strip()
                tt.set(XS, 'preserve')
            tr.append(c)
        t.append(tr)
    return t


# ---------------------------------------------------------------- parse sources
OUTL, FIGS, TCS = [], [], []
chapter_n = 0
for path in ORDER:
    lines = open(path, encoding='utf-8').read().splitlines()
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        i += 1
        if not ln or ln.startswith('<!--'):
            continue
        if ln.startswith('@title'):
            parts = [x.strip() for x in ln[6:].split('|')]
            for k, s in zip(['title_series', 'title_no', 'title_main', 'title_sub', 'title_author'], parts):
                OUTL.append(para(k, s))
            OUTL.append(T['toc_title'])
        elif ln.startswith('@heading'):
            s = ln[8:].strip()
            p = para('heading', s); tc(p, s, 1); OUTL.append(p)
        elif ln.startswith('@part'):
            s = ln[5:].strip()
            p = para('part', s); tc(p, s, 1); OUTL.append(p)
        elif ln.startswith('@chapter'):
            no, ttl = [x.strip() for x in ln[8:].split('|', 1)]
            p = para('chno', no); tc(p, f'{no}: {ttl}', 2); OUTL.append(p)
            OUTL.append(para('chtitle', ttl))
        elif ln.startswith('@verse'):
            v, ref = [x.strip() for x in ln[6:].rsplit('|', 1)]
            OUTL.append(para('verse', v)); OUTL.append(para('vref', ref))
        elif ln.startswith('#### '):
            OUTL.append(para('sub', ln[5:].strip()))
        elif ln.startswith('### '):
            s = re.sub(r'^(\d+\.\d+)\s+', r'\1  ', ln[4:].strip())
            p = para('section', s); tc(p, s, 3); OUTL.append(p)
        elif ln.startswith('@fig'):
            spec = json.loads(ln[4:].strip())
            anchor = etree.Element(q('p'))
            OUTL.append(anchor)
            FIGS.append((anchor, spec))
        elif ln.startswith('@table'):
            rows = []
            while i < len(lines) and lines[i].strip() != '@end':
                if lines[i].strip():
                    rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')])
                i += 1
            i += 1
            OUTL.append(table(rows))
        elif ln.startswith('@note'):
            OUTL.append(para('caption', ln[5:].strip()))
        else:
            OUTL.append(para('body', ln))

for e in OUTL:
    sect.addprevious(e)

# ---------------------------------------------------------------- figures
media = f'{OUT}/word/media'
shutil.rmtree(media, ignore_errors=True)
os.makedirs(media)
rels_p = f'{OUT}/word/_rels/document.xml.rels'
rels = open(rels_p, encoding='utf-8').read()
rels = re.sub(r'<Relationship [^>]*Target="media/[^"]+"/>', '', rels)
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
FIGDIR = os.path.join(HERE, 'figs')
os.makedirs(FIGDIR, exist_ok=True)
new_rels = []
for n, (anchor, spec) in enumerate(FIGS, 1):
    path = os.path.join(FIGDIR, f'fig{n:02d}.png')
    diagrams.render(spec, path)
    shutil.copy(path, f'{media}/fig{n:02d}.png')
    w, h = Image.open(path).size
    cx = min(5900000, int(w / 200 * 914400)); cy = int(cx * h / w)
    if cy > 7600000:
        cy = 7600000; cx = int(cy * w / h)
    rid = f'rIdFig{n}'
    new_rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fig{n:02d}.png"/>')
    xml = (f'<w:p xmlns:w="{W}" xmlns:r="{R_NS}" xmlns:wp="{WP}" xmlns:a="{A_NS}" xmlns:pic="{PIC}">'
           '<w:pPr><w:keepNext/><w:bidi/><w:spacing w:before="160" w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
           f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{1000 + n}" name="fig{n}"/>'
           '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
           f'<pic:nvPicPr><pic:cNvPr id="{1000 + n}" name="fig{n}"/><pic:cNvPicPr/></pic:nvPicPr>'
           f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
           f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
           '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    ip = etree.fromstring(xml)
    anchor.addprevious(ip)
    cap = re.sub(r'^شكل[^:：]*[:：]\s*', '', spec.get('caption') or spec.get('title', ''))
    anchor.addprevious(para('caption', f'شكل {n}: {cap}'))
    body.remove(anchor)
rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
rels = re.sub(r'<Relationship [^>]*Target="comments.xml"/>', '', rels)
open(rels_p, 'w', encoding='utf-8').write(rels)
if os.path.exists(f'{OUT}/word/comments.xml'):
    os.remove(f'{OUT}/word/comments.xml')
ct_p = f'{OUT}/[Content_Types].xml'
ct = open(ct_p, encoding='utf-8').read()
ct = re.sub(r'<Override PartName="/word/comments.xml"[^>]*/>', '', ct)
if 'Extension="png"' not in ct:
    ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>')
open(ct_p, 'w', encoding='utf-8').write(ct)

# ---------------------------------------------------------------- static TOC (levels 1-2) with a live TOC field
toc_title = T['toc_title']
entries = []
for p in body.iter(q('p')):
    for it in p.iter(q('instrText')):
        m = re.search(r'TC "([^"]+)" \\l ([12])', it.text or '')
        if m:
            e = para('body', m.group(1))
            e = copy.deepcopy(TOC_L0 if m.group(2) == '1' else TOC_L1)
            for r in e.findall(q('r')):
                e.remove(r)
            r = etree.SubElement(e, q('r'))
            rp = (TOC_L0 if m.group(2) == '1' else TOC_L1).find('.//' + q('rPr'))
            if rp is not None:
                r.append(copy.deepcopy(rp))
            t = etree.SubElement(r, q('t')); t.text = m.group(1); t.set(XS, 'preserve')
            entries.append(e)


def fld(kind=None, instr=None):
    r = etree.Element(q('r'))
    if kind:
        f = etree.SubElement(r, q('fldChar')); f.set(q('fldCharType'), kind)
        if kind == 'begin':
            f.set(q('dirty'), 'true')
    else:
        it = etree.SubElement(r, q('instrText')); it.set(XS, 'preserve'); it.text = instr
    return r


if entries:
    first = entries[0]
    k0 = 1 if first.find(q('pPr')) is not None else 0
    first.insert(k0, fld('begin')); first.insert(k0 + 1, fld(instr=' TOC \\f \\l "1-2" \\z ')); first.insert(k0 + 2, fld('separate'))
    entries[-1].append(fld('end'))
    prev = toc_title
    for e in entries:
        prev.addnext(e); prev = e

# core properties
core_p = f'{OUT}/docProps/core.xml'
if os.path.exists(core_p):
    core = open(core_p, encoding='utf-8').read()
    core = re.sub(r'<dc:title>.*?</dc:title>', '<dc:title>موسوعة الإدارة القرآنية: الكتاب الرابع، الإدارة النبوية في السيرة</dc:title>', core)
    core = re.sub(r'<dc:subject>.*?</dc:subject>', '<dc:subject>الكتاب الرابع في سلسلة الإدارة القرآنية</dc:subject>', core)
    core = re.sub(r'<dc:description>.*?</dc:description>', '<dc:description></dc:description>', core)
    open(core_p, 'w', encoding='utf-8').write(core)

seen = set()
for e in body.iter():
    for a in ('paraId', 'textId'):
        v = e.get(W14 + a)
        if v is not None:
            if (a, v) in seen:
                del e.attrib[W14 + a]
            else:
                seen.add((a, v))
tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
if os.path.exists(DST):
    os.remove(DST)
subprocess.run(['zip', '-qXr', os.path.abspath(DST), '.'], cwd=OUT, check=True)
words = sum(len(text(p).split()) for p in body.iter(q('p')))
print('wrote', DST, '| words', words, '| figs', len(FIGS), '| toc', len(entries))
