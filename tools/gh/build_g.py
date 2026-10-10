"""Companion guides: apply humanization edits.
Usage: build_g.py src.docx out.docx part [part ...]

Applies the condensation edits (edits/*.json) to the print edition, then:
renumbers sections and appendices, updates cross-references, rebuilds the TOC (TC entries + static result),
recomputes index locations, compacts the layout, and writes the companion guide with the material moved out.
"""
import sys, os, re, json, glob, copy, shutil, subprocess, zipfile, collections
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from docxedit import *

SRC = os.path.abspath(sys.argv[1])
DST = os.path.abspath(sys.argv[2])
PARTS = sys.argv[3:]
OUT = os.path.join(HERE, 'build_' + os.path.basename(DST)[:-5])
shutil.rmtree(OUT, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(OUT)
tree = etree.parse(f'{OUT}/word/document.xml')
body = tree.getroot().find(q('body'))
sect = body.find(q('sectPr'))
A = {i: e for i, e in enumerate(e for e in body if e.tag != q('sectPr'))}
LOG = []

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
             'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
             'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
             'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']


def ensure(parent, tag, order):
    el = parent.find(q(tag))
    if el is not None:
        return el
    el = etree.Element(q(tag))
    pos = order.index(tag)
    for ch in parent:
        name = etree.QName(ch).localname
        if name in order and order.index(name) > pos:
            ch.addprevious(el)
            return el
    parent.append(el)
    return el


def ppr(p):
    el = p.find(q('pPr'))
    if el is None:
        el = etree.Element(q('pPr'))
        p.insert(0, el)
    return el


def text(e):
    return ptext(e).strip()


def nonempty_runs(p):
    return [r for r in p.findall(q('r')) if ''.join(t.text or '' for t in r.iter(q('t'))).strip()]


def is_bold(r):
    rp = r.find(q('rPr'))
    b = rp.find(q('b')) if rp is not None else None
    return b is not None and b.get(q('val')) not in ('false', '0')


def strip_fields(p):
    for r in list(p.findall(q('r'))):
        if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
            p.remove(r)


def smart_set(p, new):
    strip_fields(p)
    runs = nonempty_runs(p)
    if len(runs) >= 2:
        lab = ''.join(t.text or '' for t in runs[0].iter(q('t')))
        if lab.strip() and new.startswith(lab.strip()) and is_bold(runs[0]):
            rest = new[len(lab.strip()):]
            set_runs(p, [(lab if lab.endswith(' ') else lab.strip() + ' ', runs[0]), (rest.lstrip(), runs[1])])
            return
    if not p.findall(q('r')):
        p.append(etree.Element(q('r')))
    set_text(p, new)


def first_p(pred):
    return next(e for i, e in sorted(A.items()) if e.tag == q('p') and pred(text(e)))


def last_p(pred):
    return [e for i, e in sorted(A.items()) if e.tag == q('p') and pred(text(e))][-1]


def _sz(e):
    for r in e.findall(q('r')):
        v = r.find(q('rPr') + '/' + q('sz'))
        if v is not None and ''.join(t.text or '' for t in r.iter(q('t'))).strip():
            return int(v.get(q('val')))
    return 0


_tops = [e for i, e in sorted(A.items()) if e.tag == q('p')]
_body = next((e for e in _tops if len(text(e)) > 80), None)
if _body is None:
    _body = next(p for e in A.values() if e.tag == q('tbl') for p in e.iter(q('p')) if len(text(p)) > 40)
_head = next((e for e in _tops if text(e).startswith('مخرجات التعلم: الفصل')), _body)
_sub = next((e for e in _tops if re.match(r'^(تمرين|مرآة)', text(e))), _head)
TEMPL = {'body': _body, 'heading': _head, 'subheading': _sub}
TEMPL = {k: copy.deepcopy(v) for k, v in TEMPL.items()}
for v in TEMPL.values():
    strip_fields(v)


def idx_of(id_):
    return int(re.search(r'(\d+)', id_).group(1))


def resolve(id_):
    m = re.match(r'^TP(\d+)\.(\d+)\.(\d+)$', id_)
    if m:
        tbl = A.get(int(m.group(1)))
        if tbl is None or tbl.tag != q('tbl'):
            return None, None
        try:
            return tbl.findall(q('tr'))[int(m.group(2))].findall(q('tc'))[int(m.group(3))], 'cell'
        except IndexError:
            return None, None
    m = re.match(r'^(P|TBL)(\d+)$', id_)
    if m:
        e = A.get(int(m.group(2)))
        if e is None:
            return None, None
        return e, ('p' if e.tag == q('p') else 'tbl')
    return None, None


def walk(ea, eb):
    """Elements from ea to eb in current document order (handles earlier moves)."""
    if ea.getparent() is not eb.getparent():
        return None
    out, cur = [], ea
    while cur is not None and cur is not eb and len(out) < 4000:
        out.append(cur)
        cur = cur.getnext()
    return out + [eb] if cur is eb else None


# paragraphs carrying TOC entries (levels 1–2) keep them even when an edit rewrites their text
ORIG_TC = []
for i, e in sorted(A.items()):
    for it in e.iter(q('instrText')):
        m = re.search(r'TC "([^"]+)" \\l ([12])', it.text or '')
        if m:
            ORIG_TC.append((e, int(m.group(2))))

# ------------------------------------------------------------------ 1. edits
EDITS = {}
for f in [os.path.join(HERE, 'edits', p + '.json') for p in PARTS]:
    try:
        EDITS[os.path.basename(f)[:-5]] = json.load(open(f, encoding='utf-8'))
    except Exception as ex:
        LOG.append(f'BAD JSON {f}: {ex}')

stats = {}
COMPANION = []          # (label, [elements])
LASTINS = {}
ALL_DELETES = []
for part, data in EDITS.items():
    st = stats.setdefault(part, {'ok': 0, 'skip': 0})
    deletes, moves, comps = [], [], []
    for op in data.get('edits', []):
        kind = op.get('op')
        try:
            if kind == 'replace':
                el, typ = resolve(op['id'])
                if typ == 'p':
                    smart_set(el, op['text'])
                elif typ == 'cell':
                    lines = [l.strip() for l in op['text'].split(' / ')] if ' / ' in op['text'] else [op['text']]
                    ps = el.findall(q('p'))
                    if not ps:
                        raise ValueError('empty cell')
                    if not ps[0].findall(q('r')):
                        ps[0].append(etree.Element(q('r')))
                    set_text(ps[0], lines[0])
                    for extra in ps[1:]:
                        el.remove(extra)
                    prev = ps[0]
                    for ln in lines[1:]:
                        prev = clone_after(prev, ln)
                else:
                    raise ValueError('bad id')
            elif kind == 'insert_after':
                el, typ = resolve(op['id'])
                if typ != 'p':
                    raise ValueError('bad anchor')
                new = fresh(copy.deepcopy(TEMPL.get(op.get('style', 'body'), TEMPL['body'])))
                set_text(new, op['text'])
                el.addnext(new)
            elif kind == 'insert_block':
                anchor, typ = resolve(op['id'])
                if typ != 'p':
                    raise ValueError('bad anchor')
                prev = LASTINS.get(id(anchor), anchor)
                for like, txt in op['paras']:
                    tmpl, _ = resolve(like)
                    new = fresh(copy.deepcopy(tmpl))
                    strip_fields(new)
                    runs = nonempty_runs(new)
                    if len(runs) >= 2 and is_bold(runs[0]) and not is_bold(runs[1]) and ':' in txt[:40]:
                        k = txt.index(':') + 1
                        set_runs(new, [(txt[:k] + ' ', runs[0]), (txt[k:].lstrip(), runs[1])])
                    else:
                        set_text(new, txt)
                    prev.addnext(new)
                    prev = new
                LASTINS[id(anchor)] = prev
            elif kind == 'subst':
                hits = 0
                for t_ in body.iter(q('t')):
                    if t_.text and op['find'] in t_.text:
                        t_.text = t_.text.replace(op['find'], op['replace'])
                        hits += 1
                if hits != 1:
                    raise ValueError(f'subst hits {hits}')
            elif kind == 'delete':
                for i in op.get('ids', []):
                    el, typ = resolve(i)
                    if el is None or typ == 'cell':
                        raise ValueError('bad id ' + i)
                    deletes.append(el)
            elif kind == 'delete_range':
                a, b = idx_of(op['from']), idx_of(op['to'])
                deletes += [A[i] for i in range(a, b + 1) if i in A]
            elif kind == 'delete_rows':
                tbl, typ = resolve(op['id'])
                if typ != 'tbl':
                    raise ValueError('not a table')
                trs = tbl.findall(q('tr'))
                for r in sorted(set(op['rows'])):
                    if 0 <= r < len(trs):
                        deletes.append(trs[r])      # removed with the other deletes, so row ids stay stable
            elif kind == 'replace_table':
                tbl, typ = resolve(op['id'])
                if typ != 'tbl':
                    raise ValueError('not a table')
                prev = tbl
                for para in [s for s in re.split(r'\n\s*\n', op['text']) if s.strip()]:
                    new = fresh(copy.deepcopy(TEMPL['body']))
                    set_text(new, para.strip())
                    prev.addnext(new)
                    prev = new
                deletes.append(tbl)
            elif kind == 'move_range':
                a, b = idx_of(op['from']), idx_of(op['to'])
                anchor, typ = resolve(op['after'])
                if typ != 'p':
                    raise ValueError('bad anchor')
                moves.append(([A[i] for i in range(a, b + 1) if i in A], anchor, A.get(a), A.get(b)))
            elif kind == 'to_companion':
                a, b = idx_of(op['from']), idx_of(op['to'])
                comps.append((op.get('label', ''), A.get(a), A.get(b), [A[i] for i in range(a, b + 1) if i in A]))
            else:
                raise ValueError('unknown op')
            st['ok'] += 1
        except Exception as ex:
            st['skip'] += 1
            LOG.append(f'{part}: skip {kind} {op.get("id") or op.get("from")}: {ex}')
    for els, anchor, ea, eb in moves:
        w = walk(ea, eb) if ea is not None and eb is not None else None
        els = w or els
        if anchor in els:
            LOG.append(f'{part}: move anchor inside range, skipped')
            continue
        for e in reversed(els):
            anchor.addnext(e)
    for label, ea, eb, els in comps:
        w = walk(ea, eb) if ea is not None and eb is not None else None
        COMPANION.append((part, label, list(w or els)))
    ALL_DELETES += deletes

# copy companion material before deletion, then remove it and the deletes from the book
COMP_COPIES = [(part, re.sub(r'\s*\(الصفان[^)]*\)', '', label), [copy.deepcopy(e) for e in els]) for part, label, els in COMPANION]
for part, label, els in COMP_COPIES:
    for e in els:
        if e.tag == q('tbl'):
            trs = e.findall(q('tr'))
            if trs and ptext(trs[0]).strip().startswith('ما يثبته الفصل'):
                k = next((i for i, tr in enumerate(trs) if ptext(tr).strip().startswith('تمرين')), None)
                if k:
                    for tr in trs[:k]:
                        e.remove(tr)
for part, label, els in COMPANION:
    for e in els:
        if e.getparent() is not None:
            e.getparent().remove(e)
for el in ALL_DELETES:
    if el is not None and el.getparent() is not None:
        el.getparent().remove(el)
for t in body.iter(q('t')):
    if t.text and t.text != t.text.strip():
        t.set(XML_SPACE, 'preserve')
seen = set()
W14N = '{http://schemas.microsoft.com/office/word/2010/wordml}'
for e in body.iter():
    for a_ in ('paraId', 'textId'):
        v = e.get(W14N + a_)
        if v is not None:
            if (a_, v) in seen:
                del e.attrib[W14N + a_]
            else:
                seen.add((a_, v))
tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
if os.path.exists(DST):
    os.remove(DST)
subprocess.run(['zip', '-qXr', DST, '.'], cwd=OUT, check=True)
words = sum(len(ptext(e).split()) for e in body.iter(q('p')))
print('wrote', DST, '| words', words, '| stats', stats, '| log', LOG[:10])
