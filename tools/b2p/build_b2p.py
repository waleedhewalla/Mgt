"""Condensed edition of Book Two.

Applies the condensation edits (edits/*.json) to the print edition, then:
renumbers sections and appendices, updates cross-references, rebuilds the TOC (TC entries + static result),
recomputes index locations, compacts the layout, and writes the companion guide with the material moved out.
"""
import sys, os, re, json, glob, copy, shutil, subprocess, zipfile, collections
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from docxedit import *

SRC = os.path.join(HERE, 'b2u.docx')
OUT = os.path.join(HERE, 'build')
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


TEMPL = {'body': next(e for i, e in sorted(A.items()) if e.tag == q('p') and i > 120 and _sz(e) == 28 and len(text(e)) > 150),
         'heading': next(e for i, e in sorted(A.items()) if e.tag == q('p') and re.match(r'^1\.1\s\s', text(e))),
         'subheading': next(e for i, e in sorted(A.items()) if e.tag == q('p') and i > 120 and _sz(e) == 30 and 10 < len(text(e)) < 60)}
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
for f in sorted(glob.glob(os.path.join(HERE, 'edits', '*.json'))):
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

# ------------------------------------------------------------------ 2. structure helpers
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر',
       'الثالث عشر', 'الرابع عشر']
ORD_RX = '|'.join(sorted(ORD, key=len, reverse=True))
CH_HEAD = re.compile(r'^الفصل (' + ORD_RX + r')$')
SEC_HEAD = re.compile(r'^(\d{1,2})\.(\d{1,2})(?:-[٠-٩\dء-ي]+)?(\s+)(.*)$')
APP_HEAD = re.compile(r'^الملحق (' + ORD_RX + r'): (.*)$')


def top_paras():
    return [e for e in body if e.tag == q('p')]


toc_title = next(e for e in body if e.tag == q('p') and text(e) == 'المحتويات')
# the old static TOC: from the paragraph holding the TOC field begin to the one holding its end
kids = [e for e in body if e.tag != q('sectPr')]
i0 = next(i for i, e in enumerate(kids) if any('TOC' in (t.text or '') for t in e.iter(q('instrText'))))
i1 = next(i for i in range(i0, len(kids)) if any(f.get(q('fldCharType')) == 'end' for f in kids[i].iter(q('fldChar'))))
OLD_TOC = kids[i0:i1 + 1]
TOC_L0 = copy.deepcopy(OLD_TOC[0])
TOC_L1 = copy.deepcopy(next(e for e in OLD_TOC if text(e).startswith('الفصل ')))
for t_ in (TOC_L0, TOC_L1):
    strip_fields(t_)
OLD_TOC_TEXTS = {text(e) for e in OLD_TOC}
for e in OLD_TOC:
    body.remove(e)

# ------------------------------------------------------------------ 3. section renumbering
chapter = None
SECMAP = {}                # (chapter, old_k) -> new_k
count = collections.Counter()
for p in top_paras():
    t = text(p)
    m = CH_HEAD.match(t)
    if m:
        chapter = ORD.index(m.group(1)) + 1
        continue
    if t.startswith('خاتمة الكتاب الأول'):
        chapter = None
    m = SEC_HEAD.match(t)
    if chapter and m and int(m.group(1)) == chapter and len(t) < 160:
        count[chapter] += 1
        new_k = count[chapter]
        SECMAP.setdefault((chapter, int(m.group(2))), new_k)
        if new_k != int(m.group(2)) or re.match(r'^\d+\.\d+-', t):
            smart_set(p, f'{chapter}.{new_k}  {m.group(4).strip()}')
# deleted sections map to the nearest earlier surviving section of the same chapter
ORIG_SECS = collections.defaultdict(list)
for i, e in sorted(A.items()):
    if e.tag == q('p'):
        m = SEC_HEAD.match(text(e))
        if m:
            ORIG_SECS[int(m.group(1))].append(int(m.group(2)))
for ch, ks in ORIG_SECS.items():
    last = 1
    for k in ks:
        if (ch, k) in SECMAP:
            last = SECMAP[(ch, k)]
        else:
            SECMAP[(ch, k)] = last
REF = re.compile(r'(?<![\d.])(1[0-4]|[1-9])\.(\d{1,2})(?![\d.]*\d)')


def remap_refs(s):
    def f(m):
        ch, k = int(m.group(1)), int(m.group(2))
        if (ch, k) in SECMAP:
            return f'{ch}.{SECMAP[(ch, k)]}'
        return m.group(0)
    return REF.sub(f, s)


for t in body.iter(q('t')):
    if t.text and '§' in t.text:
        t.text = re.sub(r'§\s*(?=\d)', '', t.text)
ref_changes = 0
for t in body.iter(q('t')):
    if t.text and REF.search(t.text):
        par = next(t.iterancestors(q('p')), None)
        pt = text(par) if par is not None else ''
        if SEC_HEAD.match(pt) and t is next(par.iter(q('t'))):
            continue             # headings already renumbered
        if re.fullmatch(r'\(?\d{1,2}\.\d{1,2}\)?', pt):
            continue             # an item number standing alone (survey/table numbering), not a section reference
        if len(REF.findall(pt)) >= 5:
            src = t.text
            def f2(m, src=src):
                pre = src[:m.start()]
                if pre.count('(') > pre.count(')'):
                    return m.group(0)        # numbered items inside parentheses: survey numbering
                ch, k = int(m.group(1)), int(m.group(2))
                return f'{ch}.{SECMAP[(ch, k)]}' if (ch, k) in SECMAP else m.group(0)
            new = REF.sub(f2, src)
        else:
            new = remap_refs(t.text)
        if new != t.text:
            ref_changes += 1
            t.text = new

# ------------------------------------------------------------------ 3b. figure and table captions renumbered per chapter
CAP = re.compile(r'^(شكل|جدول) (\d{1,2})\.(\d{1,2})(-[ء-ي])?')
CAPMAP, ccount, chapter = {}, collections.Counter(), None
for p in top_paras():
    t = text(p)
    m = CH_HEAD.match(t)
    if m:
        chapter = ORD.index(m.group(1)) + 1
        continue
    m = CAP.match(t)
    if chapter and m and int(m.group(2)) == chapter and len(t) < 160:
        kind, k = m.group(1), int(m.group(3))
        if (kind, chapter, k) not in CAPMAP:
            ccount[(kind, chapter)] += 1
            CAPMAP[(kind, chapter, k)] = ccount[(kind, chapter)]
        new_k = CAPMAP[(kind, chapter, k)]
        if new_k != k:
            ts = [x for x in p.iter(q('t')) if x.text]
            old = f'{kind} {chapter}.{k}'
            for x in ts:
                if old in x.text:
                    x.text = x.text.replace(old, f'{kind} {chapter}.{new_k}', 1)
                    break
CAPREF = re.compile(r'((?:ال)?(?:شكل|جدول)) (\d{1,2})\.(\d{1,2})(?![\d])')
cap_changes = 0
for p in body.iter(q('p')):
    if CAP.match(text(p)):
        continue
    for x in p.iter(q('t')):
        if x.text and CAPREF.search(x.text):
            def f(m):
                kind = m.group(1).replace('ال', '', 1) if m.group(1).startswith('ال') else m.group(1)
                key = (kind, int(m.group(2)), int(m.group(3)))
                return f'{m.group(1)} {m.group(2)}.{CAPMAP[key]}' if key in CAPMAP else m.group(0)
            new = CAPREF.sub(f, x.text)
            if new != x.text:
                cap_changes += 1
                x.text = new


# ------------------------------------------------------------------ 3c. figure captions «شكل N:» renumbered across the book
fig_n = 0
for p in top_paras():
    t = text(p)
    m = re.match(r'^شكل (\d+):', t)
    if m:
        fig_n += 1
        if int(m.group(1)) != fig_n:
            for x in p.iter(q('t')):
                if x.text and f'شكل {m.group(1)}:' in x.text:
                    x.text = x.text.replace(f'شكل {m.group(1)}:', f'شكل {fig_n}:', 1)
                    break

# ------------------------------------------------------------------ 4. appendix (أ) moved to the companion: references follow
APPA_REF = re.compile(r'(?:في )?(?:ال)?ملحق \(أ\)')
appa_changes = 0
for p in body.iter(q('p')):
    if text(p) == 'ملحق (أ)':
        continue
    for t in p.iter(q('t')):
        if t.text and APPA_REF.search(t.text):
            t.text = APPA_REF.sub(lambda m: ('في ' if m.group(0).startswith('في ') else '') + 'الدليل العملي المرافق', t.text)
            appa_changes += 1

# ------------------------------------------------------------------ 5. layout compaction
removed_empty = 0
for e in list(body):
    if e.tag != q('p') or text(e) or e.find('.//' + q('drawing')) is not None or e.find('.//' + q('br')) is not None:
        continue
    if e.find('.//' + q('instrText')) is not None or e.find('.//' + q('sectPr')) is not None:
        continue
    pr, nx = e.getprevious(), e.getnext()
    if pr is not None and nx is not None and pr.tag == q('tbl') and nx.tag == q('tbl'):
        continue
    if ppr(e).find(q('pageBreakBefore')) is not None:
        continue
    body.remove(e)
    removed_empty += 1
for tbl in body.iter(q('tbl')):
    for p in tbl.iter(q('p')):
        sp = p.find(q('pPr') + '/' + q('spacing'))
        if sp is not None:
            for a in ('before', 'after'):
                if sp.get(q(a)) and int(sp.get(q(a))) > 30:
                    sp.set(q(a), '20')

# ------------------------------------------------------------------ 6. index locations recomputed
from importlib.machinery import SourceFileLoader
AR_DIG = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
DIAC = re.compile(r'[ً-ْٰـ]')
idx_head = [e for e in body if e.tag == q('p') and text(e) == 'الفهارس'][-1]
UNITS, label, chapter, started = [], 'المقدمة', None, False
for e in body:
    if e is idx_head:
        break
    if e.tag == q('p'):
        t = text(e)
        if t == 'الإهداء':
            started = True
        if not started:
            continue
        m = CH_HEAD.match(t)
        if m:
            chapter = ORD.index(m.group(1)) + 1
            label = f'ف{chapter}'
        elif t.startswith('الخاتمة: المؤسسة القرآنية'):
            label, chapter = 'الخاتمة', None
        elif t == 'ملحق (ب)':
            label, chapter = 'ملحق (ب)', None
        else:
            ms = SEC_HEAD.match(t)
            if ms and chapter and int(ms.group(1)) == chapter and len(t) < 160:
                label = f'{ms.group(1)}.{ms.group(2)}'
    if e.tag in (q('p'), q('tbl')) and started:
        UNITS.append((label, ' '.join(ptext(x) for x in ([e] if e.tag == q('p') else e.iter(q('p'))))))
idx_tables = [e for e in idx_head.itersiblings() if e.tag == q('tbl')][:2]
verse_tbl, hadith_tbl = idx_tables
dropped = []


def set_cell(tc, s):
    ps = tc.findall(q('p'))
    set_text(ps[0], s)
    for extra in ps[1:]:
        tc.remove(extra)


for tr in verse_tbl.findall(q('tr'))[1:]:
    tcs = tr.findall(q('tc'))
    ref = ptext(tcs[0]).translate(AR_DIG).strip()
    m = re.match(r'^(.+?)\s*:\s*(\d+)', ref)
    if not m:
        continue
    sura, ayah = m.group(1).strip(), m.group(2)
    mr = re.match(r'^(.+?)\s*:\s*(\d+)\s*[-–]\s*(\d+)', ref)
    if mr:
        rx = re.compile(re.escape(sura) + r'\s*:\s*' + mr.group(2) + r'\s*[-–]\s*' + mr.group(3) + r'(?!\d)')
    else:
        rx = re.compile(re.escape(sura) + r'\s*:\s*(?:\d+\s*[-–،و]\s*)?' + ayah + r'(?!\d)')
    locs = []
    for lab, s in UNITS:
        if rx.search(s.translate(AR_DIG)) and lab not in locs:
            locs.append(lab)
    if not locs:
        verse_tbl.remove(tr)
        dropped.append(ref)
        continue
    set_cell(tcs[2], '، '.join(locs))
for tr in hadith_tbl.findall(q('tr'))[1:]:
    tcs = tr.findall(q('tc'))
    h = DIAC.sub('', ptext(tcs[0]))
    m = re.search(r'«([^»]+)»', h)
    if not m:
        continue
    words = [w for w in re.sub(r'[^\w\s]', ' ', m.group(1)).split() if len(w) > 1][:4]
    if len(words) < 2:
        continue
    rx = re.compile(r'\W+'.join(re.escape(w) for w in words))
    locs = []
    for lab, s in UNITS:
        if rx.search(DIAC.sub('', s)) and lab not in locs:
            locs.append(lab)
    if not locs:
        hadith_tbl.remove(tr)
        dropped.append('hadith: ' + ' '.join(words))
        continue
    set_cell(tcs[2], '، '.join(locs))

# ------------------------------------------------------------------ 7. section TC entries refreshed; static TOC rebuilt from TC levels 1–2


def tc_field(p, label, level):
    anchor = p.findall(q('r'))[-1] if p.findall(q('r')) else None
    for kind in ('begin', None, 'end'):
        r = etree.Element(q('r'))
        if kind:
            fc = etree.SubElement(r, q('fldChar'))
            fc.set(q('fldCharType'), kind)
        else:
            it = etree.SubElement(r, q('instrText'))
            it.set(XML_SPACE, 'preserve')
            it.text = f' TC "{label}" \\l {level} '
        if anchor is not None:
            anchor.addnext(r)
            anchor = r
        else:
            p.append(r)


def has_tc(p):
    return any((it.text or '').strip().startswith('TC ') for it in p.iter(q('instrText')))


chapter = None
for p in top_paras():
    t = text(p)
    m = CH_HEAD.match(t)
    if m:
        chapter = ORD.index(m.group(1)) + 1
    ms = SEC_HEAD.match(t)
    if chapter and ms and int(ms.group(1)) == chapter and len(t) < 160:
        if has_tc(p):
            strip_fields(p)
        tc_field(p, t, 3)
for e, lvl in ORIG_TC:
    if e.getparent() is body and not has_tc(e) and text(e):
        tc_field(e, text(e), lvl)
TC_LIST = []
for p in top_paras():
    for it in p.iter(q('instrText')):
        m = re.search(r'TC "([^"]+)" \\l (\d)', it.text or '')
        if m and m.group(2) in '12':
            TC_LIST.append((int(m.group(2)) - 1, m.group(1)))
entries = []
for lv, t in TC_LIST:
    p = fresh(copy.deepcopy(TOC_L0 if lv == 0 else TOC_L1))
    set_text(p, t)
    entries.append(p)


def fld_run(kind=None, instr=None):
    r = etree.Element(q('r'))
    if kind:
        f = etree.SubElement(r, q('fldChar'))
        f.set(q('fldCharType'), kind)
        if kind == 'begin':
            f.set(q('dirty'), 'true')
    else:
        it = etree.SubElement(r, q('instrText'))
        it.set(XML_SPACE, 'preserve')
        it.text = instr
    return r


first = entries[0]
k0 = 1 if first.find(q('pPr')) is not None else 0
first.insert(k0, fld_run('begin'))
first.insert(k0 + 1, fld_run(instr=' TOC \\f \\l "1-2" \\z '))
first.insert(k0 + 2, fld_run('separate'))
entries[-1].append(fld_run('end'))
prev = toc_title
for p in entries:
    prev.addnext(p)
    prev = p

for t in body.iter(q('t')):
    if t.text and t.text != t.text.strip():
        t.set(XML_SPACE, 'preserve')
seen = set()
W14N = '{http://schemas.microsoft.com/office/word/2010/wordml}'
for e in body.iter():
    for a in ('paraId', 'textId'):
        v = e.get(W14N + a)
        if v is not None:
            if (a, v) in seen:
                del e.attrib[W14N + a]
            else:
                seen.add((a, v))
tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
out = os.path.join(HERE, 'b2p_text.docx')
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-qXr', out, '.'], cwd=OUT, check=True)

words = sum(len(ptext(e).split()) for e in body.iter(q('p')))
json.dump({'stats': stats, 'log': LOG, 'ref_changes': ref_changes, 'cap_changes': cap_changes, 'appa_changes': appa_changes,
           'removed_empty': removed_empty, 'index_dropped': dropped, 'figures': fig_n,
           'companion_blocks': [(p, l, len(e)) for p, l, e in COMP_COPIES], 'words_after': words, 'toc_entries': len(TC_LIST)},
          open(os.path.join(HERE, 'build_log.json'), 'w'), ensure_ascii=False, indent=1)
print('wrote', out, '| words', words, '| skips', sum(v['skip'] for v in stats.values()), '| empty removed', removed_empty,
      '| refs', ref_changes, '| toc', len(TC_LIST), '| figs', fig_n)
