"""Book Three: reviewed full edition (mode full) and condensed print edition with companion guide (mode cond).

Applies the review edits (edits/*.json) — and in cond mode the condensation edits (edits_c/*.json) after them — to the
combined base, removes the material marked for mechanical removal, then normalises headings and formatting, sets the
tables right-to-left, inserts part headings, figures, indexes, the table of contents and page numbers.
"""
import sys, os, re, json, glob, copy, shutil, subprocess, zipfile, collections
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from docxedit import *

MODE = sys.argv[1] if len(sys.argv) > 1 else 'full'
SRC = os.path.join(HERE, 'base.docx')
OUT = os.path.join(HERE, f'build_{MODE}')
META = json.load(open(os.path.join(HERE, 'base_meta.json')))
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


def sz_of(e):
    for r in e.findall(q('r')):
        v = r.find(q('rPr') + '/' + q('sz'))
        if v is not None and ''.join(t.text or '' for t in r.iter(q('t'))).strip():
            return int(v.get(q('val')))
    return 0


def tmpl(pred):
    e = next(e for i, e in sorted(A.items()) if e.tag == q('p') and pred(e))
    c = copy.deepcopy(e)
    strip_fields(c)
    return c


TEMPL = {'body': tmpl(lambda e: sz_of(e) == 28 and len(text(e)) > 120),
         'heading': tmpl(lambda e: text(e).startswith('القسم الثاني: الفجوة الميدانية')),
         'subheading': tmpl(lambda e: sz_of(e) == 30 and text(e).startswith('٢/١'))}
T_CHAP = tmpl(lambda e: text(e) == 'الفصل الثاني' and sz_of(e) >= 44)
T_CHSUB = tmpl(lambda e: text(e).startswith('مقياس النضج القطاعي القرآني'))
T_PART = tmpl(lambda e: sz_of(e) == 40)
T_TITLE = [copy.deepcopy(A[i]) for i in range(0, 4)]
T_CAPTION = tmpl(lambda e: text(e).startswith('سورة البقرة: الآية 148'))
T_TABLE3 = copy.deepcopy(next(e for i, e in sorted(A.items()) if e.tag == q('tbl') and len(e.find(q('tr')).findall(q('tc'))) == 3))
MECH = [A[i] for i in META['mech'] if i in A]
CH1_TABLES = [A[i] for i in range(dict(META['marks'])['ch01'], dict(META['marks'])['ch02']) if A[i].tag == q('tbl')]


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


# ------------------------------------------------------------------ 1. edits
EDITS = collections.OrderedDict()
for d in (['edits'] + (['edits_c'] if MODE == 'cond' else [])):
    for f in sorted(glob.glob(os.path.join(HERE, d, '*.json'))):
        try:
            EDITS[('c:' if d == 'edits_c' else 'r:') + os.path.basename(f)[:-5]] = json.load(open(f, encoding='utf-8'))
        except Exception as ex:
            LOG.append(f'BAD JSON {f}: {ex}')
DIAG_SPECS = []

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
                prev_ = LASTINS.get(id(el), el)
                prev_.addnext(new)
                LASTINS[id(el)] = new                 # several inserts after one anchor keep their order
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
    if part.startswith('r:'):
        for dg in data.get('diagrams', []):
            el, typ = resolve(dg.get('after', ''))
            if typ == 'p':
                DIAG_SPECS.append((el, dg, part))
            else:
                LOG.append(f'{part}: diagram anchor bad {dg.get("after")}')
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

for el in MECH:
    if el.getparent() is not None:
        el.getparent().remove(el)

# ================================================================== 2. structure
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر',
       'الثالث عشر', 'الرابع عشر']
ORD_RX = '|'.join(sorted(ORD, key=len, reverse=True))
CH_HEAD = re.compile(r'^الفصل (' + ORD_RX + r')$')
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing',
               'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']


def top_paras():
    return [e for e in body if e.tag == q('p')]


def new_p(tmpl_, s):
    p = fresh(copy.deepcopy(tmpl_))
    strip_fields(p)
    set_text(p, s)
    return p


# chapter one has no heading in the source
first_ch1 = next(e for e in body if e.tag == q('p') and text(e).startswith(('وَلِكُلٍّ وِجْهَةٌ', 'القسم الأول: الآية المحورية')))
h1 = new_p(T_CHAP, 'الفصل الأول')
first_ch1.addprevious(h1)
h1.addnext(new_p(T_CHSUB, 'QMF في البيئة القطاعية'))

# part headings
PARTS = {'الأول': 'الباب التمهيدي: من المؤسسي إلى القطاعي', 'الثالث': 'الباب الأول: المؤسسة التعليمية',
         'السادس': 'الباب الثاني: المؤسسة الصحية', 'التاسع': 'الباب الثالث: المؤسسة الإنتاجية والمالية',
         'الثاني عشر': 'الباب الرابع: الحوكمة العامة وتقويم المرحلة'}
for p in top_paras():
    m = CH_HEAD.match(text(p))
    if m and m.group(1) in PARTS:
        p.addprevious(new_p(T_PART, PARTS[m.group(1)]))

# title page
first = next(e for e in body if e.tag == q('p') and text(e))
for tm, s in zip(T_TITLE + [T_TITLE[3]], ['موسوعة الإدارة القرآنية', 'الكتاب الثالث', 'التطبيقات القطاعية المتخصصة',
                                          'التعليم والصحة والصناعة والمالية والحوكمة العامة', 'م. وليد عبدالله حواله']):
    first.addprevious(new_p(tm, s))

# ================================================================== 3. headings: sections numbered per chapter
ORDW = r'(?:أول|ثان|ثالث|رابع|خامس|سادس|سابع|ثامن|تاسع|عاشر)(?:اً|ًا|ا|ي)?(?:\s+عشر)?|حادي\s+عشر|ثاني\s+عشر|ثالث\s+عشر|رابع\s+عشر|خامس\s+عشر|سادس\s+عشر'
PREFIX = re.compile(r'^\s*(?:القسم\s+(?:' + ORD_RX + r')\s*[:：]\s*|(?:' + ORDW + r')\s*[:：.]\s*|[\d٠-٩]+\s*[.\-—–]\s*|[\d٠-٩]+\s+)')
SUBPFX = re.compile(r'^\s*[\d٠-٩]+\s*/\s*[\d٠-٩]+\s*')
# a chapter's section level: size-34 headings, or size-30 where the chapter has (almost) no size-34 headings;
# size-30 headings before the first size-34 heading are sections too
n34 = collections.Counter()
chapter, after_head = None, 0
for p in top_paras():
    t = text(p)
    if not t:
        continue
    m = CH_HEAD.match(t)
    if m and sz_of(p) >= 44:
        chapter, after_head = ORD.index(m.group(1)) + 1, 1
        continue
    if chapter and sz_of(p) == 34:
        if after_head:
            after_head = 0
            continue
        n34[chapter] += 1
    elif sz_of(p) < 34 and after_head and sz_of(p) < 30:
        after_head = 0
chapter, sec_n, after_head, seen34 = None, 0, 0, False
n_sections = collections.Counter()


def number(p, t):
    global sec_n
    title = t
    for _ in range(3):
        title = PREFIX.sub('', title).strip()
    sec_n += 1
    n_sections[chapter] += 1
    smart_set(p, f'{chapter}.{sec_n}  {title}')


for p in top_paras():
    t = text(p)
    if not t:
        continue
    s = sz_of(p)
    m = CH_HEAD.match(t)
    if m and s >= 44:
        chapter, sec_n, after_head, seen34 = ORD.index(m.group(1)) + 1, 0, 1, False
        continue
    if s >= 34 and t.startswith(('الخاتمة', 'الملحق')) or (s >= 40 and not CH_HEAD.match(t)):
        chapter = None
        continue
    if not chapter:
        continue
    if s == 34:
        if after_head:
            after_head = 0
            continue
        seen34 = True
        number(p, t)
    elif s == 30 and '=' not in t:
        after_head = 0
        if n34[chapter] < 2 or not seen34:
            number(p, t)
        elif SUBPFX.match(t):
            smart_set(p, SUBPFX.sub('', t))
    elif after_head and s < 30:
        after_head = 0
SEC_HEAD = re.compile(r'^(\d{1,2})\.(\d{1,2})\s\s')

# ================================================================== 4. clean-up and style
MARK = re.compile(r'^[\s●◆◈▌■□▪►•✦✓✔☑]+')
EMOJI = re.compile('[\U0001F000-\U0001FFFF☀-⛿✀-➿️]')
for p in body.iter(q('p')):
    ts = [t for t in p.iter(q('t')) if t.text]
    if ts and MARK.match(ts[0].text):
        ts[0].text = MARK.sub('', ts[0].text)
AI_PHRASES = [(r'\s*[—\-–]\s*من منهج مهارة لسان', ''), (r'\s*من منهج مهارة لسان', ''), (r'استنادًا إلى مهارة mgt،?\s*', ''),
              (r'استناداً إلى مهارة mgt،?\s*', ''), (r'وليد هوالة|وليد حوالة', 'وليد حواله')]
for t in body.iter(q('t')):
    if t.text and re.search('[◈►●◆▌]', t.text):
        t.text = re.sub(r'[\u200e\u200f]*[◈►●◆▌]\s*', '', t.text)
for t in body.iter(q('t')):
    if t.text:
        s_ = EMOJI.sub(lambda m: m.group(0) if m.group(0) in '★✓' else '', t.text)
        for a, b in AI_PHRASES:
            s_ = re.sub(a, b, s_)
        t.text = s_
left_ai = [text(p)[:80] for p in body.iter(q('p')) if re.search(r'مهارة (لسان|mgt)|WRKOH|لجنة الخبراء', text(p))]
if left_ai:
    LOG.append('AI/committee mentions left: ' + ' | '.join(left_ai[:10]))
CAPTION = re.compile(r'^(جدول|شكل)\s')
NOT_BUT = re.compile(r'((?:^|[\s(«])(?:و|ف)?(?:ليس|ليست|ليسوا)\s[^.؟!؛:]{1,80}?)\s*[،:؛]?\s+بل\s')
dash_n = nb_n = 0


def fix_sentence_dashes(s, sep_single):
    out = []
    for sent in re.split(r'(?<=[.؟!])\s', s):
        n = sent.count('—')
        if n >= 2:
            sent = re.sub(r'\s*—\s*', '، ', sent)
        elif n == 1:
            sent = re.sub(r'\s*—\s*', sep_single, sent)
        out.append(sent)
    return ' '.join(out)


for p in body.iter(q('p')):
    full = text(p)
    if '—' not in full and 'بل' not in full:
        continue
    ts = [t for t in p.iter(q('t')) if t.text]
    sep = ': ' if len(full) < 70 else '؛ '
    for n, t in enumerate(ts):
        s0 = t.text
        s_ = s0
        if '—' in s_ and full.strip('— ') :
            s_ = fix_sentence_dashes(s_, sep)
            if s_.startswith((':', '؛', '،')) and n > 0:
                ts[n - 1].text = ts[n - 1].text.rstrip()
            if s_.endswith(' ') and n + 1 < len(ts):
                ts[n + 1].text = ts[n + 1].text.lstrip()
        s2 = NOT_BUT.sub(r'\1، وإنما ', s_)
        nb_n += s2 != s_
        dash_n += s0.count('—') - s2.count('—')
        t.text = s2.replace('؛ ،', '؛').replace('، ،', '،').replace('  ', ' ') if not SEC_HEAD.match(s2) else s2
for p in list(body.iter(q('p'))):                 # end-of-chapter markers such as «— نهاية الفصل الثاني —»
    if re.match(r'^[\s:؛،]*نهاية الفصل', text(p)):
        p.getparent().remove(p)

# body text regular weight: the sources set every paragraph bold
QURAN = re.compile('[ً-ْ]{6,}|﴿')
unbold = 0
for p in top_paras():
    t = text(p)
    if sz_of(p) not in (28, 24) or len(t) < 60 or QURAN.search(t[:80]):
        continue
    runs = nonempty_runs(p)
    keep_first = len(runs) > 1 and len(''.join(x.text or '' for x in runs[0].iter(q('t')))) < 45 and \
        ''.join(x.text or '' for x in runs[0].iter(q('t'))).rstrip().endswith((':', '：', '؛'))
    for k, r in enumerate(runs):
        if k == 0 and keep_first:
            continue
        rp = r.find(q('rPr'))
        if rp is not None:
            for tag in ('b', 'bCs'):
                el = rp.find(q(tag))
                if el is not None:
                    rp.remove(el)
                    unbold += 1
for tbl in body.iter(q('tbl')):
    for tr in tbl.findall(q('tr'))[1:]:
        for tc in tr.findall(q('tc')):
            if len(ptext(tc)) > 40:
                for rp in tc.iter(q('rPr')):
                    for tag in ('b', 'bCs'):
                        el = rp.find(q(tag))
                        if el is not None:
                            rp.remove(el)

# ================================================================== 5. tables right-to-left
KEY_HEADS = {'#', 'م', 'رقم', 'رمز', 'الرمز', 'رقم CLO', 'رمز CLO', 'CLO', 'المرحلة', 'المستوى', 'الفصل', 'الباب', 'المكوِّن', 'المكون',
             'مكوِّن QMF', 'مكون QMF', 'المعيار', 'المحك', 'المحور', 'الأداة', 'اسم الأداة', 'البُعد', 'البعد', 'الركن', 'المفهوم', 'المصطلح',
             'الحالة', 'الخطوة', 'الطبقة', 'السبب', 'النمط', 'الفئة', 'العنصر', 'الكتاب', 'الآية', 'الحديث', 'النموذج', 'الإطار', 'المؤشر',
             'السؤال', 'القرار', 'المدرسة', 'البند', 'الجانب', 'الوظيفة', 'الدور', 'النوع', 'الشهر', 'الوحدة', 'الفترة', 'السنة', 'المجال',
             'القطاع', 'الطرف', 'الفجوة', 'التحدي', 'الدرجة', 'الجذر', 'العضو', 'المقصد', 'الأصل', 'التوتر', 'الاعتراض', 'الميثاق', 'المبدأ'}
nh = lambda s: re.sub(r'[ً-ْـ]', '', s).strip()
KEYN = {nh(k) for k in KEY_HEADS}


def is_key(s):
    s = nh(s)
    return s in KEYN or any(s.startswith(k + ' ') for k in KEYN if len(k) > 2) or s.startswith(('رمز', 'رقم'))


def codeish(s):
    s = s.strip()
    if not s:
        return False
    return bool(re.match(r'^(\d+|[٠-٩]+|CLO\s?[\d.]+|[أ-ي]-[\d٠-٩]+|ف[\d٠-٩]+|[①-⑳]|المستوى\s|المرحلة\s)', s)) or len(s) <= 3


def reversed_score(tbl):
    rows = [[ptext(tc).strip() for tc in tr.findall(q('tc'))] for tr in tbl.findall(q('tr'))]
    rows = [r for r in rows if len(r) >= 2]
    if not rows:
        return 0
    hdr = rows[0]
    score = 0
    kl, kf = is_key(hdr[-1]), is_key(hdr[0])
    if not hdr[-1].strip() and hdr[0].strip():
        score += 3
    if not hdr[0].strip() and hdr[-1].strip():
        score -= 3
    if kl and not kf:
        score += 3
    if kf and not kl:
        score -= 3
    br = rows[1:] or rows
    lc, fc = sum(codeish(r[-1]) for r in br), sum(codeish(r[0]) for r in br)
    if lc > fc and lc >= len(br) * 0.6:
        score += 2
    if fc > lc and fc >= len(br) * 0.6:
        score -= 2
    return score


rev_n, amb_list = 0, []
CH1_IDS = {id(e) for e in CH1_TABLES}
for tbl in body.iter(q('tbl')):
    trs = tbl.findall(q('tr'))
    ncols = max(len(tr.findall(q('tc'))) for tr in trs)
    sc = reversed_score(tbl) if ncols >= 2 else 0
    if ncols >= 2 and (sc >= 2 or id(tbl) in CH1_IDS):
        for tr in trs:
            tcs = tr.findall(q('tc'))
            for tc in tcs:
                tr.remove(tc)
            for tc in reversed(tcs):
                tr.append(tc)
        grid = tbl.find(q('tblGrid'))
        if grid is not None:
            cols = grid.findall(q('gridCol'))
            for c in cols:
                grid.remove(c)
            for c in reversed(cols):
                grid.append(c)
        rev_n += 1
    elif ncols >= 2 and sc > -2:
        amb_list.append(ptext(trs[0])[:60])
    tp = tbl.find(q('tblPr'))
    if tp is None:
        tp = etree.Element(q('tblPr'))
        tbl.insert(0, tp)
    ensure(tp, 'bidiVisual', TBLPR_ORDER)
    for p in tbl.iter(q('p')):
        if re.search('[؀-ۿ]', ptext(p)):
            ensure(ppr(p), 'bidi', PPR_ORDER)
            jc = ppr(p).find(q('jc'))
            if jc is not None and jc.get(q('val')) == 'left':
                jc.set(q('val'), 'right')

# ================================================================== 6. figures
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'b2'))
import diagrams
from PIL import Image
os.makedirs(f'{OUT}/word/media', exist_ok=True)
FIGDIR = os.path.join(HERE, f'figs_{MODE}')
os.makedirs(FIGDIR, exist_ok=True)
rels_p = f'{OUT}/word/_rels/document.xml.rels'
rels = open(rels_p, encoding='utf-8').read()
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'


def image_para(rid, cx, cy, n, name):
    xml = (f'<w:p xmlns:w="{W}" xmlns:r="{R_NS}" xmlns:wp="{WP}" xmlns:a="{A_NS}" xmlns:pic="{PIC}">'
           '<w:pPr><w:keepNext/><w:bidi/><w:spacing w:before="160" w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
           f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{1000 + n}" name="{name}"/>'
           '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
           f'<pic:nvPicPr><pic:cNvPr id="{1000 + n}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
           f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
           f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
           '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    return etree.fromstring(xml)


fig_no, new_rels = 0, []
CAPS = []
for p in top_paras():                              # figure labels in the source text would clash with the figure numbering
    ts = [x for x in p.iter(q('t')) if x.text]
    if ts and re.match(r'^شكل\s*[\d٠-٩]+\s*[:：]\s*', ts[0].text):
        ts[0].text = re.sub(r'^شكل\s*[\d٠-٩]+\s*[:：]\s*', '', ts[0].text)
for anchor, spec, part in DIAG_SPECS:
    if anchor.getparent() is None:
        LOG.append(f'diagram anchor gone: {spec.get("title")}')
        continue
    fig_no += 1
    path = os.path.join(FIGDIR, f'fig{fig_no:02d}.png')
    try:
        diagrams.render(spec, path)
    except Exception as ex:
        LOG.append(f'diagram failed {spec.get("title")}: {ex}')
        fig_no -= 1
        continue
    shutil.copy(path, f'{OUT}/word/media/fig{fig_no:02d}.png')
    w, h = Image.open(path).size
    cx = min(5900000, int(w / 200 * 914400))
    cy = int(cx * h / w)
    if cy > 7600000:
        cy = 7600000
        cx = int(cy * w / h)
    rid = f'rIdFig{fig_no}'
    new_rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    f'Target="media/fig{fig_no:02d}.png"/>')
    ip = image_para(rid, cx, cy, fig_no, f'fig{fig_no}')
    cap_txt = re.sub(r'^شكل[^:：]*[:：]\s*', '', spec.get('caption') or spec.get('title', ''))
    anchor.addnext(ip)
    cap_el = new_p(T_CAPTION, f'شكل {fig_no}: {cap_txt}')
    ip.addnext(cap_el)
    CAPS.append(cap_el)
rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
open(rels_p, 'w', encoding='utf-8').write(rels)
cap_ids = {id(c) for c in CAPS}
k = 0
for p in top_paras():                              # figures numbered in page order
    if id(p) in cap_ids and p.getparent() is body:
        k += 1
        ts = [x for x in p.iter(q('t')) if x.text]
        ts[0].text = re.sub(r'^شكل \d+:', f'شكل {k}:', ts[0].text)
ct_p = f'{OUT}/[Content_Types].xml'
ct = open(ct_p, encoding='utf-8').read()
if 'Extension="png"' not in ct:
    ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>')
open(ct_p, 'w', encoding='utf-8').write(ct)

# ================================================================== 7. back matter order: conclusion, appendices, indexes, references
bib_head = next((p for p in top_paras() if re.match(r'^المراجع\s*[:：]?\s*الكتاب الثالث', text(p)) or text(p) == 'المراجع والمصادر'), None)
first_app = next((p for p in top_paras() if text(p).startswith(('الملاحق', 'الملحق أ')) and len(text(p)) < 90), None)
bib = []
if bib_head is not None:
    cur = bib_head
    while cur is not None and cur is not first_app and cur.tag != q('sectPr'):
        bib.append(cur)
        cur = cur.getnext()
    smart_set(bib_head, 'المراجع والمصادر')
    for e in bib:
        sect.addprevious(e)
else:
    LOG.append('bibliography heading not found')
INDEX_HEAD = new_p(TEMPL['heading'], 'الفهارس')
(bib_head if bib_head is not None else sect).addprevious(INDEX_HEAD)
if first_app is not None:
    if text(first_app).startswith('الملاحق'):
        smart_set(first_app, 'الملاحق')
    else:
        first_app.addprevious(new_p(TEMPL['heading'], 'الملاحق'))

# ================================================================== 8. verse and hadith indexes
from importlib.machinery import SourceFileLoader
SURAHS = SourceFileLoader('surahs', os.path.join(os.path.dirname(HERE), 'surahs.py')).load_module().SURAHS
UNITS, loc, chapter = [], 'التقديم', None
for e in body:
    if e is INDEX_HEAD:
        break
    if e.tag == q('p'):
        t = text(e)
        m = CH_HEAD.match(t)
        if m:
            chapter = ORD.index(m.group(1)) + 1
            loc = f'ف{chapter}'
        elif t == 'الخاتمة' or t.startswith('الخاتمة'):
            chapter, loc = None, 'الخاتمة'
        elif t.startswith('الملحق ') and len(t) < 90:
            chapter, loc = None, t.split(':')[0]
        elif SEC_HEAD.match(t):
            loc = SEC_HEAD.match(t).group(0).strip()
        if t:
            UNITS.append((loc, t))
    elif e.tag == q('tbl'):
        for p in e.iter(q('p')):
            if text(p):
                UNITS.append((loc, text(p)))
DIG = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
alt = '|'.join(sorted(map(re.escape, SURAHS), key=len, reverse=True))
VREF = re.compile(r'(?<![\w])(?:سورة\s+)?(' + alt + r')\s*[:：]\s*(?:الآية\s*)?([\d٠-٩]+)(?:\s*[-–]\s*([\d٠-٩]+))?')
QUOTE = re.compile(r'﴿([^﴾]+)﴾')
V, prev_q = {}, None
for where, s in UNITS:
    quotes = [(m.start(), m.end(), m.group(1)) for m in QUOTE.finditer(s)]
    bare = s if (QURAN.search(s) and not quotes and len(s) < 260 and not re.search(r'[a-zA-Z]', s)) else None
    for m in VREF.finditer(s):
        a = int(m.group(2).translate(DIG))
        b = int(m.group(3).translate(DIG)) if m.group(3) else None
        key = (SURAHS.index(m.group(1)), a, b)
        before = [qq for qq in quotes if qq[1] <= m.start() and m.start() - qq[1] < 8]
        txt = before[-1][2] if before else (prev_q if len(s) < 60 else None)
        ent = V.setdefault(key, {'text': None, 'locs': []})
        if txt and not ent['text']:
            ent['text'] = txt
        if where not in ent['locs']:
            ent['locs'].append(where)
    prev_q = quotes[-1][2] if quotes else (bare if bare else (prev_q if len(s) < 60 else None))


def short(t, n=7):
    w = t.split()
    return ' '.join(w[:n]) + ('…' if len(w) > n else '')


VERSES = [('الآية', 'طرفها', 'مواضعها')]
for key in sorted(V, key=lambda k: (k[0], k[1], k[2] or 0)):
    si, a, b = key
    VERSES.append((f'{SURAHS[si]}: {a}' + (f'–{b}' if b else ''), f'﴿{short(V[key]["text"])}﴾' if V[key]['text'] else 'إحالة دون نص',
                   '، '.join(V[key]['locs'][:12])))
HAD = re.compile(r'«([^»]{8,160})»\s*[؛،:\-]?\s*\(?\s*(?:رواه|متفق عليه|أخرجه|البخاري|مسلم|الترمذي|أبو داود|النسائي|ابن ماجه|أحمد)')
H = {}
for where, s in UNITS:
    for m in HAD.finditer(s):
        key = re.sub(r'[ً-ْـ]', '', m.group(1))[:28]
        src = re.split(r'[).؛]', s[m.end(1) + 1:m.end(1) + 60].replace('(', ' '))[0].strip(' ،-')
        ent = H.setdefault(key, {'text': m.group(1), 'src': src, 'locs': []})
        if where not in ent['locs']:
            ent['locs'].append(where)
HADITHS = [('طرف الحديث', 'التخريج كما ورد', 'مواضعه')] + [(f'«{short(v["text"], 9)}»', v['src'][:40], '، '.join(v['locs'][:10]))
                                                        for v in H.values()]


def make_table(rows3):
    t = fresh(copy.deepcopy(T_TABLE3))
    trs = t.findall(q('tr'))
    for tr in trs[2:]:
        t.remove(tr)
    while len(t.findall(q('tr'))) < len(rows3):
        add_row(t, 1, ['', '', ''])
    for r, row in enumerate(rows3):
        for c, val in enumerate(row):
            cell_text(t, r, c, val)
    tp = t.find(q('tblPr'))
    if tp is None:
        tp = etree.Element(q('tblPr'))
        t.insert(0, tp)
    ensure(tp, 'bidiVisual', TBLPR_ORDER)
    for p in t.iter(q('p')):
        ensure(ppr(p), 'bidi', PPR_ORDER)
    return t


idx_block = [new_p(TEMPL['subheading'], 'أولًا: فهرس الآيات القرآنية'),
             new_p(TEMPL['body'], 'الآيات مرتبة بترتيب المصحف، ومع كل آية طرفها كما ورد في الكتاب، ومواضعها برقم القسم (مثل 6.3 أي الفصل السادس، القسم الثالث) أو برقم الفصل.'),
             make_table(VERSES),
             new_p(TEMPL['subheading'], 'ثانيًا: فهرس الأحاديث والآثار'),
             new_p(TEMPL['body'], 'الأحاديث والآثار مرتبة بحسب أول ورودها، والتخريج هو ما ذكره الكتاب في موضعه.'),
             make_table(HADITHS)]
prev = INDEX_HEAD
for e in idx_block:
    prev.addnext(e)
    prev = e

# ================================================================== 9. layout: empty paragraphs, pages, TOC, page numbers
for e in list(body):
    if e.tag == q('p') and not text(e) and e.find('.//' + q('drawing')) is None and e.find('.//' + q('sectPr')) is None:
        pr, nx = e.getprevious(), e.getnext()
        if MODE == 'cond' or (pr is not None and pr.tag == q('p') and not text(pr) and pr.find('.//' + q('drawing')) is None):
            if not (pr is not None and nx is not None and pr.tag == q('tbl') and nx.tag == q('tbl')):
                body.remove(e)
PART = re.compile(r'^الباب (التمهيدي|الأول|الثاني|الثالث|الرابع): ')
TOP = re.compile(r'^(تقديم|الخاتمة|الملاحق|الملحق [أ-ي]{1,2}:|الفهارس|المراجع والمصادر)')
TC_LIST = []


def tc(p, label, level):
    label = re.sub(r'\s*\n\s*', '، ', label).replace('"', '')
    TC_LIST.append((level, label))
    last = p.findall(q('r'))[-1] if p.findall(q('r')) else None
    for kind in ('begin', 'instr', 'end'):
        r = etree.Element(q('r'))
        if kind == 'instr':
            it = etree.SubElement(r, q('instrText'))
            it.set(XML_SPACE, 'preserve')
            it.text = f' TC "{label}" \\l {level} '
        else:
            f_ = etree.SubElement(r, q('fldChar'))
            f_.set(q('fldCharType'), kind)
        if last is not None:
            last.addnext(r)
            last = r
        else:
            p.append(r)


prev_text = ''
children = top_paras()
for i, p in enumerate(children):
    t = text(p)
    if not t:
        continue
    m = CH_HEAD.match(t)
    if PART.match(t) and len(t) < 90:
        ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        tc(p, t, 1)
    elif m and sz_of(p) >= 44:
        if not PART.match(prev_text):
            ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        nxt = next((text(x) for x in children[i + 1:i + 3] if text(x)), '')
        tc(p, f'{t}: {nxt}', 2)
    elif TOP.match(t) and len(t) < 100 and sz_of(p) >= 30:
        if prev_text != 'الملاحق':
            ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        tc(p, t, 2 if t.startswith('الملحق ') else 1)
    elif SEC_HEAD.match(t) and len(t) < 160:
        tc(p, t, 3)
    prev_text = t
intro = next(p for p in top_paras() if text(p).startswith('تقديم'))
toc_head = new_p(TEMPL['heading'], 'المحتويات')
ensure(ppr(toc_head), 'pageBreakBefore', PPR_ORDER)
entries = []
for lv, lab in TC_LIST:
    if lv == 3:
        continue
    p_ = new_p(TEMPL['subheading'] if lv == 1 else TEMPL['body'], lab)
    for r in p_.findall(q('r')):
        rp = r.find(q('rPr'))
        if rp is not None and lv == 2:
            for tag in ('b', 'bCs'):
                el = rp.find(q(tag))
                if el is not None:
                    rp.remove(el)
    if lv == 2:
        ind = ensure(ppr(p_), 'ind', PPR_ORDER)
        ind.set(q('right'), '567')
        sp_ = ensure(ppr(p_), 'spacing', PPR_ORDER)
        sp_.set(q('before'), '0')
        sp_.set(q('after'), '0')
    entries.append(p_)


def fld_run(kind=None, instr=None):
    r = etree.Element(q('r'))
    if kind:
        f_ = etree.SubElement(r, q('fldChar'))
        f_.set(q('fldCharType'), kind)
        if kind == 'begin':
            f_.set(q('dirty'), 'true')
    else:
        it = etree.SubElement(r, q('instrText'))
        it.set(XML_SPACE, 'preserve')
        it.text = instr
    return r


pos = 1 if entries[0].find(q('pPr')) is not None else 0
entries[0].insert(pos, fld_run('begin'))
entries[0].insert(pos + 1, fld_run(instr=' TOC \\f \\l "1-2" \\z '))
entries[0].insert(pos + 2, fld_run('separate'))
entries[-1].append(fld_run('end'))
intro.addprevious(toc_head)
for e in entries:
    intro.addprevious(e)
ensure(ppr(intro), 'pageBreakBefore', PPR_ORDER)
footer = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}" xmlns:r="{R_NS}"><w:p><w:pPr><w:bidi/><w:jc w:val="center"/></w:pPr>'
          '<w:r><w:rPr><w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic" w:cs="Traditional Arabic"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
          '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
          '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')
open(f'{OUT}/word/footer1.xml', 'w', encoding='utf-8').write(footer)
rels = open(rels_p, encoding='utf-8').read()
if 'footer1.xml' not in rels:
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
    open(rels_p, 'w', encoding='utf-8').write(rels)
ct = open(ct_p, encoding='utf-8').read()
if 'footer1.xml' not in ct:
    ct = ct.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    open(ct_p, 'w', encoding='utf-8').write(ct)
SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
              'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']
for old in sect.findall(q('footerReference')):
    sect.remove(old)
fr = etree.Element(q('footerReference'))
fr.set(q('type'), 'default')
fr.set('{%s}id' % R_NS, 'rIdFtr1')
sect.insert(0, fr)
ensure(sect, 'titlePg', SECT_ORDER)
ensure(sect, 'bidi', SECT_ORDER)
st_p = f'{OUT}/word/settings.xml'
st = etree.parse(st_p)
sr = st.getroot()
if sr.find(q('updateFields')) is None:
    uf = etree.Element(q('updateFields'))
    uf.set(q('val'), 'true')
    AFTER = ['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang',
             'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown',
             'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator']
    nxt = next((ch for ch in sr if etree.QName(ch).localname in AFTER), None)
    (nxt.addprevious(uf) if nxt is not None else sr.append(uf))
st.write(st_p, xml_declaration=True, encoding='UTF-8', standalone=True)
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
NAME = {'full': 'موسوعة_الإدارة_القرآنية_الكتاب_الثالث_النسخة_الكاملة.docx', 'cond': 'موسوعة_الإدارة_القرآنية_الكتاب_الثالث_للطباعة.docx'}[MODE]
out = os.path.join(HERE, NAME)
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-qXr', out, '.'], cwd=OUT, check=True)

# ================================================================== 10. companion guide (cond mode)
if MODE == 'cond' and COMP_COPIES:
    CDIR = os.path.join(HERE, 'build_comp')
    shutil.rmtree(CDIR, ignore_errors=True)
    with zipfile.ZipFile(SRC) as z:
        z.extractall(CDIR)
    ctree = etree.parse(f'{CDIR}/word/document.xml')
    cbody = ctree.getroot().find(q('body'))
    csect = copy.deepcopy(cbody.find(q('sectPr')))
    for e in list(cbody):
        cbody.remove(e)
    for tm, s in zip(T_TITLE[:3] + [T_TITLE[3]], ['موسوعة الإدارة القرآنية', 'الكتاب الثالث', 'الدليل العملي المرافق', 'م. وليد عبدالله حواله']):
        cbody.append(new_p(tm, s))
    cbody.append(new_p(TEMPL['body'], 'يضم هذا الدليل المواد التعليمية والعملية التي فُصلت عن الكتاب الثالث عند اختصاره: مخرجات التعلم والتمارين لكل فصل، '
                                      'وما نُقل من نماذج وأدوات مفصّلة. وهو يُستعمل مع الكتاب ولا يغني عنه.'))
    for part, label, els in COMP_COPIES:
        h = new_p(TEMPL['heading'], label)
        cbody.append(h)
        for e in els:
            if e.tag == q('p'):
                strip_fields(e)
            if e.tag == q('tbl'):
                tp = e.find(q('tblPr'))
                if tp is None:
                    tp = etree.Element(q('tblPr'))
                    e.insert(0, tp)
                trs = e.findall(q('tr'))
                if reversed_score(e) >= 2:
                    for tr in trs:
                        tcs = tr.findall(q('tc'))
                        for tc_ in tcs:
                            tr.remove(tc_)
                        for tc_ in reversed(tcs):
                            tr.append(tc_)
                ensure(tp, 'bidiVisual', TBLPR_ORDER)
            cbody.append(fresh(e))
    cbody.append(csect)
    ctree.write(f'{CDIR}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
    cout = os.path.join(HERE, 'الدليل_العملي_المرافق_للكتاب_الثالث.docx')
    if os.path.exists(cout):
        os.remove(cout)
    subprocess.run(['zip', '-qXr', cout, '.'], cwd=CDIR, check=True)

words = sum(len(ptext(e).split()) for e in body.iter(q('p')))
json.dump({'mode': MODE, 'stats': stats, 'log': LOG, 'words': words, 'figures': fig_no, 'reversed_tables': rev_n, 'ambiguous_tables': amb_list,
           'dash_fixed': dash_n, 'notbut_fixed': nb_n, 'unbold_runs': unbold, 'sections': dict(n_sections), 'verses': len(VERSES) - 1,
           'hadiths': len(HADITHS) - 1, 'toc_entries': len(TC_LIST), 'companion_blocks': [(p, l, len(e)) for p, l, e in COMP_COPIES]},
          open(os.path.join(HERE, f'build_log_{MODE}.json'), 'w'), ensure_ascii=False, indent=1)
print('wrote', out, '| words', words, '| skips', sum(v['skip'] for v in stats.values()), '| figs', fig_no, '| reversed', rev_n,
      '| ambiguous', len(amb_list), '| verses', len(VERSES) - 1, '| hadiths', len(HADITHS) - 1)
