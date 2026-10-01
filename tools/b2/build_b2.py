"""Book Two, single print-ready file.

Sources: QME_B2_All_Chapters_Complete.docx (chapters 1-14 + appendices A, B; the 1-8 file is a subset of it)
and QME_B2_Front_Back.docx (front matter, conclusion, indexes, references).
Steps: apply reviewers' ID-based edits -> assemble -> renumber sections -> style clean-up -> tables RTL (un-reversing
fake-RTL tables) -> diagrams -> verse/hadith indexes -> TOC field, page breaks, page numbers.
"""
import sys, os, re, json, glob, copy, shutil, subprocess, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..'))
sys.path.insert(0, HERE)
from docxedit import *
import diagrams
from surahs import SURAHS

os.chdir(HERE)
OUT = 'b2_build'
shutil.rmtree(OUT, ignore_errors=True)
shutil.copytree('all_x', OUT)
tree = etree.parse(f'{OUT}/word/document.xml')
body = tree.getroot().find(q('body'))
sect = body.find(q('sectPr'))
A = {i: e for i, e in enumerate(body) if e.tag != q('sectPr')}
fb_tree = etree.parse('fb_x/word/document.xml')
F = {i: fresh(copy.deepcopy(e)) for i, e in enumerate(fb_tree.getroot().find(q('body'))) if e.tag != q('sectPr')}
HOLD = etree.Element(q('body'))       # keeps the front/back elements attached so edits can insert and delete
for _i in sorted(F):
    HOLD.append(F[_i])
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
LOG = []

# ------------------------------------------------------------------ helpers
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
             'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing',
               'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']


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
    if rp is None:
        return False
    b = rp.find(q('b'))
    return b is not None and b.get(q('val')) not in ('false', '0')


def smart_set(p, new):
    """Replace paragraph text; keep a bold lead-in label run if the new text starts with it."""
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


# ------------------------------------------------------------------ 1. edits
def resolve(id_, part):
    fb = part == 'front_back'
    src = F if fb else A
    m = re.match(r'^T(?:BL)?([PF])?(\d+)(?:\.(\d+)\.(\d+))?$', id_)
    if id_.startswith(('TP', 'TF', 'TBL')) and m:
        idx = int(m.group(2))
        tbl = (F if (m.group(1) == 'F' or (m.group(1) is None and fb)) else A).get(idx)
        if tbl is None or tbl.tag != q('tbl'):
            return None, None
        if m.group(3) is None:
            return tbl, 'tbl'
        try:
            tc = tbl.findall(q('tr'))[int(m.group(3))].findall(q('tc'))[int(m.group(4))]
        except IndexError:
            return None, None
        return tc, 'cell'
    m = re.match(r'^([PF])(\d+)$', id_)
    if m:
        e = (F if m.group(1) == 'F' else A).get(int(m.group(2)))
        return e, ('p' if e is not None and e.tag == q('p') else ('tbl' if e is not None else None))
    return None, None


def idx_of(id_):
    m = re.search(r'(\d+)', id_)
    return int(m.group(1))


TEMPL = {'body': A[42], 'heading': A[26], 'subheading': A[47]}
EDITS = {}
for f in sorted(glob.glob('edits/*.json')):
    try:
        EDITS[os.path.basename(f)[:-5]] = json.load(open(f, encoding='utf-8'))
    except Exception as ex:
        LOG.append(f'BAD JSON {f}: {ex}')

DIAG_SPECS = []          # (anchor element, spec)
stats = {}
for part, data in EDITS.items():
    src = F if part == 'front_back' else A
    ops = data.get('edits', [])
    st = stats.setdefault(part, {'ok': 0, 'skip': 0})
    deletes, moves = [], []
    for op in ops:
        kind = op.get('op')
        try:
            if kind == 'replace':
                el, typ = resolve(op['id'], part)
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
                el, typ = resolve(op['id'], part)
                if el is None:
                    raise ValueError('bad id')
                tmpl = TEMPL.get(op.get('style', 'body'), TEMPL['body'])
                new = fresh(copy.deepcopy(tmpl))
                set_text(new, op['text'])
                (el if typ != 'cell' else el.getparent().getparent().getparent()).addnext(new)
            elif kind == 'delete':
                for i in op.get('ids', []):
                    el, typ = resolve(i, part)
                    if el is None:
                        raise ValueError('bad id ' + i)
                    deletes.append(el if typ != 'cell' else None)
            elif kind == 'delete_range':
                a, b = idx_of(op['from']), idx_of(op['to'])
                deletes += [src[i] for i in range(a, b + 1) if i in src]
            elif kind == 'move_range':
                a, b = idx_of(op['from']), idx_of(op['to'])
                anchor, _ = resolve(op['after'], part)
                if anchor is None:
                    raise ValueError('bad anchor')
                moves.append(([src[i] for i in range(a, b + 1) if i in src], anchor, src.get(a), src.get(b)))
            else:
                raise ValueError('unknown op')
            st['ok'] += 1
        except Exception as ex:
            st['skip'] += 1
            LOG.append(f'{part}: skip {kind} {op.get("id") or op.get("from")}: {ex}')
    for els, anchor, ea, eb in moves:
        # walk the current document order from ea to eb (handles chained moves)
        if ea is not None and eb is not None and ea.getparent() is eb.getparent():
            walk, cur = [], ea
            while cur is not None and cur is not eb and len(walk) < 2000:
                walk.append(cur); cur = cur.getnext()
            if cur is eb:
                els = walk + [eb]
        if anchor in els:
            LOG.append(f'{part}: move anchor inside range, skipped')
            continue
        for e in reversed(els):
            anchor.addnext(e)
    for d in data.get('diagrams', []):
        el, typ = resolve(d.get('after', ''), part)
        if el is not None and typ == 'p':
            DIAG_SPECS.append((el, d, part))
        else:
            LOG.append(f'{part}: diagram anchor bad {d.get("after")}')
    for el in deletes:
        if el is not None and el.getparent() is not None:
            el.getparent().remove(el)

# ------------------------------------------------------------------ 2. assemble
def fidx_text(start_pred):
    for i, e in sorted(F.items()):
        if start_pred(text(e)):
            return i
    return None


f_concl = fidx_text(lambda t: t.startswith('الخاتمة'))
f_index = fidx_text(lambda t: t == 'الفهارس')
f_refs = fidx_text(lambda t: 'المصادر والمراجع' in t and len(t) < 40)
f_done = fidx_text(lambda t: t.startswith('الكتاب الثاني مكتمل'))
SEQ = list(HOLD)
def fpos(i):
    if i is None:
        return len(SEQ)
    for j in sorted(F):
        if j >= i and F[j].getparent() is HOLD:
            return SEQ.index(F[j])
    return len(SEQ)
pc, pi, pr, pd = fpos(f_concl), fpos(f_index), fpos(f_refs), fpos(f_done)
front, conclusion, fb_indexes, fb_refs, fb_end = SEQ[:pc], SEQ[pc:pi], SEQ[pi:pr], SEQ[pr:pd], SEQ[pd:]
# drop fb's verse and hadith index sections (regenerated from the whole book) and the page-placeholder tables
keep_idx, skipping = [], False
for e in fb_indexes:
    t = text(e)
    if re.match(r'^(أوَّلاً|أولًا|ثانيًا): فهرس (الآيات|الأحاديث)', t):
        skipping = True
        continue
    if skipping and re.match(r'^(ثالثًا|رابعًا|خامسًا|سادسًا|سابعًا|ثامنًا|تاسعًا):', t):
        skipping = False
    if skipping:
        continue
    if e.tag == q('tbl'):
        hdr = ' '.join(ptext(tc) for tc in e.find(q('tr')).findall(q('tc')))
        if 'الصفحات' in hdr:
            continue
    if re.match(r'^ثامنًا: فهرس الفصول', t):
        continue
    keep_idx.append(e)
fb_indexes = keep_idx

all_children = [e for e in body if e.tag != q('sectPr')]
ch1 = next(e for e in all_children if text(e) == 'الفصل الأول')
appA = next(e for e in all_children if text(e) == 'ملحق (أ)')
for e in all_children[:all_children.index(ch1)]:               # the chapters file's own mini title
    if e.getparent() is not None and not text(e).startswith('الباب'):
        body.remove(e)
for e in front:
    ch1.addprevious(e)
for e in conclusion:
    appA.addprevious(e)
INDEX_ANCHOR = etree.Element(q('p'))
sect.addprevious(INDEX_ANCHOR)
for e in fb_indexes:
    INDEX_ANCHOR.addprevious(e)
for e in fb_refs + fb_end:
    sect.addprevious(e)

# ------------------------------------------------------------------ 1b. references
def _add_refs(after_id, tmpl_id, items):
    prev = F[after_id]
    for t in items:
        new = fresh(copy.deepcopy(F[tmpl_id]))
        set_text(new, t)
        prev.addnext(new)
        prev = new

set_text(F[168], 'أوَّلاً: القرآن الكريم والتفسير والحديث')
_add_refs(173, 169, [
    '6. البخاري، محمَّد بن إسماعيل. الجامع الصحيح (صحيح البخاري). دار طوق النجاة، ١٤٢٢ هـ.',
    '7. مسلم بن الحجاج النيسابوري. المسند الصحيح (صحيح مسلم). دار إحياء التراث العربي، د.ت.',
    '8. الترمذي، محمَّد بن عيسى. السنن (الجامع). مطبعة مصطفى البابي الحلبي، ١٩٧٥.',
    '9. أبو داود، سليمان بن الأشعث. السنن. المكتبة العصرية، بيروت، د.ت.',
])
set_text(F[184], '4. الغزالي، أبو حامد. إحياء علوم الدين. دار المعرفة، بيروت، د.ت.')
_add_refs(184, 184, [
    '5. حواله، وليد عبدالله. موسوعة الإدارة القرآنية، الكتاب الأول: الأسس الوجودية والمعرفية والمنهجية للإدارة القرآنية، من النص القرآني إلى النظرية الإدارية.',
])
_add_refs(195, 187, [
    '10. Arrow, K.J. (1951). Social Choice and Individual Values. Wiley.',
    '11. Coombs, W.T. (2007). Ongoing Crisis Communication: Planning, Managing, and Responding (2nd ed.). Sage.',
    '12. Davis, J.H., Schoorman, F.D., & Donaldson, L. (1997). Toward a Stewardship Theory of Management. Academy of Management Review, 22(1), 20–47.',
    '13. Edmans, A. (2011). Does the Stock Market Fully Value Intangibles? Employee Satisfaction and Equity Prices. Journal of Financial Economics, 101(3), 621–640.',
    '14. Edmondson, A. (1999). Psychological Safety and Learning Behavior in Work Teams. Administrative Science Quarterly, 44(2), 350–383.',
    '15. Friedman, M. (1962). Capitalism and Freedom. University of Chicago Press.',
    '16. Global Reporting Initiative. (2021). GRI Universal Standards. GRI.',
    '17. Greenleaf, R.K. (1977). Servant Leadership. Paulist Press.',
    '18. Habermas, J. (1984). The Theory of Communicative Action, Vol. 1. Beacon Press.',
    '19. Hiatt, J. (2006). ADKAR: A Model for Change in Business, Government and Our Community. Prosci.',
    '20. Jensen, M.C., & Meckling, W.H. (1976). Theory of the Firm: Managerial Behavior, Agency Costs and Ownership Structure. Journal of Financial Economics, 3(4), 305–360.',
    '21. Kahneman, D. (2011). Thinking, Fast and Slow. Farrar, Straus and Giroux.',
    '22. Kim, W.C., & Mauborgne, R. (1997). Fair Process: Managing in the Knowledge Economy. Harvard Business Review, 75(4), 65–75.',
    '23. Nicholls, J., Lawlor, E., Neitzert, E., & Goodspeed, T. (2012). A Guide to Social Return on Investment. The SROI Network.',
    '24. Porter, M.E., & Kramer, M.R. (2011). Creating Shared Value. Harvard Business Review, 89(1/2), 62–77.',
    '25. Schein, E.H. (2010). Organizational Culture and Leadership (4th ed.). Jossey-Bass.',
    '26. Taleb, N.N. (2012). Antifragile: Things That Gain from Disorder. Random House.',
    '27. Tedeschi, R.G., & Calhoun, L.G. (1996). The Posttraumatic Growth Inventory. Journal of Traumatic Stress, 9(3), 455–471.',
    '28. Zak, P.J. (2017). The Neuroscience of Trust. Harvard Business Review, 95(1), 84–90.',
])

# ------------------------------------------------------------------ 3. clean-up
STATUS = re.compile(r'(مكتمل\s*✅|✅\s*مكتمل|⏳|^التالي\s*:|^التالي \()')
for p in list(body.iter(q('p'))):
    t = text(p)
    if not t or p.getparent() is not body:
        continue
    if t.startswith('§') and re.search(r'\s*—\s*الباب \S+ مكتمل\s*$', t):
        smart_set(p, re.sub(r'\s*—\s*الباب \S+ مكتمل\s*$', '', t))
        continue
    if (STATUS.search(t) or re.search(r'مكتمل\s*$', t) or ('✅' in t and len(t) < 40)) and len(t) < 160 and not t.startswith('§'):
        body.remove(p)
        LOG.append('status line removed: ' + t[:60])
    elif '✅' in t:
        smart_set(p, re.sub(r'\s*✅\s*', '؛ ', re.sub(r':\s*✅\s*', ': ', t)).strip('؛ '))
MARK = re.compile(r'^[\s●◆◈▌■□▪►•✦✓✔☑]+')
EMOJI = re.compile('[\U0001F000-\U0001FFFF☀-⛿✀-➿️]')
for p in body.iter(q('p')):
    ts = [t for t in p.iter(q('t')) if t.text]
    if ts and MARK.match(ts[0].text):
        ts[0].text = MARK.sub('', ts[0].text)
for t in body.iter(q('t')):
    if t.text:
        s = EMOJI.sub(lambda m: m.group(0) if m.group(0) in '★✓' else '', t.text)
        s = s.replace('وليد حوالة', 'وليد حواله')
        t.text = s

# ------------------------------------------------------------------ 4. section renumbering
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر',
       'الثالث عشر', 'الرابع عشر']
CH_HEAD = re.compile(r'^الفصل (' + '|'.join(sorted(ORD, key=len, reverse=True)) + r')(\s.*)?$')
SEC = re.compile(r'^§\s?(\d+)\.(\d+)((?:-[ء-ي]+)?)\s+')
chapter_of = {}
cur = None
SECMAP = {}
counter = {}
for p in body:
    if p.tag != q('p'):
        continue
    t = text(p)
    m = CH_HEAD.match(t)
    if m and len(t) < 40:
        cur = ORD.index(m.group(1)) + 1
        counter[cur] = 0
        continue
    if text(p).startswith('ملحق (') or t == 'الخاتمة' or t.startswith('الخاتمة —'):
        cur = None
    ms = SEC.match(t)
    if ms and cur:
        counter[cur] += 1
        new = f'{cur}.{counter[cur]}'
        old = f'{ms.group(1)}.{ms.group(2)}{ms.group(3)}'
        SECMAP.setdefault((cur, old), new)
        ts = [x for x in p.iter(q('t')) if x.text]
        full = ''.join(x.text for x in ts)
        rest = SEC.sub('', full, count=1)
        ts[0].text = f'{new}  ' + rest
        for x in ts[1:]:
            x.text = ''
        chapter_of[p] = cur
REF = re.compile(r'§\s?(\d+)\.(\d+)((?:-[ء-ي]+)?)')


def map_ref(m):
    a = int(m.group(1))
    key = (a, f'{m.group(1)}.{m.group(2)}{m.group(3)}')
    if key in SECMAP:
        return SECMAP[key]
    key2 = (a, f'{m.group(1)}.{m.group(2)}')
    return SECMAP.get(key2, f'{m.group(1)}.{m.group(2)}')


for t in body.iter(q('t')):
    if t.text and '§' in t.text:
        t.text = REF.sub(map_ref, t.text).replace('§', '')

# ------------------------------------------------------------------ 5. style: dashes, «ليس… بل»
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
    caption = bool(CAPTION.match(full))
    sep = ': ' if len(full) < 70 else '؛ '
    for n, t in enumerate(ts):
        s0 = t.text
        s = s0
        if caption and '—' in s:
            s = re.sub(r'\s*—\s*', ': ', s, count=1)
        if '—' in s and not caption and full != '—':
            s = fix_sentence_dashes(s, sep)
            if s.startswith((':', '؛', '،')) and n > 0:
                ts[n - 1].text = ts[n - 1].text.rstrip()
            if s.endswith(' ') and n + 1 < len(ts):
                ts[n + 1].text = ts[n + 1].text.lstrip()
        s2 = NOT_BUT.sub(r'\1، وإنما ', s)
        nb_n += s2 != s
        dash_n += s0.count('—') - s2.count('—')
        t.text = s2.replace('؛ ،', '؛').replace('، ،', '،').replace('  ', ' ') if not re.match(r'^\d+\.\d+  ', s2) else s2

# ------------------------------------------------------------------ 6. tables right-to-left
KEY_HEADS = {'#', 'م', 'رمز', 'الرمز', 'رمز CLO', 'CLO', 'المرحلة', 'المستوى', 'الفصل', 'الباب', 'المكوِّن', 'المكون', 'المعيار', 'المحك',
             'المحكّ', 'المحور', 'الأداة', 'البُعد', 'البعد', 'الركن', 'المفهوم', 'المصطلح', 'الحالة', 'الخطوة', 'الطبقة', 'السبب',
             'الاستنتاج', 'النمط', 'الفئة', 'العنصر', 'الكتاب', 'العَلَم', 'العلم', 'الآية', 'الحديث', 'النموذج', 'الإطار', 'المؤشر',
             'السؤال', 'القرار', 'المدرسة', 'البند', 'المنصب', 'المرشح', 'نوع الأزمة', 'الوتيرة', 'نمط المؤسسة', 'الفارق', 'الإحالة', 'الجانب', 'الوظيفة', 'الدور', 'الصنف', 'النوع', 'الأسبوع', 'الشهر', 'اليوم',
             'الوحدة', 'الجلسة', 'التاريخ', 'الفترة', 'السنة', 'الربع', 'المجال', 'القطاع', 'الطرف', 'صاحب المصلحة', 'الأصل'}


def norm_head(s):
    return re.sub(r'[ً-ْـ]', '', s).strip()


KEY_NORM = {norm_head(k) for k in KEY_HEADS}


def is_key(s):
    s = norm_head(s)
    return s in KEY_NORM or any(s.startswith(k + ' ') for k in KEY_NORM if len(k) > 2) or s.startswith('رمز')


def codeish(s):
    s = s.strip()
    if not s or set(s) <= set('_./ ') :
        return False
    return bool(re.match(r'^(\d+|[٠-٩]+|CLO\s?[\d.]+|[أ-ي]-[\d٠-٩]+|ف[\d٠-٩]+|\(?[أ-ي]\)|[①-⑳]|م\d)\b', s)) or len(s) <= 3


def reversed_score(tbl):
    rows = [[ptext(tc).strip() for tc in tr.findall(q('tc'))] for tr in tbl.findall(q('tr'))]
    rows = [r for r in rows if len(r) >= 2]
    if not rows:
        return 0
    hdr = rows[0]
    score = 0
    kl, kf = is_key(hdr[-1]) or norm_head(hdr[-1]).startswith(('نمط', 'نوع', 'بعد ', 'مجال')), is_key(hdr[0])
    if not hdr[-1].strip() and hdr[0].strip():
        score += 3                                         # empty corner cell belongs at the start
    if kl and not kf:
        score += 3
    if kf and not kl:
        score -= 3
    body_rows = rows[1:] or rows
    last_code = sum(codeish(r[-1]) for r in body_rows)
    first_code = sum(codeish(r[0]) for r in body_rows)
    if last_code > first_code and last_code >= len(body_rows) * 0.6:
        score += 2
    if first_code > last_code and first_code >= len(body_rows) * 0.6:
        score -= 2
    if len(hdr) >= 3 and re.match(r'^(درجتك|التقييم|ملاحظ|الشاهد|نتيجتي|الصفحات)', hdr[0]):
        score += 2
    return score


rev_n = amb = 0
for tbl in body.iter(q('tbl')):
    trs = tbl.findall(q('tr'))
    ncols = max(len(tr.findall(q('tc'))) for tr in trs)
    if ncols >= 2:
        sc = reversed_score(tbl)
        if sc >= 2:
            # un-reverse: cells, grid columns and per-cell widths
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
        elif sc > -2:
            amb += 1                                       # ambiguous: the generator mostly wrote columns reversed
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
    tp = tbl.find(q('tblPr'))
    if tp is None:
        tp = etree.Element(q('tblPr'))
        tbl.insert(0, tp)
    ensure(tp, 'bidiVisual', TBLPR_ORDER)
    for p in tbl.iter(q('p')):
        if re.search('[؀-ۿ]', ptext(p)):
            ensure(ppr(p), 'bidi', PPR_ORDER)

# ------------------------------------------------------------------ 7. diagrams
os.makedirs(f'{OUT}/word/media', exist_ok=True)
os.makedirs('figs', exist_ok=True)
rels_p = f'{OUT}/word/_rels/document.xml.rels'
rels = open(rels_p, encoding='utf-8').read()
CAPTION_TMPL = A[6]                     # centred gold subtitle style
fig_no = 0
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


from PIL import Image
new_rels = []
for anchor, spec, part in DIAG_SPECS:
    if anchor.getparent() is None:
        LOG.append(f'diagram anchor deleted: {spec.get("title")}')
        continue
    fig_no += 1
    path = f'figs/fig{fig_no:02d}.png'
    try:
        diagrams.render(spec, path)
    except Exception as ex:
        LOG.append(f'diagram failed {spec.get("title")}: {ex}')
        fig_no -= 1
        continue
    shutil.copy(path, f'{OUT}/word/media/fig{fig_no:02d}.png')
    w, h = Image.open(path).size
    maxw = 5900000                              # ~16.4 cm
    cx = min(maxw, int(w / 200 * 914400))
    cy = int(cx * h / w)
    if cy > 7600000:
        cy = 7600000
        cx = int(cy * w / h)
    rid = f'rIdFig{fig_no}'
    new_rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    f'Target="media/fig{fig_no:02d}.png"/>')
    ip = image_para(rid, cx, cy, fig_no, f'fig{fig_no}')
    cap = fresh(copy.deepcopy(CAPTION_TMPL))
    caption = spec.get('caption') or spec.get('title', '')
    caption = re.sub(r'\s*\(§[^)]*\)', '', caption)
    caption = re.sub(r'^شكل[^:：]*[:：]\s*', '', caption)
    set_text(cap, f'شكل {fig_no}: {caption}')
    anchor.addnext(ip)
    ip.addnext(cap)
rels = rels.replace('</Relationships>', ''.join(new_rels) + '</Relationships>')
open(rels_p, 'w', encoding='utf-8').write(rels)
ct_p = f'{OUT}/[Content_Types].xml'
ct = open(ct_p, encoding='utf-8').read()
if 'Extension="png"' not in ct:
    ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/>')
open(ct_p, 'w', encoding='utf-8').write(ct)

# ------------------------------------------------------------------ 8. verse and hadith indexes
SEC_HEAD = re.compile(r'^(\d+)\.(\d+)\s\s')
loc = None
chap = None
UNITS = []
for e in body:
    if e is INDEX_ANCHOR:
        break
    if e.tag == q('p'):
        t = text(e)
        m = CH_HEAD.match(t)
        if m and len(t) < 40:
            chap = ORD.index(m.group(1)) + 1
            loc = f'ف{chap}'
        elif t.startswith('الخاتمة'):
            chap, loc = None, 'الخاتمة'
        elif t.startswith('ملحق ('):
            chap, loc = None, t[:9]
        elif SEC_HEAD.match(t):
            loc = SEC_HEAD.match(t).group(0).strip()
        if t:
            UNITS.append((loc or 'المقدمة', t))
    elif e.tag == q('tbl'):
        for p in e.iter(q('p')):
            t = text(p)
            if t:
                UNITS.append((loc or 'المقدمة', t))
DIG = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
alt = '|'.join(sorted(map(re.escape, SURAHS), key=len, reverse=True))
VREF = re.compile(r'(?<![\w])(' + alt + r')\s*[:：]\s*([\d٠-٩]+)(?:\s*[-–]\s*([\d٠-٩]+))?')
QUOTE = re.compile(r'﴿([^﴾]+)﴾')
V = {}
prev_q = None
for where, s in UNITS:
    quotes = [(m.start(), m.end(), m.group(1)) for m in QUOTE.finditer(s)]
    for m in VREF.finditer(s):
        a = int(m.group(2).translate(DIG))
        b = int(m.group(3).translate(DIG)) if m.group(3) else None
        key = (SURAHS.index(m.group(1)), a, b)
        before = [qq for qq in quotes if qq[1] <= m.start() and m.start() - qq[1] < 8]
        txt = before[-1][2] if before else (prev_q if len(s) < 60 else None)
        e_ = V.setdefault(key, {'text': None, 'locs': []})
        if txt and not e_['text']:
            e_['text'] = txt
        if where not in e_['locs']:
            e_['locs'].append(where)
    prev_q = quotes[-1][2] if quotes else (prev_q if len(s) < 60 else None)


def short(t, n=7):
    w = t.split()
    return ' '.join(w[:n]) + ('…' if len(w) > n else '')


VERSES = [('الآية', 'طرفها', 'مواضعها')]
for key in sorted(V, key=lambda k: (k[0], k[1], k[2] or 0)):
    si, a, b = key
    VERSES.append((f'{SURAHS[si]}: {a}' + (f'–{b}' if b else ''), f'﴿{short(V[key]["text"])}﴾' if V[key]['text'] else 'إحالة دون نص',
                   '، '.join(V[key]['locs'][:12])))
HAD = re.compile(r'«([^»]{8,160})»\s*[؛،:—\-]?\s*\(?\s*(?:رواه|متفق عليه|أخرجه|البخاري|مسلم|الترمذي|أبو داود|النسائي|ابن ماجه|أحمد)')
H = {}
for where, s in UNITS:
    for m in HAD.finditer(s):
        key = re.sub(r'[ً-ْـ]', '', m.group(1))[:28]
        src = s[m.end(1) + 1:m.end(1) + 60]
        src = re.split(r'[).؛]', src.replace('(', ' '))[0].strip(' ،—-')
        e_ = H.setdefault(key, {'text': m.group(1), 'src': src, 'locs': []})
        if where not in e_['locs']:
            e_['locs'].append(where)
HADITHS = [('طرف الحديث', 'التخريج كما ورد', 'مواضعه')] + [(f'«{short(v["text"], 9)}»', v['src'][:40], '، '.join(v['locs'][:10]))
                                                        for v in H.values()]


def make_table(rows3):
    tmpl = next(t for t in fb_tree.getroot().iter(q('tbl')) if len(t.find(q('tr')).findall(q('tc'))) == 3)
    t = fresh(copy.deepcopy(tmpl))
    trs = t.findall(q('tr'))
    for tr in trs[2:]:
        t.remove(tr)
    while len(t.findall(q('tr'))) < len(rows3):
        add_row(t, 1, ['', '', ''])
    for r, row in enumerate(rows3):
        for c, val in enumerate(row):
            cell_text(t, r, c, val)
    tp = t.find(q('tblPr'))
    ensure(tp, 'bidiVisual', TBLPR_ORDER)
    return t


H1 = next(e for e in fb_indexes if text(e) == 'الفهارس') if any(text(e) == 'الفهارس' for e in fb_indexes) else None
HEAD_T = next(e for e in F.values() if text(e).startswith('أوَّلاً: فهرس') or text(e).startswith('أولًا: فهرس'))
BODY_T = next(e for e in F.values() if e.tag == q('p') and len(text(e)) > 80)


def head_p(s):
    p = fresh(copy.deepcopy(HEAD_T)); set_text(p, s); return p


def body_p(s):
    p = fresh(copy.deepcopy(BODY_T)); set_text(p, s); return p


new_idx = [head_p('فهرس الآيات القرآنية'),
           body_p('الآيات مرتبة بترتيب المصحف، ومع كل آية طرفها كما ورد في الكتاب، ومواضعها برقم القسم (مثل 4.3 أي الفصل الرابع، القسم الثالث) أو برقم الفصل.'),
           make_table(VERSES),
           head_p('فهرس الأحاديث والآثار'),
           body_p('الأحاديث والآثار مرتبة بحسب أول ورودها، والتخريج هو ما ذكره الكتاب في موضعه.'),
           make_table(HADITHS)]
if H1 is not None:
    a_ = H1
    for e in new_idx:
        a_.addnext(e); a_ = e
else:
    for e in new_idx:
        INDEX_ANCHOR.addprevious(e)
body.remove(INDEX_ANCHOR)
set_text(new_idx[0], 'أولًا: فهرس الآيات القرآنية')
set_text(new_idx[3], 'ثانيًا: فهرس الأحاديث والآثار')
# renumber the ordinal labels of the kept fb index sections
ORDW = ['ثالثًا', 'رابعًا', 'خامسًا', 'سادسًا', 'سابعًا', 'ثامنًا', 'تاسعًا', 'عاشرًا']
k = 0
for e in fb_indexes:
    if e.tag != q('p'):
        continue
    t = text(e)
    m = re.match(r'^(ثالثًا|رابعًا|خامسًا|سادسًا|سابعًا|ثامنًا|تاسعًا): (.*)$', t)
    if m and len(t) < 90 and k < len(ORDW):
        set_text(e, f'{ORDW[k]}: {m.group(2)}')
        k += 1
for e in fb_refs:
    if e.tag == q('p') and re.match(r'^[^:]{2,8}: المصادر والمراجع$', text(e)):
        set_text(e, 'المصادر والمراجع')

# one banner and one heading per part
BANNER = {'موسوعة الإدارة القرآنية', 'الكتاب الثاني: التطبيقات المؤسسية التفصيلية'}
ch1h = next(e for e in body if e.tag == q('p') and text(e) == 'الفصل الأول')
seen_parts, after_ch1 = set(), False
part_tmpl = None
for e in list(body):
    if e is ch1h:
        after_ch1 = True
    if e.tag != q('p') or not after_ch1:
        continue
    t = text(e)
    if t in BANNER:
        body.remove(e)
    elif re.match(r'^الباب (الأول|الثاني|الثالث|الرابع): ', t) and len(t) < 80:
        part_tmpl = part_tmpl if part_tmpl is not None else e
        if t in seen_parts:
            body.remove(e)
        seen_parts.add(t)
intro_part = fresh(copy.deepcopy(part_tmpl))
set_text(intro_part, 'الباب التمهيدي: لماذا يتعثر التطبيق؟ وأين تقف مؤسستك؟')
ch1h.addprevious(intro_part)

# ------------------------------------------------------------------ 9. pages, TOC, page numbers
for p in list(body.iter(q('p'))):
    brs = [b for b in p.iter(q('br')) if b.get(q('type')) == 'page']
    if not brs:
        continue
    if not text(p) and p.getparent() is body and p.find('.//' + q('drawing')) is None:
        body.remove(p)
        continue
    for b in brs:
        b.getparent().remove(b)
def _empty(e):
    return (e is not None and e.tag == q('p') and not text(e) and e.find('.//' + q('drawing')) is None
            and e.find('.//' + q('sectPr')) is None and e.find('.//' + q('instrText')) is None)
for e in list(body):
    if _empty(e) and _empty(e.getprevious()):
        body.remove(e)
PART = re.compile(r'^الباب (التمهيدي|الأول|الثاني|الثالث|الرابع): ')
TOP = re.compile(r'^(الإهداء|كلمة المؤلِّف|كلمة المؤلف|التوطئة|الكتاب الثاني في سياق|المنهج والمصادر|دليل قراءة|خريطة الكتاب|المفاهيم والمصطلحات|الكتاب الثاني والموروث|الخاتمة|الفهارس|(أولًا: )?فهرس الآيات القرآنية|(ثانيًا: )?فهرس الأحاديث والآثار|تاسعًا: المصادر|عاشرًا: المصادر|.*المصادر والمراجع$|ملحق \((أ|ب)\)$)')
TC_LIST = []


def tc(p, label, level):
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
children = [e for e in body if e.tag == q('p')]
for i, p in enumerate(children):
    t = text(p)
    if not t:
        continue
    m = CH_HEAD.match(t)
    if PART.match(t) and len(t) < 80:
        ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        tc(p, t, 1)
    elif m and len(t) < 40:
        if not PART.match(prev_text):
            ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        nxt = next((text(x) for x in children[i + 1:i + 3] if text(x)), '')
        tc(p, f'{t}: {nxt}', 2)
    elif TOP.match(t) and len(t) < 90 and p.getparent() is body:
        if prev_text != 'الفهارس':
            ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)
        tc(p, t, 1)
    elif SEC_HEAD.match(t) and len(t) < 140:
        tc(p, t, 3)
    prev_text = t
# TOC at the front, after the title page
first_front = next(e for e in body if e.tag == q('p') and text(e).startswith('الإهداء'))
toc_head = fresh(copy.deepcopy(first_front))
for r in toc_head.findall(q('r')):
    if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
        toc_head.remove(r)
set_text(toc_head, 'المحتويات')
ensure(ppr(toc_head), 'pageBreakBefore', PPR_ORDER)
L1 = TEMPL['subheading']
L2 = TEMPL['body']
entries = []
for lv, lab in TC_LIST:
    if lv == 3:
        continue                                    # chapter sections are listed inside chapters; TOC keeps 2 levels
    p_ = fresh(copy.deepcopy(L1 if lv == 1 else L2))
    for r in p_.findall(q('r')):
        if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
            p_.remove(r)
    set_text(p_, lab)
    if lv == 2:
        ind = ensure(ppr(p_), 'ind', PPR_ORDER)
        ind.set(q('right'), '567')
        sp = ensure(ppr(p_), 'spacing', PPR_ORDER)
        sp.set(q('after'), '0')
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
first_front.addprevious(toc_head)
for e in entries:
    first_front.addprevious(e)

# footer with page numbers
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
for old in sect.findall(q('footerReference')):
    sect.remove(old)
fr = etree.Element(q('footerReference'))
fr.set(q('type'), 'default')
fr.set('{%s}id' % R_NS, 'rIdFtr1')
sect.insert(0, fr)
SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
              'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']
ensure(sect, 'titlePg', SECT_ORDER)
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
# unique w14 paraIds
seen = set()
for e in body.iter(q('p')):
    pid = e.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
    if pid:
        if pid in seen:
            del e.attrib['{http://schemas.microsoft.com/office/word/2010/wordml}paraId']
        seen.add(pid)

tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
out = os.path.abspath('موسوعة_الإدارة_القرآنية_الكتاب_الثاني_للطباعة.docx')
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-qXr', out, '.'], cwd=OUT, check=True)
json.dump({'stats': stats, 'log': LOG, 'dash': dash_n, 'notbut': nb_n, 'reversed_tables': rev_n, 'ambiguous_tables': amb,
           'figures': fig_no, 'verses': len(VERSES) - 1, 'hadiths': len(HADITHS) - 1, 'sections': len(SECMAP)},
          open('build_log.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('wrote', out, '| figs', fig_no, '| reversed tables', rev_n, 'ambiguous', amb, '| verses', len(VERSES) - 1, '| hadiths', len(HADITHS) - 1)
print(stats)
