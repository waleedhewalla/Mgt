"""Combine the three Book Three files into one base document and write ID-tagged dumps for review."""
import zipfile, re, json, copy
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
pt = lambda e: ''.join(t.text or '' for t in e.iter(W + 't')).strip()

def sz(e):
    for r in e.findall(W + 'r'):
        v = r.find(W + 'rPr/' + W + 'sz')
        if v is not None and ''.join(t.text or '' for t in r.iter(W + 't')).strip():
            return int(v.get(W + 'val'))
    return 0

docs = []
for f in ['f1', 'f2', 'f3']:
    x = etree.fromstring(zipfile.ZipFile(f'src/{f}_fixed.docx').read('word/document.xml'))
    docs.append(x)
base = docs[0]
bbody = base.find(W + 'body')
sect = bbody.find(W + 'sectPr')
for x in docs[1:]:
    for e in list(x.find(W + 'body')):
        if e.tag != W + 'sectPr':
            sect.addprevious(e)
for s in list(bbody.iter(W + 'sectPr')):          # section breaks inside paragraphs: one section for the book
    if s is not sect:
        s.getparent().remove(s)
for br in list(bbody.iter(W + 'br')):             # hard page breaks: the build sets page breaks itself
    if br.get(W + 'type') == 'page':
        br.getparent().remove(br)
z = zipfile.ZipFile('src/f1_fixed.docx')
with zipfile.ZipFile('base.docx', 'w', zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        o.writestr(it, etree.tostring(base, xml_declaration=True, encoding='UTF-8', standalone=True)
                   if it.filename == 'word/document.xml' else z.read(it.filename))

body = [e for e in bbody if e.tag != W + 'sectPr']
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر',
       'الثالث عشر', 'الرابع عشر']
CH = re.compile(r'^الفصل (' + '|'.join(sorted(ORD, key=len, reverse=True)) + r')$')
marks, mech = {}, set()
f2_start = None
for i, e in enumerate(body):
    t = pt(e)
    if e.tag != W + 'p':
        continue
    m = CH.match(t)
    if m and sz(e) >= 44:
        marks['ch%02d' % (ORD.index(m.group(1)) + 1)] = i
    if t == 'فهرس الفصل الأول':
        marks['ch01'] = i
    if t.startswith('تقديم — الكتاب الثالث'):
        marks['front'] = i
    if t.startswith('الخاتمة الكبرى'):
        marks['concl'] = i
    if t.startswith('الملحق أ:'):
        marks['apps'] = i
    if t == 'بسم الله الرحمن الرحيم' and f2_start is None and i > 100:
        f2_start = i
# mechanical removals: title pages and static tables of contents of the source files
for i in range(0, marks['front']):
    mech.add(i)
toc = next(i for i, e in enumerate(body) if pt(e) == 'فهرس المحتويات')
for i in range(toc, marks['ch01'] + 2):          # also the first chapter's own contents table
    mech.add(i)
for i in range(f2_start, marks['ch06']):
    mech.add(i)
# expert-committee (WRKOH) review sections, the external-review section, and per-chapter reference lists
order = sorted(marks.items(), key=lambda kv: kv[1])
def sec_end(i):
    for j in range(i + 1, len(body)):
        if body[j].tag == W + 'p' and pt(body[j]) and sz(body[j]) >= 34:
            return j
    return len(body)
for i, e in enumerate(body):
    t = pt(e)
    if e.tag == W + 'p' and sz(e) >= 34 and t and (
            'WRKOH' in t or t.startswith('المراجعة الخارجية المتعددة') or re.search(r'(المراجع والمصادر|المصادر القرآنية والتفسيرية)$', t)):
        for j in range(i, sec_end(i)):
            mech.add(j)
        # chapter reference lists run to the next chapter heading
        if 'المراجع' in t or 'المصادر' in t:
            j = sec_end(i)
            while j < len(body) and not (body[j].tag == W + 'p' and sz(body[j]) >= 44) and not (pt(body[j]).startswith(('الخاتمة', 'الملحق'))):
                if body[j].tag == W + 'p' and sz(body[j]) >= 34 and not re.search(r'(أولاً|ثانياً|ثالثاً|رابعاً|خامساً)', pt(body[j])[:8]):
                    break
                mech.add(j)
                j += 1
mech -= {250, 251, 252}           # chapter 2 closing verse stays
json.dump({'marks': order, 'mech': sorted(mech)}, open('base_meta.json', 'w'))
print(order, 'mechanical removals', len(mech))

def dump(a, b):
    out = []
    for i in range(a, b):
        e = body[i]
        if i in mech:
            if out and out[-1].startswith('[حُذف آليًا'):
                continue
            out.append('[حُذف آليًا: صفحات العنوان والفهارس القديمة، ومراجعات اللجان، وقوائم مراجع الفصول (تُجمع في آخر الكتاب)]')
            continue
        if e.tag == W + 'p':
            t = pt(e)
            if t:
                s = sz(e)
                tag = '[عنوان فصل] ' if s >= 44 else '[عنوان قسم] ' if s >= 34 else '[عنوان فرعي] ' if s >= 30 else ''
                out.append(f'P{i}\t{tag}{t}')
        elif e.tag == W + 'tbl':
            out.append(f'TBL{i}\t[جدول]')
            for r, tr in enumerate(e.findall(W + 'tr')):
                for c, tc in enumerate(tr.findall(W + 'tc')):
                    out.append(f'TP{i}.{r}.{c}\t' + ' / '.join(pt(p) for p in tc.findall(W + 'p') if pt(p)))
    return out
for k, (name, a) in enumerate(order):
    b = order[k + 1][1] if k + 1 < len(order) else len(body)
    open(f'dumps/{name}.txt', 'w').write('\n'.join(dump(a, b)))
