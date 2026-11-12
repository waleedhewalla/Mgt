"""Add an imprint page (after the title page) and, for التأسيس المعرفي, the scope note in the Introduction.

usage: python3 front_matter.py <src.docx> <out.docx> <key>   (key: b1 b2 b3 b4 ta)
Run AFTER the reviewed edits are applied (it inserts paragraphs, which shifts paragraph IDs).
"""
import sys, copy
import docx
from docx.oxml.ns import qn

XS = '{http://www.w3.org/XML/1998/namespace}space'

TITLES = {
    'b1': ('موسوعة الإدارة القرآنية', 'الكتاب الأول: الأسس الوجودية والمعرفية والمنهجية للإدارة القرآنية'),
    'b2': ('موسوعة الإدارة القرآنية', 'الكتاب الثاني: التطبيقات المؤسسية التفصيلية'),
    'b3': ('موسوعة الإدارة القرآنية', 'الكتاب الثالث: التطبيقات القطاعية المتخصصة'),
    'b4': ('موسوعة الإدارة القرآنية', 'الكتاب الرابع: الإدارة النبوية في السيرة'),
    'ta': ('سلسلة المعرفة الحضارية', 'التأسيس المعرفي: القرآن الكريم وبناء العلوم'),
}

def lines(key):
    series, title = TITLES[key]
    return [
        (series, True),
        (title, True),
        ('تأليف: م. وليد عبدالله حواله', False),
        ('الطبعة الأولى: 1448هـ / 2026م', False),
        ('جميع الحقوق محفوظة للمؤلف. لا يجوز نسخ أي جزء من هذا الكتاب أو نقله أو تخزينه بأي وسيلة دون إذن كتابي مسبق من المؤلف، '
         'إلا الاقتباس القصير لأغراض البحث والنقد مع ذكر المصدر.', False),
        ('رقم الإيداع: ........................     الترقيم الدولي (ISBN): ........................', False),
        ('طوبقت الآيات القرآنية على نص المصحف برواية حفص عن عاصم، وطوبقت الأحاديث على الكتب الستة وموطأ مالك بترقيم طبعاتها '
         'المتداولة. والآراء الواردة في الكتاب اجتهاد مؤلفه.', False),
    ]

SCOPE = ('حدود العنوان. لا تعني عبارة «بناء العلوم» في العنوان أن الكتاب يبني علمًا بعينه، بل أنه يقترح إجراءً معلنًا للتمييز '
         'بين أنحاء حضور النص في بنية العلم، إجراءً قابلًا للتخطئة يُختبر في كل علم على حدة. والحالة الإدارية في الفصل العشرين '
         'مثال تطبيقي واحد على هذا الإجراء، لا برهان على صحته.')


def first_run_rpr(p):
    r = p.find(qn('w:r'))
    return copy.deepcopy(r.find(qn('w:rPr'))) if r is not None and r.find(qn('w:rPr')) is not None else None


def make_par(tmpl_ppr, rpr, text, bold, size, page_break=False, space_after=120):
    p = docx.oxml.parse_xml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    ppr = copy.deepcopy(tmpl_ppr)
    for tag in ('w:pageBreakBefore', 'w:shd', 'w:pBdr', 'w:spacing'):
        for e in ppr.findall(qn(tag)):
            ppr.remove(e)
    sp = ppr.makeelement(qn('w:spacing'), {qn('w:before'): '0', qn('w:after'): str(space_after)})
    bidi = ppr.find(qn('w:bidi'))
    (bidi.addnext(sp) if bidi is not None else ppr.insert(0, sp))
    if page_break:
        ppr.insert(0, ppr.makeelement(qn('w:pageBreakBefore'), {}))
    p.append(ppr)
    r = p.makeelement(qn('w:r'), {})
    rp = copy.deepcopy(rpr)
    for tag in ('w:b', 'w:bCs', 'w:sz', 'w:szCs', 'w:color', 'w:i', 'w:iCs'):
        for e in rp.findall(qn(tag)):
            rp.remove(e)
    pos = 1 if rp.find(qn('w:rFonts')) is not None else 0
    if bold:
        rp.insert(pos, rp.makeelement(qn('w:bCs'), {})); rp.insert(pos, rp.makeelement(qn('w:b'), {}))
    sz = rp.makeelement(qn('w:sz'), {qn('w:val'): str(size)}); szc = rp.makeelement(qn('w:szCs'), {qn('w:val'): str(size)})
    col = rp.makeelement(qn('w:color'), {qn('w:val'): '404040'})
    rtl = rp.find(qn('w:rtl'))
    for e in (col, sz, szc):
        (rtl.addprevious(e) if rtl is not None else rp.append(e))
    r.append(rp)
    t = r.makeelement(qn('w:t'), {}); t.text = text; t.set(XS, 'preserve'); r.append(t)
    p.append(r)
    return p


def main():
    src, out, key = sys.argv[1:4]
    d = docx.Document(src)
    P = [e for e in d.element.body.iterchildren() if e.tag.endswith('}p')]
    txt = lambda e: ''.join(t.text or '' for t in e.iter(qn('w:t'))).strip()
    contents = next(e for e in P if txt(e) in ('المحتويات', 'فهرس المحتويات'))
    body_tmpl = next(e for e in P[P.index(contents):] if len(txt(e)) > 200 and e.find(qn('w:r')) is not None)
    rpr = first_run_rpr(body_tmpl)
    tmpl_ppr = copy.deepcopy(contents.find(qn('w:pPr')))
    jc = tmpl_ppr.find(qn('w:jc'))
    if jc is None:
        jc = tmpl_ppr.makeelement(qn('w:jc'), {})
        tmpl_ppr.append(jc)
    jc.set(qn('w:val'), 'center')
    for i, (text, bold) in enumerate(lines(key)):
        contents.addprevious(make_par(tmpl_ppr, rpr, text, bold, 26 if bold else 22, page_break=(i == 0),
                                      space_after=240 if i in (1, 3, 4, 5) else 80))
    if key == 'ta':
        anchor = next(e for e in P if txt(e).startswith('ما أطلبه من القارئ'))
        new = copy.deepcopy(anchor)
        for r in new.findall(qn('w:r'))[1:]:
            new.remove(r)
        r0 = new.find(qn('w:r'))
        for t in r0.findall(qn('w:t')):
            r0.remove(t)
        t = r0.makeelement(qn('w:t'), {}); t.text = SCOPE; t.set(XS, 'preserve'); r0.append(t)
        anchor.addprevious(new)
    d.save(out)
    print('front matter added', key)


if __name__ == '__main__':
    main()
