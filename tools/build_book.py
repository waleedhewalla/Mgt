"""Print edition of Book One, built on the author's combined file.

1. Review: restore chapter 13, unify style (dashes, «ليس… بل»), update cross-references to the new layout.
2. Every table right-to-left (bidiVisual) with RTL cell paragraphs.
3. Chapter reference lists removed; one reference list at the end of the book.
4. A short introduction (goal and content) at the head of each chapter.
Print set-up: table of contents at the front (Word TOC field), each part on a new page, page numbers.
"""
import sys, os, re, json, glob, shutil, subprocess, copy, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxedit import *

UP = '/root/.claude/uploads/be2474ec-46f6-5fbc-9b19-f8201f104606'
SRC_DOCX = glob.glob(f'{UP}/a88d261d*.docx')[0]
CH13 = '/home/user/Mgt/chapters/الفصل_الثالث_عشر_النهائي.docx'
INDEXES = 'الفهارس_النهائية.docx'
OUT = 'book_build'
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT)
with zipfile.ZipFile(SRC_DOCX) as z:
    z.extractall(OUT)

tree = etree.parse(f'{OUT}/word/document.xml')
body = tree.getroot().find(q('body'))
sect = body.find(q('sectPr'))
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def kids():
    return [e for e in body if e.tag != q('sectPr')]


def text(e):
    return ptext(e).strip()


def find(pred, start=0):
    ks = kids()
    for i in range(start, len(ks)):
        if pred(ks[i]):
            return ks[i]
    return None


def is_para(e, s):
    return e.tag == q('p') and text(e) == s


def body_of(path):
    x = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    return [fresh(copy.deepcopy(e)) for e in x.find(q('body')) if e.tag != q('sectPr')]


PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
             'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing',
               'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
             'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
             'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']


def ensure(parent, tag, order):
    """Return child <tag>, inserting it at its schema position if missing."""
    el = parent.find(q(tag))
    if el is not None:
        return el
    el = etree.Element(q(tag))
    pos = order.index(tag)
    for i, ch in enumerate(parent):
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


# ================================================================== 1. restore chapter 13
ch14 = find(lambda e: is_para(e, 'الفصل الرابع عشر'))
prev = ch14.getprevious()
anchor = prev if (prev is not None and not text(prev) and prev.find('.//' + q('br')) is not None) else ch14
for e in body_of(CH13):
    anchor.addprevious(e)

# ================================================================== 3b. indexes and references move to the end
idx_start = find(lambda e: is_para(e, 'الفهارس'))
app_title = None
for e in kids():
    if is_para(e, 'الملاحق') and e.getnext() is not None and 'أدوات عملية' in text(e.getnext()):
        app_title = e
e = idx_start
while e is not app_title:
    nxt = e.getnext()
    body.remove(e)
    e = nxt
for e in body_of(INDEXES):
    sect.addprevious(e)

# ================================================================== 3a. remove chapter reference lists
REF_HEAD = re.compile(r'^(\d+\.\d+\s+)?قائمة المراجع المذكورة في هذا الفصل$')
CH_HEAD = re.compile(r'^الفصل (الأول|الثاني|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر|الحادي عشر|الثاني عشر|الثالث عشر|الرابع عشر)$')
removed_refs = 0
for head in [e for e in kids() if e.tag == q('p') and REF_HEAD.match(text(e))]:
    e = head
    while e is not None and e.tag != q('sectPr'):
        t = text(e)
        if CH_HEAD.match(t) or t == 'خاتمة الكتاب الأول':
            break
        if not t and e.find('.//' + q('br')) is not None:      # page break before the next chapter
            break
        nxt = e.getnext()
        body.remove(e)
        removed_refs += 1
        e = nxt

# ================================================================== page layout: one mechanism for new pages
for p in list(body.iter(q('p'))):
    brs = [b for b in p.iter(q('br')) if b.get(q('type')) == 'page']
    if not brs:
        continue
    if not text(p) and p.getparent() is body:
        body.remove(p)                                            # page-break-only paragraph
        continue
    for b in brs:
        r = b.getparent()
        r.remove(b)
        if not ''.join(r.itertext()).strip() and r.find(q('t')) is None:
            r.getparent().remove(r)

APP_HEAD = re.compile(r'^الملحق (الأول|الثاني|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر|الحادي عشر|الثاني عشر): ')
new_page = [find(lambda e: is_para(e, 'المقدمة'))]
new_page += [e for e in kids() if e.tag == q('p') and CH_HEAD.match(text(e))]
new_page += [find(lambda e: is_para(e, 'خاتمة الكتاب الأول')), app_title]
new_page += [e for e in kids() if e.tag == q('p') and APP_HEAD.match(text(e))]
new_page += [find(lambda e: is_para(e, t)) for t in ('الفهارس', 'فهرس الأحاديث والآثار', 'قائمة المراجع')]
for p in new_page:
    ensure(ppr(p), 'pageBreakBefore', PPR_ORDER)

# ================================================================== 4. chapter introductions
INTROS = {
    'الأول': 'يحدد هذا الفصل ما يسعى إليه الكتاب وما لا يدّعيه. يبدأ بالسؤال المؤسِّس، ثم يعرض الإطار الابستيمولوجي الذي يعمل داخله: ما نسلّم به وما '
             'نختبره، وأقومية القرآن، وهرم المصادر في صورته الكلية. ثم يشخّص مشكلة القولبة، ويبيّن لماذا يبدو المشروع ممكنًا الآن، ويختم بحدود الادعاء.',
    'الثاني': 'يحدد هذا الفصل موقع المشروع بين خمسة تيارات: الفكر الإداري الغربي الحديث، والتجربة الإسلامية التاريخية، والإدارة الإسلامية المعاصرة، '
              'ومشروع أسلمة المعرفة، ثم هذا المشروع. ويشرح سبعة أسباب لتعثر الأسلمة في الإدارة، ويعمّق تشخيص القولبة، ويعرض النقد الداخلي للإدارة '
              'الغربية، ثم يختبر ذلك في حالة التمويل الإسلامي.',
    'الثالث': 'يضع هذا الفصل للمشروع شرطًا يمكن أن يُحكم عليه به بالفشل: الفرضية الصفرية. يعرض صياغتها ومحكّاتها الثلاثة (الأصالة، والاستقلالية، '
              'والتمييز) وقائمة النفي الثمانية، ثم يطبقها على أربعة أمثلة. ويصف بروتوكول الاختبار التجريبي وجدول القبول والرفض، ويقرر قاعدة فض '
              'التعارض بين النص والتجربة، ويواجه ثلاثة اعتراضات على المنهجية كلها.',
    'الرابع': 'يرتّب هذا الفصل المصادر التي يستدل بها الكتاب: القرآن دستورًا، والسنة بيانًا لا منشأً، وخمس مدارس تفسيرية بمعايير اختيارها، ثم مصادر '
              'السيرة بحسب ثبوتها، والمصادر اللغوية. ويعالج مسألة الاستفادة من الطباطبائي، ويطبق الهرم كاملًا على مفهوم العدل، ثم يقارنه بنظريات '
              'السلطة المعرفية والشهادة الغربية، ويناقش أربعة اعتراضات.',
    'الخامس': 'يعرض هذا الفصل الأداة المركزية في الكتاب: سلم الاشتقاق السباعي، من الدلالة القرآنية إلى التطبيق، وبين مستوييه الرابع والخامس بوابة '
              'المقاصد. ويؤصّله في منهج الاستنباط، ويحدد معايير الترقية والرفض، ثم يطبقه كاملًا على الشورى والأمانة. ويقدّم بطاقة الهوية المعرفية '
              'بحقولها الخمسة عشر، ويقارن السلم بمنهجيات بناء النظرية الغربية.',
    'السادس': 'يحدد هذا الفصل كيف يُقرأ النص قبل دخول السلم: ثماني خطوات إلزامية، من السياق واللغة إلى إعلان المستوى. ويقف عند المحكم والمتشابه '
              'والمثاني، وخطوات الاستقراء الموضوعي، ثم يبيّن الخطوات الاختيارية والمحظورات الصريحة. ويطبق البروتوكول كاملًا على آية الإحسان، '
              'ويقارنه بنظريات صحة التأويل الغربية، وينتهي بقائمة تحقق للاستخدام اليومي.',
    'السابع': 'يجيب هذا الفصل عن سؤال يطرحه كل ناقد جاد: ماذا نفعل حين يخالف الواقعُ النموذج؟ ينطلق من إشكالية دوهيم-كواين، ويؤصّل الجواب في باب '
              'تعارض الأدلة، ثم يصنّف التعارض أربعة أنواع، ويقدّم خوارزمية للقرار مع حد التدهور. ويطبقها على ثلاث حالات، ويقارنها بكون ولاكاتوش، '
              'ويبيّن كيف يتجنب فخ اللاقابلية للتكذيب.',
    'الثامن': 'يفتتح هذا الفصل الباب الثاني بسؤال: من الإنسان الذي يُدار؟ ويشتق بالبروتوكول الكامل ثلاثة أركان: الاستخلاف، والتكليف بالوسع، '
              'والابتلاء، ثم يركّبها في تصور واحد. ويطبقها على قرار تسريح جماعي وعلى نظام لتقييم الأداء، ثم يقارنها بنظرية أصحاب المصلحة وهرم '
              'ماسلو والأنثروبولوجيات الأربع، ويواجه خمسة اعتراضات.',
    'التاسع': 'ينتقل هذا الفصل من الإنسان إلى معرفته: كيف يعرف المدير ما يبني عليه قراره؟ ويشتق ثلاثة أركان: الفطرة إشارةً أخلاقية أولى، والتثبت '
              'تحققًا من الخبر، والتدبر نظرًا في العواقب. ثم يركّبها في نظرية واحدة، ويطبقها على قرار استراتيجي وعلى الإبلاغ عن مخالفة، ويقارنها '
              'بسايمون وكانمان والمدارس الإبستيمولوجية الأربع.',
    'العاشر': 'يسأل هذا الفصل: بأي ميزان يتصرف المدير؟ ويشتق ثلاث قيم: الرحمة الفاعلة، والحكمة التطبيقية، والصدق، ثم يجمعها مع الأمانة والشورى '
              'والإحسان والعدل في منظومة سباعية، ويختبر تكاملها مع أركان الإنسان والمعرفة. ويطبقها على ثلاث حالات، ويعالج تعارض الرحمة والعدل، '
              'ويقارن المنظومة بالنظريات الأخلاقية الغربية الكبرى.',
    'الحادي عشر': 'يجمع هذا الفصل ما بُني في الفصول السابقة في بنية واحدة: الإطار الإداري القرآني QMF بمكوّناته الثمانية، من الرؤية الكلية إلى '
                  'الأثر والمساءلة. ويعرّف كل مكوّن ومصدره، ويفصّل الرؤية والغاية والمؤسسة، ويضع السلطة داخل المساءلة. ثم يطبق الإطار على مثالين، '
                  'ويقارنه بأربعة أطر بنائية غربية، ويختبر اتساق مستويات ادعائه.',
    'الثاني عشر': 'ينقل هذا الفصل الإطار من البنية إلى الحركة: كيف يُتخذ القرار؟ ويعرض الدورة السباعية من الاستشعار إلى المتابعة والمحاسبة، مع '
                  'تعميق مراحل الاستشعار والعزم والتوكل والمتابعة. ويقدّم سجل القرار وقاعدة لاختصار الدورة حين يضيق الوقت، ثم يطبقها على حالتين، '
                  'ويقارنها بأربعة نماذج قرار غربية.',
    'الثالث عشر': 'يسأل هذا الفصل: كيف نعرف أن نموذجًا ما قرآني فعلًا، وكيف نطوّره دون أن يفقد مرجعيته؟ ويعرض تسعة معايير، أربعة منها إلزامية، '
                  'ويبيّن حالات العلاقة مع الأدبيات الإدارية: التوافق والاستقلال والتعارض. ويقرر أن فشل التطبيق يُراجَع فيه النموذج البشري لا النص، '
                  'ثم يختبر نموذجين مقترحين، ويطبق المعايير على نفسها.',
    'الرابع عشر': 'يعرض هذا الفصل الجهاز المنهجي كاملًا في ثماني حالات مركبة، افتراضية الأسماء واقعية السياق: أربع من القطاع الخاص، وأربع من ميادين '
                  'الحياة الشخصية والأعمال الخاصة والمشاريع. وفي كل حالة المعضلة، والأركان الأوثق صلة، والدورة السباعية، والقرار، وما يبقى من كلفة. '
                  'ثم يقارن الحالات، ويقدّم دليلًا عمليًا وأجندة بحثية.',
}
TEMPLATE = find(lambda e: e.tag == q('p') and text(e).startswith('أقومية القرآن:'))     # shaded label + text paragraph
VERSE_REF = re.compile(r'^[ء-ي ]+\s?:\s?\d+$')
BLANK = find(lambda e: e.tag == q('p') and not text(e) and e.find('.//' + q('br')) is None and e.find('.//' + q('drawing')) is None)
chapter_titles = {}
for head in [e for e in kids() if e.tag == q('p') and CH_HEAD.match(text(e))]:
    name = CH_HEAD.match(text(head)).group(1)
    t1, t2 = head.getnext(), head.getnext().getnext()
    chapter_titles[name] = (text(t1), text(t2))
    e, spot = head, None
    for _ in range(8):
        e = e.getnext()
        if e.tag == q('p') and VERSE_REF.match(text(e)):
            spot = e
            break
    spot = spot if spot is not None else t2
    intro = fresh(copy.deepcopy(TEMPLATE))
    runs = [r for r in intro.findall(q('r')) if ''.join(r.itertext()).strip()]
    set_runs(intro, [('مدخل الفصل: ', runs[0]), (INTROS[name], runs[1])])
    spot.addnext(intro)
    intro.addprevious(fresh(copy.deepcopy(BLANK)))

# ================================================================== 1b. text updates for the new layout
UPDATES = {
    'وبعد الفصول تأتي خاتمة الكتاب، ثم الملاحق والفهارس.': 'وبعد الفصول تأتي خاتمة الكتاب، ثم الملاحق، ففهرسا الآيات والأحاديث، ثم قائمة المراجع.',
}
for p in body.iter(q('p')):
    if text(p) in UPDATES:
        set_text(p, UPDATES[text(p)])
for t in body.iter(q('t')):
    if t.text and 'وتجمع القائمة الموحدة في الفهارس مراجع الفصول كلها' in t.text:
        t.text = t.text.replace('وتجمع القائمة الموحدة في الفهارس مراجع الفصول كلها', 'وتجمع قائمة المراجع في آخر الكتاب مراجع الفصول كلها')

# ================================================================== 1c. style: long dashes and «ليس… بل»
CAPTION = re.compile(r'^(جدول|شكل)\s')
NOT_BUT = re.compile(r'((?:^|[\s(«])(?:و|ف)?(?:ليس|ليست|ليسوا)\s[^.؟!؛:]{1,80}?)،?\s+بل\s')
dash_n = notbut_n = 0


def fix_sentence_dashes(s, sep_single):
    out = []
    for sent in re.split(r'(?<=[.؟!])\s', s):
        n = sent.count('—')
        if n >= 2:
            sent = re.sub(r'\s*—\s*', '، ', sent)
        elif n == 1:
            sent = re.sub(r'\s*\(([^)]*?)\s*—\s*([^)]*?)\)', r' (\1، \2)', sent)
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
        if '—' in s and not caption and full != '—':
            s = fix_sentence_dashes(s, sep)
            if s.startswith((':', '؛', '،')) and n > 0:
                ts[n - 1].text = ts[n - 1].text.rstrip()
            if s.endswith(' ') and n + 1 < len(ts):
                ts[n + 1].text = ts[n + 1].text.lstrip()
        s2 = NOT_BUT.sub(r'\1، وإنما ', s)
        notbut_n += s2 != s
        dash_n += s0.count('—') - s2.count('—')
        s2 = s2.replace('؛ ،', '؛').replace('، ،', '،').replace('  ', ' ')
        t.text = s2

# ================================================================== 2. tables right-to-left
tables = list(body.iter(q('tbl')))
for tbl in tables:
    tp = tbl.find(q('tblPr'))
    if tp is None:
        tp = etree.Element(q('tblPr'))
        tbl.insert(0, tp)
    ensure(tp, 'bidiVisual', TBLPR_ORDER)
    for p in tbl.iter(q('p')):
        if re.search('[؀-ۿ]', ptext(p)):
            ensure(ppr(p), 'bidi', PPR_ORDER)

# ================================================================== table of contents at the front (Word field) + TC entries
TOC = json.load(open('toc.json', encoding='utf-8'))


TC_LIST = []


def tc_field(p, label, level):
    TC_LIST.append((level - 1, label))
    runs = []
    for kind in ('begin', None, 'end'):
        r = etree.Element(q('r'))
        if kind:
            f = etree.SubElement(r, q('fldChar'))
            f.set(q('fldCharType'), kind)
        else:
            it = etree.SubElement(r, q('instrText'))
            it.set(XML_SPACE, 'preserve')
            it.text = f' TC "{label}" \\l {level} '
        runs.append(r)
    anchor_run = p.findall(q('r'))[-1] if p.findall(q('r')) else None
    for r in runs:
        if anchor_run is not None:
            anchor_run.addnext(r)
            anchor_run = r
        else:
            p.append(r)


ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'الثاني عشر',
       'الثالث عشر', 'الرابع عشر']
INTRO_HEADS = {t for lv, t in TOC[:12] if lv == 1}
level1 = {t for lv, t in TOC if lv == 1}
in_intro = True
for p in list(body.iter(q('p'))):
    if p.getparent() is not body:
        continue
    t = text(p)
    if t == 'المقدمة':
        tc_field(p, 'المقدمة', 1)
    elif CH_HEAD.match(t):
        in_intro = False
        name = CH_HEAD.match(t).group(1)
        a, b = text(p.getnext()), text(p.getnext().getnext())
        sep_ = ' ' if a.endswith(('؟', '?')) else '، '
        tc_field(p, f'{t}: {a}{sep_}{b}', 1)
    elif t == 'خاتمة الكتاب الأول':
        tc_field(p, 'خاتمة الكتاب الأول: من السؤال إلى المنظومة', 1)
    elif p is app_title:
        tc_field(p, 'الملاحق', 1)
    elif t in ('الفهارس', 'قائمة المراجع'):
        tc_field(p, t, 1)
    elif re.match(r'^\d+\.\d+\s', t) and len(t) < 130:
        tc_field(p, t, 2)
    elif t in level1 and (APP_HEAD.match(t) or t in ('فهرس الآيات القرآنية', 'فهرس الأحاديث والآثار') or t in INTRO_HEADS
                          or t in {x for lv, x in TOC if lv == 1}):
        tc_field(p, t, 2)

intro_title = find(lambda e: is_para(e, 'المقدمة'))
toc_head = fresh(copy.deepcopy(intro_title))
for r in toc_head.findall(q('r')):
    if r.find(q('instrText')) is not None or r.find(q('fldChar')) is not None:
        toc_head.remove(r)
set_text(toc_head, 'المحتويات')
L0 = find(lambda e: e.tag == q('p') and text(e).startswith('الباب التمهيدي: المشروع وموقعه'))
L1 = find(lambda e: e.tag == q('p') and text(e).startswith('منذ أكثر من مئة عام'))
entries = []
for lv, t in TC_LIST:
    p = fresh(copy.deepcopy(L0 if lv == 0 else L1))
    for r in p.findall(q('r')):
        if r.find(q('fldChar')) is not None or r.find(q('instrText')) is not None:
            p.remove(r)
    set_text(p, t)
    pp = ppr(p)
    for tag in ('pBdr', 'shd'):
        el = pp.find(q(tag))
        if el is not None:
            pp.remove(el)
    if lv == 1:
        ind = ensure(pp, 'ind', PPR_ORDER)
        ind.set(q('right'), '567')
        sp = ensure(pp, 'spacing', PPR_ORDER)
        sp.set(q('after'), '0')
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
first.insert(1 if first.find(q('pPr')) is not None else 0, fld_run('begin'))
first.insert(2 if first.find(q('pPr')) is not None else 1, fld_run(instr=' TOC \\f \\l "1-2" \\z '))
first.insert(3 if first.find(q('pPr')) is not None else 2, fld_run('separate'))
entries[-1].append(fld_run('end'))
ensure(ppr(toc_head), 'pageBreakBefore', PPR_ORDER)
intro_title.addprevious(toc_head)
for p in entries:
    intro_title.addprevious(p)

# ================================================================== page numbers
footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:ftr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:pPr><w:bidi/><w:jc w:val="center"/></w:pPr>'
          '<w:r><w:rPr><w:rFonts w:ascii="Sakkal Majalla" w:hAnsi="Sakkal Majalla" w:cs="Sakkal Majalla"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
          '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>1</w:t></w:r>'
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')
open(f'{OUT}/word/footer1.xml', 'w', encoding='utf-8').write(footer)
rels_p = f'{OUT}/word/_rels/document.xml.rels'
rels = open(rels_p, encoding='utf-8').read()
if 'footer1.xml' not in rels:
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
                                            'Target="footer1.xml"/></Relationships>')
    open(rels_p, 'w', encoding='utf-8').write(rels)
ct_p = f'{OUT}/[Content_Types].xml'
ct = open(ct_p, encoding='utf-8').read()
if 'footer1.xml' not in ct:
    ct = ct.replace('</Types>', '<Override PartName="/word/footer1.xml" '
                                'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
    open(ct_p, 'w', encoding='utf-8').write(ct)
for old in sect.findall(q('footerReference')):
    sect.remove(old)
fr = etree.Element(q('footerReference'))
fr.set(q('type'), 'default')
fr.set('{%s}id' % R, 'rIdFtr1')
sect.insert(0, fr)
SECT_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
              'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings']
ensure(sect, 'titlePg', SECT_ORDER)

# settings: ask Word to refresh the TOC and page numbers on open
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

# whitespace preservation for edited runs
for t in body.iter(q('t')):
    if t.text and t.text != t.text.strip():
        t.set(XML_SPACE, 'preserve')

tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
out = os.path.abspath('موسوعة_الإدارة_القرآنية_الكتاب_الأول_للطباعة.docx')
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-qXr', out, '.'], cwd=OUT, check=True)
print('wrote', out, '| ref elements removed', removed_refs, '| dashes fixed', dash_n, '| ليس…بل', notbut_n, '| tables', len(tables))
