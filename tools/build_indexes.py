"""Indexes of Book One: contents, Qur'anic verses, hadiths and athar, unified references.
Generated from the final chapter files and the introduction/conclusion file, inside a copy of the manuscript."""
import sys, os, re, json, shutil, subprocess, copy, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxedit import *
from surahs import SURAHS

CHAPTERS = '/home/user/Mgt/chapters'
NAMES = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر',
         'الحادي_عشر', 'الثاني_عشر', 'الثالث_عشر', 'الرابع_عشر']
INTRO = 'المقدمة_والخاتمة_النهائي.docx'
APPENDIX_TITLES = [
    'الملحق الأول: خريطة الكتاب',
    'الملحق الثاني: معجم الأركان الثلاثة عشر',
    'الملحق الثالث: ثمانية نماذج عمل',
    'الملحق الرابع: عشرة مبادئ عملية',
    'الملحق الخامس: التشخيص الذاتي لمراحل التطبيق',
    'الملحق السادس: خطة تسعين يومًا',
    'الملحق السابع: ثلاثة سيناريوهات تطبيقية',
    'الملحق الثامن: أسئلة شائعة',
    'الملحق التاسع: دليل المدرب',
    'الملحق العاشر: متابعة الأثر',
    'الملحق الحادي عشر: مصفوفة الأركان والسياقات',
    'الملحق الثاني عشر: معجم المصطلحات',
]


def q_(t):
    return q(t)


def units(path):
    """Paragraph texts in document order; table cells flattened. Yields (text, is_body_para)."""
    x = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    for e in x.find(q('body')):
        if e.tag == q('p'):
            yield ptext(e).strip(), True
        elif e.tag == q('tbl'):
            for p in e.iter(q('p')):
                s = ptext(p).strip()
                if s:
                    yield s, False


# ------------------------------------------------------------------ collect
FILES = [('intro', INTRO)] + [(str(k), f'{CHAPTERS}/الفصل_{n}_النهائي.docx') for k, n in enumerate(NAMES, 1)]
U = []                       # (chapter, section, text)
TOC = []                     # (level, text)   level 0 = part/chapter title, 1 = section
INTRO_HEADS = {'سؤال قبل الأسئلة', 'التشخيص: ثلاث مفارقات في حقل الإدارة الإسلامية', 'لماذا الآن؟', 'ما يقدّمه هذا الكتاب وما لا يدّعيه',
               'ثلاث أدوات في قلب المنهج', 'خارطة الكتاب: أربعة أبواب وأربعة عشر فصلًا', 'لمن يُكتب هذا الكتاب؟', 'بين يدي الكتاب: كلمة من المؤلف',
               'ملحق المقدمة: خريطة المفاهيم الأساسية'}
CONCL_HEADS = {'السؤال الذي بدأنا به', 'الرحلة في أربعة أبواب', 'ما أنجزه الكتاب وما لم ينجزه', 'ما لم يثبته الكتاب',
               'ما كان يمكن أن أفعله بطريقة أخرى', 'ما بعد هذا الكتاب', 'كيف يُقرأ الكتاب بحسب القارئ', 'كلمة أخيرة'}
concl_toc = []
REFSECS = set()
for ch, path in FILES:
    sec = None
    texts = list(units(path))
    if ch == 'intro':
        TOC.append((0, 'المقدمة'))
    else:
        body_ps = [t for t, isp in texts if t and isp]
        sep = ' ' if body_ps[1].endswith(('؟', '?')) else '، '
        TOC.append((0, f'{body_ps[0]}: {body_ps[1]}{sep}{body_ps[2]}'.replace(' — ', '، ')))
    for s, isp in texts:
        if not s:
            continue
        if ch == 'intro' and s == 'خاتمة الكتاب الأول':
            ch = 'concl'
        m = re.match(r'^(\d+\.\d+)\s', s)
        if m and len(s) < 130 and isp:
            sec = m.group(1)
            if 'قائمة المراجع' in s:          # chapter reference lists are merged into the book's list
                REFSECS.add(sec)
            else:
                TOC.append((1, s))
        if isp and ch == 'intro' and s in INTRO_HEADS:
            TOC.append((1, s))
        if isp and ch == 'concl' and s in CONCL_HEADS:
            concl_toc.append((1, s))
        U.append((ch, sec, s))
TOC.append((0, 'خاتمة الكتاب الأول: من السؤال إلى المنظومة'))
TOC += concl_toc
TOC.append((0, 'الملاحق'))
TOC += [(1, t) for t in APPENDIX_TITLES]
TOC.append((0, 'الفهارس'))
TOC += [(1, t) for t in ('فهرس الآيات القرآنية', 'فهرس الأحاديث والآثار')]
TOC.append((0, 'قائمة المراجع'))
json.dump(TOC, open('toc.json', 'w', encoding='utf-8'), ensure_ascii=False)

ORD = {'1': 'الأول', '2': 'الثاني', '3': 'الثالث', '4': 'الرابع', '5': 'الخامس', '6': 'السادس', '7': 'السابع', '8': 'الثامن',
       '9': 'التاسع', '10': 'العاشر', '11': 'الحادي عشر', '12': 'الثاني عشر', '13': 'الثالث عشر', '14': 'الرابع عشر'}


def where(ch, sec):
    if ch == 'intro':
        return 'المقدمة'
    if ch == 'concl':
        return 'الخاتمة'
    return sec if sec else f'افتتاح الفصل {ORD[ch]}'


def loc_key(l):
    if l == 'المقدمة':
        return (0, 0)
    if l == 'الخاتمة':
        return (99, 0)
    m = re.match(r'(\d+)\.(\d+)', l)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    n = [k for k, v in ORD.items() if l.endswith(v)]
    return (int(n[0]) if n else 50, -1)


def join_locs(locs):
    seen = []
    for l in locs:
        if l not in seen:
            seen.append(l)
    return '، '.join(sorted(seen, key=loc_key))


# ------------------------------------------------------------------ verses
alt = '|'.join(sorted(map(re.escape, SURAHS), key=len, reverse=True))
VREF = re.compile(r'(?<![\w])(' + alt + r')\s*:\s*(\d+)(?:\s*[-–]\s*(\d+))?')
QUOTE = re.compile(r'﴿([^﴾]+)﴾')
V = {}
prev_quote = None
for ch, sec, s in U:
    if sec in REFSECS or re.match(r'^\d+\.\s', s):
        continue
    quotes = [(m.start(), m.end(), m.group(1)) for m in QUOTE.finditer(s)]
    for m in VREF.finditer(s):
        key = (SURAHS.index(m.group(1)), int(m.group(2)), int(m.group(3)) if m.group(3) else None)
        text = None
        before = [qq for qq in quotes if qq[1] <= m.start() and m.start() - qq[1] < 6]
        if before:
            text = before[-1][2]
        elif len(s) < 25 and prev_quote:
            text = prev_quote
        e = V.setdefault(key, {'text': None, 'locs': []})
        if text and not e['text']:
            e['text'] = text
        e['locs'].append(where(ch, sec))
    if quotes:
        prev_quote = quotes[-1][2]
    elif len(s) >= 25:
        prev_quote = None


def short(t, n=7):
    w = t.split()
    return ' '.join(w[:n]) + ('…' if len(w) > n else '')


VERSES = [('الآية', 'طرفها كما ورد في الكتاب', 'مواضعها')]
for key in sorted(V, key=lambda k: (k[0], k[1], k[2] or 0)):
    si, a, b = key
    ref = f'{SURAHS[si]}: {a}' + (f'–{b}' if b else '')
    t = V[key]['text']
    VERSES.append((ref, f'﴿{short(t)}﴾' if t else 'إحالة دون نص', join_locs(V[key]['locs'])))

# ------------------------------------------------------------------ hadiths and athar
HADITH = [
    ('«الإحسان أن تعبد الله كأنك تراه، فإن لم تكن تراه فإنه يراك» (حديث جبريل)', 'متفق عليه', r'كأنك تراه|حديث جبريل'),
    ('«كلكم راعٍ وكلكم مسؤول عن رعيته»', 'متفق عليه', r'كلكم راع'),
    ('«لو أن فاطمة بنت محمد سرقت لقطعت يدها» (حديث المخزومية)', 'متفق عليه', r'فاطمة بنت محمد سرقت|المخزومية'),
    ('«أعطوا الأجير أجره قبل أن يجف عرقه»', 'ابن ماجه', r'أعطوا الأجير|حديث الأجير'),
    ('«أدِّ الأمانة إلى من ائتمنك، ولا تخن من خانك»', 'أبو داود', r'أدِّ الأمانة|أد الأمانة إلى من'),
    ('استشارته ﷺ أصحابه قبل بدر في لقاء العدو', 'مسلم', r'قبل بدر'),
    ('مشاورته ﷺ أصحابه يوم أحد في المقام والخروج', 'البخاري تعليقًا، كتاب الاعتصام', r'يوم أحد في المقام'),
    ('«إن الدين يسر، ولن يشادّ الدين أحد إلا غلبه»', 'البخاري', r'إن الدين يسر'),
    ('«إن الله وضع عن أمتي الخطأ والنسيان وما استُكرهوا عليه»', 'ابن ماجه، وفي ثبوته خلاف', r'وضع عن أمتي|حديث الوضع'),
    ('«عجبًا لأمر المؤمن، إن أمره كله خير»', 'مسلم', r'عجبًا لأمر المؤمن'),
    ('«كل مولود يولد على الفطرة، فأبواه يهودانه أو ينصرانه أو يمجسانه»', 'متفق عليه', r'يولد على الفطرة|حديث الفطرة'),
    ('«كفى بالمرء كذبًا أن يحدّث بكل ما سمع»', 'مسلم، مقدمة الصحيح', r'كفى بالمرء كذب'),
    ('حديث حذيفة: «إذا مرّ بآية فيها تسبيح سبّح…»', 'مسلم', r'إذا مرّ بآية|حديث حذيفة'),
    ('حديث أبي ذر: قيامه ﷺ بآية يرددها حتى أصبح', 'النسائي وابن ماجه وأحمد', r'حديث أبي ذر|يرددها حتى أصبح'),
    ('«الراحمون يرحمهم الرحمن، ارحموا من في الأرض يرحمكم من في السماء»', 'أبو داود والترمذي', r'الراحمون يرحمهم'),
    ('«من لا يَرحم لا يُرحم»', 'متفق عليه', r'من لا يَرحم|من لا يرحم'),
    ('«لا حسد إلا في اثنتين»', 'متفق عليه', r'لا حسد إلا في اثنتين'),
    ('«عليكم بالصدق، فإن الصدق يهدي إلى البر…»', 'متفق عليه، واللفظ لمسلم', r'عليكم بالصدق|إن الصدق يهدي'),
    ('حديث أم كلثوم بنت عقبة في الإصلاح بين الناس', 'متفق عليه', r'أم كلثوم'),
    ('حديث كعب بن مالك في تخلّفه عن تبوك', 'متفق عليه', r'كعب بن مالك'),
    ('«ما من مسلم يغرس غرسًا أو يزرع زرعًا…»', 'متفق عليه', r'يغرس غرس|حديث الغرس'),
    ('«مثل المؤمنين في توادّهم وتراحمهم وتعاطفهم مثل الجسد…»', 'متفق عليه، واللفظ لمسلم', r'توادّهم|توادهم|الجسد الواحد'),
    ('«ما من عبد يسترعيه الله رعية، يموت يوم يموت وهو غاشّ لرعيته…» (حديث معقل بن يسار)', 'متفق عليه، واللفظ لمسلم', r'يسترعيه الله|معقل بن يسار'),
    ('«لا طاعة في معصية، إنما الطاعة في المعروف»', 'متفق عليه', r'لا طاعة في معصية'),
    ('«على المرء المسلم السمع والطاعة فيما أحب وكره، ما لم يؤمر بمعصية»', 'متفق عليه', r'السمع والطاعة فيما أحب'),
    ('«لا يُلدغ المؤمن من جحر واحد مرتين»', 'متفق عليه', r'يُلدغ المؤمن|يلدغ المؤمن'),
    ('«اعقلها وتوكل»', 'الترمذي وابن حبان', r'اعقلها'),
    ('«لو أنكم تتوكلون على الله حق توكله لرزقكم كما يرزق الطير»', 'الترمذي', r'حق توكله'),
    ('«احرص على ما ينفعك، واستعن بالله ولا تعجز»', 'مسلم', r'احرص على ما ينفعك'),
    ('حديث الاستخارة: «إذا همّ أحدكم بالأمر فليركع ركعتين…»', 'البخاري، من حديث جابر', r'إذا همّ أحدكم|الاستخارة'),
    ('«أنتم شهداء الله في الأرض»', 'متفق عليه', r'شهداء الله في الأرض'),
    ('«الكيّس من دان نفسه وعمل لما بعد الموت»', 'الترمذي', r'الكيّس من دان|الكيس من دان'),
    ('«اللهم انفعني بما علمتني، وعلمني ما ينفعني، وزدني علمًا»', 'الترمذي وابن ماجه', r'انفعني بما علمتني'),
    ('«قد أفلح من أسلم، ورُزق كفافًا، وقنّعه الله بما آتاه»', 'مسلم', r'قد أفلح من أسلم'),
]
ATHAR = [
    ('«حاسبوا أنفسكم قبل أن تحاسبوا» (عمر بن الخطاب)', 'الترمذي تعليقًا', r'حاسبوا أنفسكم'),
    ('كتاب عمر بن الخطاب إلى أبي موسى الأشعري في القضاء', 'الدارقطني والبيهقي', r'أبي موسى'),
    ('المسألة المشتركة (قضاء عمر في الفرائض)', 'البيهقي والدارمي', r'المسألة المشتركة|المشرّكة'),
]


def hadith_rows(items):
    rows = []
    for text, src, pat in items:
        rx = re.compile(pat)
        locs = [where(ch, sec) for ch, sec, s in U if rx.search(s) and not re.match(r'^\d+\.\s', s) and sec not in REFSECS]
        if locs:
            rows.append((text, src, join_locs(locs)))
        else:
            print('NOT FOUND:', text)
    return rows


H_ROWS = [('طرف الحديث', 'التخريج كما ورد في الكتاب', 'مواضعه')] + hadith_rows(HADITH)
A_ROWS = [('الأثر', 'التخريج كما ورد في الكتاب', 'مواضعه')] + hadith_rows(ATHAR)

# ------------------------------------------------------------------ references
REF_AR = {
    'القرآن الكريم وعلومه وتفسيره': [
        'القرآن الكريم.',
        'ابن تيمية، أحمد بن عبد الحليم (ت.728هـ). مقدمة في أصول التفسير. تحقيق: عدنان زرزور. دار القرآن الكريم، الكويت، 1971.',
        'ابن عاشور، محمد الطاهر (ت.1393هـ). التحرير والتنوير. الدار التونسية للنشر، تونس، 1984م.',
        'ابن كثير، إسماعيل بن عمر (ت.774هـ). تفسير القرآن العظيم. تحقيق: سامي بن محمد سلامة. دار طيبة.',
        'الرازي، فخر الدين محمد بن عمر (ت.606هـ). مفاتيح الغيب (التفسير الكبير).',
        'الزحيلي، وهبة بن مصطفى (ت.2015م). التفسير المنير في العقيدة والشريعة والمنهج. دار الفكر المعاصر، بيروت.',
        'الزركشي، بدر الدين (ت.794هـ). البرهان في علوم القرآن.',
        'الزمخشري، محمود بن عمر (ت.538هـ). الكشّاف عن حقائق التنزيل. دار الكتاب العربي.',
        'سيد قطب، إبراهيم حسين (ت.1966م). في ظلال القرآن. دار الشروق، القاهرة.',
        'السيوطي، جلال الدين (ت.911هـ). الإتقان في علوم القرآن.',
        'السيوطي، جلال الدين (ت.911هـ). لباب النقول في أسباب النزول. مؤسسة الكتب الثقافية.',
        'الطباطبائي، محمد حسين (ت.1981م). الميزان في تفسير القرآن. مؤسسة الأعلمي للمطبوعات، بيروت.',
        'الطبري، محمد بن جرير (ت.310هـ). جامع البيان عن تأويل آي القرآن. تحقيق: أحمد شاكر ومحمود شاكر. مؤسسة الرسالة.',
        'عبد الباقي، محمد فؤاد. المعجم المفهرس لألفاظ القرآن الكريم. دار الكتب المصرية، القاهرة.',
        'القرطبي، محمد بن أحمد (ت.671هـ). الجامع لأحكام القرآن. دار الكتب المصرية.',
        'مسلم، مصطفى (1997). مباحث في التفسير الموضوعي. دار القلم، دمشق.',
        'الواحدي، علي بن أحمد (ت.468هـ). أسباب النزول. تحقيق: كمال بسيوني زغلول. دار الكتب العلمية.',
    ],
    'الحديث وعلومه والسيرة': [
        'ابن حبان، محمد (ت.354هـ). الصحيح، بترتيب ابن بلبان.',
        'ابن الصلاح، عثمان بن عبد الرحمن (ت.643هـ). مقدمة ابن الصلاح في علوم الحديث.',
        'ابن ماجه، محمد بن يزيد (ت.273هـ). السنن.',
        'ابن هشام، عبد الملك (ت.218هـ). السيرة النبوية. تحقيق: عمر عبد السلام تدمري. دار الكتاب العربي.',
        'أبو داود، سليمان بن الأشعث (ت.275هـ). السنن.',
        'أحمد بن حنبل (ت.241هـ). المسند.',
        'الأعظمي، محمد مصطفى (1978). دراسات في الحديث النبوي وتاريخ تدوينه. المكتب الإسلامي.',
        'البخاري، محمد بن إسماعيل (ت.256هـ). الصحيح.',
        'البيهقي، أحمد بن الحسين (ت.458هـ). السنن الكبرى.',
        'الترمذي، محمد بن عيسى (ت.279هـ). الجامع. دار الغرب الإسلامي.',
        'الدارقطني، علي بن عمر (ت.385هـ). السنن.',
        'الدارمي، عبد الله بن عبد الرحمن (ت.255هـ). السنن.',
        'الغزالي، محمد (ت.1996م). فقه السيرة.',
        'الغضبان، منير محمد. الفقه الحركي للسيرة النبوية.',
        'المباركفوري، صفي الرحمن. الرحيق المختوم.',
        'مسلم بن الحجاج (ت.261هـ). الصحيح.',
        'مهدي رزق الله أحمد. السيرة النبوية في ضوء المصادر الأصلية.',
        'النسائي، أحمد بن شعيب (ت.303هـ). السنن.',
    ],
    'أصول الفقه والمقاصد والفكر الإسلامي': [
        'ابن تيمية، أحمد بن عبد الحليم (ت.728هـ). درء تعارض العقل والنقل. تحقيق: محمد رشاد سالم. جامعة الإمام محمد بن سعود الإسلامية، الرياض.',
        'ابن تيمية، تقي الدين أحمد (ت.728هـ). السياسة الشرعية في إصلاح الراعي والرعية. دار الكتب العلمية.',
        'ابن خلدون، عبد الرحمن بن محمد (ت.808هـ). المقدمة. دار الكتب العلمية، بيروت.',
        'ابن قيم الجوزية، محمد بن أبي بكر (ت.751هـ). إعلام الموقعين عن رب العالمين. تحقيق: طه عبد الرؤوف سعد. دار الجيل، بيروت.',
        'ابن مسكويه، أحمد بن محمد (ت.421هـ). تهذيب الأخلاق وتطهير الأعراق. تحقيق: ابن الخطيب. مكتبة الثقافة الدينية.',
        'الزركشي، بدر الدين (ت.794هـ). البحر المحيط في أصول الفقه. تحقيق: عبد القادر عبد الله العاني. وزارة الأوقاف الكويتية، 1992م.',
        'الشاطبي، إبراهيم بن موسى (ت.790هـ). الموافقات في أصول الشريعة. تحقيق: مشهور بن حسن آل سلمان. دار ابن عفان.',
        'الشافعي، محمد بن إدريس (ت.204هـ). الرسالة. تحقيق: أحمد محمد شاكر. مكتبة الحلبي، القاهرة، 1940م.',
        'الغزالي، أبو حامد محمد (ت.505هـ). إحياء علوم الدين. دار الحديث، القاهرة.',
        'الغزالي، أبو حامد محمد (ت.505هـ). المستصفى من علم الأصول. تحقيق: محمد عبد السلام عبد الشافي. دار الكتب العلمية، بيروت، 1993م.',
        'الغزالي، أبو حامد محمد (ت.505هـ). المنقذ من الضلال. تحقيق: جميل صليبا وكامل عياد. دار الأندلس.',
        'الماوردي، أبو الحسن علي (ت.450هـ). الأحكام السلطانية. دار الكتب العلمية.',
    ],
    'المعاجم اللغوية': [
        'ابن فارس، أحمد (ت.395هـ). معجم مقاييس اللغة. تحقيق: عبد السلام محمد هارون. دار الفكر.',
        'ابن منظور، محمد بن مكرم (ت.711هـ). لسان العرب. دار صادر، بيروت.',
        'الراغب الأصفهاني، الحسين بن محمد (ت.502هـ). المفردات في غريب القرآن. تحقيق: صفوان عدنان الداودي. دار القلم، دمشق.',
    ],
}

# Western references: every entry from the chapter lists, one per work, fullest variant kept
refs = []
for ch, sec, s in U:
    pass
REFS_RAW = []
inref = None
for ch, sec, s in U:
    if 'قائمة المراجع' in s and re.match(r'^\d+\.\d+', s):
        inref = ch
        continue
    if inref and ch != inref:
        inref = None
    if inref:
        m = re.match(r'^\d+\.\s+(.*)', s)
        if m and re.match(r'[A-Za-z]', m.group(1)):
            REFS_RAW.append(m.group(1).strip())
MERGE = {                     # variant -> canonical
    'Aristotle. Nicomachean Ethics, Book VI. Translated by W.D. Ross. Oxford University Press.': 'Aristotle. Nicomachean Ethics. Trans. W.D. Ross. Oxford University Press.',
    'Popper, K. (1934/1959). The Logic of Scientific Discovery. Hutchinson & Co.': 'Popper, K. (1959). The Logic of Scientific Discovery. Hutchinson.',
    'Mayo, E. (1933). The Human Problems of Industrial Civilization. Macmillan.': 'Mayo, E. (1933). The Human Problems of an Industrial Civilization. Macmillan.',
}


def wkey(r):
    n = r.lower()
    au = re.split(r'[ ,.]', n)[0]
    yr = re.search(r'\((\d{4})', n)
    t = re.search(r'\)\.?\s*([^.]+)', n)
    return (au, yr.group(1) if yr else '', re.sub(r'\W', '', t.group(1))[:20] if t else n[:30])


W_G = {}
for r in REFS_RAW:
    r = MERGE.get(r, r)
    W_G.setdefault(wkey(r), []).append(r)
WESTERN = sorted({max(v, key=len) for v in W_G.values()}, key=lambda s: s.lower().lstrip("'"))
RELIGION = ('McGrath', 'Neusner', 'Second Vatican', 'Pontifical', 'Elon', 'Laozi')
W_REL = [r for r in WESTERN if r.startswith(RELIGION)]
W_OTH = [r for r in WESTERN if not r.startswith(RELIGION)]

# ================================================================== build
SRC, OUT = 'orig', 'indexes_build'
shutil.rmtree(OUT, ignore_errors=True)
shutil.copytree(SRC, OUT)
tree = etree.parse(f'{SRC}/word/document.xml')
body = tree.getroot().find(q('body'))
kids = list(body)
sect = body.find(q('sectPr'))
K = {i: kids[i] for i in (10, 11, 17, 18, 20, 28, 102, 498, 3376, 3377, 3078)}


def para(like, text):
    p = fresh(copy.deepcopy(K[like])); set_text(p, text); return p


def blank():
    return fresh(copy.deepcopy(K[17]))


def heading(text):
    return [blank(), para(18, text), blank()]


def ref_line(n, text):
    p = fresh(copy.deepcopy(K[3377]))
    r = [x for x in p.findall(q('r')) if ptext(x)]
    set_runs(p, [(f'{n}.  ', r[0]), (text, r[1])])
    return p


def table(tmpl, data):
    t = fresh(copy.deepcopy(K[tmpl]))
    ncol = len(data[0])
    while len(t.findall(q('tr'))) > len(data):
        delete_row(t, len(t.findall(q('tr'))) - 1)
    while len(t.findall(q('tr'))) < len(data):
        add_row(t, 1, [''] * ncol)
    for r, row in enumerate(data):
        for c, text in enumerate(row):
            cell_text(t, r, c, text)
    return t


OUTP = [para(11, 'الفهارس')]
OUTP += heading('فهرس الآيات القرآنية')
OUTP.append(para(20, 'الآيات مرتبة بترتيب المصحف، ومع كل آية طرفها كما ورد في الكتاب، ومواضعها برقم القسم (مثل 8.2 أي الفصل الثامن، القسم الثاني). '
                     'و«إحالة دون نص» تعني أن الكتاب أحال إلى الآية دون أن ينقل لفظها.'))
OUTP += [table(498, VERSES), blank()]
OUTP += [fresh(copy.deepcopy(K[10]))]
OUTP += heading('فهرس الأحاديث والآثار')
OUTP.append(para(20, 'الأحاديث مرتبة بحسب أول ورودها في الكتاب، والتخريج هو ما ذكره الكتاب في موضعه.'))
OUTP += [table(498, H_ROWS), blank(), para(28, 'الآثار'), table(498, A_ROWS), blank()]
OUTP += [fresh(copy.deepcopy(K[10]))]
OUTP += [para(11, 'قائمة المراجع')]
OUTP.append(para(20, 'تجمع هذه القائمة مراجع الفصول الأربعة عشر في موضع واحد، بعد حذف المكرر وتوحيد بيانات النشر، مرتبةً بحسب نوع المصدر. '
                     'وتُذكر المراجع في متن الفصول باسم المؤلف وسنة النشر.'))
n = 0
for title, items in list(REF_AR.items()) + [('المصادر الدينية المقارنة', W_REL), ('المراجع الأجنبية', W_OTH)]:
    OUTP.append(blank())
    OUTP.append(para(3376, title))
    for it in items:
        n += 1
        OUTP.append(ref_line(n, it))

for ch in list(body):
    body.remove(ch)
for e in OUTP:
    body.append(e)
body.append(sect)
tree.write(f'{OUT}/word/document.xml', xml_declaration=True, encoding='UTF-8', standalone=True)
out = os.path.abspath('الفهارس_النهائية.docx')
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-qXr', out, '.'], cwd=OUT, check=True)
print('wrote', out, '| verses', len(VERSES) - 1, '| hadiths', len(H_ROWS) - 1, '| athar', len(A_ROWS) - 1, '| refs', n)
