"""Builds src/17_index.md: verse index (mushaf order) and hadith index (first occurrence order), located by section number."""
import re, os, glob, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from surahs import SURAHS
AD = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
AR = str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩')
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر',
       'الثاني عشر', 'الثالث عشر', 'الرابع عشر']
SUR_RX = '|'.join(sorted(SURAHS, key=len, reverse=True))
VREF = re.compile(r'(?:سورة\s+)?(' + SUR_RX + r')\s*[:،]\s*(?:الآية\s*)?([\d٠-٩]+(?:\s*[-–]\s*[\d٠-٩]+)?)')

verses, hadith, order = {}, {}, []
files = sorted(f for f in glob.glob(os.path.join(HERE, 'src', '[01][0-9]_ch*.md')) if not f.endswith('.sources.md'))
for f in files:
    loc, ch = None, None
    for ln in open(f, encoding='utf-8'):
        ln = ln.strip()
        if ln.startswith('@chapter'):
            no = ln[8:].split('|')[0].strip().replace('الفصل ', '')
            ch = ORD.index(no) + 1
            loc = f'ف{ch}'
        m = re.match(r'^### (\d+\.\d+)', ln)
        if m:
            loc = m.group(1)
        if ln.startswith('@verse'):
            v, ref = ln[6:].rsplit('|', 1)
            for mm in VREF.finditer(ref):
                key = (SURAHS.index(mm.group(1)), mm.group(2).translate(AD).replace(' ', ''))
                w_=v.strip().strip('﴿﴾').split(); verses.setdefault(key, ['﴿' + ' '.join(w_[:7]) + ('…﴾' if len(w_) > 7 else '﴾'), []])[1].append(f'ف{ch}')
            continue
        # verses: ﴿text﴾ followed (within 40 chars) by a reference in parentheses
        for mv in re.finditer(r'﴿([^﴾]+)﴾\s*\(([^)]{3,40})\)', ln):
            for mm in VREF.finditer(mv.group(2)):
                key = (SURAHS.index(mm.group(1)), mm.group(2).translate(AD).replace(' ', ''))
                ent = verses.setdefault(key, ['﴿' + ' '.join(mv.group(1).split()[:7]) + ('…﴾' if len(mv.group(1).split()) > 7 else '﴾'), []])
                ent[1].append(loc)
        # hadith: «quote» … (رواه X)
        for mh in re.finditer(r'«([^»]{8,})»[^«]{0,80}?\((رواه[^)]{2,70})\)', ln):
            words = mh.group(1).split()
            k = ' '.join(words[:6])
            src = re.sub(r'،.*$', '', mh.group(2)).strip()
            if k not in hadith:
                hadith[k] = [' '.join(words[:9]) + ('…' if len(words) > 9 else ''), src, []]
                order.append(k)
            hadith[k][2].append(loc)


def uniq(xs):
    out = []
    for x in xs:
        if x and x not in out:
            out.append(x)
    return out


lines = ['@heading الفهارس', '#### أولًا: فهرس الآيات القرآنية',
         'الآيات مرتبة بترتيب المصحف، ومع كل آية طرفها كما ورد في الكتاب، ومواضعها برقم القسم (مثل 7.3 أي الفصل السابع، القسم الثالث) أو برقم الفصل.',
         '@table', 'السورة والآية | طرف الآية | المواضع']
for (s, a), (txt, locs) in sorted(verses.items(), key=lambda kv: (kv[0][0], int(re.match(r'\d+', kv[0][1]).group()))):
    lines.append(f'{SURAHS[s]}: {a.translate(AR)} | {txt.replace("|", " ")} | {"، ".join(uniq(locs))}')
lines += ['@end', '#### ثانيًا: فهرس الأحاديث والآثار',
          'الأحاديث والآثار مرتبة بحسب أول ورودها، والتخريج هو ما ذكره الكتاب في موضعه.', '@table', 'طرف الحديث أو الأثر | التخريج | المواضع']
for k in order:
    t, src, locs = hadith[k]
    lines.append(f'«{t.replace("|", " ")}» | {src} | {"، ".join(uniq(locs))}')
lines.append('@end')
open(os.path.join(HERE, 'src', '17_index.md'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
print('verses', len(verses), 'hadith', len(order))
