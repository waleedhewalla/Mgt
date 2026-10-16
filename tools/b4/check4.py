"""Checks a Book Four markup source: syntax, length, banned items, style metrics, citations list."""
import sys, re, json, os
HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.dirname(HERE)
BANNED = ['—', 'الاستدعاء المجزأ', 'الاستدعاء التجزيئي', 'صناعة الفارق', 'ترشيد الوعي', 'الصحوة إلى اليقظة', 'النموذج الكلي',
          'انتقام الأفكار', 'قراءة قيمية راشدة', 'قادة النهضة', 'في عالمنا المتسارع', 'لا يخفى على أحد', 'تجدر الإشارة', 'السيارة',
          'الشطرنج', '**', 'شكل ١', 'شكل 1:']
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر',
       'الثاني عشر', 'الثالث عشر', 'الرابع عشر']


def norm(s): return re.sub(r'[ً-ْـ«»"“”،؛:.؟!()\[\]﴿﴾]', ' ', s)


def grams(s, n=5):
    w = norm(s).split()
    return {' '.join(w[i:i + n]) for i in range(len(w) - n + 1) if all(len(t) > 1 for t in w[i:i + n])}


path = sys.argv[1]
lines = open(path, encoding='utf-8').read().splitlines()
errors, body, all_text, chapter, secs, tables, figs = [], [], [], None, [], 0, 0
in_tab = False
for n, raw in enumerate(lines, 1):
    ln = raw.strip()
    if not ln:
        continue
    if in_tab:
        if ln == '@end':
            in_tab = False
        else:
            all_text.append(ln.replace('|', ' '))
        continue
    if ln.startswith('@table'):
        in_tab, tables = True, tables + 1
    elif ln.startswith('@fig'):
        figs += 1
        try:
            spec = json.loads(ln[4:].strip())
            assert spec.get('kind') in ('flow', 'cycle', 'levels', 'hub', 'matrix'), 'kind'
            assert spec.get('items'), 'items'
        except Exception as ex:
            errors.append(f'L{n}: bad @fig json ({ex})')
    elif ln.startswith('@chapter'):
        if '|' not in ln:
            errors.append(f'L{n}: @chapter needs "الفصل X | العنوان"')
        else:
            no = ln[8:].split('|')[0].strip().replace('الفصل ', '')
            chapter = ORD.index(no) + 1 if no in ORD else None
            if chapter is None:
                errors.append(f'L{n}: unknown chapter ordinal {no}')
    elif ln.startswith('@verse'):
        if '|' not in ln or '﴿' not in ln:
            errors.append(f'L{n}: @verse needs ﴿…﴾ | ref')
    elif ln.startswith('### '):
        m = re.match(r'^### (\d+)\.(\d+)\s', ln)
        if not m:
            errors.append(f'L{n}: section must start with N.M')
        else:
            if chapter and int(m.group(1)) != chapter:
                errors.append(f'L{n}: section {m.group(0)} not in chapter {chapter}')
            secs.append(int(m.group(2)))
        all_text.append(ln[4:])
    elif ln.startswith('#### '):
        all_text.append(ln[5:])
    elif ln.startswith(('@part', '@heading', '@title', '@note')):
        all_text.append(ln.split(' ', 1)[1] if ' ' in ln else '')
    elif ln.startswith('@'):
        errors.append(f'L{n}: unknown directive')
    elif ln.startswith(('-', '*', '#')):
        errors.append(f'L{n}: markdown list/heading not allowed: {ln[:30]}')
    else:
        body.append(ln)
        all_text.append(ln)
if in_tab:
    errors.append('unclosed @table')
if secs and secs != list(range(1, len(secs) + 1)):
    errors.append(f'section numbers not consecutive: {secs}')
full = '\n'.join(all_text)
for b in BANNED:
    if b in full:
        errors.append(f'banned: {b} x{full.count(b)}')
words = len(full.split())
txt = ' '.join(re.sub(r'﴿[^﴾]*﴾', '', p) for p in body if len(p.split()) > 25)
W = max(1, len(txt.split()))
ss = [s.split() for s in re.split(r'[.؟!]', txt) if len(s.split()) > 1]
L = [len(s) for s in ss] or [0]
style = dict(mean=round(sum(L) / len(L), 1), long=round(100 * sum(l >= 40 for l in L) / len(L)),
             short=round(100 * sum(l <= 8 for l in L) / len(L)), q=round(1000 * txt.count('؟') / W, 1),
             semi=round(1000 * txt.count('؛') / W, 1), colon=round(1000 * txt.count(':') / W, 1),
             inna=round(1000 * len(re.findall(r'(^|[.؟!]\s*)إنّ?\s', txt)) / W, 1))
if not 18 <= style['mean'] <= 28: errors.append(f"mean sentence {style['mean']} outside 18-28")
if style['long'] < 12: errors.append(f"long sentences {style['long']}% < 12%")
if style['semi'] > 8: errors.append(f"semicolons {style['semi']}/1000 > 8")
if style['q'] < 4: errors.append(f"questions {style['q']}/1000 < 4")
# repeated paragraph openers
op = [' '.join(p.split()[:2]) for p in body]
rep = [op[i] for i in range(2, len(op)) if op[i] == op[i - 1] == op[i - 2]]
if rep: errors.append(f'three paragraphs in a row open the same way: {set(rep)}')
ref = set()
for f in ('a_n.txt', 'b_n.txt'):
    pth = os.path.join(S, 'style', 'src', f)
    if os.path.exists(pth):
        ref |= grams(open(pth, encoding='utf-8').read())
ov = grams(full) & ref
if ov: errors.append(f'5-gram overlap with reference texts: {list(ov)[:5]}')
cites = re.findall(r'\((?:رواه|أخرجه)[^)]{2,60}\)', full)
print(f'{os.path.basename(path)}: errors {len(errors)} | words {words} | sections {len(secs)} | tables {tables} | figs {figs}')
print('style', style)
print('citations', len(cites), sorted(set(cites)))
for e in errors:
    print('  -', e)
