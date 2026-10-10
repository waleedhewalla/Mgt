"""Style metrics for a book docx: template phrases, tables, sentence-length spread, paragraph openings."""
import sys, zipfile, re, statistics, json
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
x = etree.fromstring(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))
b = x.find(W + 'body')
els = list(b)
start = next(i for i, e in enumerate(els) if e.tag == W + 'p' and ''.join(t.text or '' for t in e.iter(W + 't')).strip() == 'المقدمة'
             and any('TC' in (it.text or '') for it in e.iter(W + 'instrText')))
end = next(i for i, e in enumerate(els) if e.tag == W + 'p' and ''.join(t.text or '' for t in e.iter(W + 't')).strip() == 'الملاحق' and i > start)
paras, tables, cells = [], 0, []
for e in els[start:end]:
    if e.tag == W + 'tbl':
        tables += 1
        cells.append(' '.join(t.text or '' for t in e.iter(W + 't')))
    elif e.tag == W + 'p':
        t = ''.join(t.text or '' for t in e.iter(W + 't')).strip()
        if t:
            paras.append(t)
txt = '\n'.join(paras)
pats = ['حدود الادعاء', 'خاتمة الفصل', 'مدخل الفصل', 'أدّعي', 'اجتهادي', 'اجتهاد مني', 'لم يُختبر', 'ولا يُثبتها', 'هامش عدم اليقين',
        'هذا الفصل', '—', 'مثال توضيحي']
alltxt = txt + '\n' + '\n'.join(cells)
m = {p: alltxt.count(p) for p in pats}
long = [p for p in paras if len(p.split()) > 25]
sents = [s for s in re.split(r'[.؟!:]\s+', ' '.join(long)) if len(s.split()) > 1]
lens = [len(s.split()) for s in sents]
m.update({'words': len(alltxt.split()), 'prose_words': len(txt.split()), 'tables': tables, 'paras_starting_wa_%': round(100 * sum(p.startswith('و') for p in long) / len(long), 1),
          'sent_len_mean': round(statistics.mean(lens), 1), 'sent_len_sd': round(statistics.pstdev(lens), 1),
          'short_sents_%': round(100 * sum(l <= 6 for l in lens) / len(lens), 1), 'long_sents_%': round(100 * sum(l >= 40 for l in lens) / len(lens), 1),
          'questions': txt.count('؟')})
print(json.dumps(m, ensure_ascii=False))
