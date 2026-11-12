"""Coordinator pass on التأسيس المعرفي after the reviewers' edits (works by content, not IDs).

usage: python3 post_ta.py <in.docx> <out.docx>
- reading-guide table: drop the ⚠ row, redefine ✓ for the print edition
- one Ibn Khaldun edition everywhere (Dar al-Fikr, 2nd ed. 1408/1988, per the OpenITI/Shamela metadata)
- glossary: move the paragraph-form entries into the glossary table
- remove the drafting table «ملاحظات على الملاحق» and the Zilal reference no chapter cites any more
"""
import sys, copy, re
import docx
from docx.oxml.ns import qn

XS = '{http://www.w3.org/XML/1998/namespace}space'
txt = lambda e: ''.join(t.text or '' for t in e.iter(qn('w:t')))


def set_cell(tc, text):
    ps = tc.findall(qn('w:p'))
    for p in ps[1:]:
        tc.remove(p)
    p = ps[0]
    runs = p.findall(qn('w:r'))
    r0 = copy.deepcopy(runs[0]) if runs else p.makeelement(qn('w:r'), {})
    for ch in list(p):
        if ch.tag != qn('w:pPr'):
            p.remove(ch)
    for ch in list(r0):
        if ch.tag != qn('w:rPr'):
            r0.remove(ch)
    t = r0.makeelement(qn('w:t'), {}); t.text = text; t.set(XS, 'preserve'); r0.append(t); p.append(r0)


def replace_all_text(body, old, new):
    n = 0
    for t in body.iter(qn('w:t')):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new); n += 1
    return n


def main():
    src, out = sys.argv[1], sys.argv[2]
    d = docx.Document(src)
    body = d.element.body
    log = []
    tables = [e for e in body.iterchildren() if e.tag.endswith('}tbl')]

    # 1. reading guide (first table whose header is «الرمز / التسمية»)
    guide = next(t for t in tables if txt(t.find(qn('w:tr'))).startswith('الرمز'))
    for tr in guide.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        k = txt(tcs[0]).strip()
        if k == '⚠':
            guide.remove(tr); log.append('guide: ⚠ row removed')
        elif k == '✓':
            set_cell(tcs[1], 'نصّ طوبق لفظه أو موضعه على مصدره في هذه النشرة: المصحف، أو كتب الحديث، أو التفاسير، أو نصوص التراث الرقمية')
            log.append('guide: ✓ redefined')

    # 2. Ibn Khaldun edition
    n = replace_all_text(body, '١٤٠١هـ/١٩٨١م', 'ط٢، ١٤٠٨هـ/١٩٨٨م') + replace_all_text(body, '1401هـ/1981م', 'ط2، 1408هـ/1988م')
    log.append(f'Ibn Khaldun edition unified in {n} places')

    # 3. glossary paragraphs -> table rows
    P = [e for e in body.iterchildren() if e.tag.endswith('}p')]
    intro = next((p for p in P if txt(p).strip().startswith('ومن مصطلحات الأبواب الأولى')), None)
    if intro is not None:
        gl = intro.getprevious()
        while gl is not None and not gl.tag.endswith('}tbl'):
            gl = gl.getprevious()
        rows = gl.findall(qn('w:tr'))
        tmpl = rows[-1]
        el = intro.getnext(); moved = 0; last = tmpl
        while el is not None and el.tag.endswith('}p') and ':' in txt(el) and not (el.find(qn('w:pPr')) is not None and el.find(qn('w:pPr')).find(qn('w:pStyle')) is not None):
            term, defin = txt(el).split(':', 1)
            defin = defin.strip().replace('وهو أعلى المراتب، لكنه الأضيق', 'وهو أعمق المراتب حضورًا، لكنه الأضيق')
            new = copy.deepcopy(tmpl)
            tcs = new.findall(qn('w:tc'))
            set_cell(tcs[0], term.strip()); set_cell(tcs[1], defin)
            last.addnext(new); last = new
            nxt = el.getnext(); body.remove(el); el = nxt; moved += 1
        body.remove(intro)
        log.append(f'glossary: {moved} entries moved into the table')

    # 4. drafting table «ملاحظات على الملاحق»
    P = [e for e in body.iterchildren() if e.tag.endswith('}p')]
    h = next((p for p in P if txt(p).strip() == 'ملاحظات على الملاحق'), None)
    if h is not None:
        nxt = h.getnext()
        body.remove(h)
        if nxt is not None and nxt.tag.endswith('}tbl') and txt(nxt.find(qn('w:tr'))).startswith('#'):
            body.remove(nxt)
        log.append('drafting table removed')

    # 5. references: Zilal no longer cited in any chapter
    P = [e for e in body.iterchildren() if e.tag.endswith('}p')]
    for p in P:
        if txt(p).strip().startswith('سيد قطب. في ظلال القرآن'):
            # the only other mentions are the chapters' «لم يُنسب … إلى … سيد قطب» disclaimers
            others = sum(1 for q in P if 'قطب' in txt(q) and not re.match(r'\s*لم (تُ|يُ|ت|ي)?نسب', txt(q))) - 1
            if others == 0:
                body.remove(p); log.append('reference to Zilal removed (no chapter cites it)')
    d.save(out)
    print('\n'.join(log))


if __name__ == '__main__':
    main()
