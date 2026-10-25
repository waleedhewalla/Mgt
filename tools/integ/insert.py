"""Insert new paragraphs/tables and verse-index rows into a book docx, cloning existing formatting.

usage: python3 insert.py <src.docx> <out.docx> <spec_module> <book_key>
The spec module defines SPECS[book_key] = dict(templates=..., ops=[...]).
"""
import sys, copy, re, importlib.util
import docx
from docx.oxml.ns import qn
sys.path.insert(0, '/home/user/Mgt/tools')
from surahs import SURAHS

W_T = qn('w:t')
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def ptext(el):
    return ''.join(x.text or '' for x in el.iter(W_T))


def set_par_text(p, text, bold=None):
    """Keep pPr and the first run's rPr; replace all content with one run holding text."""
    runs = p.findall(qn('w:r'))
    r0 = copy.deepcopy(runs[0]) if runs else None
    for ch in list(p):
        if ch.tag != qn('w:pPr'):
            p.remove(ch)
    if r0 is None:
        r0 = p.makeelement(qn('w:r'), {})
    for ch in list(r0):
        if ch.tag != qn('w:rPr'):
            r0.remove(ch)
    if bold is not None:
        rpr = r0.find(qn('w:rPr'))
        if rpr is None:
            rpr = r0.makeelement(qn('w:rPr'), {})
            r0.insert(0, rpr)
        for tag in ('w:b', 'w:bCs'):
            e = rpr.find(qn(tag))
            if e is not None:
                rpr.remove(e)
        if bold:
            # w:b must precede colour/size in CT_RPr order: put after rFonts if present
            pos = 1 if rpr.find(qn('w:rFonts')) is not None else 0
            for tag in ('w:bCs', 'w:b'):
                rpr.insert(pos, rpr.makeelement(qn(tag), {}))
    t = r0.makeelement(W_T, {})
    t.text = text
    t.set(XML_SPACE, 'preserve')
    r0.append(t)
    p.append(r0)
    return p


def set_cell_text(tc, text, bold=None):
    ps = tc.findall(qn('w:p'))
    for extra in ps[1:]:
        tc.remove(extra)
    set_par_text(ps[0], text, bold)


def build_table(tmpl, header, rows):
    tbl = copy.deepcopy(tmpl)
    trs = tbl.findall(qn('w:tr'))
    head, body = trs[0], trs[1]
    for tr in trs[1:]:
        tbl.remove(tr)
    for tc, h in zip(head.findall(qn('w:tc')), header):
        set_cell_text(tc, h)
    for row in rows:
        tr = copy.deepcopy(body)
        tcs = tr.findall(qn('w:tc'))
        assert len(tcs) == len(row), (len(tcs), row)
        for tc, v in zip(tcs, row):
            set_cell_text(tc, v)
        tbl.append(tr)
    return tbl


AR_DIG = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')


def verse_key(label):
    m = re.match(r'\s*([^:]+?)\s*:\s*([0-9٠-٩]+)', label)
    if not m:
        return None
    s = m.group(1).strip()
    if s not in SURAHS:
        return None
    return (SURAHS.index(s), int(m.group(2).translate(AR_DIG)))


def main():
    src, out, specpath, key = sys.argv[1:5]
    spec_mod = importlib.util.spec_from_file_location('spec', specpath)
    m = importlib.util.module_from_spec(spec_mod)
    spec_mod.loader.exec_module(m)
    S = m.SPECS[key]
    d = docx.Document(src)
    body = d.element.body
    P = [el for el in body.iterchildren() if el.tag.endswith('}p')]
    T = [el for el in body.iterchildren() if el.tag.endswith('}tbl')]
    tp = {k: P[v] for k, v in S['templates'].items() if isinstance(v, int)}
    tt = {k: T[v - 1] for k, v in S.get('tables', {}).items()}
    log = []

    def find_par(text, min_idx=0):
        hits = [i for i, el in enumerate(P) if i >= min_idx and ptext(el).strip().startswith(text)]
        if len(hits) != 1:
            raise SystemExit(f'anchor {text[:40]!r} min={min_idx}: {len(hits)} hits')
        return P[hits[0]]

    def make(item):
        kind = item[0]
        if kind == 'table':
            _, tkey, header, rows = item
            return build_table(tt[tkey], header, rows)
        text = item[1]
        tmpl = tp[kind]
        bold = S.get('bold', {}).get(kind)
        return set_par_text(copy.deepcopy(tmpl), text, bold)

    for op in S['ops']:
        if op['op'] in ('after', 'before'):
            anchor = find_par(op['anchor'], op.get('min', 0))
            new = [make(it) for it in op['items']]
            if op['op'] == 'after':
                ref = anchor
                for el in new:
                    ref.addnext(el)
                    ref = el
            else:
                for el in new:
                    anchor.addprevious(el)
            log.append(f"{op['op']} {op['anchor'][:40]}: +{len(new)}")
        elif op['op'] == 'edit':
            el = find_par(op['anchor'], op.get('min', 0))
            p = docx.text.paragraph.Paragraph(el, d)
            full = ''.join(r.text for r in p.runs)
            assert full.count(op['old']) == 1, (op['old'], full.count(op['old']))
            sys.path.insert(0, '../ar')
            from apply_docx import replace_once
            ok, why = replace_once(p, op['old'], op['new'])
            assert ok, why
            log.append(f"edit {op['old'][:30]}")
        elif op['op'] == 'index':
            head = [i for i, el in enumerate(P) if ptext(el).strip() == op['heading']][-1]
            hel = P[head]
            nxt = hel.getnext()
            while nxt is not None and not nxt.tag.endswith('}tbl'):
                nxt = nxt.getnext()
            tbl = nxt
            for label, snippet, loc in op['rows']:
                k = verse_key(label)
                assert k, label
                trs = tbl.findall(qn('w:tr'))[1:]
                done = False
                for tr in trs:
                    tcs = tr.findall(qn('w:tc'))
                    if verse_key(ptext(tcs[0])) == k and ptext(tcs[0]).strip().translate(AR_DIG).replace(' ', '') == label.translate(AR_DIG).replace(' ', ''):
                        cur = ptext(tcs[2]).strip()
                        if loc not in cur:
                            set_cell_text(tcs[2], cur + '، ' + loc)
                        done = True
                        log.append(f'index merge {label}')
                        break
                if done:
                    continue
                pos = None
                for tr in trs:
                    kk = verse_key(ptext(tr.findall(qn('w:tc'))[0]))
                    if kk and kk > k:
                        pos = tr
                        break
                tmpl_tr = trs[0]
                new = copy.deepcopy(tmpl_tr)
                for tc, v in zip(new.findall(qn('w:tc')), (label, snippet, loc)):
                    set_cell_text(tc, v)
                if pos is not None:
                    pos.addprevious(new)
                else:
                    trs[-1].addnext(new)
                log.append(f'index add {label}')
    d.save(out)
    print('\n'.join(log))


if __name__ == '__main__':
    main()
