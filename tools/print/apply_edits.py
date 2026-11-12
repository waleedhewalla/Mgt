"""Apply reviewed edits to a docx. IDs are those printed by dump.py on the SAME source file.

usage: python3 apply_edits.py <src.docx> <out.docx> <edits.json> [more.json ...]
Edit kinds (all IDs resolved against the source before any change):
  {"id": "P65", "old": "...", "new": "..."}                    replace a substring occurring exactly once (run-aware)
  {"id": "T3.1.2", "old": "...", "new": "..."}                 same, inside a table cell
  {"id": "P65", "op": "delete"}                                 delete the paragraph
  {"id": "P65", "op": "insert_after", "text": "...", "like": "P60"}   new paragraph after P65, formatted like P60
  {"id": "P65", "op": "insert_before", "text": "...", "like": "P60"}
Edits that cannot be applied are listed and skipped; nothing is silently changed.
"""
import sys, json, copy
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

W_T = qn('w:t')
XS = '{http://www.w3.org/XML/1998/namespace}space'


def replace_once(p, old, new):
    runs = list(p.runs)
    texts = [r.text for r in runs]
    full = ''.join(texts)
    n = full.count(old)
    if n != 1:
        return f'occurs {n} times'
    s = full.find(old); e = s + len(old); pos = 0; first = True
    for r, t in zip(runs, texts):
        a, b = pos, pos + len(t); pos = b
        if b <= s or a >= e:
            continue
        pre = t[:max(0, s - a)]
        post = t[max(0, e - a):] if e < b else ''
        r.text = pre + (new if first else '') + post
        first = False
    return None


def clone_par(tmpl, text):
    p = copy.deepcopy(tmpl)
    runs = p.findall(qn('w:r'))
    r0 = copy.deepcopy(runs[0]) if runs else p.makeelement(qn('w:r'), {})
    for ch in list(p):
        if ch.tag != qn('w:pPr'):
            p.remove(ch)
    for ch in list(r0):
        if ch.tag != qn('w:rPr'):
            r0.remove(ch)
    t = r0.makeelement(W_T, {}); t.text = text; t.set(XS, 'preserve'); r0.append(t); p.append(r0)
    return p


def main():
    src, out, files = sys.argv[1], sys.argv[2], sys.argv[3:]
    d = docx.Document(src)
    body = d.element.body
    P = [el for el in body.iterchildren() if el.tag.endswith('}p')]
    T = [el for el in body.iterchildren() if el.tag.endswith('}tbl')]
    edits = []
    for f in files:
        for e in json.load(open(f, encoding='utf-8')):
            e['_f'] = f.split('/')[-1]; edits.append(e)
    ok, bad = 0, []
    deleted = []
    lastins = {}
    for e in edits:
        eid = e['id']; op = e.get('op', 'replace')
        try:
            if eid.startswith('P'):
                el = P[int(eid[1:])]
            elif eid.startswith('T'):
                k, r, c = eid[1:].split('.')
                tbl = T[int(k) - 1]
                tc = None
                for ri, tr in enumerate(tbl.findall(qn('w:tr'))):
                    if ri == int(r):
                        seen = 0
                        for x in tr.findall(qn('w:tc')):
                            if seen == int(c): tc = x
                            seen += 1
                if tc is None:
                    raise ValueError('cell not found')
                cands = [p for p in tc.findall(qn('w:p')) if e.get('old', '') in ''.join(t.text or '' for t in p.iter(W_T))]
                if not cands: raise ValueError('old not in cell')
                el = cands[0]
            else:
                raise ValueError('bad id')
            if op == 'replace':
                why = replace_once(Paragraph(el, d), e['old'], e.get('new', ''))
                if why: raise ValueError(why)
            elif op == 'delete':
                deleted.append(el)
            elif op in ('insert_after', 'insert_before'):
                tmpl = P[int(e['like'][1:])] if e.get('like') else el
                newp = clone_par(tmpl, e['text'])
                if op == 'insert_after':
                    # keep order of several inserts after the same anchor
                    last = lastins.get(id(el), el)
                    last.addnext(newp)
                    lastins[id(el)] = newp
                else:
                    el.addprevious(newp)
            ok += 1
        except Exception as ex:
            bad.append((e['_f'], eid, str(ex), (e.get('old') or e.get('text') or '')[:60]))
    for el in deleted:
        if el.getparent() is not None:
            el.getparent().remove(el)
    d.save(out)
    print('applied', ok, 'skipped', len(bad))
    for b in bad:
        print('  SKIP', b)


if __name__ == '__main__':
    main()
