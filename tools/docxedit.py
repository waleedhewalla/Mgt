"""Shared helpers for editing chapters inside a copy of the manuscript's document.xml."""
import copy
from lib import *

W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def _run(rpr, text):
    r = etree.Element(q('r'))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, q('t'))
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return r


def set_text(p, text):
    """Replace a paragraph's text, keeping paragraph props and the first non-empty run's formatting."""
    runs = p.findall(q('r'))
    if runs:
        tmpl = next((r for r in runs if ''.join(t.text or '' for t in r.iter(q('t'))).strip()), runs[0])
        rpr = tmpl.find(q('rPr'))
    else:                                   # empty paragraph: take the run props stored on the paragraph mark
        ppr = p.find(q('pPr'))
        rpr = ppr.find(q('rPr')) if ppr is not None else None
    for r in runs:
        p.remove(r)
    p.append(_run(rpr, text))
    return p


def set_runs(p, parts):
    """parts: [(text, source_run)] — each new run copies the source run's formatting."""
    rprs = [src.find(q('rPr')) for _, src in parts]
    rprs = [copy.deepcopy(x) if x is not None else None for x in rprs]
    for r in p.findall(q('r')):
        p.remove(r)
    for (text, _), rpr in zip(parts, rprs):
        p.append(_run(rpr, text))


def fresh(el):
    for e in el.iter():
        for a in ('paraId', 'textId'):
            e.attrib.pop('{%s}%s' % (W14, a), None)
    return el


def clone_after(el, text):
    new = fresh(copy.deepcopy(el))
    set_text(new, text)
    el.addnext(new)
    return new


def cell(tbl, row, col):
    return tbl.findall(q('tr'))[row].findall(q('tc'))[col]


def cell_text(tbl, row, col, text):
    tc = cell(tbl, row, col)
    ps = tc.findall(q('p'))
    set_text(ps[0], text)
    for extra in ps[1:]:
        tc.remove(extra)


def cell_lines(tbl, row, col, lines):
    """Put several lines in a cell as separate paragraphs cloned from the first one."""
    cell_text(tbl, row, col, lines[0])
    prev = cell(tbl, row, col).findall(q('p'))[0]
    for ln in lines[1:]:
        prev = clone_after(prev, ln)


def story_box(tbl, row, col, before, after):
    """Opening-story cells: text, a separator line run, then the reflection."""
    p = cell(tbl, row, col).findall(q('p'))[0]
    runs = p.findall(q('r'))
    sep = next(r for r in runs if '───' in ''.join(t.text or '' for t in r.iter(q('t'))))
    body = runs[0]
    set_runs(p, [(before, body), ('─' * 53, sep), (after, body)])


def delete_row(tbl, row):
    tr = tbl.findall(q('tr'))[row]
    tbl.remove(tr)


def add_row(tbl, like_row, texts):
    rows = tbl.findall(q('tr'))
    new = fresh(copy.deepcopy(rows[like_row]))
    rows[-1].addnext(new)
    for tc, text in zip(new.findall(q('tc')), texts):
        ps = tc.findall(q('p'))
        set_text(ps[0], text)
        for extra in ps[1:]:
            tc.remove(extra)
    return new


def insert_column(tbl, after_col, texts):
    """Clone column `after_col` and insert the copy right after it; keep the table's total width."""
    grid = tbl.find(q('tblGrid'))
    cols = grid.findall(q('gridCol'))
    widths = [int(c.get(q('w'))) for c in cols]
    total = sum(widths)
    widths.insert(after_col + 1, widths[after_col])
    scale = total / sum(widths)
    widths = [int(w * scale) for w in widths]
    widths[-1] += total - sum(widths)
    new_col = copy.deepcopy(cols[after_col])
    cols[after_col].addnext(new_col)
    for c, w in zip(grid.findall(q('gridCol')), widths):
        c.set(q('w'), str(w))
    for tr, text in zip(tbl.findall(q('tr')), texts):
        tcs = tr.findall(q('tc'))
        new = fresh(copy.deepcopy(tcs[after_col]))
        tcs[after_col].addnext(new)
        ps = new.findall(q('p'))
        set_text(ps[0], text)
        for extra in ps[1:]:
            new.remove(extra)
        for tc, w in zip(tr.findall(q('tc')), widths):
            tcw = tc.find('.//' + q('tcW'))
            if tcw is not None:
                tcw.set(q('w'), str(w))
