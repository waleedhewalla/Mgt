"""Dump a docx to ID-tagged text: P<i> for body paragraphs (document order), T<k>.<r>.<c> for table cells.
Adds [style] after the id when the paragraph has a named style."""
import sys, docx
from docx.table import Table
from docx.text.paragraph import Paragraph
src, out = sys.argv[1], sys.argv[2]
d = docx.Document(src)
lines, pi, ti = [], 0, 0
for el in d.element.body.iterchildren():
    if el.tag.endswith('}p'):
        p = Paragraph(el, d)
        t = p.text.strip()
        if t:
            st = p.style.name if p.style is not None and p.style.name != 'Normal' else ''
            lines.append(f'P{pi}{" ["+st+"]" if st else ""}: {t}')
        pi += 1
    elif el.tag.endswith('}tbl'):
        ti += 1
        t = Table(el, d)
        for r, row in enumerate(t.rows):
            seen = set()
            for c, cell in enumerate(row.cells):
                if id(cell._tc) in seen: continue
                seen.add(id(cell._tc))
                tx = cell.text.strip()
                if tx: lines.append(f'T{ti}.{r}.{c}: {tx}')
open(out, 'w').write('\n'.join(lines) + '\n')
print('lines', len(lines), 'paragraphs', pi, 'tables', ti)
