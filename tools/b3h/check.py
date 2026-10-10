"""Check an edits file against its dump: JSON, ids, overlaps, edits on deleted items, word count, leftover template phrases.
Usage: python3 check.py <part>"""
import sys, json, re
part = sys.argv[1]
lines = [l.split('\t', 1) for l in open(f'dumps/{part}.txt', encoding='utf-8').read().split('\n') if '\t' in l]
D = dict(lines)
order = [k for k, _ in lines]
pos = {k: i for i, k in enumerate(order)}
num = lambda k: int(re.search(r'(\d+)', k).group(1))
data = json.load(open(f'edits/{part}.json', encoding='utf-8'))
errs = []
deleted = set()
def tbl_ids(t):
    return [k for k in order if k == t or k.startswith('TP' + t[3:] + '.')]
ranges = []
for op in data['edits']:
    k = op.get('op')
    if k == 'delete':
        for i in op['ids']:
            if i not in D: errs.append('missing ' + i)
            deleted.update(tbl_ids(i) if i.startswith('TBL') else [i])
    elif k == 'delete_range':
        a, b = num(op['from']), num(op['to'])
        if op['from'] not in D or op['to'] not in D: errs.append(f'missing range end {op["from"]}..{op["to"]}')
        for r in ranges:
            if not (b < r[0] or a > r[1]): errs.append(f'overlap {a}-{b} with {r}')
        ranges.append((a, b))
        deleted.update(x for x in order if a <= num(x) <= b)
    elif k == 'replace_table':
        if op['id'] not in D: errs.append('missing ' + op['id'])
        deleted.update(x for x in tbl_ids(op['id']) if x != op['id'])
    elif k == 'delete_rows':
        t = op['id'][3:]
        for r in op['rows']:
            if r == 0: errs.append('row 0 deleted in ' + op['id'])
            deleted.update(x for x in order if x.startswith(f'TP{t}.{r}.'))
for op in data['edits']:
    k = op.get('op')
    if k in ('replace', 'insert_after'):
        if op['id'] not in D: errs.append('missing ' + op['id'])
        elif op['id'] in deleted: errs.append('edit on deleted ' + op['id'])
        if k == 'insert_after' and not op['id'].startswith('P'): errs.append('insert anchor not P ' + op['id'])
    elif k not in ('delete', 'delete_range', 'replace_table', 'delete_rows'):
        errs.append('unknown op ' + str(k))
new = dict(D)
for x in deleted:
    new.pop(x, None)
extra = []
for op in data['edits']:
    if op.get('op') == 'replace' and op['id'] in new:
        new[op['id']] = op['text']
    elif op.get('op') in ('insert_after', 'replace_table'):
        extra.append(op['text'])
    if op.get('op') == 'replace_table':
        new.pop(op['id'], None)
final = [v for k, v in new.items() if not k.startswith('TBL')] + extra
wb = sum(len(v.split()) for k, v in D.items() if not k.startswith('TBL'))
wa = sum(len(v.split()) for v in final)
txt = '\n'.join(final)
bad = {p: txt.count(p) for p in ['—', 'ولا يُثبت', 'الاعتراضات الثمانية', 'ثمانية اعتراضات', 'خلاصة الفصل', 'تأطير الفصل', 'مقدمة الفصل', 'يفتتح هذا الفصل', 'هذا الفصل']}
tables_before = sum(1 for k in D if k.startswith('TBL'))
tables_after = sum(1 for k in new if k.startswith('TBL'))
print(f'{part}: errors {len(errs)} | words {wb} -> {wa} ({(wa - wb) * 100 / wb:+.1f}%) | tables {tables_before} -> {tables_after} | leftovers {bad}')
for e in errs[:30]:
    print('  ', e)
